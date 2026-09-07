Attribute VB_Name = "modPlayOps"
Option Explicit

' ============================================================================
' modPlayOps - PL-03 プリフライト診断の実行制御(app層・T-25)
' ----------------------------------------------------------------------------
' 正: 14章§6(RunPreflight の宣言)/14章§1(Stepレジストリ: step=pf・play=PL-03・
'   呼び出し元は modPlayOps)/14章§5(防衛線の順: 抽出 -> NormalizeLlmJson ->
'   CheckPF -> 修復1回 -> 確定または失敗)/15章§6・§6.1(PFのプロンプトと注入
'   テキストの1行書式)/13章§2.6(pf_json / pf_survival / pf_pred_types)/
'   16章 E-40(一括診断の途中失敗)・E-07(ID幻覚)。
' 責務: 受信箱1件を読み、ナレッジを注入してPFを1回呼び、防衛線を通して結果を
'   受信箱へ返す。**シートに直接触れない**(R4。受信箱I/Oは modInboxStore が
'   唯一の口)。プロンプト本文は持たない(15章の本文は modPromptsOps)。
' 判定核の分離: 診断結果の要約列(pf_survival / pf_pred_types)の取り出しと、
'   不合格の内訳をエラーコードへ落とす切分け、ID実在検査へ渡す一覧の組立は
'   シートもログもLLMも触らない純関数として公開する(14章§6)。層(a)=
'   modTestsPure から直接叩けることが要件。
' run_log(13章§2.4): 受信箱起点なので round_no は空・case_id 列に inbox_id を
'   入れる。1呼び出し=1行(SetRunContext -> CallStep -> TakeLastRun -> LogRun)。
' ============================================================================

Private Const PO_SRC As String = "modPlayOps"

' 14章§1 Stepレジストリ(pf / PL-03)。
Private Const PO_STEP_PF As String = "pf"
Private Const PO_PLAY_PF As String = "PL-03"

' 13章§2.6 pf_pred_types の区切り。
Private Const PO_TYPE_SEP As String = ";"

' 16章 E-07(ID幻覚=E0301)にあたるPFの検証ケースID(15章§6)。他の不合格は E0302。
Private Const PO_GHOST_CASES As String = "|V-PF-03|"

' config 既定(13章§2.3・NFR-M3)。
Private Const PO_REPAIR_DFLT As Long = 1

' CallGuarded の戻り値(合否と経路失敗の3値)。
Private Const PO_RES_OK As String = "ok"
Private Const PO_RES_FAILED As String = "failed"
Private Const PO_RES_CALL As String = "call_failed"

' --- 純関数(シート・ログ・LLMに触れない判定核。14章§6) ---

' PfSurvivalOf - 診断JSONから受信箱の要約列 pf_survival を取り出す(13章§2.6)。
'   enum(high / mid / low)以外は "" を返す(未定義値を列へ書かない)。
Public Function PfSurvivalOf(ByVal pfJson As String) As String
    Dim v As String
    v = Trim$(modJsonLite.GetStr(pfJson, "survival"))
    Select Case v
        Case "high", "mid", "low"
            PfSurvivalOf = v
    End Select
End Function

' PfPredTypesOf - 診断JSONから pf_pred_types(";"区切り T1からT10)を組む
'   (13章§2.6)。enum外の値は落とす。重複は1件へ寄せる(modUtil.AppendIdList)。
'   T0 は PF の predicted_drop_types には現れない(19章§3。フィードバック入力用)。
Public Function PfPredTypesOf(ByVal pfJson As String) As String
    Dim acc As String
    Dim it As Variant
    For Each it In modJsonLite.GetArrayItems(pfJson, "predicted_drop_types")
        Dim v As String
        v = Trim$(CStr(it))
        If IsDropTypeValue(v) Then acc = modUtil.AppendIdList(acc, v)
    Next it
    PfPredTypesOf = acc
End Function

' PfRefIds - CheckPF の V-PF-03(ref_id 実在検査)へ渡す一覧テキストを組む。
'   15章§6 が PF へ注入する5種(判断基準 / メニュー要約 / 型ライブラリ /
'   座組パターン / 研究中テーマ)をそのまま連結したもので、行頭が "[ID] " の
'   1行書式(15章§6.1)を保つ。空の注入は行ごと落とす(空行を作らない)。
'   **一覧を渡さない呼び出しは fail-closed で不合格**(裁定書7 A-2)なので、
'   PFを結線する側は必ず本関数の戻り値を渡す。
Public Function PfRefIds(ByVal rulesText As String, ByVal menusSummary As String, _
                         ByVal schemesText As String, ByVal patternsText As String, _
                         ByVal researchingText As String) As String
    Dim acc As String
    acc = AppendBlock(acc, rulesText)
    acc = AppendBlock(acc, menusSummary)
    acc = AppendBlock(acc, schemesText)
    acc = AppendBlock(acc, patternsText)
    acc = AppendBlock(acc, researchingText)
    PfRefIds = acc
End Function

' PfFailCodeOf - PFの不合格の内訳をエラーコードへ。**ID幻覚(V-PF-03)を含めば
'   E0301(16章 E-07)、それ以外は E0302(E-06)**。判定は検証エラー行の先頭の
'   ケースID(15章§0 原則10の書式)で行う。errText が空でも E0302 を返す
'   (呼び出し側は不合格のときしか使わない)。
Public Function PfFailCodeOf(ByVal errText As String) As String
    Dim parts() As String
    Dim i As Long
    PfFailCodeOf = "E0302"
    If LenB(errText) = 0 Then Exit Function
    parts = Split(errText, vbLf)
    For i = LBound(parts) To UBound(parts)
        If InStr(1, PO_GHOST_CASES, "|" & CaseIdOfPfLine(parts(i)) & "|", vbTextCompare) > 0 Then
            PfFailCodeOf = "E0301"
            Exit Function
        End If
    Next i
End Function

' CaseIdOfPfLine - 検証エラー1行の先頭 `[ケースID] ` からケースIDを取り出す
'   (15章§0 原則10の書式)。不一致は ""。
Public Function CaseIdOfPfLine(ByVal lineText As String) As String
    Dim t As String, p As Long

    t = Trim$(lineText)
    If Left$(t, 1) <> "[" Then Exit Function
    p = InStr(1, t, "]", vbBinaryCompare)
    If p <= 2 Then Exit Function
    CaseIdOfPfLine = Mid$(t, 2, p - 2)
End Function

' 19章§3 の drop_type のうち PF が返しうる T1からT10 か。
Private Function IsDropTypeValue(ByVal v As String) As Boolean
    IsDropTypeValue = (InStr(1, "|T1|T2|T3|T4|T5|T6|T7|T8|T9|T10|", _
                             "|" & v & "|", vbBinaryCompare) > 0)
End Function

' 空でない注入テキストだけを vbLf で継ぐ(空行を作らない)。
Private Function AppendBlock(ByVal acc As String, ByVal blockText As String) As String
    Dim t As String
    t = Trim$(blockText)
    If LenB(t) = 0 Then
        AppendBlock = acc
        Exit Function
    End If
    If LenB(acc) = 0 Then
        AppendBlock = t
    Else
        ' 15章§6.1 は1行1件なので vbLf で継ぐ(空行を挟まない)。
        AppendBlock = acc & vbLf & t
    End If
End Function

' --- 実行制御 ---

' RunPreflight - 受信箱1件のプリフライト診断(14章§6・15章§6)。
'   True=検証合格まで到達し pf_json / pf_survival / pf_pred_types を保存できた。
'   False=前提不足・経路失敗・検証不合格(**当該行は undiagnosed のまま残す**。
'   16章 E-40 の「失敗分は undiagnosed のまま」を1件単位でも守る)。
Public Function RunPreflight(ByVal inboxId As String) As Boolean
    On Error GoTo Failed

    If Not modInboxStore.IsValidInboxId(inboxId) Then
        modLog.LogError "E0101", PO_SRC & ".RunPreflight", "invalid_inbox_id"
        Exit Function
    End If

    Dim themeText As String, bodyText As String, statusText As String
    If Not modInboxStore.ReadInboxItem(inboxId, themeText, bodyText, statusText) Then
        modLog.LogUsage "inbox_item_unavailable", inboxId, "受信箱を読めないため診断中止"
        Exit Function
    End If
    If LenB(Trim$(themeText)) = 0 And LenB(Trim$(bodyText)) = 0 Then
        modLog.LogError "E0101", PO_SRC & ".RunPreflight", "empty_post"
        Exit Function
    End If

    ' 15章§6.1 の5種を注入する。注入IDの累積はStep開始時に初期化(14章§6)。
    modKnowledge.ResetInjectedIds
    Dim rulesText As String, menusSummary As String, schemesText As String
    Dim patternsText As String, researchingText As String
    rulesText = modKnowledge.RulesText()
    menusSummary = modKnowledge.MenusSummaryFor(vbNullString)
    schemesText = modKnowledge.SchemesFor(vbNullString)
    patternsText = modKnowledge.PatternsText()
    researchingText = modKnowledge.ResearchingText()

    ' 16章 E-04: 外部由来テキスト(投函本文)は組立の前に無害化する。
    Dim detailPre As String
    Dim safeTheme As String, safeBody As String
    safeTheme = Sanitized(themeText, detailPre)
    safeBody = Sanitized(bodyText, detailPre)

    Dim sysText As String, userText As String, schemaText As String
    sysText = modPromptsOps.BuildPFSystem()
    userText = modPromptsOps.AsmPFUser(safeTheme, safeBody, rulesText, menusSummary, _
                                       schemesText, patternsText, researchingText)
    schemaText = modSchemas.SchemaPF()

    Dim refIds As String
    refIds = PfRefIds(rulesText, menusSummary, schemesText, patternsText, researchingText)

    ' run_log の「gatewayが知り得ない列」を預ける。受信箱起点は round_no 空
    ' (13章§2.4)。case_type はPFに無いので空(mockのバリアントは common)。
    modGatewayRPN.SetRunContext inboxId, vbNullString, vbNullString, _
                                modKnowledge.LastInjectedIds(), vbNullString

    Dim okJson As String
    Dim resultText As String
    resultText = CallGuarded(inboxId, sysText, userText, schemaText, detailPre, _
                             refIds, okJson)
    modGatewayRPN.ClearRunContext

    If resultText <> PO_RES_OK Then Exit Function

    If Not modInboxStore.SavePfResult(inboxId, okJson, PfSurvivalOf(okJson), _
                                      PfPredTypesOf(okJson)) Then
        modLog.LogError "E0604", PO_SRC & ".RunPreflight", "save_failed:" & inboxId
        Exit Function
    End If

    RunPreflight = True
    Exit Function

Failed:
    modLog.LogError "E0603", PO_SRC & ".RunPreflight", "unexpected", Err.Number
    RunPreflight = False
End Function

' RunPreflightAll - 未診断の投函をまとめて診断する(17章 T-25 の一括診断)。
'   16章 E-40: **診断できた分は保存し、失敗した分は undiagnosed のまま残す**。
'   途中で止めず最後まで回し、1件でも落ちたら E0701 を件数つきで記録する
'   (どの件が落ちたかは各件の err_log に残る)。戻り値=診断できた件数。
'   Step間で DoEvents を挟む(16章 E-50(c))。
Public Function RunPreflightAll() As Long
    On Error GoTo Failed

    Dim idList As String
    idList = modInboxStore.UndiagnosedIds()
    If LenB(idList) = 0 Then Exit Function

    Dim ids() As String
    ids = modUtil.SplitKeepNonEmpty(idList, PO_TYPE_SEP)

    Dim okCount As Long, ngCount As Long
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If RunPreflight(ids(i)) Then
            okCount = okCount + 1
        Else
            ngCount = ngCount + 1
        End If
        DoEvents
    Next i

    If ngCount > 0 Then
        modLog.LogError "E0701", PO_SRC & ".RunPreflightAll", _
                        "diagnosed=" & CStr(okCount) & " left_undiagnosed=" & CStr(ngCount)
    End If
    modLog.LogUsage "inbox_batch_diagnose", vbNullString, _
                    "ok=" & CStr(okCount) & " ng=" & CStr(ngCount)
    RunPreflightAll = okCount
    Exit Function

Failed:
    modLog.LogError "E0701", PO_SRC & ".RunPreflightAll", "unexpected", Err.Number
    RunPreflightAll = 0
End Function

' --- 内部ヘルパー(防衛線と記録) ---

' CallGuarded - PF1回分の呼出と防衛線(14章§5)。修復リトライは1回まで。
'   戻り値: "ok"(初回または修復後に合格) / "failed"(不合格) /
'   "call_failed"(経路そのものの失敗)。合格したJSONは okJson へ返す。
'   初回合格か修復後合格かは run_log の validate_result(ok / repaired)が持つ
'   (13章§2.4)ので、呼び出し側へは合否だけを返す。
Private Function CallGuarded(ByVal inboxId As String, ByVal sysText As String, _
                             ByVal userText As String, ByVal schemaText As String, _
                             ByVal detailPre As String, ByVal refIds As String, _
                             ByRef okJson As String) As String
    Dim errText As String, rawText As String
    Dim callOk As Boolean

    errText = OneCall(sysText, userText, schemaText, detailPre, refIds, False, _
                      okJson, rawText, callOk)
    If Not callOk Then
        CallGuarded = PO_RES_CALL
        Exit Function
    End If
    If LenB(errText) = 0 Then
        CallGuarded = PO_RES_OK
        Exit Function
    End If

    If modPipeline.NeedsRepair(errText, RepairBudget()) Then
        errText = OneCall(sysText, _
                          userText & FillOne(modPromptsOps.RepairSuffix(), _
                                             "validationErrors", errText), _
                          schemaText, "repair=1", refIds, True, okJson, rawText, callOk)
        If Not callOk Then
            CallGuarded = PO_RES_CALL
            Exit Function
        End If
        If LenB(errText) = 0 Then
            CallGuarded = PO_RES_OK
            Exit Function
        End If
    End If

    ' 不合格。生応答は受信箱の pf_json を汚さない(13章§2.6 の pf_json には
    ' 検証に合格した診断だけを入れる)。内訳のコードだけを err_log へ残す。
    ' 16章 E-63 の印(裁定書33 C-4)。modPipeline.FailStep と同じ1語を添える。
    modLog.LogError PfFailCodeOf(errText), PO_SRC & "." & PO_STEP_PF, _
                    "validate_failed:" & inboxId & modRibbonWire.CutNote(rawText)
    CallGuarded = PO_RES_FAILED
End Function

' OneCall - 1回の CallStep と防衛線(1)(2.5)(3)。run_log を必ず1行書く。
'   戻り値=検証エラー文(""=合格)。callOk=False は経路そのものの失敗。
Private Function OneCall(ByVal sysText As String, ByVal userText As String, _
                         ByVal schemaText As String, ByVal detailPre As String, _
                         ByVal refIds As String, ByVal isRepair As Boolean, _
                         ByRef outJson As String, ByRef outRaw As String, _
                         ByRef callOk As Boolean) As String
    Dim latencyMs As Long
    Dim detailAcc As String
    Dim errText As String

    detailAcc = detailPre
    outRaw = modGatewayRPN.CallStep(PO_STEP_PF, PO_PLAY_PF, sysText, userText, _
                                    schemaText, callOk, latencyMs)
    If Not callOk Then
        ' 検証へ到達していないので validate_result は空のまま記録する。
        RecordRun vbNullString, detailAcc
        Exit Function
    End If

    errText = Defend(outRaw, refIds, outJson, detailAcc)
    If LenB(errText) > 0 Then AddNote detailAcc, "verr=" & PfFailCodeOf(errText)
    RecordRun modPipeline.ClassifyResult((LenB(errText) = 0), isRepair, _
                                         (LenB(errText) = 0)), detailAcc
    OneCall = errText
End Function

' 防衛線(1)抽出 -> (2.5)重複排除 -> (3)CheckPF。件数のみ detail へ。
Private Function Defend(ByVal rawText As String, ByVal refIds As String, _
                        ByRef outJson As String, ByRef detailAcc As String) As String
    Dim extracted As String
    Dim removedCount As Long, fwCount As Long, extraCount As Long

    extracted = modJsonLite.ExtractJsonBlock(rawText)
    fwCount = modJsonLite.LastFwNormalized()
    extraCount = modJsonLite.LastExtraJson()
    If fwCount > 0 Then AddNote detailAcc, "fw_normalized=" & CStr(fwCount)
    If extraCount > 0 Then AddNote detailAcc, "extra_json=" & CStr(extraCount)

    If LenB(extracted) = 0 Then
        Defend = "[E0302] 応答からJSONを抽出できませんでした（説明文のみ・括弧の欠落など）"
        Exit Function
    End If

    outJson = modValidate.NormalizeLlmJson(PO_STEP_PF, extracted, removedCount)
    If removedCount > 0 Then
        ' E-49: 黙殺しない。件数を detail に残し E0303 を警告として記録。
        AddNote detailAcc, "e49_removed=" & CStr(removedCount)
        modLog.LogError "E0303", PO_SRC & "." & PO_STEP_PF, "e49_removed=" & CStr(removedCount)
    End If
    Defend = modValidate.CheckPF(outJson, refIds)
End Function

' run_log を1行書く(gatewayの保留行へ validate_result と detail を足す)。
Private Sub RecordRun(ByVal validateResult As String, ByVal detailText As String)
    Dim rec As TRunLogRec

    If Not modGatewayRPN.TakeLastRun(rec) Then Exit Sub
    rec.validate_result = validateResult
    rec.injected_kb_ids = modKnowledge.LastInjectedIds()
    If LenB(detailText) > 0 Then
        If LenB(rec.detail) > 0 Then rec.detail = rec.detail & ";"
        rec.detail = rec.detail & detailText
    End If
    modLog.LogRun rec
End Sub

' 外部由来テキストの無害化(16章 E-04)。件数だけ detail へ(本文は書かない)。
Private Function Sanitized(ByVal rawText As String, ByRef detailAcc As String) As String
    Dim removedN As Long, markerN As Long

    Sanitized = modUtilText.SanitizeInput(rawText, removedN, markerN)
    If removedN > 0 Then AddNote detailAcc, "e04_removed=" & CStr(removedN)
    If markerN > 0 Then AddNote detailAcc, "e04_marker=" & CStr(markerN)
End Function

' run_log detail の積み方(400字上限は modLog.TruncDetail)。
Private Sub AddNote(ByRef acc As String, ByVal noteText As String)
    If LenB(noteText) = 0 Then Exit Sub
    If LenB(acc) > 0 Then acc = acc & ";"
    acc = acc & noteText
End Sub

' テンプレートの {{name}} を1つ埋める(置換の口は modPromptsOps.Fill のみ)。
Private Function FillOne(ByVal tpl As String, ByVal nameText As String, _
                         ByVal valText As String) As String
    Dim names(0 To 0) As String, vals(0 To 0) As String

    names(0) = nameText
    vals(0) = valText
    FillOne = modPromptsOps.Fill(tpl, names, vals)
End Function

Private Function RepairBudget() As Long
    RepairBudget = modConfig.GetLong("json_repair_retry", PO_REPAIR_DFLT)
End Function
