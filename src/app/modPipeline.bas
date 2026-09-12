Attribute VB_Name = "modPipeline"
Option Explicit

' ============================================================================
' modPipeline - PL-01/PL-02 の実行制御(S1からS4。T-24)
' ---
' 正: 14章§6(RunAll/RunStep の宣言)/14章§5(防衛線の順: P0つき抽出 ->
'   NormalizeLlmJson -> CheckSN -> 修復1回 -> 確定または失敗)/16章E-03・E-06・
'   E-07・E-49/13章§2.2(保存 data_key と参照優先)/15章§0.7/17章 T-24。
' 責務: 入力読取と打切り / ナレッジ注入と切詰め / プロンプト組立の委譲(15章の本文は
'   持たない) / CallStep / 防衛線と修復リトライ / 保存 / run_log を1呼び出し1行。
'   **シートに直接触れない**(R4)。案件一覧の読取は modCaseRead.ReadCaseCtx
'   (14章§6・裁定書7 B-7)が唯一の口で、読めなければ fail-closed で中止する。
' 判定核の分離(T-24): 修復要否・validate_result の3値分類・deep分岐・E-03の打切り
'   計画・E0301/E0302の切分けはシートもログも触らない純関数で、**16本は14章§6が
'   公開を宣言**(裁定書7 B-6)。層(a)=modTestsPure が直接叩く。
' last_ok_step / failed_step は modCaseStore.SetStepOutcome が唯一の書込口
'   (16章E-06・裁定書8 A-2)。成功経路=(stepNo, "")・失敗経路=(-1, sN) で呼ぶ。
' 入念モード(15章§4.5-4.7・16章E-35/E-36)の批判・改訂は modPipeline2 が持つ
'   (裁定書8 A-1)。ここは DeepEnabled が True のときの**1行の委譲**だけを置く。
' quality_mode は案件単位で上書きできる(13章§2.3・§2.10)。ui層が
'   hm_quality_mode を読んで RunStep / RunAll の qualityOverride へ渡す。
' ============================================================================

Private Const PL_SRC As String = "modPipeline"

' 13章§2.1 の enum と 14章§1 の play(10章§5: PL-01新規 / PL-02更新)。
Private Const PL_PLAY_NEW As String = "PL-01"
Private Const PL_PLAY_RENEWAL As String = "PL-02"
Private Const PL_TYPE_RENEWAL As String = "renewal"
Private Const PL_TIER_T1 As String = "t1_quick"
Private Const PL_MODE_DEEP As String = "deep"
Private Const PL_MODE_STD As String = "standard"
Private Const PL_STATUS_ERROR As String = "error"

' validate_result の3値(13章§2.4)と経路自体が失敗した場合の内部値。
Private Const PL_RES_OK As String = "ok"
Private Const PL_RES_REPAIRED As String = "repaired"
Private Const PL_RES_FAILED As String = "failed"
Private Const PL_RES_CALL As String = "call_failed"

Private Const PL_NONE_TEXT As String = "なし"
Private Const PL_OMIT_TEXT As String = "（一部省略）"

' 15章§2 S1 user の9貼付ブロック(AsmS1User 引数順)
Private Const PL_S1_KEYS As String = "input_hp|input_yuho|input_memo|input_contract|" & _
    "input_prev_renewal|input_dossier|input_field_notes|input_coverage_note|input_hearing_answers"
' 16章E-03(2) 切る順(ドシエ/前回更新/有報/メモ/HP)を上の添字で表す。
Private Const PL_CUT_IDX As String = "5|4|1|2|0"
Private Const PL_CUT_LABELS As String = "dossier|prev_renewal|yuho|memo|hp"
' 16章E-03(3) 打切らない4欄(現契約/現場メモ/付保見立て/ヒアリング回答)。
Private Const PL_KEEP_IDX As String = "3|6|7|8"
' 15章§0.7 切詰め順(成功事例/型/メニュー/種目/リスクライブラリ)。

' 16章E-07(ID幻覚=E0301)の検証ケースID(15章§11)。他の不合格は E0302。
Private Const PL_GHOST_CASES As String = "|V-S2-06|V-S3-03|V-S3-04|V-S3-05|V-S3-06|"

' 13章§2.3 の既定値(NFR-M3)と§0.7の予算配分。
Private Const PL_MAX_CTX_DFLT As Long = 40000
Private Const PL_T2_MAX_CTX_DFLT As Long = 100000
Private Const PL_PPT_MAX_DFLT As Long = 10
Private Const PL_REPAIR_DFLT As Long = 1
Private Const PL_PCT_KB As Long = 3
' 15章§6.1: 0行に切り詰めたスロットの代替文言(KbRowCount が0行と数える)。
Private Const PL_KB_ZERO As String = "なし"
Private Const PL_PCT_PASTE As Long = 7

' 1Step分の検証文脈。Check系は純関数でJSONの外側の文脈を引数で受ける(14章§6)。
Private Type TChkCtx
    stepNo As Long
    stepName As String
    caseId As String
    caseType As String
    dossierTier As String
    menusText As String
    linesText As String
    schemesText As String
    casesText As String
    incidentsText As String
    riskLibText As String
    s1Json As String
    s2Json As String
    prevS2Json As String
    fieldNoteProvided As Boolean
    lastErrs As String
End Type

' 案件一覧側の値のうち TCaseCtx に無い3つ(mQualityMode は RunStep の
' qualityOverride で上書きされたあとの【実効値】を持つ)。
Private mRoundNo As String
Private mS4Variant As String
Private mQualityMode As String

' RunAll - S1からS4の通し実行(14章§6)。1つでも失敗したらそこで止める。Step間で
'   DoEvents を挟む(16章E-50(c))。ui_lock と進捗表示は ui層の責務(R1)。
'   qualityOverride は素通しで各Stepへ渡す(4Step同じ品質モードで走らせる)。
Public Function RunAll(ByVal caseId As String, Optional ByVal qualityOverride As String) As Boolean
    Dim i As Long

    For i = 1 To 4
        If Not RunStep(caseId, i, qualityOverride) Then Exit Function
        DoEvents
    Next i
    RunAll = True
End Function

' RunStep - 1Stepの実行(14章§6)。True=検証合格まで到達
'   qualityOverride: 案件単位の quality_mode 上書き(13章§2.3・§2.10。ui層が
'     hm_quality_mode を読んで渡す)。空=config・ティア連動。案件一覧に保存は
'     しない。LibreOffice制約により Optional に既定値リテラルは書かない。
Public Function RunStep(ByVal caseId As String, ByVal stepNo As Long, _
                        Optional ByVal qualityOverride As String) As Boolean
    On Error GoTo Failed

    ' 裁定書10 M1: ここでリセットしない(RunAll の Step4 が Step2/3 の結末を
    ' 消すため)。リセット口は modPipeline2.ResetDeepOutcome(N9)だけ。
    If Not modCaseStore.IsValidCaseId(caseId) Then
        modLog.LogError "E0101", PL_SRC & ".RunStep", "invalid_case_id"
        Exit Function
    End If
    If stepNo < 1 Or stepNo > 4 Then
        modLog.LogError "E0101", PL_SRC & ".RunStep", "invalid_step:" & CStr(stepNo)
        Exit Function
    End If

    Dim ctx As TCaseCtx
    If Not LoadCtx(caseId, ctx) Then Exit Function
    ' 案件単位の上書きは config より優先(空なら config・ティア連動のまま)。
    If LenB(Trim$(qualityOverride)) > 0 Then mQualityMode = Trim$(qualityOverride)
    RunStep = ExecStep(caseId, ctx, stepNo)
    Exit Function

Failed:
    modLog.LogError "E0603", PL_SRC & ".RunStep", "step=" & CStr(stepNo), Err.Number
    RunStep = False
End Function

' LoadCtx - 1案件の文脈を読む【唯一の口】。modCaseRead.ReadCaseCtx へ委譲し、読め
'   なければ fail-closed(既定値で走らせるとAI利用枠を誤った前提で消費する)。
Private Function LoadCtx(ByVal caseId As String, ByRef ctx As TCaseCtx) As Boolean
    Dim rNo As Long
    Dim tierText As String

    ctx.case_type = vbNullString
    mRoundNo = vbNullString
    mS4Variant = vbNullString
    mQualityMode = vbNullString

    ' 第6引数は ctx.dossier_tier と同値。UDTの要素を ByRef で別名にしない
    ' (LibreOffice Basic で別名参照が壊れるのを避ける)。
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, rNo, mQualityMode, mS4Variant, _
                                   tierText) Then
        modLog.LogUsage "case_ctx_unavailable", caseId, "案件一覧を読めないため実行中止"
        Exit Function
    End If

    If rNo > 0 Then mRoundNo = CStr(rNo)
    If LenB(ctx.dossier_tier) = 0 Then ctx.dossier_tier = tierText
    LoadCtx = True
End Function

' ExecStep - 組立 -> 呼出 -> 防衛線 -> 保存 -> 状態遷移
Private Function ExecStep(ByVal caseId As String, ByRef ctx As TCaseCtx, _
                          ByVal stepNo As Long) As Boolean
    Dim c As TChkCtx
    Dim sysText As String, userText As String, schemaText As String
    Dim detailPre As String, resultText As String
    Dim okJson As String, failRaw As String

    c.stepNo = stepNo
    c.stepName = StepNameOf(stepNo)
    c.caseId = caseId
    c.caseType = ctx.case_type
    c.dossierTier = ctx.dossier_tier

    ' 注入IDの累積はStep開始時に初期化(14章§6 ResetInjectedIds)。
    modKnowledge.ResetInjectedIds
    If Not BuildPrompts(caseId, ctx, c, sysText, userText, schemaText, detailPre) Then Exit Function

    ' gatewayが知り得ない列を預ける(14章§6の run_log 受け渡し)。注入IDは組立で
    ' 確定済み。operator は config の同名キー(13章§2.3)。
    modGatewayRPN.SetRunContext caseId, mRoundNo, ctx.case_type, _
                                modKnowledge.LastInjectedIds(), _
                                Trim$(modConfig.GetStr("operator", vbNullString))

    resultText = CallGuarded(c, PlayIdOf(ctx.case_type), sysText, userText, _
                             schemaText, detailPre, okJson, failRaw)
    modGatewayRPN.ClearRunContext

    ' E-14/15/16: 経路側の失敗。案件状態は保持する(errorにしない)。
    If resultText = PL_RES_CALL Then Exit Function
    If resultText = PL_RES_FAILED Then
        FailStep caseId, c, failRaw
        Exit Function
    End If

    ' 13章§2.2: sN_json に入るのは検証合格済みの出力だけ。
    If Not modCaseStore.SaveData(caseId, c.stepName & "_json", okJson) Then
        modLog.LogError "E0604", PL_SRC & ".ExecStep", "save_failed:" & c.stepName
        Exit Function
    End If

    ' E-06: 成功したStepまでを last_ok_step へ確定し failed_step を空へ戻す。
    modCaseStore.SetStepOutcome caseId, stepNo, vbNullString
    modCaseStore.SetStatus caseId, StatusForStep(stepNo)

    ' 裁定書9 B10(10章 FR-13 Must): S3成功直後、確定した s3_json の
    ' unmatched_risks を新サービス候補としてナレッジへ自動記録する。
    If stepNo = 3 Then RecordServiceGaps caseId, ctx.industry_code, okJson

    ' 入念モード(15章§4.5-4.7)は modPipeline2 へ1行で委譲(裁定書8 A-1)。戻り値は
    ' 握りつぶす: 批判・改訂の不首尾で本体Stepを落とさない(E-35/E-36)。
    If DeepEnabled(ResolveQualityMode(mQualityMode, ctx.dossier_tier), stepNo) Then modPipeline2.RunDeep caseId, stepNo

    ExecStep = True
End Function

' RecordServiceGaps - 裁定書9 B10。確定した s3_json の unmatched_risks を1件ずつ
'   modKnowledge.AppendServiceGap へ渡す(FR-13。書込失敗時の退避・再送は
'   AppendServiceGap 側の責務で、本体Stepの成否は動かさない)。
Private Sub RecordServiceGaps(ByVal caseId As String, ByVal industryCode As String, _
                              ByVal s3Json As String)
    Dim it As Variant

    For Each it In modJsonLite.GetArrayItems(s3Json, "unmatched_risks")
        modKnowledge.AppendServiceGap caseId, industryCode, _
            modJsonLite.GetStr(CStr(it), "risk_name") & " / " & _
            modJsonLite.GetStr(CStr(it), "why_unmatched")
    Next it
End Sub

' BuildPrompts - Step別のsystem/user/schemaを組む(15章の本文は持たない)
Private Function BuildPrompts(ByVal caseId As String, ByRef ctx As TCaseCtx, _
                              ByRef c As TChkCtx, ByRef sysText As String, _
                              ByRef userText As String, ByRef schemaText As String, _
                              ByRef detailAcc As String) As Boolean
    Dim limitChars As Long
    Dim pasted(0 To 8) As String
    Dim s3Json As String, hearing As String, variantNote As String

    limitChars = ContextLimitOf(ctx.dossier_tier)
    If c.stepNo >= 2 Then
        c.s1Json = modCaseStore.ResolveStepJson(caseId, 1)
        If Not Upstream(c.s1Json, "s1") Then Exit Function
    End If
    If c.stepNo >= 3 Then
        c.s2Json = modCaseStore.ResolveStepJson(caseId, 2)
        If Not Upstream(c.s2Json, "s2") Then Exit Function
    End If

    Select Case c.stepNo
    Case 1
        LoadPasted caseId, limitChars, pasted, detailAcc, c
        sysText = modPromptsOps.AsmGuarded(modPromptsCore.BuildS1System())
        userText = modPipeline3.S1UserText(ctx, caseId, pasted)
        schemaText = modSchemas.SchemaS1()

    Case 2
        LoadKb ctx, c, limitChars, detailAcc
        c.prevS2Json = OrNone(modCaseStore.LoadData(caseId, "s2_prev_json"))
        hearing = OrNone(Sanitized(caseId, "input_hearing_answers", detailAcc))
        sysText = modPromptsOps.AsmGuarded(modPromptsCore.BuildS2System())
        userText = modPipeline3.S2UserText(ctx, caseId, c.s1Json, c.riskLibText, _
                                           c.menusText, c.prevS2Json, hearing, _
                                           c.incidentsText)
        schemaText = modSchemas.SchemaS2()

    Case 3
        LoadKb ctx, c, limitChars, detailAcc
        sysText = modPromptsOps.AsmGuarded(modPromptsCore.BuildS3System())
        userText = modPipeline3.S3UserText(ctx, caseId, S1SummaryOf(c.s1Json), c.s2Json, _
                                           c.menusText, c.linesText, c.schemesText, c.casesText)
        schemaText = modSchemas.SchemaS3()

    Case 4
        s3Json = modCaseStore.ResolveStepJson(caseId, 3)
        If Not Upstream(s3Json, "s3") Then Exit Function
        ' 15章§5: 想定外の s4_variant は proposal 扱い。黙って落とさず記録するのが
        ' 呼び出し側の責務。
        sysText = modPromptsOps.AsmS4System(mS4Variant, ctx.dossier_tier, variantNote)
        If LenB(variantNote) > 0 Then
            AddNote detailAcc, variantNote
            modLog.LogUsage "s4_variant_fallback", caseId, variantNote
        End If
        userText = modPromptsOps.AsmS4User(ctx, c.s1Json, c.s2Json, s3Json)
        schemaText = modSchemas.SchemaS4()
    End Select

    BuildPrompts = (LenB(userText) > 0)
End Function

' 上流Stepの成果物(解決は ResolveStepJson)が無ければ実行しない。
Private Function Upstream(ByVal jsonText As String, ByVal stepKey As String) As Boolean
    If LenB(Trim$(jsonText)) > 0 Then
        Upstream = True
        Exit Function
    End If
    modLog.LogError "E0101", PL_SRC & ".Upstream", "upstream_missing:" & stepKey
End Function

' LoadPasted - 9貼付欄の読取・無害化(E-04)・入力打切り(E-03)。E-03の順どおり
'   (1)ティアで上限 (2)切る順に1欄ずつ削り再計測 (3)4欄は削らない (4)4欄だけで超過
'   なら削らずE0102警告 (6)切詰め欄に注記+usage_log 記録(欄名と字数のみ)。
Private Sub LoadPasted(ByVal caseId As String, ByVal limitChars As Long, _
                       ByRef outText() As String, ByRef detailAcc As String, _
                       ByRef c As TChkCtx)
    Dim lens(0 To 5) As Long
    Dim plan As Variant
    Dim i As Long, idx As Long, budget As Long, labelText As String

    For i = 0 To 8
        outText(i) = Sanitized(caseId, PickAt(PL_S1_KEYS, i), detailAcc)
    Next i
    c.fieldNoteProvided = (LenB(Trim$(outText(6))) > 0)

    For i = 0 To 4
        lens(i) = Len(outText(IdxAt(PL_CUT_IDX, i)))
    Next i
    For i = 0 To 3
        lens(5) = lens(5) + Len(outText(IdxAt(PL_KEEP_IDX, i)))
    Next i

    budget = BudgetOf(limitChars, PL_PCT_PASTE)
    If ProtectedOverBudget(lens(5), budget) Then
        ' E-03(4): 自動では削らない。警告して人に削減を求める。
        AddNote detailAcc, "E0102:keep_over=" & CStr(lens(5))
        modLog.LogUsage "E0102", caseId, "keep_only_over_limit=" & CStr(lens(5))
    End If

    plan = TrimInputPlan(lens, budget)
    For i = 0 To 4
        If plan(i) < lens(i) Then
            idx = IdxAt(PL_CUT_IDX, i)
            outText(idx) = TruncField(outText(idx), plan(i))
            labelText = PickAt(PL_CUT_LABELS, i)
            AddNote detailAcc, "truncated:" & labelText & "=" & CStr(plan(i))
            modLog.LogUsage "E0102", caseId, labelText & "=" & CStr(plan(i))
        End If
    Next i
End Sub

' 外部由来テキストの読取と無害化(16章E-04)。件数だけ detail へ。
Private Function Sanitized(ByVal caseId As String, ByVal dataKey As String, _
                           ByRef detailAcc As String) As String
    Dim removedN As Long, markerN As Long

    Sanitized = modUtilText.SanitizeInput(modCaseStore.LoadData(caseId, dataKey), removedN, markerN)
    If removedN > 0 Then AddNote detailAcc, "e04_removed=" & CStr(removedN)
    If markerN > 0 Then AddNote detailAcc, "e04_marker=" & CStr(markerN)
End Function

' LoadKb - ナレッジ注入と15章§0.7の切詰め(6段)。手順の実体は modPipeline4 が
'   唯一持つ(T-57。modPipeline2 と写経していたものを畳んだ)。
Private Sub LoadKb(ByRef ctx As TCaseCtx, ByRef c As TChkCtx, _
                   ByVal limitChars As Long, ByRef detailAcc As String)
    Dim txt() As String

    modPipeline4.LoadKbSlots ctx, c.stepNo, BudgetOf(limitChars, PL_PCT_KB), txt, detailAcc
    c.casesText = txt(0)
    c.incidentsText = txt(1)
    c.schemesText = txt(2)
    c.menusText = txt(3)
    c.linesText = txt(4)
    c.riskLibText = txt(5)
End Sub

' 12章§3: S2=リスクライブラリ+メニュー要約+事故事例(15章§3。裁定書25 S6) /
'   S3=メニュー/種目/型/事例。slot番号は15章§0.7 の切詰め順(modPipeline4 冒頭)。
Public Function UsesSlot(ByVal stepNo As Long, ByVal slot As Long) As Boolean
    If stepNo = 2 Then UsesSlot = (slot = 1 Or slot = 3 Or slot = 5)
    If stepNo = 3 Then UsesSlot = (slot = 0 Or (slot >= 2 And slot <= 4))
End Function

' CallGuarded - 1Step分の呼出と防衛線(14章§5)。修復リトライは1回まで。
'   戻り値: ok / repaired / failed / call_failed。
Private Function CallGuarded(ByRef c As TChkCtx, ByVal playId As String, _
                             ByVal sysText As String, ByVal userText As String, _
                             ByVal schemaText As String, ByVal detailPre As String, _
                             ByRef okJson As String, ByRef failRaw As String) As String
    Dim rawText As String, errText As String, callOk As Boolean

    errText = OneCall(c, playId, sysText, userText, schemaText, detailPre, _
                      False, okJson, rawText, callOk)
    If Not callOk Then
        CallGuarded = PL_RES_CALL
        Exit Function
    End If
    If LenB(errText) = 0 Then
        CallGuarded = PL_RES_OK
        Exit Function
    End If

    c.lastErrs = errText
    failRaw = rawText
    If Not NeedsRepair(errText, RepairBudget()) Then
        CallGuarded = PL_RES_FAILED
        Exit Function
    End If

    ' 15章§7: 元のuser末尾へ修復サフィックスを足して同Stepを再呼び出しする。
    errText = OneCall(c, playId, sysText, _
                      userText & FillOne(modPromptsOps.RepairSuffix(), "validationErrors", errText), _
                      schemaText, "repair=1", True, okJson, rawText, callOk)
    If Not callOk Then
        CallGuarded = PL_RES_CALL
        Exit Function
    End If
    If LenB(errText) = 0 Then
        CallGuarded = PL_RES_REPAIRED
        Exit Function
    End If

    c.lastErrs = errText
    failRaw = rawText
    CallGuarded = PL_RES_FAILED
End Function

' OneCall - 1回の CallStep と防衛線(1)(2.5)(3)。run_log を必ず1行書く。
'   戻り値=検証エラー文(""=合格)。callOk=False は経路そのものの失敗。
Private Function OneCall(ByRef c As TChkCtx, ByVal playId As String, _
                         ByVal sysText As String, ByVal userText As String, _
                         ByVal schemaText As String, ByVal detailPre As String, _
                         ByVal isRepair As Boolean, ByRef outJson As String, _
                         ByRef outRaw As String, ByRef callOk As Boolean) As String
    Dim latencyMs As Long, detailAcc As String, errText As String

    detailAcc = detailPre
    outRaw = modGatewayRPN.CallStep(c.stepName, playId, sysText, userText, _
                                    schemaText, callOk, latencyMs)
    If Not callOk Then
        ' 検証へ到達していないので validate_result は空のまま記録する。
        RecordRun vbNullString, detailAcc
        Exit Function
    End If

    errText = Defend(c, outRaw, outJson, detailAcc)
    ' 裁定書37 B-03/B-05: 検証の**後ろ**で原文照合(S2)と充足度(S1)を detail へ
    ' 1行で足す。**落とさない・修復リトライを起こさない**(実体は modPipeline3)。
    modPipeline3.DefendNotes c.stepNo, c.caseId, outJson, c.s1Json, (LenB(errText) = 0), detailAcc
    If LenB(errText) > 0 Then AddNote detailAcc, "verr=" & FailCodeOf(errText)
    RecordRun ClassifyResult((LenB(errText) = 0), isRepair, (LenB(errText) = 0)), detailAcc
    OneCall = errText
End Function

' 防衛線(1)抽出 -> (2.5)重複排除 -> (3)検証。件数のみ detail へ。
Private Function Defend(ByRef c As TChkCtx, ByVal rawText As String, _
                        ByRef outJson As String, ByRef detailAcc As String) As String
    Dim extracted As String, removedCount As Long, fwCount As Long, extraCount As Long

    extracted = modJsonLite.ExtractJsonBlock(rawText)
    fwCount = modJsonLite.LastFwNormalized()
    extraCount = modJsonLite.LastExtraJson()
    If fwCount > 0 Then AddNote detailAcc, "fw_normalized=" & CStr(fwCount)
    If extraCount > 0 Then AddNote detailAcc, "extra_json=" & CStr(extraCount)

    If LenB(extracted) = 0 Then
        Defend = modPipeline3.MsgE0302()
        Exit Function
    End If

    outJson = modValidate.NormalizeLlmJson(c.stepName, extracted, removedCount)
    If removedCount > 0 Then
        ' E-49: 黙殺しない。件数を detail に残し E0303 を警告として記録。
        AddNote detailAcc, "e49_removed=" & CStr(removedCount)
        modLog.LogError "E0303", PL_SRC & "." & c.stepName, "e49_removed=" & CStr(removedCount)
    End If
    ' Step別の検証(14章§6の Check*)。改訂版(s2r/s3r)は本体と同じスキーマ
    ' (15章§4.7)。批判(s2c/s3c)の検証は T-28。
    Select Case c.stepName
    Case "s1"
        Defend = modValidate.CheckS1(outJson, c.caseType, c.fieldNoteProvided)
    Case "s2", "s2r"
        Defend = modValidate.CheckS2(outJson, c.caseType, c.menusText, c.prevS2Json, c.s1Json)
    Case "s3", "s3r"
        Defend = modValidate.CheckS3(outJson, c.s2Json, c.menusText, c.linesText, _
                                     c.schemesText, c.casesText, c.caseType, _
                                     S1SummaryOf(c.s1Json))
    Case "s4"
        Defend = modValidate.CheckS4(outJson, c.dossierTier, PptMaxSlides())
    End Select
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

' FailStep - E-06: 生応答は sN_json を汚さず sN_json_failed へ。状態は error。
'   ID幻覚(E-07)なら E0301、それ以外は E0302。
Private Sub FailStep(ByVal caseId As String, ByRef c As TChkCtx, ByVal rawText As String)
    modCaseStore.SaveData caseId, c.stepName & "_json_failed", rawText
    ' 16章 E-63: 経路側で本文が切られた疑いがあれば detail に印を1語足す
    '   (新しいエラーコードは作らない。裁定書33 C-4)。
    modLog.LogError FailCodeOf(c.lastErrs), PL_SRC & "." & c.stepName, _
                    "validate_failed:" & c.stepName & modRibbonWire.CutNote(rawText)
    ' E-06: failed_step に当該Stepを書き、last_ok_step は更新しない(-1=据置)。
    modCaseStore.SetStepOutcome caseId, -1, c.stepName
    modCaseStore.SetStatus caseId, PL_STATUS_ERROR
End Sub

' === 純関数(シート・ログ・LLMに触れない判定核)。14章§6が公開を宣言(裁定書7 B-6) ===

' 19章§4 の step 値(s1..s4。範囲外は "")。
Public Function StepNameOf(ByVal stepNo As Long) As String
    If stepNo < 1 Or stepNo > 4 Then Exit Function
    StepNameOf = "s" & CStr(stepNo)
End Function

' 13章§2.1 の status(Step成功時の遷移先)。
Public Function StatusForStep(ByVal stepNo As Long) As String
    If LenB(StepNameOf(stepNo)) = 0 Then Exit Function
    StatusForStep = StepNameOf(stepNo) & "_done"
End Function

Public Function PlayIdOf(ByVal caseType As String) As String
    If Trim$(caseType) = PL_TYPE_RENEWAL Then
        PlayIdOf = PL_PLAY_RENEWAL
    Else
        PlayIdOf = PL_PLAY_NEW
    End If
End Function

' 修復要否(14章§5(4))。検証エラーがあり再試行枠が残るときだけ。
Public Function NeedsRepair(ByVal errText As String, ByVal retryBudget As Long) As Boolean
    NeedsRepair = (LenB(errText) > 0 And retryBudget > 0)
End Function

' validate_result の3値(14章§5)。初回合格=ok / 修復後合格=repaired / 他=failed。
Public Function ClassifyResult(ByVal firstOk As Boolean, ByVal repairTried As Boolean, _
                                ByVal repairOk As Boolean) As String
    If Not repairTried Then
        If firstOk Then
            ClassifyResult = PL_RES_OK
        Else
            ClassifyResult = PL_RES_FAILED
        End If
        Exit Function
    End If
    If repairOk Then
        ClassifyResult = PL_RES_REPAIRED
    Else
        ClassifyResult = PL_RES_FAILED
    End If
End Function

' 不合格の内訳をコードへ(16章E-07=E0301 / E-06=E0302)。
Public Function FailCodeOf(ByVal errText As String) As String
    Dim parts() As String
    Dim i As Long
    FailCodeOf = "E0302"
    If LenB(errText) = 0 Then Exit Function
    parts = Split(errText, vbLf)
    For i = LBound(parts) To UBound(parts)
        If InStr(1, PL_GHOST_CASES, "|" & CaseIdOfLine(parts(i)) & "|", vbTextCompare) > 0 Then
            FailCodeOf = "E0301"
            Exit Function
        End If
    Next i
End Function

' 検証エラー行の先頭からケースID(15章§0 原則10)。
Public Function CaseIdOfLine(ByVal lineText As String) As String
    Dim t As String, p As Long

    t = Trim$(lineText)
    If Left$(t, 1) <> "[" Then Exit Function
    p = InStr(1, t, "]", vbBinaryCompare)
    If p <= 2 Then Exit Function
    CaseIdOfLine = Mid$(t, 2, p - 2)
End Function

' 13章§2.3: quality_mode は案件で上書き可。空はティア連動(t1=standard/他=deep)。
Public Function ResolveQualityMode(ByVal cfgMode As String, ByVal tier As String) As String
    Dim m As String

    m = LCase$(Trim$(cfgMode))
    If m = PL_MODE_DEEP Or m = PL_MODE_STD Then
        ResolveQualityMode = m
        Exit Function
    End If
    If Trim$(tier) = PL_TIER_T1 Then
        ResolveQualityMode = PL_MODE_STD
    Else
        ResolveQualityMode = PL_MODE_DEEP
    End If
End Function

' deep分岐は S2/S3 だけ(15章§4.5-4.7)。
Public Function DeepEnabled(ByVal qualityMode As String, ByVal stepNo As Long) As Boolean
    DeepEnabled = (qualityMode = PL_MODE_DEEP And (stepNo = 2 Or stepNo = 3))
End Function

' 16章E-03(1): ティアで上限config。
Private Function ContextLimitOf(ByVal tier As String) As Long
    If Trim$(tier) = PL_TIER_T1 Then
        ContextLimitOf = modConfig.GetLong("max_context_chars", PL_MAX_CTX_DFLT)
    Else
        ContextLimitOf = modConfig.GetLong("t2_max_context_chars", PL_T2_MAX_CTX_DFLT)
    End If
End Function

' 15章§0.7 の予算配分(貼付7割 / ナレッジ3割)。
Public Function BudgetOf(ByVal limitChars As Long, ByVal pct As Long) As Long
    If limitChars <= 0 Then Exit Function
    BudgetOf = CLng(Fix(limitChars / 10)) * pct
End Function

' 打切らない4欄だけで上限超過か(E-03(4)の実行前警告の条件)。
Public Function ProtectedOverBudget(ByVal keepChars As Long, ByVal budgetChars As Long) As Boolean
    If budgetChars <= 0 Then Exit Function
    ProtectedOverBudget = (keepChars > budgetChars)
End Function

' TrimInputPlan - 16章E-03(2)の打切りを【計画するだけ】の純関数。lens(0..4)=切る順の
'   現在字数 / lens(5)=打切らない4欄の合計。budgetChars<=0 は上限なし。戻り値=5要素
'   の「残してよい字数」。1欄ずつ削り再計測し上限を下回った時点で止める(4欄だけで
'   超過する場合は E-03(4) により自動では削らない)。
Public Function TrimInputPlan(ByRef lens() As Long, ByVal budgetChars As Long) As Long()
    Dim res(0 To 4) As Long
    Dim i As Long, total As Long, over As Long, cut As Long

    For i = 0 To 4
        res(i) = lens(i)
        total = total + res(i)
    Next i
    total = total + lens(5)

    If budgetChars <= 0 Or total <= budgetChars Or lens(5) > budgetChars Then
        TrimInputPlan = res
        Exit Function
    End If

    For i = 0 To 4
        over = total - budgetChars
        If over <= 0 Then Exit For
        cut = over
        If cut > res(i) Then cut = res(i)
        res(i) = res(i) - cut
        total = total - cut
    Next i
    TrimInputPlan = res
End Function

' 切詰めた欄には必ず注記を付す(E-03(6))。
Public Function TruncField(ByVal s As String, ByVal allowedChars As Long) As String
    If allowedChars <= 0 Then
        TruncField = PL_OMIT_TEXT
        Exit Function
    End If
    If Len(s) <= allowedChars Then
        TruncField = s
        Exit Function
    End If
    TruncField = modUtil.SafeLeft(s, allowedChars) & PL_OMIT_TEXT
End Function

' 注入テキストの行数(1行1件・vbLf区切り)。0行の既定文言は0。
Public Function KbRowCount(ByVal s As String) As Long
    Dim t As String

    t = Trim$(s)
    If LenB(t) = 0 Then Exit Function
    If Left$(t, 1) = "(" Or t = PL_KB_ZERO Then Exit Function
    KbRowCount = UBound(Split(t, vbLf)) + 1
End Function

' 15章 S3/S3C の {{s1SummaryJson}}(4キーだけのJSON。他のキーは含めない)。
Public Function S1SummaryOf(ByVal s1Json As String) As String
    Dim s As String

    s = "{""business_summary"": " & JStr(modJsonLite.GetStr(s1Json, "business_summary"))
    s = s & ", ""strategy_outlook"": {""mvv"": " & JStr(modJsonLite.GetStr(s1Json, "mvv"))
    s = s & ", ""aspirations"": " & JArr(s1Json, "aspirations", True)
    s = s & ", ""market_context"": " & JStr(modJsonLite.GetStr(s1Json, "market_context")) & "}"
    s = s & ", ""current_coverage"": " & JArr(s1Json, "current_coverage", False)
    s = s & ", ""field_insights"": " & JArr(s1Json, "field_insights", False)
    S1SummaryOf = s & "}"
End Function

' "a|b|c" の idx 番目(0始まり。範囲外は "")。
Private Function PickAt(ByVal listText As String, ByVal idx As Long) As String
    Dim parts() As String

    parts = Split(listText, "|")
    If idx < 0 Or idx > UBound(parts) Then Exit Function
    PickAt = parts(idx)
End Function

Private Function IdxAt(ByVal listText As String, ByVal idx As Long) As Long
    IdxAt = CLng(Val(PickAt(listText, idx)))
End Function

Private Function OrNone(ByVal s As String) As String
    If LenB(Trim$(s)) = 0 Then
        OrNone = PL_NONE_TEXT
    Else
        OrNone = s
    End If
End Function

Private Function JStr(ByVal s As String) As String
    JStr = """" & modJsonLite.EscapeJsonStr(s) & """"
End Function

' 配列を素のJSONへ(asText=True は要素を文字列として囲む)。
Private Function JArr(ByVal srcJson As String, ByVal keyName As String, _
                      ByVal asText As Boolean) As String
    Dim acc As String, n As Long
    Dim it As Variant

    For Each it In modJsonLite.GetArrayItems(srcJson, keyName)
        If n > 0 Then acc = acc & ", "
        If asText Then
            acc = acc & JStr(CStr(it))
        Else
            acc = acc & CStr(it)
        End If
        n = n + 1
    Next it
    JArr = "[" & acc & "]"
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
    RepairBudget = modConfig.GetLong("json_repair_retry", PL_REPAIR_DFLT)
End Function

Private Function PptMaxSlides() As Long
    PptMaxSlides = modConfig.GetLong("ppt_max_slides_t2", PL_PPT_MAX_DFLT)
End Function
