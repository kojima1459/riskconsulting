Attribute VB_Name = "modPipeline5"
Option Explicit

' ============================================================================
' modPipeline5 - Step5(顧客向け提案書 S5)の実行(15章§5.6・20章・docs/29 §5)
' ----------------------------------------------------------------------------
' 12章§2 の30,000字契約による modPipeline の分割先。S1からS4 が modPipeline の
'   RunStep で回るのに対し、S5 は
'     ・入力が S1+S2+S3(+S4 があれば)で、案件の status を動かさない
'     ・呼び出しは**1回＋修復1回**(docs/29 §5.1)
'     ・禁止語だけが残ったときは機械置換で生成を続ける(docs/29 §5.3)
'   という別の流れなので、modPipeline の Select Case を太らせず本モジュールへ置く。
'
' 責務の分界(答えを2箇所に書かない):
'   プロンプト本文 = modPromptsS5 / スキーマ = modSchemas2.SchemaS5 /
'   検証 = modValidate4.CheckS5 / 実数の数え上げ = modExportProposal.StatsText /
'   修復サフィックス = modPromptsOps.RepairSuffix / 呼び出し = modGatewayRPN。
'   本モジュールが持つのは**組立と流れ**だけである。
'
' R4(12章§2): シート・ブックに触れない。案件データは modCaseStore / modCaseRead
'   経由、config は modConfig 経由。CP932準拠(15章§0 原則7)。
' ============================================================================

Private Const P5_SRC As String = "modPipeline5"
Private Const P5_STEP As String = "s5"
Private Const P5_REPAIR_DFLT As Long = 1
Private Const P5_RES_OK As String = "ok"
Private Const P5_RES_REPAIRED As String = "repaired"
Private Const P5_RES_FAILED As String = "failed"
Private Const P5_RES_CALL As String = "call_failed"
' 15章§5.6 のプレースホルダ名(組立層が埋める5つ)。
Private Const P5_PH As String = "company|s1Json|s2Json|s3Json|statsText"
Private Const P5_PH_SEP As String = "|"

' 直近の検証エラー(FailCodeOf の材料)。run_log と err_log の両方で使う。
' モジュール変数は全プロシージャより前に置く(VBAの宣言部の規約)。
Private mLastErrs As String

' ============================================================================
' RunStep5 - 1案件ぶんの S5 実行。True=検証合格まで到達。
'   案件の status は動かさない(提案書の生成はパイプラインの段ではなく出力の
'   準備であり、16章 E-48 と同じ扱い)。失敗しても S1からS4 の成果物は残る。
' ============================================================================
Public Function RunStep5(ByVal caseId As String) As Boolean
    On Error GoTo Failed

    If Not modCaseStore.IsValidCaseId(caseId) Then
        modLog.LogError "E0101", P5_SRC & ".RunStep5", "invalid_case_id"
        Exit Function
    End If

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String, s4Variant As String, tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        modLog.LogUsage "case_ctx_unavailable", caseId, "案件一覧を読めないため実行中止"
        Exit Function
    End If

    Dim s1Json As String, s2Json As String, s3Json As String
    s1Json = modCaseStore.ResolveStepJson(caseId, 1)
    s2Json = modCaseStore.ResolveStepJson(caseId, 2)
    s3Json = modCaseStore.ResolveStepJson(caseId, 3)
    If LenB(Trim$(s1Json)) = 0 Or LenB(Trim$(s2Json)) = 0 Or LenB(Trim$(s3Json)) = 0 Then
        modLog.LogUsage "s5_skipped", caseId, "upstream_missing"
        Exit Function
    End If

    ' 15章§1.3: system の末尾へデータ境界規律を必ず差し込む(AsmGuarded)。
    Dim sysText As String, userText As String
    sysText = modPromptsOps.AsmGuarded(modPromptsS5.BuildS5System())
    userText = AsmS5User(ctx.company, s1Json, s2Json, s3Json, _
                         modExportProposal.StatsText(s2Json, s3Json))
    If LenB(userText) = 0 Then Exit Function

    modGatewayRPN.SetRunContext caseId, CStr(roundNo), ctx.case_type, _
                                vbNullString, _
                                Trim$(modConfig.GetStr("operator", vbNullString))

    Dim okJson As String, failRaw As String, resultText As String
    resultText = CallGuarded(caseId, ctx.case_type, sysText, userText, s2Json, okJson, failRaw)
    modGatewayRPN.ClearRunContext

    If resultText = P5_RES_CALL Then Exit Function
    If resultText = P5_RES_FAILED Then
        ' 16章 E-06 と同じ扱い: 生応答は s5_json を汚さず s5_json_failed へ。
        ' **案件の status は error にしない**(S4 までの成果物は健在で、営業は
        ' 骨子(S4)で商談に行ける。16章 E-71)。
        modCaseStore.SaveData caseId, "s5_json_failed", failRaw
        modLog.LogError modPipeline.FailCodeOf(LastErrs()), P5_SRC & "." & P5_STEP, _
                        "validate_failed:" & P5_STEP & modRibbonWire.CutNote(failRaw)
        Exit Function
    End If

    If Not modCaseStore.SaveData(caseId, "s5_json", okJson) Then
        modLog.LogError "E0604", P5_SRC & ".RunStep5", "save_failed:s5"
        Exit Function
    End If
    modLog.LogUsage "s5_done", caseId, "chars=" & CStr(Len(okJson))
    RunStep5 = True
    Exit Function

Failed:
    modLog.LogError "E0603", P5_SRC & ".RunStep5", "s5", Err.Number
    RunStep5 = False
End Function

' ============================================================================
' AsmS5User - 15章§5.6 user の組立層(14章§6の二層分離)。本文は1文字も持たない。
' ============================================================================
Public Function AsmS5User(ByVal company As String, ByVal s1Json As String, _
                          ByVal s2Json As String, ByVal s3Json As String, _
                          ByVal statsText As String) As String
    Dim names() As String
    Dim vals(0 To 4) As String

    names = Split(P5_PH, P5_PH_SEP)
    vals(0) = company
    vals(1) = s1Json
    vals(2) = s2Json
    vals(3) = s3Json
    vals(4) = statsText
    AsmS5User = modPromptsOps.Fill(modPromptsS5.BuildS5User(), names, vals)
End Function

' ============================================================================
' 呼び出しと防衛線(docs/29 §5.1「1回＋修復1回」)
' ============================================================================

Public Function LastErrs() As String
    LastErrs = mLastErrs
End Function

Private Function CallGuarded(ByVal caseId As String, ByVal caseType As String, _
                             ByVal sysText As String, ByVal userText As String, _
                             ByVal s2Json As String, ByRef okJson As String, _
                             ByRef failRaw As String) As String
    Dim errText As String, rawText As String, callOk As Boolean

    mLastErrs = vbNullString
    errText = OneCall(caseId, caseType, sysText, userText, s2Json, _
                      vbNullString, False, okJson, rawText, callOk)
    If Not callOk Then
        CallGuarded = P5_RES_CALL
        Exit Function
    End If
    If LenB(errText) = 0 Then
        CallGuarded = P5_RES_OK
        Exit Function
    End If

    mLastErrs = errText
    failRaw = rawText
    If Not modPipeline.NeedsRepair(errText, RepairBudget()) Then
        CallGuarded = SoftenOrFail(caseId, errText, rawText, s2Json, okJson)
        Exit Function
    End If

    ' 15章§7: 元の user の末尾へ修復サフィックスを足して1回だけ再呼び出しする。
    errText = OneCall(caseId, caseType, sysText, _
                      userText & FillOne(modPromptsOps.RepairSuffix(), "validationErrors", errText), _
                      s2Json, "repair=1", True, okJson, rawText, callOk)
    If Not callOk Then
        CallGuarded = P5_RES_CALL
        Exit Function
    End If
    If LenB(errText) = 0 Then
        CallGuarded = P5_RES_REPAIRED
        Exit Function
    End If

    mLastErrs = errText
    failRaw = rawText
    CallGuarded = SoftenOrFail(caseId, errText, rawText, s2Json, okJson)
End Function

' ============================================================================
' SoftenOrFail - docs/29 §5.3 の最後の砦。
'   **残った不合格が V-S5-12(禁止語)だけ**なら、対訳表で機械置換してから
'   もう一度 CheckS5 に掛け、通れば合格として続行する(顧客向け提案書の生成を
'   語1つで止めない)。置換したことは run_log に taboo_softened=n で残す
'   (黙って直さない)。それ以外の不合格が混じっていれば従来どおり失敗させる。
'
'   **CheckS5 の再検証に通らなければ必ず失敗**(裁定書40 S-M4。16章 E-71③
'   「機械置換してから再検証し、**通れば**成功として続行する」の逐語)。
'   W15 Round2 の Fix 波で、ここに「残りが一般語だけなら再検証が不合格でも
'   通す」という未申告の緩和が入っていた。検証に不合格の S5 が s5_json として
'   保存され顧客向け提案書の材料になるため、**戻した**。
'   一般語(移転・保有・抜け)やサ変語幹(付保・保険化ほか)が本文に残る場合、
'   modValidate4 はそれらを機械置換しないので CheckS5 の V-S5-12 が残り、
'   S5 は失敗する(利用者には16章 E-71 の①②=案件は壊さず提案骨子を案内)。
'   顧客向け提案書の本文に社内語が残ったときに**警告だけで続ける**のは、
'   S2 由来の自由文を扱う出力側の経路(modExportProposal.TabooLeftNote)であり、
'   検証の戻り値を通す経路ではない(裁定書40 §0 の警告チャネル規約)。
' ============================================================================
Private Function SoftenOrFail(ByVal caseId As String, ByVal errText As String, _
                              ByVal rawText As String, ByVal s2Json As String, _
                              ByRef okJson As String) As String
    Dim softened As String, changed As Long

    SoftenOrFail = P5_RES_FAILED
    softened = SoftenedOrEmpty(errText, rawText, s2Json, changed)
    If LenB(softened) = 0 Then Exit Function

    okJson = softened
    mLastErrs = vbNullString
    modLog.LogUsage "s5_taboo_softened", caseId, "taboo_softened=" & CStr(changed)
    SoftenOrFail = P5_RES_REPAIRED
End Function

' ============================================================================
' SoftenedOrEmpty - 上の**判断だけ**を取り出した純関数(層(a)から直接叩ける)。
'   ""=採用しない(=S5 は失敗) / 非空=採用してよい S5 JSON。changed は置換箇所数。
'   採用の条件は3つで、1つでも欠けたら採用しない(裁定書40 S-M4):
'     (1) 残った不合格が V-S5-12 だけであること
'     (2) 機械置換が実際に1件以上効いていること
'     (3) **置換後の CheckS5 が1件も発火しないこと**(16章 E-71③「通れば」)
'   (3) を緩めると、検証に不合格の S5 が s5_json として保存され顧客向け
'   提案書の材料になる。前波の未申告の緩和はここだった。判断を切り出したのは
'   層(b)の呼び口(modGatewayRPN)無しで両方向を回帰で押さえるためである。
' ============================================================================
Public Function SoftenedOrEmpty(ByVal errText As String, ByVal rawText As String, _
                                ByVal s2Json As String, ByRef changed As Long) As String
    Dim softened As String

    changed = 0
    If Not OnlyTabooLeft(errText) Then Exit Function

    softened = modValidate4.SoftenTaboo(modJsonLite.ExtractJsonBlock(rawText), changed)
    ' 1件も置換していないなら再検証しても同じ結果にしかならない(無駄な
    ' CheckS5 を回さない)。前波で消えていた早期 Exit を戻した。
    If changed = 0 Then Exit Function
    If LenB(modValidate4.CheckS5(softened, s2Json)) > 0 Then Exit Function

    SoftenedOrEmpty = softened
End Function

' 残ったエラー行が V-S5-12 だけか(1行でも他のケースIDがあれば False)。
Public Function OnlyTabooLeft(ByVal errText As String) As Boolean
    Dim rows() As String
    Dim i As Long
    Dim t As String
    Dim n As Long

    If LenB(Trim$(errText)) = 0 Then Exit Function
    rows = Split(errText, vbLf)
    For i = LBound(rows) To UBound(rows)
        t = Trim$(rows(i))
        If LenB(t) > 0 Then
            n = n + 1
            If InStr(1, t, "[V-S5-12]", vbBinaryCompare) <> 1 Then Exit Function
        End If
    Next i
    OnlyTabooLeft = (n > 0)
End Function

' OneCall - 1回の CallStep と防衛線。run_log を必ず1行書く。
'   戻り値=検証エラー文(""=合格)。callOk=False は経路そのものの失敗。
Private Function OneCall(ByVal caseId As String, ByVal caseType As String, _
                         ByVal sysText As String, ByVal userText As String, _
                         ByVal s2Json As String, ByVal detailPre As String, _
                         ByVal isRepair As Boolean, ByRef outJson As String, _
                         ByRef outRaw As String, ByRef callOk As Boolean) As String
    Dim latencyMs As Long, detailAcc As String, errText As String
    Dim extracted As String, removedCount As Long

    detailAcc = detailPre
    outRaw = modGatewayRPN.CallStep(P5_STEP, modPipeline.PlayIdOf(caseType), _
                                    sysText, userText, modSchemas2.SchemaS5(), _
                                    callOk, latencyMs)
    If Not callOk Then
        RecordRun vbNullString, detailAcc
        Exit Function
    End If

    extracted = modJsonLite.ExtractJsonBlock(outRaw)
    If LenB(extracted) = 0 Then
        errText = modPipeline3.MsgE0302()
    Else
        outJson = modValidate.NormalizeLlmJson(P5_STEP, extracted, removedCount)
        If removedCount > 0 Then AddNote detailAcc, "e49_removed=" & CStr(removedCount)
        errText = modValidate4.CheckS5(outJson, s2Json)
    End If
    ' 裁定書47 追補(2): modPipeline/modPipeline2 と同じ VerrNoteOf へ揃える。
    If LenB(errText) > 0 Then AddNote detailAcc, modPipeline3.VerrNoteOf(errText)
    RecordRun modPipeline.ClassifyResult((LenB(errText) = 0), isRepair, (LenB(errText) = 0)), _
              detailAcc
    OneCall = errText
End Function

' run_log を1行書く(gatewayの保留行へ validate_result と detail を足す)。
Private Sub RecordRun(ByVal validateResult As String, ByVal detailText As String)
    Dim rec As TRunLogRec

    If Not modGatewayRPN.TakeLastRun(rec) Then Exit Sub
    rec.validate_result = validateResult
    If LenB(detailText) > 0 Then
        If LenB(rec.detail) > 0 Then rec.detail = rec.detail & ";"
        rec.detail = rec.detail & detailText
    End If
    modLog.LogRun rec
End Sub

Private Sub AddNote(ByRef acc As String, ByVal noteText As String)
    If LenB(noteText) = 0 Then Exit Sub
    If LenB(acc) > 0 Then acc = acc & ";"
    acc = acc & noteText
End Sub

Private Function FillOne(ByVal tpl As String, ByVal nameText As String, _
                         ByVal valText As String) As String
    Dim names(0 To 0) As String
    Dim vals(0 To 0) As String

    names(0) = nameText
    vals(0) = valText
    FillOne = modPromptsOps.Fill(tpl, names, vals)
End Function

' 修復リトライの回数(config json_repair_retry。既定1)。
Private Function RepairBudget() As Long
    RepairBudget = modConfig.GetLong("json_repair_retry", P5_REPAIR_DFLT)
End Function
