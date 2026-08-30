Attribute VB_Name = "modPipeline2"
Option Explicit

' ============================================================================
' modPipeline2 - 入念モード(deep)の批判・改訂パイプ(app層・T-28。裁定書8 A-1)
' ----------------------------------------------------------------------------
' なぜ2本に分かれているのか:
'   modPipeline(S1からS4の実行制御)が 30,000字契約(12章§2)の警告帯へ達し、
'   15章§4.5-4.7 の批判(S2C/S3C)・改訂(ReviseSuffix)の呼び出しを足す余地が
'   無くなった。modValidate/modValidate2、modCompanyFile/modCompanyFile2 と
'   同じく【分割の継ぎ目】であり、呼んでよいのは modPipeline だけ(制御の依存は
'   modPipeline -> modPipeline2 の一方向で、本モジュールが modPipeline の
'   実行制御を呼び返すことはしない)。
'
' 正: 14章§6(RunDeep の宣言)/15章§4.5(S2C)・§4.6(S3C)・§4.7(改訂パス)/
'   16章 E-35(批判不合格=批判をスキップして生成版を確定)・E-36(改訂不合格=
'   改訂を破棄して改訂前を採用)/13章§2.2(保存 data_key と参照優先)・§2.4
'   (run_log の step は s2c/s3c/s2r/s3r)/14章§5(防衛線の順)。
'
' 責務: 生成版が確定した直後の 批判 -> (指摘があれば)改訂 を回し、s2c/s3c と
'   s2r/s3r を modCaseStore 経由で保存して run_log へ1呼び出し1行を記録する。
'   **本体Stepの成否は動かさない**(E-35/E-36 とも failed_step を立てず status も
'   触らない)。
' R4(12章§4): シートに直接触れない。案件データは modCaseStore / modCaseRead、
'   ナレッジは modKnowledge が唯一の口。
'
' 【改訂プロンプトの組み直しについて】15章§4.7 は改訂を「元のsystemのまま、元の
'   userプロンプト末尾へ追記して再生成する」と定める。RunDeep の引数は14章§6が
'   (caseId, stepNo) と定めており modPipeline が組んだ本文は渡ってこないため、
'   **同じ単一の値源から同じ手順で組み直す**(ctx=modCaseRead / JSON=
'   modCaseStore / ナレッジ=modKnowledge / 本文=modPromptsOps)。規約そのもの
'   (予算配分・行数の数え方・Step別のナレッジ枠・S1要約の射影・playの決定・
'   修復要否・validate_result の3値)は modPipeline が14章§6で公開した【状態を
'   持たない判定核】をそのまま呼ぶ。答えを2箇所に書かないためであり、制御の
'   向きは一方向のまま(判定核はシートもログもLLMも触らない)。
' ============================================================================

Private Const P2_SRC As String = "modPipeline2"

' 15章§0.7 の予算配分(ナレッジ=上限の3割)。
Private Const P2_PCT_KB As Long = 3

' 13章§2.3 の既定値(NFR-M3。modPipeline と同じキーを見る)。
Private Const P2_MAX_CTX_DFLT As Long = 40000
Private Const P2_T2_MAX_CTX_DFLT As Long = 100000
Private Const P2_REPAIR_DFLT As Long = 1
Private Const P2_TIER_T1 As String = "t1_quick"

' ナレッジ0行・入力なしの既定文言(modPipeline と同じ扱い)。
Private Const P2_KB_ZERO As String = "なし"
Private Const P2_NONE_TEXT As String = "なし"

' パイプの結末(13章§2.4 detail・16章 E-35/E-36)。
Private Const P2_OUT_CRITIQUE_SKIPPED As String = "critique_skipped"
Private Const P2_OUT_REVISION_SKIPPED As String = "revision_skipped"
Private Const P2_OUT_REVISED As String = "revised"
Private Const P2_OUT_REVISION_DISCARDED As String = "revision_discarded"

' CallGuarded の戻り値(modPipeline と同じ3値+経路失敗)。
Private Const P2_RES_OK As String = "ok"
Private Const P2_RES_FAILED As String = "failed"
Private Const P2_RES_CALL As String = "call_failed"

' 1パイプ分の文脈。Check系は純関数でJSONの外側の文脈を引数で受ける(14章§6)
' ため、注入テキストと上流JSONをここへまとめて持ち回る。
Private Type TDeepCtx
    caseId As String
    stepNo As Long
    stepName As String
    roundText As String
    baseJson As String
    s1Json As String
    s2Json As String
    prevS2Json As String
    menusText As String
    linesText As String
    schemesText As String
    casesText As String
    riskLibText As String
End Type

' ============================================================================
' RunDeep - 入念モードの批判・改訂パイプの入口(14章§6)
' ----------------------------------------------------------------------------
'   modPipeline は sN_json を確定した直後、DeepEnabled が True のときだけ
'   1行で委譲する。stepNo は 2 または 3(それ以外は False で何もしない)。
'   戻り値 True = 15章§4.5-4.7 のパイプを規約どおり最後まで回した(改訂の採否・
'     E-35 のスキップ・E-36 の破棄を含む。どれも設計どおりの終わり方)。
'   戻り値 False = 対象Stepでない / 生成版や案件文脈が読めない / 経路そのものが
'     失敗した(16章 E-14からE-16)。
'   **戻り値で本体Stepの成否を左右しない**のが契約。呼び出し側は握りつぶしてよい。
' ============================================================================
Public Function RunDeep(ByVal caseId As String, ByVal stepNo As Long) As Boolean
    On Error GoTo Failed

    If stepNo <> 2 And stepNo <> 3 Then Exit Function

    Dim d As TDeepCtx
    d.caseId = caseId
    d.stepNo = stepNo
    d.stepName = "s" & CStr(stepNo)

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String, s4Variant As String, tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, _
                                   tierText) Then
        modLog.LogUsage "deep_ctx_unavailable", caseId, "案件一覧を読めないため入念モード中止"
        Exit Function
    End If
    If LenB(ctx.dossier_tier) = 0 Then ctx.dossier_tier = tierText
    If roundNo > 0 Then d.roundText = CStr(roundNo)

    ' 批判の審査対象は【今確定した生成版】そのもの(sN_json)。ResolveStepJson を
    ' 使うと前ラウンドの改訂版(sNr_json)を拾いうるので直接読む。
    d.baseJson = modCaseStore.LoadData(caseId, d.stepName & "_json")
    If LenB(Trim$(d.baseJson)) = 0 Then
        modLog.LogUsage "deep_base_missing", caseId, d.stepName & "_json が無いため中止"
        Exit Function
    End If

    ' 前ラウンドの批判・改訂を残さない(13章§2.2の参照優先で古い改訂が下流Step
    ' に拾われるのを防ぐ)。
    modCaseStore.SaveData caseId, d.stepName & "c_json", vbNullString
    modCaseStore.SaveData caseId, d.stepName & "r_json", vbNullString

    NoteDeepRoute caseId

    d.s1Json = modCaseStore.ResolveStepJson(caseId, 1)
    If stepNo = 3 Then
        d.s2Json = modCaseStore.ResolveStepJson(caseId, 2)
    Else
        d.s2Json = d.baseJson
        d.prevS2Json = OrNone(modCaseStore.LoadData(caseId, "s2_prev_json"))
    End If
    LoadKb ctx, d

    RunDeep = RunPipe(ctx, d)
    Exit Function

Failed:
    modLog.LogError "E0603", P2_SRC & ".RunDeep", "step=" & CStr(stepNo), Err.Number
    RunDeep = False
End Function

' RunPipe - 批判 -> (指摘があれば)改訂 の本体。結末の分類は DeepOutcomeOf が
'   唯一の判定点で、HOMEへ出す警告文は DeepWarningOf が唯一の値源。
Private Function RunPipe(ByRef ctx As TCaseCtx, ByRef d As TDeepCtx) As Boolean
    Dim critiqueJson As String
    Dim revisedJson As String
    Dim resultText As String

    ' --- 批判(15章§4.5 / §4.6) ---
    resultText = RunCritique(ctx, d, critiqueJson)
    If resultText <> P2_RES_OK Then
        ' E-35: 批判をスキップして生成版を確定。本体Stepは成功のまま。経路その
        ' ものが落ちた場合(E-14からE-16・E-15の上限を含む)も生成版の確定は同じで、
        ' 戻り値だけを False にして「回し切っていない」ことを呼び出し側へ伝える。
        RecordOutcome d, DeepOutcomeOf(False, False, False)
        RunPipe = (resultText = P2_RES_FAILED)
        Exit Function
    End If

    ' V-S2C-05 / V-S3C-05: 指摘が無ければ改訂パスをスキップし呼び出しを節約する。
    If Not NeedsRevision(critiqueJson, d.stepNo) Then
        RecordOutcome d, DeepOutcomeOf(True, False, False)
        RunPipe = True
        Exit Function
    End If

    ' --- 改訂(15章§4.7) ---
    resultText = RunRevision(ctx, d, critiqueJson, revisedJson)
    If resultText = P2_RES_OK Then
        ' 13章§2.2: 改訂版は sN_json を上書きせず sNr_json へ入れる。下流Stepは
        ' sN_edited > sNr_json > sN_json の優先で参照する。
        If Not modCaseStore.SaveData(d.caseId, d.stepName & "r_json", revisedJson) Then
            modLog.LogError "E0604", P2_SRC & ".RunPipe", "save_failed:" & d.stepName
        End If
    End If
    RecordOutcome d, DeepOutcomeOf(True, True, (resultText = P2_RES_OK))
    RunPipe = (resultText <> P2_RES_CALL)
End Function

' RunCritique - S2C / S3C の1往復。合格した批判JSONを critiqueJson へ返す。
'   戻り値は CallGuarded と同じ ok / failed / call_failed。
'   生の批判応答は合否によらず sNc_json へ残す(16章 E-35)。
Private Function RunCritique(ByRef ctx As TCaseCtx, ByRef d As TDeepCtx, _
                             ByRef critiqueJson As String) As String
    Dim sysText As String, userText As String, schemaText As String
    Dim rawText As String

    If d.stepNo = 2 Then
        sysText = modPromptsOps.BuildS2CriticSystem()
        userText = modPromptsOps.AsmS2CriticUser(d.s1Json, d.baseJson, d.riskLibText)
        schemaText = modSchemas.SchemaS2C()
    Else
        sysText = modPromptsOps.BuildS3CriticSystem()
        userText = modPromptsOps.AsmS3CriticUser(ctx, modPipeline.S1SummaryOf(d.s1Json), _
                                                 d.s2Json, d.baseJson)
        schemaText = modSchemas.SchemaS3C()
    End If

    Dim resultText As String
    resultText = CallGuarded(ctx, d, CritiqueStepOf(d.stepNo), sysText, userText, _
                             schemaText, critiqueJson, rawText)

    ' 生出力を残す(合否によらず。E-35)。経路失敗で空のときは書かない。
    If LenB(Trim$(rawText)) > 0 Then
        modCaseStore.SaveData d.caseId, d.stepName & "c_json", rawText
    End If
    RunCritique = resultText
End Function

' RunRevision - 改訂の1往復。**元のsystemのまま、元のuser末尾へ ReviseSuffix を
'   足して再生成**する(15章§4.7)。検証は改訂前と同じ本体スキーマ(CheckS2 /
'   CheckS3)。合格した改訂版を revisedJson へ返す。戻り値は CallGuarded と同じ
'   ok / failed / call_failed。
'   不合格(E-36)は改訂版を捨て、生応答を sN_json_failed へ退避する(sNr_json
'   には本体スキーマに合格した改訂版だけを入れる)。
Private Function RunRevision(ByRef ctx As TCaseCtx, ByRef d As TDeepCtx, _
                             ByVal critiqueJson As String, _
                             ByRef revisedJson As String) As String
    Dim sysText As String, userText As String, schemaText As String
    Dim rawText As String

    If d.stepNo = 2 Then
        sysText = modPromptsCore.BuildS2System()
        userText = modPromptsOps.AsmS2User(ctx, d.s1Json, d.riskLibText, d.menusText, _
                                           d.prevS2Json, HearingOf(d.caseId))
        schemaText = modSchemas.SchemaS2()
    Else
        sysText = modPromptsCore.BuildS3System()
        userText = modPromptsOps.AsmS3User(ctx, modPipeline.S1SummaryOf(d.s1Json), _
                                           d.s2Json, d.menusText, d.linesText, _
                                           d.schemesText, d.casesText)
        schemaText = modSchemas.SchemaS3()
    End If
    userText = userText & FillOne(modPromptsOps.ReviseSuffix(), "critiqueDigest", _
                                  CritiqueDigest(critiqueJson, d.stepNo))

    Dim resultText As String
    resultText = CallGuarded(ctx, d, ReviseStepOf(d.stepNo), sysText, userText, _
                             schemaText, revisedJson, rawText)
    RunRevision = resultText
    If resultText = P2_RES_OK Then Exit Function

    ' E-36: 改訂を破棄し改訂前を採用する。生応答は sN_json を汚さず退避する。
    revisedJson = vbNullString
    If resultText = P2_RES_FAILED Then
        If LenB(Trim$(rawText)) > 0 Then
            modCaseStore.SaveData d.caseId, d.stepName & "_json_failed", rawText
        End If
    End If
End Function

' RecordOutcome - パイプの結末を記録する。**failed_step も status も動かさない**
'   (E-35/E-36: 本体Stepは成功のまま)。HOMEの hm_warning は ui層 が
'   DeepWarningOf から取る(app層から ui層 は呼べない=R1)。
Private Sub RecordOutcome(ByRef d As TDeepCtx, ByVal outcome As String)
    Dim detailText As String
    Dim warnText As String

    detailText = d.stepName
    warnText = DeepWarningOf(outcome)
    If LenB(warnText) > 0 Then detailText = detailText & " " & warnText
    modLog.LogUsage "deep_" & outcome, d.caseId, detailText
End Sub

' ============================================================================
' 純関数(シート・ログ・LLMに触れない判定核)。14章§6が公開を宣言する
' ============================================================================

' CritiqueStepOf - 批判呼び出しの step 値(19章§4)。2->s2c / 3->s3c / 他は ""。
Public Function CritiqueStepOf(ByVal stepNo As Long) As String
    If stepNo <> 2 And stepNo <> 3 Then Exit Function
    CritiqueStepOf = "s" & CStr(stepNo) & "c"
End Function

' ReviseStepOf - 改訂呼び出しの step 値(19章§4)。2->s2r / 3->s3r / 他は ""。
Public Function ReviseStepOf(ByVal stepNo As Long) As String
    If stepNo <> 2 And stepNo <> 3 Then Exit Function
    ReviseStepOf = "s" & CStr(stepNo) & "r"
End Function

' NeedsRevision - 改訂パスへ進むか(15章§4.5 V-S2C-05 / §4.6 V-S3C-05)。
'   共通: issues が1件でもあれば進む。
'   S2C : 加えて additional_risks(見落としの指摘)が1件でもあれば進む。
'   S3C : 加えて lands=false の反応が1件でもあれば進む(V-S3C-05 のスキップ条件
'         が「lands が3件とも true **かつ** issues 0件」だから)。
'   範囲外の stepNo・空JSONは False(無駄な呼び出しを増やさない)。
Public Function NeedsRevision(ByVal critiqueJson As String, ByVal stepNo As Long) As Boolean
    If stepNo <> 2 And stepNo <> 3 Then Exit Function
    If LenB(Trim$(critiqueJson)) = 0 Then Exit Function

    If modJsonLite.GetArrayItems(critiqueJson, "issues").count > 0 Then
        NeedsRevision = True
        Exit Function
    End If
    If stepNo = 2 Then
        NeedsRevision = (modJsonLite.GetArrayItems(critiqueJson, "additional_risks").count > 0)
        Exit Function
    End If

    Dim it As Variant
    For Each it In modJsonLite.GetArrayItems(critiqueJson, "executive_reactions")
        If Not modJsonLite.GetBoolJ(CStr(it), "lands", True) Then
            NeedsRevision = True
            Exit Function
        End If
    Next it
End Function

' CritiqueDigest - 15章§4.7 の {{critiqueDigest}}。批判JSONを日本語の箇条書きへ
'   整形した文字列(行区切りは vbLf)。**キー名をそのまま出さず**審査結果として
'   読める文にする。空・範囲外は ""。
'   S2: issues -> additional_risks / S3: lands=false の反応 -> issues。
Public Function CritiqueDigest(ByVal critiqueJson As String, ByVal stepNo As Long) As String
    Dim acc As String
    Dim it As Variant

    If stepNo <> 2 And stepNo <> 3 Then Exit Function
    If LenB(Trim$(critiqueJson)) = 0 Then Exit Function

    If stepNo = 3 Then
        For Each it In modJsonLite.GetArrayItems(critiqueJson, "executive_reactions")
            If Not modJsonLite.GetBoolJ(CStr(it), "lands", True) Then
                acc = AddLine(acc, "・経営者に刺さらないと判定 story_no:" & _
                    CStr(modJsonLite.GetLong(CStr(it), "story_no", 0)) & ": " & _
                    modJsonLite.GetStr(CStr(it), "reaction"))
            End If
        Next it
    End If

    For Each it In modJsonLite.GetArrayItems(critiqueJson, "issues")
        acc = AddLine(acc, "・[" & modJsonLite.GetStr(CStr(it), "target") & "] " & _
            modJsonLite.GetStr(CStr(it), "issue_type") & ": " & _
            modJsonLite.GetStr(CStr(it), "detail") & " / 改善の方向: " & _
            modJsonLite.GetStr(CStr(it), "suggestion"))
    Next it

    If stepNo = 2 Then
        For Each it In modJsonLite.GetArrayItems(critiqueJson, "additional_risks")
            acc = AddLine(acc, "・見落としの指摘: " & _
                modJsonLite.GetStr(CStr(it), "risk_name") & " / 根拠: " & _
                modJsonLite.GetStr(CStr(it), "why"))
        Next it
    End If
    CritiqueDigest = acc
End Function

' DeepOutcomeOf - パイプの結末の唯一の分類点(16章 E-35/E-36)。
'   critique_skipped   : 批判が修復1回で直らずスキップ(E-35。生成版を確定)
'   revision_skipped   : 批判に指摘が無く改訂を省略(V-S2C-05 / V-S3C-05)
'   revised            : 改訂版が本体スキーマに合格し sNr_json へ入った
'   revision_discarded : 改訂版が不合格で破棄し改訂前を採用(E-36)
'   4値のいずれでも**本体Stepは成功のまま**で failed_step は立てない。
Public Function DeepOutcomeOf(ByVal critiqueOk As Boolean, ByVal revisionTried As Boolean, _
                              ByVal revisionOk As Boolean) As String
    If Not critiqueOk Then
        DeepOutcomeOf = P2_OUT_CRITIQUE_SKIPPED
        Exit Function
    End If
    If Not revisionTried Then
        DeepOutcomeOf = P2_OUT_REVISION_SKIPPED
        Exit Function
    End If
    If revisionOk Then
        DeepOutcomeOf = P2_OUT_REVISED
    Else
        DeepOutcomeOf = P2_OUT_REVISION_DISCARDED
    End If
End Function

' DeepWarningOf - HOMEの hm_warning へ出す文言(16章 E-35/E-36 の逐語)。
'   警告の要らない結末は ""。ui層(modUIHome)はこの1本だけを見る(同じ文言を
'   2箇所に書かない。app層から ui層 は呼べないので値だけを供給する)。
Public Function DeepWarningOf(ByVal outcome As String) As String
    Select Case Trim$(outcome)
        Case P2_OUT_CRITIQUE_SKIPPED
            DeepWarningOf = "入念モードの審査を省略しました（生成版で続行）"
        Case P2_OUT_REVISION_DISCARDED
            DeepWarningOf = "入念モードの改訂を破棄しました（改訂前で続行）"
    End Select
End Function

' DeepRouteOf - config `deep_transport`(13章§2.3)の解決。空=通常経路のまま。
'   "direct" 指定のときだけ批判・改訂の呼び出しを direct へ寄せる意思表示で、
'   想定外の値は空(通常経路)へ倒す(設定の打ち間違いで経路を壊さない)。
Public Function DeepRouteOf(ByVal cfgDeepTransport As String) As String
    If LCase$(Trim$(cfgDeepTransport)) = "direct" Then DeepRouteOf = "direct"
End Function

' vbLf 区切りで1行積む(空行は作らない)。
Private Function AddLine(ByVal acc As String, ByVal lineText As String) As String
    If LenB(Trim$(lineText)) = 0 Then
        AddLine = acc
        Exit Function
    End If
    If LenB(acc) = 0 Then
        AddLine = lineText
    Else
        AddLine = acc & vbLf & lineText
    End If
End Function

' ============================================================================
' 呼び出しと防衛線(14章§5)。modPipeline と同じ順序を保つ
' ============================================================================

' CallGuarded - 1呼び出し分の防衛線。修復リトライは1回まで(config
'   json_repair_retry)。戻り値: ok / failed / call_failed。合格したJSONは
'   okJson、最後に受け取った生応答は rawOut へ返す。
Private Function CallGuarded(ByRef ctx As TCaseCtx, ByRef d As TDeepCtx, _
                             ByVal stepKey As String, ByVal sysText As String, _
                             ByVal userText As String, ByVal schemaText As String, _
                             ByRef okJson As String, ByRef rawOut As String) As String
    Dim errText As String
    Dim callOk As Boolean

    errText = OneCall(ctx, d, stepKey, sysText, userText, schemaText, _
                      vbNullString, False, okJson, rawOut, callOk)
    If Not callOk Then
        CallGuarded = P2_RES_CALL
        Exit Function
    End If
    If LenB(errText) = 0 Then
        CallGuarded = P2_RES_OK
        Exit Function
    End If

    If modPipeline.NeedsRepair(errText, RepairBudget()) Then
        errText = OneCall(ctx, d, stepKey, sysText, _
                          userText & FillOne(modPromptsOps.RepairSuffix(), _
                                             "validationErrors", errText), _
                          schemaText, "repair=1", True, okJson, rawOut, callOk)
        If Not callOk Then
            CallGuarded = P2_RES_CALL
            Exit Function
        End If
        If LenB(errText) = 0 Then
            CallGuarded = P2_RES_OK
            Exit Function
        End If
    End If

    ' 不合格。E-35/E-36 の分岐は呼び出し側が行うので、ここではコードだけ残す。
    modLog.LogError "E0302", P2_SRC & "." & stepKey, "validate_failed:" & stepKey
    CallGuarded = P2_RES_FAILED
End Function

' OneCall - 1回の CallStep と防衛線(1)(2.5)(3)。run_log を必ず1行書く
'   (13章§2.4 の step は s2c/s3c/s2r/s3r)。戻り値=検証エラー文(""=合格)。
'   callOk=False は経路そのものの失敗(16章 E-14からE-16)。
Private Function OneCall(ByRef ctx As TCaseCtx, ByRef d As TDeepCtx, _
                         ByVal stepKey As String, ByVal sysText As String, _
                         ByVal userText As String, ByVal schemaText As String, _
                         ByVal detailPre As String, ByVal isRepair As Boolean, _
                         ByRef outJson As String, ByRef outRaw As String, _
                         ByRef callOk As Boolean) As String
    Dim latencyMs As Long
    Dim detailAcc As String
    Dim errText As String

    detailAcc = detailPre
    modGatewayRPN.SetRunContext d.caseId, d.roundText, ctx.case_type, _
                                modKnowledge.LastInjectedIds(), vbNullString
    outRaw = modGatewayRPN.CallStep(stepKey, modPipeline.PlayIdOf(ctx.case_type), _
                                    sysText, userText, schemaText, callOk, latencyMs)
    modGatewayRPN.ClearRunContext
    If Not callOk Then
        ' 検証へ到達していないので validate_result は空のまま記録する。
        RecordRun vbNullString, detailAcc
        Exit Function
    End If

    errText = Defend(ctx, d, stepKey, outRaw, outJson, detailAcc)
    If LenB(errText) > 0 Then AddNote detailAcc, "verr=E0302"
    RecordRun modPipeline.ClassifyResult((LenB(errText) = 0), isRepair, _
                                         (LenB(errText) = 0)), detailAcc
    OneCall = errText
End Function

' 防衛線(1)抽出 -> (2.5)重複排除 -> (3)検証。批判は CheckS2C / CheckS3C、
'   改訂は**改訂前と同じ本体スキーマ**(CheckS2 / CheckS3。15章§4.7)。
Private Function Defend(ByRef ctx As TCaseCtx, ByRef d As TDeepCtx, _
                        ByVal stepKey As String, ByVal rawText As String, _
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

    outJson = modValidate.NormalizeLlmJson(stepKey, extracted, removedCount)
    If removedCount > 0 Then
        ' E-49: 黙殺しない。件数を detail に残し E0303 を警告として記録。
        AddNote detailAcc, "e49_removed=" & CStr(removedCount)
        modLog.LogError "E0303", P2_SRC & "." & stepKey, "e49_removed=" & CStr(removedCount)
    End If

    Select Case stepKey
        Case "s2c"
            ' V-S2C-03 は「指摘が指す risk_no / gap_no が審査対象のS2に在るか」。
            Defend = modValidate.CheckS2C(outJson, d.baseJson)
        Case "s3c"
            Defend = modValidate.CheckS3C(outJson)
        Case "s2r"
            Defend = modValidate.CheckS2(outJson, ctx.case_type, d.menusText, d.prevS2Json)
        Case "s3r"
            Defend = modValidate.CheckS3(outJson, d.s2Json, d.menusText, d.linesText, _
                                         d.schemesText, d.casesText, ctx.case_type)
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

' run_log detail の積み方(400字上限は modLog.TruncDetail)。
Private Sub AddNote(ByRef acc As String, ByVal noteText As String)
    If LenB(noteText) = 0 Then Exit Sub
    If LenB(acc) > 0 Then acc = acc & ";"
    acc = acc & noteText
End Sub

' ============================================================================
' 入力の取り直し(単一の値源から同じ手順で組み直す。冒頭の注記を参照)
' ============================================================================

' LoadKb - ナレッジ注入と15章§0.7の切詰め。行数の数え方(KbRowCount)・Step別の
'   枠(UsesSlot)・予算配分(BudgetOf)は modPipeline が公開した判定核、切詰めの
'   計画は modKnowledgeFmt.TrimPlan が唯一の実装(答えを2箇所に書かない)。
Private Sub LoadKb(ByRef ctx As TCaseCtx, ByRef d As TDeepCtx)
    Dim txt(0 To 4) As String
    Dim counts(0 To 9) As Long
    Dim plan As Variant
    Dim i As Long
    Dim trimmed As Boolean

    modKnowledge.ResetInjectedIds
    FetchKb ctx, d.stepNo, txt, -1, 0
    For i = 0 To 4
        counts(i) = modPipeline.KbRowCount(txt(i))
        counts(5 + i) = Len(txt(i))
    Next i

    plan = modKnowledgeFmt.TrimPlan(counts, _
        modPipeline.BudgetOf(ContextLimitOf(ctx.dossier_tier), P2_PCT_KB))
    For i = 0 To 4
        If plan(i) < counts(i) Then trimmed = True
    Next i
    If trimmed Then
        modKnowledge.ResetInjectedIds
        For i = 0 To 4
            If counts(i) > 0 Then FetchKb ctx, d.stepNo, txt, i, plan(i)
        Next i
    End If

    d.casesText = txt(0)
    d.schemesText = txt(1)
    d.menusText = txt(2)
    d.linesText = txt(3)
    d.riskLibText = txt(4)
End Sub

' slotIdx=-1 は全スロットを既定行数(config)で、0以上はそのスロットだけ maxRows
'   行(0なら既定文言)で取り直す。メニューはStepで別物(12章§3)。
Private Sub FetchKb(ByRef ctx As TCaseCtx, ByVal stepNo As Long, ByRef txt() As String, _
                    ByVal slotIdx As Long, ByVal maxRows As Long)
    Dim i As Long, n As Long

    For i = 0 To 4
        If (slotIdx < 0 Or slotIdx = i) And modPipeline.UsesSlot(stepNo, i) Then
            n = 0
            If slotIdx >= 0 Then n = maxRows
            If slotIdx >= 0 And maxRows <= 0 Then
                txt(i) = P2_KB_ZERO
            ElseIf i = 0 Then
                txt(i) = modKnowledge.CasesFor(ctx.industry_code, n)
            ElseIf i = 1 Then
                txt(i) = modKnowledge.SchemesFor(ctx.industry_code, n)
            ElseIf i = 2 And stepNo = 2 Then
                txt(i) = modKnowledge.MenusSummaryFor(ctx.industry_code, n)
            ElseIf i = 2 Then
                txt(i) = modKnowledge.MenusFor(ctx.industry_code, n)
            ElseIf i = 3 Then
                txt(i) = modKnowledge.LinesText(n)
            Else
                txt(i) = modKnowledge.RiskLibFor(ctx.industry_code, n)
            End If
        End If
    Next i
End Sub

' 16章E-03(1): ティアで上限config(modPipeline と同じキーを見る)。
Private Function ContextLimitOf(ByVal tier As String) As Long
    If Trim$(tier) = P2_TIER_T1 Then
        ContextLimitOf = modConfig.GetLong("max_context_chars", P2_MAX_CTX_DFLT)
    Else
        ContextLimitOf = modConfig.GetLong("t2_max_context_chars", P2_T2_MAX_CTX_DFLT)
    End If
End Function

' S2 user の {{hearingAnswers}}。外部由来なので E-04 の無害化を通す。
Private Function HearingOf(ByVal caseId As String) As String
    Dim removedN As Long, markerN As Long

    HearingOf = OrNone(modUtilText.SanitizeInput( _
        modCaseStore.LoadData(caseId, "input_hearing_answers"), removedN, markerN))
End Function

Private Function OrNone(ByVal s As String) As String
    If LenB(Trim$(s)) = 0 Then
        OrNone = P2_NONE_TEXT
    Else
        OrNone = s
    End If
End Function

' テンプレートの {{name}} を1つ埋める(置換の口は modPromptsOps.Fill のみ)。
Private Function FillOne(ByVal tpl As String, ByVal nameText As String, _
                         ByVal valText As String) As String
    Dim names(0 To 0) As String, vals(0 To 0) As String

    names(0) = nameText
    vals(0) = valText
    FillOne = modPromptsOps.Fill(tpl, names, vals)
End Function

Private Function RepairBudget() As Long
    RepairBudget = modConfig.GetLong("json_repair_retry", P2_REPAIR_DFLT)
End Function

' deep_transport(13章§2.3)は「批判・改訂の呼び出しだけを direct へ寄せる」設定
'   だが、14章§6 の CallStep には**呼び出し単位で経路を上書きする口が無い**
'   (経路は modGatewayRPN が config から決める)。黙って無視すると利用者は
'   設定したつもりの経路で走り続けるため、指定がある間は毎回その事実を記録する
'   (結線には§6への経路上書き口の追加が要る=裁定事項)。
Private Sub NoteDeepRoute(ByVal caseId As String)
    Dim routeText As String

    routeText = DeepRouteOf(modConfig.GetStr("deep_transport", vbNullString))
    If LenB(routeText) = 0 Then Exit Sub
    modLog.LogUsage "deep_transport_unwired", caseId, _
                    "deep_transport=" & routeText & " は経路上書き口が無く未適用"
End Sub
