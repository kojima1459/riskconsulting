Attribute VB_Name = "modNaviActions"
Option Explicit

' err_log の出所(16章§1)。関数名をリテラルで散らさないための値源。
Private Const NA_SRC As String = "modNaviActions"

' F-16 phase 1: snapshots are taken only for sheets opened by this host.
Private gSheetCase(1 To 4) As String
Private gSheetSnapshot(1 To 4) As String

' [NAVI] Specification 7.2 / 7.3. Only this dispatcher owns long-operation locks.
Public Function Dispatch(ByVal action As String, ByVal data As String, ByRef caseId As String) As String
    On Error GoTo Failed
    Dim requested As String, ownedLock As Boolean, response As String, srcValue As String
    Dim ok As Boolean, n As Long, note As String, errCode As String
    requested = modJsonLite.GetStr(data, "case_id")
    Select Case action
    Case "new_case", "open_case", "open_step_sheet", "import_case", "company_open", "archive_case", "close"
        If Not FlushAllSheets(note) Then
            Dispatch = Failure(note, "E0302")
            Exit Function
        End If
    End Select
    If LenB(requested) > 0 And requested <> caseId Then
        If Not FlushAllSheets(note) Then
            Dispatch = Failure(note, "E0302")
            Exit Function
        End If
    End If
    If LenB(requested) > 0 Then
        If Not modNaviStore.CaseExists(requested) Then
            Dispatch = Failure("案件が見つかりません。", "E0101")
            Exit Function
        End If
        caseId = requested
    End If
    If NeedsCase(action) Then
        If Not modNaviStore.CaseExists(caseId) Then
            Dispatch = Failure("会社情報を入力し、登録内容を保存してください。", "E0101")
            Exit Function
        End If
    End If
    Select Case action
    Case "chat", "sparring_resume", "sparring_send", "start_round2", "company_save", "export_case"
        If Not FlushOwnedSheets(caseId, 5, note) Then
            Dispatch = Failure(note, "E0302")
            Exit Function
        End If
    End Select
    If IsLongAction(action) Then
        ownedLock = modUIProgress.TryEnterUiLock("HTML:" & action)
        If Not ownedLock Then
            Dispatch = "{""ok"":false,""busy"":true,""message"":""処理中です""}"
            Exit Function
        End If
    End If
    Select Case action
    Case "initialize"
        response = Success("準備できました。")
    Case "new_case"
        caseId = ""
        modNaviState.SetDraft "{}"
        response = Success("新規案件の会社情報を入力してください。")
    Case "open_case"
        If modNaviStore.HydrateCaseIfNeeded(caseId, note) Then
            response = Success("案件を開きました。")
        Else
            response = Failure(note, "E0603")
        End If
    Case "save_basics"
        response = ActSaveBasics(caseId, data)
    Case "paste_material"
        response = ActPasteMaterial(caseId, data)
    Case "save_materials"
        response = ActSaveMaterials(caseId, data)
    Case "clear_material"
        response = ActClearMaterial(caseId, data)
    Case "copy_prompt"
        response = ActCopyPrompt(caseId, data)
    Case "open_url"
        response = ActOpenUrl(data)
    Case "run_pipeline"
        response = ActRunPipeline(caseId, data)
    Case "open_step_sheet"
        n = modJsonLite.GetLong(data, "step_no", 1)
        If n < 1 Or n > 4 Then
            response = Failure("段番号が不正です。", "E0101")
        Else
            ok = modUICase2.DrawStep(caseId, n)
            If ok Then
                gSheetCase(n) = caseId
                gSheetSnapshot(n) = modUICase2.SerializeStep(n)
                ok = modUISheet.ShowSheet(modUICase2.SheetNameOf(n))
            End If
            response = ResultOf(ok, "シートで編集できます。", "シートを表示できませんでした。")
        End If
    Case "save_step_edit"
        response = PhaseTwoStepEdit(caseId, data)
    Case "export_report"
        response = ActExportReport(caseId, modJsonLite.GetStr(data, "reviewedBy"))
    Case "export_hearing"
        response = ActExportHearing(caseId, data)
    Case "report_mail"
        response = modReportMail.SendReportMail(caseId)
    Case "open_report"
        response = ActOpenReport(caseId, data)
    Case "chat"
        modNaviHost.ShowBusy "案件チャットの回答を確認しています", ProgressJson(1, 1, "案件チャット")
        DoEvents
        ok = modNaviChat.Ask(caseId, modJsonLite.GetStr(data, "question"), _
                            modNaviJson.ObjectField(data, "include"), srcValue, errCode)
        If ok Then
            response = SavedResult(caseId, "回答を保存しました。")
        Else
            note = modGatewayRPN.ErrMessageFor(errCode)
            If errCode = "E0103" Then note = "個人情報らしき記述を検知したため送信しませんでした。"
            response = Failure(note, errCode)
        End If
    Case "clear_chat"
        If Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
            response = Confirm("案件チャットの履歴をクリアします。")
        Else
            ok = modNaviChat.Clear(caseId, modJsonLite.GetLong(data, "expected_count", -1))
            If ok Then
                response = SavedResult(caseId, "履歴をクリアしました。")
            Else
                response = Failure("履歴が更新されています。最新の履歴を確認してやり直してください。", "E0604")
            End If
        End If
    Case "sparring_resume"
        n = modSparring.ResumeSparring(caseId, note)
        If n >= 0 Then
            modNaviState.SetSparringNote caseId, note
            response = SavedResult(caseId, "商談の予行演習を開始しました。調べる深さを「商談の予行演習あり」に変更しました。")
        Else
            response = Failure("商談の予行演習を開始できませんでした。", "E0101")
        End If
    Case "sparring_send"
        modNaviHost.ShowBusy "商談の予行演習の回答を確認しています", ProgressJson(1, 1, "商談の予行演習")
        DoEvents
        ok = modSparring.SendSparring(caseId, modJsonLite.GetStr(data, "utterance"), srcValue, errCode)
        If ok Then
            response = SavedResult(caseId, "商談の予行演習の回答を保存しました。")
        Else
            note = modGatewayRPN.ErrMessageFor(errCode)
            If errCode = "E0103" Then note = "個人情報らしき記述を検知したため送信しませんでした。"
            response = Failure(note, errCode)
        End If
    Case "sparring_to_inbox"
        srcValue = modJsonLite.GetStr(data, "role")
        If srcValue = "assistant" Then srcValue = "ai"
        note = modSparring.SendToInbox(caseId, srcValue, modJsonLite.GetLong(data, "seq", 0))
        response = ResultOf(LenB(note) > 0, "受信箱へ登録しました。", "受信箱へ登録できませんでした。")
    Case "start_round2", "company_save", "company_open", "feedback_add", _
         "inbox_list", "inbox_post", "inbox_diagnose_all", "inbox_judge", _
         "judge_list", "judge_add", "judge_result", "logs", "reload_kb", "run_tests", _
         "rename_case", "archive_case", "export_case", "import_case", "save_settings"
        response = modNaviActions2.DispatchMore(action, data, caseId)
    Case "resize", "close"
        ' Host owns form lifetime and performs these after successful dispatch.
        response = Success("")
    Case Else
        response = Failure("この操作は受け付けられません。", "E0101")
    End Select
    If ownedLock Then modUIProgress.ReleaseUiLock
    Dispatch = response
    Exit Function
Failed:
    note = Err.Description
    modLog.LogError "E0603", NA_SRC & ".Dispatch", caseId & ":" & action, Err.Number
    If ownedLock Then ReleaseAfterError
    Dispatch = Failure("処理を完了できませんでした。" & note, "E0603")
End Function

Private Sub ReleaseAfterError()
    On Error GoTo Done
    modUIProgress.ReleaseUiLock
Done:
End Sub

Public Function IsLongAction(ByVal action As String) As Boolean
    Select Case action
    Case "run_pipeline", "chat", "sparring_send", "inbox_diagnose_all", "run_tests"
        IsLongAction = True
    End Select
End Function

Public Function NeedsCase(ByVal action As String) As Boolean
    Select Case action
    Case "open_case", "clear_material", "run_pipeline", "open_step_sheet", "save_step_edit", _
         "export_report", "open_report", "export_hearing", "chat", "clear_chat", "report_mail", _
         "sparring_resume", "sparring_send", "sparring_to_inbox", "start_round2", _
         "company_save", "company_open", "feedback_add", "rename_case", "archive_case", "export_case"
        NeedsCase = True
    End Select
End Function

Public Function ValidSlot(ByVal slot As String) As Boolean
    ValidSlot = (LenB(slot) > 0 And InStr(1, ";" & Replace(modUICase6.AreaKeys(), vbLf, ";") & ";", _
                                                 ";" & slot & ";", 0) > 0)
End Function

Public Function Success(ByVal message As String) As String
    Success = "{""ok"":true,""message"":" & modNaviJson.Q(message) & "}"
End Function

Public Function Failure(ByVal message As String, ByVal code As String) As String
    Failure = "{""ok"":false,""message"":" & modNaviJson.Q(message) & _
              ",""error_code"":" & modNaviJson.Q(code) & "}"
End Function

Public Function Confirm(ByVal message As String) As String
    Confirm = "{""ok"":true,""confirm"":{""message"":" & modNaviJson.Q(message) & _
              "},""message"":" & modNaviJson.Q(message) & "}"
End Function

Public Function ResultOf(ByVal ok As Boolean, ByVal successText As String, ByVal failureText As String) As String
    If ok Then ResultOf = Success(successText) Else ResultOf = Failure(failureText, "E0603")
End Function

Public Function SavedResult(ByVal caseId As String, ByVal message As String) As String
    SavedResult = SavedResultWithWarning(caseId, message, vbNullString)
End Function

' SavedResultWithWarning - 裁定書38 Z-46。企業ファイル自動保存の警告に加えて、
'   呼び出し側が持つ追加の警告(policy_noのみ検知時の案内等)を1つのJSONへ
'   まとめる(両方あれば vbLf で連結。どちらも無ければ warning キー自体を
'   持たない Success と同じ形)。
Public Function SavedResultWithWarning(ByVal caseId As String, ByVal message As String, _
                                       ByVal extraWarning As String) As String
    Dim result As String, warnText As String
    result = modCompanyFile3.AutoSaveCase(caseId)
    warnText = extraWarning
    If result <> "saved" Then
        If LenB(warnText) > 0 Then warnText = warnText & vbLf
        warnText = warnText & "企業ファイルへ保存できませんでした。ブックを閉じる前に保存先を確認してください。"
    End If
    If LenB(warnText) = 0 Then
        SavedResultWithWarning = Success(message)
    Else
        SavedResultWithWarning = "{""ok"":true,""message"":" & modNaviJson.Q(message) & _
                                 ",""warning"":" & modNaviJson.Q(warnText) & "}"
    End If
End Function

Private Function BasicValue(ByVal data As String, ByVal srcName As String) As String
    If LenB(modNaviJson.RawField(data, srcName)) > 0 Then
        BasicValue = modJsonLite.GetStr(data, srcName)
    Else
        BasicValue = modJsonLite.GetStr(modNaviState.DraftJson(), srcName)
    End If
End Function

Public Function ValidateBasics(ByVal company As String, ByVal caseType As String) As String
    If LenB(Trim$(company)) = 0 Then
        ValidateBasics = "会社名を入力してください。"
    ElseIf Len(company) > 200 Then
        ValidateBasics = "会社名は200文字以内で入力してください。"
    ElseIf caseType <> "new" And caseType <> "renewal" Then
        ValidateBasics = "種別を選択してください。"
    End If
End Function

Private Function EnsureCase(ByRef caseId As String, ByVal data As String, ByRef reason As String) As Boolean
    Dim company As String, caseType As String
    If LenB(caseId) > 0 Then
        EnsureCase = modNaviStore.CaseExists(caseId)
        Exit Function
    End If
    company = modUtilText.SanitizeInput(BasicValue(data, "company"))
    caseType = BasicValue(data, "case_type")
    reason = ValidateBasics(company, caseType)
    If LenB(reason) > 0 Then Exit Function
    caseId = modCaseStore.NewCase(company, BasicValue(data, "industry_code"), caseType)
    If LenB(caseId) = 0 Then
        reason = "案件を登録できませんでした。"
        Exit Function
    End If
    If Not modNaviStore.SaveBasics(caseId, modNaviState.DraftJson()) Then Exit Function
    If Not modNaviStore.SaveBasics(caseId, data) Then Exit Function
    EnsureCase = True
End Function

Public Function ActSaveBasics(ByRef caseId As String, ByVal data As String) As String
    Dim reason As String
    reason = ValidateBasics(modJsonLite.GetStr(data, "company"), modJsonLite.GetStr(data, "case_type"))
    If LenB(reason) > 0 Then
        ActSaveBasics = Failure(reason, "E0101")
        Exit Function
    End If
    If LenB(caseId) = 0 Then
        modNaviState.SetDraft data
        ActSaveBasics = Success("会社情報を反映しました。資料の登録時に案件を採番します。")
    ElseIf modNaviStore.SaveBasics(caseId, data) Then
        ActSaveBasics = SavedResult(caseId, "会社情報を保存しました。")
    Else
        ActSaveBasics = Failure("会社情報を保存できませんでした。", "E0603")
    End If
End Function

Public Function ActPasteMaterial(ByRef caseId As String, ByVal data As String) As String
    Dim slot As String, body As String, reason As String, company As String, hits As Long
    Dim memo As String, others As String, coverage As String, ok As Boolean, oldBody As String
    Dim piiKinds As String, piiWarn As String
    slot = modJsonLite.GetStr(data, "slot")
    If Not ValidSlot(slot) Then
        ActPasteMaterial = Failure("登録先が不正です。", "E0101")
        Exit Function
    End If
    body = modNavText.NormalizeEol(modNavText.StripDrFooter(modJsonLite.GetStr(data, "text")))
    If LenB(Trim$(body)) = 0 Then
        ActPasteMaterial = Failure("登録する本文を入力してください。", "E0101")
        Exit Function
    End If
    If Len(body) > 100000 Then
        ActPasteMaterial = Failure("1欄の登録上限は100,000文字です。分割して内容を整理してください。", "E0101")
        Exit Function
    End If
    ' 裁定書38 Z-46(伝書鳩3-2): 検知種別が policy_no だけなら登録は止めず警告に
    '   留める(DR出力は出典URLを必ず含み、URL断片を契約番号と誤検知しやすい)。
    '   人名・メール・電話を1件でも含む混在は従来どおりブロックする。
    piiKinds = modPii.KindsOf(body)
    If LenB(piiKinds) > 0 Then
        If piiKinds = "policy_no" Then
            piiWarn = "契約番号らしき数字列(" & CStr(modPii.DetectionCount(body)) & _
                      "箇所: " & modUtil.SafeLeft(body, 30) & "…)を検知しました。" & _
                      "伏せ字にするか、そのままでよいか確認してください。"
            modLog.LogUsage "pii_policy_warning", caseId, modPii.ScanReport(body, "貼付/" & slot)
        Else
            modLog.LogError "E0103", "modNaviActions.ActPasteMaterial", _
                            modPii.ScanReport(body, "貼付/" & slot)
            ActPasteMaterial = Failure("個人情報らしき記述を検知したため登録しませんでした。伏せ字にして登録してください。", "E0103")
            Exit Function
        End If
    End If
    If Not EnsureCase(caseId, data, reason) Then
        ActPasteMaterial = Failure(reason, "E0101")
        Exit Function
    End If
    If slot = "hearing_answers" Then
        If Val(modCaseRead.CaseColumnOf(caseId, "round_no")) < 2 Then
            ActPasteMaterial = Failure("ヒアリング回答は第2ラウンドから登録できます。", "E0101")
            Exit Function
        End If
    End If
    company = modCaseRead.CaseColumnOf(caseId, "company")
    body = modUICase.AnonymizeText(modUtilText.SanitizeInput(body), company, hits)
    If slot = "field_notes" Then
        modNavText.SplitFieldNotes body, memo, others
        coverage = modNavText.CoverageNoteOf(body)
        ok = SaveFieldNotes(caseId, memo, others, coverage)
    Else
        ok = modCaseStore.SaveData(caseId, modUICase6.AreaField(slot, 1), body)
    End If
    If ok Then
        modLog.LogUsage "material_registered", caseId, "slot=" & slot & ";chars=" & CStr(Len(body))
        ActPasteMaterial = SavedResultWithWarning(caseId, "資料を登録しました。", piiWarn)
    Else
        ActPasteMaterial = Failure("資料を登録できませんでした。", "E0604")
    End If
End Function

Private Function SaveFieldNotes(ByVal caseId As String, ByVal memo As String, _
                                ByVal others As String, ByVal coverage As String) As Boolean
    Dim oldMemo As String, oldOther As String, oldCoverage As String
    oldMemo = modCaseStore.LoadData(caseId, "input_memo")
    oldOther = modCaseStore.LoadData(caseId, "input_field_notes")
    oldCoverage = modCaseStore.LoadData(caseId, "input_coverage_note")
    If Not modCaseStore.SaveData(caseId, "input_memo", memo) Then Exit Function
    If Not modCaseStore.SaveData(caseId, "input_field_notes", others) Then GoTo RollBack
    If Not modCaseStore.SaveData(caseId, "input_coverage_note", coverage) Then GoTo RollBack
    SaveFieldNotes = True
    Exit Function
RollBack:
    modCaseStore.SaveData caseId, "input_memo", oldMemo
    modCaseStore.SaveData caseId, "input_field_notes", oldOther
    modCaseStore.SaveData caseId, "input_coverage_note", oldCoverage
End Function

Public Function ActClearMaterial(ByVal caseId As String, ByVal data As String) As String
    Dim slot As String, ok As Boolean
    slot = modJsonLite.GetStr(data, "slot")
    If Not ValidSlot(slot) Then
        ActClearMaterial = Failure("登録先が不正です。", "E0101")
        Exit Function
    End If
    If Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
        ActClearMaterial = Confirm("この資料の登録内容を消去します。")
        Exit Function
    End If
    If slot = "field_notes" Then
        ok = SaveFieldNotes(caseId, "", "", "")
    Else
        ok = modCaseStore.SaveData(caseId, modUICase6.AreaField(slot, 1), "")
    End If
    If ok Then
        ActClearMaterial = SavedResult(caseId, "登録内容を消去しました。")
    Else
        ActClearMaterial = Failure("登録内容を消去できませんでした。", "E0604")
    End If
End Function

Public Function ActSaveMaterials(ByRef caseId As String, ByVal data As String) As String
    Dim reason As String, tier As String
    If Not EnsureCase(caseId, data, reason) Then
        ActSaveMaterials = Failure(reason, "E0101")
        Exit Function
    End If
    If Not modNaviStore.SaveBasics(caseId, data) Then
        ActSaveMaterials = Failure("会社情報を保存できませんでした。", "E0603")
        Exit Function
    End If
    tier = modCaseRead.CaseColumnOf(caseId, "dossier_tier")
    If tier <> "t3_sparring" Then
        If LenB(modCaseStore.LoadData(caseId, "input_dossier")) > 0 Then tier = "t2_full" Else tier = "t1_quick"
        If Not modCaseStore.PromoteTier(caseId, tier) Then
            ActSaveMaterials = Failure("調べる深さを保存できませんでした。", "E0603")
            Exit Function
        End If
    End If
    ActSaveMaterials = SavedResult(caseId, "登録内容を保存しました。")
End Function

Public Function ActRunPipeline(ByVal caseId As String, ByVal data As String) As String
    Dim fromStep As Long, n As Long, downstream As Boolean, warning As String, saveResult As String
    Dim editError As String, lastError As String
    fromStep = modJsonLite.GetLong(data, "from_step", 1)
    If fromStep < 1 Or fromStep > 4 Then
        ActRunPipeline = Failure("段番号が不正です。", "E0101")
        Exit Function
    End If
    For n = fromStep To 4
        If LenB(modCaseStore.ResolveStepJson(caseId, n)) > 0 Then downstream = True
    Next n
    If downstream And Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
        ActRunPipeline = Confirm("この段から再分析します。下流の下書きと編集内容が更新されます。")
        Exit Function
    End If
    If Not FlushOwnedSheets(caseId, fromStep, editError) Then
        ActRunPipeline = Failure(editError, "E0302")
        Exit Function
    End If
    modPipeline2.ResetDeepOutcome
    If downstream Then modCaseStore.InvalidateDownstream caseId, fromStep - 1
    For n = fromStep To 4
        If gSheetCase(n) = caseId Then
            gSheetCase(n) = ""
            gSheetSnapshot(n) = ""
        End If
    Next n
    For n = fromStep To 4
        modNaviHost.ShowBusy modUIHome2.StageNameOf(n), ProgressJson(n, 4, modUIHome2.StageNameOf(n))
        DoEvents
        If Not modPipeline.RunStep(caseId, n, "") Then
            lastError = modNaviStore.LastErrorJson()
            modLog.LogError "E0603", NA_SRC & ".ActRunPipeline", caseId & ":s" & CStr(n)
            ActRunPipeline = "{""ok"":false,""message"":""分析が完了しませんでした。ログを確認してください。""," & _
              """failed_step"":" & CStr(n) & ",""failed_raw_available"":" & _
              modNaviJson.Flag(LenB(modCaseStore.LoadData(caseId, "s" & CStr(n) & "_json_failed")) > 0) & _
              ",""last_error"":" & lastError & "}"
            Exit Function
        End If
        saveResult = modCompanyFile3.AutoSaveCase(caseId)
        If saveResult <> "saved" Then warning = "企業ファイルへの自動保存に失敗しました。保存先を確認してください。"
        If n = 1 Then
            ' 裁定書37 B-05(表示側)。S1成功直後に充足度を見て iq=low なら
            ' 16章 E-02 の警告(一般論に近い出力になります+advice)を warning へ
            ' 足し、続行する(モーダル無し)。値源は modPipeline3.SufficiencyNoteOf。
            Dim suffNote As String, s1Now As String, adviceText As String
            s1Now = modCaseStore.ResolveStepJson(caseId, 1)
            suffNote = modPipeline3.SufficiencyNoteOf(s1Now)
            If Left$(suffNote, 6) = "iq=low" Then
                adviceText = Trim$(modJsonLite.GetStr(s1Now, "advice"))
                If LenB(warning) > 0 Then warning = warning & vbLf
                warning = warning & "入力が薄いため、一般論に近い出力になります。"
                If LenB(adviceText) > 0 Then warning = warning & " 助言: " & adviceText
            End If
        End If
        DoEvents
    Next n
    If LenB(modPipeline2.LastDeepOutcome()) > 0 Then
        If LenB(warning) > 0 Then warning = warning & vbLf
        warning = warning & modPipeline2.DeepWarningOf(modPipeline2.LastDeepOutcome())
    End If
    ActRunPipeline = "{""ok"":true,""message"":""分析が完了しました。"",""warning"":" & modNaviJson.Q(warning) & "}"
End Function

Public Function ProgressJson(ByVal n As Long, ByVal total As Long, ByVal srcName As String) As String
    ProgressJson = "{""stage_no"":" & CStr(n) & ",""stage_count"":" & CStr(total) & _
                   ",""step"":" & CStr(n) & ",""total"":" & CStr(total) & ",""name"":" & modNaviJson.Q(srcName) & _
                   ",""started_at"":" & modNaviJson.Q(modUtil.NowStamp()) & ",""max_wait_sec"":" & _
                   CStr(modGatewayRPN2.ResolveWaitSec(modConfig.GetLong("llm_wait_sec", 1200))) & "}"
End Function

Public Function ActExportReport(ByVal caseId As String, ByVal reviewedBy As String) As String
    Dim path As String, reason As String, status As String
    If Not FlushOwnedSheets(caseId, 5, reason) Then
        ActExportReport = Failure(reason, "E0302")
        Exit Function
    End If
    If LenB(modCaseStore.ResolveStepJson(caseId, 1)) = 0 Then
        ActExportReport = Failure("先に企業プロファイル分析を実行してください。", "E0101")
        Exit Function
    End If
    reason = ExportWithReview(caseId, reviewedBy, path)
    If LenB(reason) > 0 Or LenB(path) = 0 Then
        ActExportReport = Failure(reason, "E0603")
        Exit Function
    End If
    status = modCaseRead.CaseColumnOf(caseId, "status")
    If modCaseStore.CanTransition(status, "exported") Then modCaseStore.SetStatus caseId, "exported"
    ActExportReport = SavedResult(caseId, "レポートを出力しました。")
End Function

' 裁定書37 B-06(UI側)。reviewedBy が空なら従来どおり確認前の免責のまま出す
' (GenerateHtmlReport)。非空なら班2の GenerateHtmlReportEx(caseId, reviewedBy)
' で確認済みの免責へ切り替える契約(裁定書37 §2 班2欄)。呼び分けはこの1関数に
' 閉じ、マージ時はここだけを差し替える。
Private Function ExportWithReview(ByVal caseId As String, ByVal reviewedBy As String, ByRef path As String) As String
    If LenB(reviewedBy) > 0 Then
        ExportWithReview = modExportHtml.GenerateHtmlReportEx(caseId, path, reviewedBy)
    Else
        ExportWithReview = modExportHtml.GenerateHtmlReport(caseId, path)
    End If
End Function

Public Function ActExportHearing(ByVal caseId As String, ByVal data As String) As String
    Dim editError As String
    If Not FlushOwnedSheets(caseId, 5, editError) Then
        ActExportHearing = Failure(editError, "E0302")
        Exit Function
    End If
    If modExportHearing.AnswerMemoCount(caseId) > 0 And Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
        ActExportHearing = Confirm("ヒアリングシートの手書き回答を上書きします。")
        Exit Function
    End If
    If modExportHearing.BuildHearingSheet(caseId) Then
        modUISheet.ShowSheet "ヒアリングシート"
        modLog.LogUsage "hearing_built", caseId, "ok"
        ActExportHearing = Success("ヒアリングシートを出力しました。Excelで印刷できます。")
    Else
        ActExportHearing = Failure("ヒアリングシートを出力できませんでした。骨子を確認してください。", "E0603")
    End If
End Function

Public Function ActOpenReport(ByVal caseId As String, ByVal data As String) As String
    Dim path As String, given As String, ok As Boolean
    path = modCaseRead.CaseColumnOf(caseId, "report_path")
    given = modJsonLite.GetStr(data, "path")
    If LenB(path) = 0 Or (LenB(given) > 0 And given <> path) Then
        ActOpenReport = Failure("この案件に登録されたレポートがありません。", "E0101")
        Exit Function
    End If
    If Not modUtil.FileExistsAt(path) Then
        ActOpenReport = Failure("レポートファイルが見つかりません。", "E0603")
        Exit Function
    End If
    If modJsonLite.GetStr(data, "kind") = "folder" Then path = modUtil.ParentDirOf(path, modUtil.PathSep())
    ok = modUIResearch.OpenUrl(path)
    ActOpenReport = ResultOf(ok, "レポートを開きました。", "レポートを開けませんでした。")
End Function

Public Function ActOpenUrl(ByVal data As String) As String
    Dim kind As String, url As String
    kind = modJsonLite.GetStr(data, "kind")
    Select Case kind
    Case "full", "quick", "menu"
        url = modUIResearch.DrUrlOf(kind, modConfig.GetStr("dr_url_" & kind, ""))
    Case "portal"
        url = modConfig.GetStr("portal_url", "")
    Case "help"
        ActOpenUrl = ResultOf(modUISheet.ShowSheet("使い方"), "使い方を開きました。", "使い方を開けませんでした。")
        Exit Function
    Case Else
        ActOpenUrl = Failure("このページは開けません。", "E0101")
        Exit Function
    End Select
    ActOpenUrl = ResultOf(modUIResearch.OpenUrl(url), "ページを開きました。", "ページを開けませんでした。")
End Function

Public Function ActCopyPrompt(ByVal caseId As String, ByVal data As String) As String
    Dim n As Long, stamp As String, json As String
    n = modJsonLite.GetLong(data, "prompt_no", 0)
    If n < 1 Or n > 8 Then
        ActCopyPrompt = Failure("調査指示文の番号が不正です。", "E0101")
        Exit Function
    End If
    stamp = modUtil.NowStamp()
    json = "{""copied_" & CStr(n) & """:" & modNaviJson.Q(stamp) & "}"
    If LenB(caseId) > 0 Then
        json = modNaviStore.MergeBasics(modCaseStore.LoadData(caseId, "nav_basics"), json)
        modCaseStore.SaveData caseId, "nav_basics", json
        modCompanyFile3.AutoSaveCase caseId
    Else
        modNaviState.SetDraft modNaviJson.ReplaceTextField(modNaviState.DraftJson(), "copied_" & CStr(n), stamp)
    End If
    modLog.LogUsage "research_prompt_copied", caseId, "no=" & CStr(n)
    If modConfig.GetBool("dr_open_after_copy", False) Then
        modUIResearch.OpenUrl modUIResearch.DrUrlOf("full", modConfig.GetStr("dr_url_full", ""))
    End If
    ActCopyPrompt = Success("調査指示文をコピーしました。")
End Function

Public Function PhaseTwoStepEdit(ByVal caseId As String, ByVal data As String) As String
    ' F-16 phase 2 extension. Business JSON is not parsed or rewritten here.
    PhaseTwoStepEdit = Failure("直接編集は第2段の機能です。「シートで編集」を利用してください。", "E0101")
End Function


Public Function FlushOwnedSheets(ByVal caseId As String, ByVal beforeStep As Long, ByRef reason As String) As Boolean
    Dim n As Long, current As String, saveState As String
    FlushOwnedSheets = True
    For n = 1 To 4
        If n < beforeStep And gSheetCase(n) = caseId And LenB(caseId) > 0 Then
            current = modUICase2.SerializeStep(n)
            If current <> gSheetSnapshot(n) Then
                If Not modUICase2.SaveEditedStep(caseId, n, reason) Then
                    If LenB(reason) = 0 Then reason = "表示案件または表示行数を確認してください。"
                    reason = "S" & CStr(n) & "のシート編集を保存できません。" & reason
                    FlushOwnedSheets = False
                    Exit Function
                End If
                saveState = modCompanyFile3.AutoSaveCase(caseId)
                If saveState <> "saved" Then
                    reason = "シート編集を企業ファイルへ保存できません。保存先を確認してください。"
                    FlushOwnedSheets = False
                    Exit Function
                End If
                gSheetSnapshot(n) = current
            End If
        End If
    Next n
End Function

Private Function FlushAllSheets(ByRef reason As String) As Boolean
    Dim n As Long
    FlushAllSheets = True
    For n = 1 To 4
        If LenB(gSheetCase(n)) > 0 Then
            If Not FlushOwnedSheets(gSheetCase(n), 5, reason) Then
                FlushAllSheets = False
                Exit Function
            End If
        End If
    Next n
End Function
