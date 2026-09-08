Attribute VB_Name = "modNaviActions2"
Option Explicit

' [NAVI] Specification 7.2. Advanced actions are separated for the 30000-character contract.
Public Function DispatchMore(ByVal action As String, ByVal data As String, ByRef caseId As String) As String
    Dim ok As Boolean, note As String, result As String, n As Long, id As String
    Select Case action
    Case "start_round2"
        If Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
            result = modNaviActions.Confirm("現在の分析を保存し、次のラウンドを始めます。")
        Else
            n = modCaseStore.FreezeRound(caseId)
            If n > 0 Then
                result = modNaviActions.SavedResult(caseId, CStr(n) & "巡目を開始しました。")
            Else
                result = modNaviActions.Failure("ラウンドを開始できませんでした。", "E0603")
            End If
        End If
    Case "company_save"
        result = ActCompanySave(caseId, data)
    Case "company_open"
        result = ActCompanyOpen(caseId)
    Case "feedback_add"
        note = modJsonLite.GetStr(data, "customer_quote")
        If modPii.HasPii(note) And Not modJsonLite.GetBoolJ(data, "mask_pii", True) And _
           Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
            result = modNaviActions.Confirm("個人情報らしき記述を原文のまま保存します。伏せ字にする場合は戻って選択してください。")
        ElseIf modNaviStore.AppendFeedback(caseId, data) Then
            modCaseStore.SetStatus caseId, "feedback_done"
            modLog.LogUsage "feedback_saved", caseId, "html"
            result = modNaviActions.SavedResult(caseId, "商談の記録を保存しました。")
        Else
            result = modNaviActions.Failure("商談の記録を保存できませんでした。", "E0603")
        End If
    Case "inbox_list"
        result = "{""ok"":true,""inbox"":" & modNaviStore.ListInbox() & "}"
    Case "inbox_post"
        result = ActInboxPost(data)
    Case "inbox_diagnose_all"
        If Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
            result = modNaviActions.Confirm("未診断の受信箱をまとめて診断します。案件ごとにAI利用枠を使います。")
        Else
            modNaviHost.ShowBusy "受信箱を診断しています", modNaviActions.ProgressJson(1, 1, "受信箱の診断")
            DoEvents
            n = modPlayOps.RunPreflightAll()
            result = "{""ok"":true,""message"":" & modNaviJson.Q(CStr(n) & "件の診断が完了しました。") & _
                     ",""inbox"":" & modNaviStore.ListInbox() & "}"
        End If
    Case "inbox_judge"
        result = ActInboxJudge(data)
    Case "judge_list"
        result = "{""ok"":true,""judgements"":" & modNaviStore.ListJudgements(caseId) & "}"
    Case "judge_add"
        result = ActJudgeAdd(caseId, data)
    Case "judge_result"
        id = modJsonLite.GetStr(data, "judge_id")
        If Not JudgeBelongsToCase(id, caseId) Then
            result = modNaviActions.Failure("この案件の判断が見つかりません。", "E0101")
        Else
            ok = modJudgeStore.SetJudgementResult(id, modJsonLite.GetStr(data, "result"), _
                                                   modJsonLite.GetStr(data, "post_loss"))
            If ok Then
                result = modNaviActions.SavedResult(caseId, "判断の結果を記録しました。")
            Else
                result = modNaviActions.Failure("判断の結果を記録できませんでした。", "E0101")
            End If
        End If
    Case "logs"
        result = "{""ok"":true,""logs"":" & modNaviStore.LogRowsOf(caseId) & "}"
    Case "reload_kb"
        result = modNaviActions.ResultOf(modKnowledge.LoadKnowledge(), _
                                         "ナレッジを読み直しました。", "ナレッジを読み込めませんでした。")
    Case "run_tests"
        result = ActRunTests(data)
    Case "rename_case"
        note = modJsonLite.GetStr(data, "display_name")
        If LenB(modNaviJson.RawField(data, "display_name")) = 0 Then note = modJsonLite.GetStr(data, "name")
        If modNaviStore.SetDisplayName(caseId, note) Then
            result = modNaviActions.SavedResult(caseId, "案件の表示名を変更しました。")
        Else
            result = modNaviActions.Failure("表示名を変更できませんでした。200文字以内で入力してください。", "E0101")
        End If
    Case "archive_case"
        If Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
            result = modNaviActions.Confirm("案件を一覧から非表示にします。企業ファイルと案件データは保持されます。")
        Else
            If modJsonLite.GetBoolJ(data, "export_first", False) Then
                result = ActExportCase(caseId, data)
                If Not modJsonLite.GetBoolJ(result, "ok", False) Then
                    DispatchMore = result
                    Exit Function
                End If
            End If
            If modNaviStore.SetArchived(caseId) Then
                result = modNaviActions.SavedResult(caseId, "案件を非表示にしました。")
                caseId = ""
                modNaviState.SetDraft "{}"
            Else
                result = modNaviActions.Failure("案件を非表示にできませんでした。archived_at列を確認してください。", "E0603")
            End If
        End If
    Case "export_case"
        result = ActExportCase(caseId, data)
    Case "import_case"
        result = ActImportCase(caseId)
    Case "save_settings"
        ok = modNaviStore.SaveSettings(data, note)
        result = modNaviActions.ResultOf(ok, "設定を保存しました。画面モードは次回起動から反映されます。", note)
    Case Else
        result = modNaviActions.Failure("この操作は受け付けられません。", "E0101")
    End Select
    DispatchMore = result
End Function

Public Function ActCompanySave(ByVal caseId As String, ByVal data As String) As String
    Dim pii As String, stamp As String, path As String
    pii = modCompanyFile.ScanCaseForPii(caseId)
    If LenB(pii) > 0 And Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
        ActCompanySave = modNaviActions.Confirm("個人情報らしき記述があります。企業ファイルへの保存内容を確認してください。")
        Exit Function
    End If
    If modJsonLite.GetBoolJ(data, "confirmed", False) Then stamp = modUtil.NowStamp()
    path = modCompanyFile.ExportCompanyFile(caseId, modCompanyFile3.CompanyDir(), stamp)
    If LenB(path) = 0 Then
        ActCompanySave = modNaviActions.Failure("企業ファイルへ保存できませんでした。", "E0603")
    Else
        ActCompanySave = "{""ok"":true,""message"":""企業ファイルへ保存しました。"",""path"":" & modNaviJson.Q(path) & "}"
    End If
End Function

Public Function ActCompanyOpen(ByVal caseId As String) As String
    Dim selected As Variant
    selected = Application.GetOpenFilename("企業ファイル (*.xlsx),*.xlsx", 1, "企業ファイルを選択")
    If VarType(selected) = 11 Then
        ActCompanyOpen = "{""ok"":true,""cancelled"":true}"
        Exit Function
    End If
    If modCompanyFile.ImportCompanyFile(CStr(selected), caseId) Then
        ActCompanyOpen = modNaviActions.SavedResult(caseId, "企業ファイルを取り込みました。")
    Else
        ActCompanyOpen = modNaviActions.Failure("企業ファイルを取り込めませんでした。", "E0603")
    End If
End Function

Public Function ActInboxPost(ByVal data As String) As String
    Dim sourceKind As String, theme As String, body As String, id As String
    sourceKind = modJsonLite.GetStr(data, "source_kind")
    If LenB(sourceKind) = 0 Then sourceKind = modJsonLite.GetStr(data, "source")
    theme = modUtilText.SanitizeInput(modJsonLite.GetStr(data, "theme"))
    body = modUtilText.SanitizeInput(modJsonLite.GetStr(data, "body"))
    If LenB(Trim$(theme)) = 0 Or LenB(Trim$(body)) = 0 Then
        ActInboxPost = modNaviActions.Failure("テーマと本文を入力してください。", "E0101")
        Exit Function
    End If
    If modPii.HasPii(body) Or modPii.HasPii(theme) Then
        ActInboxPost = modNaviActions.Failure("個人情報らしき記述を検知したため登録しませんでした。", "E0103")
        Exit Function
    End If
    id = modInboxStore.NewInboxItem(sourceKind, theme, body)
    If LenB(id) = 0 Then
        ActInboxPost = modNaviActions.Failure("受信箱へ登録できませんでした。", "E0101")
    Else
        ActInboxPost = "{""ok"":true,""message"":""受信箱へ登録しました。"",""inbox"":" & modNaviStore.ListInbox() & "}"
    End If
End Function

Public Function ParseIsoDate(ByVal srcText As String, ByRef srcValue As Date) As Boolean
    On Error GoTo Invalid
    Dim yy As Long, mm As Long, dd As Long
    If Len(srcText) <> 10 Or Mid$(srcText, 5, 1) <> "-" Or Mid$(srcText, 8, 1) <> "-" Then Exit Function
    yy = CLng(Left$(srcText, 4))
    mm = CLng(Mid$(srcText, 6, 2))
    dd = CLng(Right$(srcText, 2))
    If yy < 1900 Or yy > 9999 Or mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then Exit Function
    srcValue = DateSerial(yy, mm, dd)
    ParseIsoDate = (modUtilText.IsoDate(srcValue) = srcText)
    Exit Function
Invalid:
    ParseIsoDate = False
End Function

Public Function ActInboxJudge(ByVal data As String) As String
    Dim status As String, dropType As String, reviveTag As String, dueText As String
    Dim due As Date, reason As String, ok As Boolean
    status = modJsonLite.GetStr(data, "status")
    If LenB(status) = 0 Then status = modJsonLite.GetStr(data, "judgement")
    dropType = modJsonLite.GetStr(data, "drop_type")
    reviveTag = modJsonLite.GetStr(data, "revive_tag")
    dueText = modJsonLite.GetStr(data, "revive_due")
    If LenB(dueText) > 0 Then
        If Not ParseIsoDate(dueText, due) Then
            ActInboxJudge = modNaviActions.Failure("見直し期日を年-月-日の形式で入力してください。", "E0101")
            Exit Function
        End If
    End If
    reason = modInboxStore.JudgementError(status, dropType, reviveTag, LenB(dueText) > 0)
    If LenB(reason) > 0 Then
        ActInboxJudge = modNaviActions.Failure(reason, "E0101")
        Exit Function
    End If
    ok = modInboxStore.SetInboxJudgement(modJsonLite.GetStr(data, "inbox_id"), status, dropType, reviveTag, due)
    If ok Then
        ActInboxJudge = "{""ok"":true,""message"":""受信箱の判定を保存しました。"",""inbox"":" & modNaviStore.ListInbox() & "}"
    Else
        ActInboxJudge = modNaviActions.Failure("受信箱の判定を保存できませんでした。状態を確認してください。", "E0101")
    End If
End Function

Private Function JudgeBelongsToCase(ByVal judgeId As String, ByVal caseId As String) As Boolean
    Dim rec As TJudgement
    If Not modJudgeStore.ReadJudgement(judgeId, rec) Then Exit Function
    JudgeBelongsToCase = (rec.case_ref = caseId)
End Function

Public Function ActJudgeAdd(ByVal caseId As String, ByVal data As String) As String
    Dim rec As TJudgement, id As String
    rec.line_id = modJsonLite.GetStr(data, "line_id")
    rec.case_ref = caseId
    rec.situation = modUtilText.SanitizeInput(modJsonLite.GetStr(data, "situation"))
    rec.decision = modJsonLite.GetStr(data, "decision")
    rec.factor_note = modUtilText.SanitizeInput(modJsonLite.GetStr(data, "factor_note"))
    rec.key_reason = modUtilText.SanitizeInput(modJsonLite.GetStr(data, "key_reason"))
    rec.result = modJsonLite.GetStr(data, "result")
    rec.post_loss = modUtilText.SanitizeInput(modJsonLite.GetStr(data, "post_loss"))
    rec.recorded_by = modCaseStore2.OwnerName()
    id = modJudgeStore.NewJudgement(rec)
    If LenB(id) = 0 Then
        ActJudgeAdd = modNaviActions.Failure("判断を登録できませんでした。必須欄と種目・判断の値を確認してください。", "E0101")
    Else
        ActJudgeAdd = modNaviActions.SavedResult(caseId, "判断を登録しました。")
    End If
End Function

Public Function ActExportCase(ByVal caseId As String, ByVal data As String) As String
    Dim dirPath As String, path As String, payload As String, pii As String
    pii = modCompanyFile.ScanCaseForPii(caseId)
    If LenB(pii) > 0 And Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
        ActExportCase = modNaviActions.Confirm("個人情報らしき記述があります。案件JSONの出力内容を確認してください。")
        Exit Function
    End If
    dirPath = modUtil.ResolveDataDir(modConfig.GetStr("data_dir", ""), ThisWorkbook.Path)
    dirPath = modUtilPath.JoinPath(dirPath, "案件")
    If Not modUtil.EnsureFolder(dirPath) Then
        ActExportCase = modNaviActions.Failure("出力フォルダを用意できませんでした。", "E0603")
        Exit Function
    End If
    path = modUtilPath.JoinPath(dirPath, caseId & ".json")
    If modUtil.FileExistsAt(path) And Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
        ActExportCase = modNaviActions.Confirm("既存の案件JSONを上書きします。")
        Exit Function
    End If
    payload = modNaviStore.ExportCaseJson(caseId)
    If LenB(payload) = 0 Or Not modUtil.WriteUtf8File(path, payload, True) Then
        ActExportCase = modNaviActions.Failure("案件JSONを出力できませんでした。", "E0603")
    Else
        ActExportCase = "{""ok"":true,""message"":""案件JSONを出力しました。"",""path"":" & modNaviJson.Q(path) & "}"
    End If
End Function

Public Function ActImportCase(ByRef caseId As String) As String
    Dim selected As Variant, payload As String, reason As String, newId As String
    selected = Application.GetOpenFilename("案件JSON (*.json),*.json", 1, "案件JSONを選択")
    If VarType(selected) = 11 Then
        ActImportCase = "{""ok"":true,""cancelled"":true}"
        Exit Function
    End If
    payload = modUtil.ReadUtf8File(CStr(selected))
    newId = modNaviStore.ImportCaseJson(payload, reason)
    If LenB(newId) = 0 Then
        ActImportCase = modNaviActions.Failure(reason, "E0101")
    Else
        caseId = newId
        ActImportCase = modNaviActions.SavedResult(caseId, "新しい案件IDで取り込みました。")
    End If
End Function

' 自己テストの実行(17章 T-48)。
'   R1(製品コードからテスト層を参照しない・12章§4)の唯一の逃げ道は
'   vba_lint の R1_TEST_LAYER_EXCEPTIONS に名指しで登録された組だけである。
'   既存の[テストを実行]ボタン(modUIGuide -> modTestsRunnerUi.RunAllTestsFromBook)
'   と同じ扱いで、HTML画面用の 2本 だけを登録した。
'   合否の4条件・集計・gd_test_result への書込は modTestsRunnerUi 側が持つ
'   (判定を2箇所に置かない)。髙橋さん版はここで modTestRunner / modTestsExcel を
'   直に呼び、期待本数と4条件の判定をこの関数の中で作り直していた。
Public Function ActRunTests(ByVal data As String) As String
    Dim summary As String, report As String, ok As Boolean
    If Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
        ActRunTests = modNaviActions.Confirm("自己テストを実行します。数分かかり、検査の行がログへ残ります。")
        Exit Function
    End If
    modNaviHost.ShowBusy "自己テストを実行しています", modNaviActions.ProgressJson(1, 1, "自己テスト")
    DoEvents
    On Error GoTo Failed
    summary = modTestsRunnerUi.RunAllTestsHeadless()
    report = modTestsRunnerUi.LastReportText()
    ok = (InStr(1, summary, "全PASS", vbBinaryCompare) = 1)
    ActRunTests = "{""ok"":" & modNaviJson.Flag(ok) & ",""message"":" & modNaviJson.Q(summary) & _
                  ",""test_report"":" & modNaviJson.Q(report) & "}"
    Exit Function
Failed:
    ActRunTests = modNaviActions.Failure("自己テストを実行できませんでした。", "E0603")
End Function
