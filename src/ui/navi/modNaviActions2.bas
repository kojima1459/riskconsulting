Attribute VB_Name = "modNaviActions2"
Option Explicit

' 裁定書39 R1-06: modNaviActions が30,000字契約に達したため、区画①の
'   [コピー](copy_prompt)と、その直前の下見をこちらへ分けた(12章§2)。
'   入口は modNaviActions.Dispatch の Case "copy_prompt" の1本のまま。
' [コピー]直前の下見(16章 E-69・裁定書39 R1-06)。
'   NA_COPY_MIN_LEN : 一致とみなす断片の最短長(20字)。
'   NA_COPY_SCAN_MAX: 下敷き(現契約・営業メモ・現場メモ)の走査上限。
Private Const NA_COPY_MIN_LEN As Long = 20
Private Const NA_COPY_SCAN_MAX As Long = 30000

' [コピー]の逐語案内(裁定書44 A-3・docs/08 §1y)。貼り付け先と「調査の広がり」の
'   選択まで言い切る(社内の調査ページの操作を1回で終えられるように)。
Private Const NA2_COPY_GUIDE As String = _
    "社内の調査ページで「レポート調査」→「指示文」欄に貼り付け、" & _
    "「調査の広がり」は「集中（1観点）」を選んでください。"

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
' 裁定書44 A-2(F-2): run_tests は IsLongAction から外れ、UiLockを握らなくなった
'   (modNaviActions.IsLongAction)。その代わり、この関数自身の再入(HTML画面から
'   連打)を Static running で防ぐ。UiLockと違い**この関数の実行中だけ**を
'   ブロックするので、層(b)の「貼ったものの保存」等が自分自身のロックとぶつかって
'   SKIPされる事故(F-2)が起きない。
Public Function ActRunTests(ByVal data As String) As String
    Static running As Boolean
    Dim summary As String, report As String, ok As Boolean
    If running Then
        ActRunTests = "{""ok"":false,""busy"":true,""message"":""処理中です""}"
        Exit Function
    End If
    If Not modJsonLite.GetBoolJ(data, "confirmed", False) Then
        ActRunTests = modNaviActions.Confirm("自己テストを実行します。数分かかり、検査の行がログへ残ります。")
        Exit Function
    End If
    running = True
    modNaviHost.ShowBusy "自己テストを実行しています", modNaviActions.ProgressJson(1, 1, "自己テスト")
    DoEvents
    On Error GoTo Failed
    summary = modTestsRunnerUi.RunAllTestsHeadless()
    report = modTestsRunnerUi.LastReportText()
    ok = (InStr(1, summary, "全PASS", vbBinaryCompare) = 1)
    ActRunTests = "{""ok"":" & modNaviJson.Flag(ok) & ",""message"":" & modNaviJson.Q(summary) & _
                  ",""test_report"":" & modNaviJson.Q(report) & "}"
    running = False
    Exit Function
Failed:
    running = False
    ActRunTests = modNaviActions.Failure("自己テストを実行できませんでした。", "E0603")
End Function

' ActCopyPrompt - 区画①[コピー](裁定書44 A-3・F-3)。
'   旧実装は HTML の window.clipboardData に頼っていたが、WebBrowser 制御内では
'   成功を返してもクリップボードへ入らない環境がある(EDR/クリップボード制限)。
'   ここでは **Excel 自身のコピー**(modUICase7.ClipCopyText。シート予備画面の
'   [コピー]と同じ実装)を使う。応答の "copied" が false のときだけ、画面側
'   (ui/app.js::copyPrompt)が旧来の window.clipboardData → 手動コピーの
'   モーダルへ落ちる(片方だけ直さない・二段構え)。
Public Function ActCopyPrompt(ByVal caseId As String, ByVal data As String) As String
    Dim n As Long, stamp As String, json As String, warnText As String, basicsJson As String
    Dim bodyText As String, copied As Boolean, openedUrl As String, message As String
    n = modJsonLite.GetLong(data, "prompt_no", 0)
    If n < 1 Or n > 8 Then
        ActCopyPrompt = modNaviActions.Failure("調査指示文の番号が不正です。", "E0101")
        Exit Function
    End If
    stamp = modUtil.NowStamp()
    json = "{""copied_" & CStr(n) & """:" & modNaviJson.Q(stamp) & "}"
    ' data.text は画面(状態から組み立て済みの本文)。空のときはサーバ側が正を
    ' 持つ PromptTextOf で作り直す(画面のJSを書き換えて空文字を送っても、
    ' 個人情報の下見(CopyWarningOf)を素通りできない)。
    bodyText = modJsonLite.GetStr(data, "text")
    If LenB(caseId) > 0 Then
        basicsJson = modCaseStore.LoadData(caseId, "nav_basics")
        If Not modNaviJson.IsValidJson(basicsJson) Then basicsJson = "{}"
        If LenB(bodyText) = 0 Then
            bodyText = modNaviState.PromptTextOf(n, modCaseRead.CaseColumnOf(caseId, "company"), _
                                                 basicsJson, modCaseRead.CaseColumnOf(caseId, "industry_name"))
        End If
        ' 裁定書39 R1-06 / 16章 E-69: 下見はコピーの直前=ここだけで行う。
        ' 画面更新(BuildCaseState)では走らせない(8本×最大30万字になるため)。
        warnText = CopyWarningOf(bodyText, CopySourceTextOf(caseId))
        json = modNaviStore.MergeBasics(basicsJson, json)
        modCaseStore.SaveData caseId, "nav_basics", json
        modCompanyFile3.AutoSaveCase caseId
    Else
        If LenB(bodyText) = 0 Then
            bodyText = modNaviState.PromptTextOf(n, modJsonLite.GetStr(modNaviState.DraftJson(), "company"), _
                                                 modNaviState.DraftJson(), _
                                                 modJsonLite.GetStr(modNaviState.DraftJson(), "industry_name"))
        End If
        modNaviState.SetDraft modNaviJson.ReplaceTextField(modNaviState.DraftJson(), "copied_" & CStr(n), stamp)
    End If
    copied = modUICase7.ClipCopyText(bodyText)
    modLog.LogUsage "research_prompt_copied", caseId, "no=" & CStr(n) & ";copied=" & CStr(copied)
    If modConfig.GetBool("dr_open_after_copy", False) Then
        openedUrl = modUIResearch.DrUrlOf("full", modConfig.GetStr("dr_url_full", ""))
        modUIResearch.OpenUrl openedUrl
    End If
    message = "調査指示文をコピーしました。" & NA2_COPY_GUIDE
    If LenB(warnText) > 0 Then message = message & warnText
    ActCopyPrompt = "{""ok"":true,""copied"":" & modNaviJson.Flag(copied) & _
        ",""opened_url"":" & modNaviJson.Q(openedUrl) & _
        IIf(LenB(warnText) > 0, ",""kind"":""warn""", "") & _
        ",""message"":" & modNaviJson.Q(message) & "}"
End Function

' CopyWarningOf - [コピー]の直前の下見(16章 E-69・裁定書39 R1-06)。
'   (a) 個人情報らしき記述があるか (b) 現契約サマリ・営業メモ(sourceText)と
'   NA_COPY_MIN_LEN 字以上一致する断片があるか、を見て警告文を返す
'   (どちらも無ければ空文字)。**コピー自体は止めない**。
'   判定は純関数なので modTestsPureNavi が直接叩ける。
Public Function CopyWarningOf(ByVal bodyText As String, ByVal sourceText As String) As String
    If modPii.HasPii(bodyText) Then
        CopyWarningOf = "個人情報らしき記述が含まれています。"
        Exit Function
    End If
    If LenB(sourceText) = 0 Then Exit Function
    If modPii.SharesLongFragment(bodyText, sourceText, NA_COPY_MIN_LEN) Then
        CopyWarningOf = "現契約・営業メモと20字以上一致する記述が含まれています。"
    End If
End Function

' 裁定書44 B-1(16章 E-05・F-4): 貼付欄のPII方針。person/email/phone を
'   1件でも含む混在は config pii_paste_policy に従い warn(既定。登録を続ける)
'   /block(従来どおり止める)を切り替える。policy_no**だけ**の検知は
'   このキーに関わらず常に警告のみ(裁定書38 Z-46・変更なし)。
'   modPii だけを呼ぶ純関数なので modTestsPureNavi が直接叩ける
'   (config読取・ログ・保存はActPasteMaterial/ImportDirectPastes側の責務)。
' ----------------------------------------------------------------------------
' PiiPasteWarnText - 検知結果から利用者向けの案内文を1本作る。
'   isBlocked: True=登録を止める(呼び出し側はここでブロックしてE0103を記録)。
'              False=登録は続ける(呼び出し側は警告として保存する)。
'   検知なし("")のときは isBlocked=False・戻り値も "" (呼び出し側は何もしない)。
Public Function PiiPasteWarnText(ByVal bodyText As String, ByVal blockPolicy As Boolean, _
                                 ByRef isBlocked As Boolean) As String
    isBlocked = False
    Dim kinds As String
    kinds = modPii.KindsOf(bodyText)
    If LenB(kinds) = 0 Then Exit Function

    If kinds = "policy_no" Then
        ' 裁定書38 Z-46(既存挙動・変更なし)。文言も変えない。
        PiiPasteWarnText = "契約番号らしき数字列(" & CStr(modPii.DetectionCount(bodyText)) & _
            "箇所: " & modUtil.SafeLeft(bodyText, 30) & "…)を検知しました。" & _
            "伏せ字にするか、そのままでよいか確認してください。"
        Exit Function
    End If

    Dim breakdown As String
    breakdown = PiiBreakdownJa(bodyText)
    If blockPolicy Then
        isBlocked = True
        PiiPasteWarnText = "個人情報らしき記述を検知したため登録しませんでした（" & breakdown & _
            "）。伏せ字にして登録してください。"
    Else
        PiiPasteWarnText = "個人情報らしき記述を検知しました（" & breakdown & _
            "）。公開情報ならそのままで構いません。個人の連絡先なら伏せ字にしてから" & _
            "登録し直してください。登録は完了しています。"
    End If
End Function

' 裁定書44 B-1: policy_no以外(person/email/phone)を1件でも含み、かつ
'   pii_paste_policy=blockのときだけ止める(既定warnは止めない。policy_no
'   単独は裁定書38 Z-46のまま常に止めない)。ActPasteMaterialと
'   modUICase7.ImportDirectPastesの両方が同じ口を呼ぶ(片方だけ直さない)。
'   敵対的検証: "And blockPolicy"を外して常時Trueへ戻すとNAVI-U7-01/02の
'   いずれかが赤くなる。Excelトークンを持たない純関数。
Public Function BlocksGeneralPii(ByVal piiKinds As String, ByVal blockPolicy As Boolean) As Boolean
    BlocksGeneralPii = (LenB(piiKinds) > 0 And piiKinds <> "policy_no" And blockPolicy)
End Function

' PiiBreakdownJa - "人名らしき語 2件: 山田太郎様、佐藤様／電話番号 1件: 053-4xx-…"
'   の形を組む。modPii.SnippetsOf(先頭500件=PII_MAX_SPANS相当)から、種別ごとの
'   **全体件数**と、検知順で全体3件までの断片を作る(件数は全体、断片は3件が
'   予算。裁定書44 B-1「先頭3件の断片」)。
Private Function PiiBreakdownJa(ByVal bodyText As String) As String
    Const NA_PII_SHOW_ITEMS As Long = 3
    Const NA_PII_SNIP_LEN As Long = 12
    Const NA_PII_SCAN_MAX As Long = 500

    Dim allEntries As String
    allEntries = modPii.SnippetsOf(bodyText, NA_PII_SCAN_MAX, NA_PII_SNIP_LEN)
    If LenB(allEntries) = 0 Then Exit Function

    Dim rows() As String
    rows = Split(allEntries, ";")

    Dim kindsOrder(0 To 10) As String, counts(0 To 10) As Long, samples(0 To 10) As String
    Dim nKinds As Long, shownTotal As Long
    Dim i As Long, k As Long, idx As Long, tabPos As Long
    Dim kindText As String, snippet As String

    For i = LBound(rows) To UBound(rows)
        tabPos = InStr(1, rows(i), vbTab, vbBinaryCompare)
        If tabPos > 0 Then
            kindText = Left$(rows(i), tabPos - 1)
            snippet = Mid$(rows(i), tabPos + 1)
            idx = -1
            For k = 0 To nKinds - 1
                If kindsOrder(k) = kindText Then
                    idx = k
                    Exit For
                End If
            Next k
            If idx = -1 And nKinds <= UBound(kindsOrder) Then
                idx = nKinds
                kindsOrder(idx) = kindText
                nKinds = nKinds + 1
            End If
            If idx >= 0 Then
                counts(idx) = counts(idx) + 1
                If shownTotal < NA_PII_SHOW_ITEMS Then
                    If LenB(samples(idx)) > 0 Then samples(idx) = samples(idx) & "、"
                    samples(idx) = samples(idx) & snippet
                    shownTotal = shownTotal + 1
                End If
            End If
        End If
    Next i

    Dim acc As String
    For k = 0 To nKinds - 1
        If LenB(acc) > 0 Then acc = acc & "／"
        acc = acc & PiiKindJa(kindsOrder(k)) & " " & CStr(counts(k)) & "件"
        If LenB(samples(k)) > 0 Then acc = acc & ": " & samples(k)
    Next k
    PiiBreakdownJa = acc
End Function

' 検知種別(機械値)の日本語化(16章E-05注記どおりUI層の責務)。
Private Function PiiKindJa(ByVal kindText As String) As String
    Select Case kindText
        Case "person": PiiKindJa = "人名らしき語"
        Case "email": PiiKindJa = "メールアドレス"
        Case "phone": PiiKindJa = "電話番号"
        Case "policy_no": PiiKindJa = "契約番号らしき数字列"
        Case Else: PiiKindJa = kindText
    End Select
End Function

' CopySourceTextOf - 下見の下敷き(現契約サマリ・営業メモ・現場メモ)。
'   裁定書39 R1-06: 走査長に上限(先頭 NA_COPY_SCAN_MAX 字)を置く。
'   SharesLongFragment は O(|本文|×|下敷き|) なので、上限が無いと
'   3欄合計30万字の案件でコピーのたびに数十秒待たされる。
Private Function CopySourceTextOf(ByVal caseId As String) As String
    If LenB(caseId) = 0 Then Exit Function
    CopySourceTextOf = CapSourceText(modCaseStore.LoadData(caseId, "input_contract") & vbLf & _
                                     modCaseStore.LoadData(caseId, "input_memo") & vbLf & _
                                     modCaseStore.LoadData(caseId, "input_field_notes"))
End Function

' CapSourceText - 下見の下敷きを先頭 NA_COPY_SCAN_MAX 字で打ち切る純関数
'   (裁定書39 R1-06 の走査上限を層(a)から見えるところへ出した=裁定書40 の
'   横展開。上限を消しても機械で気づけないままにしない)。
Public Function CapSourceText(ByVal srcText As String) As String
    If Len(srcText) > NA_COPY_SCAN_MAX Then
        CapSourceText = Left$(srcText, NA_COPY_SCAN_MAX)
    Else
        CapSourceText = srcText
    End If
End Function

' OpenTargetOf - 出力一覧の[ブラウザで開く]/[フォルダを開く]で**開いてよいパス**を
'   決める純関数(裁定書39 R2-06 の防波堤を層(a)へ出した=裁定書40 の横展開)。
'   画面から来た givenPath は、案件に登録された2本(レポート・提案書)の
'   **どちらかと一致したときだけ**採る。一致しなければ空文字=開かない。
'   givenPath が空のとき(=一覧の行を指定しない呼び)はレポートを開く。
'   ここが崩れると、画面のJSを書き換えるだけで任意のファイルを開かせられる。
Public Function OpenTargetOf(ByVal reportPath As String, ByVal proposalPath As String, _
                             ByVal givenPath As String) As String
    If LenB(givenPath) = 0 Then
        OpenTargetOf = reportPath
        Exit Function
    End If
    If LenB(proposalPath) > 0 And givenPath = proposalPath Then
        OpenTargetOf = proposalPath
        Exit Function
    End If
    If LenB(reportPath) > 0 And givenPath = reportPath Then OpenTargetOf = reportPath
End Function

' OpenLabelOf - [ブラウザで開く]/[フォルダを開く]が返す文言の**出力の種類名**
'   (裁定書42 §2-2)。裁定書39 R2-06 で提案書の行も ActOpenReport を通るように
'   したのに、返る文言は「レポート」固定だった。区画④で「提案書(お客さま向け)」
'   の[ブラウザで開く]を押したのに「この案件に登録されたレポートがありません。」
'   と返るので、利用者は自分が何を開こうとしたのか分からなくなる(会社PCは
'   エクスプローラ操作が限られるため、この文言が唯一の手掛かり)。
'   判断の材料は OpenTargetOf と同じ2つ(案件の提案書パスと画面から来たパス)。
'   提案書の行から来たときだけ「提案書」、それ以外(行を指定しない呼び・
'   レポートの行・登録された2本のどちらとも一致しないパス)は「レポート」。
'   ヒアリングシートは出力一覧に[開く]を持たない(Excelのシートとして開く=
'   modNaviActions.ActExportHearing)ので、ここでは扱わない。
'   文言の値源はこの1本で、ActOpenReport の4つの文言すべてがこれを使う。
Public Function OpenLabelOf(ByVal proposalPath As String, ByVal givenPath As String) As String
    OpenLabelOf = "レポート"
    If LenB(proposalPath) > 0 And givenPath = proposalPath Then OpenLabelOf = "提案書"
End Function

' 区画①[調査ページを開く]/[クイック調査を開く](12章§2 の分割。裁定書39 R1-06 と同時)。
Public Function ActOpenUrl(ByVal data As String) As String
    Dim kind As String, url As String
    kind = modJsonLite.GetStr(data, "kind")
    Select Case kind
    Case "full", "quick", "menu"
        url = modUIResearch.DrUrlOf(kind, modConfig.GetStr("dr_url_" & kind, ""))
    Case "portal"
        url = modConfig.GetStr("portal_url", "")
    Case "help"
        ActOpenUrl = modNaviActions.ResultOf(modUISheet.ShowSheet("使い方"), "使い方を開きました。", "使い方を開けませんでした。")
        Exit Function
    Case Else
        ActOpenUrl = modNaviActions.Failure("このページは開けません。", "E0101")
        Exit Function
    End Select
    ActOpenUrl = modNaviActions.ResultOf(modUIResearch.OpenUrl(url), "ページを開きました。", "ページを開けませんでした。")
End Function

' ==========================================================
' 提案書(区画④)の出力判断(裁定書40 R-M2)。配線の実体は
' modNaviActions.ActExportProposal にあるが、**断る/断らないの判断と、
' 利用者へ返す文言・error_code はここの純関数に寄せる**(modNaviActions は
' 30,000字契約の残りが少なく、新しい関数を置けない=12章§2の分割先が本)。
' 層(a)の modTestsPureNavi がこの2本を直接叩く。
' ==========================================================

' ProposalBlockOf - [提案書（お客さま向け）を出す]を**出力させずに断る**ときの
'   応答(断らないときは空文字)。断る理由は2つだけで、順番にも意味がある。
'   (1) reviewNote が非空 = 確認者名が無い(16章 E-71 ④・20章§8-3)。S5 の
'       AI呼出を始める**前**に断る(作ってから断ると1回むだになる)。
'       error_code は E0603(裁定書40 R-m3: W15 Round2 で申告なく E0101 へ
'       変わっていたので戻した。16章 E-71 ④は「エラーではない」としか書いて
'       おらず code を登記していないので、登記の無い割当てを実装だけで
'       増やさない)。
'   (2) plan = "upstream_missing" = S1/S2/S3 のどれかが未了。error_code は
'       E0101(必須入力の欠落。16章 E-01。ActExportReport の同型と同じ)。
'   両方あてはまるときは (1) が勝つ(確認していない資料はそもそも出せない)。
Public Function ProposalBlockOf(ByVal reviewNote As String, ByVal plan As String) As String
    If LenB(reviewNote) > 0 Then
        ProposalBlockOf = modNaviActions.Failure(reviewNote, "E0603")
        Exit Function
    End If
    If plan = "upstream_missing" Then
        ProposalBlockOf = modNaviActions.Failure("先に[まとめて分析]を実行してください。", "E0101")
    End If
End Function

' ProposalS5FailedJson - S5(お客さま向け提案書の文章)が作れなかったときの応答
'   (16章 E-71 ②の逐語)。案件は壊れていない(生応答は s5_json_failed へ退避し
'   status も動かさない)ので、骨子(4. 提案の骨子)で商談へ行けることを案内する。
'   error_code は E-71 の表のとおり E0302(E-06 のまま。新しいコードは作らない)。
Public Function ProposalS5FailedJson() As String
    ProposalS5FailedJson = modNaviActions.Failure("提案書を作れませんでした。" & _
        "今回は提案骨子（4. 提案の骨子）をご利用ください。", "E0302")
End Function
