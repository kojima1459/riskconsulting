Attribute VB_Name = "modNaviHost"
Option Explicit
' Specifications 7.1-7.3, 8.2. Only this module owns the HTML request lifetime.
Private gForm As frmNaviHtml
Private gBusy As Boolean
Private gCurrent As String
Private gCloseAt As Date
Private gCloseScheduled As Boolean
Private gWorkbookClosing As Boolean
Private gCloseAllowed As Boolean
Public Sub OpenNaviTool()
    On Error GoTo Failed
    If gForm Is Nothing Then Set gForm = New frmNaviHtml
    gCloseAllowed = False
    gForm.Show 0
    gForm.EnsureBrowser
    Exit Sub
Failed:
    HostFailure "open", Err.Number
    DropForm
    Application.StatusBar = "HTML画面を開けませんでした。従来のナビ画面を使用してください。"
End Sub
Public Function HostIsBusy() As Boolean
    HostIsBusy = gBusy Or modUIProgress.IsUiLocked()
End Function
Public Function CurrentCaseId() As String
    CurrentCaseId = gCurrent
End Function
Public Function IsAllowed(ByVal action As String) As Boolean
    Select Case action
    Case "initialize", "open_case", "new_case", "save_basics", "copy_prompt", "open_url"
        IsAllowed = True
    Case "paste_material", "clear_material", "save_materials", "run_pipeline", "save_step_edit", "open_step_sheet"
        IsAllowed = True
    Case "export_report", "open_report", "export_hearing", "chat", "clear_chat", "report_mail"
        IsAllowed = True
    Case "sparring_resume", "sparring_send", "sparring_to_inbox", "start_round2", "company_save", "company_open"
        IsAllowed = True
    Case "feedback_add", "inbox_list", "inbox_post", "inbox_diagnose_all", "inbox_judge"
        IsAllowed = True
    Case "judge_list", "judge_add", "judge_result", "logs", "reload_kb", "run_tests"
        IsAllowed = True
    Case "rename_case", "archive_case", "export_case", "import_case", "save_settings", "resize", "close"
        IsAllowed = True
    End Select
End Function
Public Sub HostDispatchPending()
    Dim request As String, response As String
    On Error GoTo Failed
    If gForm Is Nothing Then Exit Sub
    ' Always consume and erase the DOM payload, including rejected re-entrant requests.
    request = gForm.TakePendingText()
    response = HostRequestJson(request)
    If Not gForm Is Nothing Then gForm.DeliverResponse response
    Exit Sub
Failed:
    HostFailure "dispatch_dom", Err.Number
End Sub
Public Function HostRequestJson(ByVal request As String) As String
    Dim action As String, data As String, response As String, requestedCase As String
    On Error GoTo Failed
    If HostIsBusy() Then
        HostRequestJson = "{""ok"":false,""busy"":true,""message"":" & modNaviJson.Q("処理中です") & "}"
        Exit Function
    End If
    If Not modNaviJson.IsValidJson(request) Then
        HostRequestJson = ErrorResponse("要求のJSONが不正です。", "E0101"): Exit Function
    End If
    action = modNaviJson.StringField(request, "action")
    If Not IsAllowed(action) Then
        HostRequestJson = ErrorResponse("この操作は受け付けません。", "E0101"): Exit Function
    End If
    data = modNaviJson.RawField(request, "data")
    If Left$(data, 1) <> "{" Then
        HostRequestJson = ErrorResponse("操作データが不正です。", "E0101"): Exit Function
    End If
    requestedCase = modNaviJson.StringField(data, "case_id")
    If Len(requestedCase) > 0 Then
        If Not modCaseStore.IsValidCaseId(requestedCase) Then
            HostRequestJson = ErrorResponse("案件IDが不正です。", "E0101"): Exit Function
        End If
        If Len(modCaseRead.CaseColumnOf(requestedCase, "case_id")) = 0 Then
            HostRequestJson = ErrorResponse("案件が見つかりません。", "E0101"): Exit Function
        End If
        If Len(modCaseRead.CaseColumnOf(requestedCase, "archived_at")) > 0 Then
            HostRequestJson = ErrorResponse("削除済みの案件は操作できません。", "E0101"): Exit Function
        End If
    End If
    gBusy = True
    Select Case action
    Case "initialize"
        response = "{""ok"":true,""message"":" & modNaviJson.Q("準備完了です。") & "}"
    Case "resize"
        If Not gForm Is Nothing Then gForm.CycleSize
        response = "{""ok"":true}"
    Case "close"
        response = modNaviActions.Dispatch(action, data, gCurrent)
        If modJsonLite.GetBoolJ(response, "ok", False) Then ScheduleClose
    Case Else
        response = modNaviActions.Dispatch(action, data, gCurrent)
    End Select
    If action = "save_settings" Then
        If Not gForm Is Nothing Then gForm.ApplyTextScale modConfig.GetStr("ui_font_scale", "medium")
    End If
    If Not modNaviJson.IsValidJson(response) Then Err.Raise 5, "modNaviHost", "invalid_action_response"
    response = NormalizeResponse(response)
    If action <> "close" Then
        response = Left$(Trim$(response), Len(Trim$(response)) - 1) & ",""state"":" & modNaviState.BuildAppState(gCurrent) & "}"
    End If
    gBusy = False
    HostRequestJson = response
    Exit Function
Failed:
    HostFailure "request", Err.Number
    gBusy = False
    HostRequestJson = ErrorResponse("処理に失敗しました。err_logタブの最後の行を開発担当へ送ってください。", "E0603")
End Function
Public Function ErrorResponse(ByVal message As String, ByVal code As String) As String
    ErrorResponse = "{""ok"":false,""kind"":""error"",""error_code"":" & modNaviJson.Q(code) & _
                    ",""message"":" & modNaviJson.Q(DisplayMessage(message)) & "}"
End Function
Private Function NormalizeResponse(ByVal response As String) As String
    Dim keys As Variant, k As Variant, v As String
    keys = Array("message", "warning", "confirm")
    For Each k In keys
        v = modNaviJson.StringField(response, CStr(k))
        If Len(v) > 0 Then response = modNaviJson.ReplaceTextField(response, CStr(k), DisplayMessage(v))
    Next k
    NormalizeResponse = response
End Function
' HTML画面が出す文言の言い換え(髙橋さん版 Specification 5.8)。
'   VBA 側が持っている業務文言だけを画面の語彙へ寄せる。お客様のテキストや
'   JSON は通さない(NormalizeResponse が message/warning/confirm の3つだけを通す)。
'   裁定書34 §0.4: 製品名は config app_display_name の1箇所が値源であり、
'   ここに "リスコンNavi" を焼かない。既定値のときは何も置き換わらない。
Public Function DisplayMessage(ByVal srcText As String) As String
    Dim source As Variant, target As Variant, i As Long
    Dim shownName As String
    srcText = Replace(srcText, "リスク提案ナビを起動", "★起動.bat")
    shownName = modBootNavi.AppDisplayName()
    If LenB(shownName) > 0 Then srcText = Replace(srcText, "リスク提案ナビ", shownName)
    source = Array("まとめて作る", "ヒアリングシートを出す", "レポートを出す", _
        "貼ったものを保存する", "貼る欄", "貼った資料", "貼ったもの", "作り直す", "作っています", _
        "作っています", "作ります", "貼ると", "貼る", "作る", "出す", "経路")
    target = Array("まとめて分析", "ヒアリングシート出力", "レポート出力", _
        "登録内容を保存", "資料登録欄", "登録資料", "登録内容", "再分析する", "分析しています", _
        "分析しています", "分析します", "登録すると", "登録", "分析", "出力", "AI接続")
    For i = LBound(source) To UBound(source)
        srcText = Replace(srcText, CStr(source(i)), CStr(target(i)))
    Next i
    DisplayMessage = srcText
End Function
Public Sub ShowBusy(ByVal message As String, ByVal progressJson As String)
    If gForm Is Nothing Then Exit Sub
    gForm.ShowBusy DisplayMessage(message), progressJson
    DoEvents
End Sub
Public Sub HostFailure(ByVal phase As String, ByVal errNo As Long)
    On Error GoTo Done
    modLog.LogError "E0603", "modNaviHost." & phase, "html_host_failed", errNo
Done:
End Sub
Public Function HostReadPage() As String
    Dim folder As String, html As String, css As String, js As String
    folder = modUtilPath.JoinPath(ThisWorkbook.Path, "ui")
    html = ReadAsset(folder, "index.html")
    css = ReadAsset(folder, "style.css")
    js = ReadAsset(folder, "markdown.js") & vbCrLf & ReadAsset(folder, "views.js") & vbCrLf & ReadAsset(folder, "app.js")
    html = RemoveExternalBlock(html, "STYLE")
    html = RemoveExternalBlock(html, "SCRIPT")
    If InStr(1, html, "<!--INLINE_STYLE-->", 0) = 0 Then Err.Raise 5, "HostReadPage", "missing_style_marker"
    If InStr(1, html, "<!--INLINE_SCRIPT-->", 0) = 0 Then Err.Raise 5, "HostReadPage", "missing_script_marker"
    html = Replace(html, "<!--INLINE_STYLE-->", "<style>" & css & "</style>")
    html = Replace(html, "<!--INLINE_SCRIPT-->", "<script>" & js & "</script>")
    HostReadPage = html
End Function
Private Function RemoveExternalBlock(ByVal html As String, ByVal kind As String) As String
    Dim first As Long, last As Long, closing As String
    first = InStr(1, html, "<!--EXTERNAL_" & kind & "_BEGIN-->", 0)
    closing = "<!--EXTERNAL_" & kind & "_END-->"
    last = InStr(1, html, closing, 0)
    If first > 0 And last > first Then html = Left$(html, first - 1) & Mid$(html, last + Len(closing))
    RemoveExternalBlock = html
End Function
Private Function ReadAsset(ByVal folder As String, ByVal srcName As String) As String
    Dim path As String
    path = modUtilPath.JoinPath(folder, srcName)
    If Not modUtil.FileExistsAt(path) Then Err.Raise 53, "HostReadPage", srcName
    ReadAsset = modUtil.ReadUtf8File(path)
    If Len(ReadAsset) = 0 Then Err.Raise 5, "HostReadPage", "empty_asset:" & srcName
End Function
Private Sub ScheduleClose()
    If gCloseScheduled Then Exit Sub
    gCloseAt = Now
    gCloseScheduled = True
    Application.OnTime gCloseAt, "CloseNaviTool"
End Sub
Public Sub CloseNaviTool()
    On Error GoTo Failed
    If HostIsBusy() Then gCloseScheduled = False: Exit Sub
    gCloseScheduled = False
    gCloseAllowed = True
    DropForm
    ThisWorkbook.Saved = True
    ThisWorkbook.Close False
    Exit Sub
Failed:
    HostFailure "close", Err.Number
End Sub
Public Function HostWorkbookClose() As Boolean
    If gCloseAllowed Then HostWorkbookClose = True: Exit Function
    If HostIsBusy() Then Exit Function
    If gForm Is Nothing Then HostWorkbookClose = True: Exit Function
    gWorkbookClosing = True
    gForm.RequestClose
End Function
Private Sub DropForm()
    On Error GoTo Done
    If Not gForm Is Nothing Then
        gForm.AllowClose = True
        Unload gForm
    End If
    Set gForm = Nothing
Done:
End Sub
Public Sub ShowSheet(ByVal sheetName As String)
    Dim ws As Object
    Set ws = modCaseStore2.SheetOf(sheetName)
    If ws Is Nothing Then Exit Sub
    ws.Visible = -1
    ws.Activate
    ' Keep the modeless window available; the sheet action is in Excel's task window.
End Sub
' Stage two extension points: intentionally no writes / no AI calls.
Public Function RoundDiffExtension(ByVal caseId As String) As String
End Function
Public Function EmbeddingSearchExtension(ByVal question As String) As String
End Function

' Read-only diagnostics for development verification, specification 12.
Public Function HostDiagnostics() As String
    Dim naviReady As Boolean, documentMode As Long, bodyChars As Long
    Dim fontScale As String, browserZoom As Long
    If Not gForm Is Nothing Then
        naviReady = gForm.IsReady: documentMode = gForm.DocumentMode: bodyChars = gForm.HtmlLength
        fontScale = gForm.TextScale: browserZoom = gForm.OpticalZoom
    End If
    HostDiagnostics = "{""ready"":" & modNaviJson.Flag(naviReady) & ",""document_mode"":" & CStr(documentMode) & _
        ",""html_length"":" & CStr(bodyChars) & ",""text_scale"":" & modNaviJson.Q(fontScale) & ",""zoom"":" & CStr(browserZoom) & "}"
End Function
