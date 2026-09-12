Attribute VB_Name = "modNaviStore"
Option Explicit

' err_log の出所(16章§1)。関数名をリテラルで散らさないための値源。
Private Const NST_SRC As String = "modNaviStore"

' [NAVI] Frontend specification 8.1 / 8.3.
' Thin adapters only. Business writes stay in existing stores.
Public Function Q(ByVal srcValue As String) As String
    Q = modNaviJson.Q(srcValue)
End Function

Public Function CaseExists(ByVal caseId As String) As Boolean
    Dim ws As Object, blk As Variant, r As Long
    If Not modCaseStore.IsValidCaseId(caseId) Then Exit Function
    CaseExists = modCaseStore2.LocateRow("案件一覧", caseId, ws, blk, r)
End Function

Public Function CellValue(ByVal blk As Variant, ByVal r As Long, ByVal srcName As String) As String
    Dim c As Long
    c = modUtil.FindHeaderCol(blk, srcName)
    If c > 0 Then CellValue = modKnowledge2.CellRaw(blk, r, c)
End Function

Public Function RowJson(ByVal blk As Variant, ByVal r As Long) As String
    Dim c As Long, key As String, buf() As String, n As Long
    modUtil.BufInit buf, n
    For c = LBound(blk, 2) To UBound(blk, 2)
        key = modKnowledge2.CellRaw(blk, 1, c)
        If LenB(key) > 0 Then
            If n > 0 Then modUtil.BufAdd buf, n, ","
            modUtil.BufAdd buf, n, Q(key) & ":" & Q(modKnowledge2.CellRaw(blk, r, c))
        End If
    Next c
    RowJson = "{" & JoinPieces(buf, n) & "}"
End Function

Public Function JoinPieces(ByRef buf() As String, ByVal n As Long) As String
    ' BufText inserts LF: legal JSON whitespace, but not safe inside scalar tokens.
    Dim i As Long, result As String
    For i = 0 To n - 1
        result = result & buf(i)
    Next i
    JoinPieces = result
End Function

Public Function RowsJson(ByVal sheetName As String, ByVal caseId As String) As String
    Dim filterName As String
    If sheetName = "判断台帳" Then
        filterName = "case_ref"
    Else
        filterName = "case_id"
    End If
    RowsJson = FilterRows(sheetName, filterName, caseId, False)
End Function

Public Function ListCases() As String
    ListCases = FilterRows("案件一覧", "", "", True)
End Function

Public Function ListInbox() As String
    ListInbox = FilterRows("受信箱", "", "", False)
End Function

Public Function ListJudgements(ByVal caseId As String) As String
    ListJudgements = RowsJson("判断台帳", caseId)
End Function

Public Function CaseRowJson(ByVal caseId As String) As String
    Dim ws As Object, blk As Variant, r As Long
    CaseRowJson = "{}"
    If modCaseStore2.LocateRow("案件一覧", caseId, ws, blk, r) Then CaseRowJson = RowJson(blk, r)
End Function

Private Function FilterRows(ByVal sheetName As String, ByVal filterName As String, _
                            ByVal caseId As String, ByVal omitArchived As Boolean) As String
    On Error GoTo Failed
    Dim ws As Object, blk As Variant, lastRow As Long, r As Long
    Dim take As Boolean, rowText As String, buf() As String, n As Long
    FilterRows = "[]"
    Set ws = modCaseStore2.SheetOf(sheetName)
    If ws Is Nothing Then Exit Function
    lastRow = modCaseStore2.LastRowOf(ws)
    blk = modCaseStore2.ReadBlock(ws, lastRow)
    If IsEmpty(blk) Then Exit Function
    modUtil.BufInit buf, n
    For r = 2 To lastRow
        take = True
        If LenB(caseId) > 0 And LenB(filterName) > 0 Then
            take = (CellValue(blk, r, filterName) = caseId)
        End If
        If sheetName = "案件一覧" Then
            If Not modCaseStore.IsValidCaseId(CellValue(blk, r, "case_id")) Then take = False
        ElseIf sheetName = "受信箱" Then
            If Not modInboxStore.IsValidInboxId(CellValue(blk, r, "inbox_id")) Then take = False
        End If
        If omitArchived Then
            If LenB(CellValue(blk, r, "archived_at")) > 0 Then take = False
        End If
        If take Then
            rowText = RowJson(blk, r)
            If n > 0 Then modUtil.BufAdd buf, n, ","
            modUtil.BufAdd buf, n, rowText
        End If
    Next r
    FilterRows = "[" & JoinPieces(buf, n) & "]"
    Exit Function
Failed:
    modLog.LogError "E0603", NST_SRC & ".FilterRows", sheetName, Err.Number
    FilterRows = "[]"
End Function

Public Function LogRowsOf(ByVal caseId As String) As String
    LogRowsOf = "{""run_log"":" & RowsJson("run_log", caseId) & _
                ",""err_log"":" & ErrorRowsOf(caseId) & _
                ",""usage_log"":" & RowsJson("usage_log", caseId) & "}"
End Function

' err_log has no case_id column. Associate rows two ways(裁定書36): detail に
' caseId を含む行(従来どおり) / その案件の run_log 行の run_at ±60秒の窓に
' logged_at が入る行(modPipeline・modGatewayRPN が書く原因行を拾うため)。
' run_log 行が1本もない案件は従来どおり detail 一致だけで判定する。
Private Function ErrorRowsOf(ByVal caseId As String) As String
    On Error GoTo Failed
    Dim ws As Object, blk As Variant, lastRow As Long, r As Long
    Dim buf() As String, n As Long
    Dim fromAt As String, toAt As String
    ErrorRowsOf = "[]"
    If LenB(caseId) = 0 Then Exit Function
    RunLogWindowOf caseId, fromAt, toAt
    Set ws = modCaseStore2.SheetOf("err_log")
    If ws Is Nothing Then Exit Function
    lastRow = modCaseStore2.LastRowOf(ws)
    blk = modCaseStore2.ReadBlock(ws, lastRow)
    If IsEmpty(blk) Then Exit Function
    modUtil.BufInit buf, n
    For r = 2 To lastRow
        If InStr(1, CellValue(blk, r, "detail"), caseId, 0) > 0 _
                Or InWindow(CellValue(blk, r, "logged_at"), fromAt, toAt) Then
            If n > 0 Then modUtil.BufAdd buf, n, ","
            modUtil.BufAdd buf, n, RowJson(blk, r)
        End If
    Next r
    ErrorRowsOf = "[" & JoinPieces(buf, n) & "]"
    Exit Function
Failed:
    ErrorRowsOf = "[]"
End Function

' RunLogWindowOf - 案件 caseId の run_log 行(13章§2.4: run_at / case_id 列)から
' run_at の最小値-60秒・最大値+60秒を求める(±60秒の加減は呼び出し側で
' DateAdd する方式を選択。理由は最終報告に記載)。行が無ければ両方空文字。
' 変換は Format$ ではなく modUtilText.IsoDateTime を使う(和暦端末対策。
' modUtil.NowStamp の既存注記と同じ理由)。
Private Sub RunLogWindowOf(ByVal caseId As String, ByRef fromAt As String, ByRef toAt As String)
    Dim ws As Object, blk As Variant, lastRow As Long, r As Long
    Dim runAt As String
    fromAt = vbNullString
    toAt = vbNullString
    Set ws = modCaseStore2.SheetOf("run_log")
    If ws Is Nothing Then Exit Sub
    lastRow = modCaseStore2.LastRowOf(ws)
    blk = modCaseStore2.ReadBlock(ws, lastRow)
    If IsEmpty(blk) Then Exit Sub
    For r = 2 To lastRow
        If CellValue(blk, r, "case_id") = caseId Then
            runAt = CellValue(blk, r, "run_at")
            If IsDate(runAt) Then
                If LenB(fromAt) = 0 Then
                    fromAt = runAt
                    toAt = runAt
                Else
                    If CDate(runAt) < CDate(fromAt) Then fromAt = runAt
                    If CDate(runAt) > CDate(toAt) Then toAt = runAt
                End If
            End If
        End If
    Next r
    If LenB(fromAt) > 0 Then
        fromAt = modUtilText.IsoDateTime(DateAdd("s", -60, CDate(fromAt)))
        toAt = modUtilText.IsoDateTime(DateAdd("s", 60, CDate(toAt)))
    End If
End Sub

' InWindow - loggedAt が [fromAt, toAt] の範囲内(両端含む)かどうか。
' すべて "yyyy-mm-dd hh:nn:ss" の文字列。IsDate で変換できないものは
' すべて False(不正日付は「窓に入らない」= fail-closed。裁定書36)。
Public Function InWindow(ByVal loggedAt As String, ByVal fromAt As String, ByVal toAt As String) As Boolean
    If LenB(loggedAt) = 0 Or LenB(fromAt) = 0 Or LenB(toAt) = 0 Then Exit Function
    If Not IsDate(loggedAt) Then Exit Function
    If Not IsDate(fromAt) Then Exit Function
    If Not IsDate(toAt) Then Exit Function
    InWindow = (CDate(loggedAt) >= CDate(fromAt)) And (CDate(loggedAt) <= CDate(toAt))
End Function

Public Function LastErrorJson() As String
    Dim ws As Object, blk As Variant, r As Long
    LastErrorJson = "{}"
    Set ws = modCaseStore2.SheetOf("err_log")
    If ws Is Nothing Then Exit Function
    r = modCaseStore2.LastRowOf(ws)
    If r < 2 Then Exit Function
    blk = modCaseStore2.ReadBlock(ws, r)
    If Not IsEmpty(blk) Then LastErrorJson = RowJson(blk, r)
End Function

Public Function SavedAt(ByVal caseId As String, ByVal dataKey As String) As String
    Dim ws As Object, blk As Variant, r As Long, lastRow As Long
    Set ws = modCaseStore2.SheetOf("case_data")
    If ws Is Nothing Then Exit Function
    lastRow = modCaseStore2.LastRowOf(ws)
    blk = modCaseStore2.ReadBlock(ws, lastRow)
    If IsEmpty(blk) Then Exit Function
    For r = 2 To lastRow
        If CellValue(blk, r, "case_id") = caseId And CellValue(blk, r, "data_key") = dataKey Then
            SavedAt = CellValue(blk, r, "saved_at")
        End If
    Next r
End Function

Public Function SetDisplayName(ByVal caseId As String, ByVal srcValue As String) As Boolean
    If Len(srcValue) > 200 Then Exit Function
    SetDisplayName = PutCaseField(caseId, "note", modUtilText.SanitizeInput(srcValue))
End Function

Public Function SetArchived(ByVal caseId As String) As Boolean
    ' Missing archived_at fails closed; import procedure adds the column.
    SetArchived = PutCaseField(caseId, "archived_at", modUtil.NowStamp())
End Function

Public Function PutCaseField(ByVal caseId As String, ByVal fieldName As String, _
                             ByVal srcValue As String) As Boolean
    On Error GoTo Failed
    Dim ws As Object, blk As Variant, r As Long
    If Not modCaseStore2.LocateRow("案件一覧", caseId, ws, blk, r) Then Exit Function
    If modUtil.FindHeaderCol(blk, fieldName) = 0 Then Exit Function
    modCaseStore2.PutText ws, blk, r, fieldName, srcValue
    modCaseStore2.PutText ws, blk, r, "updated_at", modUtil.NowStamp()
    PutCaseField = True
    Exit Function
Failed:
    modLog.LogError "E0603", NST_SRC & ".PutCaseField", caseId & ":" & fieldName, Err.Number
End Function

Public Function SaveBasics(ByVal caseId As String, ByVal data As String) As Boolean
    Dim keys() As String, i As Long, key As String, srcValue As String
    Dim basicJson As String, result As Boolean
    If Not CaseExists(caseId) Then Exit Function
    keys = Split("case_type;company;industry_code;industry_name", ";")
    result = True
    For i = LBound(keys) To UBound(keys)
        key = keys(i)
        If LenB(modNaviJson.RawField(data, key)) > 0 Then
            srcValue = modUtilText.SanitizeInput(modJsonLite.GetStr(data, key))
            If Not PutCaseField(caseId, key, srcValue) Then result = False
        End If
    Next i
    basicJson = MergeBasics(modCaseStore.LoadData(caseId, "nav_basics"), data)
    If Not modCaseStore.SaveData(caseId, "nav_basics", basicJson) Then result = False
    SaveBasics = result
End Function

Public Function MergeBasics(ByVal stored As String, ByVal data As String) As String
    Dim keys() As String, i As Long, key As String, srcValue As String, result As String
    keys = Split("address;sec_code;sites;copied_1;copied_2;copied_3;copied_4;copied_5;copied_6;copied_7;copied_8", ";")
    For i = LBound(keys) To UBound(keys)
        key = keys(i)
        If LenB(modNaviJson.RawField(data, key)) > 0 Then
            srcValue = modUtilText.SanitizeInput(modJsonLite.GetStr(data, key))
        Else
            srcValue = modJsonLite.GetStr(stored, key)
        End If
        If LenB(result) > 0 Then result = result & ","
        result = result & Q(key) & ":" & Q(srcValue)
    Next i
    MergeBasics = "{" & result & "}"
End Function

Public Function AppendFeedback(ByVal caseId As String, ByVal data As String) As Boolean
    On Error GoTo Failed
    Dim ws As Object, blk As Variant, r As Long, keys() As String, i As Long
    Dim body As String, key As String
    If Not CaseExists(caseId) Then Exit Function
    Set ws = modCaseStore2.SheetOf("フィードバック")
    If ws Is Nothing Then Exit Function
    r = modCaseStore2.LastRowOf(ws) + 1
    blk = modCaseStore2.ReadBlock(ws, r)
    If modUtil.FindHeaderCol(blk, "case_id") = 0 Then Exit Function
    keys = Split("visited_at;event;used_proposals;customer_quote;terms_summary;loss_note", ";")
    For i = LBound(keys) To UBound(keys)
        key = keys(i)
        body = modUtilText.SanitizeInput(modJsonLite.GetStr(data, key))
        If modJsonLite.GetBoolJ(data, "mask_pii", True) Then body = modPii.MaskText(body)
        modCaseStore2.PutText ws, blk, r, key, body
    Next i
    modCaseStore2.PutText ws, blk, r, "case_id", caseId
    modCaseStore2.PutText ws, blk, r, "recorded_by", modCaseStore2.OwnerName()
    modCaseStore2.PutText ws, blk, r, "recorded_at", modUtil.NowStamp()
    AppendFeedback = True
    Exit Function
Failed:
    modLog.LogError "E0603", NST_SRC & ".AppendFeedback", caseId, Err.Number
End Function

Public Function ExportCaseJson(ByVal caseId As String) As String
    Dim keys() As String, i As Long, rowText As String, buf() As String, n As Long
    If Not CaseExists(caseId) Then Exit Function
    keys = Split(modCaseStore3.DataKeys(), ";")
    modUtil.BufInit buf, n
    For i = LBound(keys) To UBound(keys)
        If LenB(keys(i)) > 0 Then
            rowText = "{""key"":" & Q(keys(i)) & ",""content"":" & _
                      Q(modCaseStore.LoadData(caseId, keys(i))) & "}"
            If n > 0 Then modUtil.BufAdd buf, n, ","
            modUtil.BufAdd buf, n, rowText
        End If
    Next i
    ExportCaseJson = "{""format"":""riscon-navi-case"",""version"":1,""case"":" & _
                     CaseRowJson(caseId) & ",""data"":[" & JoinPieces(buf, n) & "]}"
End Function

Public Function ValidTransfer(ByVal json As String, ByRef reason As String) As Boolean
    Dim ctx As String, items As Collection, item As Variant, key As String, seen As String
    reason = ""
    If Not modNaviJson.IsValidJson(json) Then
        reason = "案件JSONの構文が不正です。"
        Exit Function
    End If
    If modJsonLite.GetStr(json, "format") <> "riscon-navi-case" Or _
       modJsonLite.GetLong(json, "version", 0) <> 1 Then
        reason = "案件JSONの形式または版が一致しません。"
        Exit Function
    End If
    ctx = modNaviJson.ObjectField(json, "case")
    If LenB(Trim$(modJsonLite.GetStr(ctx, "company"))) = 0 Then
        reason = "会社名がありません。"
        Exit Function
    End If
    key = modJsonLite.GetStr(ctx, "case_type")
    If key <> "new" And key <> "renewal" Then
        reason = "種別が不正です。"
        Exit Function
    End If
    Set items = modJsonLite.GetArrayItems(json, "data")
    For Each item In items
        key = modJsonLite.GetStr(CStr(item), "key")
        If LenB(key) = 0 Or InStr(1, ";" & modCaseStore3.DataKeys() & ";", ";" & key & ";", 0) = 0 Then
            reason = "許可されていない保存キーです。"
            Exit Function
        End If
        If InStr(1, seen, ";" & key & ";", 0) > 0 Then
            reason = "保存キーが重複しています。"
            Exit Function
        End If
        seen = seen & ";" & key & ";"
    Next item
    If items.Count = 0 Then
        reason = "案件データがありません。"
        Exit Function
    End If
    ValidTransfer = True
End Function

Public Function ImportCaseJson(ByVal json As String, ByRef reason As String) As String
    On Error GoTo Failed
    Dim ctx As String, newId As String, items As Collection, item As Variant
    Dim keys() As String, i As Long, key As String, pairs As String, srcValue As String
    If Not ValidTransfer(json, reason) Then Exit Function
    ctx = modNaviJson.ObjectField(json, "case")
    newId = modCaseStore.NewCase(modJsonLite.GetStr(ctx, "company"), _
             modJsonLite.GetStr(ctx, "industry_code"), modJsonLite.GetStr(ctx, "case_type"))
    If LenB(newId) = 0 Then
        reason = "案件の採番ができませんでした。"
        Exit Function
    End If
    Set items = modJsonLite.GetArrayItems(json, "data")
    For Each item In items
        If Not modCaseStore.SaveData(newId, modJsonLite.GetStr(CStr(item), "key"), _
                                     modJsonLite.GetStr(CStr(item), "content")) Then GoTo FailedImport
    Next item
    ' Paths, IDs, owner and timestamps are local. Import cannot open a remote path.
    keys = Split("industry_name;channel;kanji;bid;reins;other_insurers;dossier_tier;note;s4_variant;" & _
                 "round_no;last_ok_step;failed_step;adopted_story_nos;focus_line_ids", ";")
    For i = LBound(keys) To UBound(keys)
        key = keys(i)
        srcValue = modJsonLite.GetStr(ctx, key)
        If InStr(srcValue, vbTab) > 0 Or InStr(srcValue, vbLf) > 0 Or InStr(srcValue, vbCr) > 0 Then
            ' Pair format has no escape; use direct checked field adapter for free text.
            If Not PutCaseField(newId, key, srcValue) Then GoTo FailedImport
        ElseIf LenB(srcValue) > 0 Then
            pairs = pairs & key & vbTab & srcValue & vbLf
        End If
    Next i
    If Not modCaseStore3.UpsertCaseRow(newId, pairs) Then GoTo FailedImport
    srcValue = modJsonLite.GetStr(ctx, "status")
    If srcValue = "exported" Or srcValue = "feedback_done" Then srcValue = "s4_done"
    If Not modCaseStore.SetStatus(newId, srcValue) Then GoTo FailedImport
    ImportCaseJson = newId
    Exit Function
FailedImport:
    reason = "取り込みが完了しませんでした。途中の案件は非表示にしました。"
    SetArchived newId
    Exit Function
Failed:
    modLog.LogError "E0603", NST_SRC & ".ImportCaseJson", newId, Err.Number
    reason = "案件JSONを取り込めませんでした。"
    If LenB(newId) > 0 Then HideIncompleteImport newId
End Function

Private Sub HideIncompleteImport(ByVal caseId As String)
    On Error GoTo Done
    SetArchived caseId
Done:
End Sub

Public Function SaveSettings(ByVal data As String, ByRef reason As String) As Boolean
    Dim fontScale As String, mode As String, inc As Boolean, dirPath As String, path As String
    Dim original As String, body As String
    fontScale = modJsonLite.GetStr(data, "ui_font_scale")
    mode = modJsonLite.GetStr(data, "ui_mode")
    inc = modJsonLite.GetBoolJ(data, "chat_include_materials", False)
    If fontScale <> "small" And fontScale <> "medium" And fontScale <> "large" Then
        reason = "表示倍率が不正です。"
        Exit Function
    End If
    If mode <> "html" And mode <> "sheet" Then
        reason = "画面モードが不正です。"
        Exit Function
    End If
    dirPath = modUtil.ResolveDataDir(modConfig.GetStr("data_dir", ""), ThisWorkbook.Path)
    path = modUtilPath.JoinPath(dirPath, "設定.txt")
    If modUtil.FileExistsAt(path) Then original = modUtil.ReadUtf8File(path)
    body = original & vbCrLf & "ui_font_scale=" & fontScale & vbCrLf & "ui_mode=" & mode & _
           vbCrLf & "chat_include_materials=" & UCase$(CStr(inc)) & vbCrLf
    If Not modUtil.WriteUtf8File(path, body, True) Then
        reason = "設定.txtへ保存できませんでした。"
        Exit Function
    End If
    If Not modConfig.SetValue("ui_font_scale", fontScale) Then Exit Function
    If Not modConfig.SetValue("ui_mode", mode) Then Exit Function
    If Not modConfig.SetValue("chat_include_materials", UCase$(CStr(inc))) Then Exit Function
    SaveSettings = True
End Function

' @unused: Phase2 予約(裁定書38)
Public Function BuildRoundDiff(ByVal caseId As String) As String
    ' Specification F-13: reserved for phase 2. Do not parse S1-S4 in VBA.
End Function

' @unused: Phase2 予約(裁定書38)
Public Function FindSimilarByEmbedding(ByVal question As String) As String
    ' Phase 2 extension only. No embedding API call in phase 1.
End Function


Public Function HasCaseData(ByVal caseId As String) As Boolean
    Dim ws As Object, blk As Variant, r As Long, lastRow As Long
    Set ws = modCaseStore2.SheetOf("case_data")
    If ws Is Nothing Then Exit Function
    lastRow = modCaseStore2.LastRowOf(ws)
    blk = modCaseStore2.ReadBlock(ws, lastRow)
    If IsEmpty(blk) Then Exit Function
    For r = 2 To lastRow
        If CellValue(blk, r, "case_id") = caseId Then
            HasCaseData = True
            Exit Function
        End If
    Next r
End Function

Public Function HydrateCaseIfNeeded(ByVal caseId As String, ByRef reason As String) As Boolean
    ' Spec 7.2 open_case; A 1.5 / 1.10. Boot reconstructs headers only.
    ' Only an empty local cache is populated; intentional local clears are never refilled.
    On Error GoTo Failed
    Dim path As String, company As String, WB As Object, ws As Object, blk As Variant
    Dim r As Long, lastRow As Long, matchRow As Long, n As Long
    Dim mainWs As Object, mainBlk As Variant, mainRow As Long, c As Long, key As String
    reason = ""
    HydrateCaseIfNeeded = False
    If Not CaseExists(caseId) Then Exit Function
    If HasCaseData(caseId) Then
        HydrateCaseIfNeeded = True
        Exit Function
    End If
    company = modCaseRead.CaseColumnOf(caseId, "company")
    path = modUtilPath.JoinPath(modCompanyFile3.CompanyDir(), modCompanyFile3.CompanyFileNameOf(company))
    If Not modUtil.FileExistsAt(path) Then
        HydrateCaseIfNeeded = True
        Exit Function
    End If
    Set WB = modCompanyFile2.DossierOpen(path, True)
    If WB Is Nothing Then
        reason = "企業ファイルを読み込めませんでした。ファイルの利用状況を確認してください。"
        Exit Function
    End If
    Set ws = modCompanyFile2.DossierSheet(WB, "dossier_case")
    If ws Is Nothing Then
        reason = "企業ファイルに案件IDの見出しがありません。"
        GoTo Done
    End If
    lastRow = modCompanyFile2.SheetLastRow(ws)
    blk = modCompanyFile2.SheetBlock(ws, lastRow)
    If IsEmpty(blk) Then GoTo Done
    For r = 2 To lastRow
        If CellValue(blk, r, "case_id") = caseId Then matchRow = r
    Next r
    If matchRow = 0 Then
        reason = "企業ファイル内に同じ案件IDがありません。別の案件は取り込みませんでした。"
        GoTo Done
    End If
    If Not modCompanyFile3.IsSchemaReadable(CellValue(blk, matchRow, "schema_version")) Then
        reason = "この企業ファイルの版には対応していません。"
        GoTo Done
    End If
    If Not modCaseStore2.LocateRow("案件一覧", caseId, mainWs, mainBlk, mainRow) Then GoTo Done
    For c = LBound(mainBlk, 2) To UBound(mainBlk, 2)
        key = modKnowledge2.CellRaw(mainBlk, 1, c)
        If LenB(key) > 0 And key <> "case_id" Then
            If modUtil.FindHeaderCol(blk, key) > 0 Then
                modCaseStore2.PutText mainWs, mainBlk, mainRow, key, CellValue(blk, matchRow, key)
            End If
        End If
    Next c
    ' ImportExtensions addresses dossier_data by exact case_id and never overwrites nonempty local data.
    n = modCompanyFile3.ImportExtensions(WB, caseId)
    RestoreCompanyRecords WB, "dossier_facts", "フィードバック", "case_id", caseId
    RestoreCompanyRecords WB, "dossier_judge", "判断台帳", "case_ref", caseId
    HydrateCaseIfNeeded = True
Done:
    modCompanyFile2.DossierClose WB
    Set WB = Nothing
    If Not HydrateCaseIfNeeded And LenB(reason) = 0 Then reason = "案件の内容を読み込めませんでした。"
    Exit Function
Failed:
    reason = "案件の内容を読み込めませんでした。"
    modLog.LogError "E0603", NST_SRC & ".HydrateCaseIfNeeded", caseId, Err.Number
    CloseReadBook WB
End Function

Private Sub CloseReadBook(ByVal WB As Object)
    On Error GoTo Done
    If Not WB Is Nothing Then modCompanyFile2.DossierClose WB
Done:
End Sub

Private Sub RestoreCompanyRecords(ByVal WB As Object, ByVal sourceName As String, _
                                  ByVal targetName As String, ByVal targetKey As String, ByVal caseId As String)
    ' An existing local record set wins as a whole. Never duplicate user-edited rows.
    Dim sourceWs As Object, targetWs As Object, sourceBlk As Variant, targetBlk As Variant
    Dim sourceLast As Long, targetLast As Long, r As Long, c As Long, key As String, wr As Long
    Set sourceWs = modCompanyFile2.DossierSheet(WB, sourceName)
    Set targetWs = modCaseStore2.SheetOf(targetName)
    If sourceWs Is Nothing Or targetWs Is Nothing Then Exit Sub
    targetLast = modCaseStore2.LastRowOf(targetWs)
    targetBlk = modCaseStore2.ReadBlock(targetWs, targetLast)
    If IsEmpty(targetBlk) Then Exit Sub
    For r = 2 To targetLast
        If CellValue(targetBlk, r, targetKey) = caseId Then Exit Sub
    Next r
    sourceLast = modCompanyFile2.SheetLastRow(sourceWs)
    sourceBlk = modCompanyFile2.SheetBlock(sourceWs, sourceLast)
    If IsEmpty(sourceBlk) Then Exit Sub
    wr = targetLast
    For r = 2 To sourceLast
        If CellValue(sourceBlk, r, "case_id") = caseId Then
            wr = wr + 1
            For c = LBound(targetBlk, 2) To UBound(targetBlk, 2)
                key = modKnowledge2.CellRaw(targetBlk, 1, c)
                If LenB(key) > 0 Then
                    modCaseStore2.PutText targetWs, targetBlk, wr, key, CellValue(sourceBlk, r, key)
                End If
            Next c
        End If
    Next r
End Sub
