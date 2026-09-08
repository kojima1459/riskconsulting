Attribute VB_Name = "modNaviState2"
Option Explicit
' Specification 7.4: read-only formatting adapters; no interpretation of Step JSON.
Public Function CaseContext(ByVal caseId As String) As String
    Dim headers As Variant, h As Variant, result As String, srcValue As String
    headers = Split("case_id case_type dossier_tier channel kanji bid reins other_insurers company industry_code industry_name status round_no last_ok_step failed_step updated_at report_path note archived_at", " ")
    result = "{"
    For Each h In headers
        If Len(result) > 1 Then result = result & ","
        srcValue = modCaseRead.CaseColumnOf(caseId, CStr(h))
        result = result & modNaviJson.Q(CStr(h)) & ":" & modNaviJson.Q(srcValue)
    Next h
    CaseContext = result & "}"
End Function
Public Function MaterialsJson(ByVal caseId As String, ByVal company As String, ByRef anyArea As Boolean) As String
    Dim rows As Variant, row As Variant, f As Variant, raw As String, restored As String
    Dim result As String, preview As String, one As Variant, extra As String
    rows = Split(modUICase6.AreaTable(), vbLf)
    result = "{": anyArea = False
    For Each row In rows
        f = Split(CStr(row), "|")
        raw = "": extra = ""
        If Len(caseId) > 0 Then
            raw = modCaseStore.LoadData(caseId, CStr(f(1)))
            If CStr(f(0)) = "field_notes" Then
                extra = ",""memo"":" & modNaviJson.Q(modUICase.RestoreNames(raw, company))
                raw = "【営業メモ】" & vbLf & raw
                raw = raw & vbLf & "【そのほか】" & vbLf & modCaseStore.LoadData(caseId, "input_field_notes")
                raw = raw & vbLf & "【付保の見立て】" & vbLf & modCaseStore.LoadData(caseId, "input_coverage_note")
                extra = extra & ",""field_notes"":" & modNaviJson.Q(modUICase.RestoreNames(modCaseStore.LoadData(caseId, "input_field_notes"), company))
                extra = extra & ",""coverage_note"":" & modNaviJson.Q(modUICase.RestoreNames(modCaseStore.LoadData(caseId, "input_coverage_note"), company))
                If Len(modCaseStore.LoadData(caseId, "input_memo")) + Len(modCaseStore.LoadData(caseId, "input_field_notes")) + Len(modCaseStore.LoadData(caseId, "input_coverage_note")) = 0 Then raw = ""
            End If
        End If
        If Len(raw) > 0 Then anyArea = True
        restored = modUICase.RestoreNames(raw, company)
        preview = "["
        For Each one In Split(modNavText.PreviewLines(restored, 3), vbLf)
            If Len(preview) > 1 Then preview = preview & ","
            preview = preview & modNaviJson.Q(CStr(one))
        Next one
        If Len(result) > 1 Then result = result & ","
        result = result & modNaviJson.Q(CStr(f(0))) & ":{""label"":" & modNaviJson.Q(CStr(f(2))) & ",""chars"":" & CStr(Len(raw)) & _
            ",""text"":" & modNaviJson.Q(restored) & ",""ai_text"":" & modNaviJson.Q(raw) & ",""preview"":" & preview & "],""pasted_at"":" & modNaviJson.Q(SavedAt(caseId, CStr(f(1)))) & extra & "}"
    Next row
    MaterialsJson = result & "}"
End Function
Public Function SavedAt(ByVal caseId As String, ByVal dataKey As String) As String
    Dim ws As Object, blk As Variant, r As Long, c As Long, k As Long, t As Long
    If Len(caseId) = 0 Then Exit Function
    Set ws = modCaseStore2.SheetOf("case_data")
    If ws Is Nothing Then Exit Function
    blk = modCaseStore2.ReadBlock(ws, modCaseStore2.LastRowOf(ws))
    If Not IsArray(blk) Then Exit Function
    c = modUtil.FindHeaderCol(blk, "case_id"): k = modUtil.FindHeaderCol(blk, "data_key")
    t = modUtil.FindHeaderCol(blk, "saved_at")
    If c = 0 Or k = 0 Or t = 0 Then Exit Function
    For r = 2 To UBound(blk, 1)
        If CStr(blk(r, c)) = caseId And CStr(blk(r, k)) = dataKey Then SavedAt = CStr(blk(r, t))
    Next r
End Function
Public Function SparringJson(ByVal caseId As String, ByVal company As String) As String
    SparringJson = modNaviChat.MergeHistoryJson(modSparring.HistoryOf(caseId, "user"), _
                                               modSparring.HistoryOf(caseId, "ai"), company)
End Function
Public Function BuildEnums() As String
    Dim groups As Variant, g As Variant, rows As Variant, row As Variant, f As Variant
    Dim result As String, pairs As String
    groups = Split(modUICase.EnumGroups(), ";")
    rows = Split(modUICase.EnumPairsCsv(), vbLf)
    result = "{"
    For Each g In groups
        If Len(CStr(g)) > 0 Then
            pairs = "["
            For Each row In rows
                f = Split(CStr(row), ",")
                If UBound(f) >= 2 Then
                    If CStr(f(0)) = CStr(g) Then
                        If Len(pairs) > 1 Then pairs = pairs & ","
                        pairs = pairs & "[" & modNaviJson.Q(CStr(f(1))) & "," & modNaviJson.Q(CStr(f(2))) & "]"
                    End If
                End If
            Next row
            If Len(result) > 1 Then result = result & ","
            result = result & modNaviJson.Q(CStr(g)) & ":" & pairs & "]"
        End If
    Next g
    BuildEnums = result & "}"
End Function
Public Function BuildIndustries() As String
    Dim codes As Object, names As Object, i As Long, result As String
    On Error GoTo Missing
    Set codes = ThisWorkbook.Names("enum_industry_code").RefersToRange
    Set names = ThisWorkbook.Names("enum_industry_name").RefersToRange
    result = "["
    For i = 1 To codes.Cells.Count
        If i > names.Cells.Count Then Exit For
        If Len(CStr(codes.Cells(i).Value2)) > 0 Then
            If Len(result) > 1 Then result = result & ","
            result = result & "[" & modNaviJson.Q(CStr(codes.Cells(i).Value2)) & "," & modNaviJson.Q(CStr(names.Cells(i).Value2)) & "]"
        End If
    Next i
    BuildIndustries = result & "]"
    Exit Function
Missing:
    BuildIndustries = "[]"
End Function
Public Function HearingBuiltAt(ByVal caseId As String) As String
    Dim rows As Collection, row As Variant
    If Len(caseId) = 0 Then Exit Function
    Set rows = modJsonLite.GetArrayItems("{""items"":" & modNaviStore.RowsJson("usage_log", caseId) & "}", "items")
    For Each row In rows
        If modJsonLite.GetStr(CStr(row), "event") = "hearing_sheet" Then
            HearingBuiltAt = modJsonLite.GetStr(CStr(row), "logged_at")
        End If
    Next row
End Function
