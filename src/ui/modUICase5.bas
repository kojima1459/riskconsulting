Attribute VB_Name = "modUICase5"
Option Explicit

' ============================================================================
' modUICase5 - S1～S4シートの逆シリアライズ(読む側)とブロックの幾何(ui層・T-31)
' ----------------------------------------------------------------------------
' 30,000字契約(12章§2)による modUICase2 の分割先(裁定書9 W4.1 の司令塔裁定)。
' **呼んでよいのは modUICase2** だけであり、変換表(19章§3)と列定義(13章§2.12-
' §2.15)は持たない(modUICase / modUICaseFmt が唯一持つ)。
'
' 本モジュールの責務は2つだけ:
'   (1) シート -> JSON の逆シリアライズ本体(13章§2.2 規約1/2/3/5)。入口は
'       SerializeBody で、失敗時のE-code記録と "" 返しは呼び出し側
'       (modUICase2.SerializeStep)が持つ(捏造しない口を1本に保つ)。
'   (2) ブロックの幾何(ColIndexes / ColCount / RoomOf)。描画側(modUICase2)と
'       読取側(本モジュール)が**同じ1本**を呼ぶことで、列の引き当て方と部屋の
'       数え方が2箇所へ分かれないようにする。
'
' 部屋(RoomOf)は13章§2.9が「行番号を仮定しない」と定めるため、次のブロックの
' アンカー行から動的に決める(確保行数を定数で持たない)。
' ============================================================================

Private Const U5_MAX_ROOM As Long = 200        ' 最終ブロックの部屋の上限
Private Const U5_HDR_WIDTH As Long = 24        ' 見出し行の探索幅(列番号ではない)
Private Const U5_FIRST_COL As Long = 1         ' ブロックの左端列(build/sheets_main.json)

' ============================================================================
' SerializeBody - シート -> JSON 本体(13章§2.2)。組めなければ ""。
'   例外の捕捉と E0302 の記録は呼び出し側 modUICase2.SerializeStep が持つ。
' ============================================================================
Public Function SerializeBody(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 1
        SerializeBody = SerializeS1()
    Case 2
        SerializeBody = SerializeS2()
    Case 3
        SerializeBody = SerializeS3()
    Case 4
        SerializeBody = SerializeS4()
    End Select
End Function

' S1: 15章 SchemaS1 のプロパティ順に組む(13章§2.2 規約1)。
Private Function SerializeS1() As String
    Dim basicSpec As String
    basicSpec = modUICaseFmt.ColsS1Basic()

    Dim vals As Variant
    vals = ReadSingleRow("s1_basic", basicSpec)
    If IsEmpty(vals) Then Exit Function

    Dim s As String
    s = ""
    modUICaseFmt.AddFrag s, "company_name", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "company_name"))
    modUICaseFmt.AddFrag s, "business_summary", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "business_summary"))
    modUICaseFmt.AddFrag s, "main_products", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "main_products"))
    modUICaseFmt.AddFrag s, "processes", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "processes"))
    modUICaseFmt.AddFrag s, "locations", _
            BlockArrJson("s1_locations", modUICaseFmt.ColsS1Locations())

    Dim sub1 As String
    sub1 = ""
    modUICaseFmt.AddFrag sub1, "key_materials", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "supply_chain_key_materials"))
    modUICaseFmt.AddFrag sub1, "notes", modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "supply_chain_notes"))
    modUICaseFmt.AddFrag s, "supply_chain", "{" & sub1 & "}"

    sub1 = ""
    modUICaseFmt.AddFrag sub1, "segments", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "customers_segments"))
    modUICaseFmt.AddFrag sub1, "channels", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "customers_channels"))
    modUICaseFmt.AddFrag s, "customers", "{" & sub1 & "}"

    modUICaseFmt.AddFrag s, "workforce_notes", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "workforce_notes"))
    modUICaseFmt.AddFrag s, "management_notes", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "management_notes"))

    sub1 = ""
    modUICaseFmt.AddFrag sub1, "mvv", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "mvv"))
    modUICaseFmt.AddFrag sub1, "aspirations", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "aspirations"))
    modUICaseFmt.AddFrag sub1, "market_context", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "market_context"))
    modUICaseFmt.AddFrag s, "strategy_outlook", "{" & sub1 & "}"

    modUICaseFmt.AddFrag s, "current_coverage", _
            BlockArrJson("s1_current_coverage", modUICaseFmt.ColsS1Coverage())
    modUICaseFmt.AddFrag s, "field_insights", _
            BlockArrJson("s1_field_insights", modUICaseFmt.ColsS1Insights())
    modUICaseFmt.AddFrag s, "missing_info", _
            BlockArrJson("s1_missing_info", modUICaseFmt.ColsS1Missing())

    sub1 = ""
    modUICaseFmt.AddFrag sub1, "coverage", _
            BlockArrJson("s1_input_quality", modUICaseFmt.ColsS1Quality())
    modUICaseFmt.AddFrag sub1, "overall", _
            modUICaseFmt.EnumJson("input_quality_overall", ValueOf(basicSpec, vals, "input_quality_overall"))
    modUICaseFmt.AddFrag sub1, "advice", modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "input_quality_advice"))
    modUICaseFmt.AddFrag s, "input_quality", "{" & sub1 & "}"

    modUICaseFmt.AddFrag s, "research_requests", _
            BlockArrJson("s1_research_requests", modUICaseFmt.ColsS1Research())
    SerializeS1 = "{" & s & "}"
End Function

Private Function SerializeS2() As String
    Dim s As String
    s = ""
    modUICaseFmt.AddFrag s, "risks", BlockArrJson("s2_risks", modUICaseFmt.ColsS2Risks())
    modUICaseFmt.AddFrag s, "gaps", BlockArrJson("s2_gaps", modUICaseFmt.ColsS2Gaps())
    modUICaseFmt.AddFrag s, "emerging_risks", BlockArrJson("s2_emerging", modUICaseFmt.ColsS2Emerging())
    modUICaseFmt.AddFrag s, "open_questions", ScalarArrJson("s2_open_questions", "question")
    SerializeS2 = "{" & s & "}"
End Function

Private Function SerializeS3() As String
    Dim s As String
    s = ""
    modUICaseFmt.AddFrag s, "stories", BlockArrJson("s3_stories", modUICaseFmt.ColsS3Stories())
    modUICaseFmt.AddFrag s, "unmatched_risks", BlockArrJson("s3_unmatched_risks", modUICaseFmt.ColsS3Unmatched())
    modUICaseFmt.AddFrag s, "do_not_propose", BlockArrJson("s3_do_not_propose", modUICaseFmt.ColsS3DoNot())
    ' 13章§2.14 v2.6: 15章 SchemaS3 は全プロパティ required なので、この2本を
    ' 落とすと sN_edited が CheckS3(V-S3-14 / V-S3-19)で必ず不合格になる。
    modUICaseFmt.AddFrag s, "growth_ideas", _
            BlockArrJson("s3_growth_ideas", modUICaseFmt.ColsS3Growth())
    modUICaseFmt.AddFrag s, "talk_script", _
            BlockObjJson("s3_talk_script", modUICaseFmt.ColsS3Talk())
    SerializeS3 = "{" & s & "}"
End Function

Private Function SerializeS4() As String
    Dim vals As Variant
    vals = ReadSingleRow("s4_meta", modUICaseFmt.ColsS4Meta())
    If IsEmpty(vals) Then Exit Function

    Dim s As String
    s = ""
    modUICaseFmt.AddFrag s, "file_title", modUICaseFmt.StrJson(ValueOf(modUICaseFmt.ColsS4Meta(), vals, "file_title"))
    modUICaseFmt.AddFrag s, "slides", BlockArrJson("s4_slides", modUICaseFmt.ColsS4Slides())
    modUICaseFmt.AddFrag s, "hearing_questions", BlockArrJson("s4_hearing_questions", modUICaseFmt.ColsS4Hearing())
    SerializeS4 = "{" & s & "}"
End Function

' 単一行ブロック(s3_talk_script)を1つのJSONオブジェクトにする(13章§2.14)。
'   行が読めなければ空のオブジェクト "{}" を返す(キーごと落とすと 15章 SchemaS3
'   の required を割るため、器だけは必ず出す)。
Private Function BlockObjJson(ByVal anchorName As String, ByVal colSpec As String) As String
    BlockObjJson = "{}"

    Dim vals As Variant
    vals = ReadSingleRow(anchorName, colSpec)
    If IsEmpty(vals) Then Exit Function

    BlockObjJson = "{" & modUICaseFmt.RowObjJson(colSpec, vals) & "}"
End Function

' 単一行ブロック(s1_basic / s4_meta)の1行を読む。読めなければ Empty。
Private Function ReadSingleRow(ByVal anchorName As String, ByVal colSpec As String) As Variant
    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Function

    Dim cols As Variant
    cols = ColIndexes(ws, headerRow, colSpec)
    If IsEmpty(cols) Then Exit Function

    ReadSingleRow = RowValues(ws, headerRow + 1, cols)
End Function

' 配列ブロックを読んでJSON配列本文にする。0行なら "[]"(13章§2.2 空配列規約)。
Private Function BlockArrJson(ByVal anchorName As String, ByVal colSpec As String) As String
    BlockArrJson = "[]"

    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Function

    Dim cols As Variant
    cols = ColIndexes(ws, headerRow, colSpec)
    If IsEmpty(cols) Then Exit Function

    Dim lastRow As Long
    lastRow = modUISheet.BlockLastRow(ws, headerRow, U5_FIRST_COL, _
                                      ColCount(colSpec), RoomOf(anchorName))

    ' 13章§2.2 逆シリアライズ規約2「行順=配列順。順序列の昇順に並べ替えてから
    ' 配列化する」。並べ替えの鍵は当該表が持つ順序列(risk_no / gap_no / story_no /
    ' slide_no / emg_no / seq)で、無ければ行順のまま。
    Dim keyCol As Long
    keyCol = modUICaseFmt.OrderColOf(colSpec)

    Dim items() As String
    Dim keys() As Long
    Dim n As Long
    ReDim items(0 To lastRow - headerRow)
    ReDim keys(0 To lastRow - headerRow)

    Dim r As Long
    Dim vals As Variant
    For r = headerRow + 1 To lastRow
        vals = RowValues(ws, r, cols)
        items(n) = modUICaseFmt.RowObjJson(colSpec, vals)
        If keyCol >= 0 Then
            keys(n) = CLng(Val(CStr(vals(keyCol + LBound(vals)))))
        Else
            keys(n) = n + 1
        End If
        n = n + 1
    Next r
    If n = 0 Then Exit Function

    modUICaseFmt.SortItems items, keys, n

    Dim acc As String
    Dim i As Long
    For i = 0 To n - 1
        If LenB(acc) > 0 Then acc = acc & ","
        acc = acc & items(i)
    Next i

    BlockArrJson = "[" & acc & "]"
End Function

' 文字列だけの配列ブロック(s2_open_questions)。
Private Function ScalarArrJson(ByVal anchorName As String, ByVal colName As String) As String
    ScalarArrJson = "[]"

    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Function

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, U5_FIRST_COL, U5_HDR_WIDTH)
    Dim colNo As Long
    colNo = modUISheet.ColOf(hdr, U5_FIRST_COL, colName)
    If colNo <= 0 Then Exit Function

    Dim lastRow As Long
    lastRow = modUISheet.BlockLastRow(ws, headerRow, U5_FIRST_COL, 2, RoomOf(anchorName))

    Dim seqCol As Long
    seqCol = modUISheet.ColOf(hdr, U5_FIRST_COL, "seq")

    Dim items() As String
    Dim keys() As Long
    Dim n As Long
    ReDim items(0 To lastRow - headerRow)
    ReDim keys(0 To lastRow - headerRow)

    Dim r As Long
    Dim v As String
    For r = headerRow + 1 To lastRow
        v = modUISheet.CellText(ws, r, colNo)
        If LenB(Trim$(v)) > 0 Then
            items(n) = modUICaseFmt.StrJson(v)
            keys(n) = n + 1
            If seqCol > 0 Then keys(n) = CLng(Val(modUISheet.CellText(ws, r, seqCol)))
            n = n + 1
        End If
    Next r
    If n = 0 Then Exit Function

    modUICaseFmt.SortItems items, keys, n

    Dim acc As String
    Dim i As Long
    For i = 0 To n - 1
        If LenB(acc) > 0 Then acc = acc & ","
        acc = acc & items(i)
    Next i

    ScalarArrJson = "[" & acc & "]"
End Function

' 列定義に対応する絶対列番号の配列。1つでも見つからなければ Empty。
Public Function ColIndexes(ByVal ws As Object, ByVal headerRow As Long, _
                            ByVal colSpec As String) As Variant
    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, U5_FIRST_COL, U5_HDR_WIDTH)
    If IsEmpty(hdr) Then Exit Function

    Dim specs() As String
    specs = Split(colSpec, ";")

    Dim cols() As Long
    ReDim cols(0 To UBound(specs) - LBound(specs))

    Dim i As Long
    Dim physName As String
    Dim kindText As String
    Dim extraText As String
    Dim pathText As String
    For i = LBound(specs) To UBound(specs)
        modUICaseFmt.SplitCol specs(i), physName, kindText, extraText, pathText
        cols(i - LBound(specs)) = modUISheet.ColOf(hdr, U5_FIRST_COL, physName)
        If cols(i - LBound(specs)) <= 0 Then Exit Function
    Next i

    ColIndexes = cols
End Function

' 1行の値を列定義順に読む。
Private Function RowValues(ByVal ws As Object, ByVal rowNo As Long, _
                           ByVal cols As Variant) As Variant
    Dim vals() As String
    ReDim vals(LBound(cols) To UBound(cols))

    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        vals(i) = modUISheet.CellText(ws, rowNo, CLng(cols(i)))
    Next i
    RowValues = vals
End Function

' 列定義の物理名で値を引く(位置ではなく名前で引く=列順の変更に強い)。
Private Function ValueOf(ByVal colSpec As String, ByVal vals As Variant, _
                         ByVal physWanted As String) As String
    Dim specs() As String
    specs = Split(colSpec, ";")

    Dim i As Long
    Dim physName As String
    Dim kindText As String
    Dim extraText As String
    Dim pathText As String
    For i = LBound(specs) To UBound(specs)
        modUICaseFmt.SplitCol specs(i), physName, kindText, extraText, pathText
        If physName = physWanted Then
            ValueOf = CStr(vals(i - LBound(specs) + LBound(vals)))
            Exit Function
        End If
    Next i
End Function

Public Function ColCount(ByVal colSpec As String) As Long
    Dim specs() As String
    specs = Split(colSpec, ";")
    ColCount = UBound(specs) - LBound(specs) + 1
End Function

' ブロックの「部屋」(見出しの下に使ってよい行数)。次のブロックの見出し行の
'   1行手前(空行=表の終端)までを上限にする。最後のブロックは U5_MAX_ROOM。
Public Function RoomOf(ByVal anchorName As String) As Long
    RoomOf = U5_MAX_ROOM

    Dim stepNo As Long
    Dim names() As String
    Dim i As Long
    Dim here As Long
    Dim nxt As Long

    For stepNo = 1 To 4
        names = Split(modUICaseFmt.AnchorsOf(stepNo), ";")
        For i = LBound(names) To UBound(names)
            If names(i) = anchorName Then
                If i < UBound(names) Then
                    here = modUISheet.BlockRow(anchorName)
                    nxt = modUISheet.BlockRow(names(i + 1))
                    If here > 0 And nxt > here Then
                        RoomOf = nxt - here - 2      ' 間の空行を1行残す
                        If RoomOf < 1 Then RoomOf = 1
                    End If
                End If
                Exit Function
            End If
        Next i
    Next stepNo
End Function
