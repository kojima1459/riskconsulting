Attribute VB_Name = "modCaseStore2"
Option Explicit

' ============================================================================
' modCaseStore2 - 案件一覧 / case_data の下位シートI/O(app層・T-20。裁定書8 A-2)
' ----------------------------------------------------------------------------
' なぜ2本に分かれているのか:
'   modCaseStore が 30,000字契約(12章§2)の警告帯(28,000字)へ達し、16章E-06の
'   書込口 SetStepOutcome を足す余地が無くなった。そこで modCompanyFile /
'   modCompanyFile2、modValidate / modValidate2 と【同じ切り口】で割った。
'     modCaseStore  = 「何を書くか」(採番・参照優先・状態遷移・分割保存の規約)
'     modCaseStore2 = 「どこへどう書くか」(シートを取る・最終行・矩形読み・
'                      行を消す・列名で1セル書く・セル値をLongへ)
'   本モジュールは 13章の列定義も data_key の enum も知らない。呼んでよいのは
'   modCaseStore だけ(依存は modCaseStore -> modCaseStore2 の一方向)であり、
'   14章§6の公開契約面には載せない(modCompanyFile2 と同じ扱い。vba_lint の
'   CONTRACT は required=[] で登録する)。
'
' R4(12章§4): 案件一覧 と case_data のシートI/Oが責務そのもの。許可の幅は
'   modCaseStore と同じ「本体ブックの案件2枚」で広がっていない。
'   セル書込は全て modUtilText.SetCellSafe(16章 NFR-S7①)。数値列だけは内部
'   生成のLongなので ' SAFE:const を明示する。
' ============================================================================

' xlUp の数値(組込定数名を書かず LO の構文チェックで未定義名にしない)。
Private Const CS2_DIR_UP As Long = -4162

' 見出し行を読む幅。案件一覧23列に余裕を見た「探索範囲」で列番号ではない。
Private Const CS2_SCAN_COLS As Long = 32

' 名前でシートを取る。無ければ Nothing(実行時に生やすと列定義の欠けた表になる)。
Public Function SheetOf(ByVal sheetName As String) As Object
    On Error GoTo NoSheet
    Set SheetOf = ThisWorkbook.Worksheets(sheetName)
    Exit Function
NoSheet:
    Set SheetOf = Nothing
End Function

' Application.UserName(13章§2.1 owner)。取得できない環境では ""。
Public Function OwnerName() As String
    On Error GoTo NoName
    OwnerName = CStr(Application.UserName)
    Exit Function
NoName:
    OwnerName = vbNullString
End Function

' A列基準の最終行。データが無ければ1(見出し行)。
Public Function LastRowOf(ByVal ws As Object) As Long
    Dim r As Long
    On Error GoTo One1
    r = ws.Cells(ws.Rows.count, 1).End(CS2_DIR_UP).row
    If r < 1 Then r = 1
    LastRowOf = r
    Exit Function
One1:
    LastRowOf = 1
End Function

' 見出し行を含む矩形を一度だけ読む。1セルだけの Range は2次元配列にならないので
' 必ず2行以上を読む。
Public Function ReadBlock(ByVal ws As Object, ByVal lastRow As Long) As Variant
    On Error GoTo Empty0
    Dim hi As Long
    hi = lastRow
    If hi < 2 Then hi = 2
    ReadBlock = ws.Range(ws.Cells(1, 1), ws.Cells(hi, CS2_SCAN_COLS)).Value
    Exit Function
Empty0:
    ReadBlock = Empty
End Function

' 読み込み済みブロックから case_id 一致行を探す。見つからなければ0。
Public Function RowOfCase(ByVal blk As Variant, ByVal lastRow As Long, _
                          ByVal caseCol As Long, ByVal caseId As String) As Long
    On Error GoTo NotFound
    If caseCol <= 0 Then Exit Function
    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(blk(r, caseCol))) = caseId Then
            RowOfCase = r
            Exit Function
        End If
    Next r
    Exit Function
NotFound:
    RowOfCase = 0
End Function

' LocateRow - 「案件一覧の1行を書く」ための場所決めを1本にまとめたもの。
'   シートを取る -> 最終行 -> 矩形読み -> case_id 列 -> 該当行、までを行い
'   ws / blk / rowNo を帯域外で返す。どこかで欠けたら False(呼び出し側は
'   これまでどおり黙って中止する。ここでログは書かない)。
'   SetStatus / SetStepOutcome / ApplyRepairedState が同じ12行を3度書いて
'   いたのを畳んだもので、判断は1行も持たない。
Public Function LocateRow(ByVal sheetName As String, ByVal caseId As String, _
                          ByRef ws As Object, ByRef blk As Variant, _
                          ByRef rowNo As Long) As Boolean
    rowNo = 0
    Set ws = SheetOf(sheetName)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    blk = ReadBlock(ws, lastRow)

    Dim cId As Long
    cId = modUtil.FindHeaderCol(blk, "case_id")
    If cId <= 0 Then Exit Function

    rowNo = RowOfCase(blk, lastRow, cId, Trim$(caseId))
    LocateRow = (rowNo > 0)
End Function

' case_data の1行が (case_id, data_key) に一致するか。
Public Function MatchesRow(ByVal blk As Variant, ByVal r As Long, _
                           ByVal caseCol As Long, ByVal keyCol As Long, _
                           ByVal caseId As String, ByVal dataKey As String) As Boolean
    On Error GoTo NoMatch
    If Trim$(CStr(blk(r, caseCol))) <> caseId Then Exit Function
    If Trim$(CStr(blk(r, keyCol))) <> dataKey Then Exit Function
    MatchesRow = True
    Exit Function
NoMatch:
    MatchesRow = False
End Function

' 同じ (case_id, data_key) の行を全削除し件数を返す(下から回すのは行ずれ回避)。
Public Function DropRowsOf(ByVal ws As Object, ByVal caseId As String, _
                           ByVal dataKey As String) As Long
    On Error GoTo Done0
    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = ReadBlock(ws, lastRow)
    Dim cCase As Long, cKey As Long
    cCase = modUtil.FindHeaderCol(blk, "case_id")
    cKey = modUtil.FindHeaderCol(blk, "data_key")
    If cCase <= 0 Or cKey <= 0 Then Exit Function

    Dim n As Long
    Dim r As Long
    For r = lastRow To 2 Step -1
        If MatchesRow(blk, r, cCase, cKey, caseId, Trim$(dataKey)) Then
            ws.Rows(r).Delete
            n = n + 1
        End If
    Next r
    DropRowsOf = n
    Exit Function
Done0:
    DropRowsOf = 0
End Function

' 列名で位置を引いて1セル書く。列が無ければ何もしない。書込は必ず SetCellSafe。
Public Sub PutText(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                   ByVal colName As String, ByVal textValue As String)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(r, c), textValue, ws.Name & "/" & colName
End Sub

' 数値列(seq / round_no / last_ok_step)。内部生成のLongで数値型を保つ必要があり
' SetCellSafe(文字列を返す)は通さない(modLog.PutNum と同じ理由)。
Public Sub PutNum(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                  ByVal colName As String, ByVal numValue As Long)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    ws.Cells(r, c).Value = numValue   ' SAFE:const 内部生成のLong(外部由来テキストではない)
End Sub

' セル値を Long へ。空・非数値・取得失敗は0。
Public Function ToLongSafe(ByVal v As Variant) As Long
    On Error GoTo Zero0
    Dim t As String
    t = Trim$(CStr(v))
    If LenB(t) = 0 Then Exit Function
    If Not IsNumeric(t) Then Exit Function
    ToLongSafe = CLng(t)
    Exit Function
Zero0:
    ToLongSafe = 0
End Function
