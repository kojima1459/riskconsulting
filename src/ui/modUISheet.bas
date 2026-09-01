Attribute VB_Name = "modUISheet"
Option Explicit

' ============================================================================
' modUISheet - ui層のシート操作プリミティブ(T-30/31/32/34の共有部品)
' ----------------------------------------------------------------------------
' なぜ独立したモジュールなのか:
'   modUIHome / modUICase / modUICase2 / modUICase3 / modUIInbox /
'   modUISparring は「名前付きレンジを1点読み書きする」「ブロックアンカーから
'   見出し行を解決する」「図形ボタンを置いてOnActionを配線する」の3つを必ず
'   行う。各モジュールに同じ12行を6度書くと、13章§2.9(行番号を仮定しない)と
'   11章§5(図形ボタン+OnAction)の規約が6箇所に散る。**規約の実装を1箇所に
'   まとめる**ために切り出した(modCaseStore2 が案件2枚の下位I/Oを引き受けて
'   いるのと同じ切り口。12章§2に登録済み)。
'
'   本モジュールは 13章の列定義も 19章のenumも知らない。**何を書くかは
'   呼び出し側が決め、本モジュールはどこへどう書くかだけを持つ**。
'
' R4(12章§4): ui層なのでExcelトークンを使ってよい。
' NFR-S7(1)(16章): セル書込は modUtilText.SetCellSafe 経由のみ。見出し・
'   採番など内部生成の定数書込には行末へ ' SAFE:const を明示する。
' 11章§5: シートProtectは使わない(Locked + EnableSelection + ParkFocus)。
'   本モジュールは Protect / Unprotect を1度も呼ばない。
' ============================================================================

' Excel組み込み定数の数値(名前を書かず、LibreOffice側の構文チェックで未定義名に
' ならないようにする。modBoot / modCaseStore2 と同じ流儀)。
Private Const US_DIR_UP As Long = -4162             ' xlUp
Private Const US_DV_TYPE_LIST As Long = 3           ' xlValidateList
Private Const US_DV_ALERT_STOP As Long = 1          ' xlValidAlertStop
Private Const US_SHAPE_ROUNDED As Long = 5          ' msoShapeRoundedRectangle
Private Const US_SHAPE_TEXTBOX As Long = 1          ' msoTextOrientationHorizontal
Private Const US_SHEET_VERY_HIDDEN As Long = 2      ' xlSheetVeryHidden
Private Const US_SHEET_VISIBLE As Long = -1         ' xlSheetVisible
Private Const US_ALIGN_CENTER As Long = 2           ' xlCenter / msoAlignCenter相当
Private Const US_COLOR_BTN As Long = 16777215&      ' ボタン地=白(RGB 255,255,255)
Private Const US_COLOR_WARN As Long = 13551615&     ' 淡い赤(超過表示)
Private Const US_COLOR_NONE As Long = 16777215&     ' 白

' 裁定書14 裁定7(ボタンの刷新)。値は「RGB(r,g,b) = r + g*256 + b*65536」。
Private Const US_COLOR_BTN_LINE As Long = 13616055& ' 枠線(RGB 183,195,207)
Private Const US_COLOR_BTN_TEXT As Long = 2562065&  ' 文字(RGB 17,24,39)
Private Const US_COLOR_BTN_PRIMARY As Long = 5990145&  ' 主要動線の地(RGB 1,103,91)
Private Const US_COLOR_BTN_DANGER As Long = 2500316&   ' 取り消し系の地(RGB 220,38,38)
Private Const US_BTN_FONT As String = "Yu Gothic UI"
Private Const US_BTN_FONT_SIZE As Double = 10#
Private Const US_PLACEMENT_FREE As Long = 3         ' xlFreeFloating
Private Const US_BTN_ROUND As Double = 0.35         ' 角丸の深さ(Adjustments(1))
Private Const US_BTN_ROW_PAD As Double = 4#         ' アンカー行に足す余白

Private Const US_BTN_HEIGHT As Double = 26#
Private Const US_LABEL_HEIGHT As Double = 16#
Private Const US_SRC As String = "modUISheet"

' 見出し行を読む既定の探索幅(列番号ではなく「この幅までを見出しとして読む」)。
Private Const US_SCAN_COLS As Long = 24

' ============================================================================
' シート・名前付きレンジ
' ============================================================================

' 名前でシートを取る。無ければ Nothing(実行時に生やすと列定義の欠けた表になる)。
Public Function SheetOf(ByVal sheetName As String) As Object
    On Error GoTo NoSheet
    Set SheetOf = ThisWorkbook.Worksheets(sheetName)
    Exit Function
NoSheet:
    Set SheetOf = Nothing
End Function

' 名前付きレンジの左上1セル。不在は Nothing。
Public Function NamedCell(ByVal rangeName As String) As Object
    On Error GoTo NoName
    Set NamedCell = ThisWorkbook.Names(rangeName).RefersToRange.Cells(1, 1)
    Exit Function
NoName:
    Set NamedCell = Nothing
End Function

' 名前付きレンジ1点の値(文字列)。不在・読めないときは ""。
Public Function ReadNamed(ByVal rangeName As String) As String
    On Error GoTo NoValue
    Dim cell As Object
    Set cell = NamedCell(rangeName)
    If cell Is Nothing Then Exit Function
    ReadNamed = Trim$(CStr(cell.Value))
    Exit Function
NoValue:
    ReadNamed = vbNullString
End Function

' 名前付きレンジ1点への書込。外部由来かどうかに関わらず SetCellSafe を通す
' (16章NFR-S7(1): 書込口を分岐させない)。不在は無音で何もしない。
Public Sub WriteNamed(ByVal rangeName As String, ByVal valueText As String)
    Dim cell As Object
    Set cell = NamedCell(rangeName)
    If cell Is Nothing Then Exit Sub
    modUtilText.SetCellSafe cell, valueText, US_SRC & "/" & rangeName
End Sub

' 名前付きレンジ1点の背景色。超過表示(13章§2.11 ci_count_total)に使う。
'   over=True で淡い赤、False で白へ戻す(色を塗りっぱなしにしない)。
Public Sub MarkNamed(ByVal rangeName As String, ByVal over As Boolean)
    On Error Resume Next
    Dim cell As Object
    Set cell = NamedCell(rangeName)
    If cell Is Nothing Then Exit Sub
    If over Then
        cell.Interior.Color = US_COLOR_WARN
    Else
        cell.Interior.Color = US_COLOR_NONE
    End If
End Sub

' ============================================================================
' ブロックアンカー(13章§2.9: 複数ブロックのシートは行番号を仮定しない)
' ============================================================================

' ブロックアンカーが指す見出し行の行番号。不在は0。
Public Function BlockRow(ByVal anchorName As String) As Long
    Dim cell As Object
    Set cell = NamedCell(anchorName)
    If cell Is Nothing Then Exit Function
    On Error GoTo NoRow
    BlockRow = cell.row
    Exit Function
NoRow:
    BlockRow = 0
End Function

' ブロックアンカーが指す見出し行の左端列。不在は0。
Public Function BlockCol(ByVal anchorName As String) As Long
    Dim cell As Object
    Set cell = NamedCell(anchorName)
    If cell Is Nothing Then Exit Function
    On Error GoTo NoCol
    BlockCol = cell.Column
    Exit Function
NoCol:
    BlockCol = 0
End Function

' ブロックアンカーが属するシート。不在は Nothing。
Public Function BlockSheet(ByVal anchorName As String) As Object
    Dim cell As Object
    Set cell = NamedCell(anchorName)
    If cell Is Nothing Then Exit Function
    On Error GoTo NoSheet
    Set BlockSheet = cell.Worksheet
    Exit Function
NoSheet:
    Set BlockSheet = Nothing
End Function

' 見出し行を1回で読む(FindHeaderCol へ渡す2次元配列)。読めなければ Empty。
Public Function HeaderOf(ByVal ws As Object, ByVal headerRow As Long, _
                         ByVal firstCol As Long, ByVal widthCols As Long) As Variant
    On Error GoTo Empty0
    Dim w As Long
    w = widthCols
    If w <= 0 Then w = US_SCAN_COLS
    HeaderOf = ws.Range(ws.Cells(headerRow, firstCol), _
                        ws.Cells(headerRow, firstCol + w - 1)).Value
    Exit Function
Empty0:
    HeaderOf = Empty
End Function

' 見出し配列から列名の**絶対列番号**を返す(firstCol基準の相対位置を足す)。0=不在。
Public Function ColOf(ByVal hdr As Variant, ByVal firstCol As Long, _
                      ByVal colName As String) As Long
    Dim rel As Long
    rel = modUtil.FindHeaderCol(hdr, colName)
    If rel <= 0 Then Exit Function
    ColOf = firstCol + rel - 1
End Function

' ブロックの最終データ行(見出しの直下から最初の空行の手前まで。13章§2.2の
' 逆シリアライズ規約5「空行が表の終端」)。データが無ければ headerRow を返す。
Public Function BlockLastRow(ByVal ws As Object, ByVal headerRow As Long, _
                             ByVal firstCol As Long, ByVal widthCols As Long, _
                             ByVal maxRows As Long) As Long
    BlockLastRow = headerRow
    On Error GoTo Done

    Dim w As Long
    w = widthCols
    If w <= 0 Then w = 1

    Dim limitRows As Long
    limitRows = maxRows
    If limitRows <= 0 Then limitRows = 200

    Dim r As Long
    For r = headerRow + 1 To headerRow + limitRows
        If Not RowHasValue(ws, r, firstCol, w) Then Exit For
        BlockLastRow = r
    Next r
Done:
End Function

' 1行が「全列空」でないか(空行=表の終端)。
Public Function RowHasValue(ByVal ws As Object, ByVal rowNo As Long, _
                            ByVal firstCol As Long, ByVal widthCols As Long) As Boolean
    On Error GoTo NoValue
    Dim c As Long
    For c = firstCol To firstCol + widthCols - 1
        If LenB(Trim$(CStr(ws.Cells(rowNo, c).Value))) > 0 Then
            RowHasValue = True
            Exit Function
        End If
    Next c
    Exit Function
NoValue:
    RowHasValue = False
End Function

' ブロックの本文行を空へ戻す(再描画で古い行が残って混ざらないようにする)。
Public Sub ClearBlock(ByVal ws As Object, ByVal headerRow As Long, _
                      ByVal firstCol As Long, ByVal widthCols As Long, _
                      ByVal rowCount As Long)
    On Error Resume Next
    Dim r As Long
    For r = headerRow + 1 To headerRow + rowCount
        ws.Range(ws.Cells(r, firstCol), _
                 ws.Cells(r, firstCol + widthCols - 1)).ClearContents
    Next r
End Sub

' 1セルの読取(表示用に切り詰めない生の値)。範囲外・読めないときは ""。
Public Function CellText(ByVal ws As Object, ByVal rowNo As Long, ByVal colNo As Long) As String
    On Error GoTo NoValue
    If colNo <= 0 Or rowNo <= 0 Then Exit Function
    CellText = CStr(ws.Cells(rowNo, colNo).Value)
    Exit Function
NoValue:
    CellText = vbNullString
End Function

' 1セルの書込(外部由来テキストの唯一の書込口 SetCellSafe 経由)。
Public Sub PutText(ByVal ws As Object, ByVal rowNo As Long, ByVal colNo As Long, _
                   ByVal valueText As String, ByVal whereNote As String)
    If colNo <= 0 Or rowNo <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(rowNo, colNo), valueText, whereNote
End Sub

' ============================================================================
' 単票シート(1行目=物理名ヘッダ)の行探索。受信箱・判断台帳・フィードバックが使う。
' ============================================================================

' A列基準の最終行。データが無ければ1(見出し行)。
Public Function LastRowOf(ByVal ws As Object) As Long
    On Error GoTo One1
    Dim r As Long
    r = ws.Cells(ws.Rows.count, 1).End(US_DIR_UP).row
    If r < 1 Then r = 1
    LastRowOf = r
    Exit Function
One1:
    LastRowOf = 1
End Function

' いま選択されているセルの行番号(利用者が選んだ行に対して操作する画面のため)。
'   別シートが選択されている・取得できない場合は0。
Public Function SelectedRow(ByVal ws As Object) As Long
    On Error GoTo NoRow
    Dim sel As Object
    Set sel = Application.Selection
    If sel Is Nothing Then Exit Function
    If sel.Worksheet.Name <> ws.Name Then Exit Function
    SelectedRow = sel.Cells(1, 1).row
    Exit Function
NoRow:
    SelectedRow = 0
End Function

' 画面遷移(11章§3)。ui層の責務であり、失敗しても黙って固まらないよう
' 戻り値で成否を返す(呼び出し側が案内を出せる)。
Public Function ShowSheet(ByVal sheetName As String) As Boolean
    On Error GoTo Failed
    Dim ws As Object
    Set ws = SheetOf(sheetName)
    If ws Is Nothing Then Exit Function
    ws.Visible = US_SHEET_VISIBLE
    ws.Activate
    ShowSheet = True
    Exit Function
Failed:
    ShowSheet = False
End Function

' ============================================================================
' 入力規則(データの入力規則リスト)の隠しレンジ方式(11章§5)
' ----------------------------------------------------------------------------
' インラインの値並びは255字上限があるため、値はveryHiddenシートへ複製して
' 名前付きレンジで参照する(セル参照は255字制限の対象外)。
' ============================================================================

' veryHidden の作業シートを用意する(無ければ作る)。13章§2.9より照合対象外。
Public Function EnsureHiddenSheet(ByVal sheetName As String) As Object
    On Error GoTo Failed
    Dim ws As Object
    Set ws = SheetOf(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        If ws Is Nothing Then Exit Function
        ws.Name = sheetName
    End If
    ws.Visible = US_SHEET_VERY_HIDDEN
    Set EnsureHiddenSheet = ws
    Exit Function
Failed:
    Set EnsureHiddenSheet = Nothing
End Function

' 値の並びを隠しシートの1列へ書き、その範囲を指す名前付きレンジを張り直す。
'   戻り値=張れた名前(失敗は "")。値は内部定数(19章§3のレジストリ)だが、
'   書込口は分岐させず SetCellSafe を通す。
Public Function PutEnumRange(ByVal ws As Object, ByVal colNo As Long, _
                             ByVal rangeName As String, ByVal valuesText As String, _
                             ByVal sepText As String) As String
    On Error GoTo Failed
    If ws Is Nothing Then Exit Function
    If colNo <= 0 Then Exit Function

    Dim vals() As String
    vals = Split(valuesText, sepText)

    Dim n As Long
    n = UBound(vals) - LBound(vals) + 1
    If n <= 0 Then Exit Function

    ' 前回より短くなったときに古い値が残らないよう、余白を先に消す。
    ws.Range(ws.Cells(1, colNo), ws.Cells(n + 40, colNo)).ClearContents

    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        modUtilText.SetCellSafe ws.Cells(i - LBound(vals) + 1, colNo), vals(i), _
                                US_SRC & "/" & rangeName
    Next i

    Dim target As Object
    Set target = ws.Range(ws.Cells(1, colNo), ws.Cells(n, colNo))
    ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=target
    PutEnumRange = rangeName
    Exit Function
Failed:
    PutEnumRange = vbNullString
End Function

' 対象レンジへリスト入力規則(名前付きレンジ参照)を張り直す。
Public Sub BindListValidation(ByVal target As Object, ByVal rangeName As String, _
                              ByVal errTitle As String, ByVal errMessage As String)
    On Error Resume Next
    If target Is Nothing Then Exit Sub
    If LenB(rangeName) = 0 Then Exit Sub
    target.Validation.Delete
    target.Validation.Add Type:=US_DV_TYPE_LIST, AlertStyle:=US_DV_ALERT_STOP, _
        Formula1:="=" & rangeName
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
    target.Validation.ErrorTitle = errTitle
    target.Validation.ErrorMessage = errMessage
End Sub

' 名前付きレンジ1点へ入力規則を張る(帳票型の属性欄。13章§2.11)。
Public Sub BindNamedValidation(ByVal cellName As String, ByVal rangeName As String, _
                               ByVal errTitle As String, ByVal errMessage As String)
    Dim cell As Object
    Set cell = NamedCell(cellName)
    If cell Is Nothing Then Exit Sub
    BindListValidation cell, rangeName, errTitle, errMessage
End Sub

' ブロックの1列(見出しの下 rowCount 行)へ入力規則を張る(テーブル型。13章§2.12-)。
Public Sub BindBlockValidation(ByVal ws As Object, ByVal headerRow As Long, _
                               ByVal colNo As Long, ByVal rowCount As Long, _
                               ByVal rangeName As String, ByVal errTitle As String, _
                               ByVal errMessage As String)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If colNo <= 0 Or rowCount <= 0 Then Exit Sub
    Dim target As Object
    Set target = ws.Range(ws.Cells(headerRow + 1, colNo), _
                          ws.Cells(headerRow + rowCount, colNo))
    BindListValidation target, rangeName, errTitle, errMessage
End Sub

' ============================================================================
' 図形ボタン(11章§5: UserFormは使わない。図形＋OnAction)
' ----------------------------------------------------------------------------
' 同名の図形を消してから置き直す(再実行で二重に生えないこと・キャプションと
' 配線の変更が必ず反映されることの両方を、同じ1本で担保する)。
' ============================================================================

' 図形ボタンを1つ用意する。戻り値 True=置けた。
'   shapeKey は Shape.Name(ブック内で安定・一意にする)。
'   onActionName は "modUIHome.HomeRunAll" 形式の**Publicプロシージャ名**。
'   見た目(書体・角丸・配色)は EnsureButtonEx の "plain" と同一であり、
'   呼び出し側のシグネチャは裁定書14 裁定7の前後で変えていない。
Public Function EnsureButton(ByVal ws As Object, ByVal shapeKey As String, _
                             ByVal caption As String, ByVal anchorRow As Long, _
                             ByVal anchorCol As Long, ByVal widthPt As Double, _
                             ByVal onActionName As String) As Boolean
    EnsureButton = EnsureButtonEx(ws, shapeKey, caption, anchorRow, anchorCol, _
                                  widthPt, onActionName, "plain")
End Function

' 種別つきの図形ボタン(裁定書14 裁定7)。kind:
'   "primary" = 主要動線(濃緑地＋白の太字) / "plain" = 既定(白地＋枠) /
'   "danger"  = 取り消し系(赤地＋白の太字)。未知の値は "plain" として扱う。
' 高さは US_BTN_HEIGHT(26pt)で固定し、**アンカー行の行高をボタンが収まる高さまで
'   広げてから**置く。行高がボタンより低いと隣接行のボタンと縦に重なる(実機で
'   起きた不具合。裁定書14 追補1)ため、重なりの芽をここで摘む。
Public Function EnsureButtonEx(ByVal ws As Object, ByVal shapeKey As String, _
                               ByVal caption As String, ByVal anchorRow As Long, _
                               ByVal anchorCol As Long, ByVal widthPt As Double, _
                               ByVal onActionName As String, _
                               ByVal kind As String) As Boolean
    On Error GoTo Failed
    If ws Is Nothing Then Exit Function

    DropShape ws, shapeKey
    GrowRowForButton ws, anchorRow

    Dim anchor As Object
    Set anchor = ws.Cells(anchorRow, anchorCol)

    Dim shp As Object
    Set shp = ws.Shapes.AddShape(US_SHAPE_ROUNDED, anchor.Left, anchor.Top, _
                                 widthPt, US_BTN_HEIGHT)
    If shp Is Nothing Then Exit Function

    shp.Name = shapeKey
    shp.Placement = US_PLACEMENT_FREE
    On Error Resume Next
    shp.Adjustments(1) = US_BTN_ROUND
    On Error GoTo Failed

    Dim fillColor As Long
    Dim textColor As Long
    Dim boldText As Boolean
    fillColor = US_COLOR_BTN
    textColor = US_COLOR_BTN_TEXT
    boldText = False
    If StrComp(kind, "primary", vbBinaryCompare) = 0 Then
        fillColor = US_COLOR_BTN_PRIMARY
        textColor = US_COLOR_NONE
        boldText = True
    ElseIf StrComp(kind, "danger", vbBinaryCompare) = 0 Then
        fillColor = US_COLOR_BTN_DANGER
        textColor = US_COLOR_NONE
        boldText = True
    End If

    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = True
    shp.Line.ForeColor.RGB = US_COLOR_BTN_LINE
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.HorizontalAlignment = US_ALIGN_CENTER
    shp.TextFrame.Characters.Font.Name = US_BTN_FONT
    shp.TextFrame.Characters.Font.Size = US_BTN_FONT_SIZE
    shp.TextFrame.Characters.Font.Color = textColor
    shp.TextFrame.Characters.Font.Bold = boldText
    shp.OnAction = onActionName
    EnsureButtonEx = True
    Exit Function
Failed:
    EnsureButtonEx = False
End Function

' アンカー行をボタンが収まる高さまで広げる(縮めはしない)。
Private Sub GrowRowForButton(ByVal ws As Object, ByVal anchorRow As Long)
    On Error Resume Next
    If anchorRow <= 0 Then Exit Sub
    Dim needed As Double
    needed = US_BTN_HEIGHT + US_BTN_ROW_PAD
    If ws.Rows(anchorRow).RowHeight < needed Then ws.Rows(anchorRow).RowHeight = needed
End Sub

' ボタンの左端(pt)。幾何計算を呼び出し側で組み立てるための読み取り口。
'   取れないときは -1(呼び出し側は配置をあきらめて既定の列アンカーへ落とす)。
Public Function CellLeft(ByVal ws As Object, ByVal rowNo As Long, ByVal colNo As Long) As Double
    On Error GoTo Failed
    CellLeft = ws.Cells(rowNo, colNo).Left
    Exit Function
Failed:
    CellLeft = -1#
End Function

' 表示専用のラベル図形(名前付きレンジを持たない表示欄。11章のワイヤーで
'   「関心度」「診断全文」等、13章が列も名前付きレンジも定義していない
'   読み取り専用の表示に使う。**シートのデータ面を汚さない**ための選択)。
Public Function EnsureLabel(ByVal ws As Object, ByVal shapeKey As String, _
                            ByVal caption As String, ByVal anchorRow As Long, _
                            ByVal anchorCol As Long, ByVal widthPt As Double, _
                            ByVal heightPt As Double) As Boolean
    On Error GoTo Failed
    If ws Is Nothing Then Exit Function

    DropShape ws, shapeKey

    Dim anchor As Object
    Set anchor = ws.Cells(anchorRow, anchorCol)

    Dim h As Double
    h = heightPt
    If h <= 0 Then h = US_LABEL_HEIGHT

    Dim shp As Object
    Set shp = ws.Shapes.AddTextbox(US_SHAPE_TEXTBOX, anchor.Left, anchor.Top, widthPt, h)
    If shp Is Nothing Then Exit Function

    shp.Name = shapeKey
    shp.Line.Visible = False
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Size = 9
    EnsureLabel = True
    Exit Function
Failed:
    EnsureLabel = False
End Function

' ラベル図形の文言を差し替える(無ければ何もしない=描画前でも落ちない)。
Public Sub SetLabelText(ByVal ws As Object, ByVal shapeKey As String, ByVal caption As String)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    ws.Shapes(shapeKey).TextFrame.Characters.Text = caption
End Sub

' 同名図形の削除(存在しなくても落ちない)。
Public Sub DropShape(ByVal ws As Object, ByVal shapeKey As String)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    ws.Shapes(shapeKey).Delete
End Sub

' 接頭辞が一致する図形をまとめて消す(行数可変のブロックに置いた行ボタンの
'   後始末。逆順に回すのはコレクションの添字がずれるため)。
Public Sub DropShapesByPrefix(ByVal ws As Object, ByVal prefixText As String)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        If InStr(1, ws.Shapes(i).Name, prefixText, vbBinaryCompare) = 1 Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

' ============================================================================
' クリップボード(11章 追加収集ブロックの[コピー]。17章 T-31 DoD)
' ----------------------------------------------------------------------------
' UserFormは使わないため参照設定を足さず、MSForms.DataObject を遅延生成して
' 使う(GUIDによる New: 生成。参照設定なしで動く定番手)。失敗したら戻り値
' False を返し、呼び出し側が「セルを選ぶのでCtrl+Cしてください」へ落とす。
' ============================================================================
Public Function CopyToClipboard(ByVal payloadText As String) As Boolean
    On Error GoTo Failed
    If LenB(payloadText) = 0 Then Exit Function
    Dim dobj As Object
    Set dobj = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If dobj Is Nothing Then Exit Function
    dobj.SetText payloadText
    dobj.PutInClipboard
    CopyToClipboard = True
    Exit Function
Failed:
    CopyToClipboard = False
End Function
