Attribute VB_Name = "modUINavDraw"
Option Explicit

' ============================================================================
' modUINavDraw - ナビの描画(ui層・T-49)
' ----------------------------------------------------------------------------
' 11章§3.1 のワイヤーと §8.5 の再現可否表・§8.6 の流用表が正。
' 状態の決定(いまどのSTEPか)は modUINav が持ち、本モジュールは**描くだけ**。
'
' 流用元(11章§8.6。ファイル:行を明記する):
'   帯の地   notebook/src/ui/modUINexusDraw.bas:354 HeaderButton ＋
'            notebook/src/ui/modSkin.bas:89 ApplyHeaderDepth(2色グラデーション)
'            -> DrawCoachBar。色は本製品の主色へ差し替えた(濃緑を持ち込まない)
'   丸ピル   notebook/src/ui/modUINexusDraw.bas:382 HeaderTheme(楕円)
'            -> 角丸四角へ改変(モックの .coach-no は角丸8px)
'   パネル   notebook/src/ui/modHub.bas:448 DrawNavButtons(角丸・Adjustments 0.08)
'            notebook/src/ui/modHelp.bas:145 ShowHelpCard(AutoSizeで高さ実測)
'            -> DrawSections。**AutoSize は使わず行から実測**する
'   影       notebook/src/ui/modSkin.bas:123 ApplyLightShadow / :161 ApplySoftShadow
'            -> modUISheet.ApplyLightShadow / ApplySoftShadow へ移植済み
'   一括削除 notebook/src/ui/modDash.bas:621 RemoveShapesByPrefix
'            -> 既存の modUISheet.DropShapesByPrefix が同等(名前を集めてから消す)
'
' 移植時の禁忌(11章§8.6。1つでも破ったらビルドを止める):
'   Hyperlinks.Add を使わない / msoTrue 等の定数名を書かず数値リテラル＋コメント /
'   Shapes を列挙しながら削除しない / 絵文字リテラルを書かない /
'   1物理行のCP932バイト長を1,000以下に保つ / 影・グラデーションの失敗が
'   業務を止めない。
' ============================================================================

Private Const UD_SRC As String = "modUINavDraw"
Private Const UD_SHEET As String = "ナビ"

' 図形の接頭辞。描き直す前にまとめて落とす(孤児防止)。
Private Const UD_PANEL_PREFIX As String = "nv_panel"
Private Const UD_BAND_PREFIX As String = "nv_band"
Private Const UD_FOCUS As String = "nv_focus"
Private Const UD_WAIT_PREFIX As String = "nv_wait"
Private Const UD_BTN_PREFIX As String = "btn_nv_"

' Excel/Office組み込み定数の数値(名前を書かない。11章§8.6 禁忌2)。
Private Const UD_SHAPE_ROUNDED As Long = 5      ' msoShapeRoundedRectangle
Private Const UD_PLACEMENT_FREE As Long = 3     ' xlFreeFloating
Private Const UD_ALIGN_LEFT As Long = 1         ' msoAlignLeft
Private Const UD_ALIGN_CENTER As Long = 2       ' msoAlignCenter
Private Const UD_ANCHOR_MIDDLE As Long = 3      ' msoAnchorMiddle
Private Const UD_TRUE As Long = -1              ' msoTrue
Private Const UD_FALSE As Long = 0              ' msoFalse
Private Const UD_ZORDER_BACK As Long = 1        ' msoSendToBack
Private Const UD_ZORDER_FRONT As Long = 0       ' msoBringToFront

' 色(RGB(r,g,b) = r + g*256 + b*65536)。本製品の主色。
Private Const UD_COLOR_BRAND As Long = 5990145&     ' RGB(1,103,91)   主色
Private Const UD_COLOR_BRAND_DK As Long = 5205761&  ' RGB(1,77,79)    主色(濃)
Private Const UD_COLOR_PANEL As Long = 16777215&    ' 白
Private Const UD_COLOR_PANEL_LN As Long = 14939616& ' RGB(224,231,227)
Private Const UD_COLOR_TEXT As Long = 2562065&      ' RGB(17,24,39)
Private Const UD_COLOR_WHITE As Long = 16777215&
Private Const UD_COLOR_FOCUS As Long = 47359&       ' RGB(255,184,0) 橙(強調枠)
' 現場メモの見出し行の地色(11章§3.3.3(b) の「灰色地」。RGB(235,235,235))。
Private Const UD_COLOR_HEAD_BG As Long = 15461355&
Private Const UD_FONT As String = "Yu Gothic UI"

' コーチ帯の文字(裁定書26 A)。白・太字・左寄せ・上下中央、内側余白12pt。
Private Const UD_BAND_FONT_PT As Double = 13#
Private Const UD_BAND_PAD As Double = 12#

' 強調枠(11章§3.1.1)。枠線のみ・塗りなし・線幅2.25pt・**点滅させない**。
Private Const UD_FOCUS_WEIGHT As Double = 2.25
Private Const UD_FOCUS_PAD As Double = 4#

' 待ちカード(11章§4.2(2))。幅420pt。高さは文字量から実測する。
Private Const UD_WAIT_WIDTH As Double = 420#
Private Const UD_WAIT_COLS As Long = 30
Private Const UD_WAIT_LINE_PT As Double = 15#
Private Const UD_WAIT_PAD_PT As Double = 28#
Private Const UD_WAIT_MIN_PT As Double = 120#
Private Const UD_WAIT_MAX_PT As Double = 260#
Private Const UD_WAIT_NOTE As String = _
    "この画面は20分ほど固まったように見えることがあります。閉じずにお待ちください。"

' ボタンの並び。列アンカーの実測左端から求め、列幅を仮定しない。
Private Const UD_BTN_COL_FIRST As Long = 4
Private Const UD_BTN_COL_LAST As Long = 60
Private Const UD_BTN_GAP As Double = 8#

' 状態行・プレビューの規約。
Private Const UD_PREVIEW_LINES As Long = 5
Private Const UD_MSG_NOT_YET As String = "まだ貼っていません"

' ============================================================================
' DrawCoachBar - コーチ帯(11章§3.1.1)。最上段3行はビルドが枠固定してある。
'   セルへ書くのは STEP番号・進捗ドット・いま何をするかの1文の3つだけで、
'   図形は「帯の地」と3行目のボタン4本だけにする(セル1つで済むものを図形に
'   しない=再描画が要らない。§8.5 #2 #3)。
' ============================================================================
Public Sub DrawCoachBar(ByVal stepNo As Long, ByVal stepCount As Long, _
                        ByVal actionText As String)
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UD_SHEET)
    If ws Is Nothing Then Exit Sub

    modUISheet.WriteNamed "nv_step_no", "STEP " & CStr(stepNo) & "/" & CStr(stepCount)
    modUISheet.WriteNamed "nv_step_dots", modUIGeom.StepDots(stepNo, stepCount)
    modUISheet.WriteNamed "hm_next_action", actionText

    ' 帯の地(2色グラデーション)。最背面へ送り、セルの文字を隠さない。
    Dim topRow As Long
    topRow = modUISheet.BlockRow("nv_step_no")
    If topRow <= 0 Then Exit Sub

    Dim btnRow As Long
    btnRow = modUISheet.BlockRow("hm_next_action")
    If btnRow <= 0 Then btnRow = topRow

    Dim band As Object
    Set band = ws.Shapes.AddShape(UD_SHAPE_ROUNDED, ws.Cells(topRow, 1).Left, _
                                  ws.Cells(topRow, 1).Top, 520#, _
                                  ws.Cells(btnRow + 1, 1).Top - ws.Cells(topRow, 1).Top)
    If Not band Is Nothing Then
        band.Name = UD_BAND_PREFIX & "_bg"
        band.Placement = UD_PLACEMENT_FREE
        band.Adjustments(1) = 0.06
        band.Line.Visible = UD_FALSE
        modUISheet.ApplyBandGradient band, UD_COLOR_BRAND, UD_COLOR_BRAND_DK
        modUISheet.ApplyLightShadow band
        ' 帯の文字は**図形の中に持たせる**(裁定書26 A・11章§8.5 の禁忌)。
        ' 図形は色や重ね順に関わらず常にセルの上へ描かれるため、セルへ書いた
        ' STEP・ドット・次の一手は帯の下に隠れて実機で1文字も見えない。
        band.TextFrame.Characters.Text = _
            modUIGeom.CoachBandText(stepNo, stepCount, actionText)
        band.TextFrame.HorizontalAlignment = UD_ALIGN_LEFT
        band.TextFrame.VerticalAlignment = UD_ANCHOR_MIDDLE
        band.TextFrame.MarginLeft = UD_BAND_PAD
        band.TextFrame.MarginRight = UD_BAND_PAD
        band.TextFrame.MarginTop = UD_BAND_PAD
        band.TextFrame.MarginBottom = UD_BAND_PAD
        band.TextFrame.Characters.Font.Name = UD_FONT
        band.TextFrame.Characters.Font.Size = UD_BAND_FONT_PT
        band.TextFrame.Characters.Font.Bold = True
        band.TextFrame.Characters.Font.Color = UD_COLOR_WHITE
        band.ZOrder UD_ZORDER_BACK
    End If

    EnsureCoachButtons ws, btnRow
End Sub

' 帯の3行目のボタン4本(13章§2.10(f))。配置表の値源は modUINav。
'   **ここで btn_nv_ をまとめて落とさない**。同じ接頭辞を[コピー]8本と
'   [＋ もっと調べる]も使っており、帯を描くたびにそれらが消えてしまうため。
'   孤児の一括削除は DropNavShapes が描画の**いちばん最初**に1回だけ行う。
Private Sub EnsureCoachButtons(ByVal ws As Object, ByVal rowNo As Long)
    On Error Resume Next
    PlaceButtonRow ws, rowNo, modUINav.NavRowCoach(), "plain"
End Sub

' DropNavShapes - ナビの図形の孤児を落とす(描き直しの**いちばん最初**に1回)。
'   名前を配列へ集めてから削除する方式(modUISheet.DropShapesByPrefix)であり、
'   Shapes を列挙しながら削除しない(11章§8.6 禁忌3)。
Public Sub DropNavShapes()
    On Error Resume Next
    Dim ws As Object
    Set ws = modUISheet.SheetOf(UD_SHEET)
    If ws Is Nothing Then Exit Sub
    modUISheet.DropShapesByPrefix ws, UD_BTN_PREFIX
    modUISheet.DropShapesByPrefix ws, UD_PANEL_PREFIX
    modUISheet.DropShapesByPrefix ws, UD_BAND_PREFIX
End Sub

' ============================================================================
' DrawSections - 4区画の角丸パネルと、区画ごとのボタン(11章§3.1)。
'   パネルは背面の角丸四角で、セルはその上に載る(モックの .panel の再現)。
'   高さは**区画の見出し行から次の区画の見出し行まで**を実測して決める
'   (AutoSize を使わない。行高・列幅を仮定しない)。
' ============================================================================
Public Sub DrawSections()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UD_SHEET)
    If ws Is Nothing Then Exit Sub

    DrawOnePanel ws, "nv_sec1", "nv_sec2"
    DrawOnePanel ws, "nv_sec2", "nv_sec3"
    DrawOnePanel ws, "nv_sec3", "nv_sec4"
    DrawOnePanel ws, "nv_sec4", vbNullString

    ' 区画①③④のボタン(区画②の18本は modUICase6 の表が値源)。
    ' 区画①の見出し行の直下=調べる場所の案内1行(裁定書26 C)。
    Dim r1b As Long
    r1b = modUISheet.BlockRow("nv_sec1")
    If r1b > 0 Then PlaceButtonRow ws, r1b + 1, modUINav.NavRowSec1B(), "plain"

    Dim r1 As Long
    r1 = modUISheet.BlockRow("dr_copied_seq")
    If r1 > 0 Then PlaceButtonRow ws, r1, modUINav.NavRowSec1(), "plain"

    Dim r2 As Long
    r2 = modUISheet.BlockRow("ci_count_total")
    If r2 > 0 Then PlaceButtonRow ws, r2, modUINav.NavRowSec2(), "primary"

    Dim r3 As Long
    r3 = modUISheet.BlockRow("nv_sec3")
    If r3 > 0 Then PlaceButtonRow ws, r3 + 1, modUINav.NavRowSec3(), "primary"

    Dim r4 As Long
    r4 = modUISheet.BlockRow("nv_sec4")
    If r4 > 0 Then PlaceButtonRow ws, r4 + 1, modUINav.NavRowSec4(), "primary"
End Sub

' 1区画ぶんの角丸パネル。endAnchor が空なら見出しから12行ぶんを目安にする。
Private Sub DrawOnePanel(ByVal ws As Object, ByVal headAnchor As String, _
                         ByVal endAnchor As String)
    On Error Resume Next

    Dim r1 As Long
    r1 = modUISheet.BlockRow(headAnchor)
    If r1 <= 0 Then Exit Sub

    Dim r2 As Long
    If LenB(endAnchor) > 0 Then
        r2 = modUISheet.BlockRow(endAnchor) - 1
    Else
        r2 = r1 + 12
    End If
    If r2 <= r1 Then Exit Sub

    Dim topPt As Double
    Dim botPt As Double
    topPt = ws.Cells(r1, 1).Top
    botPt = ws.Cells(r2, 1).Top
    If botPt <= topPt Then Exit Sub

    ' 区画②は1,600行を超える(直貼り枠300行×5)。全体を1枚のパネルで囲むと
    ' 図形が縦に長くなりすぎて描画が重くなるため、上限で丸める。
    If botPt - topPt > 900# Then botPt = topPt + 900#

    Dim shp As Object
    Set shp = ws.Shapes.AddShape(UD_SHAPE_ROUNDED, ws.Cells(r1, 1).Left, topPt, _
                                 560#, botPt - topPt)
    If shp Is Nothing Then Exit Sub
    shp.Name = UD_PANEL_PREFIX & "_" & headAnchor
    shp.Placement = UD_PLACEMENT_FREE
    shp.Adjustments(1) = 0.02
    shp.Fill.ForeColor.RGB = UD_COLOR_PANEL
    shp.Fill.Transparency = 1#                  ' 面は透明にしてセルを隠さない
    shp.Line.Visible = UD_TRUE
    shp.Line.ForeColor.RGB = UD_COLOR_PANEL_LN
    modUISheet.ApplyLightShadow shp
    shp.ZOrder UD_ZORDER_BACK
End Sub

' 1行ぶんのボタンを左から右へ、重ならない位置に置く(modUIHome と同じ作法)。
'   specText の1件 = "図形名;キャプション;OnAction;幅pt" を vbLf 区切り。
Public Sub PlaceButtonRow(ByVal ws As Object, ByVal rowNo As Long, _
                          ByVal specText As String, ByVal kind As String)
    ' W9.2 N8: 起動時の描画経路(EnsureScreens -> DrawNav -> DrawSections)の末端。
    ' ここで投げると起動直後に生ダイアログが出るので、必ず受け止めて記録する。
    On Error GoTo Failed

    If rowNo <= 0 Then Exit Sub
    If LenB(specText) = 0 Then Exit Sub

    Dim specs() As String
    specs = Split(specText, vbLf)

    Dim col As Long
    Dim minLeft As Double
    col = UD_BTN_COL_FIRST
    minLeft = modUISheet.CellLeft(ws, rowNo, UD_BTN_COL_FIRST)
    If minLeft < 0 Then Exit Sub

    Dim i As Long
    For i = LBound(specs) To UBound(specs)
        Dim flds() As String
        flds = Split(specs(i), ";")
        If UBound(flds) - LBound(flds) >= 3 Then
            Dim widthPt As Double
            widthPt = Val(flds(3))
            col = FitCol(ws, rowNo, col, minLeft)
            If col <= 0 Then Exit Sub
            modUISheet.EnsureButtonEx ws, flds(0), flds(1), rowNo, col, widthPt, _
                                      flds(2), kind
            minLeft = modUISheet.CellLeft(ws, rowNo, col) + widthPt + UD_BTN_GAP
        End If
    Next i
    Exit Sub

Failed:
    ' ハンドラ稼働中は On Error Resume Next が効かないので、記録は別Subへ。
    LogPlaceFailure rowNo, Err.Number
End Sub

' PlaceButtonRow の Failed: から呼ぶ記録専用(W9.2 N8)。
Private Sub LogPlaceFailure(ByVal rowNo As Long, ByVal errNo As Long)
    On Error Resume Next
    modLog.LogError "E0603", UD_SRC & ".PlaceButtonRow", _
                    "place_button_row_failed:" & CStr(rowNo), errNo
End Sub

' 左端が minLeft 以上になる最初のアンカー列(0=範囲内に無い)。
Private Function FitCol(ByVal ws As Object, ByVal rowNo As Long, ByVal fromCol As Long, _
                        ByVal minLeft As Double) As Long
    Dim c As Long
    Dim x As Double
    For c = fromCol To UD_BTN_COL_LAST
        x = modUISheet.CellLeft(ws, rowNo, c)
        If x < 0 Then Exit Function
        If x >= minLeft Then
            FitCol = c
            Exit Function
        End If
    Next c
End Function

' ============================================================================
' MoveFocusFrame - 黄色い強調枠(11章§3.1.1・§8.5 #5)。
'   図形は**常に1つだけ**(描く前に必ず削除する)。枠線のみ・塗りなし・
'   線色は橙・線幅2.25pt。**点滅させない**(§8.5 #6: タイマーを増やさない)。
'   クリックできない(OnAction を付けない。原則①)。
' ============================================================================
Public Sub MoveFocusFrame(ByVal anchorName As String)
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UD_SHEET)
    If ws Is Nothing Then Exit Sub

    modUISheet.DropShape ws, UD_FOCUS
    If LenB(anchorName) = 0 Then Exit Sub

    Dim r1 As Long
    r1 = modUISheet.BlockRow(anchorName)
    If r1 <= 0 Then Exit Sub

    Dim shp As Object
    Set shp = ws.Shapes.AddShape(UD_SHAPE_ROUNDED, ws.Cells(r1, 1).Left - UD_FOCUS_PAD, _
                                 ws.Cells(r1, 1).Top - UD_FOCUS_PAD, _
                                 568#, ws.Rows(r1).RowHeight + UD_FOCUS_PAD * 2#)
    If shp Is Nothing Then Exit Sub
    shp.Name = UD_FOCUS
    shp.Placement = UD_PLACEMENT_FREE
    shp.Adjustments(1) = 0.06
    shp.Fill.Visible = UD_FALSE
    shp.Line.Visible = UD_TRUE
    shp.Line.ForeColor.RGB = UD_COLOR_FOCUS
    shp.Line.Weight = UD_FOCUS_WEIGHT
    shp.ZOrder UD_ZORDER_FRONT
End Sub

' ============================================================================
' DrawWaitCard - 待ち時間の(2)事前描画カード(11章§4.2(2))。
' ----------------------------------------------------------------------------
' 実装規約(順序が正・1つでも入れ替えない):
'   1. 既存の nv_wait* を全て削除する(名前を配列へ集めてから削除)
'   2. ActiveWindow.VisibleRange から中央を求めて置く(列幅・行高・DPIを仮定しない)
'   3. 段の進捗ドットは図形の中の文字で書く(アニメーションさせない)
'   4. ScreenUpdating = True を明示してから DoEvents を1回入れ、**描画が画面に
'      出たことを確定させてから**リボンを呼ぶ(ここを飛ばすと1度も表示されない)
'   5. 戻ったらカードを削除する(失敗経路でも必ず。呼び出し側の責務)
' ============================================================================
Public Sub DrawWaitCard(ByVal stageName As String, ByVal startedAt As String, _
                        ByVal maxWaitText As String, ByVal stageNo As Long, _
                        ByVal stageCount As Long)
    On Error Resume Next

    HideWaitCard

    Dim ws As Object
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Dim bodyText As String
    bodyText = "いま作っています。" & vbLf & vbLf & _
               stageName & vbLf & _
               "開始 " & startedAt & " ／ 最大 " & maxWaitText & vbLf & vbLf & _
               modUIGeom.StepDots(stageNo, stageCount) & _
               "　（" & CStr(stageCount) & "段のうち" & CStr(stageNo) & "段目）" & vbLf & vbLf & _
               "画面が白くなっても、止まっていません。" & vbLf & _
               "そのままお待ちください。" & vbLf & vbLf & UD_WAIT_NOTE

    Dim heightPt As Double
    heightPt = modUIGeom.CardHeightFor(bodyText, UD_WAIT_COLS, UD_WAIT_LINE_PT, _
                                       UD_WAIT_PAD_PT, UD_WAIT_MIN_PT, UD_WAIT_MAX_PT)

    Dim leftPt As Double
    Dim topPt As Double
    leftPt = CardLeft(heightPt)
    topPt = CardTop(heightPt)

    Dim card As Object
    Set card = ws.Shapes.AddShape(UD_SHAPE_ROUNDED, leftPt, topPt, UD_WAIT_WIDTH, heightPt)
    If card Is Nothing Then Exit Sub
    card.Name = UD_WAIT_PREFIX & "_card"
    card.Placement = UD_PLACEMENT_FREE
    card.Adjustments(1) = 0.05
    card.Fill.ForeColor.RGB = UD_COLOR_WHITE
    card.Line.Visible = UD_TRUE
    card.Line.ForeColor.RGB = UD_COLOR_BRAND
    card.Line.Weight = 1.5
    card.TextFrame.Characters.Text = bodyText
    card.TextFrame.HorizontalAlignment = UD_ALIGN_LEFT
    card.TextFrame.VerticalAlignment = UD_ANCHOR_MIDDLE
    card.TextFrame.Characters.Font.Name = UD_FONT
    card.TextFrame.Characters.Font.Size = 11#
    card.TextFrame.Characters.Font.Color = UD_COLOR_TEXT
    modUISheet.ApplySoftShadow card
    card.ZOrder UD_ZORDER_FRONT

    ' 4: 描き切ったことを確定させてから呼び出し側がリボンを呼ぶ。
    Application.ScreenUpdating = True
    DoEvents
End Sub

' HideWaitCard - カードを消す。**失敗経路でも必ず通す**(11章§4.2(2) 規約5)。
Public Sub HideWaitCard()
    On Error Resume Next
    Dim ws As Object
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    modUISheet.DropShapesByPrefix ws, UD_WAIT_PREFIX
End Sub

' カードの左端・上端(表示領域の中央へ寄せる。取れないときは既定値)。
Private Function CardLeft(ByVal heightPt As Double) As Double
    On Error GoTo Fallback1
    Dim x As Double
    x = ActiveWindow.VisibleRange.Left + _
        (ActiveWindow.VisibleRange.Width - UD_WAIT_WIDTH) / 2#
    If x < 24# Then x = 24#
    CardLeft = x
    Exit Function
Fallback1:
    CardLeft = 24#
End Function

Private Function CardTop(ByVal heightPt As Double) As Double
    On Error GoTo Fallback2
    Dim y As Double
    y = ActiveWindow.VisibleRange.Top + _
        (ActiveWindow.VisibleRange.Height - heightPt) / 2#
    If y < 24# Then y = 24#
    CardTop = y
    Exit Function
Fallback2:
    CardTop = 24#
End Function

' ============================================================================
' 区画②の描画(状態行・プレビュー・3ボタン・出る条件)
' ============================================================================

' RefreshAreas - 7欄すべての状態行とプレビューと合計字数を書き直す(v2.6)。
Public Sub RefreshAreas(ByVal caseId As String)
    On Error Resume Next

    Dim keys() As String
    keys = Split(modUICase6.AreaKeys(), vbLf)

    Dim i As Long
    Dim total As Long
    For i = LBound(keys) To UBound(keys)
        total = total + RefreshArea(caseId, keys(i))
    Next i

    Dim maxChars As Long
    maxChars = modConfig.GetLong("t2_max_context_chars", 100000)
    modUISheet.WriteNamed "ci_count_total", "全部で " & Format$(total, "#,##0") & _
        "字 / 目安 " & Format$(maxChars, "#,##0") & "字"
    modUISheet.MarkNamed "ci_count_total", (total > maxChars)
End Sub

' RefreshArea - 1欄ぶん。戻り値はその欄の字数(合計に足す)。
Public Function RefreshArea(ByVal caseId As String, ByVal areaKey As String) As Long
    On Error Resume Next

    Dim countRange As String
    countRange = modUICase6.AreaField(areaKey, 6)
    If LenB(countRange) = 0 Then Exit Function

    Dim body As String
    body = modUICase6.AreaBody(caseId, areaKey)

    ' 現場メモの枠は見出しと例文が先置きされているので、ひな型のままなら
    ' 「まだ書いていません」と出す(字数だけを見ると常に「入っている」になる)。
    If StrComp(areaKey, "field_notes", vbBinaryCompare) = 0 Then
        If Not modUINav.FieldNotesWritten() Then body = vbNullString
    End If

    If LenB(body) = 0 Then
        modUISheet.WriteNamed countRange, UD_MSG_NOT_YET
    Else
        modUISheet.WriteNamed countRange, _
            PastedLine(Len(body), modUISheet.ReadNamed(countRange))
    End If

    Dim prevRange As String
    prevRange = modUICase6.AreaField(areaKey, 3)
    If LenB(prevRange) > 0 Then WritePreview prevRange, body
    RefreshArea = Len(body)
End Function

' ============================================================================
' PastedLine - 状態行の1行(11章§3.3.2 の逐語「貼り付け済み　9,812字　
'   （14:03 に貼りました）」)。
' ----------------------------------------------------------------------------
' 裁定書22 m1: **字数が変わらない描き直しでは時刻を書き換えない**。以前は
'   描き直すたびに Now を入れていたため、3日前に貼った欄が「たったいま貼った」
'   ように見えていた(貼った時刻は貼った事実の記録であって描画時刻ではない)。
'   前の状態行から字数と時刻を読み、字数が同じならその時刻を持ち越す。
' ============================================================================
Private Function PastedLine(ByVal charCount As Long, ByVal prevLine As String) As String
    Dim timeText As String
    timeText = KeptTime(charCount, prevLine)
    If LenB(timeText) = 0 Then timeText = Format$(Now, "hh:nn")
    PastedLine = "貼り付け済み　" & Format$(charCount, "#,##0") & _
                 "字　（" & timeText & " に貼りました）"
End Function

' 前の状態行が同じ字数を示していれば、その時刻(hh:nn)を返す。違えば ""。
Private Function KeptTime(ByVal charCount As Long, ByVal prevLine As String) As String
    On Error Resume Next
    If LenB(prevLine) = 0 Then Exit Function

    Dim wantHead As String
    wantHead = "貼り付け済み　" & Format$(charCount, "#,##0") & "字　（"
    If InStr(1, prevLine, wantHead, vbBinaryCompare) <> 1 Then Exit Function

    Dim tail As String
    tail = Mid$(prevLine, Len(wantHead) + 1)
    Dim pos As Long
    pos = InStr(1, tail, " ", vbBinaryCompare)
    If pos <= 1 Then Exit Function
    KeptTime = Left$(tail, pos - 1)
End Function

' プレビュー5行を書く(6行目以降は書かない。足りない行は空にする)。
Private Sub WritePreview(ByVal prevRange As String, ByVal body As String)
    On Error Resume Next
    Dim head As Object
    Set head = modUISheet.NamedCell(prevRange)
    If head Is Nothing Then Exit Sub

    Dim ws As Object
    Set ws = head.Worksheet

    Dim lines() As String
    lines = Split(modNavText.PreviewLines(body, UD_PREVIEW_LINES), vbLf)

    Dim i As Long
    For i = 1 To UD_PREVIEW_LINES
        Dim s As String
        s = vbNullString
        If LenB(body) > 0 Then
            If i - 1 <= UBound(lines) Then s = lines(i - 1)
        End If
        modUISheet.PutText ws, head.row + i - 1, head.Column, s, UD_SRC & "/" & prevRange
    Next i
End Sub

' EnsureAreaButtons - 6欄×3種の図形ボタン(13章§2.11(b))。
'   置き直す前に ci_btn_ をまとめて落とす(名前を集めてから消す)。
'   **未貼付の欄には[消す]を出さない**(押しても何も起きないボタンを置かない)。
Public Sub EnsureAreaButtons(ByVal caseId As String)
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UD_SHEET)
    If ws Is Nothing Then Exit Sub

    modUISheet.DropShapesByPrefix ws, "ci_btn_"

    Dim keys() As String
    keys = Split(modUICase6.AreaKeys(), vbLf)

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        If Not modUICase6.AreaHidden(keys(i)) Then
            Dim rowNo As Long
            rowNo = modUISheet.BlockRow(modUICase6.AreaField(keys(i), 6))
            If rowNo > 0 Then
                ' 裁定書22 m5: 現場メモは枠に見出しと例文が先置きされているので
                '   字数で見ると常に「入っている」になり、まだ書いていない欄に
                '   [消す]が出ていた。判定は FieldNotesWritten() に一本化する。
                Dim hasBody As Boolean
                If StrComp(keys(i), "field_notes", vbBinaryCompare) = 0 Then
                    hasBody = modUINav.FieldNotesWritten()
                Else
                    hasBody = (LenB(modUICase6.AreaBody(caseId, keys(i))) > 0)
                End If
                PlaceAreaRow ws, rowNo, keys(i), hasBody
            End If
        End If
    Next i
End Sub

' 1欄ぶんの3ボタンを、重ならない位置へ左から置く。
Private Sub PlaceAreaRow(ByVal ws As Object, ByVal rowNo As Long, _
                         ByVal areaKey As String, ByVal hasBody As Boolean)
    Dim specText As String
    specText = "ci_btn_paste_" & areaKey & ";ここに貼る;modUICase6." & _
               modUICase6.HandlerName("PasteInto", areaKey) & ";96" & vbLf & _
               "ci_btn_show_" & areaKey & ";中身を見る;modUICase6." & _
               modUICase6.HandlerName("ShowArea", areaKey) & ";96"
    If hasBody Then
        specText = specText & vbLf & "ci_btn_clear_" & areaKey & ";消す;modUICase6." & _
                   modUICase6.HandlerName("ClearArea", areaKey) & ";64"
    End If
    PlaceButtonRow ws, rowNo, specText, "plain"
End Sub

' ApplyAreaVisibility - 出る条件が偽の欄を行ごと隠す(中身は消さない)。
Public Sub ApplyAreaVisibility()
    On Error Resume Next
    ' v2.6(裁定書25 S1): 「いまの契約」は常時表示になった。旧版で隠れたままの
    '   行を出し直すため、False を渡して明示的に表示へ戻す(呼び続ける)。
    HideAreaRows "contract", modUICase6.AreaHidden("contract")
    HideAreaRows "finance", modUICase6.AreaHidden("finance")
    HideAreaRows "hearing_answers", modUICase6.AreaHidden("hearing_answers")
End Sub

' 1欄ぶんの行(見出し・説明・例文2行・状態行からプレビュー・直貼り枠・見張り行
'   まで)を1かたまりで隠す/出す。行番号は名前付きレンジから実測する。
Private Sub HideAreaRows(ByVal areaKey As String, ByVal hideIt As Boolean)
    On Error Resume Next

    Dim topRow As Long
    Dim endRow As Long
    topRow = modUISheet.BlockRow(modUICase6.AreaField(areaKey, 6)) - 4
    endRow = modUISheet.BlockRow(modUICase6.AreaField(areaKey, 5))
    If topRow <= 0 Or endRow < topRow Then Exit Sub

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UD_SHEET)
    If ws Is Nothing Then Exit Sub
    ws.Rows(topRow & ":" & endRow).Hidden = hideIt
End Sub

' ============================================================================
' 現場メモの先置き(11章§3.3.3(b))
' ----------------------------------------------------------------------------
' 枠には最初から4つの見出しと例文が入っている。営業は**見出しの下に書き足す
' だけ**でよい。例文は「例: 」で始まる行で、保存時に落ちる(modNavText)。
' 見出しを消しても壊れない(見出しが1つも無ければ全文が営業メモへ入る)。
' ============================================================================
Public Function FieldNotesTemplate() As String
    Dim s As String
    s = s & "【営業メモ】" & vbLf
    s = s & "  例: 社長はワンマンで決裁は即断。3年前から毎年訪問している。" & vbLf
    s = s & "【前回更新メモ】" & vbLf
    s = s & "  例: 前回は火災だけ更新。地震は見送りになった。" & vbLf
    s = s & "【付保の見立て】" & vbLf
    s = s & "  例: たぶん火災は他社さん。労災上乗せは元請の包括に乗っている模様。" & vbLf
    s = s & "【そのほか】" & vbLf
    s = s & "  例: 工場の裏の川が気になる。◯◯損保の営業が最近よく来ているらしい。"
    FieldNotesTemplate = s
End Function

' ResetForNewCase - 案件が変わったときに画面の枠を入れ替える(裁定書13 W1と同じ
'   趣旨: 前の案件を描いた画面のまま次の案件を確定させない)。
'   **case_data は触らない**(消すのは画面の枠だけ。原則⑤)。
'   現場メモの枠だけは**その案件の保存済みの内容を読み戻す**。この欄は画面が
'   唯一の入力面なので、案件を切り替えたときに空へ戻すと書いたものが消える。
Public Sub ResetForNewCase(ByVal caseId As String)
    On Error Resume Next

    Dim keys() As String
    keys = Split(modUICase6.AreaKeys(), vbLf)

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim rawRange As String
        rawRange = modUICase6.AreaField(keys(i), 4)
        If LenB(rawRange) > 0 Then
            If StrComp(keys(i), "field_notes", vbBinaryCompare) <> 0 Then
                ThisWorkbook.Names(rawRange).RefersToRange.ClearContents
            End If
        End If
        Dim sentRange As String
        sentRange = modUICase6.AreaField(keys(i), 5)
        If LenB(sentRange) > 0 Then modUISheet.WriteNamed sentRange, vbNullString
    Next i

    LoadFieldNotesFor caseId
End Sub

' LoadFieldNotesFor - 現場メモの枠を、その案件の保存済みの内容で描き直す。
'   まだ何も書かれていない案件では**見出しと例文を先置きしたひな型**を入れる
'   (11章§3.3.3(b))。何か書いてある案件では例文を出さない
'   (「書き始めたら例文は消える」の実装。消える単位は枠ぜんたい)。
Public Sub LoadFieldNotesFor(ByVal caseId As String)
    On Error Resume Next

    If LenB(caseId) = 0 Then
        modUICase6.WriteFieldNotesArea FieldNotesTemplate()
        StyleFieldNotesHeads
        Exit Sub
    End If

    Dim memoText As String
    Dim othersText As String
    memoText = modUICase6.LoadArea(caseId, "input_memo")
    othersText = modUICase6.LoadArea(caseId, "input_field_notes")

    If LenB(memoText) = 0 And LenB(othersText) = 0 Then
        modUICase6.WriteFieldNotesArea FieldNotesTemplate()
        StyleFieldNotesHeads
        Exit Sub
    End If
    modUICase6.WriteFieldNotesArea modNavText.JoinFieldNotes(memoText, othersText)
    StyleFieldNotesHeads
End Sub

' ============================================================================
' StyleFieldNotesHeads - 現場メモ枠の見出し行(【…】)を灰色地・Locked=True に
'   する(11章§3.3.3(b)・裁定書22 D10)。利用者が見出しを消せないようにし、
'   「どこに書けばよいか」を地色で示す。見出し以外の行は書ける状態へ戻す
'   (前回の描画で見出しだった行が本文になったときに固まらないようにする)。
' ============================================================================
Public Sub StyleFieldNotesHeads()
    On Error Resume Next

    Dim head As Object
    Set head = modUISheet.NamedCell("ci_area_field_notes")
    If head Is Nothing Then Exit Sub

    Dim rows As Long
    rows = modUISheet.NamedRows("ci_area_field_notes")
    If rows <= 0 Then Exit Sub

    Dim ws As Object
    Set ws = head.Worksheet

    Dim i As Long
    For i = 1 To rows
        Dim cell As Object
        Set cell = ws.Cells(head.row + i - 1, head.Column)
        If IsHeadLine(CStr(cell.Value)) Then
            cell.Interior.Color = UD_COLOR_HEAD_BG
            cell.Locked = True
        Else
            cell.Interior.ColorIndex = -4142        ' xlColorIndexNone
            cell.Locked = False
        End If
    Next i
End Sub

' 【…】だけの行か(前後の空白を除いた完全一致の形)。
Private Function IsHeadLine(ByVal lineText As String) As Boolean
    Dim s As String
    s = Trim$(lineText)
    If Len(s) < 3 Then Exit Function
    If StrComp(Left$(s, 1), "【", vbBinaryCompare) <> 0 Then Exit Function
    IsHeadLine = (StrComp(Right$(s, 1), "】", vbBinaryCompare) = 0)
End Function

' ShowMoreRows - 区画①の補助5本([＋ もっと調べる])の行を出す/隠す(11章§3.2)。
'   閉じている間は行ごと非表示にし、図形は作らない。
Public Sub ShowMoreRows(ByVal showIt As Boolean)
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UD_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim topRow As Long
    Dim endRow As Long
    topRow = modUISheet.BlockRow("dr_copied_seq") + 1
    endRow = modUISheet.BlockRow("dr_body_08")
    If topRow <= 0 Or endRow < topRow Then Exit Sub
    ws.Rows(topRow & ":" & endRow).Hidden = Not showIt
End Sub
