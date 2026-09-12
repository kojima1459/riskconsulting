Attribute VB_Name = "modUIToast"
Option Explicit

' ============================================================================
' modUIToast - 押したことが必ず見えるトースト(ui層・裁定書17 H2)
' ----------------------------------------------------------------------------
' なぜ要るのか(実機フィードバック第2報):
'   HOMEのボタンを押しても「セルがA1へ動くだけ」で、案内を書いた hm_warning の
'   セルは利用者の目に入っていなかった。**押せるものは必ず反応する**を、
'   セルではなく「画面の右上に出て自分で消えるカード」で成り立たせる。
'
' 移植元: PoC「マイ本棚AI」 src/ui/modSkin.bas の ShowToast(図形カード＋
'   TextFrame2 の AutoSize による実測高さ)。RPNでは待機ループ(DoEvents)を使わず
'   **Application.OnTime による自動消去**へ置き換えた(呼び出し側を待たせない)。
'
' 約束:
'   1. 図形名は全て "ts_" 接頭辞。既存の btn_ / lbl_ / btncopy_ / gt_ と衝突
'      しない。削除は modUISheet.DropShapesByPrefix に委ねる。
'   2. **全Publicの先頭で On Error Resume Next**。トーストの失敗は「出ない・
'      消えない」に留め、本体機能へ波及させない(modUIGuide と同じ作法)。
'   3. 連続表示は前のカードを消してから置く(重ならない・最後の1枚だけが残る)。
'   4. 予約は常に1本(ShowToast のたびに前の予約を取り消してから張り直す)。
'   5. **HideToast は対象図形が無ければ何もしない**。ブックを閉じたあとに予約が
'      残って発火しても、消すものが無ければ黙って戻る(Excelが本ブックを開き直す
'      経路を作らない)。**CancelToast は終了処理から呼ぶ**
'      (ThisWorkbook.Workbook_BeforeClose が RestoreScreen の前に呼ぶ。W9.2/W9.3)。
'   6. 自ブックがアクティブなときだけ描く(他人のブックへ図形を作らない)。
'
' 文言はここが唯一の値源(modUIHome は残量が少ないため文言定数を持たない。
'   裁定書17 H2)。hm_warning へ書く1行の組み立て(WarnLine)も本モジュールが持つ。
' ============================================================================

' Shape名(接頭辞は1つ。消し込みはこの接頭辞でまとめて行う)。
Private Const UT_PREFIX As String = "ts_"
Private Const UT_CARD As String = "ts_card"

' 【廃止】UT_GUIDE_SHEET / UT_CH7_ANCHOR / UT_LOCK_NAME / UT_MSG_RESEARCH と
' ShowResearchPrompts は W6.1(裁定書22)で撤去した。v3.2 で HOME そのものが無く
' なり、調べる文はナビの区画①にその場で出るため、使い方タブ⑦へ飛ばす導線は
' 二重動線になっていた(押す口も無かった)。14章§6は「廃止」として残す。

' 見た目。Yu Gothic UI 11pt・角丸カード・右上に固定幅。
Private Const UT_FONT As String = "Yu Gothic UI"
Private Const UT_FONT_SIZE As Double = 11#
Private Const UT_WIDTH As Double = 400#
Private Const UT_MIN_H As Double = 40#
Private Const UT_MAX_H As Double = 132#
Private Const UT_EDGE As Double = 24#

' Excel組み込み定数の数値(名前を書かず、LibreOffice側の構文チェックで未定義名に
' ならないようにする。modUIGuide / modUISheet と同じ流儀)。
Private Const UT_SHAPE_ROUNDED As Long = 5      ' msoShapeRoundedRectangle
Private Const UT_PLACEMENT_FREE As Long = 3     ' xlFreeFloating
Private Const UT_ALIGN_LEFT As Long = 1         ' msoAlignLeft
Private Const UT_ANCHOR_MIDDLE As Long = 3      ' msoAnchorMiddle
Private Const UT_WRAP_ON As Long = -1           ' msoTrue
Private Const UT_AUTOSIZE_ON As Long = 1        ' msoAutoSizeShapeToFitText
Private Const UT_AUTOSIZE_OFF As Long = 0       ' msoAutoSizeNone
Private Const UT_FRONT As Long = 0              ' msoBringToFront

' 配色。値は「RGB(r,g,b) = r + g*256 + b*65536」(modUIGuide と同じ流儀)。
' コントラスト比(WCAG相対輝度で計算。いずれも 4.5:1 以上):
'   info  文字 RGB(17,24,39) / 地 白           = 16.9:1
'   warn  文字 RGB(17,24,39) / 地 RGB(255,244,214) = 16.6:1
'   error 文字 白            / 地 RGB(185,28,28)   =  6.4:1
Private Const UT_WHITE As Long = 16777215&        ' 白
Private Const UT_TEXT As Long = 2562065&          ' 文字 RGB(17,24,39)
Private Const UT_ACCENT As Long = 5990145&        ' 濃緑 RGB(1,103,91)
Private Const UT_WARN_BG As Long = 14087423&      ' 淡黄 RGB(255,244,214)
Private Const UT_WARN_LINE As Long = 30900&       ' 黄枠 RGB(180,120,0)
Private Const UT_ERROR_BG As Long = 1842361&      ' 赤   RGB(185,28,28)

' 失敗の警告文の末尾に足す案内(裁定書17 H4)。err_log を可視にしたので、
' 利用者が自分で最後の1行を送れる。
Private Const UT_ERRLOG_NOTE As String = _
    "（err_logタブの最後の行を開発担当へ送ってください）"

' 予約済み OnTime の時刻(0=予約なし)。永続しない画面制御変数であり
' 14章§6「状態保持の例外」には当たらない(modUIGuide の gStep と同じ)。
Private gHideAt As Date
' 直近に出した文(表示秒数を文の長さから決めるため。裁定書22 m9)。
Private gShownText As String

' ============================================================================
' ShowToast - 画面の右上へカードを1枚出し、文の長さに応じた秒数(3～9秒。
'   modUIGeom.ToastSecondsFor)で自分で消える。
'   kind: "info"(白地・緑枠) / "warn"(黄地) / "error"(赤地・白字)。
'   空文字のときは何もしない(消し込みの呼び出しでカードを出さない)。
' ============================================================================
Public Sub ShowToast(ByVal messageText As String, Optional ByVal kind As String = "info")
    On Error Resume Next

    If LenB(Trim$(messageText)) = 0 Then Exit Sub
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub

    Dim ws As Object
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    ' 連続表示は前のカードと前の予約を片づけてから(重ねない・予約は1本)。
    CancelToast
    modUISheet.DropShapesByPrefix ws, UT_PREFIX

    Dim card As Object
    Set card = ws.Shapes.AddShape(UT_SHAPE_ROUNDED, CardLeft(), CardTop(), _
                                  UT_WIDTH, UT_MIN_H)
    If card Is Nothing Then Exit Sub

    card.Name = UT_CARD
    card.Placement = UT_PLACEMENT_FREE
    card.Adjustments(1) = 0.18
    card.Fill.ForeColor.RGB = FillOf(kind)
    card.Line.Visible = True
    card.Line.Weight = 1.25
    card.Line.ForeColor.RGB = LineOf(kind)

    With card.TextFrame2
        .WordWrap = UT_WRAP_ON
        .MarginLeft = 14#
        .MarginRight = 14#
        .MarginTop = 6#
        .MarginBottom = 6#
        .VerticalAnchor = UT_ANCHOR_MIDDLE
        .TextRange.Text = messageText
        .TextRange.Font.Name = UT_FONT
        .TextRange.Font.Size = UT_FONT_SIZE
        .TextRange.ParagraphFormat.Alignment = UT_ALIGN_LEFT
    End With
    card.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = TextOf(kind)

    FitCardHeight card
    card.ZOrder UT_FRONT
    gShownText = messageText
    ScheduleHide
End Sub

' ============================================================================
' ShowNext - 主要4ボタンの成功経路に出す「次の一手」(文言はここが唯一の値源)。
' ============================================================================
Public Sub ShowNext(ByVal stepNo As Long)
    On Error Resume Next
    ShowToast NextTextOf(stepNo), "info"
End Sub

' ============================================================================
' HideToast - OnTime のコールバック本体(Public必須)。**消すものが無ければ
'   何もしない**(ブックを閉じたあとに残った予約が発火しても無害にする)。
' ============================================================================
Public Sub HideToast()
    On Error Resume Next

    gHideAt = 0

    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        modUISheet.DropShapesByPrefix ws, UT_PREFIX
    Next ws
End Sub

' ============================================================================
' CancelToast - 未消化の OnTime 予約を取り消す。ShowToast が張り替えの前に
'   呼ぶほか、ブックの終了処理からも呼ぶ(裁定書17 H2・W9.2)。終了時の呼び口は
'   ThisWorkbook.Workbook_BeforeClose / Workbook_Deactivate で、RestoreScreen より前に呼ぶ。
' ============================================================================
Public Sub CancelToast()
    On Error Resume Next
    If CDbl(gHideAt) = 0# Then Exit Sub
    Application.OnTime EarliestTime:=gHideAt, Procedure:=HideProcName(), _
                       Schedule:=False
    Err.Clear
    gHideAt = 0
End Sub

' ============================================================================
' WarnLine - hm_warning へ書く1行。失敗系(kind="error")のときだけ、
'   err_log の送り方を末尾へ足す(裁定書17 H4)。空文字は空文字のまま返す。
' ============================================================================
Public Function WarnLine(ByVal messageText As String, ByVal kind As String) As String
    WarnLine = messageText
    If LenB(messageText) = 0 Then Exit Function
    If StrComp(kind, "error", vbBinaryCompare) <> 0 Then Exit Function
    WarnLine = messageText & UT_ERRLOG_NOTE
End Function

' ============================================================================
' 内部
' ============================================================================

' 次の一手の文言(1ボタン=1行)。11章§2/§3.1.1 の逐語に合わせる(v3.2:
'   ボタン名から番号を外し、区画の番号で呼ぶ)。
Private Function NextTextOf(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 1
        NextTextOf = "会社のことを2の枠へ貼って、[貼ったものを保存する]を押してください"
    Case 2
        NextTextOf = "できました。4の[レポートを出す]を押してください"
    Case 3
        NextTextOf = "レポートを出しました。次は4の[ヒアリングシートを出す]です"
    Case 4
        NextTextOf = "ヒアリングシートができました。Ctrl+Pで印刷できます"
    End Select
End Function

' カードの左端(表示領域の右上に寄せる)。取れないときは既定値。
Private Function CardLeft() As Double
    On Error GoTo Fallback1
    Dim x As Double
    x = ActiveWindow.VisibleRange.Left + ActiveWindow.VisibleRange.Width _
        - UT_WIDTH - UT_EDGE
    If x < UT_EDGE Then x = UT_EDGE
    CardLeft = x
    Exit Function
Fallback1:
    CardLeft = UT_EDGE
End Function

Private Function CardTop() As Double
    On Error GoTo Fallback2
    CardTop = ActiveWindow.VisibleRange.Top + UT_EDGE
    Exit Function
Fallback2:
    CardTop = UT_EDGE
End Function

' 文字量に合わせて高さを実測する(AutoSizeで一度伸ばし、読み取ってから解除)。
'   AutoSizeが効かない環境では下限のままになるので、下限・上限で丸める。
Private Sub FitCardHeight(ByVal card As Object)
    On Error Resume Next
    card.TextFrame2.AutoSize = UT_AUTOSIZE_ON
    Dim h As Double
    h = card.Height
    card.TextFrame2.AutoSize = UT_AUTOSIZE_OFF
    If h < UT_MIN_H Then h = UT_MIN_H
    If h > UT_MAX_H Then h = UT_MAX_H
    card.Height = h
End Sub

' 自動消去の予約(1本だけ持つ。予約できないのは非致命=出しっぱなしにしない
'   ために次の ShowToast が消す)。
Private Sub ScheduleHide()
    On Error Resume Next
    Dim t As Date
    t = Now + modUIGeom.ToastSecondsFor(gShownText) / 86400#
    Application.OnTime EarliestTime:=t, Procedure:=HideProcName()
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    gHideAt = t
End Sub

' 予約と取り消しで必ず同じ文字列になるように1箇所で組む。
'   **ブック名で修飾しない**(W9.2・実機第3報)。呼び先は自ブック内の Public Sub
'   であり、ブック名で修飾すると Application.OnTime の予約・取り消しの両方で
'   **ホスト側が非ASCII(日本語)のブック名を文字列解決する**経路ができる。Mac の
'   実Excel では起動直後に「実行時エラー 5」の生ダイアログが出た。修飾を外すと
'   VBA が自分のプロジェクト内で解決するので、その経路が消える。
Private Function HideProcName() As String
    HideProcName = "modUIToast.HideToast"
End Function

Private Function FillOf(ByVal kind As String) As Long
    Select Case LCase$(Trim$(kind))
    Case "warn"
        FillOf = UT_WARN_BG
    Case "error"
        FillOf = UT_ERROR_BG
    Case Else
        FillOf = UT_WHITE
    End Select
End Function

Private Function LineOf(ByVal kind As String) As Long
    Select Case LCase$(Trim$(kind))
    Case "warn"
        LineOf = UT_WARN_LINE
    Case "error"
        LineOf = UT_ERROR_BG
    Case Else
        LineOf = UT_ACCENT
    End Select
End Function

Private Function TextOf(ByVal kind As String) As Long
    If StrComp(LCase$(Trim$(kind)), "error", vbBinaryCompare) = 0 Then
        TextOf = UT_WHITE
    Else
        TextOf = UT_TEXT
    End If
End Function
