Attribute VB_Name = "modUIViewport"
Option Explicit

' ============================================================================
' modUIViewport - 全画面表示(ui層・裁定書26 B)
' ----------------------------------------------------------------------------
' 目的(11章§0 体験原則): ふだんのExcelの部品(数式バー・罫線・行列見出し)を
'   消し、ナビの1画面だけが見える状態にする。**シートタブは表示のまま**に
'   する(ナビと使い方の2枚を行き来するため)。
'
' 適用の条件と作法(裁定書26 B・16章の既存原則):
'   ・本ブックが前面のときだけ触る(他のブックの画面を壊さない)。
'   ・config `ui_fullscreen`(既定TRUE)が FALSE なら何もしない。
'   ・変える前の値を1度だけ退避し、RestoreScreen で退避した値へ戻す。
'   ・全て On Error Resume Next 配下(画面設定の失敗で業務を止めない)。
'
' 誰が呼ぶか(裁定書26 B / W9.3 でクラスから ThisWorkbook へ移した):
'   ApplyFullScreen = modBoot の起動シーケンスの最後(ナビ描画後)に1回 ＋
'     ThisWorkbook.Workbook_Activate(本ブックが前面へ戻ったとき)
'   RestoreScreen   = ThisWorkbook.Workbook_Deactivate /
'     ThisWorkbook.Workbook_BeforeClose(本ブックから離れたとき・閉じるとき)
'   **イベントの受け口は ThisWorkbook 文書モジュールのみ**で、クラスモジュール
'   (旧 clsAppEvents)は使わない。Mac の実Excel でクラスを1本含めるだけで
'   読み込み時に Err 5 の生ダイアログが出たため撤去した(17章 Z-24)。
'   ThisWorkbook の中身はビルドが焼く(値源は build/build_rpn.py の
'   _BAKED_THISWORKBOOK_TEXT)ので、焼き付け済み(baked)のファイルでも効く。
' ============================================================================

Private Const VP_SRC As String = "modUIViewport"
Private Const VP_CFG_KEY As String = "ui_fullscreen"

' 退避した元の値(1度だけ取る)。gSaved=False のときは戻さない。
Private gSaved As Boolean
Private gFullScreen As Boolean
Private gFormulaBar As Boolean
Private gGridlines As Boolean
Private gHeadings As Boolean

' ============================================================================
' ApplyFullScreen - 全画面へ切り替える(冪等)。
'   順序は「アプリ側(全画面・数式バー) -> 窓側(罫線・見出し)」で固定する。
'   全画面の切り替えで窓が作り直されるため、先に窓を触っても消える。
' ============================================================================
Public Sub ApplyFullScreen()
    On Error Resume Next

    If Not modConfig.GetBool(VP_CFG_KEY, True) Then Exit Sub
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub

    SaveOriginals

    If Not Application.DisplayFullScreen Then Application.DisplayFullScreen = True
    If Application.DisplayFormulaBar Then Application.DisplayFormulaBar = False

    Dim win As Object
    Set win = ActiveWindow
    If win Is Nothing Then Exit Sub
    If win.DisplayGridlines Then win.DisplayGridlines = False
    If win.DisplayHeadings Then win.DisplayHeadings = False
    ' シートタブは**触らない**(裁定書26 B。ナビと使い方を行き来する導線)。

    ' W9.2 N3: On Error Resume Next が握りつぶした失敗を記録だけは残す。
    If Err.Number <> 0 Then
        modLog.LogError "E0603", VP_SRC & ".ApplyFullScreen", "fullscreen_failed", Err.Number
        Err.Clear
    End If
End Sub

' ============================================================================
' RestoreScreen - 退避した元の値へ戻す。退避が無ければ何もしない。
' ============================================================================
Public Sub RestoreScreen()
    On Error Resume Next

    If Not gSaved Then Exit Sub

    Application.DisplayFullScreen = gFullScreen
    Application.DisplayFormulaBar = gFormulaBar

    Dim win As Object
    Set win = ActiveWindow
    If Not win Is Nothing Then
        win.DisplayGridlines = gGridlines
        win.DisplayHeadings = gHeadings
    End If

    gSaved = False
End Sub

' 元の値の退避(1度だけ)。2度目以降は上書きしない(全画面にした後の値を
'   「元の値」として覚えてしまうと、戻す先が全画面になる)。
Private Sub SaveOriginals()
    On Error Resume Next

    If gSaved Then Exit Sub

    gFullScreen = Application.DisplayFullScreen
    gFormulaBar = Application.DisplayFormulaBar

    Dim win As Object
    Set win = ActiveWindow
    If win Is Nothing Then Exit Sub
    gGridlines = win.DisplayGridlines
    gHeadings = win.DisplayHeadings

    gSaved = True
End Sub
