Attribute VB_Name = "modNaviWindow"
Option Explicit
' ============================================================================
' modNaviWindow - UserForm のネイティブ窓化(裁定書44 A-6)
' ----------------------------------------------------------------------------
' なぜ要るか: HTML画面(`frmNaviHtml`)はモードレスの UserForm であり、既定では
'   最小化・最大化ボタンや境界ドラッグでのリサイズができない(`ThunderDFrame`
'   ウィンドウの既定スタイルに無い)。利用者(小島さん・PJの意思決定者)が
'   2026-09-14 に「案1で」と明示決定した経路――user32 の表示系APIで
'   ウィンドウスタイルへ厚みを足す方式――をここへ実装する。
'
' なぜ Declare が使えるのか(裁定書27 W9-B7 との関係): `Declare` の撤去は
'   **予防措置**であり、2026-09-02 に社内AV(AMSI)が実際に検知したのは
'   「隠しシートの文字列をVBAプロジェクトへ注入するループ」であって
'   (`docs/28_開発担当専用_検証PCの設定.md` §0)、`Declare` 自体ではない
'   (`src/ui/modBoot.bas` の `ApplyGhostingGuard` の訂正コメント参照)。
'   姉妹PJ MyBookshelf は `Declare PtrSafe Sub DisableProcessWindowsGhosting
'   Lib "user32" ()` を2026-08-20から会社PCで運用し、2026-09-12にDefender
'   警告なしを実機確認済み(このリポジトリ外の事実。司令塔確認済み)。
'   これらを根拠に、**表示系4本だけ**をこのモジュールに限って再許可した
'   (`tools/vba_lint.py` の `FORBIDDEN_API_DECLARE_ALLOW`)。このモジュール
'   以外・この4本以外の `Declare` は従来どおり禁止のままである。
'
' 失敗時の扱い: `EnableNativeWindow` の全体を `On Error GoTo Failed` で包む。
'   `FindWindowA` が hWnd を取れない(タイミング・環境差)・API呼び出し自体が
'   失敗する、のどちらでも**画面表示は続ける**(呼ばなかったのと同じ状態へ
'   フォールバック)。失敗の記録だけ `modLog.LogUsage` へ残す。
'
' 有効/無効の切替: config `ui_window_native`(既定 TRUE)。呼び口は
'   `modNaviHost.OpenNaviTool` の `gForm.Show 0` の**直後**(Show の前は
'   `ThunderDFrame` ウィンドウが存在せず `FindWindowA` が取れないため)。
' ============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongA Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLongA Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" _
        (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLongA Lib "user32" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" _
        (ByVal hWnd As Long) As Long
#End If

' UserForm(ThunderDFrame)のウィンドウスタイル定数(user32)。
Private Const NW_GWL_STYLE As Long = -16
Private Const NW_WS_THICKFRAME As Long = &H40000&
Private Const NW_WS_MINIMIZEBOX As Long = &H20000&
Private Const NW_WS_MAXIMIZEBOX As Long = &H10000&
Private Const NW_WS_SYSMENU As Long = &H80000&
Private Const NW_CLASS As String = "ThunderDFrame"

' EnableNativeWindow - frm(表示済みのUserForm)へ最小化・最大化・境界ドラッグの
'   スタイルを足す。config ui_window_native が FALSE のときは何もしない。
'   hWnd が取れない・API呼び出しが失敗するなど、どの段でも**例外を外へ
'   出さず**画面表示は続ける(呼び出し側はこの呼び出しの成否を見ない)。
'   hWndは`Variant`で持つ(手続き内で `#If VBA7` を分岐させると LibreOffice
'   のBasicコンパイラがハングする実測があったため。モジュール冒頭の
'   宣言部の `#If VBA7`(Declare の戻り値型の分岐)は問題なくコンパイルが
'   通る――ハングするのは**手続き本体の中**の条件付きコンパイルだけである。
'   Variant はDeclare呼び出しの実引数として渡すとき、宣言側の型
'   (LongPtr/Long)へ暗黙変換されるため実害は無い)。
Public Sub EnableNativeWindow(ByVal frm As Object)
    On Error GoTo Failed
    If Not modConfig.GetBool("ui_window_native", True) Then Exit Sub

    Dim hWnd As Variant
    hWnd = FindWindowA(NW_CLASS, frm.Caption)
    If hWnd = 0 Then Exit Sub

    Dim styleValue As Long
    styleValue = GetWindowLongA(hWnd, NW_GWL_STYLE)
    styleValue = styleValue Or NW_WS_THICKFRAME Or NW_WS_MINIMIZEBOX Or _
                 NW_WS_MAXIMIZEBOX Or NW_WS_SYSMENU
    SetWindowLongA hWnd, NW_GWL_STYLE, styleValue
    DrawMenuBar hWnd
    Exit Sub
Failed:
    modLog.LogUsage "ui_native_window", vbNullString, "failed:" & Err.Number
End Sub
