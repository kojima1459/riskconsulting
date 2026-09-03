VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' clsAppEvents - アプリのブックイベントを受ける唯一のクラス(ui層・裁定書26 B)
' ----------------------------------------------------------------------------
' なぜクラスなのか(裁定書26追補):
'   全画面表示は「本ブックが前面のときだけ当て、離れたら元へ戻す」必要がある
'   (16章の既存原則: 他のブックの画面を壊さない)。ThisWorkbook モジュールは
'   ビルドの外科パッチ対象(ASCII・バイト長固定)で触れず、焼き付け後は中身が
'   変わらないため、**WithEvents を持つクラスを起動時に1つ生成**して
'   Application のブックイベントを受ける。ThisWorkbook に一切依存しないので
'   焼き付け済みファイルでも効く。
'
' 生成と保持は modBoot が1箇所で行う(gAppEvents。参照を捨てるとイベントが
'   来なくなるのでモジュール変数で持つ)。本クラスは**判断を持たず**、
'   modUIViewport の2本を呼ぶだけにする。
'
' App は Public WithEvents である(modBoot が `Set gAppEvents.App = Application`
'   で結線するため。Private では外から代入できない)。
' 画面設定の失敗で業務を止めないよう、3本とも On Error Resume Next で包む。
' ============================================================================

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

' 本ブックが前面になった: 全画面へ戻す(config ui_fullscreen が FALSE のときは
'   ApplyFullScreen 側が何もしない)。
Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    If Wb Is ThisWorkbook Then modUIViewport.ApplyFullScreen
End Sub

' 本ブックから離れた: 退避した元の値へ戻す(他のブックを全画面のままにしない)。
Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    On Error Resume Next
    If Wb Is ThisWorkbook Then modUIViewport.RestoreScreen
End Sub

' 本ブックを閉じる: 同上。閉じたあとのExcelを全画面のままにしない。
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    If Wb Is ThisWorkbook Then modUIViewport.RestoreScreen
End Sub
