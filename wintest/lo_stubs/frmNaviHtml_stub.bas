Attribute VB_Name = "frmNaviHtml"
Option Explicit

' ============================================================================
' frmNaviHtml_stub - LibreOffice 実行テスト**専用**の空の frmNaviHtml
' ----------------------------------------------------------------------------
' 裁定書34 §1.4(W12-A)。
'
' なぜ要るのか:
'   LibreOffice Basic は UserForm(MSForms)も SHDocVw.WebBrowser も知らない。
'   本物の src/ui/navi/frmNaviHtml.frm は LO へ持ち込めないので、modNaviHost が
'   `Private gForm As frmNaviHtml` と書いている型名が LO 側で未定義になる。
'   本ファイルは**同じ名前の空の標準モジュール**で、その名前を埋めるだけの物。
'
' どこへ入るのか:
'   tools/run_lo_tests.py が LO へ注入するときだけ足す(LO_STUB_DIR)。
'   **src/ の下に置かない**ので build/modules.json にも載らず、配布 bin にも
'   入らない(第1段ビルドはそもそも type=form を読み飛ばす)。
'
' 何を保証するのか / しないのか:
'   保証する : 実フォームの Public な口(下の一覧)と名前が1対1で揃っていること。
'              照合は tools/ui_check.py (5) が機械で行う。ここが食い違うと
'              「LOでは通るのに実Excelで落ちる(またはその逆)」が起きる。
'   しない   : 画面が出ること・WebBrowser が動くこと・HTML が描かれること。
'              HTML画面の**描画**は Windows の実Excel でしか確かめられない
'              (17章§7 Z-43)。層(b)も層(a)もフォームを起こさない。
'
' 中身を空にしてある理由: 何か返す stub にすると、LO のテストが「本物が
'   動いた」と勘違いできてしまう。呼ばれたら何もしない(=テストが実際に
'   フォームを叩いたら結果が空になって落ちる)ほうが安全側である。
' ============================================================================

' 実フォームでは Public 変数(UserForm_QueryClose の抑止フラグ)。
Public AllowClose As Boolean

Public Sub CycleSize()
End Sub

Public Sub EnsureBrowser()
End Sub

Public Function TakePendingText() As String
End Function

Public Sub DeliverResponse(ByVal json As String)
End Sub

Public Sub ShowBusy(ByVal message As String, ByVal progressJson As String)
End Sub

Public Sub ApplyTextScale(ByVal fontScale As String)
End Sub

Public Sub RequestClose()
End Sub

Public Property Get IsReady() As Boolean
End Property

Public Property Get DocumentMode() As Long
End Property

Public Property Get HtmlLength() As Long
End Property

Public Property Get TextScale() As String
End Property

Public Property Get OpticalZoom() As Long
End Property
