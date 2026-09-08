Attribute VB_Name = "modBootNavi"
Option Explicit

' ============================================================================
' modBootNavi - modBoot の分割先(ui層・T-63・裁定書34 §1.2)
' ----------------------------------------------------------------------------
' なぜ新設したか(30,000字契約・12章§2):
'   W12-A で modBoot に (a) HTML画面まわりの config 既定値7件 (b) HTML画面を
'   起動するかどうかの判断 が加わるが、modBoot は 29,643字で残り 357字だった。
'   **コメントを削って字数を稼ぐことはしない**規約(CLAUDE.md)に従い、
'   「起動の7手順そのもの」を持つ modBoot から、次の3つを本モジュールへ移した。
'
'   (a) RegisterNaviDefaults  HTML画面・案件チャット・表示名の config 既定値
'   (b) LaunchIfHtml          ui_mode=html のときモードレス画面を予約する判断
'   (c) RestoreDataKeyHiddenRange  case_data!data_key の入力規則(11章§5)。
'                             data_key の内蔵定数(19章§3)ごと移した。
'
'   modBoot からは (a)(b)(c) を1行ずつ呼ぶ。起動手順の順序と番号(12章§2.1)は
'   変えていない。
'
' R1/R4(12章§4): ui層。Excel・シート・ThisWorkbook を触ってよい。
' ============================================================================

' 19章§3レジストリ(data_key・全32値・13章§2.2と完全一致)の内蔵定数複製。
' modCaseStore3.DataKeys() と同値だが、起動時に app層へ依存せず復元できるよう
' 独立して保持する(値は19章§3/sheets_main.json enums.data_key と完全一致
' させること)。裁定書34 §1.2 で chat_u / chat_a / nav_basics の3値を足した。
Private Const BN_DATA_KEYS As String = _
    "input_hp;input_yuho;input_memo;input_contract;input_prev_renewal;" & _
    "input_dossier;input_field_notes;input_coverage_note;input_finance;" & _
    "input_hearing_answers;s1_json;s2_json;s3_json;s4_json;s2c_json;" & _
    "s3c_json;s2r_json;s3r_json;s2_prev_json;s1_edited;s2_edited;" & _
    "s3_edited;s4_edited;s1_json_failed;s2_json_failed;s3_json_failed;" & _
    "s4_json_failed;sparring_u;sparring_a;chat_u;chat_a;nav_basics"

Private Const BN_ENUM_SHEET As String = "enum_hidden"
Private Const BN_NAME_DATA_KEY As String = "enum_data_key"
Private Const BN_DATA_SHEET As String = "case_data"
Private Const BN_DATA_SCAN_COLS As Long = 20

' Excel組み込み定数の数値(名前を書かずLibreOffice側の構文チェックで未定義名に
' ならないようにする。modBoot の BOOT_* と同じ流儀)。
Private Const BN_SHEET_VERY_HIDDEN As Long = 2   ' xlSheetVeryHidden
Private Const BN_DV_TYPE_LIST As Long = 3        ' xlValidateList
Private Const BN_DV_ALERT_STOP As Long = 1       ' xlValidAlertStop

Private Const BN_DV_ERROR_TITLE As String = "入力できない値です"
Private Const BN_DV_ERROR_MSG As String = "一覧から選んでください。"

' HTML画面の資産(裁定書34 §0.1)。本体xlsm と同じフォルダの ui フォルダに置く。
Private Const BN_UI_DIR As String = "ui"
Private Const BN_UI_INDEX As String = "index.html"
' ui フォルダが無いときの案内(16章 E-64)。HTML画面は諦めてシート画面で続ける。
Private Const BN_MSG_NO_UI As String = _
    "ui フォルダが見つからないため従来画面で起動しました。"

' ============================================================================
' RegisterNaviDefaults - HTML画面まわりの config 既定値(13章§2.3・裁定書34)。
' ----------------------------------------------------------------------------
'   modBoot.RegisterConfigDefaults と同じ規律で「製品固有の既定値は core へ
'   焼かずここで登録し、その後 modConfig.LoadFromSheet がシート値で上書きする」。
'   ui_mode の既定は **html**(配布の既定。sheet は ActiveX を止められた端末や
'   Mac のための予備)。
'   app_display_name は画面に出す製品名の**唯一の値源**で、PJ が名前を決めたら
'   ここ1行(と config シート1行)を書き換える(裁定書34 §0.4)。
' ============================================================================
Public Sub RegisterNaviDefaults()
    modConfig.RegisterDefault "ui_mode", "html"
    modConfig.RegisterDefault "ui_font_scale", "medium"
    modConfig.RegisterDefault "chat_max_turns", "12"
    modConfig.RegisterDefault "chat_include_materials", "FALSE"
    modConfig.RegisterDefault "ch_effort", "medium"
    modConfig.RegisterDefault "ch_verbosity", "low"
    modConfig.RegisterDefault "app_display_name", "リスク提案ナビ"
End Sub

' ============================================================================
' AppDisplayName - 画面へ出す製品名(裁定書34 §0.4)。値源は config 1箇所。
' ============================================================================
Public Function AppDisplayName() As String
    AppDisplayName = modConfig.GetStr("app_display_name", "リスク提案ナビ")
End Function

' ============================================================================
' LaunchIfHtml - HTML画面(モードレスの1枚窓)を開く予約(裁定書34 §1.2(b))。
' ----------------------------------------------------------------------------
'   条件は2つ。どちらか欠ければ**シート画面のまま起動を続ける**(16章 E-64):
'     (1) config ui_mode が "html"(大小文字は問わない)
'     (2) 本体xlsm と同じフォルダに ui\index.html が実在する
'   (2)が欠けたときは hm_warning に理由を書いて利用者へ知らせる。ここで
'   MsgBox は出さない(起動を止めない・無人実行を壊さない)。
'
'   開き方は Application.OnTime Now(=起動シーケンスを抜けてから開く)。
'   Workbook_Open の中で直接 Show すると、Excel がまだ描画を終えていない状態で
'   モードレスのフォームが載って表示が崩れる。文字列ディスパッチの名前は
'   **ブック名修飾なし**の "OpenNaviTool"(W9.2 の規約。ブック名に空白や
'   全角が入っても壊れない)。
' ============================================================================
Public Sub LaunchIfHtml()
    On Error GoTo Failed

    If LCase$(Trim$(modConfig.GetStr("ui_mode", "sheet"))) <> "html" Then Exit Sub

    If Not UiIndexExists() Then
        modUISheet.WriteNamed "hm_warning", BN_MSG_NO_UI
        Exit Sub
    End If

    Application.OnTime Now, "OpenNaviTool"
    Exit Sub

Failed:
    modLog.LogError "E0603", "modBootNavi.LaunchIfHtml", "html_launch_failed", Err.Number
    Resume Ignore0
Ignore0:
End Sub

' 本体xlsm と同じフォルダの ui\index.html が読めるか(16章 E-64 の判定)。
Private Function UiIndexExists() As Boolean
    Dim folder As String
    folder = modUtilPath.JoinPath(ThisWorkbook.Path, BN_UI_DIR)
    UiIndexExists = modUtil.FileExistsAt(modUtilPath.JoinPath(folder, BN_UI_INDEX))
End Function

' ============================================================================
' RestoreDataKeyHiddenRange - case_data!data_key の入力規則を隠しレンジ参照で
'   有効化する(12章§2.1 の手順(4)。modBoot から移設・挙動は不変)。
'   veryHiddenシート enum_hidden へ19章§3内蔵定数(全32値)を複製し、その範囲を
'   指す名前付きレンジを data_key 列の Formula1 に張る(セル参照は入力規則
'   インライン255字制限の対象外。11章§5)。
' ============================================================================
Public Sub RestoreDataKeyHiddenRange()
    On Error Resume Next
    Dim ws As Object
    Set ws = EnsureEnumHiddenSheet()
    If ws Is Nothing Then Exit Sub

    Dim vals() As String
    vals = Split(BN_DATA_KEYS, ";")

    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        modUtilText.SetCellSafe ws.Cells(i + 1, 1), vals(i), "modBootNavi/enum_data_key"
    Next i

    Dim n As Long
    n = UBound(vals) - LBound(vals) + 1
    Dim target As Object
    Set target = ws.Range(ws.Cells(1, 1), ws.Cells(n, 1))
    ThisWorkbook.Names.Add Name:=BN_NAME_DATA_KEY, RefersTo:=target

    ApplyDataKeyValidation
    On Error GoTo 0
End Sub

Private Function EnsureEnumHiddenSheet() As Object
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(BN_ENUM_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Not ws Is Nothing Then
            ws.Name = BN_ENUM_SHEET
            ws.Visible = BN_SHEET_VERY_HIDDEN
        End If
    End If
    Set EnsureEnumHiddenSheet = ws
    On Error GoTo 0
End Function

Private Sub ApplyDataKeyValidation()
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(BN_DATA_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim hdr As Variant
    hdr = ws.Range(ws.Cells(1, 1), ws.Cells(1, BN_DATA_SCAN_COLS)).Value

    Dim col As Long
    col = modUtil.FindHeaderCol(hdr, "data_key")
    If col <= 0 Then Exit Sub

    Dim target As Object
    Set target = ws.Range(ws.Cells(2, col), ws.Cells(ws.Rows.Count, col))
    target.Validation.Delete
    target.Validation.Add Type:=BN_DV_TYPE_LIST, AlertStyle:=BN_DV_ALERT_STOP, _
        Formula1:="=" & BN_NAME_DATA_KEY
    target.Validation.IgnoreBlank = True
    target.Validation.ErrorTitle = BN_DV_ERROR_TITLE
    target.Validation.ErrorMessage = BN_DV_ERROR_MSG
    On Error GoTo 0
End Sub
