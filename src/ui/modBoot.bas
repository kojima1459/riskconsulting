Attribute VB_Name = "modBoot"
Option Explicit

' ============================================================================
' modBoot - 起動シーケンス(12章§2.1の7手順を固定順で実装する唯一の入口)
' ----------------------------------------------------------------------------
' 役割:
'   `Workbook_Open` (自己インストーラ。build/build_rpn.py が焼くThisWorkbook)
'   が注入完了後に `Application.Run "modBoot.Boot"` で呼ぶ唯一の入口。
'   「起動時に必ずやること」を本モジュールへ集約し、順序と失敗時挙動を
'   1箇所で固定する(12章§2.1)。modBoot以外に起動処理を書かない。
'
' 7手順(12章§2.1の表が正。番号は同表と対応):
'   (1) ガードシート非表示                          失敗=起動継続(E-code無し)
'   (2) config読込(既定値登録->シート上書き)         失敗=E0608警告・既定値続行
'   (3) modCaseStore.RepairStates()                  失敗=修復0件・E0603続行
'   (4) enum入力規則の隠しレンジ複製(11章§5)         KB未接続でも内蔵定数で続行
'   (5) modKnowledge.LoadKnowledge()+スナップショット 失敗=E0401警告・続行
'   (6) 新サービス候補のローカル退避分の再送           LoadKnowledge内部で無音実施
'   (7) modGatewayRPN.RunLimitCheck()                 Trueでも起動継続・HOME案内
'
'   手順(3)(5)(7)が呼ぶ modCaseStore.RepairStates / modKnowledge.LoadKnowledge /
'   modGatewayRPN.RunLimitCheck はいずれも失敗時の記録(E0603/E0401)とフォール
'   バックを関数内部で完結させる契約(14章§6・各モジュール本体)。手順(6)の
'   「新サービス候補の再送」も modKnowledge.LoadKnowledge の内部(FlushPending)
'   が無音で行うため、本モジュールから独立して呼ぶ処理は無い(E-13)。
'
' PoC対応: PoCの起動先(modViewport/modInstallCheck/RunFirstRunPromptEarly)は
'   RPNに存在しないため移植せず、12章§2.1の7手順として新規に組み立てた
'   (build/build_rpn.py のThisWorkbook側コメント参照)。
'
' protection_policy と ParkFocus(16章E-51・17章T-16 DoD):
'   シートProtectは使わない。ビルドは Locked 属性だけを焼き(build/
'   sheets_main.json の protection_policy)、`EnableSelection` は .xlsm に
'   永続化されない実行時プロパティのため毎起動この関数が全シートへ適用する。
'   ParkFocus(編集不可の待避セルへフォーカスを戻す)の恒久実装は
'   ui層 modUIProgress が持つ(14章§6・17章T-30)。**T-30の実装に伴い、暫定で
'   持っていた private ParkFocusAtBoot は modUIProgress.ParkFocus への委譲へ
'   差し替えた**(フォーカス退避の実装を2箇所に残さない。17章T-30 DoD)。
'
' 画面の用意(11章§5・T-30/T-31):
'   図形ボタンとOnActionの配線・enum入力規則の隠しレンジ複製・HOMEの初期表示は
'   ui層が持つ。起動シーケンスからは手順(4)で modUICase.ApplyEnumValidation を、
'   7手順の後で modUIHome.EnsureScreens を1回ずつ呼ぶだけにする(起動処理を
'   modBoot 以外へ散らさない=12章§2.1)。
'
' case_data の data_key 入力規則(11章§5・W0省略分の解消):
'   data_key の静的リスト(全28値)はExcelの入力規則インライン上限255字を
'   超える(build/sheets_main.json input_rule_omissions が宣言済み)ため、
'   W0ビルドは入力規則の付与を省略した。本モジュールは起動のたびに
'   veryHidden シート `enum_hidden` へ19章§3レジストリの内蔵定数(全28値)を
'   複製し、その範囲を指す名前付きレンジ経由でcase_data!data_key列へ
'   データの入力規則を張り直す(セル参照はインライン255字制限の対象外)。
'   本対応が本番ブックへ反映されたら、build/sheets_main.json の
'   input_rule_omissions からdata_keyのエントリを外すこと(tools/sheet_check.py
'   [7]の宣言突合。build/ 配下の変更は本タスクの範囲外のため別途対応)。
' ============================================================================

' DisableProcessWindowsGhosting(16章E-50(b)): リボンApplication.Runの同期
' 待機中にWindowsが画面を「応答なし」と誤判定して白くゴースト化するのを
' 抑止する、姉妹PoC実証済みの唯一の例外的採用API(表示系・引数なし・
' user32限定)。プロセス単位で一度呼べば以後の全呼出しに効くため、
' 呼出のたびではなく起動時に1回だけ呼ぶ。ui層に置く32/64bit両対応宣言。
#If VBA7 Then
    Private Declare PtrSafe Function DisableProcessWindowsGhosting _
        Lib "user32" () As Long
#Else
    Private Declare Function DisableProcessWindowsGhosting _
        Lib "user32" () As Long
#End If

Private Const BOOT_GUARD_SHEET As String = "はじめにお読みください"
Private Const BOOT_DATA_SHEET As String = "case_data"
Private Const BOOT_ENUM_SHEET As String = "enum_hidden"
Private Const BOOT_NAME_DATA_KEY As String = "enum_data_key"
Private Const BOOT_DATA_SCAN_COLS As Long = 20

' Excel組み込み定数の数値(名前を書かずLibreOffice側の構文チェックで未定義名に
' ならないようにする。core側 modConfig/modLog の CFG_DIR_UP 等と同じ流儀)。
Private Const BOOT_SHEET_HIDDEN As Long = 0        ' xlSheetHidden
Private Const BOOT_SHEET_VERY_HIDDEN As Long = 2   ' xlSheetVeryHidden
Private Const BOOT_DV_TYPE_LIST As Long = 3        ' xlValidateList
Private Const BOOT_DV_ALERT_STOP As Long = 1       ' xlValidAlertStop
Private Const BOOT_ENABLE_SELECTION_UNLOCKED As Long = 1   ' xlUnlockedCells

' 19章§3レジストリ(data_key・全28値・13章§2.2と完全一致)の内蔵定数複製。
' modCaseStore の同値の私有定数(CS_DATA_KEYS)とは別に、本モジュール単体で
' 起動できるよう独立して保持する(値は19章§3/sheets_main.json enums.data_key
' と完全一致させること)。
Private Const BOOT_DATA_KEYS As String = _
    "input_hp;input_yuho;input_memo;input_contract;input_prev_renewal;" & _
    "input_dossier;input_field_notes;input_coverage_note;" & _
    "input_hearing_answers;s1_json;s2_json;s3_json;s4_json;s2c_json;" & _
    "s3c_json;s2r_json;s3r_json;s2_prev_json;s1_edited;s2_edited;" & _
    "s3_edited;s4_edited;s1_json_failed;s2_json_failed;s3_json_failed;" & _
    "s4_json_failed;sparring_u;sparring_a"

Private Const BOOT_DV_ERROR_TITLE As String = "入力できない値です"
Private Const BOOT_DV_ERROR_MSG As String = "一覧から選んでください。"
Private Const BOOT_MSG_LIMIT_REACHED As String = _
    "リボンの利用上限の可能性があります。実行時に案内します。"

' ============================================================================
' Boot - 起動シーケンス本体(12章§2.1の7手順をこの順で1回ずつ実行する)。
' ============================================================================
Public Sub Boot()
    ' (1) ガードシート非表示
    HideGuardSheet

    ' (2) config読込(既定値登録->シート上書き)。keep_window_aliveは
    '     ここより前には読めないため、ゴースト化抑止も(2)の直後に置く。
    RegisterConfigDefaults
    modConfig.LoadFromSheet
    ApplyGhostingGuard

    ' (3) 案件状態の整合修復(RepairStatesが失敗時のE0603記録まで自己完結)
    modCaseStore.RepairStates

    ' (4) enum入力規則の隠しレンジ複製(11章§5)。data_key(日本語ラベルを持たない
    '     内部キー)は本モジュールが、日本語ラベルを持つ24グループは変換表を持つ
    '     modUICase が復元する(19章§3の値を2箇所に書かないため)。
    RestoreDataKeyHiddenRange
    modUICase.ApplyEnumValidation

    ' (5)+(6) ナレッジ読込・スナップショット保存・新サービス候補の再送
    '         (再送はLoadKnowledge内部のFlushPendingが無音で行う。E-13)
    modKnowledge.LoadKnowledge

    ' (7) 利用上限チェック。Trueでも起動は止めずHOMEへ案内する
    If modGatewayRPN.RunLimitCheck() Then NoticeLimitReached

    ' protection_policy(Locked+EnableSelection)の適用とParkFocus
    ' (16章E-51・17章T-16 DoD。Lockedはビルドが焼くのでここではEnableSelection
    '  のみを毎起動適用する)
    ApplyProtectionPolicy

    ' 画面の用意(図形ボタン+OnAction の配線とHOMEの初期表示。11章§5・T-30)
    modUIHome.EnsureScreens

    ' フォーカス退避(16章E-51(c))。実装は ui層 modUIProgress が唯一持つ。
    modUIProgress.ParkFocus
End Sub

' ----------------------------------------------------------------------------
' (1) ガードシート非表示
' ----------------------------------------------------------------------------
Private Sub HideGuardSheet()
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(BOOT_GUARD_SHEET)
    If Not ws Is Nothing Then ws.Visible = BOOT_SHEET_HIDDEN
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' (2) config既定値の登録(13章§2.3が正)。製品固有の既定値はcoreへ焼かず
'     ここで登録してから modConfig.LoadFromSheet がシート値で上書きする。
' ----------------------------------------------------------------------------
Private Sub RegisterConfigDefaults()
    modConfig.RegisterDefault "app_version", "2.0.0"
    modConfig.RegisterDefault "llm_transport", "ribbon"
    modConfig.RegisterDefault "mock_llm", "FALSE"
    modConfig.RegisterDefault "recommended_model", "gpt-5.5"
    modConfig.RegisterDefault "direct_model", "gpt-4.1"
    modConfig.RegisterDefault "direct_api_base", "https://api.openai.com/v1"
    modConfig.RegisterDefault "direct_key_path", "%APPDATA%\RPN\api_key.txt"
    modConfig.RegisterDefault "llm_wait_sec", "1200"
    modConfig.RegisterDefault "direct_http_timeout_ms", "120000"
    modConfig.RegisterDefault "reasoning_effort", "medium"
    modConfig.RegisterDefault "reasoning_verbosity", "low"
    modConfig.RegisterDefault "temperature", "0.3"
    modConfig.RegisterDefault "reasoning_tuning", "TRUE"
    modConfig.RegisterDefault "llm_max_tokens", "0"
    modConfig.RegisterDefault "app_tool_prefix", "リスク提案ナビ:"
    modConfig.RegisterDefault "max_context_chars", "40000"
    modConfig.RegisterDefault "t2_max_context_chars", "100000"
    modConfig.RegisterDefault "sparring_max_turns", "12"
    ' quality_modeはティア連動が既定(下流がdossier_tierから算出)であり単一の
    ' 固定既定値を持たないため、未設定を示す空文字を登録する。
    modConfig.RegisterDefault "quality_mode", vbNullString
    modConfig.RegisterDefault "deep_transport", vbNullString
    modConfig.RegisterDefault "ppt_max_slides_t2", "10"
    modConfig.RegisterDefault "kb_path", "\\...\ナレッジブック.xlsx"
    modConfig.RegisterDefault "kb_risk_rows", "20"
    modConfig.RegisterDefault "kb_menu_rows", "60"
    modConfig.RegisterDefault "kb_case_rows", "5"
    modConfig.RegisterDefault "kb_scheme_rows", "10"
    modConfig.RegisterDefault "kb_mech_rows", "40"
    modConfig.RegisterDefault "json_repair_retry", "1"
    modConfig.RegisterDefault "ppt_out_dir", "%USERPROFILE%\Documents\RPN出力"
    modConfig.RegisterDefault "ppt_template_path", vbNullString
    modConfig.RegisterDefault "html_out_dir", "%USERPROFILE%\Documents\RPN出力"
    modConfig.RegisterDefault "html_theme", "standard"
    modConfig.RegisterDefault "mock_fault", vbNullString
    modConfig.RegisterDefault "keep_window_alive", "TRUE"
    modConfig.RegisterDefault "ribbon_addin_name", "リボンちゃん"
    modConfig.RegisterDefault "limit_check", "TRUE"
    modConfig.RegisterDefault "log_max_rows", "2000"
    modConfig.RegisterDefault "anonymize_default", "TRUE"
    modConfig.RegisterDefault "feature_inbox", "TRUE"
    modConfig.RegisterDefault "feature_judgelog", "TRUE"
End Sub

' ----------------------------------------------------------------------------
' 画面ゴースト化抑止(16章E-50(b))。config keep_window_alive でオプトアウト可。
' ----------------------------------------------------------------------------
Private Sub ApplyGhostingGuard()
    If modConfig.GetBool("keep_window_alive", True) Then
        On Error Resume Next
        DisableProcessWindowsGhosting
        On Error GoTo 0
    End If
End Sub

' ----------------------------------------------------------------------------
' (4) case_data!data_key の入力規則を隠しレンジ参照で有効化する。
'     veryHiddenシート enum_hidden へ19章§3内蔵定数(全28値)を複製し、
'     その範囲を指す名前付きレンジをdata_key列のFormula1に張る
'     (セル参照は入力規則インライン255字制限の対象外。11章§5)。
' ----------------------------------------------------------------------------
Private Sub RestoreDataKeyHiddenRange()
    On Error Resume Next
    Dim ws As Object
    Set ws = EnsureEnumHiddenSheet()
    If ws Is Nothing Then Exit Sub

    Dim vals() As String
    vals = Split(BOOT_DATA_KEYS, ";")

    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        modUtilText.SetCellSafe ws.Cells(i + 1, 1), vals(i), "modBoot/enum_data_key"
    Next i

    Dim n As Long
    n = UBound(vals) - LBound(vals) + 1
    Dim target As Object
    Set target = ws.Range(ws.Cells(1, 1), ws.Cells(n, 1))
    ThisWorkbook.Names.Add Name:=BOOT_NAME_DATA_KEY, RefersTo:=target

    ApplyDataKeyValidation
    On Error GoTo 0
End Sub

Private Function EnsureEnumHiddenSheet() As Object
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(BOOT_ENUM_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Not ws Is Nothing Then
            ws.Name = BOOT_ENUM_SHEET
            ws.Visible = BOOT_SHEET_VERY_HIDDEN
        End If
    End If
    Set EnsureEnumHiddenSheet = ws
    On Error GoTo 0
End Function

Private Sub ApplyDataKeyValidation()
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(BOOT_DATA_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim hdr As Variant
    hdr = ws.Range(ws.Cells(1, 1), ws.Cells(1, BOOT_DATA_SCAN_COLS)).Value

    Dim col As Long
    col = modUtil.FindHeaderCol(hdr, "data_key")
    If col <= 0 Then Exit Sub

    Dim target As Object
    Set target = ws.Range(ws.Cells(2, col), ws.Cells(ws.Rows.Count, col))
    target.Validation.Delete
    target.Validation.Add Type:=BOOT_DV_TYPE_LIST, AlertStyle:=BOOT_DV_ALERT_STOP, _
        Formula1:="=" & BOOT_NAME_DATA_KEY
    target.Validation.IgnoreBlank = True
    target.Validation.ErrorTitle = BOOT_DV_ERROR_TITLE
    target.Validation.ErrorMessage = BOOT_DV_ERROR_MSG
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' (7) 利用上限がTrueのときのHOME案内(14章§2・12章§2.1手順(7))。
'     hm_warning が無い(未着手のui構築段階)場合は無音でスキップする。
' ----------------------------------------------------------------------------
Private Sub NoticeLimitReached()
    On Error Resume Next
    Dim target As Object
    Set target = ThisWorkbook.Names("hm_warning").RefersToRange
    If Not target Is Nothing Then
        modUtilText.SetCellSafe target, BOOT_MSG_LIMIT_REACHED, "modBoot/hm_warning"
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' protection_policy(16章E-51・build/sheets_main.json protection_policy)。
' Lockedはビルドが焼く(.xlsmに永続化される)。EnableSelectionは実行時
' プロパティで永続化されないため、毎起動このSubが全シートへ適用する。
' シートProtectは使わない(全面Protectは図形・表の再描画のたびに
' Unprotect/Protectが要り、描画失敗で保護状態が壊れるため)。
' ----------------------------------------------------------------------------
Private Sub ApplyProtectionPolicy()
    Dim ws As Object
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.EnableSelection = BOOT_ENABLE_SELECTION_UNLOCKED
        On Error GoTo 0
    Next ws
End Sub
