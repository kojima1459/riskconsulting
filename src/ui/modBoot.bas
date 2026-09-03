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

Private Const BOOT_SRC As String = "modBoot"
Private Const BOOT_GUARD_SHEET As String = "はじめにお読みください"
Private Const BOOT_DATA_SHEET As String = "case_data"
' 保存先(裁定書27 W9-C2)。既定は会社のOneDrive。html_out_dir に値が入って
' いればそちらを優先する(分けたい管理者向け)。
Private Const BOOT_DATA_DIR_KEY As String = "data_dir"
Private Const BOOT_OUT_DIR_KEY As String = "html_out_dir"
Private Const BOOT_DATA_DIR_DEFAULT As String = "%OneDriveCommercial%\リスク提案ナビ\データ"
' 逐語(裁定書27 W9-C2)。1字も変えない。
Private Const BOOT_MSG_NOT_ONEDRIVE As String = _
    "保存先がOneDriveではありません。シャットダウンで消える可能性があります。"
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

' 19章§3レジストリ(data_key・全29値・13章§2.2と完全一致)の内蔵定数複製。
' modCaseStore3.DataKeys() と同値だが、本モジュール単体で
' 起動できるよう独立して保持する(値は19章§3/sheets_main.json enums.data_key
' と完全一致させること)。
Private Const BOOT_DATA_KEYS As String = _
    "input_hp;input_yuho;input_memo;input_contract;input_prev_renewal;" & _
    "input_dossier;input_field_notes;input_coverage_note;input_finance;" & _
    "input_hearing_answers;s1_json;s2_json;s3_json;s4_json;s2c_json;" & _
    "s3c_json;s2r_json;s3r_json;s2_prev_json;s1_edited;s2_edited;" & _
    "s3_edited;s4_edited;s1_json_failed;s2_json_failed;s3_json_failed;" & _
    "s4_json_failed;sparring_u;sparring_a"

' 起動手順の本数(BootStep の Select Case と一致させること)。
Private Const BOOT_STEP_COUNT As Long = 11

' 業種入力規則(13章§2.11・裁定書9 §4)。値源はナレッジのスナップショット
' (16章 E-08 の退避先)であり、modKnowledge の内部表現を読む口を増やさない
' ためにシートを直接読む。索引1 = 業種マスタ(modKnowledge KB_SHEETS の先頭)。
Private Const BOOT_SNAP_SHEET As String = "kb_snapshot"
Private Const BOOT_SNAP_IDX_INDUSTRY As String = "1"
Private Const BOOT_IND_CODE_COL As Long = 30     ' enum_hidden の使用列(1=data_key
Private Const BOOT_IND_NAME_COL As Long = 31     '  / 2-25=19章§3の24グループ)
Private Const BOOT_IND_CODE_NAME As String = "enum_industry_code"
Private Const BOOT_IND_NAME_NAME As String = "enum_industry_name"

Private Const BOOT_DV_ERROR_TITLE As String = "入力できない値です"
Private Const BOOT_DV_ERROR_MSG As String = "一覧から選んでください。"

' ナレッジブックの自動発見(裁定書17 H1 / 裁定書19 H8(b) で改訂)。config
' kb_path の既定値は配置前のプレースホルダ(BOOT_KB_PLACEHOLDER を含む)であり、
' 実機では「本体と同じフォルダにナレッジブックを置いたのに読めない」が起きた。
' kb_path が空・プレースホルダ・不在のときは本体と同じフォルダの BOOT_KB_FILE を
' kb_path へ書いてから読み込みへ進む。**本体がOneDrive同期フォルダにあると
' ThisWorkbook.Path がURL形式("://" を含む)になり、"\" 連結+Dir$ では必ず
' 不発になる**ため、URLのときは "/" で連結し Dir$ を通さずに書く(開けるか
' どうかの判定は modKnowledge.LoadKnowledge に委ねる=fail-closedはそちら)。
Private Const BOOT_KB_KEY As String = "kb_path"
Private Const BOOT_KB_FILE As String = "ナレッジブック.xlsx"
Private Const BOOT_KB_PLACEHOLDER As String = "\\...\"
Private Const BOOT_KB_AUTO_NOTE As String = "同じフォルダのナレッジブックを読み込みました。"

' W9.2 N9: 起動手順のどれかが E0603 を記録したときに出す1枚(逐語)。
' 生ダイアログの代わりに「次に何をすればよいか」を必ず書く(11章§0.1 原則③)。
Private Const BOOT_MSG_PARTIAL As String = _
    "起動の一部が完了しませんでした。使い方タブの[記録を見る]を押してください。" & _
    vbLf & "err_log の最後の行を開発担当へお送りください。"

' 自動発見でパスを書いたか(HOMEのナレッジ欄が KbAutoNote で読む)。永続しない
' 画面制御変数であり、14章§6「状態保持の例外」には当たらない。
Private gKbAutoFound As Boolean

' アプリのブックイベントを受けるクラス(裁定書26 B・追補)。**参照を捨てると
'   イベントが来なくなる**ので、起動から終了までモジュール変数で保持する。
'   ThisWorkbook に依存しないので焼き付け済みファイルでも効く。
Private gAppEvents As clsAppEvents

' 起動手順のどれかが失敗した(=E0603 を記録した)か。W9.2 N9 のトーストの条件。
'   永続しない画面制御変数であり、14章§6「状態保持の例外」には当たらない。
Private gBootPartial As Boolean

' ============================================================================
' Boot - 起動シーケンス本体(12章§2.1の7手順をこの順で1回ずつ実行する)。
' ============================================================================
Public Sub Boot()
    ' W9.2 N1: **起動経路の最外周の網**。Boot 自身が未捕捉の実行時エラーを外へ
    ' 出すと、Excel が生のダイアログ(「実行時エラー 5」など・[OK]のみ)を出す。
    ' 利用者は何も分からず、記録も残らない。ここで必ず受け止める。
    On Error Resume Next

    ' 裁定書9 B21(12章§2.1): 各手順を**個別のエラーハンドラ**で包む。途中の
    ' 手順が未捕捉の実行時エラーを投げても、以降の手順(EnableSelectionの毎起動
    ' 適用=16章E-51(b)・ボタン配線・ParkFocus)まで必ず到達させる。
    Dim i As Long
    gBootPartial = False
    For i = 1 To BOOT_STEP_COUNT
        BootStep i
    Next i

    ' W9.2 N9: 起動の途中で1つでも E0603 を記録していたら、黙って進まずに
    ' 「一部が完了しなかった」ことを利用者へ1枚出す(生ダイアログは出さない)。
    If gBootPartial Then modUIToast.ShowToast BOOT_MSG_PARTIAL, "warn"
    Err.Clear
End Sub

' 起動手順1本ぶん。失敗しても記録して次へ進む(起動を止めない。12章§2.1)。
Private Sub BootStep(ByVal stepNo As Long)
    On Error GoTo Failed

    Select Case stepNo
    Case 1
        ' (1) ガードシート非表示
        HideGuardSheet
    Case 2
        ' (2) config読込(既定値登録->シート上書き)。keep_window_aliveは
        '     ここより前には読めないため、ゴースト化抑止も(2)の直後に置く。
        RegisterConfigDefaults
        modConfig.LoadFromSheet
        ApplyGhostingGuard
    Case 3
        ' (3) 案件状態の整合修復(RepairStatesが失敗時のE0603記録まで自己完結)
        modCaseStore.RepairStates
    Case 4
        ' (4) enum入力規則の隠しレンジ複製(11章§5)。data_key(日本語ラベルを
        '     持たない内部キー)は本モジュールが復元する。
        RestoreDataKeyHiddenRange
    Case 5
        '     日本語ラベルを持つ24グループは変換表を持つ modUICase が復元する
        '     (19章§3の値を2箇所に書かないため)。
        modUICase.ApplyEnumValidation
    Case 6
        ' (5)+(6) ナレッジ読込・スナップショット保存・新サービス候補の再送
        '         (再送はLoadKnowledge内部のFlushPendingが無音で行う。E-13)
        '         読込の**前**に kb_path の自動発見を1回だけ挟む(裁定書17 H1)。
        ResolveKbPath
        If Not modKnowledge.LoadKnowledge() Then LogKbPathTried
    Case 7
        ' 業種入力規則(13章§2.11・裁定書9 §4)。ナレッジ読込の**後**に置く
        ' (直前の読込で更新されたスナップショットを使う)。
        RestoreIndustryHiddenRange
    Case 8
        ' (7) 利用上限チェック。Trueでも起動は止めずHOMEへ案内する
        If modGatewayRPN.RunLimitCheck() Then NoticeLimitReached
    Case 9
        ' protection_policy(Locked+EnableSelection)の適用
        ' (16章E-51・17章T-16 DoD。Lockedはビルドが焼くのでここでは
        '  EnableSelectionのみを毎起動適用する)
        ApplyProtectionPolicy
    Case 10
        ' 画面の用意(図形ボタン+OnAction の配線とナビの初期表示。11章§5・T-30)。
        ' EnsureScreens の末尾が RefreshHome -> modUINav.DrawNav まで通す。
        modUIHome.EnsureScreens
        ' 裁定書27 W9-B3: [中身を見る]は %TEMP% へ一時ファイルを書かなくなった
        ' (ブック内の「中身」シートへ流し込む)ため、後始末の掃除も撤去した。
        ' 保存先がOneDriveでないときのお知らせ(裁定書27 W9-C2)。画面を描いた
        ' 後に出す(hm_warning とトーストの両方へ出るため描画済みが要る)。
        NoticeDataDir
    Case 11
        ' フォーカス退避(16章E-51(c))。実装は ui層 modUIProgress が唯一持つ。
        modUIProgress.ParkFocus
        ' 初回ガイドツアー(裁定書14 裁定6)。実装は modUIGuide が唯一持ち、
        ' 起動シーケンスからの結線はこの1行だけにする(2回目以降は何もしない)。
        modUIGuide.StartTourIfFirstRun
        ' ブックイベントの結線(裁定書26 B・追補)。**全画面を当てる前**に結線を
        ' 済ませる(Activate/Deactivate/BeforeClose で戻す口を先に用意する)。
        HookAppEvents
        ' 全画面表示(裁定書26 B)。**起動シーケンスの最後・ナビを描いた後**に
        ' 1回だけ当てる(先に当てると窓の作り直しで幾何が古い窓のまま決まる)。
        modUIViewport.ApplyFullScreen
    End Select
    Exit Sub

Failed:
    ' W9.2 N2: **ハンドラの中で起きた失敗も外へ出さない**。VBAはエラーハンドラ内で
    ' 発生したエラーを同じハンドラでは受けられない(On Error Resume Next を書いても
    ' 効かず、呼び出し元=Boot へ飛ぶ。tools/vba_lint.py の同名ルール)。したがって
    ' 記録は**別のプロシージャ**へ切り出し、そちらで網を張る。
    gBootPartial = True
    LogBootStepFailure stepNo, Err.Number
End Sub

' BootStep の Failed: から呼ぶ記録専用(W9.2 N2)。**ハンドラの外**なので
'   On Error Resume Next が効き、記録そのものが失敗しても呼び出し元へ飛ばない。
Private Sub LogBootStepFailure(ByVal stepNo As Long, ByVal errNo As Long)
    On Error Resume Next
    modLog.LogError "E0603", BOOT_SRC & ".Boot", _
                    "boot_step_failed:" & CStr(stepNo), errNo
End Sub

' ----------------------------------------------------------------------------
' ブックイベントの結線(裁定書26 B・追補)
' ----------------------------------------------------------------------------
' WithEvents を持つクラスを1つだけ作り、Application のブックイベントを受ける。
'   2回呼ばれても作り直すだけで害は無い(古い方は参照が切れて自動的に消える)。
Private Sub HookAppEvents()
    On Error Resume Next
    Set gAppEvents = New clsAppEvents
    If gAppEvents Is Nothing Then Exit Sub
    Set gAppEvents.App = Application
End Sub

' AppEventsReady - 結線できているか(層(b)の回帰が読む唯一の口)。
Public Function AppEventsReady() As Boolean
    On Error Resume Next
    If gAppEvents Is Nothing Then Exit Function
    AppEventsReady = Not (gAppEvents.App Is Nothing)
End Function

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
    modConfig.RegisterDefault "kb_incident_rows", "5"
    modConfig.RegisterDefault "kb_mech_rows", "40"
    modConfig.RegisterDefault "json_repair_retry", "1"
    modConfig.RegisterDefault "ppt_out_dir", "%USERPROFILE%\Documents\RPN出力"
    modConfig.RegisterDefault "ppt_template_path", vbNullString
    ' 裁定書27 W9-C2: 成果物の保存先は data_dir が正。html_out_dir は空を既定に
    ' して「data_dir に従う」を既定動作にし、分けたい管理者だけが値を入れる。
    modConfig.RegisterDefault "data_dir", "%OneDriveCommercial%\リスク提案ナビ\データ"
    modConfig.RegisterDefault "html_out_dir", vbNullString
    modConfig.RegisterDefault "html_theme", "standard"
    modConfig.RegisterDefault "mock_fault", vbNullString
    modConfig.RegisterDefault "keep_window_alive", "TRUE"
    modConfig.RegisterDefault "ribbon_addin_name", "リボンちゃん"
    modConfig.RegisterDefault "limit_check", "TRUE"
    modConfig.RegisterDefault "log_max_rows", "2000"
    modConfig.RegisterDefault "anonymize_default", "TRUE"
    ' 裁定書11 Q7(裁定書10 m7・13章§2.3): 実施者。run_log の operator 列と
    ' 受信箱の judged_by_group がこのキーを読む。個人名は入れない(部署・
    ' グループ名まで)ため、既定値は空とし config で記入してもらう。
    modConfig.RegisterDefault "operator", vbNullString
    modConfig.RegisterDefault "feature_inbox", "TRUE"
    modConfig.RegisterDefault "feature_judgelog", "TRUE"
    ' 裁定書14 裁定6: 初回ガイドツアーを見終えたか("1"=済)。既定は "0"。
    modConfig.RegisterDefault "guide_tour_done", "0"
    ' 裁定書26 B/C/D(13章§2.3): 全画面表示・社内ディープリサーチのURL3本と
    ' [コピー]直後に開くか・部のポータル。
    modConfig.RegisterDefault "ui_fullscreen", "TRUE"
    modConfig.RegisterDefault "dr_url_menu", "https://app.hdtech.jp/research/menu"
    modConfig.RegisterDefault "dr_url_quick", _
        "https://app.hdtech.jp/research/quick-search"
    modConfig.RegisterDefault "dr_url_full", _
        "https://app.hdtech.jp/research/instructions"
    modConfig.RegisterDefault "dr_open_after_copy", "TRUE"
    modConfig.RegisterDefault "portal_url", _
        "http://www.portal.s1.ms-ad-ins.co.jp/loader/hp/OpenContents/" & _
        "A201203280048/toppage.html"
End Sub

' ----------------------------------------------------------------------------
' 画面ゴースト化抑止(16章E-50(b))。config keep_window_alive でオプトアウト可。
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' ApplyGhostingGuard - 画面ゴースト化への備え(16章 E-50(b)。裁定書27 W9-B4)
'   撤去したもの: user32 の `DisableProcessWindowsGhosting`(`Declare PtrSafe`)。
'     Win32 APIの宣言は社内AVのAMSIがマクロ型マルウェアの特徴として重く見る形
'     であり、配布物から消す(2026-09-02 実測。裁定書27 事実)。
'   代替: **事前描画カード + DoEvents**。E-50(a) の確定表示(modUIProgress.SetStage
'     が呼出の**前**に書き切るカード)が主役であり、ここでは起動時に一度
'     `DoEvents` を通してメッセージキューを空にし、以後の待機に入る前の画面を
'     描き切らせる。config `keep_window_alive`(既定TRUE)の意味は変えない
'     (FALSE ならこの手当てもしない)。
' ----------------------------------------------------------------------------
Private Sub ApplyGhostingGuard()
    If modConfig.GetBool("keep_window_alive", True) Then
        On Error Resume Next
        DoEvents
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
' (5の前) ナレッジブックの自動発見(裁定書17 H1 / 裁定書19 H8(b))。
'   config kb_path が (a)空 (b)プレースホルダ "\\...\" を含む (c)Dir$で不在
'   のいずれかなら、本体ブックと同じフォルダの BOOT_KB_FILE を kb_path へ
'   書いてから通常の読込へ進む。**kb_path がURLでも同じ扱い**にする
'   (URLだからといって「利用者が正しく設定した」とは限らないため。kb_path が
'    URLで実在するかは Dir$ で判定できないので、判定できるとき=ローカルの
'    ときだけ実在を見て、それ以外は本体フォルダ候補を優先する)。
'   本体側 ThisWorkbook.Path がURL形式("://" を含む=OneDrive同期フォルダ)の
'   ときは "/" で連結し、**Dir$ を通さずに** kb_path へ書く(URLは Dir$ で
'   必ず不発になり、書けないまま「置いたのに読めない」になるため)。実際に
'   開けるかどうかの判定は modKnowledge.LoadKnowledge が担う(fail-closed)。
'   探索と判定はこのPrivate 1本に閉じる(modBoot以外に起動処理を置かない)。
' ----------------------------------------------------------------------------
Private Sub ResolveKbPath()
    On Error Resume Next

    gKbAutoFound = False

    Dim pathText As String
    pathText = Trim$(modConfig.GetStr(BOOT_KB_KEY, vbNullString))
    If LenB(pathText) > 0 Then
        If InStr(1, pathText, BOOT_KB_PLACEHOLDER, vbBinaryCompare) = 0 Then
            If InStr(1, pathText, "://", vbBinaryCompare) = 0 Then
                If LenB(Dir$(pathText)) > 0 Then Exit Sub
            End If
        End If
    End If

    Dim baseDir As String
    baseDir = ThisWorkbook.Path
    If LenB(baseDir) = 0 Then Exit Sub

    Dim candidate As String
    If InStr(1, baseDir, "://", vbBinaryCompare) > 0 Then
        ' OneDrive同期フォルダ。Dir$ では見えないので確かめずに書く。
        candidate = baseDir & "/" & BOOT_KB_FILE
    Else
        ' 区切りは決め打ちにしない(Mac の実Excel は "/")。W9.2。
        candidate = baseDir & Application.PathSeparator & BOOT_KB_FILE
        If LenB(Dir$(candidate)) = 0 Then Exit Sub
    End If

    modConfig.SetValue BOOT_KB_KEY, candidate
    gKbAutoFound = True

    ' W9.2 N3: On Error Resume Next が握りつぶした失敗を記録だけは残す。
    If Err.Number <> 0 Then
        modLog.LogError "E0603", BOOT_SRC & ".ResolveKbPath", "kb_resolve_failed", Err.Number
        Err.Clear
    End If
End Sub

' ----------------------------------------------------------------------------
' NoticeDataDir - 保存先がOneDriveでないときだけ、ナビのお知らせへ warn を出す
'   (裁定書27 W9-C2)。会社PCの `D:` はシャットダウンで消え、Documents が残るか
'   はOneDriveのリダイレクト設定次第で未測定であるため、**黙って Documents へ
'   書かない**。解決そのものは modUtil.ResolveDataDir が唯一持つ。
' ----------------------------------------------------------------------------
Private Sub NoticeDataDir()
    On Error Resume Next

    Dim raw As String
    raw = Trim$(modConfig.GetStr(BOOT_OUT_DIR_KEY, vbNullString))
    If LenB(raw) = 0 Then raw = Trim$(modConfig.GetStr(BOOT_DATA_DIR_KEY, BOOT_DATA_DIR_DEFAULT))
    If LenB(raw) = 0 Then raw = BOOT_DATA_DIR_DEFAULT

    Dim dirText As String
    dirText = modUtil.ResolveDataDir(raw)
    If LenB(dirText) = 0 Then Exit Sub
    If Not modUtil.DataDirNotOneDrive(dirText) Then Exit Sub

    modUIHome.ShowWarning BOOT_MSG_NOT_ONEDRIVE, "warn"
End Sub

' ----------------------------------------------------------------------------
' KbAutoNote - 起動時の自動発見でパスを書いたときだけ、HOMEのナレッジ欄の
'   先頭へ足す1文を返す(書かなかったときは空)。読むのは modUIHome の
'   KbStatusText だけ(裁定書17 H1・14章§6)。
' ----------------------------------------------------------------------------
Public Function KbAutoNote() As String
    If gKbAutoFound Then KbAutoNote = BOOT_KB_AUTO_NOTE
End Function

' ----------------------------------------------------------------------------
' LogKbPathTried - ナレッジ読込に失敗したとき、**実際に開こうとしたパス**を
'   err_log へ1行残す(裁定書19 H8(a)の代替措置)。modKnowledge は30,000字契約
'   の満杯モジュールで detail を伸ばせないため、記録はこちら(modBoot)で行う。
'   E0401 kb_open_failed(modKnowledge側)と対で読むと、探した場所が分かる。
' ----------------------------------------------------------------------------
Private Sub LogKbPathTried()
    On Error Resume Next
    modLog.LogError "E0401", BOOT_SRC & ".ResolveKbPath", _
        "kb_path_tried:" & modConfig.GetStr(BOOT_KB_KEY, vbNullString)
End Sub

' ----------------------------------------------------------------------------
' 業種入力規則(13章§2.11・裁定書9 §4。検証欠陥#1)
'   ナレッジのスナップショット(業種マスタ)から industry_code / industry_name を
'   隠しレンジへ複製し、案件入力の ci_industry_code / ci_industry_name へ入力
'   規則を張る(RestoreDataKeyHiddenRange と同作法)。
'   **ナレッジ未接続かつスナップショット無しのときは張らずに続行する**
'   (16章 E-08 準拠。起動は止めず E0401 の記録のみ)。
' ----------------------------------------------------------------------------
Private Sub RestoreIndustryHiddenRange()
    Dim codeList As String
    Dim nameList As String
    ReadIndustryMaster codeList, nameList
    If LenB(codeList) = 0 Then
        modLog.LogError "E0401", BOOT_SRC & ".RestoreIndustryHiddenRange", _
                        "industry_master_unavailable"
        Exit Sub
    End If

    Dim ws As Object
    Set ws = modUISheet.EnsureHiddenSheet(BOOT_ENUM_SHEET)
    If ws Is Nothing Then Exit Sub

    If LenB(modUISheet.PutEnumRange(ws, BOOT_IND_CODE_COL, BOOT_IND_CODE_NAME, _
                                    codeList, ";")) > 0 Then
        modUISheet.BindNamedValidation "ci_industry_code", BOOT_IND_CODE_NAME, _
                                       BOOT_DV_ERROR_TITLE, BOOT_DV_ERROR_MSG
    End If
    If LenB(modUISheet.PutEnumRange(ws, BOOT_IND_NAME_COL, BOOT_IND_NAME_NAME, _
                                    nameList, ";")) > 0 Then
        modUISheet.BindNamedValidation "ci_industry_name", BOOT_IND_NAME_NAME, _
                                       BOOT_DV_ERROR_TITLE, BOOT_DV_ERROR_MSG
    End If
End Sub

' スナップショットの業種マスタを ";" 区切りの2本にする(見出し行は読み飛ばす)。
'   1行 = [シート索引][行数][元の1行をvbTabで連結] (modKnowledge.SaveSnapshot)。
Private Sub ReadIndustryMaster(ByRef codeList As String, ByRef nameList As String)
    On Error GoTo Failed

    codeList = vbNullString
    nameList = vbNullString

    Dim ws As Object
    Set ws = modUISheet.SheetOf(BOOT_SNAP_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = modUISheet.LastRowOf(ws)
    If lastRow < 2 Then Exit Sub

    Dim r As Long
    Dim flds As Variant
    Dim cCode As Long
    Dim cName As Long
    cCode = -1
    cName = -1
    For r = 1 To lastRow
        If Trim$(modUISheet.CellText(ws, r, 1)) = BOOT_SNAP_IDX_INDUSTRY Then
            flds = Split(modUISheet.CellText(ws, r, 3), vbTab)
            If cCode < 0 Then
                cCode = FieldPos(flds, "industry_code")
                cName = FieldPos(flds, "industry_name")
                If cCode < 0 Or cName < 0 Then Exit Sub
            Else
                AppendIndustry codeList, nameList, flds, cCode, cName
            End If
        End If
    Next r
    Exit Sub

Failed:
    codeList = vbNullString
    nameList = vbNullString
End Sub

' 1行ぶんを2本のリストへ足す(どちらかが空・";"を含む行は採らない)。
Private Sub AppendIndustry(ByRef codeList As String, ByRef nameList As String, _
                           ByVal flds As Variant, ByVal cCode As Long, ByVal cName As Long)
    If cCode > UBound(flds) Or cName > UBound(flds) Then Exit Sub

    Dim codeText As String
    Dim nameText As String
    codeText = Trim$(CStr(flds(cCode)))
    nameText = Trim$(CStr(flds(cName)))
    If LenB(codeText) = 0 Or LenB(nameText) = 0 Then Exit Sub
    If InStr(1, codeText & nameText, ";", vbBinaryCompare) > 0 Then Exit Sub

    If LenB(codeList) > 0 Then codeList = codeList & ";"
    If LenB(nameList) > 0 Then nameList = nameList & ";"
    codeList = codeList & codeText
    nameList = nameList & nameText
End Sub

' 見出し行の中の位置(0起点)。見つからなければ -1。
Private Function FieldPos(ByVal flds As Variant, ByVal wanted As String) As Long
    FieldPos = -1
    Dim i As Long
    For i = LBound(flds) To UBound(flds)
        If Trim$(CStr(flds(i))) = wanted Then
            FieldPos = i
            Exit Function
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' (7) 利用上限がTrueのときのナビ区画④の「お知らせ」案内(14章§2・12章§2.1手順(7))。
'     hm_warning が無い(未着手のui構築段階)場合は無音でスキップする。
' ----------------------------------------------------------------------------
Private Sub NoticeLimitReached()
    On Error Resume Next
    Dim target As Object
    Set target = ThisWorkbook.Names("hm_warning").RefersToRange
    If Not target Is Nothing Then
        ' 文言の値源は modGatewayRPN.ErrMessageFor の1箇所(16章E-57・E0208。
        ' LimitCheck=Trueは「アドインの利用期限切れ」であって日次の利用枠
        ' (E0204)ではない。裁定書24 追補2)。
        modUtilText.SetCellSafe target, _
            modGatewayRPN.ErrMessageFor(modGatewayRPN.LimitCheckCode(True)), _
            "modBoot/hm_warning"
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
