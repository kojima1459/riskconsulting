Attribute VB_Name = "modTestsPure19"
Option Explicit

' ============================================================================
' modTestsPure19 - 裁定書28(W10 データをブックの外へ)の純層テスト
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**裁定書28 の文だけ**から手で書き出した(17章§1。実装の
'   出力を見てから期待値を合わせない)。
'
' 対象と根拠:
'   W11B data_dir の解決 8本  裁定書31 裁定1(環境変数の全撤去)。
'     解決順は (1)本体と同じフォルダの `data_dir.txt` → (2)config `data_dir`
'     (`%…%` を含む値は展開せず捨てる) → (3)本体と同じフォルダ `\データ`。
'     **環境変数は1つも見ない**(会社PCの %OneDrive% は別利用者を指しうる)。
'       01 (1)が最優先で先頭、次に(2)、最後に(3)
'       02 config に `%…%` が入っていたら候補に入らない(負例)
'       03 (1)も(2)も無ければ (3)だけ
'       04 末尾の区切りは重ねない(ポインタ・本体フォルダの両方)
'       05 ポインタが空白だけなら「無し」と同じ
'       06 本体フォルダが空なら候補ゼロ(逃げ場を勝手に作らない。負例)
'       07 DataDirIsLastResort: (3)で解決したら True(=warn を出す条件)
'       08 DataDirIsLastResort: (1)で解決したら False(負例)
'   W10B FileCandidatesIn 1本  裁定書28「ナレッジブック探索順」。
'     本体と同じフォルダ → data_dir の親(=OneDrive の配布フォルダ)。
'       01 D:\リスク提案ナビ と ...\リスク提案ナビ\データ から2候補
'   W10C CsvLineOf 4本  裁定書28「ログ csv」＋司令塔裁定「標準CSVで書く」。
'     情シス・開発担当が Excel やメモ帳でそのまま開ける形にする=区切りは
'     カンマ・**各フィールドは常に**二重引用符で囲む・値の中の " は ""。
'     組立ては**フィールドの配列**から行う(値を連結してから割り直さない)。
'       01 ふつうの値も1つずつ二重引用符で囲む
'       02 値の中の引用符は "" にする(囲みはそのまま)
'       03 改行は半角空白へ潰す(1行=1レコードを崩さない)
'       04 値の中のカンマは引用符の中に残り、1フィールドのまま(列が増えない)
'   W10D ParseSettingsText 3本  裁定書28「設定.txt」。
'       01 # で始まる行はコメントとして落ちる
'       02 空行は落ちる(前後の空白も落ちる)
'       03 許可キー以外は無視する(log_max_rows は通さない)
'
' グループ単位の失敗隔離: modTestsPure.bas と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

' 裁定書31 裁定1 の解決順から手で書き出した部品。本体と同じフォルダ(=ランチャー
'   が写した D: 側)と、そこへ作る最後の逃げ場(3)。
Private Const P19_BOOK As String = "D:\リスク提案ナビ"
Private Const P19_LAST As String = "D:\リスク提案ナビ\データ"
' 利用者が config へ手で書いた明示値(2)と、環境変数入りの無効値(捨てる側)。
Private Const P19_CFG As String = "C:\手動で決めた保存先"
Private Const P19_CFG_ENV As String = "%OneDrive%\リスク提案ナビ\データ"
'  CSVの二重引用符1文字(期待値を数え違えないための部品)。
Private Const P19_Q As String = """"
' ランチャー(.bat)が data_dir.txt へ書く値の形(裁定書28「裁定の確定」3)。
Private Const P19_POINTER As String = "C:\Users\u\OneDrive - 会社\リスク提案ナビ\データ"

Public Sub RunAll()
    On Error GoTo FA
    T_W11B_DataDirCandidates
WB:
    On Error GoTo FB
    T_W10B_FileCandidatesIn
WC:
    On Error GoTo FC
    T_W10C_CsvLineOf
WD:
    On Error GoTo FD
    T_W10D_ParseSettingsText
WE:
    On Error GoTo FE
    modTestsPure20.RunAll
WDone:
    Exit Sub
FA:
    GroupFail "W11B data_dir の解決(裁定書31 裁定1)"
    Resume WB
FB:
    GroupFail "W10B FileCandidatesIn(裁定書28 ナレッジブック探索順)"
    Resume WC
FC:
    GroupFail "W10C CsvLineOf(裁定書28 ログ csv)"
    Resume WD
FD:
    GroupFail "W10D ParseSettingsText(裁定書28 設定.txt)"
    Resume WE
FE:
    GroupFail "modTestsPure20(W10・裁定書28 企業ファイル)"
    Resume WDone
End Sub

Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), "期待=[" & want & "] 実際=[" & act & "]"
End Sub

' ============================================================================
' W11B data_dir の解決(裁定書31 裁定1。環境変数を1つも見ない)
' ============================================================================
Private Sub T_W11B_DataDirCandidates()
    ChkS "Test_W11B_01_ポインタconfig逃げ場の順_裁定書31", _
        modUtil.DataDirCandidates(P19_POINTER, P19_CFG, P19_BOOK), _
        P19_POINTER & vbLf & P19_CFG & vbLf & P19_LAST

    ChkS "Test_W11B_02_configの環境変数入りは捨てる_裁定書31", _
        modUtil.DataDirCandidates(vbNullString, P19_CFG_ENV, P19_BOOK), _
        P19_LAST

    ChkS "Test_W11B_03_ポインタもconfigも無ければ逃げ場だけ_裁定書31", _
        modUtil.DataDirCandidates(vbNullString, vbNullString, P19_BOOK), _
        P19_LAST

    ChkS "Test_W11B_04_末尾の区切りは重ねない_裁定書31", _
        modUtil.DataDirCandidates(P19_POINTER & "\", vbNullString, P19_BOOK & "\"), _
        P19_POINTER & vbLf & P19_LAST

    ChkS "Test_W11B_05_ポインタが空白だけなら無しと同じ_裁定書31", _
        modUtil.DataDirCandidates("   ", P19_CFG, P19_BOOK), _
        P19_CFG & vbLf & P19_LAST

    ChkS "Test_W11B_06_本体フォルダが空なら候補ゼロ_裁定書31", _
        modUtil.DataDirCandidates(vbNullString, P19_CFG_ENV, vbNullString), _
        vbNullString

    ChkS "Test_W11B_07_逃げ場へ落ちたらwarnを出す_裁定書31", _
        CStr(modUtil.DataDirIsLastResort(P19_LAST, P19_BOOK)), CStr(True)

    ChkS "Test_W11B_08_ポインタで解決したらwarnを出さない_裁定書31", _
        CStr(modUtil.DataDirIsLastResort(P19_POINTER, P19_BOOK)), CStr(False)
End Sub

' ============================================================================
' W10B FileCandidatesIn(本体と同じフォルダ → data_dir の親)
' ============================================================================
Private Sub T_W10B_FileCandidatesIn()
    ChkS "Test_W10B_01_本体フォルダの次はdata_dirの親_裁定書28", _
        modUtil.FileCandidatesIn("D:\リスク提案ナビ", _
                                 "C:\OD-Biz\リスク提案ナビ\データ", "\", _
                                 "ナレッジブック.xlsx"), _
        "D:\リスク提案ナビ\ナレッジブック.xlsx" & vbLf & _
        "C:\OD-Biz\リスク提案ナビ\ナレッジブック.xlsx"
End Sub

' ============================================================================
' W10C CsvLineOf(1行=1レコードを崩さない無害化)
' ============================================================================
Private Sub T_W10C_CsvLineOf()
    Dim f3(0 To 2) As String
    f3(0) = "2026-09-03 10:00:00"
    f3(1) = "E0603"
    f3(2) = "modBoot.Boot"
    ChkS "Test_W10C_01_全フィールドを二重引用符で囲む_司令塔裁定", _
        modLog.CsvLineOf(f3), _
        P19_Q & "2026-09-03 10:00:00" & P19_Q & "," & P19_Q & "E0603" & P19_Q & _
        "," & P19_Q & "modBoot.Boot" & P19_Q

    Dim fQuote(0 To 1) As String
    fQuote(0) = "a" & P19_Q & "b"
    fQuote(1) = "cd"
    ChkS "Test_W10C_02_値の中の引用符は二重化する_司令塔裁定", _
        modLog.CsvLineOf(fQuote), _
        P19_Q & "a" & P19_Q & P19_Q & "b" & P19_Q & "," & P19_Q & "cd" & P19_Q

    Dim fBreak(0 To 1) As String
    fBreak(0) = "a" & vbCrLf & "b"
    fBreak(1) = "c"
    ChkS "Test_W10C_03_値の中の改行は半角空白へ潰す_裁定書28", _
        modLog.CsvLineOf(fBreak), _
        P19_Q & "a b" & P19_Q & "," & P19_Q & "c" & P19_Q

    Dim fComma(0 To 1) As String
    fComma(0) = "a,b"
    fComma(1) = "c"
    ChkS "Test_W10C_04_値の中のカンマは1フィールドのまま_司令塔裁定", _
        modLog.CsvLineOf(fComma), _
        P19_Q & "a,b" & P19_Q & "," & P19_Q & "c" & P19_Q
End Sub

' ============================================================================
' W10D ParseSettingsText(コメント/空行/未許可キー)
' ============================================================================
Private Sub T_W10D_ParseSettingsText()
    ChkS "Test_W10D_01_シャープ始まりの行はコメント_裁定書28", _
        modConfig.ParseSettingsText("# portal_url=https://ng" & vbLf & _
                                    "portal_url=https://ok"), _
        "portal_url=https://ok"

    ChkS "Test_W10D_02_空行と前後の空白は落ちる_裁定書28", _
        modConfig.ParseSettingsText(vbLf & "  ui_fullscreen = TRUE  " & vbLf & vbLf), _
        "ui_fullscreen=TRUE"

    ChkS "Test_W10D_03_許可キー以外は無視する_裁定書28", _
        modConfig.ParseSettingsText("log_max_rows=1" & vbLf & "kb_path=D:\a.xlsx"), _
        "kb_path=D:\a.xlsx"
End Sub
