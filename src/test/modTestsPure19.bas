Attribute VB_Name = "modTestsPure19"
Option Explicit

' ============================================================================
' modTestsPure19 - 裁定書28(W10 データをブックの外へ)の純層テスト
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**裁定書28 の文だけ**から手で書き出した(17章§1。実装の
'   出力を見てから期待値を合わせない)。
'
' 対象と根拠:
'   W10A DataDirCandidates(pointer) 3本  裁定書28「裁定の確定」3・4。
'     解決順は `data_dir.txt` → config data_dir → %OneDriveCommercial% →
'     %OneDrive% → Documents。pointer が有るとき・無いとき・空文字のときで
'     並びがどう変わるかを固定する。
'       01 pointer 有 : 先頭が pointer、その後ろは従来の3候補
'       02 pointer 無 : 従来どおり(裁定書27 W9-C2 の3候補)
'       03 pointer 空 : 「無し」と同じ(空文字を候補に置かない)
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

' 13章§2.3 の data_dir 既定(表から手で写した逐語)と、そこから作る期待値の部品。
Private Const P19_DD_RAW As String = "%OneDriveCommercial%\リスク提案ナビ\データ"
Private Const P19_DD_TAIL As String = "\リスク提案ナビ\データ"
Private Const P19_DD_LAST As String = "\Documents\RPN出力"
'  CSVの二重引用符1文字(期待値を数え違えないための部品)。
Private Const P19_Q As String = """"
' ランチャー(.bat)が data_dir.txt へ書く値の形(裁定書28「裁定の確定」3)。
Private Const P19_POINTER As String = "C:\Users\u\OneDrive - 会社\リスク提案ナビ\データ"

Public Sub RunAll()
    On Error GoTo FA
    T_W10A_DataDirPointer
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
    GroupFail "W10A DataDirCandidates(裁定書28 data_dir.txt)"
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
' W10A DataDirCandidates(pointer 有/無/空)
' ============================================================================
Private Sub T_W10A_DataDirPointer()
    ChkS "Test_W10A_01_pointerが最優先で先頭に来る_裁定書28", _
        modUtil.DataDirCandidates(P19_POINTER, P19_DD_RAW, _
                                  "C:\OD-Biz", "C:\OD", "C:\Users\u"), _
        P19_POINTER & vbLf & "C:\OD-Biz" & P19_DD_TAIL & vbLf & _
        "C:\OD" & P19_DD_TAIL & vbLf & "C:\Users\u" & P19_DD_LAST

    ChkS "Test_W10A_02_pointerが無ければ従来の3候補_裁定書28", _
        modUtil.DataDirCandidates(vbNullString, P19_DD_RAW, _
                                  "C:\OD-Biz", "C:\OD", "C:\Users\u"), _
        "C:\OD-Biz" & P19_DD_TAIL & vbLf & "C:\OD" & P19_DD_TAIL & vbLf & _
        "C:\Users\u" & P19_DD_LAST

    ChkS "Test_W10A_03_pointerが空白だけなら無しと同じ_裁定書28", _
        modUtil.DataDirCandidates("   ", P19_DD_RAW, _
                                  vbNullString, vbNullString, "C:\Users\u"), _
        "C:\Users\u" & P19_DD_LAST
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
