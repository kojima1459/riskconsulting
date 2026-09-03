Attribute VB_Name = "modTestsPure18"
Option Explicit

' ============================================================================
' modTestsPure18 - W8.1(実機第1報・裁定書26)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**仕様の文だけ**から手計算した(17章§1。実装の出力を見て
'   から期待値を合わせない)。
'
' 対象と根拠:
'   W81A CoachBandText 3本  裁定書26 A・11章§3.1・§8.5 #16。
'     帯の文字は「STEP n/6 ＋空白2つ ＋進捗ドット」で1行、次の一手で1行。
'       01 n/6の桁   1行目の空白2つより前が `STEP 2/6`(2桁の 10/10 も見る)
'       02 ドット     1行目の空白2つより後が 11章§3.1.1 の `●●○○○○`
'                     (6段のうち2段目まで塗る=●2つ＋○4つ。手計算)
'       03 改行数     vbLf はちょうど1つで、2行目は hm_next_action の逐語
'   W81B DrUrlOf 2本        裁定書26 C・13章§2.3。
'     config が空・欠落のときに13章§2.3の既定URLへ倒すこと(2本)。
'   W9B1 JoinPasteCells 4本 裁定書27 W9-B1。受け皿シートに貼り付いたセルを
'     1本のテキストへ戻す規則(modNavText の宣言した4規則)から手で作った。
'       01 途中の空行は残す        A/空/B -> "A" LF LF "B"
'       02 末尾の空行は全部落とす  A/B/空/空 -> "A" LF "B"
'       03 1セル32,767字の境界     32,767字 + LF + 4字 = 32,772字(切らない)
'       04 列はタブで戻す          2列 -> "a" TAB "b"(右端の空セルは落とす)
'   W9B2 Utf8Bytes 8本      裁定書27 W9-B2。期待値は RFC 3629 の表から
'     **手計算したバイト列**で、実装の出力は見ていない。
'       01 BOM有   "A"          -> EF BB BF 41
'       02 BOM無   "Az0"        -> 41 7A 30
'       03 2バイト U+00A9       -> C2 A9      (169 = 192+2, 128+41)
'       04 3バイト U+3042(あ)   -> E3 81 82   (12354 = 224+3, 128+1, 128+2)
'       05 4バイト U+20B9F      -> F0 A0 AE 9F(134047。上位D842+下位DF9F)
'       06 CP932外 U+9FA6       -> E9 BE A6   (40870 = 224+9, 128+62, 128+38)
'       07 孤立サロゲート        -> EF BF BD  (U+FFFD へ落とす)
'       08 Utf8Len  "A"+BOM     -> 4
'   W9C2 DataDirCandidates 3本 裁定書27 W9-C2 の解決順(1)(2)(3)。
'     環境変数の有無だけで並びが決まることを、値を手で与えて固定する。
'
' グループ単位の失敗隔離: modTestsPure.bas と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

' 11章§3.1.1 の進捗ドット(6段のうち2段目まで)。表から手で写した逐語。
Private Const P18_DOTS_2_6 As String = "●●○○○○"
' 11章§3.1.1 の STEP 3/6 の文(逐語表の3行目)。帯の2行目に入る文。
Private Const P18_ACTION As String = _
    "返ってきた文章を、②の枠へ貼ってください。長くても分けなくて大丈夫です。"
' 13章§2.3 の既定URL(表から手で写した逐語)。
Private Const P18_URL_FULL As String = "https://app.hdtech.jp/research/instructions"
Private Const P18_URL_QUICK As String = "https://app.hdtech.jp/research/quick-search"
' 13章§2.3 の data_dir 既定(表から手で写した逐語)と、そこから作る期待値の部品。
Private Const P18_DD_RAW As String = "%OneDriveCommercial%\リスク提案ナビ\データ"
Private Const P18_DD_TAIL As String = "\リスク提案ナビ\データ"
Private Const P18_DD_LAST As String = "\Documents\RPN出力"

Public Sub RunAll()
    On Error GoTo FA
    T_W81A_CoachBandText
WB:
    On Error GoTo FB
    T_W81B_DrUrlDefault
WC:
    On Error GoTo FC
    T_W9B1_JoinPasteCells
WD:
    On Error GoTo FD
    T_W9B2_Utf8Bytes
WE:
    On Error GoTo FE
    T_W9C2_DataDirCandidates
WDone:
    Exit Sub
FA:
    GroupFail "W81A CoachBandText(裁定書26 A)"
    Resume WB
FB:
    GroupFail "W81B DrUrlOf の既定(裁定書26 C)"
    Resume WC
FC:
    GroupFail "W9B1 JoinPasteCells(裁定書27 W9-B1)"
    Resume WD
FD:
    GroupFail "W9B2 Utf8Bytes(裁定書27 W9-B2)"
    Resume WE
FE:
    GroupFail "W9C2 DataDirCandidates(裁定書27 W9-C2)"
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
' W81A CoachBandText(裁定書26 A)
' ============================================================================
Private Sub T_W81A_CoachBandText()
    Dim t1 As String
    Dim t2 As String
    t1 = modUIGeom.CoachBandText(2, 6, P18_ACTION)
    t2 = modUIGeom.CoachBandText(10, 10, P18_ACTION)

    ' 01 n/6の桁。1行目の「空白2つ」より前だけを見る。
    ChkS "Test_W81A_01_1行目の桁は STEP n 斜線 総数_裁定書26A", _
        HeadOf(t1) & "|" & HeadOf(t2), "STEP 2/6|STEP 10/10"

    ' 02 進捗ドット。1行目の「空白2つ」より後だけを見る。
    ChkS "Test_W81A_02_1行目のドットは6文字で2つ塗る_11章§3.1.1", _
        DotsOf(t1), P18_DOTS_2_6

    ' 03 改行はちょうど1つで、2行目は次の一手の逐語。
    ChkS "Test_W81A_03_改行1つで2行目は次の一手の逐語_裁定書26A", _
        CStr(LineCount(t1)) & "|" & TailOf(t1), "2|" & P18_ACTION
End Sub

' 1行目の「空白2つ」より前(見つからなければ全体)。
Private Function HeadOf(ByVal bandText As String) As String
    Dim line1 As String
    line1 = Split(bandText, vbLf)(0)
    Dim p As Long
    p = InStr(1, line1, "  ", vbBinaryCompare)
    If p <= 0 Then
        HeadOf = line1
    Else
        HeadOf = Left$(line1, p - 1)
    End If
End Function

' 1行目の「空白2つ」より後(見つからなければ空)。
Private Function DotsOf(ByVal bandText As String) As String
    Dim line1 As String
    line1 = Split(bandText, vbLf)(0)
    Dim p As Long
    p = InStr(1, line1, "  ", vbBinaryCompare)
    If p <= 0 Then Exit Function
    DotsOf = Mid$(line1, p + 2)
End Function

' 2行目以降(vbLf の後ろ全部)。
Private Function TailOf(ByVal bandText As String) As String
    Dim p As Long
    p = InStr(1, bandText, vbLf, vbBinaryCompare)
    If p <= 0 Then Exit Function
    TailOf = Mid$(bandText, p + 1)
End Function

' 行数(vbLf の個数＋1)。
Private Function LineCount(ByVal bandText As String) As Long
    LineCount = UBound(Split(bandText, vbLf)) + 1
End Function

' ============================================================================
' W81B DrUrlOf の既定(裁定書26 C・13章§2.3)
' ============================================================================
Private Sub T_W81B_DrUrlDefault()
    ' config の行が無い(空文字が返る)ときは13章§2.3の既定へ倒す。
    ChkS "Test_W81B_01_config欠落ならしっかり調査の既定URL_13章§2.3", _
        modUIResearch.DrUrlOf("full", vbNullString), P18_URL_FULL
    ' 空白だけの値も「未設定」として既定へ倒す。
    ChkS "Test_W81B_02_config空白ならクイック調査の既定URL_13章§2.3", _
        modUIResearch.DrUrlOf("quick", "   "), P18_URL_QUICK
End Sub

' ============================================================================
' W9B1 JoinPasteCells(裁定書27 W9-B1・modNavText の4規則)
' ============================================================================
Private Sub T_W9B1_JoinPasteCells()
    ' 01 途中の空行は残す(1列3行)。
    Dim c1(0 To 2) As String
    c1(0) = "A": c1(1) = vbNullString: c1(2) = "B"
    ChkS "Test_W9B1_01_途中の空行は残す_裁定書27W9B1", _
        modNavText.JoinPasteCells(c1, 3, 1), "A" & vbLf & vbLf & "B"

    ' 02 末尾の空行は全部落とす(1列4行)。
    Dim c2(0 To 3) As String
    c2(0) = "A": c2(1) = "B": c2(2) = vbNullString: c2(3) = vbNullString
    ChkS "Test_W9B1_02_末尾の空行は全部落とす_裁定書27W9B1", _
        modNavText.JoinPasteCells(c2, 4, 1), "A" & vbLf & "B"

    ' 03 1セル32,767字の境界(連結側は切らない)。32767 + 1 + 4 = 32772字。
    Dim c3(0 To 1) As String
    c3(0) = String$(32767, "a"): c3(1) = "bcde"
    ChkS "Test_W9B1_03_32767字の境界で切らない_裁定書27W9B1", _
        CStr(Len(modNavText.JoinPasteCells(c3, 2, 1))), "32772"

    ' 04 列はタブで戻し、右端の空セルは落とす(2列2行)。
    Dim c4(0 To 3) As String
    c4(0) = "a": c4(1) = "b"
    c4(2) = "c": c4(3) = vbNullString
    ChkS "Test_W9B1_04_列はタブで戻し右端の空セルは落とす_裁定書27W9B1", _
        modNavText.JoinPasteCells(c4, 2, 2), "a" & vbTab & "b" & vbLf & "c"
End Sub

' ============================================================================
' W9B2 Utf8Bytes(裁定書27 W9-B2・RFC 3629 から手計算)
' ============================================================================
Private Sub T_W9B2_Utf8Bytes()
    ChkS "Test_W9B2_01_BOM有りは先頭にEFBBBF_裁定書27W9B2", _
        modUtilText.Utf8Hex("A", True), "efbbbf41"
    ChkS "Test_W9B2_02_BOM無しのASCIIは1バイト_裁定書27W9B2", _
        modUtilText.Utf8Hex("Az0", False), "417a30"
    ChkS "Test_W9B2_03_U00A9は2バイト_裁定書27W9B2", _
        modUtilText.Utf8Hex(ChrW(&HA9&), False), "c2a9"
    ChkS "Test_W9B2_04_U3042は3バイト_裁定書27W9B2", _
        modUtilText.Utf8Hex(ChrW(&H3042&), False), "e38182"
    ChkS "Test_W9B2_05_サロゲートペアは4バイト_裁定書27W9B2", _
        modUtilText.Utf8Hex(ChrW(&HD842&) & ChrW(&HDF9F&), False), "f0a0ae9f"
    ChkS "Test_W9B2_06_CP932外のBMP文字も3バイト_裁定書27W9B2", _
        modUtilText.Utf8Hex(ChrW(&H9FA6&), False), "e9bea6"
    ChkS "Test_W9B2_07_孤立サロゲートはUFFFDへ落とす_裁定書27W9B2", _
        modUtilText.Utf8Hex(ChrW(&HD842&), False), "efbfbd"
    ChkS "Test_W9B2_08_Utf8LenはBOMを数える_裁定書27W9B2", _
        CStr(modUtilText.Utf8Len("A", True)), "4"
End Sub

' ============================================================================
' W9C2 DataDirCandidates(裁定書27 W9-C2 の解決順)
' ============================================================================
Private Sub T_W9C2_DataDirCandidates()
    ChkS "Test_W9C2_01_会社OneDriveがあれば3候補_裁定書27W9C2", _
        modUtil.DataDirCandidates(P18_DD_RAW, "C:\OD-Biz", "C:\OD", "C:\Users\u"), _
        "C:\OD-Biz" & P18_DD_TAIL & vbLf & "C:\OD" & P18_DD_TAIL & vbLf & _
        "C:\Users\u" & P18_DD_LAST

    ChkS "Test_W9C2_02_会社OneDriveが無ければ個人OneDriveへ_裁定書27W9C2", _
        modUtil.DataDirCandidates(P18_DD_RAW, vbNullString, "C:\OD", "C:\Users\u"), _
        "C:\OD" & P18_DD_TAIL & vbLf & "C:\Users\u" & P18_DD_LAST

    ChkS "Test_W9C2_03_OneDriveが無ければDocumentsだけ_裁定書27W9C2", _
        modUtil.DataDirCandidates(P18_DD_RAW, vbNullString, vbNullString, "C:\Users\u"), _
        "C:\Users\u" & P18_DD_LAST
End Sub
