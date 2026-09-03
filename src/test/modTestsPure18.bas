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

Public Sub RunAll()
    On Error GoTo FA
    T_W81A_CoachBandText
WB:
    On Error GoTo FB
    T_W81B_DrUrlDefault
WDone:
    Exit Sub
FA:
    GroupFail "W81A CoachBandText(裁定書26 A)"
    Resume WB
FB:
    GroupFail "W81B DrUrlOf の既定(裁定書26 C)"
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
