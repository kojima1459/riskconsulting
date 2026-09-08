Attribute VB_Name = "modTestsPure22"
Option Explicit

' ============================================================================
' modTestsPure22 - W12-A(裁定書34 §1.3)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**裁定書34 §1.3(a) の文だけ**から手で書き出した(17章§1。
'   実装の出力を見てから期待値を合わせない)。
'
' 対象と根拠(全3本。W12A 群):
'   01 UNC(`\\srv\share\x\データ`)の親は `\\srv\share\x`。会社PCの配布フォルダは
'      OneDrive の同期先だが、共有フォルダから直接起動されることもある。
'   02 ドライブ付き(`D:\a\データ`)の親は `D:\a`。ランチャーが書く形そのもの。
'   03 段が1つしか無いもの(`D:` / `データ` / 空)の親は **空文字**。
'      「親を騙って同じ場所を返す」と、実在しないフォルダを保存先として採って
'      しまうので、ここは必ず空でなければならない(fail-closed)。
'      あわせて区切りが「/」のもの・混在するものも見る(Mac の実Excel と、
'      利用者が手で書き換えた data_dir.txt のため)。
'
' グループ単位の失敗隔離: modTestsPure21 と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

Public Sub RunAll()
    On Error GoTo FA
    T_W12A_PointerParentOf
WDone:
    Exit Sub
FA:
    GroupFail "W12A data_dir ポインタの親(裁定書34 §1.3(a))"
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
' W12A-01/02/03 modUtil.PointerParentOf(data_dir.txt の1行の親フォルダ)
' ============================================================================
Private Sub T_W12A_PointerParentOf()
    ChkS "Test_W12A_01_UNCの親は共有名の下の1段まで_裁定書34§1.3(a)", _
        modUtil.PointerParentOf("\\srv\share\x\データ"), "\\srv\share\x"

    ChkS "Test_W12A_02_ドライブ付きの親は1段上_裁定書34§1.3(a)", _
        modUtil.PointerParentOf("D:\a\データ"), "D:\a"

    modTestRunner.Check _
        "Test_W12A_03_親が取れないものは空文字で区切りは両方見る_裁定書34§1.3(a)", _
        (LenB(modUtil.PointerParentOf("D:")) = 0) And _
        (LenB(modUtil.PointerParentOf("データ")) = 0) And _
        (LenB(modUtil.PointerParentOf(vbNullString)) = 0) And _
        (modUtil.PointerParentOf("D:/a/データ") = "D:/a") And _
        (modUtil.PointerParentOf("D:\a/データ") = "D:\a") And _
        (modUtil.PointerParentOf("D:/a\データ") = "D:/a"), _
        "ドライブだけ・1段だけ・空は空文字 / 「/」区切りと混在は右の区切りで落とす、の6条件"
End Sub
