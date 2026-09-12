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
'   末尾から modTestsPure23.RunAll(W14 裁定書37 班1)を呼ぶ。
'
' Z51 群(全9本。裁定書41 §4 = 17章 Z-51 効果測定の機械計測)。ここへ置いたのは
'   本モジュールが W15 の3班(W1/W2/W3)のどの担当ファイルとも重ならないため。
'   期待値は裁定書41 §4 の文と 13章§2.2/§2.4 から手で書き出した(実装の出力を
'   見てから合わせない)。
'   04 同一なら 0%・両方空も 0%(「直していない」)
'   05 片方だけ空なら 100%
'   06 半分書き足したら 33%(a=6字 b=12字 共通6字 -> 100x(18-12)/18 = 33)
'   07 共通文字が1つも無ければ 100%
'   08 四捨五入は半分を切り上げる(銀行丸めを使わない)
'   09 戻り値は必ず 0〜100 に収まる(長い入力でも)
'   10 EditRatioNote の書式 `edit_ratio_sN=R` と、範囲外(段・率)は空文字
'   11 EditedStepNoOf は sN_edited だけ N を返し、sN_json や形違いは 0
'   12 Note を段の昇順に「;」でつなぐと 13章§2.4 の detail の形になる
'      (実際につなぐのは modCaseStore3.LogEditRatiosOnReport。シートを読むので
'       層(a)からは叩けず、ここで押さえるのは書式だけ)
'
' 変異注入(出来レース禁止・裁定書41):
'   (a) modLog.EditRatio の四捨五入 `+ 0.5` を落とすと 08 が落ちる(実測:
'       期待67に対し66。06 は 33.33 の切捨てでも 33 のままなので落ちない)。
'   (b) modLog.EditRatioNote の範囲検査(stepNo/ratio)を外すと 10 が落ちる(実測)。
'   (c) modCaseStore3.EditedStepNoOf の `_edited` の照合を「3文字目が _ か」
'       だけに緩めると 11 が落ちる(実測。`s1_editeD` / `s1_json` を拾う)。
'   (d) 配線側の網(層(a)の外): modCaseStore.SaveData から
'       modCaseStore3.LogEditRatioOnSave の呼び出しを消すと
'       `python3 tools/orphan_check.py` が 22件 -> 23件になり、
'       `ERROR(docs-only) src/app/modCaseStore3.bas:392 Sub LogEditRatioOnSave`
'       を出す(実測)。**この網が効くのは Debug.Print に自分の関数名を
'       書かないため**で、`"モジュール名.関数名"` を含む文字列を置くと
'       orphan_check がそれを呼び出しと見なして救済してしまう。
' グループ単位の失敗隔離: modTestsPure21 と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

Public Sub RunAll()
    On Error GoTo FA
    T_W12A_PointerParentOf
WZ:
    On Error GoTo FZ
    T_Z51_EditRatio
WB:
    On Error GoTo FB
    modTestsPure23.RunAll
WDone:
    Exit Sub
FA:
    GroupFail "W12A data_dir ポインタの親(裁定書34 §1.3(a))"
    Resume WZ
FZ:
    GroupFail "Z51 効果測定の差分率(裁定書41 §4)"
    Resume WB
FB:
    GroupFail "W14 班1(裁定書37)への結線"
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

' ============================================================================
' Z51-04..12 17章 Z-51 効果測定の機械計測(裁定書41 §4)
'   modLog.EditRatio / modLog.EditRatioNote / modCaseStore3.EditedStepNoOf。
' ============================================================================
Private Sub T_Z51_EditRatio()
    Dim r1 As Long, r2 As Long, r3 As Long, r4 As Long
    Dim longA As String, longB As String
    Dim i As Long
    Dim noteText As String

    modTestRunner.Check "Test_Z51_04_同一と両方空は0パーセント_裁定書41§4", _
        (modLog.EditRatio("あいうえお", "あいうえお") = 0) And _
        (modLog.EditRatio(vbNullString, vbNullString) = 0), _
        "直していない=0"

    modTestRunner.Check "Test_Z51_05_片方だけ空なら100パーセント_裁定書41§4", _
        (modLog.EditRatio(vbNullString, "あいう") = 100) And _
        (modLog.EditRatio("あいう", vbNullString) = 100), _
        "全とっかえ=100"

    ' a=6字 b=12字(a をそのまま含み6字足した)。共通6字。
    ' 100 x (6 + 12 - 2 x 6) / (6 + 12) = 100 x 6 / 18 = 33.33.. -> 33
    r1 = modLog.EditRatio("あいうえおか", "あいうえおかきくけこさし")
    modTestRunner.Check "Test_Z51_06_半分書き足すと33パーセント_裁定書41§4", _
        (r1 = 33), "期待=33 実際=" & CStr(r1)

    modTestRunner.Check "Test_Z51_07_共通文字が無ければ100パーセント_裁定書41§4", _
        (modLog.EditRatio("あいう", "かきく") = 100), "重なり0=100"

    ' a=3字 b=3字 共通1字(「あ」だけ)。
    ' 100 x (3 + 3 - 2 x 1) / (3 + 3) = 100 x 4 / 6 = 66.66.. -> 67(切り上げ)。
    ' VBA の Round は銀行丸め(66.5 -> 66 の側へ寄る)なので使っていないこと。
    r2 = modLog.EditRatio("あいう", "あかき")
    modTestRunner.Check "Test_Z51_08_四捨五入は半分を切り上げる_裁定書41§4", _
        (r2 = 67), "期待=67(66.66の切り上げ) 実際=" & CStr(r2)

    longA = vbNullString
    longB = vbNullString
    For i = 1 To 400
        longA = longA & "あいうえお"
        longB = longB & "かきくけこ"
    Next i
    r3 = modLog.EditRatio(longA, longB)
    r4 = modLog.EditRatio(longA, longA)
    modTestRunner.Check "Test_Z51_09_戻り値は0から100に収まる_裁定書41§4", _
        (r3 = 100) And (r4 = 0), _
        "2000字x2 で 重なり0=100 / 同一=0。実際=" & CStr(r3) & "/" & CStr(r4)

    modTestRunner.Check "Test_Z51_10_Noteの書式と範囲外は空文字_裁定書41§4", _
        (modLog.EditRatioNote(1, 0) = "edit_ratio_s1=0") And _
        (modLog.EditRatioNote(5, 100) = "edit_ratio_s5=100") And _
        (LenB(modLog.EditRatioNote(0, 10)) = 0) And _
        (LenB(modLog.EditRatioNote(6, 10)) = 0) And _
        (LenB(modLog.EditRatioNote(2, -1)) = 0) And _
        (LenB(modLog.EditRatioNote(2, 101)) = 0), _
        "書式2件と範囲外4件"

    modTestRunner.Check "Test_Z51_11_EditedStepNoOfはsN_editedだけNを返す_13章§2.2", _
        (modCaseStore3.EditedStepNoOf("s1_edited") = 1) And _
        (modCaseStore3.EditedStepNoOf("s5_edited") = 5) And _
        (modCaseStore3.EditedStepNoOf("s1_json") = 0) And _
        (modCaseStore3.EditedStepNoOf("s6_edited") = 0) And _
        (modCaseStore3.EditedStepNoOf("s1_editeD") = 0) And _
        (modCaseStore3.EditedStepNoOf("input_hp") = 0) And _
        (modCaseStore3.EditedStepNoOf(vbNullString) = 0), _
        "sN_edited 2件 / それ以外 5件"

    noteText = modLog.EditRatioNote(1, 33) & ";" & modLog.EditRatioNote(2, 0)
    modTestRunner.Check "Test_Z51_12_レポート行は段の昇順をセミコロンでつなぐ_13章§2.4", _
        (noteText = "edit_ratio_s1=33;edit_ratio_s2=0"), _
        "期待=edit_ratio_s1=33;edit_ratio_s2=0 実際=" & noteText
End Sub
