Attribute VB_Name = "modTestsPure23"
Option Explicit

' ============================================================================
' modTestsPure23 - W14(裁定書37 班1)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**裁定書37 §2 班1契約とB_dr_quality.md §3 B-01/§4だけ**
'   から手で書き出した(17章§1。実装の出力を見てから期待値を合わせない)。
'
' 対象と根拠(全4本。W14B1 群):
'   01 BlockGuard配線(B-01)。7 step種(S1/S2/S3/S4/PF/S2C/S3C)の system組立
'      結果が、末尾にちょうど modPromptsBlocks.BlockGuard() が付いて終わる
'      こと(`Right$(s, Len(g)) = g`。`InStrRev` ではない=15章§1.3の実装注記の
'      とおり)。呼び出しは modPipeline.bas/modPipeline2.bas/modPlayOps.bas の
'      各代入点と同じ組み方(modPromptsOps.AsmGuarded / AsmS4System)に倣う。
'   02 壁打ち(BuildSparringSystem)にはBlockGuardを付けないこと(15章§1.3の
'      唯一の例外)。「終わらないこと」を見る。
'   03 modPii.MatchUrlSpan(A-08/C-2)。出典URLを含む本文でHasPii=False。
'      伝書鳩20260912 3-2 の実例4本(DR出力に必ず出るPDFリンク形)。
'   04 03の対照。証券番号・企業コードらしき文字列は従来どおり検知すること
'      (既存挙動を壊していないことの回帰)。メールアドレスも引き続き検知。
'
' グループ単位の失敗隔離: modTestsPure22 と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

Public Sub RunAll()
    On Error GoTo FA
    T_W14B1_BlockGuardWiring
WB:
    On Error GoTo FB
    T_W14B1_SparringNotGuarded
WC:
    On Error GoTo FC
    T_W14A08_UrlNotDetected
WD:
    On Error GoTo FD
    T_W14A08_PolicyStillDetected
WDone:
    Exit Sub
FA:
    GroupFail "W14B1 BlockGuard配線(裁定書37 B-01)"
    Resume WB
FB:
    GroupFail "W14B1 壁打ちは非配線(15章§1.3)"
    Resume WC
FC:
    GroupFail "W14A08 URLスパン読み飛ばし(裁定書37 A-08/C-2)"
    Resume WD
FD:
    GroupFail "W14A08 従来検知の回帰(裁定書37 A-08/C-2)"
    Resume WDone
End Sub

Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

Private Sub ChkEndsGuard(ByVal nm As String, ByVal sysText As String, ByVal g As String)
    modTestRunner.Check nm, (Right$(sysText, Len(g)) = g), _
        "system末尾がBlockGuard()と一致しない(末尾" & Len(g) & "字を比較)"
End Sub

' ============================================================================
' W14B1-01 BlockGuard配線(裁定書37 B-01。9代入点のうち区別される7 step種)
' ============================================================================
Private Sub T_W14B1_BlockGuardWiring()
    Dim g As String
    Dim s4Note As String
    g = modPromptsBlocks.BlockGuard()

    ChkEndsGuard "Test_W14B1_01a_S1のsystem末尾がBlockGuard_裁定書37B-01", _
        modPromptsOps.AsmGuarded(modPromptsCore.BuildS1System()), g

    ChkEndsGuard "Test_W14B1_01b_S2のsystem末尾がBlockGuard_裁定書37B-01", _
        modPromptsOps.AsmGuarded(modPromptsCore.BuildS2System()), g

    ChkEndsGuard "Test_W14B1_01c_S3のsystem末尾がBlockGuard_裁定書37B-01", _
        modPromptsOps.AsmGuarded(modPromptsCore.BuildS3System()), g

    ChkEndsGuard "Test_W14B1_01d_S4のsystem末尾がBlockGuard_裁定書37B-01", _
        modPromptsOps.AsmS4System("proposal", "t1_quick", s4Note), g

    ChkEndsGuard "Test_W14B1_01e_PFのsystem末尾がBlockGuard_裁定書37B-01", _
        modPromptsOps.AsmGuarded(modPromptsOps.BuildPFSystem()), g

    ChkEndsGuard "Test_W14B1_01f_S2Cのsystem末尾がBlockGuard_裁定書37B-01", _
        modPromptsOps.AsmGuarded(modPromptsOps.BuildS2CriticSystem()), g

    ChkEndsGuard "Test_W14B1_01g_S3Cのsystem末尾がBlockGuard_裁定書37B-01", _
        modPromptsOps.AsmGuarded(modPromptsOps.BuildS3CriticSystem()), g
End Sub

' ============================================================================
' W14B1-02 壁打ち(BuildSparringSystem)にはBlockGuardを付けない(15章§1.3)
' ============================================================================
Private Sub T_W14B1_SparringNotGuarded()
    Dim g As String
    Dim s As String
    g = modPromptsBlocks.BlockGuard()
    s = modPromptsOps.BuildSparringSystem()

    modTestRunner.Check _
        "Test_W14B1_02_壁打ちsystemはBlockGuardで終わらない_15章§1.3", _
        (Right$(s, Len(g)) <> g), _
        "壁打ちにBlockGuardが付いてしまっている(15章§1.3の唯一の例外)"
End Sub

' ============================================================================
' W14A08-03 modPii.HasPii: URLスパンは走査対象外(裁定書37 A-08/C-2)
' ============================================================================
Private Sub T_W14A08_UrlNotDetected()
    modTestRunner.Check _
        "Test_W14A08_03a_出典URL単体は検知しない_裁定書37A-08C-2", _
        Not modPii.HasPii("https://www.example.co.jp/ir/news/nr20260901.pdf"), _
        "URLパス中の数字列を証券番号として誤検知している"

    modTestRunner.Check _
        "Test_W14A08_03b_決算資料URLは検知しない_裁定書37A-08C-2", _
        Not modPii.HasPii("https://example.com/data/kessan-20260331.pdf"), _
        "URLパス中の数字列を証券番号として誤検知している"

    modTestRunner.Check _
        "Test_W14A08_03c_本文中に埋め込まれたURLも検知しない_裁定書37A-08C-2", _
        Not modPii.HasPii("資料 https://example.com/documents/pdf20260401001.pdf を参照"), _
        "URLパス中の数字列を証券番号として誤検知している"

    modTestRunner.Check _
        "Test_W14A08_03d_連番URLは検知しない_裁定書37A-08C-2", _
        Not modPii.HasPii("https://x.example/2024/0000123456789/"), _
        "URLパス中の連番を証券番号として誤検知している"
End Sub

' ============================================================================
' W14A08-04 対照: 証券番号・企業コード・メールは従来どおり検知(回帰)
' ============================================================================
Private Sub T_W14A08_PolicyStillDetected()
    modTestRunner.Check _
        "Test_W14A08_04a_企業コードは引き続き検知_裁定書37A-08C-2", _
        modPii.HasPii("帝国データバンク 企業コード JP123456789"), _
        "URL対策の副作用で通常の企業コード検知が壊れている"

    modTestRunner.Check _
        "Test_W14A08_04b_証券番号は引き続き検知_裁定書37A-08C-2", _
        modPii.HasPii("証券番号 AB-1234567"), _
        "URL対策の副作用で通常の証券番号検知が壊れている"

    modTestRunner.Check _
        "Test_W14A08_04c_メールアドレスは引き続き検知_裁定書37A-08C-2", _
        modPii.HasPii("担当: taro.yamada@example.co.jp までご連絡ください"), _
        "URL対策の副作用でメールアドレス検知が壊れている"
End Sub
