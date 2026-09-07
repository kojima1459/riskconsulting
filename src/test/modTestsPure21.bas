Attribute VB_Name = "modTestsPure21"
Option Explicit

' ============================================================================
' modTestsPure21 - W11-c(リボンの抽出切断・裁定書33)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**相手側の仕様(裁定書33 §0・14章§2 の 7.)だけ**から手で
'   書き出した(17章§1。実装の出力を見てから期待値を合わせない)。切断後の
'   文字列は、相手側の終点規則(`"},` の**手前**まで)を紙の上でたどって書いた。
'
' 対象と根拠(全7本。W11C 群):
'   01 1行詰めJSONは往復で**切れる**。`{"a":[{"n":"x"},{"n":"y"}]}` はボディ上で
'      `\"},\"` になり、終点候補 `"},` が `\"` の直後に部分一致する。切り出しは
'      その `"` の**手前**で終わるので、戻りは `{"a":[{"n":"x\` になる
'      (末尾の `\` は、切り口に残ったバックスラッシュを UnEscapeJSON が
'      「不正エスケープはそのまま」で戻したもの)。
'   02 同じ内容を整形(`}` `]` の直前で改行)すると往復で**一致**する
'      (`\"\n}` になり `"},` が現れないため)。C-1 の整形規則の根拠。
'   03 本文中の `": "` は往復で `":"` になる(相手側が**ボディ全体**へ掛ける
'      3置換の1つ。既知の劣化であって切断ではない)。
'   04 `\uXXXX` と `\"` が実体へ戻る(UnEscapeJSON の写しが効いていること)。
'   05 mock 11本(15章§8.1 の MK-*)が**全部**往復で一致する。1本でも切れたら
'      その ID を detail に出す(どれが1行詰めなのかが分かるように)。
'   06 mock_fault のうちJSONを返す6値も同様に一致する(JSONでない `empty` /
'      `limit` / `ribbon_*` は対象外)。
'   07 modRibbonWire.LooksRibbonCut の正例2(切れた形)・負例2(閉じている形・空)。
'
' グループ単位の失敗隔離: modTestsPure.bas と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

' 15章§8.1 の mock ID(11本。表から手で写した)。
Private Const P21_IDS As String = "MK-S1-NEW;MK-S1-RNW;MK-S2-NEW;MK-S2-RNW;" & _
    "MK-S3;MK-S4;MK-PF;MK-S2C-HIT;MK-S2C-CLEAN;MK-S3C-HIT;MK-S3C-CLEAN"

' 15章§8.2 のうち**JSONを返す**fault値と、それが効くstep(`値/step` の並び)。
Private Const P21_FAULTS As String = "broken_json/s1;broken_json_once/s1;" & _
    "enum_violation/s2;count_violation/s3;ghost_id/s3;fake_err/s1"

Public Sub RunAll()
    On Error GoTo FA
    T_W11C_CutAndPretty
WB:
    On Error GoTo FB
    T_W11C_ReplaceAndUnescape
WC:
    On Error GoTo FC
    T_W11C_MockRoundTrip
WD:
    On Error GoTo FD
    T_W11C_LooksRibbonCut
WDone:
    Exit Sub
FA:
    GroupFail "W11C 切断と整形(裁定書33 C-3 01/02)"
    Resume WB
FB:
    GroupFail "W11C 3置換とUnEscape(裁定書33 C-3 03/04)"
    Resume WC
FC:
    GroupFail "W11C mock の往復(裁定書33 C-3 05/06)"
    Resume WD
FD:
    GroupFail "W11C LooksRibbonCut(裁定書33 C-4)"
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

' 1行に詰めたS1風JSON(`"},` を含む)。
Private Function PackedJson() As String
    PackedJson = "{""a"":[{""n"":""x""},{""n"":""y""}]}"
End Function

' 同じ内容を整形したもの(`}` と `]` の直前で改行する=C-1 の規則)。
Private Function PrettyJson() As String
    PrettyJson = "{""a"":[{""n"":""x""" & vbLf & "},{""n"":""y""" & vbLf & _
                 "}" & vbLf & "]" & vbLf & "}"
End Function

' ============================================================================
' W11C-01/02 切断と整形
' ============================================================================
Private Sub T_W11C_CutAndPretty()
    Dim packed As String
    Dim cut As String

    packed = PackedJson()
    cut = modRibbonSim.SimRoundTrip(packed)

    ' 01 1行詰めは `"},` の手前で切られる(元と一致しない)。
    modTestRunner.Check "Test_W11C_01_1行詰めJSONはリボンの往復で切れる_裁定書33", _
        (cut <> packed) And (cut = "{""a"":[{""n"":""x\"), _
        "期待=[{""a"":[{""n"":""x\] 実際=[" & cut & "]"

    ' 02 整形すれば切られない。
    ChkS "Test_W11C_02_整形JSONはリボンの往復で一致する_裁定書33", _
        modRibbonSim.SimRoundTrip(PrettyJson()), PrettyJson()
End Sub

' ============================================================================
' W11C-03/04 相手側の3置換と UnEscapeJSON
' ============================================================================
Private Sub T_W11C_ReplaceAndUnescape()
    ' 03 本文中の `": "` はボディ全体への置換で `":"` になる(既知の劣化)。
    ChkS "Test_W11C_03_本文中のコロン空白が詰められる_裁定書33", _
        modRibbonSim.SimRoundTrip("見出し: 値"), "見出し:値"

    ' 04 `\uXXXX` と `\"` が実体へ戻る(ボディを手で組んで通す)。
    ChkS "Test_W11C_04_uXXXXとエスケープ引用符が実体へ戻る_裁定書33", _
        modRibbonSim.SimParseText("{""choices"":[{""message"":{""content"":" & _
            """\u3042\""x\"""",""role"":""assistant""}}]}"), _
        "あ""x"""
End Sub

' ============================================================================
' W11C-05/06 mock 応答の往復(1本でも切れたら FAIL・IDを出す)
' ============================================================================
Private Sub T_W11C_MockRoundTrip()
    Dim ids As Variant
    Dim i As Long
    Dim body As String
    Dim bad As String
    Dim cnt As Long

    ids = Split(P21_IDS, ";")
    For i = LBound(ids) To UBound(ids)
        body = modMockLlm.ResponseById(CStr(ids(i)))
        If LenB(body) = 0 Then
            bad = bad & CStr(ids(i)) & "(応答なし) "
        Else
            cnt = cnt + 1
            If modRibbonSim.SimRoundTrip(body) <> body Then bad = bad & CStr(ids(i)) & " "
        End If
    Next i

    modTestRunner.Check "Test_W11C_05_mock11本がリボンの往復で全部一致する_裁定書33", _
        (LenB(bad) = 0) And (cnt = 11), _
        "切れた/欠けたID=[" & Trim$(bad) & "] 往復した本数=" & CStr(cnt) & "(期待11)"

    ' 06 mock_fault のうちJSONを返すもの。
    Dim faults As Variant
    Dim pair As Variant
    Dim badF As String
    Dim cntF As Long

    modMockLlm.ResetFaultOnce
    faults = Split(P21_FAULTS, ";")
    For i = LBound(faults) To UBound(faults)
        pair = Split(CStr(faults(i)), "/")
        body = modMockLlm.FaultResponse(CStr(pair(0)), CStr(pair(1)))
        If LenB(body) = 0 Then
            badF = badF & CStr(pair(0)) & "(応答なし) "
        Else
            cntF = cntF + 1
            If modRibbonSim.SimRoundTrip(body) <> body Then badF = badF & CStr(pair(0)) & " "
        End If
    Next i
    ' 1bitの状態(broken_json_once)を触ったので必ず戻す(15章§8.2)。
    modMockLlm.ResetFaultOnce

    modTestRunner.Check "Test_W11C_06_mock_faultのJSON応答も往復で一致する_裁定書33", _
        (LenB(badF) = 0) And (cntF = 6), _
        "切れた/欠けたfault=[" & Trim$(badF) & "] 往復した本数=" & CStr(cntF) & "(期待6)"
End Sub

' ============================================================================
' W11C-07 LooksRibbonCut(正例2・負例2)
' ============================================================================
Private Sub T_W11C_LooksRibbonCut()
    Dim cut As String
    cut = modRibbonSim.SimRoundTrip(PackedJson())

    modTestRunner.Check "Test_W11C_07_LooksRibbonCutの正例2と負例2_裁定書33", _
        modRibbonWire.LooksRibbonCut(cut) And _
        modRibbonWire.LooksRibbonCut("{""a"":[{""n"":""x") And _
        (Not modRibbonWire.LooksRibbonCut("{""a"":[{""n"":""x""}]}")) And _
        (Not modRibbonWire.LooksRibbonCut(vbNullString)), _
        "正=切れた本文2種(実切断・引用符の途中)/負=閉じたJSON・空文字"
End Sub
