Attribute VB_Name = "modTestsPure15"
Option Explicit

' ============================================================================
' modTestsPure15 - W7(裁定書25・17章 T-55)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は 15章 v2.6 の§11 検証ルール表の規約から手計算した。
'   実装の出力を見てから期待値を合わせることは禁止(17章§1)。
'
' 対象と根拠(15章§11。新設5件・改訂2件):
'   W7A CheckS1   3本  V-S1-12(certainty の enum)/ V-S1-13(financials.source の
'                      enum)/ 正例1本(全件 assumed の新規案件で V-S1-04 が
'                      発火しないこと= 裁定書25 S1 の要)
'   W7B CheckS2   5本  V-S2-12b(新規案件の gap_type は uninsured のみ)/
'                      V-S2-18(純資産既知なのに loss_scale_note が全件空)/
'                      正例3本(新規で全件 uninsured なら不合格にならない /
'                      対比があれば V-S2-18 は出ない / emerging_risks 5件は合格
'                      = 裁定書25 S7 の閾値5)
'   W7C CheckS3   4本  V-S3-19(flow 3～5件)/ V-S3-20(opening 1～80字)/
'                      V-S3-21(constraint があるのに taboo 0件)/ 正例1本
'   計 12本
'
' 末尾から modTestsPure16.RunAll(W7・17章 T-56 の16本)を呼ぶ。
'
' 素材は 15章§8.1 の mock 応答(MK-S1-NEW / MK-S1-RNW / MK-S2-NEW / MK-S2-RNW /
'   MK-S3)であり、狙った1ケースだけを Replace で壊して**発火の差**を見る
'   (ChkFire。素材のままでは出ず、壊すと出ることの2点を同時に固定する)。
'
' ID実在検査の一覧を渡さない呼び出しは V-S2-06 / V-S3-03..06 を fail-closed で
'   出すが、ChkFire は**当該ケースIDの有無だけ**を比べるので判定に影響しない。
'
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

Private Const MK_S1N As String = "MK-S1-NEW"
Private Const MK_S1R As String = "MK-S1-RNW"
Private Const MK_S2N As String = "MK-S2-NEW"
Private Const MK_S2R As String = "MK-S2-RNW"
Private Const MK_S3 As String = "MK-S3"
Private Const CT_NEW As String = "new"
Private Const CT_RNW As String = "renewal"

Public Sub RunAll()
    On Error GoTo FA
    T_W7A_CheckS1
WB:
    On Error GoTo FB
    T_W7B_CheckS2
WC:
    On Error GoTo FC
    T_W7C_CheckS3
WD:
    On Error GoTo FD
    modTestsPure16.RunAll
WDone:
    Exit Sub
FA:
    GroupFail "W7A CheckS1(v2.6)"
    Resume WB
FB:
    GroupFail "W7B CheckS2(v2.6)"
    Resume WC
FC:
    GroupFail "W7C CheckS3(v2.6)"
    Resume WD
FD:
    GroupFail "modTestsPure16.RunAll"
    Resume WDone
End Sub

Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

' ============================================================================
' W7A CheckS1(15章§11。V-S1-12 / V-S1-13 と V-S1-04 の改訂)
' ============================================================================
Private Sub T_W7A_CheckS1()
    Dim s1n As String
    Dim okNew As String

    s1n = MockJson(MK_S1N)
    okNew = modValidate.CheckS1(s1n, CT_NEW)

    ChkFire "V-S1-12_current_coverageのcertaintyがenum外_15章§11", _
        modValidate.CheckS1(Rep(s1n, """certainty"":""assumed""", _
                                """certainty"":""maybe"""), CT_NEW), okNew

    ChkFire "V-S1-13_financialsのsourceがenum外_15章§11", _
        modValidate.CheckS1(Rep(s1n, """source"":""unknown""", _
                                """source"":""hearsay"""), CT_NEW), okNew

    ' 裁定書25 S1 の要: 新規案件でも【付保の見立て】由来の current_coverage を
    '   出してよく、全件 assumed なら V-S1-04 は発火しない(警告0件)。
    ChkNoId "CheckS1_新規で現契約が全件assumedなら確認済み警告は出ない_15章§11", _
        okNew, "V-S1-04"
End Sub

' ============================================================================
' W7B CheckS2(15章§11。V-S2-12b / V-S2-18 と V-S2-16 の閾値5)
' ============================================================================
Private Sub T_W7B_CheckS2()
    Dim s2n As String
    Dim s2r As String
    Dim s1r As String
    Dim okNew As String

    s2n = MockJson(MK_S2N)
    s2r = MockJson(MK_S2R)
    s1r = MockJson(MK_S1R)
    okNew = modValidate.CheckS2(s2n, CT_NEW)

    ChkFire "V-S2-12b_新規案件でgap_typeがuninsured以外_15章§11", _
        modValidate.CheckS2(Rep(s2n, """gap_type"":""uninsured""", _
                                """gap_type"":""underinsured"""), CT_NEW), okNew

    ' 旧 V-S2-12 の撤回(裁定書25 S1)。新規案件で gaps が2件あっても不合格に
    '   ならず、gap_type が全件 uninsured なら V-S2-12b も出ない。
    ChkNoId "CheckS2_新規でgap_typeが全件uninsuredなら不合格にならない_15章§11", _
        okNew, "V-S2-12b"

    ' MK-S1-RNW は net_assets=12億円。MK-S2-NEW は loss_scale_note が全件空。
    ChkFire "V-S2-18_純資産既知でloss_scale_noteが全件空_15章§11", _
        modValidate.CheckS2(s2n, CT_NEW, "", "", s1r), _
        modValidate.CheckS2(s2n, CT_NEW)

    ' MK-S2-RNW は risk_no=3 に「純資産12億円に対し…(概算)」の対比を持つので
    '   純資産が既知でも発火しない(15章§3 の補足: 全件空のときだけ警告)。
    ChkNoId "CheckS2_対比が1件でもあれば純資産既知でも警告は出ない_15章§11", _
        modValidate.CheckS2(s2r, CT_RNW, "", "", s1r), "V-S2-18"

    ' 裁定書25 S7: 上限は5件。5件ちょうどは合格(6件超過の負例は modTestsPure5)。
    ChkNoId "CheckS2_emerging_risksが5件なら上限超過にならない_15章§11", _
        modValidate.CheckS2(EmergingOnlyJson(5), CT_NEW), "V-S2-16"
End Sub

' ============================================================================
' W7C CheckS3(15章§11。V-S3-19 / V-S3-20 / V-S3-21)
' ============================================================================
Private Sub T_W7C_CheckS3()
    Dim s3 As String
    Dim s2n As String
    Dim s1sum As String
    Dim okBase As String

    s3 = MockJson(MK_S3)
    s2n = MockJson(MK_S2N)
    s1sum = ConstraintSummaryJson()
    okBase = modValidate.CheckS3(s3, s2n, "", "", "", "", CT_NEW, s1sum)

    ' flow を2件へ削る(3件未満)。
    ChkFire "V-S3-19_talk_scriptのflowが3件未満_15章§11", _
        modValidate.CheckS3(Rep(s3, """flow"":[", """flow"":[""1文目のみ""],""dummy_flow"":["), _
                            s2n, "", "", "", "", CT_NEW, s1sum), okBase

    ' opening を空文字にする(1字未満)。
    ChkFire "V-S3-20_talk_scriptのopeningが空_15章§11", _
        modValidate.CheckS3(Rep(s3, """opening"":""御社が掲げるEC直販の拡大について、いま一番気になっている足元のリスクからお伺いできますか。""", _
                                """opening"":"""""), s2n, "", "", "", "", CT_NEW, s1sum), okBase

    ' field_insights に constraint があるのに taboo を空配列にする。
    ChkFire "V-S3-21_constraintがあるのにtabooが0件_15章§11", _
        modValidate.CheckS3(Rep(s3, """taboo"":[""先代からの工場設備の更新を急がせる言い方""]", _
                                """taboo"":[]"), s2n, "", "", "", "", CT_NEW, s1sum), okBase

    ' 素材のままでは flow は4件なので上限・下限のどちらにも当たらない
    '   (15章§8.1 受入条件1。opening・taboo の正例は上の3本の okBase 側が兼ねる)。
    ChkNoId "CheckS3_MK-S3のtalk_scriptはflow4件で件数の不合格にならない_15章§11", _
        okBase, "V-S3-19"
End Sub

' ============================================================================
' 判定ヘルパ(modTestsPure5 と同じ作法。各 modTestsPure* が自前で持つ)
' ============================================================================

' 「壊すと出る・素材のままでは出ない」の2点を1本で固定する。ケースIDは
'   テスト名の先頭(最初の "_" まで)から取り出す。
Private Sub ChkFire(ByVal nm As String, ByVal ngOut As String, ByVal okOut As String)
    Dim cid As String

    cid = CaseIdOf(nm)
    modTestRunner.Check nm, _
        (HasCase(ngOut, cid) And Not HasCase(okOut, cid)), _
        "壊した素材の結果=[" & HeadOf(ngOut) & "] 素材のままの結果=[" & HeadOf(okOut) & "]"
End Sub

' 正例(この文脈では出てはならない)。テスト名にケースIDを含めないので
'   照合対象のIDは引数で渡す(17章§4-2 は1ケース1テストを求めるため)。
Private Sub ChkNoId(ByVal nm As String, ByVal outText As String, ByVal idText As String)
    modTestRunner.Check nm, Not HasCase(outText, idText), _
        "出てはならない " & idText & " が出ています。実際=[" & HeadOf(outText) & "]"
End Sub

Private Function CaseIdOf(ByVal nm As String) As String
    Dim p As Long

    p = InStr(nm, "_")
    If p <= 1 Then
        CaseIdOf = nm
    Else
        CaseIdOf = Left$(nm, p - 1)
    End If
End Function

' 戻り値(vbLf区切り)のどれかの行が "[cid" で始まるか(15章§0 原則10)。
Private Function HasCase(ByVal outText As String, ByVal cid As String) As Boolean
    HasCase = (InStr(outText, "[" & cid) > 0)
End Function

Private Function HeadOf(ByVal s As String) As String
    Dim t As String

    t = Replace(Replace(s, vbCrLf, " / "), vbLf, " / ")
    If Len(t) > 160 Then t = Left$(t, 160) & "..."
    HeadOf = t
End Function

' ============================================================================
' 素材ヘルパ(純文字列)
' ============================================================================

Private Function MockJson(ByVal mockId As String) As String
    MockJson = modJsonLite.ExtractJsonBlock(modMockLlm.ResponseById(mockId))
End Function

' 最初の1件だけでなく全件を置換する(狙ったキーは素材中で一意か、全件同じ
'   壊し方をしたい箇所にだけ使う)。
Private Function Rep(ByVal js As String, ByVal fromText As String, _
                     ByVal toText As String) As String
    Rep = Replace(js, fromText, toText)
End Function

' emerging_risks を n 件だけ持つ最小のS2 JSON(V-S2-16 の境界確認用)。
'   risks が0件なので V-S2-01 等は出るが、見るのは V-S2-16 の有無だけ。
Private Function EmergingOnlyJson(ByVal n As Long) As String
    Dim s As String
    Dim i As Long

    s = "{""risks"":[],""gaps"":[],""emerging_risks"":["
    For i = 1 To n
        If i > 1 Then s = s & ","
        s = s & "{""risk_name"":""新興リスク" & CStr(i) & """,""category"":""digital_info"""
        s = s & ",""horizon"":""near"",""scenario"":""この企業への当てはまりの説明"""
        s = s & ",""evidence_quote"":""引用"",""evidence_source"":""hp"",""proposal_hint"":""""}"
    Next i
    EmergingOnlyJson = s & "],""open_questions"":[]}"
End Function

' V-S3-21 用の {{s1SummaryJson}}(15章§4 の4キーのうち field_insights だけを
'   持てば足りる。tag=constraint が1件あることが taboo の根拠になる)。
Private Function ConstraintSummaryJson() As String
    ConstraintSummaryJson = "{""field_insights"":[{""note"":""先代からの工場を大事にしており設備更新には慎重"",""tag"":""constraint""}]}"
End Function
