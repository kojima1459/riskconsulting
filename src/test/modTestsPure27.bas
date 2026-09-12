Attribute VB_Name = "modTestsPure27"
Option Explicit

' ============================================================================
' modTestsPure27 - W15(裁定書38 班C)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**裁定書38 §1 班C の文と 15章§5.6 / 20章だけ**から手で
'   書き出した(17章§1。実装の出力を見てから期待値を合わせない)。
'
' 対象と根拠(全26本):
'   G1 CheckS5(15章§5.6 の検証表。**1ケース1本**。tools/validate_check.py が
'      「§11の全ケースIDに対しテストが1本ずつ」を機械で見る)
'      00 mock(MK-S5)は1件も発火しない  01..13 V-S5-01 から V-S5-13
'   G2 提案書テンプレとDATA(20章)
'      14 スライド登録表は22枚で no は 1 から 22
'      15 提案書JSONが空ならDATAを組まない(空データで落ちない)
'      16 内部の値がDATAに入らない(20章§3)
'      17 S2 由来の社内語はDATAへ入る前に機械置換される
'      18 免責は20章§8の1文
'   G3 確認必須とファイル名(裁定書38 §1 班C 2・3)
'      19 reviewedBy 空なら案内文  20 非空なら空文字(生成へ進む)
'      21 提案書のファイル名規則  22 レポートも同じ規則(先頭語だけ違う)
'   G4 対訳表(docs/design/提案書_wide/対訳表_社内語から顧客語.md)
'      23 30語以上ある  24 SoftenTaboo は置換して件数を返す
'      25 英字の禁止語は語境界で見る(IoT の中の OT を拾わない)
'
' 変異注入(出来レース禁止・裁定書38 §2):
'   (a) modExportProposal.NeedsReviewMessage を常に "" にすると 19 が落ちる。
'   (b) modProposalHtml1.SlidesJs から登録行を1本消すと 14 が落ち、
'       tools/render_proposal.py も同時に赤くなる。
'   (c) modExportProposal の Soft() を素通しにすると 17 が落ちる。
'
' 書き方の約束(LibreOffice Basic 対策): Dim はプロシージャの先頭にまとめ、
'   判定は一度ローカル変数へ入れてから modTestRunner.Check へ渡す
'   (引数の位置で関数呼び出しや複数行の論理式を組み立てない)。グループの
'   呼び分けも Select Case を介さず RunAll から直接呼ぶ(他の modTestsPure* と
'   同名の private プロシージャを作らない=同一ライブラリ内での取り違えを防ぐ)。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ============================================================================

Private Const W15_COMPANY As String = "株式会社浜松スイーツファクトリー"
Private Const W15_REVIEWER As String = "浜松支店 山田"
Private Const W15_NEED As String = "内容を確認してから出力してください。"
Private Const W15_DISC As String = "本資料は、引受・保険料・契約条件を確約するものではありません。"

Public Sub RunAll()
    On Error Resume Next

    Err.Clear
    W15CheckS5
    If Err.Number <> 0 Then W15Fail "W15-G1 CheckS5"

    Err.Clear
    W15Template
    If Err.Number <> 0 Then W15Fail "W15-G2 提案書テンプレとDATA"

    Err.Clear
    W15ReviewAndName
    If Err.Number <> 0 Then W15Fail "W15-G3 確認必須とファイル名"

    Err.Clear
    W15Glossary
    If Err.Number <> 0 Then W15Fail "W15-G4 対訳表"

    Err.Clear
    On Error GoTo 0
End Sub

' グループ全体の失敗を1件のテスト失敗として見せる(無かったことにしない)。
Private Sub W15Fail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Err.Clear
End Sub

' ============================================================================
' G1 CheckS5(15章§5.6)。素材は mock(MK-S5 / MK-S2-RNW)で、**1箇所だけ**を
'   壊して当該ケースが発火することを見る(壊した所以外が道連れで落ちない)。
' ============================================================================
Private Sub W15CheckS5()
    Dim baseJson As String
    Dim s2 As String
    Dim r As String
    Dim t As String

    baseJson = modMockLlm4.BuildS5Json()
    s2 = modMockLlm3.BuildS2RnwJson()

    r = modValidate4.CheckS5(baseJson, s2)
    modTestRunner.Check "W15 mock(MK-S5)は CheckS5 に1件も発火しない", (LenB(r) = 0), r

    t = W15KeyOff(baseJson, "themes")
    ChkFires "W15 V-S5-01 が発火する", "V-S5-01", t, s2

    t = Replace(baseJson, """riskmap""", """riskmap_off""")
    ChkFires "W15 V-S5-02 が発火する", "V-S5-02", t, s2

    t = Replace(baseJson, """risk_no"": 2,", """risk_no"": 99,")
    ChkFires "W15 V-S5-03 が発火する", "V-S5-03", t, s2

    t = W15KeyOff(baseJson, "hard_risks")
    ChkFires "W15 V-S5-04 が発火する", "V-S5-04", t, s2

    t = W15KeyOff(baseJson, "ideas")
    ChkFires "W15 V-S5-05 が発火する", "V-S5-05", t, s2

    t = Replace(baseJson, """effect"": 5", """effect"": 9")
    ChkFires "W15 V-S5-06 が発火する", "V-S5-06", t, s2

    t = W15KeyOff(baseJson, "four")
    ChkFires "W15 V-S5-07 が発火する", "V-S5-07", t, s2

    t = W15KeyOff(baseJson, "steps")
    ChkFires "W15 V-S5-08 が発火する", "V-S5-08", t, s2

    t = W15KeyOff(baseJson, "decisions")
    ChkFires "W15 V-S5-09 が発火する", "V-S5-09", t, s2

    t = W15KeyOff(baseJson, "notes")
    ChkFires "W15 V-S5-10 が発火する", "V-S5-10", t, s2

    t = W15KeyOff(baseJson, "areas")
    ChkFires "W15 V-S5-11 が発火する", "V-S5-11", t, s2

    t = Replace(baseJson, """subtitle"": """, """subtitle"": ""移転を含む")
    ChkFires "W15 V-S5-12 が発火する", "V-S5-12", t, s2

    t = Replace(baseJson, """priority"": true", """priority"": false")
    ChkFires "W15 V-S5-13 が発火する", "V-S5-13", t, s2
End Sub

' 当該ケースIDの行が出ることを見る。
'   テスト名は**呼び出し側にリテラルで書く**(tools/validate_check.py は
'   `Check "..."` / `Chk* "..."` の第1引数の文字列リテラルからテスト名を拾うため、
'   ここで連結して作るとケースIDが機械から見えなくなる)。名前にケースIDを
'   1つだけ含めるのが「1テスト1ケース」規約。
Private Sub ChkFires(ByVal testName As String, ByVal caseId As String, _
                     ByVal jsonText As String, ByVal s2 As String)
    Dim r As String
    Dim ok As Boolean

    r = modValidate4.CheckS5(jsonText, s2)
    ok = (InStr(1, r, "[" & caseId & "]", vbBinaryCompare) > 0)
    modTestRunner.Check testName, ok, "実際=" & Left$(r, 160)
End Sub

' 配列キーの名前を潰して「0件」を作る(件数系ケースの最小の壊し方)。
Private Function W15KeyOff(ByVal jsonText As String, ByVal keyName As String) As String
    Dim fromText As String
    Dim toText As String

    fromText = """" & keyName & """: ["
    toText = """" & keyName & "_off"": ["
    W15KeyOff = Replace(jsonText, fromText, toText)
End Function

' ============================================================================
' G2 提案書テンプレとDATA(20章)
' ============================================================================
Private Sub W15Template()
    Dim reg As String
    Dim i As Long
    Dim okAll As Boolean
    Dim missing As String
    Dim meta As String
    Dim s2 As String
    Dim s5 As String
    Dim s2Dirty As String
    Dim dataJson As String
    Dim emptyData As String
    Dim dirtyData As String
    Dim taboo As String
    Dim ok As Boolean
    Dim disc As String

    reg = modProposalHtml1.SlidesJs()
    okAll = True
    For i = 1 To 22
        If InStr(1, reg, "{no:" & CStr(i) & ",slug:'", vbBinaryCompare) = 0 Then
            okAll = False
            missing = missing & CStr(i) & " "
        End If
    Next i
    If InStr(1, reg, "{no:23,", vbBinaryCompare) > 0 Then
        okAll = False
        missing = missing & "23枚目がある "
    End If
    modTestRunner.Check "W15 スライド登録表は22枚で no が1から22まで揃う", okAll, _
        "欠け=" & missing

    s2 = modMockLlm3.BuildS2RnwJson()
    s5 = modMockLlm4.BuildS5Json()
    meta = modExportProposal.BuildProposalMetaJson(W15_COMPANY, s5, _
        "2026/09/12 10:00:00", "2.4.0", W15_REVIEWER, "2026/09/12 10:05:00")

    emptyData = modExportProposal.BuildProposalData(meta, s2, "")
    modTestRunner.Check "W15 提案書JSONが空ならDATAを組まない(空データで落ちない)", _
        (LenB(emptyData) = 0), "実際=" & Left$(emptyData, 80)

    dataJson = modExportProposal.BuildProposalData(meta, s2, s5)
    ok = (LenB(dataJson) > 0)
    If ok Then ok = (InStr(1, dataJson, "dossier_tier", vbBinaryCompare) = 0)
    If ok Then ok = (InStr(1, dataJson, "quality_mode", vbBinaryCompare) = 0)
    If ok Then ok = (InStr(1, dataJson, "case_type", vbBinaryCompare) = 0)
    If ok Then ok = (InStr(1, dataJson, "round_no", vbBinaryCompare) = 0)
    If ok Then ok = (InStr(1, dataJson, "talk_script", vbBinaryCompare) = 0)
    If ok Then ok = (InStr(1, dataJson, "field_insights", vbBinaryCompare) = 0)
    modTestRunner.Check "W15 内部の値がDATAに入らない(20章§3)", ok, _
        "長さ=" & CStr(Len(dataJson))

    ' S2 の control_note には社内語が入り得る。DATAへ入る前に対訳表で機械置換
    ' していれば、左列の語はDATAに1つも残らない。
    taboo = "リスクユニバース"
    s2Dirty = Replace(s2, """control_note"":""", """control_note"":""" & taboo & "と")
    dirtyData = modExportProposal.BuildProposalData(meta, s2Dirty, s5)
    ok = (InStr(1, dirtyData, taboo, vbBinaryCompare) = 0)
    modTestRunner.Check "W15 S2由来の社内語はDATAへ入る前に機械置換される", ok, _
        "長さ=" & CStr(Len(dirtyData))

    disc = modProposalHtml4.DisclaimerText()
    modTestRunner.Check "W15 免責は20章§8の1文", (disc = W15_DISC), disc
End Sub

' ============================================================================
' G3 確認必須とファイル名(裁定書38 §1 班C 2・3)
' ============================================================================
Private Sub W15ReviewAndName()
    Dim msgEmpty As String
    Dim msgNamed As String
    Dim nameP As String
    Dim nameR As String

    msgEmpty = modExportProposal.NeedsReviewMessage("")
    modTestRunner.Check "W15 確認者名が空なら提案書を生成しない", _
        (msgEmpty = W15_NEED), msgEmpty

    msgNamed = modExportProposal.NeedsReviewMessage(W15_REVIEWER)
    modTestRunner.Check "W15 確認者名があれば生成へ進む", (LenB(msgNamed) = 0), msgNamed

    nameP = modUtilPath.BuildVersionedFileName("提案書", "株式会社ABC商事", _
        "20260912", "2.4.0", "", ".html")
    modTestRunner.Check "W15 提案書のファイル名は 提案書_会社名_日付_v版", _
        (nameP = "提案書_株式会社ABC商事_20260912_v2.4.0"), nameP

    nameR = modUtilPath.BuildVersionedFileName("レポート", "株式会社ABC商事", _
        "20260912", "2.4.0", "", ".html")
    modTestRunner.Check "W15 レポートのファイル名も同じ規則(先頭語だけ違う)", _
        (nameR = "レポート_株式会社ABC商事_20260912_v2.4.0"), nameR
End Sub

' ============================================================================
' G4 対訳表(docs/design/提案書_wide/対訳表_社内語から顧客語.md)
' ============================================================================
Private Sub W15Glossary()
    Dim rows() As String
    Dim n As Long
    Dim changed As Long
    Dim after As String
    Dim hit As String
    Dim ok As Boolean

    rows = Split(modValidate4.TabooPairs(), vbLf)
    n = UBound(rows) - LBound(rows) + 1
    modTestRunner.Check "W15 対訳表は30語以上ある(裁定書38 班C)", (n >= 30), _
        "件数=" & CStr(n)

    after = modValidate4.SoftenTaboo("リスクの移転を検討します。", changed)
    ok = (changed = 1)
    If ok Then ok = (InStr(1, after, "保険で備える", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 SoftenTaboo は社内語を顧客語へ置換して件数を返す", ok, _
        "changed=" & CStr(changed) & " after=" & after

    hit = modValidate4.TabooHit("IoT機器の導入を進めます。")
    modTestRunner.Check "W15 英字の禁止語は語境界で見る(IoT の中の OT を拾わない)", _
        (LenB(hit) = 0), hit
End Sub
