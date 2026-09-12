Attribute VB_Name = "modTestsPure27"
Option Explicit

' ============================================================================
' modTestsPure27 - W15(裁定書38 班C)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**裁定書38 §1 班C の文と 15章§5.6 / 20章だけ**から手で
'   書き出した(17章§1。実装の出力を見てから期待値を合わせない)。
'
' 対象と根拠(**全61本**。G1=14 / G2=5 / G3=4 / G4=3 / G5=16 / G6=19。
'   本数は `modTestRunner.Check` と `ChkFires` の呼び出し数の実測であり、
'   wintest/tests_expected.txt の prod と必ず同時に直すこと。裁定書40 S-m で
'   「全14本」「全26本」という申告が実数と食い違っていたのを直した):
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
'   G5 W15 Round 2(裁定書39 R2-02 / R2-11 / R2-12 / R2-04受け。全16本)
'      26..33 R2-02 最長一致・一般語・冪等・英字の語境界(8本)
'      34     R2-11 機械置換の件数を呼出側が受け取る
'      35..40 R2-12 描画例外の印と ?debug=1 の赤枠・S5必須キー2本と
'             R2-04受け 空白類だけの確認者名を認めない2本(6本)
'   G6 W15 Round 2 の検収是正(裁定書40。全19本)
'      41..42 Q-M1 会社名のTABで表紙のフィールドがずれない(両方向)
'      43..45 S-M1 残った社内語を呼出側へ知らせる(出る/出ない)と警告文の値源
'      46..49 S-M2 サ変語幹は「〜する」の直前だけ置換しない(名詞の位置は置換)
'      50..54 S-m  TabooHitStrict の両方向・語境界(BIG/OTC/SLAB)・連鎖する対
'      55..57 S-M4 置換後の CheckS5 が通る/通らない素材の見分け
'      58..59 S-m  確認者名の NBSP・VT・FF(両方向)
'
' 変異注入(出来レース禁止・裁定書38 §2):
'   (a) modExportProposal.NeedsReviewMessage を常に "" にすると 19 が落ちる。
'   (b) modProposalHtml1.SlidesJs から登録行を1本消すと 14 が落ち、
'       tools/render_proposal.py も同時に赤くなる。
'   (c) modExportProposal の Soft() を素通しにすると 17 が落ちる。
'   (d) modValidate4.SoftPairs の並べ替えを昇順にすると 26..28 が落ちる。
'   (e) modExportProposal.Soft() が件数を捨てると 34 が落ちる。
'   (f) modExportProposal.StripFieldSeps の vbTab 除去をやめると 41 が落ちる
'       (裁定書40 Q-M1。レポート側の Test_W15_25/26 と同じ型の網)。
'   (g) modValidate4 の suru 印を1行でも外すと 46 が落ちる(S-M2)。
'   (h) modValidate4.SuruFollows を常に False にすると 46 が落ち、常に True に
'       すると 47(名詞の位置では置換する)が落ちる。
'   (i) BuildProposalDataEx の tabooLeft を空のままにすると 43 が落ちる(S-M1)。
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
    W15Round2Soften
    If Err.Number <> 0 Then W15Fail "W15-G5 Round2 機械置換"

    Err.Clear
    W15Round2Render
    If Err.Number <> 0 Then W15Fail "W15-G5 Round2 描画と確認者名"

    Err.Clear
    W15R40Cover
    If Err.Number <> 0 Then W15Fail "W15-G6 R40 表紙のフィールド"

    Err.Clear
    W15R40Left
    If Err.Number <> 0 Then W15Fail "W15-G6 R40 残った社内語"

    Err.Clear
    W15R40Suru
    If Err.Number <> 0 Then W15Fail "W15-G6 R40 サ変語幹"

    Err.Clear
    W15R40Strict
    If Err.Number <> 0 Then W15Fail "W15-G6 R40 担当語と語境界"

    Err.Clear
    W15R40Accept
    If Err.Number <> 0 Then W15Fail "W15-G6 R40 機械置換後の受け入れ"

    Err.Clear
    W15R40Reviewer
    If Err.Number <> 0 Then W15Fail "W15-G6 R40 確認者名の空判定"

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

    ' 「移転」は一般語なので単独では置換しない(裁定書39 R2-02)。ここは
    ' 一般語でない語で「置換して件数を返す」ことだけを見る。
    after = modValidate4.SoftenTaboo("リスクユニバースを整理します。", changed)
    ok = (changed = 1)
    If ok Then ok = (InStr(1, after, "リスクの全体像", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 SoftenTaboo は社内語を顧客語へ置換して件数を返す", ok, _
        "changed=" & CStr(changed) & " after=" & after

    hit = modValidate4.TabooHit("IoT機器の導入を進めます。")
    modTestRunner.Check "W15 英字の禁止語は語境界で見る(IoT の中の OT を拾わない)", _
        (LenB(hit) = 0), hit
End Sub

' ============================================================================
' G5-1 W15 Round 2(裁定書39 R2-02 / R2-11)。対訳表の機械置換。
'   期待値は**壊す班 R2 が実測した壊れ方**(R2_break.md §2 R2-02 の S2..S7)を
'   そのまま裏返して書いた。実装の出力を見てから合わせていない(17章§1)。
' ============================================================================
Private Sub W15Round2Soften()
    Dim changed As Long
    Dim changed2 As Long
    Dim after As String
    Dim after2 As String
    Dim hit As String
    Dim ok As Boolean
    Dim srcText As String

    ' R2-02 S2: 「未付保」が先に「付保」に食われて「未保険のご加入」になった。
    after = modValidate4.SoftenTaboo("未付保の拠点があります。", changed)
    ok = (InStr(1, after, "保険に入っていない状態", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, after, "未保険のご加入", vbBinaryCompare) = 0)
    modTestRunner.Check "W15 R2-02 未付保は最長一致で置換する(未保険のご加入を作らない)", _
        ok, "after=" & after

    ' R2-02 S3: 「付保ギャップ」が「保険のご加入ギャップ」になった。
    after = modValidate4.SoftenTaboo("付保ギャップが残ります。", changed)
    ok = (InStr(1, after, "保険で手当てできていない部分", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, after, "保険のご加入ギャップ", vbBinaryCompare) = 0)
    modTestRunner.Check "W15 R2-02 付保ギャップは最長一致で置換する", ok, "after=" & after

    ' R2-02 S4: 「リスク移転可能性」が「リスク保険で備える可能性」になった。
    after = modValidate4.SoftenTaboo("リスク移転可能性を評価します。", changed)
    ok = (InStr(1, after, "保険での備えやすさ", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, after, "リスク保険で備える可能性", vbBinaryCompare) = 0)
    modTestRunner.Check "W15 R2-02 リスク移転可能性は最長一致で置換する", ok, "after=" & after

    ' R2-02 S5: 「座組パターン」が「ご提案の構成パターン」になった。
    after = modValidate4.SoftenTaboo("座組パターンをお示しします。", changed)
    ok = (InStr(1, after, "ご提案の型", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, after, "ご提案の構成パターン", vbBinaryCompare) = 0)
    modTestRunner.Check "W15 R2-02 座組パターンは最長一致で置換する", ok, "after=" & after

    ' R2-02 S6: 一般語の「移転」(本社移転)まで潰して「本社を保険で備えるする」。
    srcText = "本社を移転する計画があります。"
    after = modValidate4.SoftenTaboo(srcText, changed)
    ok = (after = srcText)
    If ok Then ok = (changed = 0)
    modTestRunner.Check "W15 R2-02 一般語の移転は置換しない(本社移転を壊さない)", ok, _
        "changed=" & CStr(changed) & " after=" & after

    ' R2-02 S7: 一般語の「保有」(現金を保有)まで潰して「自社で負担する」。
    srcText = "現金を保有しています。"
    after = modValidate4.SoftenTaboo(srcText, changed)
    ok = (after = srcText)
    If ok Then ok = (changed = 0)
    modTestRunner.Check "W15 R2-02 一般語の保有は置換しない", ok, _
        "changed=" & CStr(changed) & " after=" & after

    ' 置換しないだけで見逃しはしない(裁定書39 R2-02「警告のみ」)。テスト名に
    ' ケースIDを書かないこと: tools/validate_check.py は「1ケース1テスト」を
    ' テスト名の文字列リテラルから数えるため、ここで書くと2本目と数えられる。
    hit = modValidate4.TabooHit("本社を移転する計画があります。")
    ok = (InStr(1, hit, "移転", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R2-02 一般語も禁止語の一覧(TabooHit)には出る", ok, "hit=" & hit

    ' 冪等: 置換後の文字列をもう一度通しても変わらない(裁定書39 R2-02)。
    after = modValidate4.SoftenTaboo("サブリミットと未付保の状況です。", changed)
    after2 = modValidate4.SoftenTaboo(after, changed2)
    ok = (after2 = after)
    If ok Then ok = (changed2 = 0)
    If ok Then ok = (InStr(1, after, "未保険のご加入", vbBinaryCompare) = 0)
    modTestRunner.Check "W15 R2-02 SoftenTaboo は冪等(2回目は1件も置換しない)", ok, _
        "1回目=" & after & " / 2回目changed=" & CStr(changed2)

    ' 英字の禁止語は語境界で見る。CBI を BI で刻まない。
    after = modValidate4.SoftenTaboo("CBI と BI を区別します。", changed)
    ok = (InStr(1, after, "取引先の被災による損害", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, after, "事業が止まったことによる利益の減少", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, after, "C事業が止まった", vbBinaryCompare) = 0)
    modTestRunner.Check "W15 R2-02 英字の禁止語は語境界で置換する(CBI を BI で刻まない)", _
        ok, "after=" & after

    W15Round2Count
End Sub

' R2-11: 機械置換の件数を呼出側が受け取れること(usage_log の値源)。
Private Sub W15Round2Count()
    Dim meta As String
    Dim s2 As String
    Dim s5 As String
    Dim s2Dirty As String
    Dim dataJson As String
    Dim softClean As Long
    Dim softDirty As Long
    Dim leftClean As String
    Dim leftDirty As String
    Dim ok As Boolean

    s2 = modMockLlm3.BuildS2RnwJson()
    s5 = modMockLlm4.BuildS5Json()
    meta = modExportProposal.BuildProposalMetaJson(W15_COMPANY, s5, _
        "2026/09/12 10:00:00", "2.4.0", W15_REVIEWER, "2026/09/12 10:05:00")

    dataJson = modExportProposal.BuildProposalDataEx(meta, s2, s5, softClean, leftClean)
    ' S2 の control_note 全件の先頭へ社内語を差し込む(置換が必ず起きる素材)。
    s2Dirty = Replace(s2, """control_note"":""", """control_note"":""リスクユニバースと")
    dataJson = modExportProposal.BuildProposalDataEx(meta, s2Dirty, s5, softDirty, leftDirty)

    ok = (softDirty > 0)
    If ok Then ok = (softDirty > softClean)
    If ok Then ok = (InStr(1, dataJson, "リスクユニバース", vbBinaryCompare) = 0)
    ' 機械置換が責任を持つ語は1語も残らない(残るのは印のある語だけ)。
    If ok Then ok = (LenB(modValidate4.TabooHitStrict(leftClean)) = 0)
    If ok Then ok = (LenB(modValidate4.TabooHitStrict(leftDirty)) = 0)
    modTestRunner.Check "W15 R2-11 提案書DATAの機械置換の件数を呼出側が受け取る", ok, _
        "clean=" & CStr(softClean) & " dirty=" & CStr(softDirty)
End Sub

' ============================================================================
' G5-2 W15 Round 2(裁定書39 R2-12 と R2-04 の受け)。描画の失敗を隠さない。
' ============================================================================
Private Sub W15Round2Render()
    Dim js As String
    Dim css As String
    Dim ok As Boolean
    Dim miss As String
    Dim msgTab As String
    Dim msgWide As String

    ' R2-12: 描画例外は「次回更新します」に化けるだけでなく印を残す。
    js = modProposalHtml1.RuntimeJs()
    ok = (InStr(1, js, "catch(err){AT(sec,'data-render-error','1');", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R2-12 描画例外は data-render-error を立てる", ok, _
        "長さ=" & CStr(Len(js))

    ' R2-12: 赤枠は ?debug=1 のときだけ(お客さまの画面には出さない)。
    css = modProposalHtml2.FormatCss()
    ok = (InStr(1, css, "body.debug", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, css, "data-render-error", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, js, "debug=1", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R2-12 描画エラーの赤枠は ?debug=1 のときだけ出る", ok, _
        "css長=" & CStr(Len(css))

    ' R2-12: S5 の必須キーが欠けたら描かずに止める(保留文に化けさせない)。
    ok = (LenB(modExportProposal.MissingProposalKeys(modMockLlm4.BuildS5Json())) = 0)
    modTestRunner.Check "W15 R2-12 mock(MK-S5)には必須キーが全部ある", ok, _
        "実際=" & modExportProposal.MissingProposalKeys(modMockLlm4.BuildS5Json())

    miss = modExportProposal.MissingProposalKeys( _
        Replace(modMockLlm4.BuildS5Json(), """steps"":", """steps_off"":"))
    ok = (InStr(1, miss, "steps", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R2-12 S5の必須キーが欠けていたら欠落キーを返す", ok, _
        "実際=[" & miss & "]"

    ' R2-04 の受け: 空白類だけの確認者名を「確認済み」と認めない。
    msgTab = modExportProposal.NeedsReviewMessage(vbTab)
    modTestRunner.Check "W15 R2-04 確認者名がTABだけなら提案書を生成しない", _
        (msgTab = W15_NEED), "実際=[" & msgTab & "]"

    msgWide = modExportProposal.NeedsReviewMessage(ChrW(12288))
    modTestRunner.Check "W15 R2-04 確認者名が全角空白だけなら提案書を生成しない", _
        (msgWide = W15_NEED), "実際=[" & msgWide & "]"
End Sub

' ============================================================================
' G6 W15 Round 2 の検収是正(裁定書40)。期待値は裁定表と検証者レポートの
'   再現手順から手で書いた(17章§1)。テスト名にケースIDを書かないこと
'   (validate_check.py が「1ケース1テスト」を名前から数えるため)。
' ============================================================================

' G6-1 Q-M1: coverFields(vbTab区切り)へ入れる前に区切りを落とす。再現手順は
'   検証者レポート「会社名を `甲斐<TAB>A<TAB>B<TAB>山田` にして提案書を出す」。
'   落とさないと <title> が「ご提案 甲斐」で切れ、<noscript> の題が会社名の
'   後半(攻撃者が決めた文字列)に化ける。
Private Sub W15R40Cover()
    Dim s2 As String
    Dim s5 As String
    Dim realTitle As String
    Dim metaTab As String
    Dim metaClean As String
    Dim docTab As String
    Dim docClean As String
    Dim company As String
    Dim ok As Boolean

    s2 = modMockLlm3.BuildS2RnwJson()
    s5 = modMockLlm4.BuildS5Json()
    realTitle = modJsonLite.GetStr(modJsonLite.ExtractJsonBlock(s5), "title")
    company = "甲斐" & vbTab & "A" & vbTab & "B" & vbTab & "C" & vbTab & "山田"

    metaTab = modExportProposal.BuildProposalMetaJson(company, s5, _
        "2026/09/12 10:00:00", "2.4.0", "", "2026/09/12 10:05:00")
    docTab = modExportProposal.BuildProposalHtml(metaTab, s2, s5)

    ' 題(FieldAt(...,1))が会社名の後半に化けていないこと。
    ok = (InStr(1, docTab, "<p>" & realTitle & "</p>", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, docTab, "<p>A</p>", vbBinaryCompare) = 0)
    If ok Then ok = (InStr(1, docTab, "<title>ご提案 甲斐ABC山田</title>", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R40 会社名にTABを混ぜても表紙のフィールドがずれない", ok, _
        "題=[" & realTitle & "]"

    ' 逆向き: TAB が無ければ会社名も題もそのまま出る(検査が常に真でないこと)。
    metaClean = modExportProposal.BuildProposalMetaJson("甲斐ABC山田", s5, _
        "2026/09/12 10:00:00", "2.4.0", "", "2026/09/12 10:05:00")
    docClean = modExportProposal.BuildProposalHtml(metaClean, s2, s5)
    ok = (InStr(1, docClean, "<title>ご提案 甲斐ABC山田</title>", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, docClean, "<p>" & realTitle & "</p>", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R40 TABの無い会社名はそのまま表紙に出る", ok, _
        "長さ=" & CStr(Len(docClean))
End Sub

' G6-2 S-M1: 置換しない語(一般語・サ変語幹)が顧客向け本文に残ったら、
'   呼出側へ**必ず知らせる**(usage_log と警告文の値源)。
Private Sub W15R40Left()
    Dim s2 As String
    Dim s5 As String
    Dim meta As String
    Dim s2Dirty As String
    Dim s2Clean As String
    Dim dummy As String
    Dim soft As Long
    Dim leftText As String
    Dim warnLine As String
    Dim ok As Boolean

    s2 = modMockLlm3.BuildS2RnwJson()
    s5 = modMockLlm4.BuildS5Json()
    meta = modExportProposal.BuildProposalMetaJson(W15_COMPANY, s5, _
        "2026/09/12 10:00:00", "2.4.0", W15_REVIEWER, "2026/09/12 10:05:00")

    ' 一般語「移転」は機械置換しない(日本語が壊れるため)。残るなら知らせる。
    s2Dirty = Replace(s2, """control_note"":""", """control_note"":""本社を移転する案もあり、")
    dummy = modExportProposal.BuildProposalDataEx(meta, s2Dirty, s5, soft, leftText)
    ok = (InStr(1, leftText, "移転", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R40 一般語が顧客向け本文に残ったら呼出側へ知らせる", ok, _
        "left=[" & leftText & "] soft=" & CStr(soft)

    ' 逆向き: 社内語を1語も含まない本文なら知らせない(常に警告が出る実装で
    '   ないこと)。mock の S2 は control_note に一般語「移転」を持つので
    '   (検証者が dist/サンプル提案書.html で実測した箇所)、そこだけを
    '   顧客語へ直した素材で見る。
    s2Clean = Replace(s2, "移転", "お引き受け")
    dummy = modExportProposal.BuildProposalDataEx(meta, s2Clean, s5, soft, leftText)
    ok = (LenB(leftText) = 0)
    modTestRunner.Check "W15 R40 社内語が残っていなければ知らせない", ok, _
        "left=[" & leftText & "]"

    ' 警告文の値源は modValidate4 の1本(不合格の行と同じ文言を使う)。
    warnLine = modValidate4.TabooWarnLine("移転;保有")
    ok = (warnLine = "[V-S5-12] 顧客向けに書き換えていない語があります: 移転;保有")
    If ok Then ok = (LenB(modValidate4.TabooWarnLine("")) = 0)
    modTestRunner.Check "W15 R40 残った社内語の警告文は語の一覧が非空のときだけ組む", _
        ok, "line=[" & warnLine & "]"
End Sub

' G6-3 S-M2: サ変語幹(「〜する」に続けて使う社内語)は、その位置では置換
'   しない。検証者が実測した「保険化する → 保険での備え方の設計する」と同型の
'   壊れ方を全語で禁じる。名詞として使われている位置では従来どおり置換する。
Private Sub W15R40Suru()
    Dim words() As String
    Dim i As Long
    Dim changed As Long
    Dim after As String
    Dim srcText As String
    Dim ng As String
    Dim ok As Boolean

    words = Split("付保|保険化|特約開発|組成|引受|ヒアリング|攻めの保険活用|" & _
                  "顕在化|与信|クロスセル", "|")
    For i = LBound(words) To UBound(words)
        srcText = words(i) & "する予定です。"
        after = modValidate4.SoftenTaboo(srcText, changed)
        If after <> srcText Then ng = ng & "[" & after & "]"
        srcText = words(i) & "しました。"
        after = modValidate4.SoftenTaboo(srcText, changed)
        If after <> srcText Then ng = ng & "[" & after & "]"
        srcText = words(i) & "されています。"
        after = modValidate4.SoftenTaboo(srcText, changed)
        If after <> srcText Then ng = ng & "[" & after & "]"
    Next i
    modTestRunner.Check "W15 R40 サ変語幹は する/し/さ の直前では置換しない", _
        (LenB(ng) = 0), "壊れた出力=" & ng

    ' 逆向き(1): 名詞として使われている位置は置換する(見逃しにしない)。
    after = modValidate4.SoftenTaboo("保険化の検討と付保の状況。", changed)
    ok = (InStr(1, after, "保険での備え方の設計の検討", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, after, "保険のご加入の状況", vbBinaryCompare) > 0)
    If ok Then ok = (changed = 2)
    modTestRunner.Check "W15 R40 サ変語幹でも名詞の位置では顧客語へ置換する", ok, _
        "after=" & after & " changed=" & CStr(changed)

    ' 逆向き(2): 顧客語がサ変名詞の対まで外していない(過剰な印を禁じる)。
    after = modValidate4.SoftenTaboo("リスクを仕分けする。", changed)
    ok = (InStr(1, after, "整理する", vbBinaryCompare) > 0)
    If ok Then ok = (changed = 1)
    modTestRunner.Check "W15 R40 顧客語がサ変名詞の対は する でも置換する", ok, _
        "after=" & after & " changed=" & CStr(changed)

    ' サ変語幹も禁止語の一覧(警告)には出る=置換しないことと見逃すことは別。
    ok = (InStr(1, modValidate4.TabooHit("保険化する予定です。"), "保険化", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R40 サ変語幹も禁止語の一覧には出る", ok, _
        "hit=" & modValidate4.TabooHit("保険化する予定です。")
End Sub

' G6-4 S-m: TabooHitStrict(機械置換が責任を持つ語だけ)の両方向。
'   この関数は S5 の受け入れ判断と提案書の記録で使うのに、テストが1本も
'   無かった(裁定書40 S-m)。
Private Sub W15R40Strict()
    Dim strictHit As String
    Dim allHit As String
    Dim ok As Boolean
    Dim ch As Long
    Dim after As String

    strictHit = modValidate4.TabooHitStrict("リスクユニバースを整理します。")
    ok = (InStr(1, strictHit, "リスクユニバース", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R40 機械置換の担当語は TabooHitStrict に出る", ok, _
        "strict=[" & strictHit & "]"

    strictHit = modValidate4.TabooHitStrict("本社を移転する。保険化する。")
    allHit = modValidate4.TabooHit("本社を移転する。保険化する。")
    ok = (LenB(strictHit) = 0)
    If ok Then ok = (InStr(1, allHit, "移転", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, allHit, "保険化", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R40 印のある語は TabooHitStrict に出ない(警告には出る)", _
        ok, "strict=[" & strictHit & "] all=[" & allHit & "]"

    ' G6-5 S-m: 語境界の検査(BoundaryOk)を**実際に通る**素材。最長一致だけでは
    '   通らない組み合わせを選ぶ(BIG / OTC / SLAB は対訳表に無い語)。
    after = modValidate4.SoftenTaboo("BIGデータとOTC医薬品とSLABの話。", ch)
    ok = (after = "BIGデータとOTC医薬品とSLABの話。")
    If ok Then ok = (ch = 0)
    modTestRunner.Check "W15 R40 英字の禁止語は語の一部に当たらない(BIG/OTC/SLAB)", _
        ok, "after=" & after & " ch=" & CStr(ch)

    after = modValidate4.SoftenTaboo("BI と OT の話。", ch)
    ok = (InStr(1, after, "事業が止まったことによる利益の減少", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, after, "工場の制御システム", vbBinaryCompare) > 0)
    If ok Then ok = (ch = 2)
    modTestRunner.Check "W15 R40 単独の英字の禁止語は置換する(語境界の逆向き)", _
        ok, "after=" & after & " ch=" & CStr(ch)

    ' G6-6 S-m: 連鎖する対の**実効出力**を固定する。宣言された顧客語
    '   「補償項目ごとの支払限度額」自体が社内語「支払限度額」を含むため、
    '   2回目の走査で言い換わる。宣言側の是正は15章§5.6 と同時にしか
    '   できない(handoff)ので、黙って漂流しないよう出力を固定する。
    after = modValidate4.SoftenTaboo("サブリミットの設定。", ch)
    ok = (after = "補償項目ごとのお支払いの上限額の設定。")
    modTestRunner.Check "W15 R40 連鎖する対の実効出力を固定する(サブリミット)", ok, _
        "after=" & after & " ch=" & CStr(ch)
End Sub

' G6-7 S-M4: 機械置換したあとの受け入れ判断の材料。**置換しても CheckS5 が
'   1件でも発火する素材は採用してはいけない**(前波はここが緩み、検証に
'   不合格の S5 が顧客向け提案書の材料になっていた)。判断そのもの
'   (modPipeline5.SoftenedOrEmpty)は PURE_ALLOWLIST に modPipeline5 が
'   無いため層(a)から呼べないので、判断が使う2つの材料(置換件数・置換後の
'   CheckS5)を両方向で固定する(allowlist へ1行=司令塔へ handoff)。
Private Sub W15R40Accept()
    Dim s2 As String
    Dim s5 As String
    Dim raw As String
    Dim softened As String
    Dim recheck As String
    Dim changed As Long
    Dim longTail As String
    Dim ok As Boolean

    s2 = modMockLlm3.BuildS2RnwJson()
    s5 = modMockLlm4.BuildS5Json()

    ' (1) 機械置換で直る素材: 置換が効き、置換後の検証に1件も発火しない。
    raw = Replace(s5, """subtitle"": """, """subtitle"": ""リスクユニバースを含む")
    softened = modValidate4.SoftenTaboo(modJsonLite.ExtractJsonBlock(raw), changed)
    recheck = modValidate4.CheckS5(softened, s2)
    ok = (changed > 0)
    If ok Then ok = (LenB(recheck) = 0)
    modTestRunner.Check "W15 R40 機械置換で直る素材は置換後の検証に通る", ok, _
        "changed=" & CStr(changed) & " recheck=" & Left$(recheck, 120)

    ' (2) 置換そのものが別の不合格を生む素材(見出しが置換で60字を超える)。
    '     置換は効くが検証は通らない=採用してはいけない最も危ない型。
    longTail = String$(52, "A")
    raw = Replace(s5, """riskmap"": """, """riskmap"": ""BI " & longTail & " ")
    softened = modValidate4.SoftenTaboo(modJsonLite.ExtractJsonBlock(raw), changed)
    recheck = modValidate4.CheckS5(softened, s2)
    ok = (changed > 0)
    If ok Then ok = (InStr(1, recheck, "[V-S5-02]", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R40 置換が別の不合格を生む素材は置換後の検証で落ちる", ok, _
        "changed=" & CStr(changed) & " recheck=" & Left$(recheck, 120)

    ' (3) 一般語だけが残った素材: 置換は1件も効かず、検証も通らないまま。
    raw = Replace(s5, """subtitle"": """, """subtitle"": ""本社を移転する")
    softened = modValidate4.SoftenTaboo(modJsonLite.ExtractJsonBlock(raw), changed)
    recheck = modValidate4.CheckS5(softened, s2)
    ok = (changed = 0)
    If ok Then ok = (InStr(1, recheck, "移転", vbBinaryCompare) > 0)
    modTestRunner.Check "W15 R40 一般語だけが残った素材は置換が効かず検証も通らない", ok, _
        "changed=" & CStr(changed) & " recheck=" & Left$(recheck, 120)
End Sub

' G6-8 S-m: 確認者名の空判定に NBSP(U+00A0)・VT(11)・FF(12)を足す
'   (Word やブラウザからの貼り付けで混入し、見た目が空の確認者名で
'    顧客提示物が「確認済み」になっていた)。
Private Sub W15R40Reviewer()
    Dim ng As String
    Dim msg As String
    Dim ok As Boolean

    msg = modExportProposal.NeedsReviewMessage(ChrW(160))
    If msg <> W15_NEED Then ng = ng & "NBSP "
    msg = modExportProposal.NeedsReviewMessage(Chr(11))
    If msg <> W15_NEED Then ng = ng & "VT "
    msg = modExportProposal.NeedsReviewMessage(Chr(12))
    If msg <> W15_NEED Then ng = ng & "FF "
    msg = modExportProposal.NeedsReviewMessage(ChrW(160) & Chr(11) & Chr(12) & vbTab)
    If msg <> W15_NEED Then ng = ng & "混在 "
    modTestRunner.Check "W15 R40 見えない空白だけの確認者名を認めない(NBSP/VT/FF)", _
        (LenB(ng) = 0), "通した文字=" & ng

    ' 逆向き: 見えない空白に実名が混じっていれば生成へ進む。
    ok = (LenB(modExportProposal.NeedsReviewMessage(ChrW(160) & "山田" & Chr(11))) = 0)
    modTestRunner.Check "W15 R40 見えない空白に実名が混じれば生成へ進む", ok, _
        "実際=[" & modExportProposal.NeedsReviewMessage(ChrW(160) & "山田" & Chr(11)) & "]"
End Sub
