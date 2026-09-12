Attribute VB_Name = "modValidate4"
Option Explicit

' ============================================================================
' modValidate4 - 顧客向け提案書(S5)の後検証(15章§5.6 CheckS5 検証ルール表)
' ----------------------------------------------------------------------------
' 12章§2 の30,000字契約による modValidate の分割先(1と2は満杯、3は班A が
'   S1 の証拠検証で使うため4を使う)。
'
' 持つもの:
'   CheckS5        15章§5.6 の13件(V-S5-01からV-S5-13)。""=合格
'   TabooPairs     対訳表(docs/design/提案書_wide/対訳表_社内語から顧客語.md)の
'                  「社内語<TAB>顧客語<TAB>mode」を vbLf で並べた**唯一の値源**
'                  (mode は全行にある。replace / warn の2区分だけ)
'   SoftenTaboo    修復後も V-S5-12 だけが残るときの機械置換(docs/29 §5.3。
'                  生成を止めない。置換件数を呼出側へ返す)
'   TabooHit       本文に残っている禁止語(";"区切り。0件なら "")
'   TabooHitStrict 同上のうち**置換の取りこぼしだけ**(mode=replace の語が、
'                  終端集合の文脈に残っているとき。=実装の欠陥)
'   TabooWarnLine  V-S5-12 の1行を組み立てる唯一の値源(不合格にも警告にも使う)
'   MissingTopKeys JSONの最外オブジェクト直下に無いキー(";"区切り)
'
' 機械置換の規約(**裁定書43 §1 で設計を変えた**。正は対訳表§6〜§6.6):
'   (1) **最長一致**(長さの降順に1回走査。宣言順の Replace をやめる)。
'   (2) **置換してよい文脈を閉じた集合で定義する**(対訳表§6)。直後の1文字が
'       終端集合(TermChars)に入るか社内語が末尾のときだけ置換する。終端でない
'       位置は**その位置の走査を打ち切る**(より短い語も当てない)。旧設計は
'       「見送る語尾」を白名簿で列挙したため直すたびに同じ型が残った(整理ている
'       →整理ながら/保険保険→額額)。活用語尾は開いた集合で列挙では閉じない。
'   (3) **mode=warn の対は一切置換しない**(警告 TabooHit には出す。旧
'       general / suru / verb はここへ集約した)。
'   (4) **冪等**(変化が無くなるまで V4_SOFT_PASS_MAX 回まで通す。顧客語が別の
'       社内語を含む対があるため1回走査では冪等にならない)。
'   (5) **取り消し規則**(UndoNeeded。対訳表§6.5 の二重の安全網)。
'
' R4準拠(12章§2): Excelトークン・Application.Run・案件データ参照を持たない
'   純関数モジュール。CP932準拠(15章§0 原則7)。
' ============================================================================

Private Const V4_TAB As String = vbTab
Private Const V4_SEP As String = ";"
' 第3列 mode(対訳表§6.1。**2区分だけ**。印を増やさない)。
'   replace = 終端集合の文脈でだけ置換する / warn = 一切置換せず警告だけ
'   (一般語・顧客語が述語・動作性名詞で顧客語が静的名詞句の対)。
Private Const V4_REPLACE As String = "replace"
Private Const V4_WARN As String = "warn"
' 終端集合(対訳表§6 の宣言と1文字ずつ対応。**これだけ。増やさない**。
'   増やしたくなった対は mode=warn にする)。
Private Const V4_TERM_JOSHI As String = "をにはがのへとでもやかねよら"
Private Const V4_TERM_MARK1 As String = "、。，．・…「」『』()（）【】[]〈〉《》"
Private Const V4_TERM_MARK2 As String = ":：;；/／|｜-－―~～!！?？""'"
' 半角空白と全角空白(タブ・CR・LF は TermChars が足す)。
Private Const V4_TERM_SPACE As String = " 　"
' 取り消し規則(対訳表§6.5)。同一文字の連続の上限と、顧客語の末尾を見る字数。
Private Const V4_RUN_MAX As Long = 3
Private Const V4_UNDO_TAIL_MAX As Long = 3
' 冪等化の走査回数の上限(対訳表の連鎖は最長でも2段。無限ループの歯止め)。
Private Const V4_SOFT_PASS_MAX As Long = 4
' V-S5-02: 見出しの上限。15章§5.6 system は「45文字程度・60文字を超えない」。
Private Const V4_HEADLINE_MAX As Long = 60
' V-S5-02 が見る headline の11キー(15章§5.6 Schema-S5 と同順)。
Private Const V4_HEADLINE_KEYS As String = _
    "riskmap|classes|priority|hard|ideas|four|themes|steps|decide|share|appendix"
' 件数の固定値(15章§5.6 の検証表)。
Private Const V4_THEMES_N As Long = 3
Private Const V4_STEPS_N As Long = 4
Private Const V4_DECISIONS_N As Long = 3
Private Const V4_NOTES_N As Long = 22
Private Const V4_BIZ_N As Long = 5
Private Const V4_HARD_MIN As Long = 1
Private Const V4_HARD_MAX As Long = 6
Private Const V4_IDEA_MIN As Long = 1
Private Const V4_IDEA_MAX As Long = 8
Private Const V4_FOUR_MIN As Long = 1
Private Const V4_FOUR_MAX As Long = 4
Private Const V4_EFFECT_MIN As Long = 1
Private Const V4_EFFECT_MAX As Long = 5
' 一覧未提供の fail-closed 文(裁定書7 A-2。15章§11と同じ趣旨)。
Private Const V4_NOLIST As String = "ID実在検査が実行できません(ID一覧未提供)"

' ============================================================================
' CheckS5 - 15章§5.6 の検証。""=合格 / 非空=エラー行(vbLf区切り・行頭[ケースID])
'   jsonText : S5の出力JSON(ExtractJsonBlock 済み)
'   s2Json   : risk_no の実在検査に使うS2のJSON。**非空の risk_no を持つのに
'              s2Json が空**なら V-S5-03 で不合格(fail-closed)
' ============================================================================
Public Function CheckS5(ByVal jsonText As String, ByVal s2Json As String) As String
    Dim r As String

    ChkCounts r, jsonText
    ChkHeadline r, jsonText
    ChkHardRisks r, jsonText, s2Json
    ChkIdeas r, jsonText
    ChkShare r, jsonText

    ' --- V-S5-12: 対訳表の社内語が本文に残っていないか ---
    Dim hit As String
    hit = TabooHit(jsonText)
    If LenB(hit) > 0 Then Ap r, TabooWarnLine(hit)

    CheckS5 = r
End Function

' V-S5-01 / V-S5-04 / V-S5-05 / V-S5-07 / V-S5-08 / V-S5-09 / V-S5-10 / V-S5-11
Private Sub ChkCounts(ByRef r As String, ByVal jsonText As String)
    Dim n As Long

    n = ArrCount(jsonText, "themes")
    If n <> V4_THEMES_N Then Ap r, "[V-S5-01] themes が" & n & "件です(3件固定)"

    n = ArrCount(jsonText, "hard_risks")
    If n < V4_HARD_MIN Or n > V4_HARD_MAX Then
        Ap r, "[V-S5-04] hard_risks が" & n & "件です(1から6件)"
    End If

    n = ArrCount(jsonText, "ideas")
    If n < V4_IDEA_MIN Or n > V4_IDEA_MAX Then
        Ap r, "[V-S5-05] ideas が" & n & "件です(1から8件)"
    End If

    n = ArrCount(jsonText, "four")
    If n < V4_FOUR_MIN Or n > V4_FOUR_MAX Then
        Ap r, "[V-S5-07] four が" & n & "件です(1から4件)"
    End If

    n = ArrCount(jsonText, "steps")
    If n <> V4_STEPS_N Then Ap r, "[V-S5-08] steps が" & n & "件です(4件固定)"

    n = ArrCount(jsonText, "decisions")
    If n <> V4_DECISIONS_N Then Ap r, "[V-S5-09] decisions が" & n & "件です(3件固定)"

    n = ArrCount(jsonText, "notes")
    If n <> V4_NOTES_N Then Ap r, "[V-S5-10] notes が" & n & "件です(22件固定)"

    Dim bizText As String
    bizText = ObjBlock(jsonText, "business")
    n = ArrCount(bizText, "areas")
    If n <> V4_BIZ_N Then Ap r, "[V-S5-11] business.areas が" & n & "件です(5件固定)"
    n = ArrCount(bizText, "factors")
    If n <> V4_BIZ_N Then Ap r, "[V-S5-11] business.factors が" & n & "件です(5件固定)"
End Sub

' V-S5-02: headline の11キーが空でなく60字以内か。headline は入れ子の
'   オブジェクトであり、themes / ideas / steps / four / share は上位にも同名の
'   キーがあるため、**必ず ObjBlock で headline の中だけを切り出してから**見る
'   (modJsonLite.GetStr は最初に現れたキーを拾うため)。
Private Sub ChkHeadline(ByRef r As String, ByVal jsonText As String)
    Dim hd As String
    Dim keys() As String
    Dim i As Long
    Dim sVal As String

    hd = ObjBlock(jsonText, "headline")
    keys = Split(V4_HEADLINE_KEYS, "|")
    For i = LBound(keys) To UBound(keys)
        sVal = Trim$(modJsonLite.GetStr(hd, keys(i)))
        If LenB(sVal) = 0 Or Len(sVal) > V4_HEADLINE_MAX Then
            Ap r, "[V-S5-02] headline." & keys(i) & " が" & Len(sVal) & _
                  "字です(空にせず60字以内)"
        End If
    Next i
End Sub

' V-S5-03: hard_risks[].risk_no がS2に実在するか(一覧未提供時も不合格)。
Private Sub ChkHardRisks(ByRef r As String, ByVal jsonText As String, ByVal s2Json As String)
    Dim it As Variant
    Dim noText As String
    Dim nosText As String

    nosText = RiskNoList(s2Json)
    For Each it In ArrItems(jsonText, "hard_risks")
        noText = Trim$(modJsonLite.GetStr(CStr(it), "risk_no"))
        If LenB(noText) > 0 Then
            If LenB(Trim$(s2Json)) = 0 Then
                Ap r, "[V-S5-03] " & V4_NOLIST
            ElseIf InStr(1, nosText, V4_SEP & noText & V4_SEP, vbBinaryCompare) = 0 Then
                Ap r, "[V-S5-03] hard_risks の risk_no " & noText & " がS2に実在しません"
            End If
        End If
    Next it
End Sub

' V-S5-06: ideas[].effect が1から5の整数か。
Private Sub ChkIdeas(ByRef r As String, ByVal jsonText As String)
    Dim it As Variant
    Dim raw As String
    Dim n As Long

    For Each it In ArrItems(jsonText, "ideas")
        raw = Trim$(modJsonLite.GetStr(CStr(it), "effect"))
        n = modJsonLite.GetLong(CStr(it), "effect", 0)
        If n < V4_EFFECT_MIN Or n > V4_EFFECT_MAX Or Not IsPositiveInt(raw) Then
            Ap r, "[V-S5-06] ideas の effect が" & raw & "です(1から5の整数)"
        End If
    Next it
End Sub

' V-S5-13(警告): share_items の優先の印が1件も無い。
Private Sub ChkShare(ByRef r As String, ByVal jsonText As String)
    Dim it As Variant
    Dim n As Long
    Dim anyPri As Boolean

    For Each it In ArrItems(jsonText, "share_items")
        n = n + 1
        If modJsonLite.GetBoolJ(CStr(it), "priority", False) Then anyPri = True
    Next it
    If n > 0 And Not anyPri Then
        Ap r, "[V-S5-13] share_items に優先の印がありません(4件程度に印を付けてください)"
    End If
End Sub

' TabooPairs - 対訳表(docs/design/提案書_wide/対訳表_社内語から顧客語.md)を
'   「社内語<TAB>顧客語<TAB>mode」で vbLf 区切りに並べた1本。**禁止語の唯一の
'   値源**。§1〜§3 の46語 + §4.1 の3語 + §6.4 の表記ゆれ2語 = 51行(後の2つは
'   番号付きの表に置かない=15章§5.6 の対訳行は46語のまま)。mode は**全行に
'   ある**。全列一致は tools/render_proposal.py の check_glossary_impl が見る。
Public Function TabooPairs() As String
    Dim s As String
    AdPair s, "付保", "保険のご加入", V4_REPLACE
    AdPair s, "未付保", "保険に入っていない状態", V4_REPLACE
    AdPair s, "付保ギャップ", "保険で手当てできていない部分", V4_REPLACE
    AdPair s, "未充足", "保険の手当てが無い", V4_WARN
    AdPair s, "移転", "保険で備える", V4_WARN
    AdPair s, "保有", "自社で負担する", V4_WARN
    AdPair s, "トリガー", "保険金をお支払いする条件", V4_REPLACE
    AdPair s, "サブリミット", "補償項目ごとの支払限度額", V4_REPLACE
    AdPair s, "待機期間", "補償が始まるまでの期間", V4_REPLACE
    AdPair s, "保険化", "保険での備え方の設計", V4_REPLACE
    AdPair s, "特約開発", "補償内容の新しい設計", V4_REPLACE
    AdPair s, "組成", "仕組みづくり", V4_REPLACE
    AdPair s, "募集スキーム", "ご加入の手続きの流れ", V4_REPLACE
    AdPair s, "料率", "保険料の水準", V4_REPLACE
    AdPair s, "相関損失", "同時に起きる損害", V4_REPLACE
    AdPair s, "引受", "保険のお引き受け", V4_REPLACE
    ' 表記ゆれ(対訳表§6.4)。用言の連用形なので warn(「引受けられる」)。
    AdPair s, "引受け", "保険のお引き受け", V4_WARN
    AdPair s, "過少保険", "補償額が損害に届かない状態", V4_REPLACE
    AdPair s, "抜け", "補償されない部分", V4_WARN
    AdPair s, "免責金額", "ご負担いただく金額", V4_REPLACE
    AdPair s, "支払限度額", "お支払いの上限額", V4_REPLACE
    AdPair s, "リスクユニバース", "リスクの全体像", V4_REPLACE
    AdPair s, "ニューリスク", "新しく生まれているリスク", V4_REPLACE
    AdPair s, "座組", "ご提案の構成", V4_REPLACE
    AdPair s, "座組み", "ご提案の構成", V4_REPLACE
    AdPair s, "ヒアリング", "お伺いしたい事項", V4_WARN
    AdPair s, "提案の核", "ご提案の前提", V4_REPLACE
    AdPair s, "攻めの保険活用", "成長を後押しする保険の活用", V4_REPLACE
    AdPair s, "発散段階", "構想段階", V4_REPLACE
    AdPair s, "実装難度", "実現までの難易度", V4_REPLACE
    AdPair s, "顕在化", "実際に起きること", V4_WARN
    AdPair s, "打ち手", "対策", V4_REPLACE
    AdPair s, "商材", "保険商品", V4_REPLACE
    AdPair s, "リスク移転可能性", "保険での備えやすさ", V4_REPLACE
    AdPair s, "与信", "取引先の支払い能力", V4_WARN
    AdPair s, "座組パターン", "ご提案の型", V4_REPLACE
    AdPair s, "PML", "想定最大損害額", V4_REPLACE
    AdPair s, "CBI", "取引先の被災による損害", V4_REPLACE
    AdPair s, "BI", "事業が止まったことによる利益の減少", V4_REPLACE
    AdPair s, "RTO", "復旧までの目標時間", V4_REPLACE
    AdPair s, "BCP", "事業継続計画", V4_REPLACE
    AdPair s, "OT", "工場の制御システム", V4_REPLACE
    AdPair s, "MFA", "多要素認証", V4_REPLACE
    AdPair s, "EDR", "端末の不審な動きを検知する仕組み", V4_REPLACE
    AdPair s, "KRI", "リスクの予兆指標", V4_REPLACE
    AdPair s, "SLA", "サービス水準の取り決め", V4_REPLACE
    AdPair s, "D&O", "会社役員賠償責任保険", V4_REPLACE
    AdPair s, "PL保険", "生産物賠償責任保険", V4_REPLACE
    AdPair s, "対話の順序", "ご説明の順序", V4_REPLACE
    AdPair s, "クロスセル", "追加でご検討いただける備え", V4_WARN
    AdPair s, "仕分け", "整理", V4_WARN
    TabooPairs = s
End Function

' TabooPairs の1行を積む(社内語<TAB>顧客語<TAB>mode。行は vbLf 区切り)。
Private Sub AdPair(ByRef acc As String, ByVal w As String, ByVal c As String, _
                   ByVal md As String)
    If LenB(acc) > 0 Then acc = acc & vbLf
    acc = acc & w & V4_TAB & c & V4_TAB & md
End Sub

' TabooHit - 本文に残っている禁止語を ";" 区切りで返す(0件なら "")。
'   半角英字だけの語(PML/BI/OT 等)は前後が英字のときに当てない(「IoT」の中の
'   OT を数えない)。**mode=warn の行もここには出す**。V-S5-12 は「書き換えて
'   いない語がある」ことの通知であり、置換するかどうかとは別の判断である。
Public Function TabooHit(ByVal bodyText As String) As String
    TabooHit = HitList(bodyText, True)
End Function

' TabooHitStrict - 上記のうち**置換の取りこぼしだけ**(対訳表§6.1 の末尾)=
'   mode=replace の語が**終端集合の文脈**(置換してよい位置)に残っているとき。
'   SoftenTaboo の後にこれが非空なら実装の欠陥である。warn の語と、終端でない
'   文脈に残った replace の語は欠陥ではないので出さない(TabooHit には出る)。
'   呼出側(modExportProposal.TabooLeftNote)が両者を分けて記録する。
Public Function TabooHitStrict(ByVal bodyText As String) As String
    TabooHitStrict = HitList(bodyText, False)
End Function

' TabooWarnLine - V-S5-12 の1行を組み立てる**唯一の値源**(15章§5.6 の
'   エラー文テンプレ)。CheckS5 は不合格の行として、提案書の出力経路
'   (modExportProposal)は警告の注記として同じ1行を使う(裁定書40 S-M1)。
Public Function TabooWarnLine(ByVal hitsText As String) As String
    If LenB(hitsText) = 0 Then Exit Function
    TabooWarnLine = "[V-S5-12] 顧客向けに書き換えていない語があります: " & hitsText
End Function

' SoftenTaboo - 修復後も V-S5-12 だけが残るときの機械置換(docs/29 §5.3)。
'   置換した**箇所数**を changed へ返す。**生成は止めない**が、置換したことは
'   呼出側が run_log と警告へ残す(黙って直さない)。規約は冒頭の(1)〜(4)。
Public Function SoftenTaboo(ByVal bodyText As String, ByRef changed As Long) As String
    Dim srcArr() As String
    Dim dstArr() As String
    Dim heads As String
    Dim termText As String
    Dim cnt As Long
    Dim pass As Long
    Dim hits As Long
    Dim t As String

    changed = 0
    t = bodyText
    If LenB(t) = 0 Then
        SoftenTaboo = t
        Exit Function
    End If

    cnt = SoftPairs(srcArr, dstArr, heads)
    If cnt = 0 Then
        SoftenTaboo = t
        Exit Function
    End If

    termText = TermChars()
    For pass = 1 To V4_SOFT_PASS_MAX
        hits = 0
        t = SoftenOnce(t, srcArr, dstArr, cnt, heads, termText, hits)
        changed = changed + hits
        If hits = 0 Then Exit For
    Next pass
    SoftenTaboo = t
End Function

' ============================================================================
' MissingTopKeys - jsonText の**最外オブジェクト直下**に無いキーを ";" 区切りで
'   返す(全部あれば "")。keyList は "|" 区切り。入れ子の同名キーを「あった」と
'   数えない(themes[].headline を headline と読まない)ので深さを数えて走る。
' ============================================================================
Public Function MissingTopKeys(ByVal jsonText As String, ByVal keyList As String) As String
    Dim keys() As String
    Dim i As Long
    Dim acc As String

    If LenB(Trim$(keyList)) = 0 Then Exit Function
    keys = Split(keyList, "|")
    For i = LBound(keys) To UBound(keys)
        If LenB(keys(i)) > 0 Then
            If Not TopKeyAt(jsonText, keys(i)) Then
                If LenB(acc) > 0 Then acc = acc & V4_SEP
                acc = acc & keys(i)
            End If
        End If
    Next i
    MissingTopKeys = acc
End Function

' ============================================================================
' 内部(すべて純関数)
' ============================================================================

' エラー行を vbLf 区切りで積む(15章§0 原則10)。
Private Sub Ap(ByRef outText As String, ByVal lineText As String)
    If LenB(outText) = 0 Then
        outText = lineText
    Else
        outText = outText & vbLf & lineText
    End If
End Sub

' 禁止語の一覧。withWarn=True なら mode を問わず本文にあるもの全部(警告用)。
'   False なら**置換の取りこぼしだけ**=mode=replace の語が終端集合の文脈に
'   残っているもの(対訳表§6.1 の末尾。=実装の欠陥)。
Private Function HitList(ByVal bodyText As String, ByVal withWarn As Boolean) As String
    Dim rows() As String
    Dim onePair() As String
    Dim i As Long
    Dim word As String
    Dim acc As String
    Dim termText As String
    Dim found As Boolean

    If LenB(bodyText) = 0 Then Exit Function
    termText = TermChars()
    rows = Split(TabooPairs(), vbLf)
    For i = LBound(rows) To UBound(rows)
        ' Split() の戻り値へ直接添字を付けない(LibreOffice Basic が解さない)。
        onePair = Split(rows(i), V4_TAB)
        word = onePair(0)
        If LenB(word) > 0 Then
            If withWarn Then
                found = WordFound(bodyText, word)
            ElseIf IsWarnRow(onePair) Then
                found = False
            Else
                found = ReplaceableFound(bodyText, word, onePair(1), termText)
            End If
            If found Then
                If LenB(acc) > 0 Then acc = acc & V4_SEP
                acc = acc & word
            End If
        End If
    Next i
    HitList = acc
End Function

' 対訳表の1行の第3列(mode)。無ければ ""(全行にあるので通常は起きない)。
Private Function ModeOf(ByRef onePair() As String) As String
    If UBound(onePair) < 2 Then Exit Function
    ModeOf = onePair(2)
End Function

' 対訳表の1行が **mode=warn**(=一切置換しない行)か。対訳表§6.1。
'   warn の語も V-S5-12 の警告(TabooHit)には出す=見逃しはしない。
Private Function IsWarnRow(ByRef onePair() As String) As Boolean
    IsWarnRow = (ModeOf(onePair) = V4_WARN)
End Function

' 終端集合(対訳表§6 の宣言と1文字ずつ対応する**閉じた集合**)。em ダッシュ
'   U+2014 は CP932 に無くソースへ直接書けないので ChrW で足す(U+2015 は
'   MARK2 に直接入っている)。
Private Function TermChars() As String
    TermChars = V4_TERM_JOSHI & V4_TERM_MARK1 & V4_TERM_MARK2 & _
                V4_TERM_SPACE & vbTab & vbCr & vbLf & ChrW(&H2014)
End Function

' pos の文字が終端集合に入るか。pos が末尾の次(=社内語が文字列の末尾)なら True。
Private Function IsTermAt(ByVal hay As String, ByVal pos As Long, _
                          ByVal termText As String) As Boolean
    If pos > Len(hay) Then
        IsTermAt = True
        Exit Function
    End If
    If pos < 1 Then Exit Function
    IsTermAt = (InStr(1, termText, Mid$(hay, pos, 1), vbBinaryCompare) > 0)
End Function

' ReplaceOk - その位置で置換してよいか(対訳表§6 + §6.5)。**判定はここ1箇所**
'   (SoftenOnce と HitList の両方が使う)。pos = 社内語の直後の位置。
Private Function ReplaceOk(ByVal hay As String, ByVal pos As Long, _
                           ByVal dstText As String, ByVal termText As String, _
                           Optional ByVal startPos As Long = 0) As Boolean
    If Not IsTermAt(hay, pos, termText) Then Exit Function
    If UndoNeeded(dstText, hay, pos) Then Exit Function
    ' 語頭側の重なり(裁定書43 検証者・W15 司令塔の手直し)。終端集合は直後しか
    '   見ないので「保険付保の状況」→「保険保険のご加入…」が素通りしていた。
    If startPos > 1 Then
        If modValidate3.HeadOverlap(dstText, hay, startPos, V4_UNDO_TAIL_MAX) Then Exit Function
    End If
    ReplaceOk = True
End Function


' UndoNeeded - 取り消し規則(対訳表§6.5)。(a) 継ぎ目で同じ文字が V4_RUN_MAX 個
'   以上続く / (b) 顧客語の末尾が直後の本文と重複する なら True(=警告へ回す)。
Private Function UndoNeeded(ByVal dstText As String, ByVal hay As String, _
                            ByVal pos As Long) As Boolean
    Dim k As Long
    Dim seam As String

    For k = 1 To V4_UNDO_TAIL_MAX
        If k <= Len(dstText) Then
            If Mid$(hay, pos, k) = Right$(dstText, k) Then
                UndoNeeded = True
                Exit Function
            End If
        End If
    Next k
    seam = Right$(dstText, V4_RUN_MAX - 1) & Mid$(hay, pos, V4_RUN_MAX - 1)
    UndoNeeded = HasRun(seam, V4_RUN_MAX)
End Function

' 同じ文字が runMax 個以上続く箇所があるか。
Private Function HasRun(ByVal s As String, ByVal runMax As Long) As Boolean
    Dim i As Long
    Dim n As Long

    n = 1
    For i = 2 To Len(s)
        If Mid$(s, i, 1) = Mid$(s, i - 1, 1) Then
            n = n + 1
            If n >= runMax Then
                HasRun = True
                Exit Function
            End If
        Else
            n = 1
        End If
    Next i
End Function

' 本文に word が**置換してよい文脈で**現れるか(=置換の取りこぼしの判定)。
Private Function ReplaceableFound(ByVal hay As String, ByVal word As String, _
                                  ByVal dstText As String, _
                                  ByVal termText As String) As Boolean
    Dim p As Long

    If LenB(word) = 0 Then Exit Function
    p = InStr(1, hay, word, vbBinaryCompare)
    Do While p > 0
        If BoundaryOk(hay, p, word) Then
            If ReplaceOk(hay, p + Len(word), dstText, termText, p) Then
                ReplaceableFound = True
                Exit Function
            End If
        End If
        p = InStr(p + 1, hay, word, vbBinaryCompare)
    Loop
End Function

' SoftPairs - 機械置換に使う対を**長さの降順**(最長一致)で返す。戻り値=件数。
'   heads には社内語の1文字目を重複なく詰める(走査の足切り用)。**mode=warn の
'   行は入れない**。置換してよいかは位置で決まる(ReplaceOk)。
Private Function SoftPairs(ByRef srcArr() As String, ByRef dstArr() As String, _
                           ByRef heads As String) As Long
    Dim rows() As String
    Dim onePair() As String
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim keySrc As String
    Dim keyDst As String
    Dim headCh As String

    heads = vbNullString
    rows = Split(TabooPairs(), vbLf)
    ReDim srcArr(0 To UBound(rows) - LBound(rows))
    ReDim dstArr(0 To UBound(rows) - LBound(rows))
    n = 0
    For i = LBound(rows) To UBound(rows)
        onePair = Split(rows(i), V4_TAB)
        If UBound(onePair) >= 1 Then
            If LenB(onePair(0)) > 0 And Not IsWarnRow(onePair) Then
                srcArr(n) = onePair(0)
                dstArr(n) = onePair(1)
                n = n + 1
            End If
        End If
    Next i

    ' 長さの降順へ挿入整列(同じ長さなら対訳表の並びを保つ=安定)。
    For i = 1 To n - 1
        keySrc = srcArr(i)
        keyDst = dstArr(i)
        j = i - 1
        Do While j >= 0
            If Len(srcArr(j)) < Len(keySrc) Then
                srcArr(j + 1) = srcArr(j)
                dstArr(j + 1) = dstArr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        srcArr(j + 1) = keySrc
        dstArr(j + 1) = keyDst
    Next i

    For i = 0 To n - 1
        headCh = Left$(srcArr(i), 1)
        If InStr(1, heads, headCh, vbBinaryCompare) = 0 Then heads = heads & headCh
    Next i
    SoftPairs = n
End Function

' SoftenOnce - 本文を左から1回走査し、その位置で**最も長く一致する**社内語を
'   顧客語へ置き換える(置換箇所数を hits へ)。置換結果は走査済みとして扱い、
'   1文字目が heads に無い位置は読み飛ばし、連結は一致した所でだけ行う(長文の
'   速度)。**置換してよい文脈でない位置はその位置の走査を打ち切る**(規約2。
'   より短い語も当てない。これが無いと「引受けの方針」で「引受け」を見送った
'   直後に「引受」が当たり「保険のお引き受けけの方針」が出る)。
Private Function SoftenOnce(ByVal bodyText As String, ByRef srcArr() As String, _
                            ByRef dstArr() As String, _
                            ByVal cnt As Long, ByVal heads As String, _
                            ByVal termText As String, ByRef hits As Long) As String
    Dim outText As String
    Dim n As Long
    Dim i As Long
    Dim k As Long
    Dim segStart As Long
    Dim matched As Long
    Dim wLen As Long
    Dim ch As String

    hits = 0
    n = Len(bodyText)
    i = 1
    segStart = 1
    Do While i <= n
        matched = 0
        ch = Mid$(bodyText, i, 1)
        If InStr(1, heads, ch, vbBinaryCompare) > 0 Then
            For k = 0 To cnt - 1
                wLen = Len(srcArr(k))
                If wLen <= n - i + 1 Then
                    If Mid$(bodyText, i, wLen) = srcArr(k) Then
                        If BoundaryOk(bodyText, i, srcArr(k)) Then
                            ' 置換してよい文脈でなければ、この位置はここで
                            '   打ち切る(規約2)。
                            If Not ReplaceOk(bodyText, i + wLen, dstArr(k), termText, i) Then
                                Exit For
                            End If
                            matched = k + 1
                            Exit For
                        End If
                    End If
                End If
            Next k
        End If
        If matched > 0 Then
            If i > segStart Then
                outText = outText & Mid$(bodyText, segStart, i - segStart)
            End If
            outText = outText & dstArr(matched - 1)
            i = i + Len(srcArr(matched - 1))
            segStart = i
            hits = hits + 1
        Else
            i = i + 1
        End If
    Loop
    If segStart <= n Then outText = outText & Mid$(bodyText, segStart, n - segStart + 1)
    SoftenOnce = outText
End Function

' 置換してよい位置か。ASCII だけの語は前後が半角英字でないときだけ当てる
'   (「IoT」の中の OT を置換しない。全角を含む語は常に True)。
Private Function BoundaryOk(ByVal hay As String, ByVal pos As Long, _
                            ByVal word As String) As Boolean
    If Not IsAsciiWord(word) Then
        BoundaryOk = True
        Exit Function
    End If
    If IsAlphaAt(hay, pos - 1) Then Exit Function
    If IsAlphaAt(hay, pos + Len(word)) Then Exit Function
    BoundaryOk = True
End Function

' 禁止語の出現判定。ASCII だけの語は前後が英字でないときだけ当てる。
Private Function WordFound(ByVal hay As String, ByVal word As String) As Boolean
    Dim p As Long

    If LenB(word) = 0 Then Exit Function
    If Not IsAsciiWord(word) Then
        WordFound = (InStr(1, hay, word, vbBinaryCompare) > 0)
        Exit Function
    End If
    p = InStr(1, hay, word, vbBinaryCompare)
    Do While p > 0
        If Not IsAlphaAt(hay, p - 1) Then
            If Not IsAlphaAt(hay, p + Len(word)) Then
                WordFound = True
                Exit Function
            End If
        End If
        p = InStr(p + 1, hay, word, vbBinaryCompare)
    Loop
End Function

' 語が半角英字と記号だけでできているか(全角が1字でもあれば False)。
Private Function IsAsciiWord(ByVal word As String) As Boolean
    Dim i As Long
    Dim c As Long

    For i = 1 To Len(word)
        c = AscW(Mid$(word, i, 1))
        If c < 32 Or c > 126 Then Exit Function
    Next i
    IsAsciiWord = True
End Function

' hay の pos 文字目が半角英字か(範囲外は False)。
Private Function IsAlphaAt(ByVal hay As String, ByVal pos As Long) As Boolean
    Dim c As Long

    If pos < 1 Or pos > Len(hay) Then Exit Function
    c = AscW(Mid$(hay, pos, 1))
    IsAlphaAt = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

' 10進の正の整数か(V-S5-06。"3.5" や "三" を弾く)。
Private Function IsPositiveInt(ByVal t As String) As Boolean
    Dim i As Long
    Dim c As Long

    If LenB(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        c = AscW(Mid$(t, i, 1))
        If c < 48 Or c > 57 Then Exit Function
    Next i
    IsPositiveInt = True
End Function

' JSONの最外オブジェクト直下(深さ1)に当該キーがあるか(MissingTopKeys の実体)。
'   入れ子の同名キーを「あった」と数えないため、深さを数えながら走る。
Private Function TopKeyAt(ByVal jsonText As String, ByVal keyName As String) As Boolean
    Dim n As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim endPos As Long
    Dim rawKey As String

    n = Len(jsonText)
    i = 1
    Do While i <= n
        ch = Mid$(jsonText, i, 1)
        If ch = """" Then
            endPos = StrEnd(jsonText, i)
            If endPos = 0 Then Exit Function
            rawKey = Mid$(jsonText, i + 1, endPos - i - 1)
            i = endPos + 1
            Do While i <= n
                If InStr(" " & vbTab & vbCr & vbLf, Mid$(jsonText, i, 1)) = 0 Then Exit Do
                i = i + 1
            Loop
            If i <= n Then
                If Mid$(jsonText, i, 1) = ":" And depth = 1 And rawKey = keyName Then
                    TopKeyAt = True
                    Exit Function
                End If
            End If
        ElseIf ch = "{" Or ch = "[" Then
            depth = depth + 1
            i = i + 1
        ElseIf ch = "}" Or ch = "]" Then
            depth = depth - 1
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End Function

' startPos の " で始まる文字列リテラルを閉じる " の位置(閉じなければ 0)。
Private Function StrEnd(ByVal jsonText As String, ByVal startPos As Long) As Long
    Dim n As Long
    Dim i As Long
    Dim ch As String

    n = Len(jsonText)
    i = startPos + 1
    Do While i <= n
        ch = Mid$(jsonText, i, 1)
        If ch = "\" Then
            i = i + 2
        ElseIf ch = """" Then
            StrEnd = i
            Exit Function
        Else
            i = i + 1
        End If
    Loop
End Function

' S2の risks[].risk_no を ";1;2;3;" の形で返す(V-S5-03 の照合表)。
Private Function RiskNoList(ByVal s2Json As String) As String
    Dim it As Variant
    Dim acc As String
    Dim noText As String

    acc = V4_SEP
    For Each it In modJsonLite.GetArrayItems(s2Json, "risks")
        noText = Trim$(modJsonLite.GetStr(CStr(it), "risk_no"))
        If LenB(noText) > 0 Then acc = acc & noText & V4_SEP
    Next it
    RiskNoList = acc
End Function

' ============================================================================
' ObjBlock - jsonText の中の "key": { ... } を波括弧の対応で切り出す純関数。
'   headline / business のように**上位と同名のキーを持つ**入れ子では、切り出さ
'   ずに modJsonLite.GetStr を呼ぶと上位の値を拾ってしまう。見つからないときは
'   ""(呼出側は空の値として扱う=検証が発火する)。
' ============================================================================
Public Function ObjBlock(ByVal jsonText As String, ByVal keyName As String) As String
    ObjBlock = Block(jsonText, keyName, "{", "}")
End Function

' ============================================================================
' ArrBlock / ArrItems / ArrCount - 配列の取り出し。15章§5.6 の出力は
'   `headline` の中に `ideas` / `four` / `steps` と**同名の文字列キー**を持ち、
'   しかも `headline` は上位の配列より前に出る。modJsonLite.GetArrayItems は
'   最初に現れたキーを拾うので、そのまま呼ぶと「見出しの一文」を配列と読んで
'   **0件**と数えてしまう。そこで**値が [ で始まる出現**だけを採る。
' ============================================================================
Public Function ArrBlock(ByVal jsonText As String, ByVal keyName As String) As String
    ArrBlock = Block(jsonText, keyName, "[", "]")
End Function

Public Function ArrItems(ByVal jsonText As String, ByVal keyName As String) As Collection
    Set ArrItems = modJsonLite.GetArrayItems(ArrBlock(jsonText, keyName), keyName)
End Function

Public Function ArrCount(ByVal jsonText As String, ByVal keyName As String) As Long
    ArrCount = ArrItems(jsonText, keyName).Count
End Function

' キーの値が openCh で始まる**最初の出現**を、括弧の対応で切り出す純関数。
'   見つからないときは ""(呼出側は空の値として扱う=検証が発火する)。
Private Function Block(ByVal jsonText As String, ByVal keyName As String, _
                       ByVal openCh As String, ByVal closeCh As String) As String
    Dim p As Long
    Dim i As Long
    Dim n As Long
    Dim depth As Long
    Dim ch As String
    Dim inStrFlag As Boolean

    n = Len(jsonText)
    ' 同名のキーが文字列値として先に現れることがある(themes[].headline が
    ' headline オブジェクトより前に出る)。**値が { で始まる出現**だけを採る。
    p = InStr(1, jsonText, """" & keyName & """", vbBinaryCompare)
    Do While p > 0
        i = p + Len(keyName) + 2
        Do While i <= n
            ch = Mid$(jsonText, i, 1)
            If ch = openCh Then Exit Do
            If ch <> ":" And ch <> " " And ch <> vbTab And ch <> vbLf And ch <> vbCr Then
                i = 0
                Exit Do
            End If
            i = i + 1
        Loop
        If i > 0 And i <= n Then Exit Do
        p = InStr(p + 1, jsonText, """" & keyName & """", vbBinaryCompare)
    Loop
    If p = 0 Then Exit Function
    If i < 1 Or i > n Then Exit Function

    depth = 0
    Do While i <= n
        ch = Mid$(jsonText, i, 1)
        If inStrFlag Then
            If ch = "\" Then
                i = i + 1
            ElseIf ch = """" Then
                inStrFlag = False
            End If
        ElseIf ch = """" Then
            inStrFlag = True
        ElseIf ch = openCh Then
            depth = depth + 1
        ElseIf ch = closeCh Then
            depth = depth - 1
            If depth = 0 Then
                Block = Mid$(jsonText, p, i - p + 1)
                Exit Function
            End If
        End If
        i = i + 1
    Loop
End Function
