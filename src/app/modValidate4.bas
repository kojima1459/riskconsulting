Attribute VB_Name = "modValidate4"
Option Explicit

' ============================================================================
' modValidate4 - 顧客向け提案書(S5)の後検証(15章§5.6 CheckS5 検証ルール表)
' ----------------------------------------------------------------------------
' 12章§2 の30,000字契約による modValidate の分割先。modValidate(29,667字)と
'   modValidate2(29,648字)はどちらも満杯であり、**modValidate3 は裁定書38 班A
'   が S1 の証拠検証(V-S1-14/15)で使うため**、班C は次番の4を使う(番号は
'   飛ばさない。同名の Public Function を2つ以上のモジュールに置かない)。
'
' 持つもの:
'   CheckS5      15章§5.6 の13件(V-S5-01からV-S5-13)。""=合格
'   TabooPairs   対訳表(docs/design/提案書_wide/対訳表_社内語から顧客語.md)の
'                「社内語<TAB>顧客語」を vbLf で並べた**唯一の値源**
'   SoftenTaboo  修復後も V-S5-12 だけが残るときの機械置換(docs/29 §5.3。
'                生成を止めない。置換件数を呼出側へ返す)
'   TabooHit     本文に残っている禁止語(";"区切り。0件なら "")
'
' R4準拠(12章§2): Excelトークン・Application.Run・案件データ参照を持たない
'   純関数モジュール。CP932準拠(15章§0 原則7)。
' ============================================================================

Private Const V4_TAB As String = vbTab
Private Const V4_SEP As String = ";"
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
    If LenB(hit) > 0 Then
        Ap r, "[V-S5-12] 顧客向けに書き換えていない語があります: " & hit
    End If

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

' ============================================================================
' TabooPairs - 対訳表(docs/design/提案書_wide/対訳表_社内語から顧客語.md)の
'   「社内語<TAB>顧客語」を vbLf で並べた1本。**禁止語の唯一の値源**であり、
'   15章§5.6 system の対訳行・本表・tools/render_proposal.py の3者を
'   突き合わせる(値源を2箇所に持たない)。
'   §4の「言い換えずに削る」3語には、機械置換の最後の砦で使う中立な代替語を
'   与える(削除すると文が崩れるため。docs/29 §5.3)。
' ============================================================================
Public Function TabooPairs() As String
    Dim s As String
    s = s & "付保" & V4_TAB & "保険のご加入" & vbLf
    s = s & "未付保" & V4_TAB & "保険に入っていない状態" & vbLf
    s = s & "付保ギャップ" & V4_TAB & "保険で手当てできていない部分" & vbLf
    s = s & "未充足" & V4_TAB & "保険の手当てが無い" & vbLf
    s = s & "移転" & V4_TAB & "保険で備える" & vbLf
    s = s & "保有" & V4_TAB & "自社で負担する" & vbLf
    s = s & "トリガー" & V4_TAB & "保険金をお支払いする条件" & vbLf
    s = s & "サブリミット" & V4_TAB & "補償項目ごとの支払限度額" & vbLf
    s = s & "待機期間" & V4_TAB & "補償が始まるまでの期間" & vbLf
    s = s & "保険化" & V4_TAB & "保険での備え方の設計" & vbLf
    s = s & "特約開発" & V4_TAB & "補償内容の新しい設計" & vbLf
    s = s & "組成" & V4_TAB & "仕組みづくり" & vbLf
    s = s & "募集スキーム" & V4_TAB & "ご加入の手続きの流れ" & vbLf
    s = s & "料率" & V4_TAB & "保険料の水準" & vbLf
    s = s & "相関損失" & V4_TAB & "同時に起きる損害" & vbLf
    s = s & "引受" & V4_TAB & "保険のお引き受け" & vbLf
    s = s & "過少保険" & V4_TAB & "補償額が損害に届かない状態" & vbLf
    s = s & "抜け" & V4_TAB & "補償されない部分" & vbLf
    s = s & "免責金額" & V4_TAB & "ご負担いただく金額" & vbLf
    s = s & "支払限度額" & V4_TAB & "お支払いの上限額" & vbLf
    s = s & "リスクユニバース" & V4_TAB & "リスクの全体像" & vbLf
    s = s & "ニューリスク" & V4_TAB & "新しく生まれているリスク" & vbLf
    s = s & "座組" & V4_TAB & "ご提案の構成" & vbLf
    s = s & "ヒアリング" & V4_TAB & "お伺いしたい事項" & vbLf
    s = s & "提案の核" & V4_TAB & "ご提案の前提" & vbLf
    s = s & "攻めの保険活用" & V4_TAB & "成長を後押しする保険の活用" & vbLf
    s = s & "発散段階" & V4_TAB & "構想段階" & vbLf
    s = s & "実装難度" & V4_TAB & "実現までの難易度" & vbLf
    s = s & "顕在化" & V4_TAB & "実際に起きること" & vbLf
    s = s & "打ち手" & V4_TAB & "対策" & vbLf
    s = s & "商材" & V4_TAB & "保険商品" & vbLf
    s = s & "リスク移転可能性" & V4_TAB & "保険での備えやすさ" & vbLf
    s = s & "与信" & V4_TAB & "取引先の支払い能力" & vbLf
    s = s & "座組パターン" & V4_TAB & "ご提案の型" & vbLf
    s = s & "PML" & V4_TAB & "想定最大損害額" & vbLf
    s = s & "CBI" & V4_TAB & "取引先の被災による損害" & vbLf
    s = s & "BI" & V4_TAB & "事業が止まったことによる利益の減少" & vbLf
    s = s & "RTO" & V4_TAB & "復旧までの目標時間" & vbLf
    s = s & "BCP" & V4_TAB & "事業継続計画" & vbLf
    s = s & "OT" & V4_TAB & "工場の制御システム" & vbLf
    s = s & "MFA" & V4_TAB & "多要素認証" & vbLf
    s = s & "EDR" & V4_TAB & "端末の不審な動きを検知する仕組み" & vbLf
    s = s & "KRI" & V4_TAB & "リスクの予兆指標" & vbLf
    s = s & "SLA" & V4_TAB & "サービス水準の取り決め" & vbLf
    s = s & "D&O" & V4_TAB & "会社役員賠償責任保険" & vbLf
    s = s & "PL保険" & V4_TAB & "生産物賠償責任保険" & vbLf
    s = s & "対話の順序" & V4_TAB & "ご説明の順序" & vbLf
    s = s & "クロスセル" & V4_TAB & "追加でご検討いただける備え" & vbLf
    s = s & "仕分け" & V4_TAB & "整理"
    TabooPairs = s
End Function

' ============================================================================
' TabooHit - 本文に残っている禁止語を ";" 区切りで返す(0件なら "")。
'   半角英字だけの語(PML/BI/OT 等)は、前後が英字のときに当たらないようにする
'   (「IoT」「BIG」のような別語の一部を禁止語と数えないため)。
' ============================================================================
Public Function TabooHit(ByVal bodyText As String) As String
    Dim rows() As String
    Dim i As Long
    Dim word As String
    Dim acc As String

    If LenB(bodyText) = 0 Then Exit Function
    rows = Split(TabooPairs(), vbLf)
    For i = LBound(rows) To UBound(rows)
        ' Split() の戻り値へ直接添字を付けない(LibreOffice Basic が解さない)。
        Dim onePair() As String
        onePair = Split(rows(i), V4_TAB)
        word = onePair(0)
        If LenB(word) > 0 Then
            If WordFound(bodyText, word) Then
                If LenB(acc) > 0 Then acc = acc & V4_SEP
                acc = acc & word
            End If
        End If
    Next i
    TabooHit = acc
End Function

' ============================================================================
' SoftenTaboo - 修復後も V-S5-12 だけが残るときの機械置換(docs/29 §5.3)。
'   置換した語数を changed へ返す。**生成は止めない**が、置換したことは
'   呼出側が run_log と警告へ残す(黙って直さない)。
' ============================================================================
Public Function SoftenTaboo(ByVal bodyText As String, ByRef changed As Long) As String
    Dim rows() As String
    Dim i As Long
    Dim pair() As String
    Dim t As String

    changed = 0
    t = bodyText
    If LenB(t) = 0 Then
        SoftenTaboo = t
        Exit Function
    End If
    rows = Split(TabooPairs(), vbLf)
    For i = LBound(rows) To UBound(rows)
        pair = Split(rows(i), V4_TAB)
        If UBound(pair) >= 1 Then
            If WordFound(t, pair(0)) Then
                t = Replace(t, pair(0), pair(1))
                changed = changed + 1
            End If
        End If
    Next i
    SoftenTaboo = t
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
'   modJsonLite は配列と値の取り出ししか持たず、入れ子オブジェクトの中だけを
'   見る口が無い。headline / business のように**上位と同名のキーを持つ**
'   入れ子では、切り出さずに GetStr すると上位の値を拾ってしまう。
'   見つからないときは ""(呼出側は空の値として扱う=検証が発火する)。
' ============================================================================
Public Function ObjBlock(ByVal jsonText As String, ByVal keyName As String) As String
    ObjBlock = Block(jsonText, keyName, "{", "}")
End Function

' ============================================================================
' ArrBlock / ArrItems / ArrCount - 配列の取り出し。
'   15章§5.6 の出力は `headline` の中に `ideas` / `four` / `steps` と**同名の
'   文字列キー**を持ち、しかも `headline` は上位の配列より前に出る。
'   modJsonLite.GetArrayItems は最初に現れたキーを拾うので、そのまま呼ぶと
'   「見出しの一文」を配列と読んで**0件**と数えてしまう(W15 の純層テストが
'   実際にこれを検出した)。そこで**値が [ で始まる出現**だけを採る。
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
