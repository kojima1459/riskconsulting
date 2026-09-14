Attribute VB_Name = "modTestsPure26"
Option Explicit

' ============================================================================
' modTestsPure26 - W15(裁定書38 班B)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は裁定書38 §1 班B(B-10 / Z-46 / Z-49)の文だけから手で
'   書き出した(実装の出力を見てから期待値を合わせない。17章§1)。
'
' 対象(全14本):
'   G1 modKnowledgeRank(B-10)
'     01/02 NgramOverlap の正負(完全一致は最大・無関係語は0)
'     03    NgramOverlap は2字と3字で数を変える(3字グラムの方が厳しい)
'     04/05 RankRows の順位(関連度降順・無関係語は最下位)
'   G2 modKnowledge2.SelectRows(B-10)
'     06 業種完全一致の該当総数(totalHits)を打切り件数と区別して返す
'     07 完全一致がmaxRowsに満たないとき、全業種から関連度の高い順に補充する
'     08 caseText/rankColsが空なら補充しない(挙動不変)
'   G3 modPii(Z-46)
'     09 policy_no**だけ**の検知はKindsOfが"policy_no"のみを返す(警告分岐)
'     10 人名を含む混在はKindsOfが"policy_no"だけにならない(ブロック分岐)
'   G4 modPii.SharesLongFragment(Z-49)
'     11 19字の一致は成立しない(minLen未満)
'     12 20字の一致は成立する
'     13 21字の一致は成立する
'     14 無関係な文字列は一致しない
'
' W15 Round2(裁定書39 §1)で追加した8本:
'   G5 modKnowledgeRank の索引化と上限(R1-03)
'     15 多重集合の共通部の値(手計算)。索引化しても数え方が変わらないこと
'     16 索引版が素朴版(このモジュールが持つ参照実装 NaiveOverlap)と一致する
'     17 順位が素朴版のスコア降順(同点は元の並び)と一致する
'     18 案件側の上限 kb_rank_case_chars で打ち切られること
'     19 行側の上限 kb_rank_row_chars で打ち切られること
'     20 上限0は無制限(既存の呼出と挙動が変わらないこと)
'   G6 打切り総数と境界(R1-04 / G-1)
'     21 補充が起きたとき totalHits は「完全一致の該当総数+補充候補の総数」
'     22 lastRow=0 でも実行時エラー(Err9)にならず0件で返ること
'
' 変異注入(出来レース禁止・裁定書38 共通規約):
'   (a) modKnowledgeRank.RankRows の並べ替えを昇順に変えると04/05が落ちる。
'   (b) modPii.SharesLongFragment の判定を `>= minLen - 1` 等へずらすと
'       11(19字)が誤ってTrueになり落ちる。
'
' CP932準拠: 本文・注釈ともにCP932内の文字だけで書く(絵文字不可)。
' ============================================================================

Public Sub RunAll()
    Dim i As Long, grpName As String
    For i = 1 To 6
        grpName = "W15B-G" & CStr(i)
        On Error Resume Next
        Err.Clear
        RunGroup i
        If Err.Number <> 0 Then
            modTestRunner.Check grpName & "(グループ全体)", False, _
                "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
                ") ※未実装/未注入の可能性"
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Sub

Private Sub RunGroup(ByVal idx As Long)
    Select Case idx
    Case 1: T_NgramRank
    Case 2: T_SelectRows
    Case 3: T_PiiKinds
    Case 4: T_SharesLongFragment
    Case 5: T_RankIndexAndCap
    Case 6: T_TotalHitsAndEmpty
    End Select
End Sub

Private Sub ChkB(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check nm, cond, detail
End Sub

Private Sub ChkN(ByVal nm As String, ByVal act As Long, ByVal want As Long)
    modTestRunner.Check nm, (act = want), "期待=" & want & " 実際=" & act
End Sub

Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), "期待=[" & want & "] 実際=[" & act & "]"
End Sub

' ============================================================================
' G1 modKnowledgeRank(B-10)
' ============================================================================
Private Sub T_NgramRank()
    ' 01 完全一致の文字列どうしはn-gram重なりが最大になる(2字7件+3字6件=13)。
    ChkN "Test_W15B_01_完全一致は最大の重なり数_裁定書38B-10", _
        modKnowledgeRank.NgramOverlap("abcdefgh", "abcdefgh", 2) + _
        modKnowledgeRank.NgramOverlap("abcdefgh", "abcdefgh", 3), 13

    ' 02 無関係な文字列は重なり0。
    ChkN "Test_W15B_02_無関係な文字列は重なり0件_裁定書38B-10", _
        modKnowledgeRank.NgramOverlap("abcdefgh", "zzzzzzzz", 2), 0

    ' 03 部分一致(先頭4字だけ共通)は2字グラムで3件・3字グラムで2件。
    ChkN "Test_W15B_03a_部分一致の2字グラム件数_裁定書38B-10", _
        modKnowledgeRank.NgramOverlap("abcdefgh", "abcdxxxx", 2), 3
    ChkN "Test_W15B_03b_部分一致の3字グラム件数(2字より厳しい)_裁定書38B-10", _
        modKnowledgeRank.NgramOverlap("abcdefgh", "abcdxxxx", 3), 2

    ' 04/05 RankRows: 無関係(0件)<部分一致(5件)<完全一致(13件)の順に並ぶ
    '   (入力順は無関係・部分一致・完全一致の順に置き、出力順が逆転すること
    '   を見る=安定ソートの検証を兼ねる)。
    Dim rowTexts(1 To 3) As String
    rowTexts(1) = "zzzzzzzz"
    rowTexts(2) = "abcdxxxx"
    rowTexts(3) = "abcdefgh"
    Dim order() As Long
    Dim n As Long
    n = modKnowledgeRank.RankRows("abcdefgh", rowTexts, order)
    ChkN "Test_W15B_04_RankRowsは全件を返す_裁定書38B-10", n, 3
    ChkB "Test_W15B_05_RankRowsは関連度降順(完全一致が1位)_裁定書38B-10", _
        (n = 3 And order(1) = 3 And order(2) = 2 And order(3) = 1), _
        "実際の順=" & JoinLongs(order, ",")
End Sub

' ============================================================================
' G2 modKnowledge2.SelectRows(B-10)
' ============================================================================
Private Function MakeBlk(ByVal headerRow As String, ByVal dataRows As String) As Variant
    Dim hdr As Variant, rows As Variant
    hdr = Split(headerRow, ";")
    rows = Split(dataRows, vbLf)
    Dim nCols As Long, nRows As Long
    nCols = UBound(hdr) - LBound(hdr) + 1
    nRows = UBound(rows) - LBound(rows) + 1
    Dim arr() As Variant
    ReDim arr(1 To nRows + 1, 1 To nCols)
    Dim c As Long, r As Long, cells As Variant
    For c = 1 To nCols
        arr(1, c) = hdr(c - 1)
    Next c
    For r = 1 To nRows
        cells = Split(rows(r - 1), ";")
        For c = 1 To nCols
            arr(r + 1, c) = cells(c - 1)
        Next c
    Next r
    MakeBlk = arr
End Function

Private Sub T_SelectRows()
    Dim blk As Variant
    blk = MakeBlk("id;industry;body", _
        "C1;09;たまご" & vbLf & _
        "C2;07;abcdxxxx" & vbLf & _
        "C3;07;zzzzzzzz" & vbLf & _
        "C4;07;abcdefgh")

    Dim selOut As Variant, idsOut As String, totalHits As Long, n As Long

    ' 06 業種完全一致は1件(C1)しかないが、maxRows=1で打ち切ればtotalHitsは
    '    その先の該当有無に関わらず「完全一致した総数」を返す(この例では1)。
    idsOut = vbNullString
    n = modKnowledge2.SelectRows(blk, 5, "id", "industry", "09", 1, selOut, idsOut, totalHits)
    ChkN "Test_W15B_06a_完全一致の使用数_裁定書38B-10", n, 1
    ChkN "Test_W15B_06b_完全一致の該当総数(totalHits)_裁定書38B-10", totalHits, 1

    ' 07 caseText/rankColsを与えると、完全一致(1件)がmaxRows(3件)に満たない
    '    ぶんを、全業種から関連度の高い順(C4→C2。C3は重なり0で対象外)に補う。
    idsOut = vbNullString
    n = modKnowledge2.SelectRows(blk, 5, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "abcdefgh", "body")
    ChkN "Test_W15B_07a_並べ替え補充後の使用数_裁定書38B-10", n, 3
    ChkS "Test_W15B_07b_並べ替え補充の順序(完全一致+関連度降順)_裁定書38B-10", _
        idsOut, "C1;C4;C2"

    ' 08 caseText/rankColsが空なら補充しない(打切りのまま。挙動不変)。
    idsOut = vbNullString
    n = modKnowledge2.SelectRows(blk, 5, "id", "industry", "09", 3, selOut, idsOut, totalHits)
    ChkN "Test_W15B_08_caseText省略時は補充しない_裁定書38B-10", n, 1
End Sub

' ============================================================================
' G3 modPii(Z-46)
' ============================================================================
Private Sub T_PiiKinds()
    ' 09 policy_noだけの検知(伝書鳩20260912 3-2・裁定書37 A-08/C-2の実例と
    '    同型)は、KindsOfが"policy_no"だけを返す(=警告のみ・登録は続行)。
    ChkS "Test_W15B_09_policy_noだけの検知_裁定書38Z-46", _
        modPii.KindsOf("証券番号 AB-1234567 の件でご連絡します。"), "policy_no"

    ' 10 人名を含む混在は"policy_no"単独にならない(=ブロック維持)。
    Dim kinds As String
    kinds = modPii.KindsOf("証券番号 AB-1234567 の件、担当の田中様に工程表をお送りください。")
    ChkB "Test_W15B_10_人名混在はpolicy_no単独にならない_裁定書38Z-46", _
        (LenB(kinds) > 0 And kinds <> "policy_no"), _
        "実際=[" & kinds & "]"
End Sub

' ============================================================================
' G4 modPii.SharesLongFragment(Z-49)
' ============================================================================
Private Sub T_SharesLongFragment()
    Dim src As String
    src = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

    ' 11 19字は成立しない(20字未満はminLenの窓すら作れない)。
    ChkB "Test_W15B_11_19字は一致とみなさない_裁定書38Z-49", _
        Not modPii.SharesLongFragment(Left$(src, 19), src, 20), _
        "19字断片で誤って一致した"

    ' 12 20字は成立する。
    ChkB "Test_W15B_12_20字は一致とみなす_裁定書38Z-49", _
        modPii.SharesLongFragment(Left$(src, 20), src, 20), _
        "20字断片が一致しなかった"

    ' 13 21字は成立する。
    ChkB "Test_W15B_13_21字は一致とみなす_裁定書38Z-49", _
        modPii.SharesLongFragment(Left$(src, 21), src, 20), _
        "21字断片が一致しなかった"

    ' 14 無関係な文字列(sourceに存在しない)は一致しない。
    ChkB "Test_W15B_14_無関係な文字列は一致しない_裁定書38Z-49", _
        Not modPii.SharesLongFragment("ZZZZZZZZZZZZZZZZZZZZ", src, 20), _
        "無関係な断片が誤って一致した"
End Sub

' ============================================================================
' G5 modKnowledgeRank の索引化と上限(裁定書39 R1-03)
' ----------------------------------------------------------------------------
' 期待値の出典: 14章§6 の NgramOverlap 契約(「共通して現れるn文字グラムの
'   延べ数=多重集合の共通部」「貪欲マッチング(先勝ち)」)と、裁定書39 §1 R1-03
'   (「比較文字列に上限」「n-gram を Collection で索引化して O(N+M) へ」)。
'
' NaiveOverlap は**この表の契約からテスト側で書き起こした参照実装**(二重ループ)。
'   実装を索引版へ入れ替えても数え方が1件も変わらないことを、これと突き合わせて
'   固定する(索引版だけを見て期待値を作らない)。
' ============================================================================
Private Function NaiveOverlap(ByVal a As String, ByVal b As String, ByVal n As Long) As Long
    If n <= 0 Then Exit Function
    Dim la As Long, lb As Long
    la = Len(a) - n + 1
    lb = Len(b) - n + 1
    If la < 1 Or lb < 1 Then Exit Function

    Dim used() As Boolean
    ReDim used(1 To lb)
    Dim i As Long, j As Long, cnt As Long, gramA As String
    For i = 1 To la
        gramA = Mid$(a, i, n)
        For j = 1 To lb
            If Not used(j) Then
                If gramA = Mid$(b, j, n) Then
                    used(j) = True
                    cnt = cnt + 1
                    Exit For
                End If
            End If
        Next j
    Next i
    NaiveOverlap = cnt
End Function

Private Sub T_RankIndexAndCap()
    ' 15 多重集合の共通部(手計算)。"aaaa" の2字グラムは "aa" が3個、"aa" は1個
    '    なので共通部は min(3,1)=1。"aaaa" どうしなら min(3,3)=3。
    ChkN "Test_R1-03_15a_重複グラムは多重集合の共通部で数える_裁定書39R1-03", _
        modKnowledgeRank.NgramOverlap("aaaa", "aa", 2), 1
    ChkN "Test_R1-03_15b_同一文字列の重複グラムは全数が共通部_裁定書39R1-03", _
        modKnowledgeRank.NgramOverlap("aaaa", "aaaa", 2), 3

    ' 16 索引版と素朴版の一致(重複の多い入力・部分一致・無関係を混ぜる)。
    Dim aTxt As String, i As Long, diffSum As Long
    Dim bTxt(1 To 4) As String
    aTxt = String$(40, "a") & "bcbcbcbc" & String$(20, "d")
    bTxt(1) = String$(15, "a") & "cbcbcb"
    bTxt(2) = "bcbcbcbc"
    bTxt(3) = String$(30, "d")
    bTxt(4) = "zzzzzzzzzz"
    For i = 1 To 4
        diffSum = diffSum + _
            Abs(modKnowledgeRank.NgramOverlap(aTxt, bTxt(i), 2) - NaiveOverlap(aTxt, bTxt(i), 2)) + _
            Abs(modKnowledgeRank.NgramOverlap(aTxt, bTxt(i), 3) - NaiveOverlap(aTxt, bTxt(i), 3))
    Next i
    ChkN "Test_R1-03_16_索引版は素朴版と同じ重なり数を返す_裁定書39R1-03", diffSum, 0

    ' 17 順位も素朴版と同じ(13 / 5 / 5 / 0 -> 3,2,4,1。同点は元の並び)。
    Dim rowTexts(1 To 4) As String
    Dim order() As Long
    Dim n As Long
    rowTexts(1) = "zzzzzzzz"
    rowTexts(2) = "abcdxxxx"
    rowTexts(3) = "abcdefgh"
    rowTexts(4) = "efghyyyy"
    n = modKnowledgeRank.RankRows("abcdefgh", rowTexts, order)
    ChkB "Test_R1-03_17_索引版でも順位は素朴版と同じ_裁定書39R1-03", _
        (n = 4 And order(1) = 3 And order(2) = 2 And order(3) = 4 And order(4) = 1), _
        "実際の順=" & JoinLongs(order, ",")

    ' 18/20 案件側の上限(kb_rank_case_chars 相当)で打ち切られること。
    '    caseText の後半(BBBBBBBB)を切ると C2 の重なりが0になり補充されない。
    Dim blkA As Variant, selOut As Variant, idsOut As String, totalHits As Long
    blkA = MakeBlk("id;industry;body", _
        "C1;09;pppppppp" & vbLf & _
        "C2;07;BBBBBBBB" & vbLf & _
        "C3;07;AAAAAAAA")

    idsOut = vbNullString
    n = modKnowledge2.SelectRows(blkA, 4, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "AAAAAAAABBBBBBBB", "body", 0, 0)
    ChkS "Test_R1-03_20_上限0は無制限で従来どおり補充する_裁定書39R1-03", idsOut, "C1;C2;C3"

    idsOut = vbNullString
    n = modKnowledge2.SelectRows(blkA, 4, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "AAAAAAAABBBBBBBB", "body", 8, 0)
    ChkS "Test_R1-03_18_案件側の上限で打ち切られる_裁定書39R1-03", idsOut, "C1;C3"

    ' 19 行側の上限(kb_rank_row_chars 相当)で打ち切られること。
    Dim blkB As Variant
    blkB = MakeBlk("id;industry;body", _
        "C1;09;pppppppp" & vbLf & _
        "C2;07;ZZZZZZZZAAAAAAAA" & vbLf & _
        "C3;07;AAAAAAAAZZZZZZZZ")
    idsOut = vbNullString
    n = modKnowledge2.SelectRows(blkB, 4, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "AAAAAAAA", "body", 0, 8)
    ChkS "Test_R1-03_19_行側の上限で打ち切られる_裁定書39R1-03", idsOut, "C1;C3"

    ' 26 14章§6 の契約「大小文字・かな漢字はそのまま比較する」(裁定書40 P-m1)。
    '    索引のキーが大小文字を畳むと、別物の行が重なったことにされて上位に来る
    '    (Basic の Collection はキー照合で大小文字を区別しない=素のグラムを
    '    キーにすると畳まれる)。**畳まれないこと**と**同一なら重なること**の
    '    両方向を固定する(片側だけだと「常に0を返す」実装が緑になる)。
    ChkB "Test_P-m1_26_大小文字違いのグラムは重ならない_裁定書40P-m1", _
        (modKnowledgeRank.NgramOverlap("ABAB", "abab", 2) = 0) And _
        (modKnowledgeRank.NgramOverlap("ABAB", "ABAB", 2) = 3), _
        "大小文字違い=" & CStr(modKnowledgeRank.NgramOverlap("ABAB", "abab", 2)) & _
        "(期待0) 同一=" & CStr(modKnowledgeRank.NgramOverlap("ABAB", "ABAB", 2)) & "(期待3)"

    ' 27 案件側の索引は**行数によらず2回**(2字と3字で各1回)しか作らない
    '    (裁定書40 P-M3(a))。行ごとに作り直す実装だと 4行=8回・12行=24回になる。
    '    1呼び出しの NgramOverlap が1回であることも併せて見る(両方向)。
    Dim rows12(1 To 12) As String
    Dim b4 As Long, b12 As Long, b1 As Long
    For i = 1 To 12
        rows12(i) = "abcd" & String$(i, "x")
    Next i
    modKnowledgeRank.ResetIndexBuilds
    n = modKnowledgeRank.RankRows("abcdefgh", rowTexts, order)
    b4 = modKnowledgeRank.IndexBuilds()
    modKnowledgeRank.ResetIndexBuilds
    n = modKnowledgeRank.RankRows("abcdefgh", rows12, order)
    b12 = modKnowledgeRank.IndexBuilds()
    modKnowledgeRank.ResetIndexBuilds
    n = modKnowledgeRank.NgramOverlap("abcdefgh", "abcd", 2)
    b1 = modKnowledgeRank.IndexBuilds()
    ChkB "Test_P-M3_27_索引は行数によらず2回しか作らない_裁定書40P-M3", _
        (b4 = 2 And b12 = 2 And b1 = 1), _
        "4行=" & CStr(b4) & " 12行=" & CStr(b12) & " 単発=" & CStr(b1) & " (期待 2/2/1)"

    ' 28 索引を使い回しても点数が変わらない(=借りた個数を必ず返している)。
    '    RankRows の並びが、1行ずつ NgramOverlap で数えた点数の降順に一致する。
    '    使い回しで個数が減る実装だと、後ろの行ほど点数が下がって並びが崩れる。
    Dim rowsX(1 To 5) As String
    Dim scX(1 To 5) As Long
    Dim caseX As String, okOrder As Boolean, scSum As Long
    caseX = "abcdefghijklmnop"
    rowsX(1) = "zzzzzzzz"
    rowsX(2) = "abcdefgh"
    rowsX(3) = "ijklmnop"
    rowsX(4) = "abcdefghijklmnop"
    rowsX(5) = "qrstuvwx"
    n = modKnowledgeRank.RankRows(caseX, rowsX, order)
    okOrder = (n = 5)
    For i = 1 To 5
        scX(i) = modKnowledgeRank.NgramOverlap(caseX, rowsX(i), 2) + _
                 modKnowledgeRank.NgramOverlap(caseX, rowsX(i), 3)
        scSum = scSum + scX(i)
    Next i
    If okOrder Then
        For i = 2 To 5
            If scX(order(i - 1)) < scX(order(i)) Then okOrder = False
        Next i
    End If
    ChkB "Test_P-M3_28_索引を使い回しても点数と順位が変わらない_裁定書40P-M3", _
        (okOrder And scSum > 0 And scX(4) > scX(2)), _
        "点数=" & CStr(scX(1)) & "," & CStr(scX(2)) & "," & CStr(scX(3)) & "," & _
        CStr(scX(4)) & "," & CStr(scX(5)) & " 順=" & JoinLongs(order, ",")

    ' 29 候補行数の上限(config kb_rank_max_rows 相当の第13引数)。上限1なら
    '    シート順で先頭の候補(C2)だけが順位付けの対象になり、C3 は採られない。
    '    上限0(無制限)では両方採られる、の両方向を1本で固定する(裁定書40 P-M3(b))。
    '    候補総数(第14引数)は上限に関係なく数える=run_log へ「打ち切った」と
    '    書けるのはこの値が上限を超えたときだけ。
    Dim blkCap As Variant, candSeen As Long
    blkCap = MakeBlk("id;industry;body", _
        "C1;09;pppppppp" & vbLf & _
        "C2;07;AAAAAAAA" & vbLf & _
        "C3;07;BBBBBBBB")
    idsOut = vbNullString
    totalHits = 0
    candSeen = 0
    n = modKnowledge2.SelectRows(blkCap, 4, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "BBBBBBBBAAAAAAAA", "body", 0, 0, 0, candSeen)
    Dim idsNoCap As String, seenNoCap As Long
    idsNoCap = idsOut
    seenNoCap = candSeen
    idsOut = vbNullString
    totalHits = 0
    candSeen = 0
    n = modKnowledge2.SelectRows(blkCap, 4, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "BBBBBBBBAAAAAAAA", "body", 0, 0, 1, candSeen)
    ChkB "Test_P-M3_29_候補行数の上限で並べ替えを打ち切る_裁定書40P-M3", _
        (idsNoCap = "C1;C2;C3" And seenNoCap = 2 And idsOut = "C1;C2" And _
         candSeen = 2 And totalHits = 2), _
        "上限なし=[" & idsNoCap & "]/" & CStr(seenNoCap) & " 上限1=[" & idsOut & _
        "]/" & CStr(candSeen) & " totalHits=" & CStr(totalHits)

    ' 30 **合算版と素朴版の等価**(裁定書41 §2)。RankRows が使う合算走査
    '    (2字と3字を1回の走査で数える Overlap23)の点数が、2字と3字を別々に
    '    数えた点数の和と一致すること。27/28 はこの枝を踏まないので、合算走査を
    '    静かに壊す変異(3字探査の枝刈り条件・3字の借用返却)が素通りしていた。
    '    点数そのものは外から取れないので、**並びで**突き合わせる(並べ替えは
    '    安定なので、点数が1つでもずれれば並びが変わる検体を選ぶ)。
    '    第3引数は手計算の並び。同点で素通りする検体を混ぜないための担保。
    '      A "abcQabQ": 行1=2字2+3字1=3 / 行2=2字3+3字1=4  -> 2,1
    '      B "abcdefghij": 行1=5+4=9 / 行2=7+6=13          -> 2,1
    Dim eq30 As String
    eq30 = Equiv23("abcQabQ", "abc|abZabZabc", "2,1")
    If LenB(eq30) = 0 Then eq30 = Equiv23("abcdefghij", "abcdef|abcdefgh", "2,1")
    If LenB(eq30) = 0 Then eq30 = Equiv23("abcabcabc", "abcabc|abc|cba|abcabcabc", "4,1,2,3")
    If LenB(eq30) = 0 Then eq30 = Equiv23("aaaaab", "aaaa|aab|baaa|ab", "1,2,3,4")
    If LenB(eq30) = 0 Then eq30 = Equiv23("abcdefghij", "zzzz|ij|hij|defghij", "4,3,2,1")
    If LenB(eq30) = 0 Then eq30 = Equiv23("xyzxyzxyz", "xyzxyz|yzxyzx|zxy|qqq", "1,2,3,4")
    ChkS "Test_P-M3_30_合算走査の点数が2字と3字の和と一致する_裁定書41§2", eq30, ""
End Sub

' ============================================================================
' Equiv23 - RankRows の並びが「NaiveOverlap(2)+NaiveOverlap(3) の安定降順」と
'   一致するかを1検体ぶん確かめる。一致すれば ""、違えば理由を返す。
' ----------------------------------------------------------------------------
'   期待値の出典は 14章§6 の RankRows 契約(「関連度の高い順」「重なり数は
'   2〜3字n-gramの多重集合の共通部」)と、このモジュールが持つ参照実装
'   NaiveOverlap だけ。**実装の出力は見ない**(17章§1)。
'   wantOrder には手計算の並びを渡す(非空のときだけ照合)。同点ばかりで
'   「点数が変わっても並びが変わらない」検体を混ぜていないことの担保であり、
'   これが無いと Test_P-M3_28 と同じ骨抜きが起きる。
' ============================================================================
Private Function Equiv23(ByVal caseText As String, ByVal rowsPipe As String, _
                         ByVal wantOrder As String) As String
    Dim parts() As String, i As Long, j As Long, cnt As Long
    parts = Split(rowsPipe, "|")
    cnt = UBound(parts) - LBound(parts) + 1
    If cnt < 2 Then
        Equiv23 = "[" & caseText & "] 検体の行数が" & CStr(cnt) & "件です(2件以上)"
        Exit Function
    End If

    Dim rw() As String, sc() As Long, ord() As Long
    ReDim rw(1 To cnt)
    ReDim sc(1 To cnt)
    ReDim ord(1 To cnt)
    For i = 1 To cnt
        rw(i) = parts(LBound(parts) + i - 1)
        sc(i) = NaiveOverlap(caseText, rw(i), 2) + NaiveOverlap(caseText, rw(i), 3)
        ord(i) = i
    Next i

    ' 安定な挿入ソート(降順・同点は元の並び)。14章§6 の「関連度の高い順」。
    Dim ks As Long, ki As Long
    For i = 2 To cnt
        ks = sc(i)
        ki = ord(i)
        j = i - 1
        Do While j >= 1
            If sc(j) < ks Then
                sc(j + 1) = sc(j)
                ord(j + 1) = ord(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        sc(j + 1) = ks
        ord(j + 1) = ki
    Next i

    Dim got() As Long, n As Long, gotText As String, wantText As String
    n = modKnowledgeRank.RankRows(caseText, rw, got)
    If n <> cnt Then
        Equiv23 = "[" & caseText & "] RankRows が" & CStr(n) & "件(期待" & CStr(cnt) & ")"
        Exit Function
    End If
    For i = 1 To cnt
        If i > 1 Then
            gotText = gotText & ","
            wantText = wantText & ","
        End If
        gotText = gotText & CStr(got(i))
        wantText = wantText & CStr(ord(i))
    Next i
    If gotText <> wantText Then
        Equiv23 = "[" & caseText & "] 合算版=" & gotText & " 素朴版=" & wantText
        Exit Function
    End If
    If LenB(wantOrder) > 0 Then
        If wantText <> wantOrder Then
            Equiv23 = "[" & caseText & "] 素朴版=" & wantText & " 手計算=" & wantOrder
        End If
    End If
End Function

' ============================================================================
' G6 打切り総数と境界(裁定書39 R1-04 / G-1)
' ============================================================================
Private Sub T_TotalHitsAndEmpty()
    Dim blk As Variant, selOut As Variant, idsOut As String, totalHits As Long, n As Long
    blk = MakeBlk("id;industry;body", _
        "C1;09;たまご" & vbLf & _
        "C2;07;abcdxxxx" & vbLf & _
        "C3;07;zzzzzzzz" & vbLf & _
        "C4;07;abcdefgh")

    ' 21 補充が起きたとき、totalHits は「完全一致の該当総数(C1の1件)+**実際に
    '    補充で採用した行数**(C4/C2の2件。C3 は重なり0で採らない)」=3。
    '    候補を作っただけで足すと、1行も打ち切っていない案件でも SEC-14 が
    '    「該当N件のうちM件を使用」と出る(裁定書40 P-M2)。使用数を下回っても
    '    ならない(kb_cut:cases=3/1 は読めない値)。
    idsOut = vbNullString
    n = modKnowledge2.SelectRows(blk, 5, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "abcdefgh", "body")
    ChkN "Test_P-M2_21_補充時のtotalHitsは完全一致+採用した補充行_裁定書40P-M2", totalHits, 3

    ' 22 lastRow=0(データ行が1本も無い)でも、補充の候補配列を ReDim(1 To 0) して
    '    実行時エラー9 を投げないこと。0件で静かに返る。
    idsOut = vbNullString
    totalHits = 0
    n = modKnowledge2.SelectRows(blk, 0, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "abcdefgh", "body")
    ChkN "Test_G-1_22_lastRow0でも実行時エラーにならない_裁定書39G-1", n, 0

    ' 23 補充候補はあるが**1件も採用しなかった**とき(重なり0の行しかない)は、
    '    totalHits = 完全一致の該当数。ここが膨らむと modExportHtml の
    '    `kbCasesTotal > kbCasesUsed` が常時真になり、打ち切っていない案件で
    '    SEC-14 が「打ち切った」と嘘をつく(裁定書40 P-M2 の再現手順そのもの)。
    Dim blkNo As Variant
    blkNo = MakeBlk("id;industry;body", _
        "C1;09;abcdefgh" & vbLf & _
        "C2;07;zzzzzzzz" & vbLf & _
        "C3;07;yyyyyyyy" & vbLf & _
        "C4;07;wwwwwwww")
    idsOut = vbNullString
    totalHits = 0
    n = modKnowledge2.SelectRows(blkNo, 5, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "abcdefgh", "body")
    ChkN "Test_P-M2_23_補充0件ならtotalHitsは完全一致数_裁定書40P-M2", totalHits, 1

    ' 24 lastRow<0(見出し行すら無いブロックを渡された)でも、入口の
    '    ReDim hits(1 To lastRow + 1) で実行時エラー9 を投げない。本関数は
    '    14章§6 の公開口であり、呼び出し側がエラーを握らない(裁定書40 P-m2)。
    idsOut = vbNullString
    totalHits = 0
    n = modKnowledge2.SelectRows(blk, -1, "id", "industry", "09", 3, selOut, idsOut, _
                                  totalHits, "abcdefgh", "body")
    ChkN "Test_P-m2_24_lastRowが負でも実行時エラーにならない_裁定書40P-m2", n, 0
End Sub

' JoinLongs - Long配列をカンマ等で連結する(裁定書44 追加裁定A-8a)。
'   実VBAの Join() は Variant/String の配列しか受け付けず、Long() 配列を渡すと
'   Err 5(プロシージャの呼び出し、または引数が不正です)になる。LibreOffice
'   Basic は型に寛容でこの差異を素通りするため、実機で初めて発覚した
'   (W15B-G1/G5 が Err 5 で落ち、以降の8本が未実行になり本数不一致も連鎖した)。
'   空配列(LBound>UBound)は空文字を返す。
Private Function JoinLongs(ByRef arr() As Long, ByVal sep As String) As String
    Dim i As Long, outText As String
    On Error GoTo Empty0
    For i = LBound(arr) To UBound(arr)
        If i > LBound(arr) Then outText = outText & sep
        outText = outText & CStr(arr(i))
    Next i
    JoinLongs = outText
    Exit Function
Empty0:
    JoinLongs = vbNullString
End Function
