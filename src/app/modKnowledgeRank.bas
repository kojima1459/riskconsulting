Attribute VB_Name = "modKnowledgeRank"
Option Explicit

' ============================================================================
' modKnowledgeRank - ナレッジ選抜の並べ替え純関数(裁定書38 §1 班B B-10)
' ----------------------------------------------------------------------------
' 背景(伝書鳩3-3・裁定書38 B-10): modKnowledge2.SelectRows は業種コード
'   「完全一致」で絞ったあと**シート上から順にmaxRows件で打ち切る**。完全一致で
'   maxRows件に満たないときに、全業種の行から**関連度の高いものを補う**ための
'   スコア付けだけを、ここに独立させる(責務: modKnowledge2 = 絞込・打切り /
'   本モジュール = 全業種補充の順位付け)。
'
' 手法: 埋め込みは重いので**2〜3字のn-gram重なり数**で代用する(伝書鳩3-3の
'   「おすすめ」の②)。日本語は分かち書きが無いので文字n-gramが実用的。
'
' R4(12章§2): Excelトークンを一切持たない純関数(Scripting.Dictionary等の
'   COMオブジェクトも使わない。CreateObjectはR4のExcel許可外だが、この
'   モジュールはCreateObject自体を使わない設計にした。配列だけで完結する)。
' 14章§6へ登記(公開口): NgramOverlap / RankRows。
'   IndexBuilds / ResetIndexBuilds は**計測専用の口**(下の(3)(c))。
' ============================================================================
'
' 性能の経緯(ここを読まずに触ると、また数分固まる実装へ戻る):
'   (1) 初版は二重ループの O(|a|x|b|)。案件本文2万字 x 行500字 x 200行で
'       Python 実測278秒。
'   (2) 裁定書39 R1-03 で a 側の n-gram を Collection で索引化し O(|a|+|b|) へ。
'       1呼び出しは約40倍速くなったが、**RankRows は候補行数 N に対して
'       NgramOverlap を 2N 回呼び、その都度 a 側の索引を作り直していた**ため、
'       本番規模では遅いままだった(裁定書40 P-M3。本モジュールの計測器
'       scratchpad/P2 による LibreOffice 実測: 案件3,000字 x 行2,000字 x
'       200行の RankRows = 77.0 秒)。
'   (3) 裁定書40 P-M3 で次の3点を直した(いずれも実測で効果を確かめた)。
'       (a) **索引は n ごとに1回だけ作り、全候補行で使い回す**。索引を受け取る
'           版(Overlap23)と、その場で作る版(NgramOverlap)に分けた。
'           作り直しの回数は IndexBuilds() で数えられる(回帰テストが
'           「行数によらず2回」を機械で押さえる)。
'       (b) 索引の実体を Collection から**開番地法のハッシュ表(配列2本)**へ
'           置き換えた。LibreOffice 実測で Collection の Item は1回 33マイクロ秒
'           (30,000回=1.0秒)、配列のハッシュ表は1回 5マイクロ秒
'           (400,000回=2.0秒)だった。行あたり2,000回の参照が要るので、
'           Collection のままでは (a) だけでは受入条件に届かない計算になる。
'       (c) 行の走査は**2字と3字を1回の走査で**数え、先頭2字が案件本文に
'           1つも無い位置では3字の探査を省く(数え方は変わらない。2字が無ければ
'           その2字で始まる3字も在りえない)。LibreOffice Basic は配列1参照が
'           約1.5マイクロ秒あり、位置あたりの参照回数が総時間を決める。
'       受入条件(裁定書40 P-M3(c)): 案件3,000字 x 行2,000字 x 200行の
'       RankRows が LibreOffice で10秒以内。**同じ計測器・同じ入力での実測
'       (置換前 -> 置換後)**:
'         ・現実的な入力(案件と行が同じ日本語語彙から成る) 86.0 -> 9.0 秒
'         ・重なりゼロ(行が案件と無縁)                     83.0 -> 7.0 秒
'         ・最悪(行が案件本文のほぼ写し。2字の9割が命中)   79.0 -> 16.0 秒
'       本番は config `kb_rank_max_rows`(既定60)で候補行を60行に切るので、
'       同じ入力で 27.0 -> 2.0 秒(現実的)/ 25.0 -> 4.0 秒(最悪)である。
'       最悪の1本だけは10秒に収まらないが、これは「KBの行が案件本文の写し」
'       という、重なりが最大になる合成入力である(実データでは起きない)。
'
' キーに生のグラム文字列を使わない理由(14章§6の契約「大小文字・かな漢字は
'   そのまま比較する」): VBA/Basic の Collection のキー照合は**大小文字を
'   区別しない**(ロケールによっては半角全角・かなカナも畳む)。そこで
'   グラムを**コードポイントの数値**(1文字16bit を 65536 進で連結した Double)
'   へ直して比較する。数値なので "AB" と "ab" は別物のまま扱われる
'   (この契約は modTestsPure26 の Test_P-m1_26 が両方向で固定している)。
'   n=3 までなら 48bit で、Double の 53bit 仮数に**誤差なく**収まる。4字以上は
'   丸められるので素朴版(WideOverlap)へ落とす。
'
' CP932準拠(15章§0 原則7)・例外を投げない。
' ============================================================================

' 比較に使う n-gram の字数(2字と3字の合算。13章§3.1)。
Private Const RK_N_MIN As Long = 2
Private Const RK_N_MAX As Long = 3

' 1文字ぶんの桁(コードポイントを 65536 進で連結する)。
Private Const RK_BASE As Double = 65536#

' 表の位置を出す多項式ハッシュの基数。**キーの下位ビットをそのまま位置に
'   使わない**ための措置(裁定書40 P-M3 の計測で判明): 位置=キーの下位ビット
'   だと「英大文字+英小文字」のような字種では2文字目だけが効いて候補が数十か所
'   へ集中し、開番地法の探査が1回あたり数百歩になって**索引化の意味が消える**
'   (実測: 案件3,000字 x 行2,000字の1回が 5.1秒)。基数131の多項式なら全文字が
'   効く。n<=3 なら ((c1*131+c2)*131+c3) <= 1.14e9 で Long に収まる。
Private Const RK_MUL As Long = 131

' 索引の構築回数(計測専用。判定には一切使わない)。
Private mIdxBuilds As Long

' ============================================================================
' NgramOverlap - a と b に共通して現れる n文字グラムの延べ数(多重集合の共通部)。
'   貪欲マッチング(先勝ち)で数える。b 側の同じ位置を二重に使わない。
'   n<=0、または a か b が n文字未満のときは 0。大小文字・かな漢字はそのまま
'   比較する(正規化は呼び出し側の責務。ここは純粋な文字列比較のみ)。
'
'   **その場で索引を作る版**(裁定書40 P-M3(a))。同じ a に対して何行も数える
'   ときは RankRows を使うこと(索引を1回だけ作って使い回す)。
'   数え方は索引化の前後で不変: 貪欲先勝ちの結果は「グラムごとの
'   min(a の個数, b の個数)の総和」と一致するため、残り個数を1つずつ減らす
'   だけで同じ値になる(modTestsPure26 が素朴版の参照実装と突き合わせて固定)。
' ============================================================================
Public Function NgramOverlap(ByVal a As String, ByVal b As String, ByVal n As Long) As Long
    Dim key2() As Double, rem2() As Long, mask2 As Long, ok2 As Boolean
    Dim key3() As Double, rem3() As Long, mask3 As Long, ok3 As Boolean
    Dim cb() As Long

    ' 2字・3字以外(1字と4字以上)は素朴版で数える。1字は本PJの経路が使わず、
    '   4字以上は数値キーが 53bit を超えて丸められるため(どちらも速度が要る
    '   経路ではないので、答えの正しさを優先して単純な二重ループにする)。
    If n <> RK_N_MIN And n <> RK_N_MAX Then
        NgramOverlap = WideOverlap(a, b, n)
        Exit Function
    End If
    If Len(b) < n Then Exit Function

    ' 使う側だけ本物の索引を作り、使わない側は**空の表を確保する**
    '   (未確保の配列を ByRef で渡さないため。BuildGramIndex は作れなくても
    '   必ず1枠の空表を確保して False を返す)。
    Dim a2 As String, a3 As String
    If n = RK_N_MIN Then a2 = a Else a3 = a
    ok2 = BuildGramIndex(a2, RK_N_MIN, key2, rem2, mask2)
    ok3 = BuildGramIndex(a3, RK_N_MAX, key3, rem3, mask3)
    If Not (ok2 Or ok3) Then Exit Function

    CharCodes b, cb
    NgramOverlap = Overlap23(key2, rem2, mask2, ok2, key3, rem3, mask3, ok3, cb, Len(b))
End Function

' ============================================================================
' RankRows - caseText との関連度(2字+3字のn-gram重なり数の合計)が高い順に
'   rowTexts の**位置(1始まり)**を order() へ並べる(安定ソート。同点は元の
'   並び順を保つ)。rowTexts は 1 To N の1次元配列を渡すこと。
'   戻り値: 並べた件数(rowTexts の要素数)。rowTexts が空/未初期化なら 0。
'
'   案件側(caseText)の索引は2字・3字とも**1回だけ**作り、全行で使い回す
'   (裁定書40 P-M3(a))。候補行数を N としても索引の構築は2回で固定であり、
'   N に比例しない(IndexBuilds() で確認できる)。
' ============================================================================
Public Function RankRows(ByVal caseText As String, ByRef rowTexts() As String, _
                          ByRef order() As Long) As Long
    On Error GoTo Empty0
    Dim lo As Long, hi As Long, cnt As Long
    lo = LBound(rowTexts)
    hi = UBound(rowTexts)
    cnt = hi - lo + 1
    If cnt <= 0 Then GoTo Empty0

    Dim scores() As Long
    Dim idx() As Long
    ReDim scores(1 To cnt)
    ReDim idx(1 To cnt)

    Dim i As Long
    For i = 1 To cnt
        idx(i) = i
    Next i

    Dim key2() As Double, rem2() As Long, mask2 As Long, ok2 As Boolean
    Dim key3() As Double, rem3() As Long, mask3 As Long, ok3 As Boolean
    Dim cb() As Long, lnB As Long
    ok2 = BuildGramIndex(caseText, RK_N_MIN, key2, rem2, mask2)
    ok3 = BuildGramIndex(caseText, RK_N_MAX, key3, rem3, mask3)
    If ok2 Or ok3 Then
        For i = 1 To cnt
            lnB = Len(rowTexts(lo + i - 1))
            If lnB >= RK_N_MIN Then
                CharCodes rowTexts(lo + i - 1), cb
                scores(i) = Overlap23(key2, rem2, mask2, ok2, _
                                      key3, rem3, mask3, ok3, cb, lnB)
            End If
        Next i
    End If

    ' 安定な挿入ソート(降順)。KB1業種あたり数十〜百行程度を想定(伝書鳩3-3)。
    Dim j As Long, keyScore As Long, keyIdx As Long
    For i = 2 To cnt
        keyScore = scores(i)
        keyIdx = idx(i)
        j = i - 1
        Do While j >= 1
            If scores(j) < keyScore Then
                scores(j + 1) = scores(j)
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        scores(j + 1) = keyScore
        idx(j + 1) = keyIdx
    Next i

    ReDim order(1 To cnt)
    For i = 1 To cnt
        order(i) = idx(i)
    Next i
    RankRows = cnt
    Exit Function
Empty0:
    RankRows = 0
End Function

' ============================================================================
' IndexBuilds / ResetIndexBuilds - 索引を作った回数(累計)。**計測専用**で、
'   重なり数の判定には一切使わない。17章§4-1 層(a) の回帰網が「索引を行ごとに
'   作り直していない」ことを機械で押さえるための唯一の口(裁定書40 P-M3(c))。
'   時間を測るテストは環境で揺れるので、回数で固定する。
' ============================================================================
Public Function IndexBuilds() As Long
    IndexBuilds = mIdxBuilds
End Function

Public Sub ResetIndexBuilds()
    mIdxBuilds = 0
End Sub

' ============================================================================
' BuildGramIndex - a の n-gram を開番地法のハッシュ表へ積む(グラムごとの個数)。
'   gKey=グラムの数値キー(-1=空き) / gRem=残り個数(a に現れた個数。
'   Overlap23 が数えるあいだ一時的に減り、走査の最後に必ず元へ戻る) /
'   mask=表の大きさ-1。
'   戻り値 False = 索引を作れない(n が範囲外、または a が n文字未満)。
'   **作れないときも空の表を確保して返す**(呼出側が ByRef の未確保配列を
'   渡したまま Overlap23 に入っても安全にするため)。
' ============================================================================
Private Function BuildGramIndex(ByVal a As String, ByVal n As Long, _
                                ByRef gKey() As Double, ByRef gRem() As Long, _
                                ByRef mask As Long) As Boolean
    Dim la As Long, size As Long, i As Long, slot As Long, cOut As Long
    Dim code As Double, powHi As Double
    Dim hp As Long, powP As Long
    Dim ca() As Long

    ReDim gKey(0 To 0)
    ReDim gRem(0 To 0)
    gKey(0) = -1
    mask = 0
    If n <= 0 Or n > RK_N_MAX Then Exit Function
    la = Len(a) - n + 1
    If la < 1 Then Exit Function

    ' 表は候補数の2倍以上の2のべき乗(負荷率0.5以下。開番地法の探査を短く保つ)。
    size = 16
    Do While size < la * 2
        size = size * 2
    Loop
    ReDim gKey(0 To size - 1)
    ReDim gRem(0 To size - 1)
    mask = size - 1
    mIdxBuilds = mIdxBuilds + 1
    ' 空き札は -1(グラムのコードは 0 以上)。個数の配列を空き判定に使わないのは、
    '   走査中に「残り0」と「空き」が区別できなくなるため。
    For i = 0 To size - 1
        gKey(i) = -1
    Next i

    CharCodes a, ca
    powHi = 1
    powP = 1
    For i = 2 To n
        powHi = powHi * RK_BASE
        powP = powP * RK_MUL
    Next i
    code = 0
    hp = 0
    For i = 1 To n
        code = code * RK_BASE + ca(i)
        hp = hp * RK_MUL + ca(i)
    Next i

    For i = 1 To la
        If i > 1 Then
            cOut = ca(i - 1)
            code = (code - cOut * powHi) * RK_BASE + ca(i + n - 1)
            hp = (hp - cOut * powP) * RK_MUL + ca(i + n - 1)
        End If
        ' 位置の出し方は Overlap23 と**必ず同じ式**にする(片方だけ変えると
        '   索引に入れた場所と探す場所がずれ、静かに重なり数が減る)。
        slot = hp And mask
        Do While gKey(slot) >= 0
            If gKey(slot) = code Then Exit Do
            slot = (slot + 1) And mask
        Loop
        gKey(slot) = code
        gRem(slot) = gRem(slot) + 1
    Next i
    BuildGramIndex = True
End Function

' ============================================================================
' Overlap23 - **索引を受け取る版**。行の文字コード配列を1回だけ走査し、2字と
'   3字の重なり数を同時に数えて合算で返す(use2/use3 の False 側は数えない)。
'   貪欲先勝ち=グラムごとに min(a の個数, b の個数)。使った印は最後に
'   **自分で元へ戻す**ので、同じ索引を次の行へそのまま使い回せる。
'
'   1文字読むごとに2字グラムと3字グラムの両方を作るので、位置あたりの配列参照
'   は1回で済む(LibreOffice Basic は配列1参照が約1.5マイクロ秒あり、ここが
'   総時間を決める。裁定書40 P-M3(c))。
' ============================================================================
Private Function Overlap23(ByRef key2() As Double, ByRef rem2() As Long, _
                           ByVal mask2 As Long, ByVal use2 As Boolean, _
                           ByRef key3() As Double, ByRef rem3() As Long, _
                           ByVal mask3 As Long, ByVal use3 As Boolean, _
                           ByRef cb() As Long, ByVal lnB As Long) As Long
    Dim j As Long, cnt As Long, slot As Long, tn2 As Long, tn3 As Long
    Dim w1 As Long, w2 As Long, w3 As Long, remAt As Long
    Dim code2 As Double, code3 As Double, keyAt As Double
    Dim hp2 As Long, hp3 As Long
    Dim has2 As Boolean, try3 As Boolean
    Dim t2() As Long, t3() As Long

    If lnB < RK_N_MIN Then Exit Function
    If Not (use2 Or use3) Then Exit Function
    ReDim t2(1 To lnB)
    ReDim t3(1 To lnB)

    w2 = cb(1)
    w3 = cb(2)
    For j = 1 To lnB - 1
        w1 = w2
        w2 = w3
        If j + 2 <= lnB Then
            w3 = cb(j + 2)
        Else
            w3 = -1
        End If
        code2 = w1 * RK_BASE + w2
        hp2 = w1 * RK_MUL + w2

        has2 = False
        If use2 Then
            slot = hp2 And mask2
            keyAt = key2(slot)
            Do While keyAt >= 0
                If keyAt = code2 Then Exit Do
                slot = (slot + 1) And mask2
                keyAt = key2(slot)
            Loop
            If keyAt >= 0 Then
                has2 = True
                remAt = rem2(slot)
                If remAt > 0 Then
                    rem2(slot) = remAt - 1
                    cnt = cnt + 1
                    tn2 = tn2 + 1
                    t2(tn2) = slot
                End If
            End If
        End If

        ' 3字の探査を省ける枝(数え方は変えない): 先頭2字が案件本文に**1つも
        '   無い**なら、その2字で始まる3字も案件本文に在りえない。案件と無縁の
        '   行では2字がほぼ全滅するので、ここで探査が半分になる。
        '   「在るが残り0」(remAt=0)は別の位置で3字が一致しうるので省かない。
        try3 = use3
        If try3 And use2 Then try3 = has2
        If try3 And w3 >= 0 Then
            code3 = code2 * RK_BASE + w3
            hp3 = hp2 * RK_MUL + w3
            slot = hp3 And mask3
            keyAt = key3(slot)
            Do While keyAt >= 0
                If keyAt = code3 Then Exit Do
                slot = (slot + 1) And mask3
                keyAt = key3(slot)
            Loop
            If keyAt >= 0 Then
                remAt = rem3(slot)
                If remAt > 0 Then
                    rem3(slot) = remAt - 1
                    cnt = cnt + 1
                    tn3 = tn3 + 1
                    t3(tn3) = slot
                End If
            End If
        End If
    Next j

    ' 借りたぶんを返して索引を呼出前の状態へ戻す(次の行でそのまま使い回す)。
    For j = 1 To tn2
        slot = t2(j)
        rem2(slot) = rem2(slot) + 1
    Next j
    For j = 1 To tn3
        slot = t3(j)
        rem3(slot) = rem3(slot) + 1
    Next j
    Overlap23 = cnt
End Function

' ============================================================================
' WideOverlap - 2字・3字以外(1字と4字以上)の素朴版(二重ループ・貪欲先勝ち)。
'   数値キーが丸められる範囲を索引で数えないための逃げ道であり、本PJの経路
'   (2字/3字)はここを通らない。数え方は Overlap23 と同じ契約(14章§6)。
' ============================================================================
Private Function WideOverlap(ByVal a As String, ByVal b As String, ByVal n As Long) As Long
    Dim la As Long, lb As Long, i As Long, j As Long, cnt As Long
    Dim gramA As String
    Dim used() As Boolean

    If n <= 0 Then Exit Function
    la = Len(a) - n + 1
    lb = Len(b) - n + 1
    If la < 1 Or lb < 1 Then Exit Function

    ReDim used(1 To lb)
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
    WideOverlap = cnt
End Function

' CharCodes - 文字列を1文字ずつのコードポイント配列(1始まり)にする。
'   AscW は符号付き16bitを返すので 65536 を足して正の値に直す
'   (modPii.CodePointOf と同じ作法)。全角英数(U+FF01..)で効く。
Private Sub CharCodes(ByVal s As String, ByRef outCodes() As Long)
    Dim i As Long, ln As Long, v As Long
    ln = Len(s)
    ReDim outCodes(1 To ln + 1)
    For i = 1 To ln
        v = AscW(Mid$(s, i, 1))
        If v < 0 Then v = v + 65536
        outCodes(i) = v
    Next i
End Sub
