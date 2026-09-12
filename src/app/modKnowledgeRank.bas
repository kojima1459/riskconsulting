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
' ============================================================================

' ============================================================================
' NgramOverlap - a と b に共通して現れる n文字グラムの延べ数(多重集合の共通部)。
'   貪欲マッチング(先勝ち)で数える。b 側の同じ位置を二重に使わない。
'   n<=0、または a か b が n文字未満のときは 0。大小文字・かな漢字はそのまま
'   比較する(正規化は呼び出し側の責務。ここは純粋な文字列比較のみ)。
'
' 裁定書39 R1-03(重大・性能): 旧実装は二重ループで O(|a|x|b|) だった。案件本文
'   20,000字 x 行500字 x 200行で Python 実測278秒(VBAはこれが下限)であり、
'   S2・S3・HTMLレポートの計3回掛かるため KB が育つほど Excel が固まる。
'   a 側の n-gram を **Collection で索引化**(キー=グラム・値=残り個数)し、
'   b を1回舐めるだけの O(|a|+|b|) へ置き換えた。
'   数え方は不変: 貪欲先勝ちの結果は「グラムごとの min(a の個数, b の個数)の
'   総和」と一致するため、残り個数を1つずつ減らすだけで同じ値になる
'   (modTestsPure26 が素朴版の参照実装と突き合わせて固定する)。
'
' キーに生のグラム文字列を使わない理由: VBA/Basic の Collection のキー照合は
'   **大小文字を区別しない**(さらにロケールによっては半角全角・かなカナも
'   畳む)。本関数の契約は「そのまま比較する」なので、グラムを UTF-16 の
'   コードポイント16進(1文字=4桁)へ直した文字列をキーにする。
' ============================================================================
Public Function NgramOverlap(ByVal a As String, ByVal b As String, ByVal n As Long) As Long
    If n <= 0 Then Exit Function
    Dim la As Long, lb As Long
    la = Len(a) - n + 1
    lb = Len(b) - n + 1
    If la < 1 Or lb < 1 Then Exit Function

    Dim idx As Collection
    Set idx = New Collection

    Dim i As Long, j As Long, cnt As Long
    Dim keyText As String
    Dim leftN As Long
    Dim hasKey As Boolean

    For i = 1 To la
        keyText = GramKey(a, i, n)
        leftN = KeyCount(idx, keyText, hasKey)
        If hasKey Then idx.Remove keyText
        idx.Add leftN + 1, keyText
    Next i

    For j = 1 To lb
        keyText = GramKey(b, j, n)
        leftN = KeyCount(idx, keyText, hasKey)
        If hasKey Then
            If leftN > 0 Then
                idx.Remove keyText
                idx.Add leftN - 1, keyText
                cnt = cnt + 1
            End If
        End If
    Next j
    NgramOverlap = cnt
End Function

' GramKey - 位置 pos から n文字ぶんのグラムを、コードポイント16進(1文字4桁)の
'   キー文字列へ直す。Collection のキー照合が大小文字を区別しないための措置
'   (上の注記)。AscW は符号付き16bitを返すので 65536 を足して正の値に直す
'   (modPii.CodePointOf と同じ作法)。
Private Function GramKey(ByVal s As String, ByVal pos As Long, ByVal n As Long) As String
    Dim k As String, i As Long, v As Long
    For i = 0 To n - 1
        v = AscW(Mid$(s, pos + i, 1))
        If v < 0 Then v = v + 65536
        k = k & Right$("000" & Hex$(v), 4)
    Next i
    GramKey = k
End Function

' KeyCount - Collection に keyText があればその値(残り個数)を、無ければ 0 を
'   返す存在チェック用ヘルパ。Collection は存在確認の口を持たず、無いキーの
'   Item はエラーになるため、**この1本の中だけ**でエラーを捕まえる
'   (呼び出し側に On Error Resume Next を撒かない)。
Private Function KeyCount(ByRef col As Collection, ByVal keyText As String, _
                          ByRef foundOut As Boolean) As Long
    foundOut = False
    On Error GoTo NoKey
    KeyCount = CLng(col.Item(keyText))
    foundOut = True
    Exit Function
NoKey:
    foundOut = False
    KeyCount = 0
End Function

' ============================================================================
' RankRows - caseText との関連度(2字+3字のn-gram重なり数の合計)が高い順に
'   rowTexts の**位置(1始まり)**を order() へ並べる(安定ソート。同点は元の
'   並び順を保つ)。rowTexts は 1 To N の1次元配列を渡すこと。
'   戻り値: 並べた件数(rowTexts の要素数)。rowTexts が空/未初期化なら 0。
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
        scores(i) = NgramOverlap(caseText, rowTexts(lo + i - 1), 2) + _
                    NgramOverlap(caseText, rowTexts(lo + i - 1), 3)
        idx(i) = i
    Next i

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
