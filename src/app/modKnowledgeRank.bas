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
' ============================================================================
Public Function NgramOverlap(ByVal a As String, ByVal b As String, ByVal n As Long) As Long
    If n <= 0 Then Exit Function
    Dim la As Long, lb As Long
    la = Len(a) - n + 1
    lb = Len(b) - n + 1
    If la < 1 Or lb < 1 Then Exit Function

    Dim used() As Boolean
    ReDim used(1 To lb)

    Dim i As Long, j As Long, cnt As Long
    Dim gramA As String
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
    NgramOverlap = cnt
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
