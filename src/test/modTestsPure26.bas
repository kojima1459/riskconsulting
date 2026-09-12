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
' 変異注入(出来レース禁止・裁定書38 共通規約):
'   (a) modKnowledgeRank.RankRows の並べ替えを昇順に変えると04/05が落ちる。
'   (b) modPii.SharesLongFragment の判定を `>= minLen - 1` 等へずらすと
'       11(19字)が誤ってTrueになり落ちる。
'
' CP932準拠: 本文・注釈ともにCP932内の文字だけで書く(絵文字不可)。
' ============================================================================

Public Sub RunAll()
    Dim i As Long, grpName As String
    For i = 1 To 4
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
        "実際の順=" & Join(order, ",")
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
