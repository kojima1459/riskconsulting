Attribute VB_Name = "modTestsPure28"
Option Explicit

' ============================================================================
' modTestsPure28 - 対訳表の設計変更(裁定書43 §1)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' なぜ別モジュールか: modTestsPure27 が 30,000字契約の上限に達しているため
'   (12章§2)。中身は27の G4/G5/G6(対訳表)の続きであり、値源は同じ
'   docs/design/提案書_wide/対訳表_社内語から顧客語.md。
'
' 執筆方針: 期待値は**対訳表§6(終端集合)・§6.1(mode)・§6.5(取り消し規則)の
'   文だけ**から手で書き出した(17章§1。実装の出力を見てから期待値を合わせない)。
'   前波まではここが「D&O保険 → 会社役員賠償責任保険」「PML額 → 想定最大損害額額」
'   のように**壊れた出力を正解として固定**していた(裁定書43 §1-6)。新しい設計では
'   どちらも「直後が漢字なので置換しない」が正解である。
'
' 対象と根拠(**全24本**。本数は modTestRunner.Check の呼び出し数の実測であり、
'   wintest/tests_expected.txt の prod と必ず同時に直すこと):
'   G1 終端集合の文脈では置換する(対訳表§6)
'      01 助詞(を)  02 記号(。)  03 文字列の末尾  04 全角空白
'   G2 終端集合でない文脈では置換しない(型ごとに1本。**列挙ではなく型**)
'      05 ひらがな(活用・サ変)  06 漢字(PML額)  07 カタカナ(ニューリスクリスク)
'      08 英数(BCP2)  09 複合語(D&O保険/BCP計画)=末尾の重なりが出ない
'   G3 mode=warn の対は一切置換しない(対訳表§6.1)
'      10 一般語(抜け)  11 顧客語が述語(未充足)  12 動作性名詞×静的名詞句
'      (ヒアリングを行う。**終端の文脈でも置換しない**のが warn の力)
'      13 用言の連用形(仕分け)
'   G4 最長一致と表記ゆれ(対訳表§6.4)
'      14 引受の方針は置換し、引受けの方針(warn)は置換しない
'      15 座組み(表記ゆれ)は代表形と同じ顧客語
'   G5 対訳表そのもの
'      16 51行・全行に mode がある・warn は10行
'      17 mode の語は replace と warn の2語だけ(印を増やしていない)
'      18 終端の文脈に残った replace の語は TabooHitStrict に出る
'      19 warn と終端でない文脈の残りは TabooHitStrict に出ない(警告には出る)
'      20 冪等(置換した文をもう一度通しても変わらない)
'   G6 新規15対(裁定書46 班F・F-6。modValidate5.TabooPairList)
'      21 てん補期間(終端の文脈で置換)  22 1事故免責金額(免責金額より最長一致で
'      先に1回。ご負担いただく金額が混ざらない)  23 No DD, No cover(warn。
'      置換せずTabooHitに出る)  24 ノンリコース型(リコース型に食われない)
'
' 変異注入(出来レース禁止・裁定書38 §2・43 §1-5):
'   (a) modValidate4 の終端集合から「を」を1字消すと 01 が落ちる
'       (tools/render_proposal.py の check_glossary_impl も同時に赤くなる)。
'   (b) 対訳表の「ヒアリング」の mode を replace にすると 12 と 16 が落ちる
'       (実装の mode と食い違うので check_glossary_impl も赤)。
'   (c) modValidate4.IsTermAt の「末尾なら True」をやめると 03 が落ちる。
'   (d) IsTermAt を常に True にすると 05..09 と 14 が落ちる(=旧設計の壊れ方が
'       そのまま戻る)。
'   (e) SoftenOnce の「終端でない位置は打ち切る」(Exit For)をやめると 14 が
'       落ちる(「引受けの方針」に「引受」が当たり送り仮名が残る)。
'   取り消し規則(§6.5)は終端集合がある限り発火しない二重の安全網なので、
'   純テストからは観測できない。**宣言(対訳表§6.5)と実装(V4_RUN_MAX /
'   V4_UNDO_TAIL_MAX と ReplaceOk が UndoNeeded を呼ぶこと)の突合**を
'   tools/render_proposal.py の check_glossary_impl が受け持つ。
'
' 書き方の約束(LibreOffice Basic 対策): Dim はプロシージャの先頭にまとめ、
'   判定は一度ローカル変数へ入れてから modTestRunner.Check へ渡す。
' ============================================================================

' ============================================================================
' RunAll - modTestRunner.RunAllPureTests から呼ばれる入口。
' ============================================================================
Public Sub RunAll()
    T28Term
    T28NotTerm
    T28Warn
    T28Longest
    T28Table
    T28NewLines46
End Sub

' --- G1 終端集合の文脈では置換する(対訳表§6) ---
Private Sub T28Term()
    Dim n As Long
    Dim after As String
    Dim ok As Boolean

    after = modValidate4.SoftenTaboo("付保を進めます。", n)
    ok = (after = "保険のご加入を進めます。") And (n = 1)
    modTestRunner.Check "W15Y1 終端集合の助詞(を)の直前では置換する", ok, _
                        "after=" & after & " n=" & n

    after = modValidate4.SoftenTaboo("今回の方針は付保。", n)
    ok = (after = "今回の方針は保険のご加入。")
    modTestRunner.Check "W15Y1 終端集合の記号(。)の直前では置換する", ok, _
                        "after=" & after

    ' 「または W が文字列の末尾のとき」(対訳表§6 の後半)。
    after = modValidate4.SoftenTaboo("今回の論点は付保", n)
    ok = (after = "今回の論点は保険のご加入") And (n = 1)
    modTestRunner.Check "W15Y1 社内語が文字列の末尾なら置換する", ok, _
                        "after=" & after & " n=" & n

    after = modValidate4.SoftenTaboo("付保　の検討", n)
    ok = (after = "保険のご加入　の検討")
    modTestRunner.Check "W15Y1 終端集合の全角空白の直前では置換する", ok, _
                        "after=" & after
End Sub

' --- G2 終端集合でない文脈では置換しない(型ごとに1本) ---
Private Sub T28NotTerm()
    Dim n As Long
    Dim after As String
    Dim ok As Boolean

    ' ひらがな(サ変・活用語尾)。語尾を1つも列挙せずに型ごと消える。
    after = modValidate4.SoftenTaboo("付保している拠点", n)
    ok = (after = "付保している拠点") And (n = 0)
    modTestRunner.Check "W15Y1 直後がひらがななら置換しない", ok, _
                        "after=" & after & " n=" & n

    ' 漢字。裁定書43 §1-6: 前波は「想定最大損害額額」を正解に固定していた。
    after = modValidate4.SoftenTaboo("PML額の試算", n)
    ok = (after = "PML額の試算") And (n = 0)
    modTestRunner.Check "W15Y1 直後が漢字なら置換しない(PML額。額額を作らない)", _
                        ok, "after=" & after & " n=" & n

    after = modValidate4.SoftenTaboo("ニューリスクリスクの整理", n)
    ok = (after = "ニューリスクリスクの整理") And (n = 0)
    modTestRunner.Check "W15Y1 直後がカタカナなら置換しない", ok, _
                        "after=" & after & " n=" & n

    after = modValidate4.SoftenTaboo("BCP2の版", n)
    ok = (after = "BCP2の版") And (n = 0)
    modTestRunner.Check "W15Y1 直後が英数なら置換しない", ok, _
                        "after=" & after & " n=" & n

    ' 複合語。終端集合があるので「顧客語の末尾を吸収する」特別規則が要らない。
    after = modValidate4.SoftenTaboo("D&O保険のご案内とBCP計画の策定", n)
    ok = (after = "D&O保険のご案内とBCP計画の策定") And (n = 0)
    modTestRunner.Check "W15Y1 複合語は置換しない(保険保険・計画計画を作らない)", _
                        ok, "after=" & after & " n=" & n
End Sub

' --- G3 mode=warn の対は一切置換しない(対訳表§6.1) ---
Private Sub T28Warn()
    Dim n As Long
    Dim after As String
    Dim ok As Boolean

    after = modValidate4.SoftenTaboo("抜け漏れがないか確認します。", n)
    ok = (after = "抜け漏れがないか確認します。") And (n = 0)
    modTestRunner.Check "W15Y1 warn(一般語)は置換しない", ok, _
                        "after=" & after & " n=" & n

    ' 顧客語が述語の対。旧§6.5 の「保険の手当てが無いの領域」を warn で解消。
    after = modValidate4.SoftenTaboo("未充足の領域があります。", n)
    ok = (after = "未充足の領域があります。") And (n = 0)
    modTestRunner.Check "W15Y1 warn(顧客語が述語)は置換しない", ok, _
                        "after=" & after & " n=" & n

    ' **終端集合の文脈(を)でも**置換しない。ここが mode の効き目。
    after = modValidate4.SoftenTaboo("ヒアリングを行う予定です。", n)
    ok = (after = "ヒアリングを行う予定です。") And (n = 0)
    modTestRunner.Check "W15Y1 warn は終端の文脈でも置換しない(お伺いしたい事項を行う)", _
                        ok, "after=" & after & " n=" & n

    after = modValidate4.SoftenTaboo("仕分けられる項目を仕分けの方法で選ぶ", n)
    ok = (after = "仕分けられる項目を仕分けの方法で選ぶ") And (n = 0)
    modTestRunner.Check "W15Y1 warn(用言の連用形)は置換しない(整理られるを作らない)", _
                        ok, "after=" & after & " n=" & n
End Sub

' --- G4 最長一致と表記ゆれ(対訳表§6.4) ---
Private Sub T28Longest()
    Dim n As Long
    Dim after As String
    Dim ok As Boolean

    ' 「引受」は replace、送り仮名つきの「引受け」は warn(用言の連用形)。
    ' 最長一致で「引受け」を見送った位置では**より短い「引受」も当てない**ので、
    ' 「保険のお引き受けけの方針」は出ない。
    after = modValidate4.SoftenTaboo("引受の方針と引受けの方針", n)
    ok = (after = "保険のお引き受けの方針と引受けの方針") And (n = 1)
    modTestRunner.Check "W15Y1 引受は置換し、引受け(warn)は送り仮名を残さず素通し", _
                        ok, "after=" & after & " n=" & n

    after = modValidate4.SoftenTaboo("見守り型の座組みで進めます。", n)
    ok = (after = "見守り型のご提案の構成で進めます。")
    modTestRunner.Check "W15Y1 表記ゆれ 座組み は代表形と同じ顧客語", ok, _
                        "after=" & after
End Sub

' --- G5 対訳表そのものと TabooHitStrict ---
Private Sub T28Table()
    Dim rows() As String
    Dim onePair() As String
    Dim i As Long
    Dim total As Long
    Dim warnN As String
    Dim modes As String
    Dim mk As String
    Dim ok As Boolean
    Dim n As Long
    Dim a1 As String
    Dim a2 As String
    Dim strictHit As String
    Dim allHit As String

    rows = Split(modValidate4.TabooPairs(), vbLf)
    For i = LBound(rows) To UBound(rows)
        If LenB(rows(i)) > 0 Then
            total = total + 1
            onePair = Split(rows(i), vbTab)
            If UBound(onePair) >= 2 Then
                mk = onePair(2)
                If mk = "warn" Then warnN = warnN & "*"
                If InStr(1, modes, "[" & mk & "]", vbBinaryCompare) = 0 Then
                    modes = modes & "[" & mk & "]"
                End If
            Else
                modes = modes & "[なし]"
            End If
        End If
    Next i
    ok = (total = 66) And (Len(warnN) = 16)
    modTestRunner.Check "W15Y1 対訳表は66行で warn は16行(全行に mode がある)", ok, _
                        "total=" & total & " warn=" & Len(warnN) & " modes=" & modes

    ok = (modes = "[replace][warn]") Or (modes = "[warn][replace]")
    modTestRunner.Check "W15Y1 mode は replace と warn の2語だけ(印を増やさない)", _
                        ok, "modes=" & modes

    ' 終端の文脈に残った replace の語=置換の取りこぼし(実装の欠陥)。
    strictHit = modValidate4.TabooHitStrict("料率を見直します。")
    ok = (InStr(1, strictHit, "料率", vbBinaryCompare) > 0)
    modTestRunner.Check "W15Y1 終端の文脈に残った replace は取りこぼしとして出る", _
                        ok, "strict=[" & strictHit & "]"

    ' warn の語(ヒアリング)と、終端でない文脈の replace(PML額)は欠陥ではない。
    strictHit = modValidate4.TabooHitStrict("ヒアリングを行い、PML額を見ます。")
    allHit = modValidate4.TabooHit("ヒアリングを行い、PML額を見ます。")
    ok = (LenB(strictHit) = 0)
    If ok Then ok = (InStr(1, allHit, "ヒアリング", vbBinaryCompare) > 0)
    If ok Then ok = (InStr(1, allHit, "PML", vbBinaryCompare) > 0)
    modTestRunner.Check "W15Y1 warn と終端外の残りは取りこぼしにしない(警告には出す)", _
                        ok, "strict=[" & strictHit & "] all=[" & allHit & "]"

    a1 = modValidate4.SoftenTaboo("付保の状況とPML額と仕分けの件", n)
    a2 = modValidate4.SoftenTaboo(a1, n)
    ok = (a1 = a2) And (a1 = "保険のご加入の状況とPML額と仕分けの件")
    modTestRunner.Check "W15Y1 置換と見送りを混ぜても冪等", ok, _
                        "a1=" & a1 & " a2=" & a2
End Sub

' --- G6 新規15対(裁定書46 班F・F-6。modValidate5.TabooPairList) ---
Private Sub T28NewLines46()
    Dim n As Long
    Dim after As String
    Dim ok As Boolean
    Dim hit As String

    ' てん補期間(47番・replace)が終端集合の文脈(を)で置換される。
    after = modValidate4.SoftenTaboo("てん補期間を確認する", n)
    ok = (after = "保険金をお支払いする期間を確認する") And (n = 1)
    modTestRunner.Check "W15Y1(F-6) てん補期間は終端の文脈で置換する", ok, _
                        "after=" & after & " n=" & n

    ' 1事故免責金額(54番)が既存の免責金額(19番)より最長一致で先に1回で置換され、
    ' 「免責金額」の顧客語(ご負担いただく金額)が混ざらない。
    after = modValidate4.SoftenTaboo("1事故免責金額は10万円とします。", n)
    ok = (after = "個別の損害ごとにお客さま負担となる金額は10万円とします。") And (n = 1)
    If ok Then ok = (InStr(1, after, "ご負担いただく金額", vbBinaryCompare) = 0)
    modTestRunner.Check "W15Y1(F-6) 1事故免責金額は免責金額より最長一致で先に1回で置換", _
                        ok, "after=" & after & " n=" & n

    ' No DD, No cover(57番・warn)は一切置換せず、TabooHit には出る。
    after = modValidate4.SoftenTaboo("No DD, No coverの原則です。", n)
    hit = modValidate4.TabooHit("No DD, No coverの原則です。")
    ok = (after = "No DD, No coverの原則です。") And (n = 0)
    If ok Then ok = (InStr(1, hit, "No DD, No cover", vbBinaryCompare) > 0)
    modTestRunner.Check "W15Y1(F-6) No DD, No coverは置換せずTabooHitに出る", ok, _
                        "after=" & after & " n=" & n & " hit=[" & hit & "]"

    ' ノンリコース型(53番)がリコース型(52番)の対に食われず正しく置換される。
    after = modValidate4.SoftenTaboo("ノンリコース型で契約する。", n)
    ok = (after = "売主が補償責任を負わない型で契約する。") And (n = 1)
    modTestRunner.Check "W15Y1(F-6) ノンリコース型はリコース型に食われず置換する", ok, _
                        "after=" & after & " n=" & n
End Sub
