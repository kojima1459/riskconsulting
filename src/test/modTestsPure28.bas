Attribute VB_Name = "modTestsPure28"
Option Explicit

' ============================================================================
' modTestsPure28 - W15 最終是正(裁定書42 §1)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' なぜ別モジュールか: modTestsPure27 が 30,000字契約の上限に達しているため
'   (12章§2)。中身は27の G4/G5/G6(対訳表)の続きであり、値源は同じ
'   docs/design/提案書_wide/対訳表_社内語から顧客語.md。
'
' 執筆方針: 期待値は**対訳表 §6.3 / §6.4 / §6.6 の文と統合レビュー(班W3)が
'   実機で再現した壊れ方**だけから手で書き出した(17章§1。実装の出力を見てから
'   期待値を合わせない)。
'
' 対象と根拠(**全17本**。本数は modTestRunner.Check の呼び出し数の実測であり、
'   wintest/tests_expected.txt の prod と必ず同時に直すこと):
'   G1 用言の連用形(対訳表§6.3。統合レビュー「仕分けている→整理ている」)
'      01 活用語尾が続く位置では置換しない(ている)
'      02 同(ました)  03 同(られる)
'      04 名詞の位置(の)は従来どおり置換する
'      05 活用ではない「など」は置換する(1文字判定の取りこぼし防止)
'      06 「仕分けする」は置換する(顧客語+する が成立するので印を付けない語)
'   G2 表記ゆれ(対訳表§6.4。統合レビュー「引受けている→保険のお引き受けけている」)
'      07 「引受けている」は置換しない=「保険のお引き受け」を1文字も出さない
'      08 「引受けの方針」は置換する
'      09 「引受する」はサ変印で見送る(裁定書40 S-M2 の据え置き確認)
'      10 「引き受ける」は対訳表に無いので素通し(ふつうの日本語を壊さない)
'      11 「座組み」は代表形と同じ顧客語へ寄る
'   G3 末尾の重なりの吸収(対訳表§6.6。統合レビュー「D&O保険→…保険保険」)
'      12 D&O保険  13 BCP計画  14 MFA認証
'      15 重なりが1文字なら吸収しない(PML額。2文字以上という規約の境界)
'   G4 対訳表そのもの
'      16 TabooPairs は51行で、印つきは15行(general3/suru10/verb2)
'      17 上の代表2件は冪等(2回通しても変わらない)
'
' 変異注入(出来レース禁止・裁定書38 §2):
'   (a) modValidate4 の「仕分け」から verb 印を外すと 01..03 が落ちる。
'   (b) V4_VERB_TAILS から「ない」「ます」を1文字の「な」「ま」に戻すと
'       05 が落ちる(「仕分けなど」を取りこぼす)。
'   (c) TabooPairs から「引受け」の行を消すと 07 が落ちる。
'   (d) SoftenOnce の「印で見送った位置は打ち切る」を元の Exit For 無しへ
'       戻すと 07 が落ちる(「保険のお引き受けけている」が出る)。
'   (e) SoftenOnce から TailOverlap の1行を消すと 12..14 が落ちる。
'   (f) V4_TAIL_MIN を 1 にすると 15 が落ちる。
'   いずれも tools/render_proposal.py の check_glossary_impl /
'   check_glossary_effective も同時に赤くなる(層を2つ持つ)。
'
' 書き方の約束(LibreOffice Basic 対策): Dim はプロシージャの先頭にまとめ、
'   判定は一度ローカル変数へ入れてから modTestRunner.Check へ渡す。
' ============================================================================

' ============================================================================
' RunAll - modTestRunner.RunAllPureTests から呼ばれる入口。
' ============================================================================
Public Sub RunAll()
    T28Verb
    T28Okurigana
    T28TailOverlap
    T28Table
End Sub

' --- G1 用言の連用形(対訳表§6.3) ---
Private Sub T28Verb()
    Dim n As Long
    Dim after As String
    Dim ok As Boolean

    after = modValidate4.SoftenTaboo("リスクを仕分けている。", n)
    ok = (after = "リスクを仕分けている。") And (n = 0)
    modTestRunner.Check "W15X1 verb 活用語尾(ている)では置換しない", ok, _
                        "after=" & after & " n=" & n

    after = modValidate4.SoftenTaboo("課題を仕分けました。", n)
    ok = (after = "課題を仕分けました。")
    modTestRunner.Check "W15X1 verb 活用語尾(ました)では置換しない", ok, _
                        "after=" & after

    after = modValidate4.SoftenTaboo("仕分けられる項目", n)
    ok = (after = "仕分けられる項目")
    modTestRunner.Check "W15X1 verb 活用語尾(られる)では置換しない", ok, _
                        "after=" & after

    after = modValidate4.SoftenTaboo("仕分けの方法", n)
    ok = (after = "整理の方法") And (n = 1)
    modTestRunner.Check "W15X1 verb 名詞の位置は置換する", ok, _
                        "after=" & after & " n=" & n

    after = modValidate4.SoftenTaboo("仕分けなどの作業", n)
    ok = (after = "整理などの作業")
    modTestRunner.Check "W15X1 verb 活用でない「など」は置換する", ok, _
                        "after=" & after

    after = modValidate4.SoftenTaboo("リスクを仕分けする。", n)
    ok = (after = "リスクを整理する。")
    modTestRunner.Check "W15X1 verb サ変語尾では見送らない(整理するは成立)", ok, _
                        "after=" & after
End Sub

' --- G2 表記ゆれ(対訳表§6.4) ---
Private Sub T28Okurigana()
    Dim n As Long
    Dim after As String
    Dim ok As Boolean

    after = modValidate4.SoftenTaboo("保険を引受けている。", n)
    ok = (after = "保険を引受けている。")
    ok = ok And (InStr(1, after, "保険のお引き受け", vbBinaryCompare) = 0)
    modTestRunner.Check "W15X1 表記ゆれ 引受けている は置換しない", ok, _
                        "after=" & after

    after = modValidate4.SoftenTaboo("引受けの方針", n)
    ok = (after = "保険のお引き受けの方針")
    modTestRunner.Check "W15X1 表記ゆれ 引受けの は置換する", ok, _
                        "after=" & after

    after = modValidate4.SoftenTaboo("保険を引受する。", n)
    ok = (after = "保険を引受する。")
    modTestRunner.Check "W15X1 サ変印は据え置き(引受する)", ok, "after=" & after

    after = modValidate4.SoftenTaboo("残余損害を保険で引き受ける", n)
    ok = (after = "残余損害を保険で引き受ける") And (n = 0)
    modTestRunner.Check "W15X1 引き受ける は対訳表に無く素通し", ok, _
                        "after=" & after & " n=" & n

    after = modValidate4.SoftenTaboo("見守り型の座組みで進めます。", n)
    ok = (after = "見守り型のご提案の構成で進めます。")
    modTestRunner.Check "W15X1 表記ゆれ 座組み は代表形と同じ顧客語", ok, _
                        "after=" & after
End Sub

' --- G3 末尾の重なりの吸収(対訳表§6.6) ---
Private Sub T28TailOverlap()
    Dim n As Long
    Dim after As String
    Dim ok As Boolean

    after = modValidate4.SoftenTaboo("D&O保険のご案内", n)
    ok = (after = "会社役員賠償責任保険のご案内")
    modTestRunner.Check "W15X1 末尾の重なりを吸収する(D&O保険)", ok, _
                        "after=" & after

    after = modValidate4.SoftenTaboo("BCP計画の策定", n)
    ok = (after = "事業継続計画の策定")
    modTestRunner.Check "W15X1 末尾の重なりを吸収する(BCP計画)", ok, _
                        "after=" & after

    after = modValidate4.SoftenTaboo("MFA認証の導入", n)
    ok = (after = "多要素認証の導入")
    modTestRunner.Check "W15X1 末尾の重なりを吸収する(MFA認証)", ok, _
                        "after=" & after

    ' 規約は「2文字以上の重なりだけ」。「額」1文字は吸収せず本文を残す
    ' (1文字の偶然の一致で顧客向け本文を削らないための境界)。
    after = modValidate4.SoftenTaboo("PML額", n)
    ok = (after = "想定最大損害額額")
    modTestRunner.Check "W15X1 重なりが1文字なら吸収しない(PML額)", ok, _
                        "after=" & after
End Sub

' --- G4 対訳表そのもの ---
Private Sub T28Table()
    Dim rows() As String
    Dim i As Long
    Dim total As Long
    Dim marked As Long
    Dim ok As Boolean
    Dim n As Long
    Dim a1 As String
    Dim a2 As String

    rows = Split(modValidate4.TabooPairs(), vbLf)
    For i = LBound(rows) To UBound(rows)
        If LenB(rows(i)) > 0 Then
            total = total + 1
            If UBound(Split(rows(i), vbTab)) >= 2 Then marked = marked + 1
        End If
    Next i
    ok = (total = 51) And (marked = 15)
    modTestRunner.Check "W15X1 対訳表は51行・印つき15行", ok, _
                        "total=" & total & " marked=" & marked

    a1 = modValidate4.SoftenTaboo("D&O保険と仕分けている件", n)
    a2 = modValidate4.SoftenTaboo(a1, n)
    ok = (a1 = a2) And (a1 = "会社役員賠償責任保険と仕分けている件")
    modTestRunner.Check "W15X1 吸収と見送りを混ぜても冪等", ok, _
                        "a1=" & a1 & " a2=" & a2
End Sub
