Attribute VB_Name = "modValidate5"
Option Explicit

' ============================================================================
' modValidate5 - 対訳表(社内語→顧客語)の唯一の値源(裁定書46 班F・F-6)
' ----------------------------------------------------------------------------
' なぜ分けたか(容量): modValidate4.bas が29,976字で満杯(12章§2の30,000字
'   契約)。対訳表の新規15対(裁定書46 F-6)を足す余地が無いため、
'   `TabooPairs()` の実体(AdPair の全行)だけをこちらへ移した。
'
' 持つもの:
'   TabooPairList  対訳表(docs/design/提案書_wide/対訳表_社内語から顧客語.md)の
'                  「社内語<TAB>顧客語<TAB>mode」を vbLf 区切りに並べた**唯一の
'                  値源**。旧 modValidate4.TabooPairs() の中身をそのまま移し、
'                  末尾に新規15対(裁定書46 F-6。対訳表§1へ番号47〜61で追記)を
'                  足した。mode は全行にある(replace / warn の2区分だけ)。
'
' 公開APIは変えない: `modValidate4.TabooPairs()` は
'   `TabooPairs = modValidate5.TabooPairList()` の1行委譲になり、
'   modTestsPure27/28・modExportProposal など既存の呼び出し側は無修正で動く。
'   `tools/render_proposal.py` の `check_glossary_impl`(`_impl_pairs` 周辺)は
'   本モジュールの `TabooPairList()` の字面から AdPair 行を読み取る
'   (modValidate4.bas はもう対を持たないため)。`TermChars()` は容量都合とは
'   無関係(終端集合の宣言は変わっていない)なので modValidate4 に残したまま
'   動かしていない。
'
' R4準拠(12章§2): Excelトークン・Application.Run・案件データ参照を持たない
'   純関数モジュール。CP932準拠(15章§0 原則7)。
' ============================================================================

Private Const V5_TAB As String = vbTab
' 第3列 mode(対訳表§6.1。**2区分だけ**。印を増やさない)。
Private Const V5_REPLACE As String = "replace"
Private Const V5_WARN As String = "warn"

' TabooPairList - 対訳表の全対(旧46語+§4.1の3語+§6.4の表記ゆれ2語=51行に、
'   裁定書46 F-6 の新規15対を足した66行)を組み立てる。
Public Function TabooPairList() As String
    Dim s As String
    AdPair s, "付保", "保険のご加入", V5_REPLACE
    AdPair s, "未付保", "保険に入っていない状態", V5_REPLACE
    AdPair s, "付保ギャップ", "保険で手当てできていない部分", V5_REPLACE
    AdPair s, "未充足", "保険の手当てが無い", V5_WARN
    AdPair s, "移転", "保険で備える", V5_WARN
    AdPair s, "保有", "自社で負担する", V5_WARN
    AdPair s, "トリガー", "保険金をお支払いする条件", V5_REPLACE
    AdPair s, "サブリミット", "補償項目ごとのお支払いの上限額", V5_REPLACE
    AdPair s, "待機期間", "補償が始まるまでの期間", V5_REPLACE
    AdPair s, "保険化", "保険での備え方の設計", V5_REPLACE
    AdPair s, "特約開発", "補償内容の新しい設計", V5_REPLACE
    AdPair s, "組成", "仕組みづくり", V5_REPLACE
    AdPair s, "募集スキーム", "ご加入の手続きの流れ", V5_REPLACE
    AdPair s, "料率", "保険料の水準", V5_REPLACE
    AdPair s, "相関損失", "同時に起きる損害", V5_REPLACE
    AdPair s, "引受", "保険のお引き受け", V5_REPLACE
    ' 表記ゆれ(対訳表§6.4)。用言の連用形なので warn(「引受けられる」)。
    AdPair s, "引受け", "保険のお引き受け", V5_WARN
    AdPair s, "過少保険", "補償額が損害に届かない状態", V5_REPLACE
    AdPair s, "抜け", "補償されない部分", V5_WARN
    AdPair s, "免責金額", "ご負担いただく金額", V5_REPLACE
    AdPair s, "支払限度額", "お支払いの上限額", V5_REPLACE
    AdPair s, "リスクユニバース", "リスクの全体像", V5_REPLACE
    AdPair s, "ニューリスク", "新しく生まれているリスク", V5_REPLACE
    AdPair s, "座組", "ご提案の構成", V5_REPLACE
    AdPair s, "座組み", "ご提案の構成", V5_REPLACE
    AdPair s, "ヒアリング", "お伺いしたい事項", V5_WARN
    AdPair s, "提案の核", "ご提案の前提", V5_REPLACE
    AdPair s, "攻めの保険活用", "成長を後押しする保険の活用", V5_REPLACE
    AdPair s, "発散段階", "構想段階", V5_REPLACE
    AdPair s, "実装難度", "実現までの難易度", V5_REPLACE
    AdPair s, "顕在化", "実際に起きること", V5_WARN
    AdPair s, "打ち手", "対策", V5_REPLACE
    AdPair s, "商材", "保険商品", V5_REPLACE
    AdPair s, "リスク移転可能性", "保険での備えやすさ", V5_REPLACE
    AdPair s, "与信", "取引先の支払い能力", V5_WARN
    AdPair s, "座組パターン", "ご提案の型", V5_REPLACE
    AdPair s, "PML", "想定最大損害額", V5_REPLACE
    AdPair s, "CBI", "取引先の被災による損害", V5_REPLACE
    AdPair s, "BI", "事業が止まったことによる利益の減少", V5_REPLACE
    AdPair s, "RTO", "復旧までの目標時間", V5_REPLACE
    AdPair s, "BCP", "事業継続計画", V5_REPLACE
    AdPair s, "OT", "工場の制御システム", V5_REPLACE
    AdPair s, "MFA", "多要素認証", V5_REPLACE
    AdPair s, "EDR", "端末の不審な動きを検知する仕組み", V5_REPLACE
    AdPair s, "KRI", "リスクの予兆指標", V5_REPLACE
    AdPair s, "SLA", "サービス水準の取り決め", V5_REPLACE
    AdPair s, "D&O", "会社役員賠償責任保険", V5_REPLACE
    AdPair s, "PL保険", "生産物賠償責任保険", V5_REPLACE
    AdPair s, "対話の順序", "ご説明の順序", V5_REPLACE
    AdPair s, "クロスセル", "追加でご検討いただける備え", V5_WARN
    AdPair s, "仕分け", "整理", V5_WARN
    ' ---- 裁定書46(班F) F-6: 新規15対(対訳表§1へ番号47〜61で追記) ----
    ' replace(9)。値源: docs/kb/対訳表候補_ニューリスク.md。
    AdPair s, "てん補期間", "保険金をお支払いする期間", V5_REPLACE
    AdPair s, "縮小支払割合", "損害額のうち保険金としてお支払いする割合", V5_REPLACE
    AdPair s, "縮小率", "経過期間に応じて保険金の支払割合を減らす仕組み", V5_REPLACE
    AdPair s, "既発生債権", "保険が始まる前から既にあった売掛金", V5_REPLACE
    AdPair s, "保証委託契約", "お客さまと保証会社との間で交わす契約", V5_REPLACE
    AdPair s, "リコース型", "売主が補償責任を負う型", V5_REPLACE
    AdPair s, "ノンリコース型", "売主が補償責任を負わない型", V5_REPLACE
    AdPair s, "1事故免責金額", "個別の損害ごとにお客さま負担となる金額", V5_REPLACE
    AdPair s, "1証券免責金額", "契約全体を通じてお客さま負担となる金額", V5_REPLACE
    ' warn(6。定型句・法的な精度が置換で落ちる懸念があるため一律置換しない。
    '   対訳表§6.1 に4つの質問を1つ追加した=裁定書46 の質問5)。
    AdPair s, "アームズ・レングス・ルール", "利害関係のない対等な立場どうしでの交渉", V5_WARN
    AdPair s, "No DD, No cover", "十分な調査をしていない部分は保険の対象にしないという原則", V5_WARN
    AdPair s, "デミニミス", "個別の損害ごとにお客さま負担となる金額", V5_WARN
    AdPair s, "エクセス", "一定額を超えた部分だけを補償する免責の設定方法", V5_WARN
    AdPair s, "フランチャイズ免責", "一定額を超えたら全額を補償する免責の設定方法", V5_WARN
    AdPair s, "求償", "保険金をお支払いした後に保険会社が本来の責任者へ請求すること", V5_WARN
    TabooPairList = s
End Function

' TabooPairList の1行を積む(社内語<TAB>顧客語<TAB>mode。行は vbLf 区切り)。
Private Sub AdPair(ByRef acc As String, ByVal w As String, ByVal c As String, _
                   ByVal md As String)
    If LenB(acc) > 0 Then acc = acc & vbLf
    acc = acc & w & V5_TAB & c & V5_TAB & md
End Sub
