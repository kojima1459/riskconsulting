Attribute VB_Name = "modHtmlTemplate6"
Option Explicit

' ==========================================================
' modHtmlTemplate6 - enum変換表(機械値 -> 日本語ラベル)のJS
' ------------------------------------------------
' 正は19章§3・15章§0の変換表(日本語ラベルの正はそちら。18章§3が「生の英字
' enumを画面に出さない」と定め、その実現手段が本モジュールの1関数)。
'
' 由来(18章§4.4の分割規約): modHtmlTemplate1 が25,000字を超えたため、
'   「25,000字を超えたら次番のモジュールへ**関数単位で**切り出す」に従い
'   LabelJs をここへ移した。関数名は変えていない(Private -> Public のみ)。
'   §4.4の既定割り当て表は modHtmlTemplate1 に BuildDocument / HeadHtml /
'   BodyShellHtml / SectionsJs / RuntimeJs の5本を置くと定めており、変換表は
'   その5本のいずれでもない(RuntimeJs の内部下請け)。編集が最も多い
'   SectionsJs を抱える1番から、19章の改訂でしか動かない表を外すのが
'   分割の筋であるためこれを選んだ。
'
' R4準拠(12章§2・18章§1): Worksheets / Range( / Application. / ThisWorkbook /
'   MsgBox / ActiveSheet に触れない純文字列。案件データもconfigも読まない。
'   CP932内の文字だけで書く(絵文字不可)。
' ==========================================================

' 19章§3のenum変換表(機械値 -> 日本語ラベル)。生の英字enumを画面に出さない
' (18章§3)。並びは15章§0の変換表の記載順に固定(§3.2)。
Public Function LabelJs() As String
    Dim s As String
    s = s & "var CATORDER=['strategy_market','supply_chain','manufacturing_quality'," & vbLf
    s = s & "'sales_customer','facility_bcp','hr_labor','digital_info','legal_regulatory'," & vbLf
    s = s & "'finance_counterparty','brand_social'];" & vbLf
    s = s & "var LCAT={strategy_market:'戦略・市場',supply_chain:'調達・供給網'," & vbLf
    s = s & "manufacturing_quality:'製造・品質',sales_customer:'販売・顧客'," & vbLf
    s = s & "facility_bcp:'施設・自然災害・BCP',hr_labor:'人材・労務'," & vbLf
    s = s & "digital_info:'デジタル・情報',legal_regulatory:'法務・規制'," & vbLf
    s = s & "finance_counterparty:'財務・取引先',brand_social:'ブランド・社会'};" & vbLf
    s = s & "var LTR={cover:'比較的移転しやすい',partial:'条件付き・部分的',hard:'保険化困難'};" & vbLf
    s = s & "var TRCLS={cover:'bdg-cover',partial:'bdg-partial',hard:'bdg-hard'};" & vbLf
    s = s & "var DOTCLS={cover:'dot-cover',partial:'dot-partial',hard:'dot-hard'};" & vbLf
    s = s & "var LST={proposed:'仮説',confirmed:'確認済み',rejected:'棄却（記録保持）','new':'新規発見'};" & vbLf
    s = s & "var LHZ={already:'既に顕在化',near:'1～3年',mid_long:'3年超'};" & vbLf
    ' HZW は SEC-09 の強度バーの長さ(%)。ラベルの正は19章§3の LHZ であり、
    ' こちらは「近いほど長い」を表す表示上の目安(文字でも併記する。18章§6(3))。
    s = s & "var HZW={already:100,near:70,mid_long:40};" & vbLf
    s = s & "var LFQ={high:'高',mid:'中',low:'低'};" & vbLf
    s = s & "var LIP={large:'大',mid:'中',small:'小'};" & vbLf
    s = s & "var LSRC={hp:'HP',yuho:'有報',memo:'営業メモ',contract:'現契約'," & vbLf
    s = s & "prev_renewal:'前回更新メモ',knowledge:'社内ナレッジ',inference:'推定'};" & vbLf
    s = s & "var LGAP={uninsured:'無保険',underinsured:'過小',overlap:'重複'};" & vbLf
    ' 19章§3 current_coverage.certainty(18章§3.8 の現契約表の確度列)。
    s = s & "var LCERT={confirmed:'確認済み',assumed:'見立て'};" & vbLf
    ' 19章§3 financials.source(15章 SchemaS1 の financials.source)。
    s = s & "var LFSRC={yuho:'有報',kessan_kokoku:'決算公告',tdb:'信用調査'," & vbLf
    s = s & "view:'VIEW情報',memo:'営業メモ',unknown:'不明'};" & vbLf
    ' 19章§3 case.outcome(成功事例シートの成否。13章§3.1)。
    s = s & "var LOUT={won:'刺さった',lost:'刺さらなかった'};" & vbLf
    s = s & "var LPK={upsell:'補償拡大',cross_sell:'新種目提案',scheme:'座組提案'};" & vbLf
    ' 19章§3 growth.difficulty(15章 SchemaS3 の growth_ideas[].difficulty)。
    ' DIFO は18章§3.7の並び「difficulty の易しい順」を決める順位表。
    s = s & "var LDIF={low:'低',mid:'中',high:'高'};" & vbLf
    s = s & "var DIFO={low:1,mid:2,high:3};" & vbLf
    s = s & "var LIQ={ok:'十分',partial:'断片的',missing:'無い'};" & vbLf
    s = s & "var IQCLS={ok:'bdg-ok',partial:'bdg-iqpartial',missing:'bdg-missing'};" & vbLf
    s = s & "var LIQO={high:'高',mid:'中',low:'低'};" & vbLf
    ' 19章§3 missing_info.kind(15章 SchemaS1 の missing_info[].kind。裁定書38 B-11)。
    s = s & "var LMK={conflict:'資料間の矛盾',undisclosed:'非開示'," & vbLf
    s = s & "not_found:'未取得',hearing_only:'ヒアリングで確認'};" & vbLf
    s = s & "var LFIT={risk_clue:'リスクの手がかり',relationship:'決裁・人間関係'," & vbLf
    s = s & "competitor:'競合・他社',constraint:'制約・NG',opportunity:'商機',other:'その他'};" & vbLf
    s = s & "var LCT={'new':'新規開拓',renewal:'更新'};" & vbLf
    s = s & "var LTIER={t1_quick:'かんたん調査',t2_full:'しっかり調査',t3_sparring:'商談の予行演習'};" & vbLf
    s = s & "var LQM={standard:'標準',deep:'入念'};" & vbLf
    s = s & "var ASPORDER=['profile','business','sites','history','news','hr'," & vbLf
    s = s & "'finance_risk','sales_memo','sns','competitors','market','finance'," & vbLf
    s = s & "'insurance_ctx','hazard'];" & vbLf
    s = s & "var LASP={profile:'会社概要',business:'事業・製品',sites:'拠点・設備'," & vbLf
    s = s & "history:'沿革',news:'直近の動き',hr:'採用・人員',finance_risk:'有報・財務リスク'," & vbLf
    s = s & "sales_memo:'営業情報',sns:'SNS評判',competitors:'競合・業界事故'," & vbLf
    s = s & "market:'市況・マクロ',finance:'財務状態',insurance_ctx:'付保・提案の経緯'," & vbLf
    s = s & "hazard:'拠点ハザード'};" & vbLf
    s = s & KickJs() ' SAFE:html
    LabelJs = s
End Function

' 18章§3.0 のキッカー(節番号+英字ラベル)。セクション登録表の7キーを増やさない
'   ため、キッカーはセクションIDから引くこの表が持つ(18章§4.2)。**節の先頭
'   セクションにだけ**付けるので、集約される側(SEC-04/08/11/12/15/16/18)は持たない。
Private Function KickJs() As String
    Dim s As String
    s = s & "var KICK={'SEC-02':'01 / Executive Summary'," & vbLf
    s = s & "'SEC-03':'02 / Business Understanding'," & vbLf
    s = s & "'SEC-05':'03 / MECE Risk Universe'," & vbLf
    s = s & "'SEC-06':'04 / Risk Map'," & vbLf
    s = s & "'SEC-07':'05 / Insurance Coverage Matrix'," & vbLf
    s = s & "'SEC-09':'06 / New Risk Radar'," & vbLf
    s = s & "'SEC-17':'07 / Insurance-enabled Growth'," & vbLf
    s = s & "'SEC-10':'08 / Executive Proposal Story'," & vbLf
    s = s & "'SEC-13':'09 / Discovery Questions'," & vbLf
    s = s & "'SEC-14':'10 / Sources and Methodology'};" & vbLf
    KickJs = s
End Function
