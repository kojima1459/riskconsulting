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
    s = s & "var LST={proposed:'仮説',confirmed:'確認済み',rejected:'棄却（記録保持）','new':'新規発見'};" & vbLf
    s = s & "var LHZ={already:'既に顕在化',near:'1～3年',mid_long:'3年超'};" & vbLf
    s = s & "var LFQ={high:'高',mid:'中',low:'低'};" & vbLf
    s = s & "var LIP={large:'大',mid:'中',small:'小'};" & vbLf
    s = s & "var LSRC={hp:'HP',yuho:'有報',memo:'営業メモ',contract:'現契約'," & vbLf
    s = s & "prev_renewal:'前回更新メモ',knowledge:'社内ナレッジ',inference:'推定'};" & vbLf
    s = s & "var LGAP={uninsured:'無保険',underinsured:'過小',overlap:'重複'};" & vbLf
    s = s & "var LPK={upsell:'補償拡大',cross_sell:'新種目提案',scheme:'座組提案'};" & vbLf
    s = s & "var LIQ={ok:'十分',partial:'断片的',missing:'無い'};" & vbLf
    s = s & "var IQCLS={ok:'bdg-ok',partial:'bdg-iqpartial',missing:'bdg-missing'};" & vbLf
    s = s & "var LIQO={high:'高',mid:'中',low:'低'};" & vbLf
    s = s & "var LCT={'new':'新規開拓',renewal:'更新'};" & vbLf
    s = s & "var LTIER={t1_quick:'かんたん調査',t2_full:'しっかり調査',t3_sparring:'壁打ち'};" & vbLf
    s = s & "var LQM={standard:'標準',deep:'入念'};" & vbLf
    s = s & "var ASPORDER=['profile','business','sites','history','news','hr'," & vbLf
    s = s & "'finance_risk','sales_memo','sns','competitors','market','finance'," & vbLf
    s = s & "'insurance_ctx','hazard'];" & vbLf
    s = s & "var LASP={profile:'会社概要',business:'事業・製品',sites:'拠点・設備'," & vbLf
    s = s & "history:'沿革',news:'直近の動き',hr:'採用・人員',finance_risk:'有報・財務リスク'," & vbLf
    s = s & "sales_memo:'営業情報',sns:'SNS評判',competitors:'競合・業界事故'," & vbLf
    s = s & "market:'市況・マクロ',finance:'財務状態',insurance_ctx:'付保・提案の経緯'," & vbLf
    s = s & "hazard:'拠点ハザード'};" & vbLf
    LabelJs = s
End Function
