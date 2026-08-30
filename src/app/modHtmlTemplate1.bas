Attribute VB_Name = "modHtmlTemplate1"
Option Explicit

' ==========================================================
' modHtmlTemplate1 - HTMLレポートの骨格・共通CSS・セクション登録表・ランタイム
' ------------------------------------------------
' 正は18章(§4.1 描画の場所 / §4.2 登録表 / §4.4 分割規約 / §5 エスケープと
' テーマ / §6 印刷とブラウザ表示の両立)。14章§6は本モジュールの関数契約を
' 持たない(18章§4.4・§5.2が正)。本モジュールが持つのは§4.4の分割表が定めた
' BuildDocument / HeadHtml / BodyShellHtml / SectionsJs / RuntimeJs の5本。
'
' 設計の要(18章§4.1): DATAはページ内のJSが JSON.parse で受け取り、セクションの
'   描画はブラウザ側のJSが行う。VBAは値ごとのHTML断片を組み立てない。JS側の描画
'   は createElement と textContent への代入のみで行い、innerHTML /
'   insertAdjacentHTML / document.write / outerHTML への代入を使わない(DATA由来
'   の文字列がマークアップとして解釈される経路を構造的に無くす。16章 E-47)。
'   出力は自己完結HTML 1ファイルで外部CSS・JS・フォント・画像・CDNを参照せず、
'   図はCSSだけで描く(18章§1)。
'
' R4準拠(12章§2・18章§1): Worksheets / Range( / Application. / ThisWorkbook /
'   MsgBox / ActiveSheet に触れない純文字列。案件データもconfigも読まない
'   (受け取るのは引数だけ)。CP932内の文字だけで書く(絵文字不可)。
' 組立方式は15章・14章§7と同じ `s = s & "..." & vbLf`(Constは1論理行1,023字と
'   行継続25本の制約に当たるため使わない。18章§4.4)。
' ==========================================================

' coverFields の書式(§4.1(b)の3値を1本で渡す取り決め)。vbTab区切りで [0]=会社名
'   [1]=案件ID [2]=生成日時。vbTab が安全なのは modUtilText.SanitizeInput が
'   制御文字を除去するため(案件データ側に残らない)。
Private Const HT1_SEP As String = vbTab

' BuildDocument - HTML全文を組み立てる(18章§1.1の手順(5))。themeName=解決済み
'   のテーマ名(未知名のフォールバックは呼出側で済ませる) / dataJson=§2のDATA /
'   coverFields=会社名・案件ID・生成日時 を vbTab で連結したもの。
Public Function BuildDocument(ByVal themeName As String, ByVal dataJson As String, _
                              ByVal coverFields As String) As String
    Dim s As String
    s = s & "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja"">" & vbLf
    s = s & HeadHtml(themeName, coverFields) ' SAFE:html
    s = s & BodyShellHtml(coverFields) ' SAFE:html

    ' §5.3(1): DATAは「1本のJS文字列リテラル + JSON.parse」の形でのみ埋める。
    s = s & "<script>var DATA=JSON.parse("""
    s = s & modUtilText.JsStringSafe(dataJson)
    s = s & """);</script>" & vbLf

    s = s & "<script>" & vbLf
    s = s & "(function(){" & vbLf
    s = s & "'use strict';" & vbLf
    s = s & SectionsJs() ' SAFE:html
    s = s & RuntimeJs() ' SAFE:html
    s = s & "})();" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body>" & vbLf
    s = s & "</html>" & vbLf
    BuildDocument = s
End Function

' HeadHtml - <head>。**最初の要素**として <meta charset="utf-8"> を置く(18章
'   §5.3(3)。BOMを見ない設定でも化けないための二重化)。共通CSSは色・書体・本文幅
'   をリテラルで書かず必ず var(--xxx) を通す(例外は #fff のみ。§5.1)。
Public Function HeadHtml(ByVal themeName As String, ByVal coverFields As String) As String
    Dim s As String
    s = s & "<head>" & vbLf
    s = s & "<meta charset=""utf-8"">" & vbLf
    s = s & "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">" & vbLf
    s = s & "<title>"
    s = s & modUtilText.HtmlSafe("リスクレポート " & FieldAt(coverFields, 0))
    s = s & "</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & modHtmlTheme.ThemeCss(themeName) & vbLf ' SAFE:html
    s = s & CommonCss() ' SAFE:html
    s = s & "</style>" & vbLf
    s = s & "</head>" & vbLf
    HeadHtml = s
End Function

' CommonCss - 共通CSS(18章§4.4「HeadHtml に一元化」)。セクション固有の
'   スタイルもクラス名 sec-<slug>-* でここに置く。
Private Function CommonCss() As String
    Dim s As String
    s = s & "*{box-sizing:border-box}html,body{margin:0;padding:0}" & vbLf
    s = s & "body{background:var(--paper);color:var(--ink);font-family:var(--font-sans);font-size:var(--font-size);line-height:var(--line-height);-webkit-print-color-adjust:exact;print-color-adjust:exact}" & vbLf
    s = s & ".wrap{max-width:var(--page-width);margin:0 auto;padding:0 var(--page-pad)}" & vbLf
    s = s & ".topbar{display:flex;align-items:center;justify-content:space-between;padding:10px 0;border-bottom:1px solid var(--line)}" & vbLf
    s = s & ".brand{font-family:var(--font-serif);font-size:12px;letter-spacing:.18em;color:var(--sub)}" & vbLf
    s = s & ".btn{font-family:inherit;font-size:12px;padding:6px 16px;border:1px solid var(--ai);background:var(--paper);color:var(--ai);border-radius:2px;cursor:pointer}" & vbLf
    s = s & ".banner{margin:14px 0 0;padding:9px 14px;background:var(--warn);border:1px solid var(--warn-line);font-size:12.5px}" & vbLf
    s = s & ".banner-h{color:var(--kaki);font-weight:600;margin:0 0 3px}.banner p{margin:0}" & vbLf
    s = s & "h2,h3,h4{font-weight:600}p{margin:0 0 8px}.sec{margin:0 0 30px}" & vbLf
    s = s & ".sec h2{font-family:var(--font-serif);font-size:19px;margin:0 0 12px;padding:0 0 5px;border-bottom:2px solid var(--ai);color:var(--ai)}" & vbLf
    s = s & ".sec h3{font-size:14.5px;margin:16px 0 6px;color:var(--ink)}" & vbLf
    s = s & ".sec h4{font-size:13px;margin:12px 0 4px;color:var(--sub)}" & vbLf
    s = s & ".muted{color:var(--sub);font-size:12.5px}.lead{font-size:14.5px}" & vbLf
    s = s & ".note{border:1px solid var(--warn-line);background:var(--warn);padding:8px 12px;font-size:12.5px;margin:0 0 12px}" & vbLf
    s = s & ".kicker{font-family:var(--font-serif);font-size:12px;letter-spacing:.28em;color:var(--ai);margin:0 0 10px}" & vbLf
    s = s & ".sec-cover{padding:34px 0 20px;border-bottom:1px solid var(--line);margin-bottom:20px}" & vbLf
    s = s & ".sec-cover h1{font-family:var(--font-serif);font-size:27px;margin:0 0 8px;line-height:1.45}" & vbLf
    s = s & ".sec-cover .cv-meta{color:var(--sub);font-size:12.5px;margin:0 0 14px}" & vbLf
    s = s & ".chips{display:flex;flex-wrap:wrap;gap:6px}" & vbLf
    s = s & ".chip{font-size:11.5px;padding:3px 11px;border:1px solid var(--line);background:var(--mist);border-radius:11px;color:var(--sub)}" & vbLf
    s = s & ".chip-ai{border-color:var(--ai);color:var(--ai);background:var(--paper)}" & vbLf
    s = s & ".toc{margin:0 0 28px;padding:12px 16px;background:var(--mist);border:1px solid var(--line)}" & vbLf
    s = s & ".toc-h{font-size:11.5px;letter-spacing:.2em;color:var(--sub);margin:0 0 7px}" & vbLf
    s = s & ".toc ol{margin:0;padding-left:1.5em;columns:2;font-size:12.5px}" & vbLf
    s = s & ".toc li{margin:0 0 3px;break-inside:avoid}.toc a{color:var(--ink);text-decoration:none}" & vbLf
    s = s & ".tblwrap{overflow-x:auto;margin:0 0 12px}" & vbLf
    s = s & "table{border-collapse:collapse;width:100%;font-size:12.5px}" & vbLf
    s = s & "th,td{border:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}" & vbLf
    s = s & "th{background:var(--mist);color:var(--sub);font-weight:600;white-space:nowrap}" & vbLf
    s = s & "tr{break-inside:avoid}td.nw,th.nw{white-space:nowrap}" & vbLf
    s = s & ".bdg{display:inline-block;font-size:11px;padding:1px 8px;border-radius:9px;color:#fff;white-space:nowrap}" & vbLf
    s = s & ".bdg-cover{background:var(--tr-cover)}.bdg-partial{background:var(--tr-partial)}.bdg-hard{background:var(--tr-hard)}" & vbLf
    s = s & ".bdg-ok{background:var(--iq-ok)}.bdg-iqpartial{background:var(--iq-partial)}.bdg-missing{background:var(--iq-missing)}" & vbLf
    s = s & ".bdg-sub{background:var(--sub)}.bdg-ai{background:var(--ai)}" & vbLf
    s = s & ".bdg-kaki{background:var(--kaki)}.bdg-matsu{background:var(--matsu)}.bdg-deep{background:var(--deep)}" & vbLf
    s = s & ".dl{display:grid;grid-template-columns:9em 1fr;gap:5px 14px;font-size:13px;margin:0 0 14px}" & vbLf
    s = s & ".dl .dt{color:var(--sub);font-size:12px;padding-top:2px}.dl .dd{margin:0}" & vbLf
    s = s & ".card{border:1px solid var(--line);border-left:3px solid var(--ai);padding:12px 15px;margin:0 0 12px;break-inside:avoid}" & vbLf
    s = s & ".card h3{margin:0 0 6px}.card-kaki{border-left-color:var(--kaki)}" & vbLf
    s = s & ".card-deep{border-left-color:var(--deep)}.card-matsu{border-left-color:var(--matsu)}" & vbLf
    s = s & ".hook{color:var(--ai);font-size:13px;margin:0 0 8px}" & vbLf
    s = s & ".rowline{display:grid;grid-template-columns:3.2em 1fr 8.6em;gap:8px;align-items:baseline;font-size:13.5px;margin:0 0 5px}" & vbLf
    s = s & ".rank-no{color:var(--ai);font-size:12px}" & vbLf
    s = s & ".sec-riskuniv-row{display:grid;grid-template-columns:11.6em 3.4em 1fr;gap:9px;align-items:center;margin:0 0 5px;font-size:12.5px}" & vbLf
    s = s & ".sec-riskuniv-bar{height:11px;background:var(--mist);border:1px solid var(--line)}" & vbLf
    s = s & ".sec-riskuniv-fill{height:100%;background:var(--ai)}" & vbLf
    s = s & ".sec-riskmap-wrap{break-inside:avoid;margin:0 0 12px}" & vbLf
    s = s & ".sec-riskmap-grid{display:grid;grid-template-columns:2.4em repeat(5,1fr);gap:3px}" & vbLf
    s = s & ".sec-riskmap-cell{position:relative;min-height:64px;border:1px solid var(--line);padding:4px 4px 13px}" & vbLf
    s = s & ".sec-riskmap-band{position:absolute;right:4px;bottom:2px;font-size:10px;color:var(--sub)}" & vbLf
    s = s & ".sec-riskmap-ax{display:flex;align-items:center;justify-content:center;font-size:11px;color:var(--sub);min-height:22px}" & vbLf
    s = s & ".sec-riskmap-axname{font-size:11px;color:var(--sub);margin:4px 0 0}" & vbLf
    s = s & ".heat1{background:var(--heat-1)}.heat2{background:var(--heat-2)}.heat3{background:var(--heat-3)}" & vbLf
    s = s & ".heat4{background:var(--heat-4)}.heat5{background:var(--heat-5)}" & vbLf
    s = s & ".rb{display:inline-block;font-size:10.5px;padding:1px 6px;margin:1px;border-radius:8px;color:#fff}" & vbLf
    s = s & ".rublock{margin:0 0 14px;break-inside:avoid}" & vbLf
    s = s & ".rubh{font-size:13.5px;margin:0 0 6px;padding:3px 9px;background:var(--mist);border-left:3px solid var(--ai)}" & vbLf
    s = s & ".rubh-rej{border-left-color:var(--kaki)}.rubh-new{border-left-color:var(--deep)}" & vbLf
    s = s & ".strike{text-decoration:line-through}" & vbLf
    s = s & ".qlist{margin:0;padding-left:1.7em;font-size:13.5px}" & vbLf
    s = s & ".qlist li{margin:0 0 9px;break-inside:avoid}" & vbLf
    s = s & ".qcat{color:var(--ai);font-size:11px;margin-right:7px}.qwhy{color:var(--sub);font-size:12px}" & vbLf
    s = s & ".disc{font-size:12px;color:var(--sub);border-top:1px solid var(--line);padding-top:10px}.disc p{margin:0 0 5px}" & vbLf
    ' 18章§6: 印刷(A4縦)との両立。@page の余白はCSS変数で解決されないため
    ' リテラルで書く(§5.1・§6(1))。
    s = s & "@page{size:A4 portrait;margin:14mm 12mm;}" & vbLf
    s = s & "@media print{.no-print{display:none}body{font-size:12px;line-height:1.7}" & vbLf
    s = s & ".wrap{max-width:none;padding:0}.tblwrap{overflow:visible}table{font-size:11px}" & vbLf
    s = s & ".sec-exec{break-after:page}a{color:var(--ink);text-decoration:none}" & vbLf
    s = s & ".card,.sec-riskmap-wrap,.rublock,tr,.qlist li{break-inside:avoid}}" & vbLf
    s = s & "@media (max-width:640px){.sec-cover h1{font-size:21px}.sec h2{font-size:17px}" & vbLf
    s = s & ".dl{grid-template-columns:1fr}.toc ol{columns:1}.rowline{grid-template-columns:2.6em 1fr}" & vbLf
    s = s & ".sec-riskuniv-row{grid-template-columns:8.2em 2.8em 1fr}" & vbLf
    s = s & ".sec-riskmap-cell{min-height:46px;padding:2px 2px 12px}.rb{font-size:9px;padding:1px 4px}}" & vbLf
    CommonCss = s
End Function

' BodyShellHtml - 骨格・操作要素・<noscript>・静的な表紙(18章§4.1)。VBAが静的
'   HTMLとして書き出すのは (a)<title> (b)表紙の会社名/案件ID/生成日時
'   (c)<noscript> の3つだけで、いずれも modUtilText.HtmlSafe を通す。
Public Function BodyShellHtml(ByVal coverFields As String) As String
    Dim s As String
    s = s & "<body>" & vbLf
    s = s & "<div class=""wrap"">" & vbLf
    s = s & "<div class=""topbar no-print"">" & vbLf
    s = s & "<span class=""brand"">RISK PROPOSAL NAVI</span>" & vbLf
    s = s & "<button type=""button"" class=""btn"" id=""btnPrint"">印刷する</button>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "<div id=""warnbox""></div>" & vbLf

    ' (c) <noscript>: スクリプトが動かない環境でもAI利用の明示だけは必ず読める
    '     ようにする(18章§4.1)。§3.5の1行目を静的HTMLで書く。
    s = s & "<noscript><div class=""note"">" & vbLf
    s = s & "<p>このレポートの表示にはJavaScriptが必要です。"
    s = s & "ファイルをローカルに保存してブラウザで開いてください。</p>" & vbLf
    s = s & "<p>本資料はAI支援により作成した骨子を人が確認・編集したものです。</p>" & vbLf
    s = s & "</div></noscript>" & vbLf

    s = s & "<main id=""doc"">" & vbLf
    ' (b) 表紙。JSが動かなくても「誰の・いつの・どの案件の資料か」は読める。
    s = s & "<section id=""sec-cover"" class=""sec sec-cover"">" & vbLf
    s = s & "<p class=""kicker"">RISK REPORT</p>" & vbLf
    s = s & "<h1>"
    s = s & modUtilText.HtmlSafe(FieldAt(coverFields, 0))
    s = s & "</h1>" & vbLf
    s = s & "<p class=""cv-meta"">案件ID "
    s = s & modUtilText.HtmlSafe(FieldAt(coverFields, 1))
    s = s & " / 作成 "
    s = s & modUtilText.HtmlSafe(FieldAt(coverFields, 2))
    s = s & "</p>" & vbLf
    s = s & "</section>" & vbLf
    s = s & "<nav id=""toc"" class=""toc""></nav>" & vbLf
    s = s & "</main>" & vbLf
    s = s & "</div>" & vbLf
    BodyShellHtml = s
End Function

' SectionsJs - セクション登録表(18章§4.2)。**編集が最も多い1関数**。(a)描画関数
'   の連結行 と (b)登録配列 の2行1組で1セクションを表す。登録行のキーは
'   id / slug / title / need / empty / note / render の7つに固定。並び順は18章§3
'   の表の上から下(10章FR-37の紙面順)。`empty` は§3の「空のときの挙動」列と
'   1対1で対応させる(§3が正)。
Public Function SectionsJs() As String
    Dim s As String
    ' (a) 描画関数の連結行
    s = s & modHtmlTemplate2.SecCoverJs() ' SAFE:html
    s = s & modHtmlTemplate2.SecExecJs() ' SAFE:html
    s = s & modHtmlTemplate2.SecProfileJs() ' SAFE:html
    s = s & modHtmlTemplate2.SecSufficiencyJs() ' SAFE:html
    s = s & modHtmlTemplate3.SecRiskUnivJs() ' SAFE:html
    s = s & modHtmlTemplate3.SecRiskMapJs() ' SAFE:html
    s = s & modHtmlTemplate3.SecRisksJs() ' SAFE:html
    s = s & modHtmlTemplate3.SecCoverageJs() ' SAFE:html
    s = s & modHtmlTemplate4.SecNewRiskJs() ' SAFE:html
    s = s & modHtmlTemplate4.SecRoundUpdateJs() ' SAFE:html
    s = s & modHtmlTemplate4.SecStoryJs() ' SAFE:html
    s = s & modHtmlTemplate5.SecPreventJs() ' SAFE:html
    s = s & modHtmlTemplate5.SecLimitJs() ' SAFE:html
    s = s & modHtmlTemplate5.SecHearingJs() ' SAFE:html
    s = s & modHtmlTemplate5.SecSourceJs() ' SAFE:html
    s = s & modHtmlTemplate5.SecDisclaimerJs() ' SAFE:html

    ' (b) 登録配列
    s = s & "var SECTIONS=[" & vbLf
    s = s & "{id:'SEC-01',slug:'cover',title:'',need:['meta'],empty:'always',render:renderCover}," & vbLf
    s = s & "{id:'SEC-02',slug:'exec',title:'エグゼクティブサマリ',need:['s1'],empty:'always',render:renderExec}," & vbLf
    s = s & "{id:'SEC-03',slug:'profile',title:'企業理解',need:['s1'],empty:'hide',render:renderProfile}," & vbLf
    s = s & "{id:'SEC-04',slug:'sufficiency',title:'入力の充足度と要確認事項',need:['s1'],empty:'hide',render:renderSufficiency}," & vbLf
    s = s & "{id:'SEC-05',slug:'riskuniv',title:'リスクユニバース10分類',need:['s2'],empty:'hide',render:renderRiskUniv}," & vbLf
    s = s & "{id:'SEC-06',slug:'riskmap',title:'2軸リスクマップ',need:['s2'],empty:'hide',render:renderRiskMap}," & vbLf
    s = s & "{id:'SEC-07',slug:'risks',title:'リスク一覧',need:['s2'],empty:'hide',render:renderRisks}," & vbLf
    s = s & "{id:'SEC-08',slug:'coverage',title:'保険カバレッジ表',need:['s1'],empty:'hide',render:renderCoverage}," & vbLf
    s = s & "{id:'SEC-09',slug:'newrisk',title:'ニューリスク',need:['s2'],empty:'note'," & vbLf
    s = s & "note:'現時点で特筆すべきニューリスクは検出されていません',render:renderNewRisk}," & vbLf
    s = s & "{id:'SEC-16',slug:'round-update',title:'訪問で分かったこと',need:['s2'],empty:'hide',render:renderRoundUpdate}," & vbLf
    s = s & "{id:'SEC-10',slug:'story',title:'提案ストーリー（当社にできること）',need:['s3'],empty:'hide',render:renderStory}," & vbLf
    s = s & "{id:'SEC-11',slug:'prevent',title:'未然防止メニュー',need:['s2'],empty:'hide',render:renderPrevent}," & vbLf
    s = s & "{id:'SEC-12',slug:'limit',title:'当社にできないこと・提案を控えること',need:['s2'],empty:'note'," & vbLf
    s = s & "note:'該当なし',render:renderLimit}," & vbLf
    s = s & "{id:'SEC-13',slug:'hearing',title:'ヒアリング事項',need:['s1'],empty:'hide',render:renderHearing}," & vbLf
    s = s & "{id:'SEC-14',slug:'source',title:'出典と根拠',need:['s2'],empty:'hide',render:renderSource}," & vbLf
    s = s & "{id:'SEC-15',slug:'disclaimer',title:'免責とご確認事項',need:['meta'],empty:'always',render:renderDisclaimer}" & vbLf
    s = s & "];" & vbLf
    SectionsJs = s
End Function

' RuntimeJs - 共通の描画ヘルパ・enum変換表・登録配列の走査・目次生成。描画は
'   createElement / textContent / setAttribute だけで行う(18章§4.1)。末尾で
'   run() を呼ぶ。関数宣言は巻き上げられるため SectionsJs が先でも問題ない。
Public Function RuntimeJs() As String
    Dim s As String
    s = s & LabelJs() ' SAFE:html
    s = s & HelperJs() ' SAFE:html
    s = s & DriverJs() ' SAFE:html
    RuntimeJs = s
End Function

' 19章§3のenum変換表(機械値 -> 日本語ラベル)。生の英字enumを画面に出さない
' (18章§3)。並びは15章§0の変換表の記載順に固定(§3.2)。
Private Function LabelJs() As String
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
    s = s & "var LIQO={high:'充足度 高',mid:'充足度 中',low:'充足度 低'};" & vbLf
    s = s & "var LCT={'new':'新規開拓',renewal:'更新'};" & vbLf
    s = s & "var LTIER={t1_quick:'クイック',t2_full:'フルドシエ',t3_sparring:'壁打ち'};" & vbLf
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

' 共通の描画ヘルパ。innerHTML系は使わない(18章§4.1)。
Private Function HelperJs() As String
    Dim s As String
    s = s & "function E(tag,cls){var n=document.createElement(tag);" & vbLf
    s = s & "if(cls){n.className=cls;}return n;}" & vbLf
    s = s & "function T(p,tag,cls,text){var n=E(tag,cls);" & vbLf
    s = s & "if(text!==undefined&&text!==null){n.textContent=String(text);}" & vbLf
    s = s & "p.appendChild(n);return n;}" & vbLf
    s = s & "function AT(n,k,v){n.setAttribute(k,v);return n;}" & vbLf
    s = s & "function S(v){return (v===undefined||v===null)?'':String(v);}" & vbLf
    s = s & "function NB(v){return S(v).replace(/^\s+|\s+$/g,'').length>0;}" & vbLf
    s = s & "function AR(v){return (v&&v.length)?v:[];}" & vbLf
    s = s & "function LB(map,k){var t=S(k);return map[t]?map[t]:t;}" & vbLf
    s = s & "function CLIP(t,n){var x=S(t);" & vbLf
    s = s & "return (x.length>n)?(x.slice(0,n)+'…'):x;}" & vbLf
    s = s & "function PARA(p,text,cls){var a=S(text).split(/\r\n|\r|\n/);" & vbLf
    s = s & "for(var i=0;i<a.length;i++){if(NB(a[i])){T(p,'p',cls,a[i]);}}}" & vbLf
    s = s & "function BDG(p,cls,text){return T(p,'span','bdg '+cls,text);}" & vbLf
    s = s & "function CHIP(p,text,cls){return T(p,'span','chip'+(cls?' '+cls:''),text);}" & vbLf
    s = s & "function TBL(host,heads,rows){var w=T(host,'div','tblwrap');" & vbLf
    s = s & "var tb=T(w,'table');var hr=T(T(tb,'thead'),'tr');" & vbLf
    s = s & "for(var i=0;i<heads.length;i++){T(hr,'th',null,heads[i]);}" & vbLf
    s = s & "var bd=T(tb,'tbody');" & vbLf
    s = s & "for(var r=0;r<rows.length;r++){var tr=T(bd,'tr');" & vbLf
    s = s & "for(var c=0;c<rows[r].length;c++){var v=rows[r][c];" & vbLf
    s = s & "if(v&&v.nodeType===1){T(tr,'td').appendChild(v);}" & vbLf
    s = s & "else{T(tr,'td',null,S(v));}}}return tb;}" & vbLf
    s = s & "function DLIST(host,pairs){var d=T(host,'div','dl');" & vbLf
    s = s & "for(var i=0;i<pairs.length;i++){if(!NB(pairs[i][1])){continue;}" & vbLf
    s = s & "T(d,'div','dt',pairs[i][0]);T(d,'div','dd',pairs[i][1]);}" & vbLf
    s = s & "if(d.childNodes.length===0){d.parentNode.removeChild(d);}return d;}" & vbLf
    s = s & "function JOIN(a,sep){var o=[];a=AR(a);" & vbLf
    s = s & "for(var i=0;i<a.length;i++){if(NB(a[i])){o.push(S(a[i]));}}return o.join(sep);}" & vbLf
    s = s & "function RISKS(D){return (D.s2&&D.s2.risks)?D.s2.risks:[];}" & vbLf
    s = s & "function RNAME(D,no){var a=RISKS(D);" & vbLf
    s = s & "for(var i=0;i<a.length;i++){if(a[i].risk_no===no){return S(a[i].risk_name);}}" & vbLf
    s = s & "return 'No.'+S(no);}" & vbLf
    s = s & "function GNAME(D,no){var a=(D.s2&&D.s2.gaps)?D.s2.gaps:[];" & vbLf
    s = s & "for(var i=0;i<a.length;i++){if(a[i].gap_no===no){return S(a[i].target);}}" & vbLf
    s = s & "return 'Gap.'+S(no);}" & vbLf
    ' 18章§3.1(2)の並び: impact*frequency の降順 -> impact の降順 -> risk_no の昇順。
    s = s & "function RANKED(D){var a=RISKS(D).slice(0);" & vbLf
    s = s & "a.sort(function(x,y){var px=(x.impact_score||0)*(x.frequency_score||0);" & vbLf
    s = s & "var py=(y.impact_score||0)*(y.frequency_score||0);" & vbLf
    s = s & "if(px!==py){return py-px;}" & vbLf
    s = s & "if((y.impact_score||0)!==(x.impact_score||0)){" & vbLf
    s = s & "return (y.impact_score||0)-(x.impact_score||0);}" & vbLf
    s = s & "return (x.risk_no||0)-(y.risk_no||0);});return a;}" & vbLf
    HelperJs = s
End Function

' 登録配列の走査・need/empty 判定・目次・警告バナー・印刷ボタン。
Private Function DriverJs() As String
    Dim s As String
    s = s & "function needOk(D,need){for(var i=0;i<need.length;i++){" & vbLf
    s = s & "if(!D[need[i]]){return false;}}return true;}" & vbLf
    ' 18章§4.2の empty は always / hide / note の閉じた3値。登録表に無い値を
    ' 黙って hide 扱いにするとセクションが理由なく消えるため always へ寄せる
    ' (「見ていない」と「見たが該当なし」を読み手が区別できる側=§3.2・§3の SEC-09
    ' と同じ判断)。3値の分岐をこの1関数に集約する。
    s = s & "function emptyMode(d){var m=S(d.empty);" & vbLf
    s = s & "if(m==='hide'){return 'hide';}if(m==='note'){return 'note';}" & vbLf
    s = s & "return 'always';}" & vbLf
    s = s & "function banner(D){var host=document.getElementById('warnbox');" & vbLf
    s = s & "var w=(D.meta&&D.meta.warnings)?D.meta.warnings:[];" & vbLf
    s = s & "if(!host||!w.length){return;}var b=T(host,'div','banner');" & vbLf
    s = s & "T(b,'p','banner-h','ご確認ください');" & vbLf
    s = s & "for(var i=0;i<w.length;i++){T(b,'p',null,w[i]);}}" & vbLf
    s = s & "function toc(list){var host=document.getElementById('toc');" & vbLf
    s = s & "if(!host){return;}var any=false;" & vbLf
    s = s & "T(host,'div','toc-h','目次');var ol=T(host,'ol');" & vbLf
    s = s & "for(var i=0;i<list.length;i++){var d=list[i];if(!d.title){continue;}" & vbLf
    s = s & "any=true;var a=T(T(ol,'li'),'a',null,d.title);" & vbLf
    s = s & "AT(a,'href','#sec-'+d.slug);}" & vbLf
    s = s & "if(!any){host.parentNode.removeChild(host);}}" & vbLf
    s = s & "function run(){var D=DATA;var doc=document.getElementById('doc');" & vbLf
    s = s & "if(!doc){return;}banner(D);var shown=[];" & vbLf
    s = s & "for(var i=0;i<SECTIONS.length;i++){var d=SECTIONS[i];" & vbLf
    s = s & "var host=document.getElementById('sec-'+d.slug);var fixed=!!host;" & vbLf
    s = s & "if(!host){host=E('section','sec sec-'+d.slug);AT(host,'id','sec-'+d.slug);}" & vbLf
    s = s & "var body=E('div',null);" & vbLf
    s = s & "if(needOk(D,d.need)){try{d.render(D,body);}" & vbLf
    s = s & "catch(err){T(body,'p','note','この節の描画に失敗しました。');}}" & vbLf
    s = s & "var isEmpty=(body.childNodes.length===0);var em=emptyMode(d);" & vbLf
    s = s & "if(isEmpty&&em==='hide'){" & vbLf
    s = s & "if(fixed&&host.parentNode){host.parentNode.removeChild(host);}continue;}" & vbLf
    s = s & "if(!fixed&&d.title){T(host,'h2',null,d.title);}" & vbLf
    s = s & "if(isEmpty){T(host,'p',(em==='note')?'note':'muted'," & vbLf
    s = s & "(em==='note'&&d.note)?d.note:'表示できるデータがありません。');}" & vbLf
    s = s & "else{while(body.firstChild){host.appendChild(body.firstChild);}}" & vbLf
    s = s & "if(!fixed){doc.appendChild(host);}shown.push(d);}" & vbLf
    s = s & "toc(shown);" & vbLf
    s = s & "var b=document.getElementById('btnPrint');" & vbLf
    s = s & "if(b){b.addEventListener('click',function(){window.print();});}}" & vbLf
    s = s & "run();" & vbLf
    DriverJs = s
End Function

' coverFields の idx 番目。範囲外は空文字列。
Private Function FieldAt(ByVal coverFields As String, ByVal idx As Long) As String
    Dim parts() As String
    parts = Split(coverFields, HT1_SEP)
    If idx < LBound(parts) Or idx > UBound(parts) Then Exit Function
    FieldAt = parts(idx)
End Function
