Attribute VB_Name = "modHtmlTemplate1"
Option Explicit

' ==========================================================
' modHtmlTemplate1 - HTMLレポートの骨格・共通CSS・セクション登録表・ランタイム
' ------------------------------------------------
' 正は18章(§3.0 / §4.1 / §4.2 / §4.4 / §5 / §6)。14章§6は本モジュールの
' 関数契約を持たない(18章§4.4・§5.2が正)。持つのは§4.4の分割表が定めた
' BuildDocument / HeadHtml / BodyShellHtml / SectionsJs / RuntimeJs の5本。
' enum変換表とキッカー表(LabelJs)は modHtmlTemplate6、見本の部品CSSと SEC-17
'   の描画(PartsCss / SecGrowthJs)は modHtmlTemplate7(§4.4の25,000字規約)。
'
' 設計の要(18章§4.1): DATAはページ内のJSが JSON.parse で受け取り、描画は
'   ブラウザ側のJSが createElement と textContent だけで行う(innerHTML /
'   insertAdjacentHTML / document.write / outerHTML は使わない。16章 E-47)。
'   出力は自己完結HTML 1ファイル(外部CSS・JS・フォント・画像・CDNを参照しない)。
'
' v1.2(裁定書21・11章§3.8): 体裁を出力見本へ寄せた。上部ナビ(sticky)+ヒーロー
'   +角丸カード+臙脂の主色。目次は上部ナビ(id=toc・画面)と本文先頭の目次
'   (id=tocprint・印刷)の2箇所へ、**同じ1本の一覧から**出す(18章§3.6)。
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
'   裁定書37 B-06 で [3]=確認者(meta.reviewed_by) [4]=確認日時(reviewed_at)を
'   足した(<noscript> の免責を SEC-15 と同じ3項分岐にするため)。無い場合は
'   FieldAt が空文字を返すので、旧来の3値だけを渡す呼出も壊れない。
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
'   CSSの呼び口は本関数1本(§4.4)。実体は CommonCss(骨格) と
'   modHtmlTemplate7.PartsCss(見本の部品) の2本で、25,000字規約による分割。
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
    s = s & modHtmlTemplate7.CommonCss() ' SAFE:html
    s = s & modHtmlTemplate7.PartsCss() ' SAFE:html
    s = s & modHtmlTemplate8.TalkCss() ' SAFE:html
    s = s & "</style>" & vbLf
    s = s & "</head>" & vbLf
    HeadHtml = s
End Function


' BodyShellHtml - 骨格・上部ナビ・ヒーロー・操作要素・<noscript>(18章§4.1)。
'   VBAが静的HTMLとして書き出すのは (a)<title> (b)ヒーローの会社名/案件ID/
'   生成日時 (c)<noscript> の3つだけで、いずれも modUtilText.HtmlSafe を通す。
Public Function BodyShellHtml(ByVal coverFields As String) As String
    Dim s As String
    s = s & "<body>" & vbLf
    ' 上部ナビ(中身はランタイムが登録表から作る。18章§3.6)
    s = s & "<nav class=""topbar no-print""><div class=""topbar-inner"" id=""toc""></div></nav>" & vbLf

    ' (b) ヒーロー。JSが動かなくても「誰の・いつの・どの案件の資料か」は読める。
    s = s & "<header class=""hero"">" & vbLf
    s = s & "<div class=""hero-inner"">" & vbLf
    s = s & "<p class=""eyebrow"">RISK PROPOSAL NAVI</p>" & vbLf
    s = s & "<h1>"
    s = s & modUtilText.HtmlSafe(FieldAt(coverFields, 0))
    s = s & "</h1>" & vbLf
    s = s & "<p class=""hero-subtitle"">経営リスク分析・総合提案構想</p>" & vbLf
    s = s & "<div class=""hero-meta"">" & vbLf
    s = s & "<span class=""pill"">案件ID "
    s = s & modUtilText.HtmlSafe(FieldAt(coverFields, 1))
    s = s & "</span>" & vbLf
    s = s & "<span class=""pill"">作成 "
    s = s & modUtilText.HtmlSafe(FieldAt(coverFields, 2))
    s = s & "</span>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "<button type=""button"" class=""btn no-print"" id=""btnPrint"">印刷する</button>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "</header>" & vbLf

    s = s & "<div id=""warnbox""></div>" & vbLf

    ' (c) <noscript>: スクリプトが動かない環境でもAI利用の明示だけは必ず読める
    '     ようにする(18章§4.1)。§3.5の1行目を静的HTMLで書く。
    s = s & "<div class=""wrap""><noscript><div class=""note"">" & vbLf
    s = s & "<p>このレポートの表示にはJavaScriptが必要です。"
    s = s & "ファイルをローカルに保存してブラウザで開いてください。</p>" & vbLf
    ' 裁定書37 B-06: §3.5 の1行目と**同じ3項分岐**を静的HTML側でも行う
    ' (JSが動かない環境で「人が確認済み」と名乗らない)。差込は HtmlSafe を通る。
    s = s & "<p>"
    If LenB(FieldAt(coverFields, 3)) > 0 Then
        s = s & "本資料はAI支援により作成した骨子を担当者が確認・編集したものです"
        s = s & "（確認: " & modUtilText.HtmlSafe(FieldAt(coverFields, 3))
        s = s & " / " & modUtilText.HtmlSafe(FieldAt(coverFields, 4)) & "）。"
    Else
        s = s & "本資料はAIが公開情報等から作成した営業担当者向けの分析資料です"
        s = s & "（AI生成・担当者確認前）。お客さまへ提示する前に、担当者が内容を"
        s = s & "確認・編集してください。"
    End If
    s = s & "</p>" & vbLf
    s = s & "</div></noscript></div>" & vbLf

    s = s & "<main id=""doc"" class=""wrap"">" & vbLf
    s = s & "<nav id=""tocprint"" class=""tocprint print-only""></nav>" & vbLf
    s = s & "<section id=""sec-cover"" class=""sec sec-cover""></section>" & vbLf
    s = s & "</main>" & vbLf
    BodyShellHtml = s
End Function

' SectionsJs - セクション登録表(18章§4.2)。**編集が最も多い1関数**。(a)描画関数
'   の連結行 と (b)登録配列 の2行1組で1セクションを表す。登録行のキーは
'   id / slug / title / need / empty / note / render の7つに固定。並び順は18章§3
'   の表の上から下(v1.2で見本の10節の流れへ並べ替えた。§3.0)。`empty` は§3の
'   「空のときの挙動」列と、`title` は§3の「見出し(既定)」列と1対1で対応させる
'   (§3が正。括弧つきのフル表記まで逐語で写す=W3.1裁定。上部ナビと目次も同じ
'   title を並べるため、ここがずれると本文と目次が同時に漂流する)。
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
    s = s & modHtmlTemplate7.SecGrowthJs() ' SAFE:html
    s = s & modHtmlTemplate8.SecTalkJs() ' SAFE:html

    ' (b) 登録配列
    s = s & "var SECTIONS=[" & vbLf
    s = s & "{id:'SEC-01',slug:'cover',title:'',need:['meta'],empty:'always',render:renderCover}," & vbLf
    s = s & "{id:'SEC-02',slug:'exec',title:'要点(1分で読む)',need:['s1'],empty:'always',render:renderExec}," & vbLf
    s = s & "{id:'SEC-03',slug:'profile',title:'企業理解',need:['s1'],empty:'hide',render:renderProfile}," & vbLf
    s = s & "{id:'SEC-04',slug:'sufficiency',title:'入力の充足度と要確認事項',need:['s1'],empty:'hide',render:renderSufficiency}," & vbLf
    s = s & "{id:'SEC-05',slug:'riskuniv',title:'リスクの全体像(10分類)',need:['s2'],empty:'hide',render:renderRiskUniv}," & vbLf
    s = s & "{id:'SEC-06',slug:'riskmap',title:'2軸リスクマップ（影響×頻度 5×5）'," & vbLf
    s = s & "need:['s2'],empty:'hide',render:renderRiskMap}," & vbLf
    s = s & "{id:'SEC-16',slug:'round-update',title:'訪問で分かったこと（ラウンド更新）'," & vbLf
    s = s & "need:['s2'],empty:'hide',render:renderRoundUpdate}," & vbLf
    s = s & "{id:'SEC-07',slug:'risks',title:'リスク一覧',need:['s2'],empty:'hide',render:renderRisks}," & vbLf
    s = s & "{id:'SEC-08',slug:'coverage',title:'いまの保険と足りないところ',need:['s2'],empty:'hide',render:renderCoverage}," & vbLf
    s = s & "{id:'SEC-11',slug:'prevent',title:'未然防止メニュー',need:['s2'],empty:'hide',render:renderPrevent}," & vbLf
    s = s & "{id:'SEC-12',slug:'limit',title:'当社にできないこと・提案を控えること',need:['s2'],empty:'note'," & vbLf
    s = s & "note:'該当なし',render:renderLimit}," & vbLf
    s = s & "{id:'SEC-09',slug:'newrisk',title:'新しく出てきたリスク（新種・新興リスク）'," & vbLf
    s = s & "need:['s2'],empty:'note'," & vbLf
    s = s & "note:'現時点で特筆すべき新しく出てきたリスクは検出されていません',render:renderNewRisk}," & vbLf
    s = s & "{id:'SEC-17',slug:'growth',title:'攻めの保険活用',need:['s3'],empty:'hide',render:renderGrowth}," & vbLf
    s = s & "{id:'SEC-10',slug:'story',title:'提案の筋書き（当社にできること）',need:['s3'],empty:'hide',render:renderStory}," & vbLf
    s = s & "{id:'SEC-18',slug:'talk',title:'経営層への話し方',need:['s3'],empty:'hide',render:renderTalk}," & vbLf
    s = s & "{id:'SEC-13',slug:'hearing',title:'ヒアリング事項',need:['s1'],empty:'hide',render:renderHearing}," & vbLf
    s = s & "{id:'SEC-14',slug:'source',title:'出典と根拠',need:['s2'],empty:'hide',render:renderSource}," & vbLf
    s = s & "{id:'SEC-15',slug:'disclaimer',title:'免責とご確認事項',need:['meta'],empty:'always',render:renderDisclaimer}" & vbLf
    s = s & "];" & vbLf
    SectionsJs = s
End Function

' RuntimeJs - 共通の描画ヘルパ・enum変換表とキッカー表(modHtmlTemplate6)・
'   登録配列の走査・上部ナビと目次の生成をこの順に連結する(§4.4)。描画は
'   createElement / textContent / setAttribute だけで行う(18章§4.1)。末尾で
'   run() を呼ぶ。関数宣言は巻き上げられるため SectionsJs が先でも問題ない。
Public Function RuntimeJs() As String
    Dim s As String
    s = s & modHtmlTemplate6.LabelJs() ' SAFE:html
    s = s & HelperJs() ' SAFE:html
    s = s & DriverJs() ' SAFE:html
    RuntimeJs = s
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
    ' 18章§3.8: 空文字の列は `-` を出す(空欄と「確認点なし」を見分けさせない)。
    s = s & "function DASH(v){return NB(v)?S(v):'-';}" & vbLf
    s = s & "function CLIP(t,n){var x=S(t);" & vbLf
    s = s & "return (x.length>n)?(x.slice(0,n)+'…'):x;}" & vbLf
    s = s & "function PARA(p,text,cls){var a=S(text).split(/\r\n|\r|\n/);" & vbLf
    s = s & "for(var i=0;i<a.length;i++){if(NB(a[i])){T(p,'p',cls,a[i]);}}}" & vbLf
    s = s & "function BDG(p,cls,text){return T(p,'span','bdg '+cls,text);}" & vbLf
    s = s & "function CHIP(p,text,cls){return T(p,'span','chip'+(cls?' '+cls:''),text);}" & vbLf
    s = s & "function CLR(n){while(n.firstChild){n.removeChild(n.firstChild);}return n;}" & vbLf
    s = s & "function TBL(host,heads,rows){var w=T(host,'div','tblwrap');" & vbLf
    s = s & "var tb=T(w,'table');var hr=T(T(tb,'thead'),'tr');" & vbLf
    s = s & "for(var i=0;i<heads.length;i++){T(hr,'th',null,heads[i]);}" & vbLf
    s = s & "var bd=T(tb,'tbody');FILL(bd,rows);return bd;}" & vbLf
    s = s & "function FILL(bd,rows){CLR(bd);" & vbLf
    s = s & "for(var r=0;r<rows.length;r++){var tr=T(bd,'tr');" & vbLf
    s = s & "for(var c=0;c<rows[r].length;c++){var v=rows[r][c];" & vbLf
    s = s & "if(v&&v.nodeType===1){T(tr,'td').appendChild(v);}" & vbLf
    s = s & "else{T(tr,'td',null,S(v));}}}return bd;}" & vbLf
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

' 登録配列の走査・need/empty 判定・上部ナビと目次・警告バナー・印刷ボタン。
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
    ' 18章§3.6: 同じ1本の一覧から、上部ナビ(toc)と本文先頭の目次(tocprint)を作る。
    s = s & "function navs(list){var bar=document.getElementById('toc');" & vbLf
    s = s & "var pr=document.getElementById('tocprint');var any=false;var ol=null;" & vbLf
    s = s & "if(pr){T(pr,'div','toc-h','目次');ol=T(pr,'ol');}" & vbLf
    s = s & "for(var i=0;i<list.length;i++){var d=list[i];if(!d.title){continue;}" & vbLf
    s = s & "any=true;" & vbLf
    s = s & "if(bar){AT(T(bar,'a',null,d.title),'href','#sec-'+d.slug);}" & vbLf
    s = s & "if(ol){AT(T(T(ol,'li'),'a',null,d.title),'href','#sec-'+d.slug);}}" & vbLf
    s = s & "if(!any&&pr&&pr.parentNode){pr.parentNode.removeChild(pr);}}" & vbLf
    ' 節の見出し(キッカー+h2)。キッカーは18章§3.0の KICK 表から引く。
    s = s & "function head(host,d){if(!d.title){return;}" & vbLf
    s = s & "var h=T(host,'div','section-head');" & vbLf
    s = s & "if(KICK[d.id]){T(h,'div','section-kicker',KICK[d.id]);}" & vbLf
    s = s & "T(h,'h2',null,d.title);}" & vbLf
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
    s = s & "if(!fixed){head(host,d);}" & vbLf
    s = s & "if(isEmpty){T(host,'p',(em==='note')?'note':'muted'," & vbLf
    s = s & "(em==='note'&&d.note)?d.note:'表示できるデータがありません。');}" & vbLf
    s = s & "else{while(body.firstChild){host.appendChild(body.firstChild);}}" & vbLf
    s = s & "if(!fixed){doc.appendChild(host);}shown.push(d);}" & vbLf
    s = s & "navs(shown);" & vbLf
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
