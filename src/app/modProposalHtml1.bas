Attribute VB_Name = "modProposalHtml1"
Option Explicit

' ==========================================================
' modProposalHtml1 - 顧客向け提案書(Wide 22枚)の骨格・スライド登録表・ランタイム
' ------------------------------------------------
' 正は20章(§1 責務分担 / §3 DATA / §4 スライド登録表 / §5 CSS / §6 印刷 /
'   §7 発表者ノート)。裁定は docs/29 §5(「AIは提案書JSONだけ・HTMLはVBAが
'   Wide 22枚テンプレで組む」)。
'
' 設計の要(18章§4.1 と同じ規律をそのまま適用する):
'   ・DATAはページ内のJSが JSON.parse で受け取り、描画は createElement と
'     textContent と setAttribute **だけ**で行う。innerHTML / insertAdjacentHTML /
'     document.write / outerHTML を1つも使わない(16章 E-47)。
'   ・自己完結HTML 1ファイル。外部CSS・JS・フォント・画像・CDNを参照しない。
'   ・**LLMにHTMLを書かせない**(docs/29 §5.1)。LLM出力はデータとしてのみ入る。
'   ・枚数は22枚で固定。データが無い枚も「次回のお打ち合わせで更新します」の
'     1行を置いて**必ず描く**(顧客へ渡す資料で頁が抜けると構成が崩れるため。
'     18章のレポートが `hide` を持つのと扱いが違う。20章§4)。
'
' 分割(20章§4.4。18章§4.4 の25,000字規約に倣う):
'   modProposalHtml1 = 全体組立 / HeadHtml / BodyShellHtml / SlidesJs / RuntimeJs
'   modProposalHtml2 = RootCss / FormatCss / ContentCss
'   modProposalHtml3 = 1..11枚目の描画   modProposalHtml4 = 12..22枚目の描画
'
' R4準拠(12章§2): Excelトークン・Application.Run・案件データ参照を持たない
'   純文字列。CP932準拠(15章§0 原則7)。
' 組立方式は15章・18章と同じ `s = s & "..." & vbLf`。
' ==========================================================

' coverFields の書式(vbTab区切り)。[0]=会社名 [1]=題 [2]=副題 [3]=日付
'   [4]=確認者。18章§5.2 の coverFields と同じ考え方(テンプレ側にJSONパーサを
'   持たせないための1本の文字列)。
Private Const PH1_SEP As String = vbTab

' BuildProposalDocument - HTML全文(20章§1.1の手順(5))。
'   dataJson=20章§3のDATA / coverFields=上記5値。
Public Function BuildProposalDocument(ByVal dataJson As String, _
                                      ByVal coverFields As String) As String
    Dim s As String
    s = s & "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja"">" & vbLf
    s = s & HeadHtml(coverFields)
    s = s & BodyShellHtml(coverFields)

    ' 18章§5.3(1) と同じ規約: DATAは「1本のJS文字列リテラル + JSON.parse」の
    ' 形でのみ埋め、modUtilText.JsStringSafe を必ず通す(生の < を1文字も残さない)。
    s = s & "<script>var DATA=JSON.parse("""
    s = s & modUtilText.JsStringSafe(dataJson)
    s = s & """);</script>" & vbLf

    s = s & "<script>" & vbLf
    s = s & "(function(){" & vbLf
    s = s & "'use strict';" & vbLf
    s = s & SlidesJs()
    s = s & RuntimeJs()
    s = s & "})();" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body>" & vbLf
    s = s & "</html>" & vbLf
    BuildProposalDocument = s
End Function

' HeadHtml - <head>。**最初の要素**として <meta charset="utf-8"> を置く。
'   CSSの呼び口は本関数1本(実体は modProposalHtml2 の3本)。
Public Function HeadHtml(ByVal coverFields As String) As String
    Dim s As String
    s = s & "<head>" & vbLf
    s = s & "<meta charset=""utf-8"">" & vbLf
    s = s & "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">" & vbLf
    s = s & "<title>"
    s = s & modUtilText.HtmlSafe("ご提案 " & FieldAt(coverFields, 0))
    s = s & "</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & modProposalHtml2.RootCss()
    s = s & modProposalHtml2.FormatCss()
    s = s & modProposalHtml2.ContentCss()
    s = s & "</style>" & vbLf
    s = s & "</head>" & vbLf
    HeadHtml = s
End Function

' BodyShellHtml - 骨格と <noscript>。VBAが静的HTMLとして書き出すのは
'   (a)<title> (b)<noscript> の宛名と題 の2つだけで、どちらも HtmlSafe を通す。
'   顧客向けなので AI 表示は出さない(社内IT・AI環境 v1.1 §7.3。担当者の確認を
'   経た顧客提示物であり、確認していない資料はそもそも出力できない
'   =modExportProposal が reviewedBy 空で生成を断る)。
Public Function BodyShellHtml(ByVal coverFields As String) As String
    Dim s As String
    s = s & "<body>" & vbLf
    s = s & "<main id=""deck""></main>" & vbLf
    s = s & "<noscript><div style=""padding:24px"">" & vbLf
    s = s & "<p>"
    s = s & modUtilText.HtmlSafe(FieldAt(coverFields, 0) & " 御中")
    s = s & "</p>" & vbLf
    s = s & "<p>"
    s = s & modUtilText.HtmlSafe(FieldAt(coverFields, 1))
    s = s & "</p>" & vbLf
    s = s & "<p>この資料の表示にはJavaScriptが必要です。"
    s = s & "ファイルを保存してブラウザで開いてください。</p>" & vbLf
    s = s & "</div></noscript>" & vbLf
    BodyShellHtml = s
End Function

' ==========================================================
' SlidesJs - スライド登録表(20章§4.2)。**編集が最も多い1関数**。
'   (a)描画関数の連結行 と (b)登録配列 の2行1組で1枚を表す。
'   登録行のキーは no / slug / master / title / label / render の**6つに固定**。
'   増やさない(18章§4.2 が登録行を7キーに固定しているのと同じ理由)。
'     no     : 1..22。ページ番号であり、発表者ノート notes[].slide_no と対応する
'     slug   : <section id="sl-<slug>"> になる。英小文字とハイフンのみ
'     master : cover / section / content / back の4値(20章§4.1)
'     title  : 本文マスターの帯に出す見出し。20章§4の表を逐語で写す
'     label  : セクション扉の「Section N」。master='section' のときだけ意味を持つ
'     render : 描画関数。引数は (DATA, bodyEl) の2つに固定し戻り値を持たない
' ==========================================================
Public Function SlidesJs() As String
    Dim s As String
    ' (a) 描画関数の連結行
    s = s & modProposalHtml3.SlCoverJs()
    s = s & modProposalHtml3.SlSummaryJs()
    s = s & modProposalHtml3.SlSectionJs()
    s = s & modProposalHtml3.SlBusinessJs()
    s = s & modProposalHtml3.SlUniverseJs()
    s = s & modProposalHtml3.SlRiskMapJs()
    s = s & modProposalHtml3.SlClassesJs()
    s = s & modProposalHtml3.SlPriorityJs()
    s = s & modProposalHtml3.SlHardJs()
    s = s & modProposalHtml4.SlIdeasJs()
    s = s & modProposalHtml4.SlFourJs()
    s = s & modProposalHtml4.SlThemesJs()
    s = s & modProposalHtml4.SlStepsJs()
    s = s & modProposalHtml4.SlDecideJs()
    s = s & modProposalHtml4.SlShareJs()
    s = s & modProposalHtml4.SlAppendixJs()
    s = s & modProposalHtml4.SlPremiseJs()
    s = s & modProposalHtml4.SlBackJs()

    ' (b) 登録配列(22行。20章§4の表と1対1)
    s = s & "var SLIDES=[" & vbLf
    s = s & "{no:1,slug:'cover',master:'cover',title:'',label:'',render:renderCover}," & vbLf
    s = s & "{no:2,slug:'summary',master:'content',title:'ご提案の要旨',label:''," & vbLf
    s = s & "render:renderSummary}," & vbLf
    s = s & "{no:3,slug:'sec1',master:'section',title:'貴社の事業構造とリスクの全体像'," & vbLf
    s = s & "label:'Section 1',render:renderSection}," & vbLf
    s = s & "{no:4,slug:'business',master:'content',title:'事業構造：リスクを生む5つの事業領域'," & vbLf
    s = s & "label:'',render:renderBusiness}," & vbLf
    s = s & "{no:5,slug:'universe',master:'content',title:'リスクの全体像：10のカテゴリーで捉える'," & vbLf
    s = s & "label:'',render:renderUniverse}," & vbLf
    s = s & "{no:6,slug:'riskmap',master:'content',title:'リスクマップ：リスクの優先順位'," & vbLf
    s = s & "label:'',render:renderRiskMap}," & vbLf
    s = s & "{no:7,slug:'sec2',master:'section',title:'既存の保険で備えられるリスクと、その条件'," & vbLf
    s = s & "label:'Section 2',render:renderSection}," & vbLf
    s = s & "{no:8,slug:'classes',master:'content',title:'保険での備えやすさ：3つの分類'," & vbLf
    s = s & "label:'',render:renderClasses}," & vbLf
    s = s & "{no:9,slug:'priority',master:'content',title:'優先リスク：対応する保険と確認事項'," & vbLf
    s = s & "label:'',render:renderPriority}," & vbLf
    s = s & "{no:10,slug:'sec3',master:'section',title:'保険だけでは備えにくいリスク'," & vbLf
    s = s & "label:'Section 3',render:renderSection}," & vbLf
    s = s & "{no:11,slug:'hard',master:'content',title:'保険だけでは備えにくいリスク'," & vbLf
    s = s & "label:'',render:renderHard}," & vbLf
    s = s & "{no:12,slug:'sec4',master:'section',title:'成長を後押しする保険の活用'," & vbLf
    s = s & "label:'Section 4',render:renderSection}," & vbLf
    s = s & "{no:13,slug:'ideas',master:'content',title:'成長支援：保険活用アイデアと優先候補'," & vbLf
    s = s & "label:'',render:renderIdeas}," & vbLf
    s = s & "{no:14,slug:'four',master:'content',title:'優先4案：狙い・仕組み・想定する保険'," & vbLf
    s = s & "label:'',render:renderFour}," & vbLf
    s = s & "{no:15,slug:'sec5',master:'section',title:'ご提案の全体像と進め方'," & vbLf
    s = s & "label:'Section 5',render:renderSection}," & vbLf
    s = s & "{no:16,slug:'themes',master:'content',title:'提案テーマ：経営課題ごとに補償と指標を束ねる'," & vbLf
    s = s & "label:'',render:renderThemes}," & vbLf
    s = s & "{no:17,slug:'steps',master:'content',title:'進め方：4つのステップ'," & vbLf
    s = s & "label:'',render:renderSteps}," & vbLf
    s = s & "{no:18,slug:'decide',master:'content',title:'本日ご判断いただきたいこと'," & vbLf
    s = s & "label:'',render:renderDecide}," & vbLf
    s = s & "{no:19,slug:'share',master:'content',title:'次のステップ：ご共有いただきたい事項'," & vbLf
    s = s & "label:'',render:renderShare}," & vbLf
    s = s & "{no:20,slug:'appendix',master:'content',title:'巻末資料：リスク一覧'," & vbLf
    s = s & "label:'',render:renderAppendix}," & vbLf
    s = s & "{no:21,slug:'premise',master:'content',title:'本資料の前提と参照した公開情報'," & vbLf
    s = s & "label:'',render:renderPremise}," & vbLf
    s = s & "{no:22,slug:'back',master:'back',title:'',label:'',render:renderBack}" & vbLf
    s = s & "];" & vbLf
    SlidesJs = s
End Function

' RuntimeJs - 共通の描画ヘルパ + 走査・マスターの器・発表者ノート。
Public Function RuntimeJs() As String
    Dim s As String
    s = s & HelperJs()
    s = s & DriverJs()
    RuntimeJs = s
End Function

' 共通の描画ヘルパ。innerHTML系は使わない(20章§1・18章§4.1)。
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
    s = s & "function P(D){return D.p||{};}" & vbLf
    s = s & "function HL(D,k){return S((P(D).headline||{})[k]);}" & vbLf
    s = s & "function ST(D){return D.stats||{};}" & vbLf
    s = s & "function RK(D){return AR(D.risks);}" & vbLf
    ' 20章§4: データが無い枚も必ず描く。共通の1行はここ1箇所が値源。
    s = s & "function TODO(el){T(el,'p','muted'," & vbLf
    s = s & "'この項目は次回のお打ち合わせで更新します。');}" & vbLf
    s = s & "function HEAD(el,text){if(NB(text)){T(el,'p','hm',text);}}" & vbLf
    ' 行が0件なら**表そのものを作らない**(見出しだけの空表を顧客へ見せない。
    ' 本文が1つも無い枚は 20章§4.1 の1行に置き換わる)。
    s = s & "function TBL(host,heads,rows,priIdx){if(!rows||!rows.length){return null;}" & vbLf
    s = s & "var t=T(host,'table','t');" & vbLf
    s = s & "var hr=T(T(t,'thead'),'tr');var i,c;" & vbLf
    s = s & "for(i=0;i<heads.length;i++){T(hr,'th',null,heads[i]);}" & vbLf
    s = s & "var bd=T(t,'tbody');" & vbLf
    s = s & "for(i=0;i<rows.length;i++){" & vbLf
    s = s & "var tr=T(bd,'tr',(priIdx&&priIdx[i])?'pri':null);" & vbLf
    s = s & "for(c=0;c<rows[i].length;c++){T(tr,'td',null,S(rows[i][c]));}}" & vbLf
    s = s & "return bd;}" & vbLf
    s = s & "function STARS(n){var k=Math.max(0,Math.min(5,Number(n)||0));" & vbLf
    s = s & "var o='';for(var i=0;i<5;i++){o+=(i<k)?'*':'-';}return o;}" & vbLf
    HelperJs = s
End Function

' 走査・マスターの器・ページ番号・発表者ノート。
Private Function DriverJs() As String
    Dim s As String
    ' マスター4種の器を作る。中身(bodyEl)だけを描画関数へ渡す。
    s = s & "function shell(d,sec){" & vbLf
    s = s & "if(d.master==='content'){" & vbLf
    s = s & "T(sec,'div','header-bar');T(sec,'div','content-title',d.title);" & vbLf
    s = s & "return T(sec,'div','content-body');}" & vbLf
    s = s & "if(d.master==='section'){" & vbLf
    s = s & "T(sec,'div','section-label',d.label);" & vbLf
    s = s & "T(sec,'div','section-title',d.title);" & vbLf
    s = s & "T(sec,'div','section-rule');" & vbLf
    s = s & "return T(sec,'div','content-body');}" & vbLf
    s = s & "return T(sec,'div','cover-body');}" & vbLf
    ' 発表者ノート(20章§7)。?notes=1 のときだけ画面に出す。印刷では常に消える。
    s = s & "function noteOf(D,no){var a=AR(P(D).notes);" & vbLf
    s = s & "for(var i=0;i<a.length;i++){if(Number(a[i].slide_no)===no){return a[i];}}" & vbLf
    s = s & "return null;}" & vbLf
    s = s & "function notes(host,D,no){var n=noteOf(D,no);if(!n){return;}" & vbLf
    s = s & "var box=T(host,'aside','note');" & vbLf
    s = s & "T(box,'b',null,'読み上げ: ');T(box,'span',null,S(n.read));" & vbLf
    s = s & "if(NB(n.ask)){T(box,'br');T(box,'b',null,'問いかけ: ');" & vbLf
    s = s & "T(box,'span',null,S(n.ask));}" & vbLf
    s = s & "var pb=AR(n.probe);if(pb.length){T(box,'br');" & vbLf
    s = s & "T(box,'b',null,'更問: ');T(box,'span',null,pb.join(' / '));}" & vbLf
    s = s & "var fl=AR(n.follow);if(fl.length){T(box,'br');" & vbLf
    s = s & "T(box,'b',null,'フォロー: ');T(box,'span',null,fl.join(' / '));}}" & vbLf
    ' 画面の縮小率(見本の viewer-css と同じ手。style.setProperty だけを使う)。
    s = s & "function fit(){var fs=document.getElementsByClassName('frame');" & vbLf
    s = s & "for(var i=0;i<fs.length;i++){var w=fs[i].clientWidth||0;" & vbLf
    s = s & "if(w&&fs[i].style&&fs[i].style.setProperty){" & vbLf
    s = s & "fs[i].style.setProperty('--s',String(w/1280));}}}" & vbLf
    s = s & "function run(){var D=DATA;var deck=document.getElementById('deck');" & vbLf
    s = s & "if(!deck){return;}" & vbLf
    s = s & "for(var i=0;i<SLIDES.length;i++){var d=SLIDES[i];" & vbLf
    s = s & "var fr=E('div','frame');" & vbLf
    s = s & "var sec=E('section','slide');AT(sec,'id','sl-'+d.slug);" & vbLf
    s = s & "AT(sec,'data-master',d.master);AT(sec,'data-no',String(d.no));" & vbLf
    s = s & "var body=shell(d,sec);" & vbLf
    ' 描画に失敗した枚は、20章§4.1 の1行に置き換えたうえで**印を残す**
    ' (裁定書39 R2-12。壊れたページが「わざと保留した項目」に見えないように
    '  section へ data-render-error を付け、?debug=1 のときだけ赤枠にする。
    '  お客さまが普通に開いた画面には何も足さない)。
    s = s & "try{d.render(D,body,d);}catch(err){AT(sec,'data-render-error','1');" & vbLf
    s = s & "TODO(body);}" & vbLf
    ' 本文マスターだけは空のとき1行を置く(表紙・扉・裏表紙は本文を持たない)。
    s = s & "if(body.childNodes.length===0&&d.master==='content'){TODO(body);}" & vbLf
    s = s & "if(d.master!=='cover'&&d.master!=='back'){" & vbLf
    s = s & "T(sec,'div','footer-rule');" & vbLf
    s = s & "T(sec,'div','page-number',String(d.no)+' / '+String(SLIDES.length));}" & vbLf
    s = s & "fr.appendChild(sec);deck.appendChild(fr);" & vbLf
    ' 発表者ノートも**同じ catch の内側**で描く(裁定書40 S-M3)。notes() の値源は
    ' S5 の notes[](LLM出力。s5_edited を人が直す経路もある)なので、ここで
    ' 例外が飛ぶと run() ごと落ちて提案書が1枚も描かれない=事実上の白紙になる。
    ' 上の d.render と同じく、印を残して次の枚へ進む。
    s = s & "try{notes(deck,D,d.no);}catch(err2){AT(sec,'data-render-error','1');}}" & vbLf
    ' 画面の切替は2つだけ。?notes=1=発表者ノート / ?debug=1=描画エラーの赤枠
    ' (社内での確認用。20章§11)。どちらも付けなければ何も起きない。
    s = s & "var q=String(location.search||'');var cls='';" & vbLf
    s = s & "if(q.indexOf('notes=1')>=0){cls='notes';}" & vbLf
    s = s & "if(q.indexOf('debug=1')>=0){cls=cls?(cls+' debug'):'debug';}" & vbLf
    s = s & "if(cls){document.body.className=cls;}" & vbLf
    s = s & "fit();" & vbLf
    s = s & "if(window.addEventListener){window.addEventListener('resize',fit,false);}}" & vbLf
    s = s & "run();" & vbLf
    DriverJs = s
End Function

' coverFields の idx 番目。範囲外は空文字列。
Private Function FieldAt(ByVal coverFields As String, ByVal idx As Long) As String
    Dim parts() As String
    parts = Split(coverFields, PH1_SEP)
    If idx < LBound(parts) Or idx > UBound(parts) Then Exit Function
    FieldAt = parts(idx)
End Function
