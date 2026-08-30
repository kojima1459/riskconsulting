Attribute VB_Name = "modHtmlTemplate2"
Option Explicit

' ==========================================================
' modHtmlTemplate2 - SEC-01 cover / SEC-02 exec / SEC-03 profile /
'                    SEC-04 sufficiency の描画スクリプト
' ------------------------------------------------
' 正は18章§3(セクションIDと読むJSONパス)・§3.1(エグゼクティブサマリの構成)・
' §4.4(分割規約の既定割り当て)。各関数は `function renderXxx(DATA,el){...}` を
' 1本返すだけで、CSSも登録行も持たない(CSSは modHtmlTemplate1.HeadHtml へ一元化・
' 登録は modHtmlTemplate1.SectionsJs へ一元化。18章§4.3の「変更は2箇所で完結」)。
'
' 描画は createElement / textContent / setAttribute だけで行う(18章§4.1)。
' 共通ヘルパ(E/T/AT/S/NB/AR/LB/CLIP/PARA/BDG/CHIP/TBL/DLIST/JOIN/RANKED/RNAME
' /GNAME)と enum変換表(LCAT/LTR/LIQ 等)は modHtmlTemplate1.RuntimeJs にある。
'
' R4準拠: Excelトークン・Application.Run・案件データ参照を持たない純文字列。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ==========================================================

' SEC-01 cover。表紙の会社名・案件ID・生成日時は静的HTML側(BodyShellHtml)に
'   あるので、ここは meta のチップ列だけを足す(18章§3の「見出し+チップ列」)。
Public Function SecCoverJs() As String
    Dim s As String
    s = s & "function renderCover(D,el){var m=D.meta||{};" & vbLf
    s = s & "var c=T(el,'div','chips');" & vbLf
    s = s & "if(NB(m.industry_name)){CHIP(c,'業種 '+S(m.industry_name));}" & vbLf
    s = s & "if(NB(m.case_type)){CHIP(c,LB(LCT,m.case_type),'chip-ai');}" & vbLf
    s = s & "if(NB(m.dossier_tier)){CHIP(c,'ドシエ '+LB(LTIER,m.dossier_tier));}" & vbLf
    s = s & "if(NB(m.quality_mode)){CHIP(c,'モード '+LB(LQM,m.quality_mode));}" & vbLf
    s = s & "if(m.round_no){CHIP(c,'第'+S(m.round_no)+'ラウンド');}" & vbLf
    s = s & "var s1=D.s1||{};" & vbLf
    s = s & "if(NB(s1.company_name)&&S(s1.company_name)!==S(m.company)){" & vbLf
    s = s & "T(el,'p','muted','Step1が読み取った社名: '+S(s1.company_name));}}" & vbLf
    SecCoverJs = s
End Function

' SEC-02 exec(18章§3.1)。リード -> 最重要リスク3件 -> 3テーマ の順に、
'   A4 1枚相当の文字中心で構成する(印刷時の break-after は共通CSSの .sec-exec)。
Public Function SecExecJs() As String
    Dim s As String
    s = s & "function renderExec(D,el){var s1=D.s1||{};var s3=D.s3;" & vbLf
    s = s & "PARA(el,s1.business_summary,'lead');" & vbLf
    s = s & "var so=s1.strategy_outlook||{};" & vbLf
    s = s & "if(NB(so.market_context)){T(el,'p','muted',S(so.market_context));}" & vbLf
    s = s & "var r=RANKED(D);" & vbLf
    s = s & "if(r.length){T(el,'h3',null,'いま最も重いリスク3件');" & vbLf
    s = s & "for(var i=0;i<r.length&&i<3;i++){var x=r[i];" & vbLf
    s = s & "var row=T(el,'div','rowline');" & vbLf
    s = s & "T(row,'span','rank-no','No.'+S(x.risk_no));" & vbLf
    s = s & "T(row,'span',null,S(x.risk_name));" & vbLf
    s = s & "T(row,'span','muted','影響'+S(x.impact_score)+' x 頻度'+S(x.frequency_score));}}" & vbLf
    s = s & "T(el,'h3',null,'ご提案の3テーマ');" & vbLf
    s = s & "if(!s3||!AR(s3.stories).length){" & vbLf
    s = s & "T(el,'p','muted','提案ストーリーは未生成です。');return;}" & vbLf
    s = s & "var st=AR(s3.stories).slice(0);" & vbLf
    s = s & "st.sort(function(a,b){return (a.story_no||0)-(b.story_no||0);});" & vbLf
    s = s & "for(var j=0;j<st.length&&j<3;j++){var y=st[j];" & vbLf
    s = s & "var card=T(el,'div','card');" & vbLf
    s = s & "T(card,'h3',null,S(y.headline));" & vbLf
    s = s & "var names=[];var tn=AR(y.target_risk_nos);" & vbLf
    s = s & "for(var k=0;k<tn.length;k++){names.push(RNAME(D,tn[k]));}" & vbLf
    s = s & "if(names.length){T(card,'p','muted','対象リスク: '+names.join('・'));}" & vbLf
    s = s & "T(card,'p',null,CLIP(y.pitch,200));}}" & vbLf
    SecExecJs = s
End Function

' SEC-03 profile。定義リスト + 拠点表(18章§3の図表種別)。
Public Function SecProfileJs() As String
    Dim s As String
    s = s & "function renderProfile(D,el){var s1=D.s1||{};" & vbLf
    s = s & "var sc=s1.supply_chain||{};var cu=s1.customers||{};" & vbLf
    s = s & "var so=s1.strategy_outlook||{};" & vbLf
    s = s & "DLIST(el,[['事業概要',S(s1.business_summary)]," & vbLf
    s = s & "['主要製品',JOIN(s1.main_products,'／')]," & vbLf
    s = s & "['主な工程',JOIN(s1.processes,'／')]," & vbLf
    s = s & "['主要な調達品',JOIN(sc.key_materials,'／')]," & vbLf
    s = s & "['供給網のメモ',S(sc.notes)]," & vbLf
    s = s & "['顧客セグメント',JOIN(cu.segments,'／')]," & vbLf
    s = s & "['販売チャネル',JOIN(cu.channels,'／')]," & vbLf
    s = s & "['人員・労務',S(s1.workforce_notes)]," & vbLf
    s = s & "['経営の動き',S(s1.management_notes)]," & vbLf
    s = s & "['理念(MVV)',S(so.mvv)]," & vbLf
    s = s & "['目指す姿',JOIN(so.aspirations,'／')]," & vbLf
    s = s & "['市況・外部環境',S(so.market_context)]]);" & vbLf
    s = s & "var lo=AR(s1.locations);if(!lo.length){return;}" & vbLf
    s = s & "T(el,'h3',null,'拠点と所在地');var rows=[];" & vbLf
    s = s & "for(var i=0;i<lo.length;i++){var x=lo[i];" & vbLf
    s = s & "rows.push([S(x.name),S(x.type),S(x.address),S(x.hazard_note),S(x.notes)]);}" & vbLf
    s = s & "TBL(el,['拠点','区分','所在地','ハザード','備考'],rows);}" & vbLf
    SecProfileJs = s
End Function

' SEC-04 sufficiency。14観点バッジ + 不足情報の表(18章§3)。該当0件の観点も
'   missing として必ず14個並べる(「見ていない」と「見たが無い」を区別する)。
Public Function SecSufficiencyJs() As String
    Dim s As String
    s = s & "function renderSufficiency(D,el){var s1=D.s1||{};" & vbLf
    s = s & "var iq=s1.input_quality||{};var cov=AR(iq.coverage);var map={};" & vbLf
    s = s & "for(var i=0;i<cov.length;i++){map[S(cov[i].aspect)]=S(cov[i].status);}" & vbLf
    s = s & "var box=T(el,'div','chips');" & vbLf
    s = s & "for(var j=0;j<ASPORDER.length;j++){var a=ASPORDER[j];" & vbLf
    s = s & "var st=map[a]?map[a]:'missing';var sp=E('span','chip');" & vbLf
    s = s & "T(sp,'span',null,LB(LASP,a)+' ');" & vbLf
    s = s & "BDG(sp,IQCLS[st]?IQCLS[st]:'bdg-missing',LB(LIQ,st));" & vbLf
    s = s & "box.appendChild(sp);}" & vbLf
    ' 裁定書10補遺P5: LIQO は19章§3どおり「高/中/低」の1字ラベルなので、
    '   単独段落では意味が立たない。表示側で「充足度: 」を前置する(18章§3)。
    s = s & "if(NB(iq.overall)){T(el,'p','muted','充足度: '+LB(LIQO,iq.overall));}" & vbLf
    s = s & "if(NB(iq.advice)){T(el,'p',null,S(iq.advice));}" & vbLf
    s = s & "var mi=AR(s1.missing_info);if(!mi.length){return;}" & vbLf
    s = s & "T(el,'h3',null,'要確認事項(いま足りていない情報)');var rows=[];" & vbLf
    s = s & "for(var k=0;k<mi.length;k++){" & vbLf
    s = s & "rows.push([S(mi[k].item),S(mi[k].why_needed)]);}" & vbLf
    s = s & "TBL(el,['不足している情報','なぜ必要か'],rows);}" & vbLf
    SecSufficiencyJs = s
End Function
