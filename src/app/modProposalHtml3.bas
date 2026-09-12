Attribute VB_Name = "modProposalHtml3"
Option Explicit

' ==========================================================
' modProposalHtml3 - 提案書 Wide の描画(1枚目から11枚目)。20章§4が正。
' ------------------------------------------------
'   1 cover / 2 summary / 3・7・10・12・15 section(共通) / 4 business /
'   5 universe / 6 riskmap / 8 classes / 9 priority / 11 hard
' 18章§4.1 と同じ規律: createElement と textContent と setAttribute だけ。
'   innerHTML / insertAdjacentHTML / document.write / outerHTML を書かない。
' 列見出し・固定文は**顧客語**で書く(docs/design/提案書_wide/対訳表_社内語から
'   顧客語.md)。「顕在化」「打ち手」「移転」などの社内語をここに書かない。
' R4準拠・CP932準拠。組立は `s = s & "..." & vbLf`。
' ==========================================================

' 1枚目 表紙。宛名・題・副題・日付・担当者。
'   担当者(meta.reviewed_by)は**必ず非空**である(modExportProposal が空なら
'   生成しない=顧客向けは確認必須。裁定書38 §1 班C)。
Public Function SlCoverJs() As String
    Dim s As String
    s = s & "function renderCover(D,el){var m=D.meta||{};" & vbLf
    s = s & "T(el,'div','cover-to',S(m.company)+' 御中');" & vbLf
    s = s & "var b=T(el,'div','cover-box');" & vbLf
    s = s & "T(b,'div','cover-title',S(P(D).title));" & vbLf
    s = s & "if(NB(P(D).subtitle)){T(b,'div','cover-sub',S(P(D).subtitle));}" & vbLf
    s = s & "T(el,'div','cover-date',S(m.date));" & vbLf
    s = s & "T(el,'div','cover-by','担当 '+S(m.reviewed_by));}" & vbLf
    SlCoverJs = s
End Function

' 2枚目 ご提案の要旨。実数の指標帯(VBAが数えた値)＋3テーマ＋本資料の構成。
Public Function SlSummaryJs() As String
    Dim s As String
    s = s & "function renderSummary(D,el){var st=ST(D);var th=AR(P(D).themes);" & vbLf
    s = s & "var band=T(el,'div','stats');" & vbLf
    s = s & "var pairs=[['リスク',st.risk_total],['保険で備えやすい',st.easy]," & vbLf
    s = s & "['補償条件の設計が必要',st.design],['保険以外の対策が中心',st.non_ins]," & vbLf
    s = s & "['成長支援の案',st.ideas],['優先候補',st.ideas_priority]];" & vbLf
    s = s & "for(var i=0;i<pairs.length;i++){var c=T(band,'div','stat');" & vbLf
    s = s & "T(c,'b',null,S(pairs[i][1]));T(c,'span',null,S(pairs[i][0]));}" & vbLf
    s = s & "var cols=T(el,'div','cols');" & vbLf
    s = s & "for(var k=0;k<th.length;k++){var col=T(cols,'div','col');" & vbLf
    s = s & "var card=T(col,'div','card');" & vbLf
    s = s & "T(card,'h3',null,'テーマ '+String(k+1)+'  '+S(th[k].name));" & vbLf
    s = s & "T(card,'p','lead',S(th[k].headline));" & vbLf
    s = s & "T(card,'p',null,S(th[k].body));}" & vbLf
    s = s & "var stc=AR(P(D).structure);" & vbLf
    s = s & "if(stc.length){T(el,'p','muted','本資料の構成: '+stc.join(' / '));}}" & vbLf
    SlSummaryJs = s
End Function

' 3・7・10・12・15枚目 セクション扉。器(ラベル・題・罫)は登録表の値から
'   modProposalHtml1.shell が描くため、本文は持たない(空のままでよい枚)。
Public Function SlSectionJs() As String
    Dim s As String
    s = s & "function renderSection(D,el,d){return;}" & vbLf
    SlSectionJs = s
End Function

' 4枚目 事業構造。左=貴社の事業領域(当社の理解)、右=事業構造から見えるリスク要因。
Public Function SlBusinessJs() As String
    Dim s As String
    s = s & "function renderBusiness(D,el){var bz=P(D).business||{};" & vbLf
    s = s & "var fa=AR(bz.facts);" & vbLf
    ' 3つとも空なら何も描かない(20章§4.1 の1行に置き換わる)。
    s = s & "if(!fa.length&&!AR(bz.areas).length&&!AR(bz.factors).length){return;}" & vbLf
    s = s & "if(fa.length){var band=T(el,'div','stats');" & vbLf
    s = s & "for(var f=0;f<fa.length;f++){var c=T(band,'div','stat');" & vbLf
    s = s & "T(c,'b',null,S(fa[f].value));" & vbLf
    s = s & "T(c,'span',null,S(fa[f].label)+(NB(fa[f].note)?('  '+S(fa[f].note)):''));}}" & vbLf
    s = s & "var cols=T(el,'div','cols');" & vbLf
    s = s & "var L=T(cols,'div','col');var R=T(cols,'div','col');" & vbLf
    s = s & "T(L,'p','sub','貴社の事業領域（公開情報に基づく当社の理解）');" & vbLf
    s = s & "var ar=AR(bz.areas);" & vbLf
    s = s & "for(var i=0;i<ar.length;i++){var kv=T(L,'div','kv');" & vbLf
    s = s & "T(kv,'b',null,S(ar[i].name));T(kv,'span',null,S(ar[i].desc));}" & vbLf
    s = s & "T(R,'p','sub','事業構造から見えるリスク要因');" & vbLf
    s = s & "var fc=AR(bz.factors);" & vbLf
    s = s & "for(var k=0;k<fc.length;k++){var kv2=T(R,'div','kv');" & vbLf
    s = s & "T(kv2,'b',null,S(fc[k].name));" & vbLf
    s = s & "var w=T(kv2,'span',null,S(fc[k].desc)+' ');" & vbLf
    s = s & "T(w,'span','pill',S(fc[k].tag));}}" & vbLf
    SlBusinessJs = s
End Function

' 5枚目 リスクの全体像。10のカテゴリーの件数バー(VBAが数えた値)。
Public Function SlUniverseJs() As String
    Dim s As String
    s = s & "function renderUniverse(D,el){var cn=P(D).categories_note||{};" & vbLf
    s = s & "if(!AR(D.categories).length&&!NB(cn.heavy)){return;}" & vbLf
    s = s & "HEAD(el,S(cn.heavy));" & vbLf
    s = s & "if(NB(cn.meaning)){T(el,'p','lead',S(cn.meaning));}" & vbLf
    s = s & "var cs=AR(D.categories);var mx=1;var i;" & vbLf
    s = s & "for(i=0;i<cs.length;i++){if((cs[i].count||0)>mx){mx=cs[i].count;}}" & vbLf
    s = s & "var box=T(el,'div','bars');" & vbLf
    s = s & "for(i=0;i<cs.length;i++){var row=T(box,'div','bar');" & vbLf
    s = s & "T(row,'span','n',String(i+1)+'. '+S(cs[i].label));" & vbLf
    s = s & "var bar=T(row,'i');" & vbLf
    s = s & "if(bar.style){bar.style.width=String(Math.round(320*(cs[i].count||0)/mx))+'px';}" & vbLf
    s = s & "T(row,'span',null,String(cs[i].count||0)+' 件');}}" & vbLf
    SlUniverseJs = s
End Function

' 6枚目 リスクマップ(影響度 x 起こりやすさ の 5x5)。色だけで意味を運ばないよう
'   セルにはリスク番号を文字で置く(18章§6(3)と同じ規律)。
Public Function SlRiskMapJs() As String
    Dim s As String
    s = s & "function bandOf(n){if(n>=16){return 'b4';}if(n>=12){return 'b3';}" & vbLf
    s = s & "if(n>=6){return 'b2';}return 'b1';}" & vbLf
    s = s & "function renderRiskMap(D,el){var rs=RK(D);if(!rs.length){return;}" & vbLf
    s = s & "HEAD(el,HL(D,'riskmap'));" & vbLf
    s = s & "var g=T(el,'div','map');var r,c,i;" & vbLf
    s = s & "for(r=5;r>=1;r--){" & vbLf
    s = s & "T(g,'div','ax','影響 '+String(r));" & vbLf
    s = s & "for(c=1;c<=5;c++){" & vbLf
    s = s & "var cell=T(g,'div','cell '+bandOf(r*c));" & vbLf
    s = s & "for(i=0;i<rs.length;i++){" & vbLf
    s = s & "if(Number(rs[i].impact)===r&&Number(rs[i].frequency)===c){" & vbLf
    s = s & "T(cell,'div',null,'No.'+S(rs[i].risk_no)+' '+S(rs[i].risk_name));}}}}" & vbLf
    s = s & "T(g,'div','ax','');" & vbLf
    s = s & "for(c=1;c<=5;c++){T(g,'div','ax','起こりやすさ '+String(c));}" & vbLf
    s = s & "T(el,'p','muted','番号は巻末のリスク一覧と対応しています。');}" & vbLf
    SlRiskMapJs = s
End Function

' 8枚目 保険での備えやすさ(3分類)。件数はVBAが数えた実数を使う。
Public Function SlClassesJs() As String
    Dim s As String
    s = s & "function renderClasses(D,el){var rs=RK(D);if(!rs.length){return;}" & vbLf
    s = s & "HEAD(el,HL(D,'classes'));" & vbLf
    s = s & "var st=ST(D);var band=T(el,'div','stats');" & vbLf
    s = s & "var pairs=[['保険で備えやすい',st.easy],['補償条件の設計が必要',st.design]," & vbLf
    s = s & "['保険以外の対策が中心',st.non_ins]];" & vbLf
    s = s & "for(var i=0;i<pairs.length;i++){var c=T(band,'div','stat');" & vbLf
    s = s & "T(c,'b',null,S(pairs[i][1])+' 件');T(c,'span',null,S(pairs[i][0]));}" & vbLf
    s = s & "var rows=[];" & vbLf
    s = s & "for(var k=0;k<rs.length;k++){" & vbLf
    s = s & "rows.push(['No.'+S(rs[k].risk_no),S(rs[k].risk_name),S(rs[k].class_label)]);}" & vbLf
    s = s & "TBL(el,['番号','リスク','備えやすさ'],rows,null);" & vbLf
    s = s & "T(el,'p','muted','補償の有無は商品名ではなく、事故の原因・損害の種類・"
    s = s & "契約の条件で決まるため、その観点で分けています。');}" & vbLf
    SlClassesJs = s
End Function

' 9枚目 優先リスク。上位8件の表(対応する保険は一般名称・貴社側の対策は例)。
Public Function SlPriorityJs() As String
    Dim s As String
    s = s & "function renderPriority(D,el){var rs=RK(D);if(!rs.length){return;}" & vbLf
    s = s & "HEAD(el,HL(D,'priority'));" & vbLf
    s = s & "var rows=[];var n=Math.min(8,rs.length);" & vbLf
    s = s & "for(var i=0;i<n;i++){var x=rs[i];" & vbLf
    s = s & "rows.push(['No.'+S(x.risk_no),S(x.risk_name),S(x.rank_label)," & vbLf
    s = s & "S(x.line_note),S(x.check_note),S(x.control_note)]);}" & vbLf
    s = s & "TBL(el,['番号','リスク','評価','対応する保険（一般名称）'," & vbLf
    s = s & "'補償の確認事項','貴社側の対策（例）'],rows,null);}" & vbLf
    SlPriorityJs = s
End Function

' 11枚目 保険だけでは備えにくいリスク。時期の見通しは顧客語で書く
'   (「顕在化」は対訳表の社内語なので列見出しに使わない)。
Public Function SlHardJs() As String
    Dim s As String
    s = s & "function renderHard(D,el){var hr=AR(P(D).hard_risks);" & vbLf
    s = s & "if(!hr.length){return;}HEAD(el,HL(D,'hard'));" & vbLf
    s = s & "var rs=RK(D);var rows=[];var i,k;" & vbLf
    s = s & "for(i=0;i<hr.length;i++){var nm='No.'+S(hr[i].risk_no);" & vbLf
    s = s & "for(k=0;k<rs.length;k++){" & vbLf
    s = s & "if(Number(rs[k].risk_no)===Number(hr[i].risk_no)){" & vbLf
    s = s & "nm=nm+' '+S(rs[k].risk_name);}}" & vbLf
    s = s & "rows.push([nm,S(hr[i].horizon),S(hr[i].background),S(hr[i].approach)]);}" & vbLf
    s = s & "TBL(el,['リスク','時期の見通し','背景','備え方（経営の施策と保険）']," & vbLf
    s = s & "rows,null);}" & vbLf
    SlHardJs = s
End Function
