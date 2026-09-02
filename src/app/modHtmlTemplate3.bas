Attribute VB_Name = "modHtmlTemplate3"
Option Explicit

' ==========================================================
' modHtmlTemplate3 - SEC-05 riskuniv / SEC-06 riskmap / SEC-07 risks /
'                    SEC-08 coverage の描画スクリプト
' ------------------------------------------------
' 正は18章§3(読むJSONパス)・§3.2(10分類の描き方)・§3.3(5x5マップの描き方)・
' §3.0(見本の10節との対応)・§4.4(分割規約の既定割り当て)。CSSも登録行も
' 持たない(18章§4.3)。描画は createElement / textContent / setAttribute だけ
' (18章§4.1)。
'
' R4準拠: Excelトークン・Application.Run・案件データ参照を持たない純文字列。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ==========================================================

' SEC-05 riskuniv(18章§3.2)。**常に10枚**のカードを並べ、該当0件の分類も0件と
'   表示する。「見ていない領域」と「見たが該当なし」を読み手が区別できるように
'   するため。並びは15章§0の変換表の記載順(CATORDER)に固定する。
'   見た目は見本の .universe(件数バー付きの10枚グリッド。11章§3.8.1)。
Public Function SecRiskUnivJs() As String
    Dim s As String
    s = s & "function renderRiskUniv(D,el){var rs=RISKS(D);" & vbLf
    s = s & "var cnt={};var tot=rs.length;" & vbLf
    s = s & "for(var i=0;i<rs.length;i++){var c=S(rs[i].category);" & vbLf
    s = s & "cnt[c]=(cnt[c]||0)+1;}" & vbLf
    s = s & "var g=T(el,'div','universe');" & vbLf
    s = s & "for(var j=0;j<CATORDER.length;j++){var k=CATORDER[j];" & vbLf
    s = s & "var n=cnt[k]||0;var u=T(g,'div','u');" & vbLf
    s = s & "T(u,'div','n','0'+String(j+1));" & vbLf
    s = s & "T(u,'h4',null,LB(LCAT,k));" & vbLf
    s = s & "T(u,'div','cnt',String(n));" & vbLf
    s = s & "T(u,'div','mini','件');" & vbLf
    s = s & "var bar=T(u,'div','bar');" & vbLf
    s = s & "var fill=T(bar,'span',null);" & vbLf
    s = s & "AT(fill,'style','width:'+((tot>0)?Math.round(n*100/tot):0)+'%');}" & vbLf
    s = s & "T(el,'p','muted','全'+tot+'件。0件の分類も「見たうえで該当なし」"
    s = s & "として並べています。');}" & vbLf
    SecRiskUnivJs = s
End Function

' SEC-06 riskmap(18章§3.3)。縦軸=impact_score(上が5)・横軸=frequency_score
'   (左が1)の5行5列。セル背景の帯は impact+frequency の和で決め(2-3=1 /
'   4-5=2 / 6-7=3 / 8-9=4 / 10=5 -> floor(sum/2))、**色に依存させない**ため
'   各セルの右下に帯番号を必ず添える。点の色は transferability。
'   1セルに6件以上入る場合は先頭5件+「+n」に切り替える(5x5の形を保つ)。
'   右側に見本の「経営優先度 Top 8」(impact x frequency の降順)を並べる
'   (11章§3.8.1。並びの規則は18章§3.1(2)の RANKED と同一)。
Public Function SecRiskMapJs() As String
    Dim s As String
    s = s & "function renderRiskMap(D,el){var rs=RISKS(D);" & vbLf
    s = s & "var w=T(el,'div','heat-wrap');" & vbLf
    s = s & "var left=T(w,'div',null);" & vbLf
    s = s & "var g=T(left,'div','heatmap');" & vbLf
    s = s & "for(var imp=5;imp>=1;imp--){" & vbLf
    s = s & "T(g,'div','y',String(imp));" & vbLf
    s = s & "for(var frq=1;frq<=5;frq++){var band=Math.floor((imp+frq)/2);" & vbLf
    s = s & "var cell=T(g,'div','cell heat'+band);var list=[];" & vbLf
    s = s & "for(var i=0;i<rs.length;i++){var x=rs[i];" & vbLf
    s = s & "if((x.impact_score||0)===imp&&(x.frequency_score||0)===frq){list.push(x);}}" & vbLf
    s = s & "list.sort(function(a,b){return (a.risk_no||0)-(b.risk_no||0);});" & vbLf
    s = s & "for(var k=0;k<list.length&&k<5;k++){var y=list[k];" & vbLf
    s = s & "var ins=y.insurability||{};var tr=S(ins.transferability);" & vbLf
    s = s & "var b=T(cell,'span','dot '+(DOTCLS[tr]?DOTCLS[tr]:''),String(y.risk_no));" & vbLf
    s = s & "AT(b,'title',S(y.risk_name)+' / '+LB(LTR,tr));}" & vbLf
    s = s & "if(list.length>5){T(cell,'span','mini','+'+(list.length-5));}" & vbLf
    s = s & "T(cell,'span','band',String(band));}}" & vbLf
    s = s & "T(g,'div','y','');" & vbLf
    s = s & "for(var f2=1;f2<=5;f2++){T(g,'div','xlab',String(f2));}" & vbLf
    s = s & "T(left,'p','axname','縦軸=影響度(上が5)・横軸=発生頻度(右が5)。"
    s = s & "各セル右下の数字は重篤度の帯(1が軽い～5が重い)で、"
    s = s & "白黒印刷や背景色印刷が無効な環境でも重篤度が読めるようにしています。');" & vbLf
    s = s & "var right=T(w,'div',null);" & vbLf
    s = s & "T(right,'h3',null,'経営優先度 Top 8');" & vbLf
    s = s & "var rk=RANKED(D);var pl=T(right,'div','priority-list');" & vbLf
    s = s & "for(var m=0;m<rk.length&&m<8;m++){var z=rk[m];" & vbLf
    s = s & "var ins2=z.insurability||{};var tr2=S(ins2.transferability);" & vbLf
    s = s & "var p=T(pl,'div','priority');" & vbLf
    s = s & "T(p,'div','rid','No.'+S(z.risk_no));" & vbLf
    s = s & "var mid=T(p,'div',null);" & vbLf
    s = s & "T(mid,'div','riskname',S(z.risk_name));" & vbLf
    s = s & "T(mid,'div','why',LB(LCAT,z.category)+' / '+LB(LTR,tr2));" & vbLf
    s = s & "T(p,'div','score',S((z.impact_score||0)*(z.frequency_score||0)));}" & vbLf
    s = s & "if(!rs.length){" & vbLf
    s = s & "T(el,'p','muted','該当するリスクがないため空のマップを表示しています。');return;}" & vbLf
    s = s & "var srt=rs.slice(0);" & vbLf
    s = s & "srt.sort(function(a,b){return (a.risk_no||0)-(b.risk_no||0);});" & vbLf
    s = s & "var rows=[];" & vbLf
    s = s & "for(var q=0;q<srt.length;q++){var v=srt[q];" & vbLf
    s = s & "var ins3=v.insurability||{};var tr3=S(ins3.transferability);" & vbLf
    s = s & "var sp=E('span');" & vbLf
    s = s & "BDG(sp,TRCLS[tr3]?TRCLS[tr3]:'bdg-sub',LB(LTR,tr3));" & vbLf
    s = s & "rows.push([S(v.risk_no),S(v.risk_name),sp]);}" & vbLf
    s = s & "TBL(el,['No','リスク名','移転可能性'],rows);}" & vbLf
    SecRiskMapJs = s
End Function

' SEC-07 risks。s2.risks[] の全項目を1枚の表に出す(横スクロールは .tblwrap 側)。
'   見本にならい**カテゴリの絞り込みボタン**(.filters)を付ける。押すと tbody を
'   作り直す(innerHTML は使わない。18章§4.1)。状態(status)は列で示す。
'   SEC-16 と違い status='new' も一覧には出す。
Public Function SecRisksJs() As String
    Dim s As String
    s = s & "function riskRows(D,list,cat){var rows=[];" & vbLf
    s = s & "for(var i=0;i<list.length;i++){var x=list[i];" & vbLf
    s = s & "if(cat&&S(x.category)!==cat){continue;}" & vbLf
    s = s & "var ins=x.insurability||{};var tr=S(ins.transferability);" & vbLf
    s = s & "var sp=E('span');" & vbLf
    s = s & "BDG(sp,TRCLS[tr]?TRCLS[tr]:'bdg-sub',LB(LTR,tr));" & vbLf
    s = s & "var pv=AR(x.preventions);var pl=[];" & vbLf
    s = s & "for(var k=0;k<pv.length;k++){pl.push(S(pv[k].measure));}" & vbLf
    s = s & "rows.push([S(x.risk_no),LB(LCAT,x.category),S(x.risk_name),S(x.scenario)," & vbLf
    s = s & "LB(LST,x.status),LB(LFQ,x.frequency)+' / '+LB(LIP,x.impact)," & vbLf
    s = s & "S(x.frequency_score)+' x '+S(x.impact_score),sp,S(ins.line_note)," & vbLf
    s = s & "S(ins.control_note),S(x.loss_scale_note),JOIN(x.check_points,'／')," & vbLf
    s = s & "pl.join('／')]);}return rows;}" & vbLf
    s = s & "function renderRisks(D,el){var rs=RISKS(D);if(!rs.length){return;}" & vbLf
    s = s & "var a=rs.slice(0);" & vbLf
    s = s & "a.sort(function(x,y){return (x.risk_no||0)-(y.risk_no||0);});" & vbLf
    s = s & "var cats=[];var seen={};" & vbLf
    s = s & "for(var i=0;i<a.length;i++){var c=S(a[i].category);" & vbLf
    s = s & "if(c&&!seen[c]){seen[c]=1;cats.push(c);}}" & vbLf
    s = s & "var fb=T(el,'div','filters');" & vbLf
    s = s & "var bd=TBL(el,['No','分類','リスク名','想定シナリオ','状態','頻度 / 影響'," & vbLf
    s = s & "'スコア(頻x影)','移転可能性','種目の見立て','管理策','損害規模'," & vbLf
    s = s & "'確認点','未然防止'],riskRows(D,a,''));" & vbLf
    s = s & "var btns=[];" & vbLf
    s = s & "function mkBtn(key,label){var b=T(fb,'button',null,label);" & vbLf
    s = s & "AT(b,'type','button');" & vbLf
    s = s & "b.addEventListener('click',function(){" & vbLf
    s = s & "for(var q=0;q<btns.length;q++){btns[q].className='';}" & vbLf
    s = s & "b.className='active';FILL(bd,riskRows(D,a,key));});" & vbLf
    s = s & "btns.push(b);return b;}" & vbLf
    s = s & "mkBtn('','すべて('+a.length+')').className='active';" & vbLf
    s = s & "for(var j=0;j<cats.length;j++){" & vbLf
    s = s & "mkBtn(cats[j],LB(LCAT,cats[j]));}}" & vbLf
    SecRisksJs = s
End Function

' SEC-08 coverage。2枚組の表(現契約 / ギャップ)。current_coverage が0件
'   (新規案件)なら注記を出してギャップ側の表だけを描く。両方0件なら何も
'   描かない(登録表の empty:'hide' が拾ってセクションごと落とす)。
'   移転可能性の内訳は s2.risks[].insurability.transferability の集計。
Public Function SecCoverageJs() As String
    Dim s As String
    s = s & "function renderCoverage(D,el){var s1=D.s1||{};var s2=D.s2||{};" & vbLf
    s = s & "var cc=AR(s1.current_coverage);var gp=AR(s2.gaps);" & vbLf
    s = s & "if(!cc.length&&!gp.length){return;}" & vbLf
    s = s & "T(el,'h3',null,'現在のご契約');" & vbLf
    ' 18章§3 SEC-08 の注記は逐語(文言を足さない)。
    s = s & "if(!cc.length){T(el,'p','note','新規案件のため現契約なし。"
    s = s & "以下は必要補償の見立て');}" & vbLf
    s = s & "else{var rows=[];" & vbLf
    s = s & "for(var i=0;i<cc.length;i++){var x=cc[i];" & vbLf
    s = s & "rows.push([S(x.line_name),S(x.coverage_summary),S(x.limit_note),"
    s = s & "S(x.special_note)]);}" & vbLf
    s = s & "TBL(el,['種目','補償の概要','限度額','特約・注記'],rows);}" & vbLf
    s = s & "if(gp.length){T(el,'h3',null,'補償のギャップ');var rows2=[];" & vbLf
    s = s & "for(var j=0;j<gp.length;j++){var y=gp[j];" & vbLf
    s = s & "rows2.push([S(y.gap_no),LB(LGAP,y.gap_type),S(y.target),S(y.description)," & vbLf
    s = s & "S(y.risk_evidence),S(y.coverage_evidence)]);}" & vbLf
    s = s & "TBL(el,['No','種別','対象','内容','リスク側の根拠','契約側の根拠'],rows2);}" & vbLf
    s = s & "var rs=RISKS(D);if(!rs.length){return;}" & vbLf
    s = s & "var c={cover:0,partial:0,hard:0};" & vbLf
    s = s & "for(var k=0;k<rs.length;k++){var ins=rs[k].insurability||{};" & vbLf
    s = s & "var t=S(ins.transferability);if(c[t]!==undefined){c[t]++;}}" & vbLf
    s = s & "T(el,'h4',null,'リスクの移転可能性の内訳');" & vbLf
    s = s & "var box=T(el,'div','chips');var keys=['cover','partial','hard'];" & vbLf
    s = s & "for(var m=0;m<keys.length;m++){var sp=E('span','chip');" & vbLf
    s = s & "BDG(sp,TRCLS[keys[m]],LB(LTR,keys[m]));" & vbLf
    s = s & "T(sp,'span',null,' '+c[keys[m]]+'件');box.appendChild(sp);}}" & vbLf
    SecCoverageJs = s
End Function
