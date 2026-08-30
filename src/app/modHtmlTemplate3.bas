Attribute VB_Name = "modHtmlTemplate3"
Option Explicit

' ==========================================================
' modHtmlTemplate3 - SEC-05 riskuniv / SEC-06 riskmap / SEC-07 risks /
'                    SEC-08 coverage の描画スクリプト
' ------------------------------------------------
' 正は18章§3(読むJSONパス)・§3.2(10分類の描き方)・§3.3(5x5マップの描き方)・
' §4.4(分割規約の既定割り当て)。CSSも登録行も持たない(18章§4.3)。
' 描画は createElement / textContent / setAttribute だけ(18章§4.1)。
'
' R4準拠: Excelトークン・Application.Run・案件データ参照を持たない純文字列。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ==========================================================

' SEC-05 riskuniv(18章§3.2)。**常に10行**並べ、該当0件の分類も0件と表示する。
'   「見ていない領域」と「見たが該当なし」を読み手が区別できるようにするため。
'   並びは15章§0の変換表の記載順(CATORDER)に固定する。
Public Function SecRiskUnivJs() As String
    Dim s As String
    s = s & "function renderRiskUniv(D,el){var rs=RISKS(D);" & vbLf
    s = s & "var cnt={};var tot=rs.length;" & vbLf
    s = s & "for(var i=0;i<rs.length;i++){var c=S(rs[i].category);" & vbLf
    s = s & "cnt[c]=(cnt[c]||0)+1;}" & vbLf
    s = s & "for(var j=0;j<CATORDER.length;j++){var k=CATORDER[j];" & vbLf
    s = s & "var n=cnt[k]||0;var row=T(el,'div','sec-riskuniv-row');" & vbLf
    s = s & "T(row,'span',null,LB(LCAT,k));" & vbLf
    s = s & "T(row,'span','muted',n+'件');" & vbLf
    s = s & "var bar=T(row,'div','sec-riskuniv-bar');" & vbLf
    s = s & "var fill=T(bar,'div','sec-riskuniv-fill');" & vbLf
    s = s & "AT(fill,'style','width:'+((tot>0)?Math.round(n*100/tot):0)+'%');}" & vbLf
    s = s & "T(el,'p','muted','全'+tot+'件。0件の分類も「見たうえで該当なし」"
    s = s & "として並べています。');}" & vbLf
    SecRiskUnivJs = s
End Function

' SEC-06 riskmap(18章§3.3)。縦軸=impact_score(上が5)・横軸=frequency_score
'   (左が1)の5行5列。セル背景の帯は impact+frequency の和で決め(2-3=1 /
'   4-5=2 / 6-7=3 / 8-9=4 / 10=5 -> floor(sum/2))、**色に依存させない**ため
'   各セルの右下に帯番号を必ず添える。バッジ色は transferability。
'   1セルに6件以上入る場合は先頭5件+「+n」に切り替える(5x5の形を保つ)。
Public Function SecRiskMapJs() As String
    Dim s As String
    s = s & "function renderRiskMap(D,el){var rs=RISKS(D);" & vbLf
    s = s & "var w=T(el,'div','sec-riskmap-wrap');" & vbLf
    s = s & "var g=T(w,'div','sec-riskmap-grid');" & vbLf
    s = s & "for(var imp=5;imp>=1;imp--){" & vbLf
    s = s & "T(g,'div','sec-riskmap-ax',String(imp));" & vbLf
    s = s & "for(var frq=1;frq<=5;frq++){var band=Math.floor((imp+frq)/2);" & vbLf
    s = s & "var cell=T(g,'div','sec-riskmap-cell heat'+band);var list=[];" & vbLf
    s = s & "for(var i=0;i<rs.length;i++){var x=rs[i];" & vbLf
    s = s & "if((x.impact_score||0)===imp&&(x.frequency_score||0)===frq){list.push(x);}}" & vbLf
    s = s & "list.sort(function(a,b){return (a.risk_no||0)-(b.risk_no||0);});" & vbLf
    s = s & "for(var k=0;k<list.length&&k<5;k++){var y=list[k];" & vbLf
    s = s & "var ins=y.insurability||{};var tr=S(ins.transferability);" & vbLf
    s = s & "var b=T(cell,'span','rb '+(TRCLS[tr]?TRCLS[tr]:'bdg-sub'),String(y.risk_no));" & vbLf
    s = s & "AT(b,'title',S(y.risk_name));}" & vbLf
    s = s & "if(list.length>5){T(cell,'span','muted','+'+(list.length-5));}" & vbLf
    s = s & "T(cell,'span','sec-riskmap-band',String(band));}}" & vbLf
    s = s & "T(g,'div','sec-riskmap-ax','');" & vbLf
    s = s & "for(var f2=1;f2<=5;f2++){T(g,'div','sec-riskmap-ax',String(f2));}" & vbLf
    s = s & "T(w,'p','sec-riskmap-axname','縦軸=影響度(上が5)・横軸=発生頻度(右が5)。"
    s = s & "各セル右下の数字は重篤度の帯(1が軽い～5が重い)で、"
    s = s & "白黒印刷や背景色印刷が無効な環境でも重篤度が読めるようにしています。');" & vbLf
    s = s & "if(!rs.length){" & vbLf
    s = s & "T(el,'p','muted','該当するリスクがないため空のマップを表示しています。');return;}" & vbLf
    s = s & "var srt=rs.slice(0);" & vbLf
    s = s & "srt.sort(function(a,b){return (a.risk_no||0)-(b.risk_no||0);});" & vbLf
    s = s & "var rows=[];" & vbLf
    s = s & "for(var m=0;m<srt.length;m++){var z=srt[m];" & vbLf
    s = s & "var ins2=z.insurability||{};var tr2=S(ins2.transferability);" & vbLf
    s = s & "var sp=E('span');" & vbLf
    s = s & "BDG(sp,TRCLS[tr2]?TRCLS[tr2]:'bdg-sub',LB(LTR,tr2));" & vbLf
    s = s & "rows.push([S(z.risk_no),S(z.risk_name),sp]);}" & vbLf
    s = s & "TBL(el,['No','リスク名','移転可能性'],rows);}" & vbLf
    SecRiskMapJs = s
End Function

' SEC-07 risks。s2.risks[] の全項目を1枚の表に出す(横スクロール可)。
'   状態(status)はバッジで示す。SEC-16 と違い status='new' も一覧には出す。
Public Function SecRisksJs() As String
    Dim s As String
    s = s & "function renderRisks(D,el){var rs=RISKS(D);if(!rs.length){return;}" & vbLf
    s = s & "var a=rs.slice(0);" & vbLf
    s = s & "a.sort(function(x,y){return (x.risk_no||0)-(y.risk_no||0);});" & vbLf
    s = s & "var rows=[];" & vbLf
    s = s & "for(var i=0;i<a.length;i++){var x=a[i];var ins=x.insurability||{};" & vbLf
    s = s & "var tr=S(ins.transferability);var sp=E('span');" & vbLf
    s = s & "BDG(sp,TRCLS[tr]?TRCLS[tr]:'bdg-sub',LB(LTR,tr));" & vbLf
    s = s & "var pv=AR(x.preventions);var pl=[];" & vbLf
    s = s & "for(var k=0;k<pv.length;k++){pl.push(S(pv[k].measure));}" & vbLf
    s = s & "rows.push([S(x.risk_no),LB(LCAT,x.category),S(x.risk_name),S(x.scenario)," & vbLf
    s = s & "LB(LST,x.status),LB(LFQ,x.frequency)+' / '+LB(LIP,x.impact)," & vbLf
    s = s & "S(x.frequency_score)+' x '+S(x.impact_score),sp,S(ins.line_note)," & vbLf
    s = s & "S(ins.control_note),S(x.loss_scale_note),JOIN(x.check_points,'／')," & vbLf
    s = s & "pl.join('／')]);}" & vbLf
    s = s & "TBL(el,['No','分類','リスク名','想定シナリオ','状態','頻度 / 影響'," & vbLf
    s = s & "'スコア(頻x影)','移転可能性','種目の見立て','管理策','損害規模'," & vbLf
    s = s & "'確認点','未然防止'],rows);}" & vbLf
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
