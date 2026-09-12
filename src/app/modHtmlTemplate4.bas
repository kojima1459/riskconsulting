Attribute VB_Name = "modHtmlTemplate4"
Option Explicit

' ==========================================================
' modHtmlTemplate4 - SEC-09 newrisk / SEC-16 round-update / SEC-10 story
' ------------------------------------------------
' 正は18章§3(読むJSONパス・空のときの挙動)・§4.4(分割規約)。18章§4.4の既定
' 割り当てはSEC-09からSEC-15までを1モジュールに置くが、25,000字を超えるため
' §4.4の分割規約(「25,000字を超えたら次番のモジュールへ関数単位で切り出す」)
' に従い SEC-11からSEC-15を modHtmlTemplate5 へ分けた。関数名は変えていない。
'
' SEC-09 と SEC-16 は別物(18章§3の注記):
'   SEC-09 = s2.emerging_risks[]。サイバー・気候変動のような**新種・新興リスク**
'            でラウンドに関係なく初回から出る。0件でも「見たうえで該当が無い」
'            ことを示すため非表示にしない(登録表の empty:'note')。
'   SEC-16 = s2.risks[].status。**第2ラウンド以降の仮説ライフサイクル**
'            (new / confirmed / rejected)。status='new' のリスクは SEC-16 と
'            SEC-07 のバッジに留め、SEC-09 には出さない。
'
' R4準拠: Excelトークン・Application.Run・案件データ参照を持たない純文字列。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ==========================================================

' SEC-09 newrisk。見本の2列レーダーカード(.radar。horizon のピル+強度バー。
'   11章§3.8.1)。0件のときは何も描かず、登録表の empty:'note' が
'   「現時点で特筆すべき新しく出てきたリスクは検出されていません」の1行を出す。
Public Function SecNewRiskJs() As String
    Dim s As String
    s = s & "function renderNewRisk(D,el){var s2=D.s2||{};" & vbLf
    s = s & "var er=AR(s2.emerging_risks);if(!er.length){return;}" & vbLf
    ' 裁定書39 R1-08: modGround.GroundNotes は emerging_risks を配列の出現順で
    ' "E1" "E2" と採番して meta.ground_unmatched へ入れる(18章§2)。SEC-14 は
    ' risks[] しか描かないので、ニューリスクの未照合はこの節で出す。捏造が最も
    ' 出やすいのが emerging_risks(一般論を書きやすい)なので、見せたい所に出す。
    s = s & "var gu=AR((D.meta||{}).ground_unmatched);var gm={};" & vbLf
    s = s & "for(var g2=0;g2<gu.length;g2++){gm[S(gu[g2])]=1;}" & vbLf
    s = s & "var g=T(el,'div','radar');" & vbLf
    s = s & "for(var i=0;i<er.length;i++){var x=er[i];" & vbLf
    s = s & "var it=T(g,'div','radar-item');" & vbLf
    s = s & "var top=T(it,'div','radar-top');" & vbLf
    s = s & "T(top,'h4',null,S(x.risk_name));" & vbLf
    s = s & "T(top,'span','horizon',LB(LHZ,x.horizon));" & vbLf
    s = s & "T(it,'div','mini',LB(LCAT,x.category));" & vbLf
    s = s & "T(it,'p',null,S(x.scenario));" & vbLf
    s = s & "var bar=T(it,'div','bar');var fill=T(bar,'span',null);" & vbLf
    s = s & "AT(fill,'style','width:'+(HZW[S(x.horizon)]||40)+'%');" & vbLf
    ' 18章§6(3): 色とバーの長さだけで意味を運ばないので文字でも書く。
    s = s & "T(it,'div','mini','近さの目安: '+LB(LHZ,x.horizon));" & vbLf
    s = s & "if(NB(x.evidence_quote)){T(it,'p','mini','根拠: 「'+S(x.evidence_quote)" & vbLf
    s = s & "+'」('+LB(LSRC,x.evidence_source)+')');}" & vbLf
    ' 裁定書39 R1-08: 採番は modGround と同じ「配列の出現順」。文言は SEC-14 の
    ' 「原文未照合」と同じ語にし、読み手が2箇所で同じ意味に読めるようにする。
    s = s & "if(gm['E'+(i+1)]){T(it,'div','mini','原文未照合: 貼り付けた資料の中に"
    s = s & "この根拠を見つけられませんでした（表記の違いで見つからないことも"
    s = s & "あります）。');}" & vbLf
    s = s & "if(NB(x.proposal_hint)){" & vbLf
    s = s & "T(it,'p','mini','提案の糸口: '+S(x.proposal_hint));}}}" & vbLf
    SecNewRiskJs = s
End Function

' SEC-16 round-update。round_no が2未満、または3つのstatusがいずれも0件なら
'   何も描かない(登録表の empty:'hide' がセクションごと目次からも落とす)。
'   rejected は見出しに取り消し表現を付し、scenario 末尾の否定理由はそのまま残す。
Public Function SecRoundUpdateJs() As String
    Dim s As String
    s = s & "function ruBlock(el,title,list,cls,strike){if(!list.length){return;}" & vbLf
    s = s & "var b=T(el,'div','rublock');" & vbLf
    s = s & "T(b,'div','rubh'+(cls?' '+cls:''),title+' '+list.length+'件');" & vbLf
    s = s & "for(var i=0;i<list.length;i++){var x=list[i];" & vbLf
    s = s & "var p=T(b,'p',null);" & vbLf
    s = s & "T(p,'span',strike?'strike':null,'No.'+S(x.risk_no)+' '+S(x.risk_name));" & vbLf
    s = s & "T(p,'span','muted','  '+LB(LCAT,x.category));" & vbLf
    s = s & "T(b,'p','muted',S(x.scenario));}}" & vbLf
    s = s & "function renderRoundUpdate(D,el){var m=D.meta||{};" & vbLf
    s = s & "if((m.round_no||0)<2){return;}var rs=RISKS(D);" & vbLf
    s = s & "var g={'new':[],confirmed:[],rejected:[]};" & vbLf
    s = s & "for(var i=0;i<rs.length;i++){var st=S(rs[i].status);" & vbLf
    s = s & "if(g[st]){g[st].push(rs[i]);}}" & vbLf
    s = s & "if(!g['new'].length&&!g.confirmed.length&&!g.rejected.length){return;}" & vbLf
    s = s & "T(el,'p','muted','第'+S(m.round_no)+'ラウンド。訪問と追加情報でリスク仮説が"
    s = s & "どう動いたかを示します。');" & vbLf
    s = s & "ruBlock(el,'新たに浮上した仮説',g['new'],'rubh-new',false);" & vbLf
    s = s & "ruBlock(el,'裏が取れたリスク',g.confirmed,'',false);" & vbLf
    s = s & "ruBlock(el,'否定された仮説(記録は残します)',g.rejected,'rubh-rej',true);}" & vbLf
    SecRoundUpdateJs = s
End Function

' SEC-10 story。見本の濃色の提案ブロック(.proposal)+ STEP1から4の流れ +
'   3枚のカード(11章§3.8.1)。target_risk_nos / target_gap_nos は s2 側を引いて
'   名称に解決する(番号だけを見せない)。hook_question / expected_objection /
'   objection_response の扱いは不変(18章§3)。
Public Function SecStoryJs() As String
    Dim s As String
    s = s & "function pfStep(p,noText,titleText,noteText){var b=T(p,'div','pf');" & vbLf
    s = s & "T(b,'div','n',noText);T(b,'b',null,titleText);" & vbLf
    s = s & "T(b,'span',null,noteText);}" & vbLf
    s = s & "function renderStory(D,el){var s3=D.s3||{};" & vbLf
    s = s & "var st=AR(s3.stories);if(!st.length){return;}" & vbLf
    s = s & "var a=st.slice(0);" & vbLf
    s = s & "a.sort(function(x,y){return (x.story_no||0)-(y.story_no||0);});" & vbLf
    s = s & "var pb=T(el,'div','proposal');" & vbLf
    s = s & "T(pb,'h3',null,'保険の見直しではなく、事業を止めないための打ち手として話す');" & vbLf
    s = s & "T(pb,'p',null,'保険種目の羅列ではなく、経営アジェンダ・リスク・打ち手・保険の順に会話します。');" & vbLf
    s = s & "var fl=T(pb,'div','proposal-flow');" & vbLf
    s = s & "pfStep(fl,'STEP 1','経営アジェンダ','守りたいもの・伸ばしたいものを確かめる');" & vbLf
    s = s & "pfStep(fl,'STEP 2','リスク','事業構造から生まれるリスクを共有する');" & vbLf
    s = s & "pfStep(fl,'STEP 3','打ち手','未然防止と当社メニューで先に手を打つ');" & vbLf
    s = s & "pfStep(fl,'STEP 4','保険','残った損害を保険で引き受ける');" & vbLf
    s = s & "for(var i=0;i<a.length;i++){var y=a[i];" & vbLf
    s = s & "var card=T(el,'div','card card-matsu');" & vbLf
    s = s & "var c=T(card,'div','chips');" & vbLf
    s = s & "CHIP(c,'テーマ'+S(y.story_no),'chip-brand');" & vbLf
    s = s & "CHIP(c,LB(LPK,y.proposal_kind));" & vbLf
    s = s & "T(card,'h3',null,S(y.headline));" & vbLf
    s = s & "if(NB(y.hook_question)){T(card,'p','hook','「'+S(y.hook_question)+'」');}" & vbLf
    s = s & "PARA(card,y.pitch,null);" & vbLf
    s = s & "var rn=[];var ta=AR(y.target_risk_nos);" & vbLf
    s = s & "for(var k=0;k<ta.length;k++){rn.push(RNAME(D,ta[k]));}" & vbLf
    s = s & "var gn=[];var tg=AR(y.target_gap_nos);" & vbLf
    s = s & "for(var m=0;m<tg.length;m++){gn.push(GNAME(D,tg[m]));}" & vbLf
    s = s & "DLIST(card,[['対象リスク',rn.join('・')],['対象ギャップ',gn.join('・')]," & vbLf
    s = s & "['関連メニュー',JOIN(y.menu_ids,'／')],['関連種目',JOIN(y.line_ids,'／')]," & vbLf
    s = s & "['座組',S(y.scheme_id)],['類似事例',S(y.similar_case_id)]," & vbLf
    s = s & "['想定される反論',S(y.expected_objection)]," & vbLf
    s = s & "['お応えの仕方',S(y.objection_response)]]);}}" & vbLf
    SecStoryJs = s
End Function
