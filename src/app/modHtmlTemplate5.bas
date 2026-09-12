Attribute VB_Name = "modHtmlTemplate5"
Option Explicit

' ==========================================================
' modHtmlTemplate5 - SEC-11 prevent / SEC-12 limit / SEC-13 hearing /
'                    SEC-14 source / SEC-15 disclaimer
' ------------------------------------------------
' 18章§4.4の分割規約(「1モジュールが25,000字を超えたら次番のモジュールへ
' 関数単位で切り出す」)により modHtmlTemplate4 から分けたもの。関数名は
' 変えていない(同名の Public Function を2つ以上のモジュールに置かない)。
' 正は18章§3(読むJSONパス)・§3.4(ヒアリング事項の生成規則)・§3.5(免責の固定文)。
'
' SEC-12 は「できないことを正直に書く」(10章FR-37)が存在理由なので、3ブロック
' とも0件でも非表示にしない(登録表 empty:'note' が「該当なし」の1行を出す)。
' SEC-15 は**常に表示**(非表示にできない唯一のセクション)。
'
' R4準拠: Excelトークン・Application.Run・案件データ参照を持たない純文字列。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ==========================================================

' SEC-11 prevent。preventions を全リスクで平らに展開した表。全リスクで0件なら
'   何も描かない(登録表の empty:'hide')。
Public Function SecPreventJs() As String
    Dim s As String
    s = s & "function renderPrevent(D,el){var rs=RISKS(D);var rows=[];" & vbLf
    s = s & "for(var i=0;i<rs.length;i++){var x=rs[i];var pv=AR(x.preventions);" & vbLf
    s = s & "for(var k=0;k<pv.length;k++){" & vbLf
    s = s & "rows.push([S(x.risk_no),S(x.risk_name),S(pv[k].measure)," & vbLf
    s = s & "S(pv[k].related_menu_id)]);}}" & vbLf
    s = s & "if(!rows.length){return;}" & vbLf
    s = s & "T(el,'p','muted','保険の前に手を打てるもの。関連メニューIDがある行は"
    s = s & "当社の実在サービスで支援できます。');" & vbLf
    s = s & "TBL(el,['対象No','リスク名','未然防止の打ち手','関連メニューID'],rows);}" & vbLf
    SecPreventJs = s
End Function

' SEC-12 limit。3ブロック(移転しにくいリスク / 受けきれないリスク / 提案を
'   控えること)。10章FR-37「できないことを正直に書く」がこの節の存在理由。
Public Function SecLimitJs() As String
    Dim s As String
    s = s & "function renderLimit(D,el){var rs=RISKS(D);var s3=D.s3||{};" & vbLf
    s = s & "var hard=[];" & vbLf
    s = s & "for(var i=0;i<rs.length;i++){var ins=rs[i].insurability||{};" & vbLf
    s = s & "if(S(ins.transferability)==='hard'){hard.push(rs[i]);}}" & vbLf
    s = s & "var um=AR(s3.unmatched_risks);var dn=AR(s3.do_not_propose);" & vbLf
    s = s & "if(!hard.length&&!um.length&&!dn.length){return;}" & vbLf
    s = s & "if(hard.length){T(el,'h3',null,'保険では移転しにくいリスク');var rows=[];" & vbLf
    s = s & "for(var j=0;j<hard.length;j++){var y=hard[j];var i2=y.insurability||{};" & vbLf
    s = s & "rows.push([S(y.risk_no),S(y.risk_name),S(i2.control_note)]);}" & vbLf
    s = s & "TBL(el,['No','リスク名','保険の代わりに効く管理策'],rows);}" & vbLf
    s = s & "if(um.length){T(el,'h3',null,'当社のメニュー・型では受けきれないもの');" & vbLf
    s = s & "var rows2=[];" & vbLf
    s = s & "for(var k=0;k<um.length;k++){" & vbLf
    s = s & "rows2.push([S(um[k].risk_no),S(um[k].risk_name),S(um[k].why_unmatched)]);}" & vbLf
    s = s & "TBL(el,['No','リスク名','受けきれない理由'],rows2);}" & vbLf
    s = s & "if(dn.length){T(el,'h3',null,'今回は提案を控えること');var rows3=[];" & vbLf
    s = s & "for(var m=0;m<dn.length;m++){" & vbLf
    s = s & "rows3.push([S(dn[m].topic),S(dn[m].reason)]);}" & vbLf
    s = s & "TBL(el,['項目','控える理由'],rows3);}}" & vbLf
    SecLimitJs = s
End Function

' SEC-13 hearing(18章§3.4)。4系統をこの順に連結して通し番号を振る:
'   (1)提案の切り口 hook_question(story_no昇順・3問まで)
'   (2)未解決の論点 open_questions(配列順・5問まで)
'   (3)不足情報 missing_info(配列順・5問まで。設問文は「{item}について
'      教えてください」・補足に why_needed)
'   (4)リスクの確認点 check_points(重要度上位5件のリスク・1件2点まで・10問まで)
'   合計20問を上限とし、超過分は切り捨てて末尾に「ほか{n}問」の1行を出す。
'   同一文字列は先に出た系統を残す(前後空白を除去した完全一致)。
Public Function SecHearingJs() As String
    Dim s As String
    s = s & "function renderHearing(D,el){var s1=D.s1||{};var s2=D.s2||{};" & vbLf
    s = s & "var s3=D.s3||{};var items=[];var seen={};" & vbLf
    s = s & "function add(cat,q,why){var t=S(q).replace(/^\s+|\s+$/g,'');" & vbLf
    s = s & "if(!t.length||seen[t]){return false;}" & vbLf
    s = s & "seen[t]=1;items.push({c:cat,q:t,w:S(why)});return true;}" & vbLf
    s = s & "var st=AR(s3.stories).slice(0);" & vbLf
    s = s & "st.sort(function(a,b){return (a.story_no||0)-(b.story_no||0);});" & vbLf
    s = s & "for(var i=0;i<st.length&&i<3;i++){add('提案の切り口',st[i].hook_question,'');}" & vbLf
    s = s & "var oq=AR(s2.open_questions);" & vbLf
    s = s & "for(var j=0;j<oq.length&&j<5;j++){add('未解決の論点',oq[j],'');}" & vbLf
    s = s & "var mi=AR(s1.missing_info);" & vbLf
    s = s & "for(var k=0;k<mi.length&&k<5;k++){" & vbLf
    s = s & "add('不足情報',S(mi[k].item)+'について教えてください',S(mi[k].why_needed));}" & vbLf
    s = s & "var rk=RANKED(D);var cp=0;" & vbLf
    s = s & "for(var m=0;m<rk.length&&m<5;m++){var ps=AR(rk[m].check_points);" & vbLf
    s = s & "for(var n=0;n<ps.length&&n<2;n++){if(cp>=10){break;}" & vbLf
    s = s & "if(add('リスクの確認点',ps[n],'No.'+S(rk[m].risk_no)+' '+S(rk[m].risk_name)))" & vbLf
    s = s & "{cp++;}}}" & vbLf
    s = s & "if(!items.length){return;}" & vbLf
    ' 見本の2段組カード(.questions の columns:2)。通し番号は Q{n}. で振る
    ' (18章§3.4 の順序と件数の規則は不変)。
    s = s & "var box=T(el,'div','questions');" & vbLf
    s = s & "for(var p=0;p<items.length&&p<20;p++){var q=T(box,'div','q');" & vbLf
    s = s & "T(q,'b',null,'Q'+(p+1)+'. ');" & vbLf
    s = s & "T(q,'span',null,items[p].q);" & vbLf
    s = s & "T(q,'span','qcat',items[p].c);" & vbLf
    s = s & "if(NB(items[p].w)){T(q,'div','qwhy',items[p].w);}}" & vbLf
    s = s & "if(items.length>20){" & vbLf
    ' 18章§3.4 の超過1行は逐語(括弧も全角)。
    s = s & "T(el,'p','muted','ほか'+(items.length-20)+'問（ヒアリングシートを参照）');}}" & vbLf
    SecHearingJs = s
End Function

' SEC-14 source。リスクごとの引用と出所(19章§3の source enum を日本語へ)。
Public Function SecSourceJs() As String
    Dim s As String
    s = s & "function renderSource(D,el){var rs=RISKS(D);var rows=[];" & vbLf
    s = s & "var a=rs.slice(0);" & vbLf
    s = s & "a.sort(function(x,y){return (x.risk_no||0)-(y.risk_no||0);});" & vbLf
    s = s & "var gu=AR((D.meta||{}).ground_unmatched);var gm={};" & vbLf
    s = s & "for(var g=0;g<gu.length;g++){gm[S(gu[g])]=1;}" & vbLf
    s = s & "for(var i=0;i<a.length;i++){var ev=a[i].evidence||{};" & vbLf
    s = s & "if(!NB(ev.quote)&&!NB(ev.source)){continue;}" & vbLf
    s = s & "rows.push([S(a[i].risk_no),S(a[i].risk_name),S(ev.quote)," & vbLf
    s = s & "LB(LSRC,ev.source),gm[S(a[i].risk_no)]?'原文未照合':'']);}" & vbLf
    ' 裁定書38 B-10: 成功事例が該当N件のうちM件だけ使われたときの1行(静かな
    ' 打切りを可視化する。totalが無い/使用数以下なら何も出さない。引用が0件の
    ' 案件でも成立しうるので rows.length の判定より前に出す)。
    s = s & "var ku=(D.meta||{}).kb_usage;" & vbLf
    s = s & "if(ku&&NB(ku.cases_total)&&ku.cases_total>ku.cases_used){" & vbLf
    s = s & "T(el,'p','muted','成功事例は該当'+ku.cases_total+'件のうち'+ku.cases_used+'件を使用しています。');}" & vbLf
    s = s & "if(!rows.length){return;}" & vbLf
    s = s & "T(el,'p','muted','各リスクの根拠にした記述と、その出どころです。"
    s = s & "「推定」は入力に直接の記述が無く当社が置いた仮定であることを示します。"
    s = s & "「原文未照合」は、貼り付けた資料の中にその記述を見つけられなかった"
    s = s & "ことを示します（表記の違いで見つからないこともあります）。');" & vbLf
    s = s & "TBL(el,['No','リスク名','引用した記述','出所','原文照合'],rows);}" & vbLf
    SecSourceJs = s
End Function

' SEC-15 disclaimer(18章§3.5)。この4行を必ず含める。**常に表示**。
'   1行目は16章NFR-S5の必須表記(<noscript>側にも同じ1行を静的に置いてある)。
'   裁定書37 B-06: 1行目は meta.reviewed_by の**空/非空で3項分岐**する。担当者の
'   確認を通していない書き出しで「人が確認・編集した」と名乗らない(社内IT環境
'   v1.1 §7.3 は顧客提示物に利用者の確認を必須と定めるが、製品側にその担保が
'   無かった)。2行目以降(仮説である旨・保険料試算は対象外・署名)は不変。
Public Function SecDisclaimerJs() As String
    Dim s As String
    s = s & "function renderDisclaimer(D,el){var m=D.meta||{};" & vbLf
    s = s & "var d=T(el,'div','disc');var rb=S(m.reviewed_by);" & vbLf
    s = s & "T(d,'p',null,rb?('本資料はAI支援により作成した骨子を担当者が確認・"
    s = s & "編集したものです（確認: '+rb+' / '+S(m.reviewed_at)+'）。'):"
    s = s & "'本資料はAIが公開情報等から作成した営業担当者向けの分析資料です"
    s = s & "（AI生成・担当者確認前）。お客さまへ提示する前に、担当者が内容を"
    s = s & "確認・編集してください。');" & vbLf
    s = s & "T(d,'p',null,'記載のリスクは公開情報と当社担当者の見立てに基づく仮説で"
    s = s & "あり、引受可否・保険料・幹事構成を確約するものではありません。');" & vbLf
    s = s & "T(d,'p',null,'保険料の試算は本資料の対象外です（要見積）。');" & vbLf
    s = s & "T(d,'p',null,S(m.company)+' 御中 / 案件ID '+S(m.case_id)+' / 作成 '" & vbLf
    s = s & "+S(m.generated_at)+' / リスク提案ナビ v'+S(m.app_version));}" & vbLf
    SecDisclaimerJs = s
End Function
