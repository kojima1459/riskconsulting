Attribute VB_Name = "modProposalHtml4"
Option Explicit

' ==========================================================
' modProposalHtml4 - 提案書 Wide の描画(13枚目から22枚目)。20章§4が正。
' ------------------------------------------------
'   13 ideas / 14 four / 16 themes / 17 steps / 18 decide / 19 share /
'   20 appendix / 21 premise / 22 back
' 18章§4.4 の25,000字規約に倣い modProposalHtml3 から分けた(関数名は変えない)。
' 免責の固定文は**1本だけ**(裁定書38 §1 班C): 20章§8。値源は本モジュールの
'   DisclaimerText() 1箇所で、21枚目と22枚目が同じ関数から取る。
' R4準拠・CP932準拠。
' ==========================================================

' 顧客向け提案書の免責(唯一の固定文。裁定書38 §1 班C・20章§8)。
'   保険料の試算・引受条件を確約しないことだけを述べ、AI利用の表示は入れない
'   (社内IT・AI環境 v1.1 §7.3: 担当者の確認を経た顧客提示物。確認していない
'   資料はそもそも出力できない=modExportProposal が reviewedBy 空を断る)。
Public Function DisclaimerText() As String
    DisclaimerText = "本資料は、引受・保険料・契約条件を確約するものではありません。"
End Function

' 13枚目 成長支援アイデアの一覧(優先候補に印)。
Public Function SlIdeasJs() As String
    Dim s As String
    s = s & "function renderIdeas(D,el){var id=AR(P(D).ideas);" & vbLf
    s = s & "if(!id.length){return;}HEAD(el,HL(D,'ideas'));" & vbLf
    s = s & "var rows=[];var pri=[];" & vbLf
    s = s & "for(var i=0;i<id.length;i++){var x=id[i];pri.push(!!x.priority);" & vbLf
    s = s & "rows.push([(x.priority?'優先':''),S(x.title),S(x.aim)," & vbLf
    s = s & "STARS(x.effect),S(x.difficulty)]);}" & vbLf
    s = s & "TBL(el,['','案','狙い','貴社の事業効果','実現までの難易度'],rows,pri);" & vbLf
    s = s & "T(el,'p','muted','実現までの難易度には、商品の準備や関係先との"
    s = s & "調整に要する期間を含みます。');}" & vbLf
    SlIdeasJs = s
End Function

' 14枚目 優先4案(狙い・仕組み・想定する保険)。
Public Function SlFourJs() As String
    Dim s As String
    s = s & "function renderFour(D,el){var fo=AR(P(D).four);" & vbLf
    s = s & "if(!fo.length){return;}HEAD(el,HL(D,'four'));" & vbLf
    s = s & "var cols=T(el,'div','cols');" & vbLf
    s = s & "var mark='ABCD';" & vbLf
    s = s & "for(var i=0;i<fo.length;i++){var col=T(cols,'div','col');" & vbLf
    s = s & "var card=T(col,'div','card');" & vbLf
    s = s & "T(card,'h3',null,mark.charAt(i)+'. '+S(fo[i].title));" & vbLf
    s = s & "T(card,'p','lead',S(fo[i].aim));" & vbLf
    s = s & "T(card,'p',null,S(fo[i].mechanism));" & vbLf
    s = s & "T(card,'p',null,'想定する保険: '+S(fo[i].insurance));}}" & vbLf
    SlFourJs = s
End Function

' 16枚目 提案テーマ(経営課題・保険・経営指標を束ねる表)。
Public Function SlThemesJs() As String
    Dim s As String
    s = s & "function renderThemes(D,el){var tt=AR(P(D).theme_table);" & vbLf
    s = s & "if(!tt.length){return;}HEAD(el,HL(D,'themes'));" & vbLf
    s = s & "var rows=[];" & vbLf
    s = s & "for(var i=0;i<tt.length;i++){" & vbLf
    s = s & "rows.push([S(tt[i].theme),S(tt[i].issues),S(tt[i].insurance),S(tt[i].kpi)]);}" & vbLf
    s = s & "TBL(el,['テーマ','経営の課題','対応する保険（一般名称）','経営指標']," & vbLf
    s = s & "rows,null);}" & vbLf
    SlThemesJs = s
End Function

' 17枚目 進め方(4つのステップ。役割と費用の扱いを明示する)。
Public Function SlStepsJs() As String
    Dim s As String
    s = s & "function renderSteps(D,el){var sp=AR(P(D).steps);" & vbLf
    s = s & "if(!sp.length){return;}HEAD(el,HL(D,'steps'));" & vbLf
    s = s & "var box=T(el,'div','steps');" & vbLf
    s = s & "for(var i=0;i<sp.length;i++){var b=T(box,'div','step');" & vbLf
    s = s & "T(b,'b',null,'STEP '+String(i+1)+'  '+S(sp[i].who));" & vbLf
    s = s & "T(b,'h3',null,S(sp[i].title));" & vbLf
    s = s & "T(b,'p',null,S(sp[i].desc));" & vbLf
    s = s & "T(b,'div','ref','本資料での該当: '+S(sp[i].ref));}" & vbLf
    s = s & "T(el,'p','muted','費用の目安は、正式なご提案の際にお示しします。');}" & vbLf
    SlStepsJs = s
End Function

' 18枚目 本日ご判断いただきたいこと(3点)。最終結論を求めない旨を添える。
Public Function SlDecideJs() As String
    Dim s As String
    s = s & "function renderDecide(D,el){var dc=AR(P(D).decisions);" & vbLf
    s = s & "if(!dc.length){return;}HEAD(el,HL(D,'decide'));" & vbLf
    s = s & "var cols=T(el,'div','cols');" & vbLf
    s = s & "for(var i=0;i<dc.length;i++){var col=T(cols,'div','col');" & vbLf
    s = s & "var card=T(col,'div','card');" & vbLf
    s = s & "T(card,'h3',null,String(i+1)+'. '+S(dc[i].title));" & vbLf
    s = s & "T(card,'p',null,S(dc[i].options));" & vbLf
    s = s & "T(card,'p','muted','確認したい点: '+S(dc[i].note));}" & vbLf
    s = s & "T(el,'p','muted','本日は最終のご結論をお願いするものではありません。');}" & vbLf
    SlDecideJs = s
End Function

' 19枚目 ご共有いただきたい事項(優先の印つき)。
Public Function SlShareJs() As String
    Dim s As String
    s = s & "function renderShare(D,el){var sh=AR(P(D).share_items);" & vbLf
    s = s & "if(!sh.length){return;}HEAD(el,HL(D,'share'));" & vbLf
    s = s & "var rows=[];var pri=[];" & vbLf
    s = s & "for(var i=0;i<sh.length;i++){pri.push(!!sh[i].priority);" & vbLf
    s = s & "rows.push([(sh[i].priority?'優先':''),String(i+1)," & vbLf
    s = s & "S(sh[i].group),S(sh[i].text)]);}" & vbLf
    s = s & "TBL(el,['','','区分','ご共有いただきたい事項'],rows,pri);" & vbLf
    s = s & "T(el,'p','muted','「優先」の項目からご共有いただければ、"
    s = s & "評価の更新に着手できます（差し支えない範囲で結構です）。');}" & vbLf
    SlShareJs = s
End Function

' 20枚目 巻末資料: リスク一覧(全件)。番号はここが値源(本文の番号はここを引く)。
Public Function SlAppendixJs() As String
    Dim s As String
    s = s & "function renderAppendix(D,el){var rs=RK(D);if(!rs.length){return;}" & vbLf
    s = s & "HEAD(el,HL(D,'appendix'));" & vbLf
    s = s & "var rows=[];var pri=[];" & vbLf
    s = s & "for(var i=0;i<rs.length;i++){var x=rs[i];pri.push(i<8);" & vbLf
    s = s & "rows.push(['No.'+S(x.risk_no),S(x.category_label),S(x.risk_name)," & vbLf
    s = s & "S(x.rank_label),S(x.class_label)]);}" & vbLf
    s = s & "TBL(el,['番号','カテゴリー','リスク','評価','備えやすさ'],rows,pri);}" & vbLf
    SlAppendixJs = s
End Function

' 21枚目 本資料の前提。前提は提案書JSONの premise、免責は固定文1本。
Public Function SlPremiseJs() As String
    Dim s As String
    s = s & "function renderPremise(D,el){" & vbLf
    s = s & "T(el,'p','sub','前提・留意事項');" & vbLf
    s = s & "if(NB(P(D).premise)){T(el,'p','lead',S(P(D).premise));}" & vbLf
    s = s & "T(el,'p','lead',DISCLAIMER);" & vbLf
    s = s & "var m=D.meta||{};" & vbLf
    s = s & "T(el,'p','muted','作成 '+S(m.date)+'  担当 '+S(m.reviewed_by));}" & vbLf
    SlPremiseJs = s
End Function

' 22枚目 裏表紙。
Public Function SlBackJs() As String
    Dim s As String
    s = s & "var DISCLAIMER='" & DisclaimerText() & "';" & vbLf
    s = s & "function renderBack(D,el){" & vbLf
    s = s & "T(el,'div','back-mark','ご清聴ありがとうございました');" & vbLf
    s = s & "T(el,'div','back-note',DISCLAIMER);}" & vbLf
    SlBackJs = s
End Function
