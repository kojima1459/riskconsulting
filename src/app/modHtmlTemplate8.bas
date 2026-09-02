Attribute VB_Name = "modHtmlTemplate8"
Option Explicit

' ==========================================================
' modHtmlTemplate8 - SEC-18 talk(経営層への話し方)の描画スクリプトと TalkCss
' ------------------------------------------------
' 正は18章§3.9(描き方)・§3.0(見本08節=SEC-10 の直後)・§4.4(分割規約の既定
' 割り当て)。CSSも登録行も持たない…のではなく、**本節専用の部品CSS TalkCss
' だけ**を持つ(§4.4。HeadHtml が PartsCss の直後に連結する。共通CSSの持ち主が
' 増えたわけではなく、25,000字規約による切り出しである)。
'
' なぜ7番の余白ではなく新モジュールなのか: 18章§4.4 の 25,000字規約
' (vba_lint の TEMPLATE_MAX_CHARS が機械強制)による。関数名は変えずに移動
' するだけの分割規約に従い、新設側へ**新しい関数だけ**を置く。
'
' 画面に「トークスクリプト」「talk_script」の語を出さない(11章§7の利用者向け
' 文の規約。18章§3.9)。taboo[] の根拠(field_insights の原文)はここに出さない
' (顧客同席の場で事故になるため。同)。
'
' R4準拠: Excelトークン・Application.Run・案件データ参照を持たない純文字列。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ==========================================================

' SEC-18 talk(18章§3.9)。s3.talk_script を描く。
'   ・opening は節の先頭に大きめの1文(見本の統合ストーリー大見出しの位置)
'   ・flow[] は**配列順のまま**番号付きカードで並べる(並べ替え・要約をしない)
'   ・closing はカードの下に1文
'   ・taboo[] は最後に注意の枠。**0件なら枠ごと出さない**
'   talk_script が無ければ何も描かない(登録表の empty:'hide' が拾って
'   セクションごと落とし、目次からも消える)。
Public Function SecTalkJs() As String
    Dim s As String
    s = s & "function renderTalk(D,el){var s3=D.s3||{};" & vbLf
    s = s & "var tk=s3.talk_script;if(!tk){return;}" & vbLf
    s = s & "var fl=AR(tk.flow);" & vbLf
    s = s & "if(!NB(tk.opening)&&!fl.length&&!NB(tk.closing)){return;}" & vbLf
    ' 18章§3.9 の固定文は逐語(前後に語を足さない)。
    s = s & "T(el,'p','note','そのまま読み上げる原稿ではありません。"
    s = s & "話す順番の下書きとしてお使いください。');" & vbLf
    s = s & "if(NB(tk.opening)){T(el,'p','talk-open',S(tk.opening));}" & vbLf
    s = s & "if(fl.length){var box=T(el,'div','talk-flow');" & vbLf
    s = s & "for(var i=0;i<fl.length;i++){var c=T(box,'div','talk-step');" & vbLf
    s = s & "T(c,'div','talk-no',String(i+1));" & vbLf
    s = s & "T(c,'p',null,S(fl[i]));}}" & vbLf
    s = s & "if(NB(tk.closing)){T(el,'p','talk-close',S(tk.closing));}" & vbLf
    s = s & "var tb=AR(tk.taboo);if(!tb.length){return;}" & vbLf
    s = s & "var w=T(el,'div','note');T(w,'b',null,'触れない事');" & vbLf
    s = s & "var ul=T(w,'ul');" & vbLf
    s = s & "for(var k=0;k<tb.length;k++){T(ul,'li',null,S(tb[k]));}}" & vbLf
    SecTalkJs = s
End Function

' TalkCss - SEC-18 の部品CSS(吹き出しと番号付きカード。18章§4.4・§5.1)。
'   色・書体・幅はリテラルで書かず必ず var(--xxx) を通す(§5.1の規約)。
'   印刷では横並びを縦へ落とす(§6)。
Public Function TalkCss() As String
    Dim s As String
    s = s & ".talk-open{font-size:17px;font-weight:800;line-height:1.6;color:var(--ink);"
    s = s & "background:var(--soft-brand);border-left:4px solid var(--brand);"
    s = s & "border-radius:0 12px 12px 0;padding:14px 16px;margin:0 0 14px}" & vbLf
    s = s & ".talk-flow{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;"
    s = s & "margin:0 0 12px}" & vbLf
    s = s & ".talk-step{position:relative;border:1px solid var(--line);border-radius:12px;"
    s = s & "background:var(--paper);box-shadow:var(--shadow);padding:14px 12px 12px}" & vbLf
    s = s & ".talk-step p{font-size:12.5px;color:var(--ink);margin:0}" & vbLf
    s = s & ".talk-no{width:24px;height:24px;border-radius:999px;background:var(--brand);"
    s = s & "color:#fff;font-size:12px;font-weight:800;display:flex;align-items:center;"
    s = s & "justify-content:center;margin:0 0 8px}" & vbLf
    s = s & ".talk-close{font-size:13.5px;font-weight:700;color:var(--navy);"
    s = s & "border-top:1px solid var(--line);padding:10px 0 0;margin:0 0 12px}" & vbLf
    s = s & "@media print{.talk-flow{grid-template-columns:1fr}}" & vbLf
    TalkCss = s
End Function
