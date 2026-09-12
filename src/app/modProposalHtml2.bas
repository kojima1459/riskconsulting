Attribute VB_Name = "modProposalHtml2"
Option Explicit

' ==========================================================
' modProposalHtml2 - 顧客向け提案書(Wide 22枚)のCSS。20章§5・§6が正。
' ------------------------------------------------
' 3本立て(見本 docs/design/提案書_wide/shunkado_proposal_wide_v0.1.src.html の
'   #format-css / #viewer-css / #content-css と同じ分け方):
'     RootCss    : CSS変数の閉じた一覧だけを定義する :root{...}(20章§5.1)
'     FormatCss  : スライドの器(16:9・マスター4種の座標)と表示・印刷(20章§6)
'     ContentCss : 本文コンポーネント(見出し・カード・表・ヒートマップ等)
'   呼び口は modProposalHtml1.HeadHtml の1本だけ(18章§4.4 と同じ一元化)。
'
' 画像は1点も持たない(docs/29 §5.2 Q-4 の軽量化裁定)。見本のヘッダー帯・
'   タイトル枠・ロゴ・透かしは、すべて CSS のグラデーションと枠線で再現する。
'   自己完結HTML 1ファイル・外部参照ゼロ(20章§1)。
' 書体は**社内PCに必ずある3書体だけ**(メイリオ / 游ゴシック / MS Pゴシック)。
'   Webフォントを読まない(社外へ出られないPCで欠字にしない)。
'
' R4準拠(12章§2): Excelトークン・案件データ・configに触れない純文字列。
' CP932準拠(15章§0 原則7): 本文・注釈ともに CP932 内の文字だけで書く。
' ==========================================================

' RootCss - 20章§5.1「テンプレが定義してよいCSS変数の閉じた一覧(19個)」。
'   これ以外を定義しない・これ以外を参照しない。色の直書きは #fff だけ許す
'   (18章§5.1 と同じ規約。tools/render_proposal.py が機械で見る)。
Public Function RootCss() As String
    Dim s As String
    s = s & ":root{" & vbLf
    s = s & "--slide-w:1280px;" & vbLf
    s = s & "--slide-h:720px;" & vbLf
    s = s & "--font-sans:'メイリオ',Meiryo,'游ゴシック','Yu Gothic','MS Pゴシック',sans-serif;" & vbLf
    s = s & "--font-base:24px;" & vbLf
    s = s & "--line-height:1.6;" & vbLf
    s = s & "--ink:#1a1a1a;" & vbLf
    s = s & "--sub:#5a6470;" & vbLf
    s = s & "--paper:#ffffff;" & vbLf
    s = s & "--line:#d8dee4;" & vbLf
    s = s & "--brand:#00655a;" & vbLf
    s = s & "--brand-deep:#00463e;" & vbLf
    s = s & "--accent:#c0392b;" & vbLf
    s = s & "--soft:#eef5f3;" & vbLf
    s = s & "--heat-1:#e3f1e4;" & vbLf
    s = s & "--heat-2:#fff2c2;" & vbLf
    s = s & "--heat-3:#ffd6a8;" & vbLf
    s = s & "--heat-4:#f7b3a6;" & vbLf
    s = s & "--note-bg:#f4f6f7;" & vbLf
    s = s & "--pad-x:4.4%;" & vbLf
    s = s & "}" & vbLf
    RootCss = s
End Function

' FormatCss - スライドの器と表示・印刷(20章§6)。見本の #format-css と
'   #viewer-css に対応する。1枚=1ページの印刷は @page{size:A4 landscape} と
'   .frame の break-after で担保する。
Public Function FormatCss() As String
    Dim s As String
    s = s & "html,body{margin:0;padding:0;background:var(--soft);" & vbLf
    s = s & "font-family:var(--font-sans);color:var(--ink);}" & vbLf
    s = s & "#deck{padding:24px 0;}" & vbLf
    s = s & ".frame{position:relative;width:min(var(--slide-w),calc(100vw - 48px));" & vbLf
    s = s & "aspect-ratio:16 / 9;margin:0 auto 24px;overflow:hidden;" & vbLf
    s = s & "background:var(--paper);box-shadow:0 1px 4px rgba(0,0,0,.25);}" & vbLf
    s = s & ".frame > section.slide{position:absolute;left:0;top:0;" & vbLf
    s = s & "transform-origin:0 0;transform:scale(var(--s,1));}" & vbLf
    s = s & "section.slide{width:var(--slide-w);height:var(--slide-h);" & vbLf
    s = s & "box-sizing:border-box;overflow:hidden;position:relative;" & vbLf
    s = s & "background:var(--paper);color:var(--ink);" & vbLf
    s = s & "font-size:var(--font-base);line-height:var(--line-height);}" & vbLf

    ' ヘッダー帯(見本の header_bar.png を CSS で再現する。画像を持たない)
    s = s & ".header-bar{position:absolute;left:0;top:0;width:100%;height:8%;" & vbLf
    s = s & "background:linear-gradient(90deg,var(--brand-deep) 0%,var(--brand) 62%," & vbLf
    s = s & "var(--brand-deep) 100%);}" & vbLf
    s = s & ".content-title{position:absolute;left:var(--pad-x);top:0;width:68%;height:8%;" & vbLf
    s = s & "display:flex;align-items:center;color:#fff;font-size:29px;font-weight:700;}" & vbLf
    s = s & ".content-body{position:absolute;left:var(--pad-x);top:11%;width:91.2%;" & vbLf
    s = s & "height:82%;}" & vbLf
    s = s & ".footer-rule{position:absolute;left:0;top:96.5%;width:100%;height:0;" & vbLf
    s = s & "border-top:1px solid var(--brand);}" & vbLf
    s = s & ".page-number{position:absolute;left:38%;top:97%;width:24%;height:2.5%;" & vbLf
    s = s & "display:flex;align-items:center;justify-content:center;" & vbLf
    s = s & "font-size:16px;color:var(--sub);}" & vbLf

    ' 表紙(cover)。見本の title_box.png は CSS の枠で再現する。
    s = s & ".cover-to{position:absolute;left:4.3%;top:13%;font-size:32px;font-weight:700;}" & vbLf
    s = s & ".cover-box{position:absolute;left:3.8%;top:33%;width:92.5%;height:33%;" & vbLf
    s = s & "background:var(--brand);border:6px solid var(--brand-deep);" & vbLf
    s = s & "display:flex;flex-direction:column;align-items:center;" & vbLf
    s = s & "justify-content:center;gap:14px;box-sizing:border-box;padding:0 40px;}" & vbLf
    s = s & ".cover-title{color:#fff;font-size:45px;font-weight:700;line-height:1.3;" & vbLf
    s = s & "text-align:center;}" & vbLf
    s = s & ".cover-sub{color:#fff;font-size:24px;text-align:center;}" & vbLf
    s = s & ".cover-date{position:absolute;left:0;top:80%;width:100%;text-align:center;" & vbLf
    s = s & "font-size:26px;font-weight:700;}" & vbLf
    s = s & ".cover-by{position:absolute;left:0;top:88%;width:100%;text-align:center;" & vbLf
    s = s & "font-size:20px;color:var(--sub);}" & vbLf

    ' セクション扉(section)
    s = s & ".section-label{position:absolute;left:8.1%;top:14%;color:var(--brand);" & vbLf
    s = s & "font-size:29px;font-weight:700;letter-spacing:.02em;}" & vbLf
    s = s & ".section-title{position:absolute;left:8.1%;top:21%;width:83.7%;" & vbLf
    s = s & "font-size:35px;font-weight:700;}" & vbLf
    s = s & ".section-rule{position:absolute;left:8.1%;top:32%;width:83.7%;height:0;" & vbLf
    s = s & "border-top:4px solid var(--brand);}" & vbLf

    ' 裏表紙(back)
    s = s & ".back-mark{position:absolute;left:0;top:44%;width:100%;text-align:center;" & vbLf
    s = s & "font-size:34px;font-weight:700;color:var(--brand);letter-spacing:.08em;}" & vbLf
    s = s & ".back-note{position:absolute;left:0;top:56%;width:100%;text-align:center;" & vbLf
    s = s & "font-size:19px;color:var(--sub);}" & vbLf

    ' 発表者ノート(20章§7)。画面では ?notes=1 のときだけ出し、印刷では消す。
    s = s & "aside.note{display:none;max-width:var(--slide-w);margin:0 auto 24px;" & vbLf
    s = s & "padding:12px 16px;background:var(--note-bg);border-left:4px solid var(--brand);" & vbLf
    s = s & "font-size:15px;line-height:1.7;color:var(--ink);}" & vbLf
    s = s & "body.notes aside.note{display:block;}" & vbLf
    s = s & "aside.note b{color:var(--brand-deep);}" & vbLf

    ' 描画エラーの可視化(20章§11。裁定書39 R2-12)。描画関数が例外を投げた枚は
    ' 本文が20章§4.1 の1行に置き換わるが、それだけでは「わざと保留した項目」と
    ' 見分けが付かない。**?debug=1 を付けたときだけ**赤枠を出す(お客さまが
    ' 普通に開いた画面は1ピクセルも変わらない)。20章§5.1 により :root の外に
    ' 生の色指定を書けないため rgba() で書く。
    s = s & "body.debug section.slide[data-render-error=""1""]{" & vbLf
    s = s & "outline:5px solid rgba(198,0,0,0.9);outline-offset:-5px;}" & vbLf

    ' 印刷(20章§6): 1枚=1ページ。用紙はA4横。
    s = s & "@page{size:A4 landscape;margin:0;}" & vbLf
    s = s & "@media print{" & vbLf
    s = s & "html,body{background:#fff;}" & vbLf
    s = s & "#deck{padding:0;}" & vbLf
    s = s & ".frame{width:297mm;height:210mm;margin:0;aspect-ratio:auto;" & vbLf
    s = s & "box-shadow:none;page-break-after:always;break-after:page;}" & vbLf
    s = s & ".frame:last-child{page-break-after:auto;break-after:auto;}" & vbLf
    s = s & ".frame > section.slide{transform:scale(0.877)!important;}" & vbLf
    s = s & "aside.note{display:none!important;}" & vbLf
    s = s & "section.slide,section.slide *{-webkit-print-color-adjust:exact!important;" & vbLf
    s = s & "print-color-adjust:exact!important;}" & vbLf
    s = s & "}" & vbLf
    FormatCss = s
End Function

' ContentCss - 本文コンポーネント(20章§5.2)。スライド固有のクラスは
'   `sl-<slug>-*` の接頭辞でここへ置く(CSSを散らさない)。
Public Function ContentCss() As String
    Dim s As String
    s = s & ".hm{font-size:26px;font-weight:700;line-height:1.45;margin:0 0 14px;}" & vbLf
    s = s & ".hm em{font-style:normal;color:var(--accent);}" & vbLf
    s = s & ".sub{font-size:19px;color:var(--sub);margin:0 0 12px;}" & vbLf
    s = s & ".cols{display:flex;gap:20px;align-items:flex-start;}" & vbLf
    s = s & ".col{flex:1 1 0;min-width:0;}" & vbLf
    s = s & ".card{border:1px solid var(--line);border-top:4px solid var(--brand);" & vbLf
    s = s & "padding:12px 14px;background:var(--paper);box-sizing:border-box;}" & vbLf
    s = s & ".card h3{margin:0 0 6px;font-size:21px;color:var(--brand-deep);}" & vbLf
    s = s & ".card p{margin:0;font-size:17px;line-height:1.6;}" & vbLf
    s = s & ".stats{display:flex;gap:12px;margin:0 0 14px;}" & vbLf
    s = s & ".stat{flex:1 1 0;text-align:center;background:var(--soft);padding:8px 4px;}" & vbLf
    s = s & ".stat b{display:block;font-size:30px;color:var(--brand-deep);}" & vbLf
    s = s & ".stat span{font-size:14px;color:var(--sub);}" & vbLf
    s = s & "table.t{width:100%;border-collapse:collapse;font-size:16px;}" & vbLf
    s = s & "table.t th{background:var(--soft);color:var(--brand-deep);text-align:left;" & vbLf
    s = s & "padding:6px 8px;border:1px solid var(--line);font-weight:700;}" & vbLf
    s = s & "table.t td{padding:6px 8px;border:1px solid var(--line);" & vbLf
    s = s & "vertical-align:top;line-height:1.5;}" & vbLf
    s = s & "table.t tr.pri td{font-weight:700;}" & vbLf
    s = s & ".pill{display:inline-block;padding:1px 8px;border-radius:10px;" & vbLf
    s = s & "font-size:14px;background:var(--soft);color:var(--brand-deep);}" & vbLf
    s = s & ".pri-mark{color:var(--accent);font-weight:700;}" & vbLf
    s = s & ".bars{display:flex;flex-direction:column;gap:5px;}" & vbLf
    s = s & ".bar{display:flex;align-items:center;gap:8px;font-size:16px;}" & vbLf
    s = s & ".bar i{display:block;height:14px;background:var(--brand);}" & vbLf
    s = s & ".bar span.n{width:180px;flex:none;}" & vbLf
    s = s & ".map{display:grid;grid-template-columns:52px repeat(5,1fr);" & vbLf
    s = s & "grid-auto-rows:64px;gap:3px;}" & vbLf
    s = s & ".map .ax{display:flex;align-items:center;justify-content:center;" & vbLf
    s = s & "font-size:14px;color:var(--sub);}" & vbLf
    s = s & ".map .cell{border:1px solid var(--line);padding:2px;font-size:13px;" & vbLf
    s = s & "overflow:hidden;}" & vbLf
    s = s & ".map .b1{background:var(--heat-1);}" & vbLf
    s = s & ".map .b2{background:var(--heat-2);}" & vbLf
    s = s & ".map .b3{background:var(--heat-3);}" & vbLf
    s = s & ".map .b4{background:var(--heat-4);}" & vbLf
    s = s & ".steps{display:flex;gap:12px;}" & vbLf
    s = s & ".step{flex:1 1 0;border:1px solid var(--line);padding:10px 12px;" & vbLf
    s = s & "box-sizing:border-box;}" & vbLf
    s = s & ".step b{display:block;color:var(--brand-deep);font-size:16px;}" & vbLf
    s = s & ".step h3{margin:4px 0 6px;font-size:19px;}" & vbLf
    s = s & ".step p{margin:0 0 6px;font-size:15px;line-height:1.55;}" & vbLf
    s = s & ".step .ref{font-size:13px;color:var(--sub);}" & vbLf
    s = s & ".lead{font-size:18px;line-height:1.7;margin:0 0 10px;}" & vbLf
    s = s & ".muted{font-size:15px;color:var(--sub);line-height:1.6;}" & vbLf
    s = s & ".fixed-note{position:absolute;left:var(--pad-x);top:90%;width:91.2%;" & vbLf
    s = s & "font-size:14px;color:var(--sub);}" & vbLf
    s = s & ".grid2{display:grid;grid-template-columns:1fr 1fr;gap:10px 20px;}" & vbLf
    s = s & ".kv{display:flex;gap:8px;font-size:16px;border-bottom:1px dotted var(--line);" & vbLf
    s = s & "padding:4px 0;}" & vbLf
    s = s & ".kv b{flex:none;width:150px;color:var(--brand-deep);font-weight:700;}" & vbLf
    s = s & ".stars{color:var(--accent);letter-spacing:2px;font-size:16px;}" & vbLf
    ContentCss = s
End Function
