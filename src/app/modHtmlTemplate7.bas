Attribute VB_Name = "modHtmlTemplate7"
Option Explicit

' ==========================================================
' modHtmlTemplate7 - 見本の部品CSS(PartsCss)と SEC-17 growth の描画
' ------------------------------------------------
' 由来(18章§4.4の分割規約): v1.2 で体裁を出力見本(docs/design/出力見本_春華堂
'   統合提案_v0.1.html)へ寄せた結果、見本の部品CSS(.universe / .heatmap /
'   .radar / .ideas / .proposal / .questions ほか)を足すと
'   modHtmlTemplate1 が25,000字を超える見込みになったため、
'   「25,000字を超えたら次番のモジュールへ**関数単位で**切り出す」に従って
'   PartsCss をここへ置いた。CSSの呼び口は従来どおり
'   modHtmlTemplate1.HeadHtml の1本だけ(§4.4「共通CSSは HeadHtml に一元化」)。
'   あわせて11章§8.4 #6「SEC-17 の描画は既存テンプレモジュールへ足さず、
'   新設のテンプレモジュールへ置く」に従い SecGrowthJs も本モジュールが持つ。
'
' 正は18章§3.7(SEC-17 の描き方)・§5.1(CSS変数の閉じた一覧)・§6(印刷)。
' 描画は createElement / textContent / setAttribute だけ(18章§4.1)。
' R4準拠: Excelトークン・Application.Run・案件データ参照を持たない純文字列。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ==========================================================

' 【18章§4.4 の分割規約による移設(v1.3・T-56)】 modHtmlTemplate1 が SEC-18 の
'   連結行と TalkCss の追加で 25,000字を超えたため、**関数名を変えずに**
'   CommonCss をここへ移した(Private -> Public の変更のみ)。CSSの持ち主が
'   増えたわけではなく、呼び口は modHtmlTemplate1.HeadHtml の1本のままである。
' CommonCss - 骨格の共通CSS(18章§4.4「HeadHtml に一元化」)。上部ナビ・ヒーロー・
'   節見出し・カード・表・バッジ・印刷。セクション固有の部品は
'   本モジュールの PartsCss にある。生の色は #fff 以外書かない(§5.1)。
'   rgba() の半透明は下地の色を変えず濃さだけを作るので例外(§5.1の規約)。
Public Function CommonCss() As String
    Dim s As String
    s = s & "*{box-sizing:border-box}html{scroll-behavior:smooth}html,body{margin:0;padding:0}" & vbLf
    s = s & "body{background:var(--bg);color:var(--ink);font-family:var(--font-sans);font-size:var(--font-size);line-height:var(--line-height);-webkit-print-color-adjust:exact;print-color-adjust:exact}" & vbLf
    s = s & "a{color:var(--brand);text-decoration:none}a:hover{text-decoration:underline}" & vbLf
    s = s & ".wrap{max-width:var(--page-width);margin:0 auto;padding:0 var(--page-pad)}" & vbLf
    s = s & ".print-only{display:none}" & vbLf
    ' --- 上部ナビ(18章§3.6・§6(7)) ---
    s = s & ".topbar{position:sticky;top:0;z-index:50;background:var(--paper);border-bottom:1px solid var(--line)}" & vbLf
    s = s & ".topbar-inner{max-width:var(--page-width);margin:auto;padding:9px var(--page-pad);display:flex;gap:8px;align-items:center;overflow-x:auto;white-space:nowrap}" & vbLf
    s = s & ".topbar a{display:inline-flex;align-items:center;font-size:12px;color:var(--ink);font-weight:600;padding:6px 10px;border:1px solid var(--line);background:var(--soft-brand);border-radius:8px}" & vbLf
    s = s & ".topbar a:hover{border-color:var(--brand);text-decoration:none}" & vbLf
    ' --- ヒーロー(18章§3 SEC-01。表紙の差替) ---
    s = s & ".hero{background:linear-gradient(135deg,var(--brand2) 0%,var(--brand) 62%,var(--kaki) 100%);color:#fff;padding:56px 0 48px}" & vbLf
    s = s & ".hero-inner{max-width:var(--page-width);margin:auto;padding:0 var(--page-pad)}" & vbLf
    s = s & ".eyebrow{font-size:11px;letter-spacing:.16em;font-weight:800;opacity:.88;margin:0 0 10px}" & vbLf
    s = s & ".hero h1{font-size:34px;line-height:1.22;margin:0 0 10px;letter-spacing:.01em}" & vbLf
    s = s & ".hero-subtitle{font-size:16px;font-weight:600;margin:0 0 16px;color:#fff;opacity:.94}" & vbLf
    s = s & ".hero-meta{display:flex;flex-wrap:wrap;gap:8px}" & vbLf
    s = s & ".pill{font-size:12px;border:1px solid rgba(255,255,255,.28);background:rgba(255,255,255,.12);padding:5px 10px;border-radius:999px}" & vbLf
    s = s & ".btn{font-family:inherit;font-size:12px;margin-top:16px;padding:7px 16px;border:1px solid rgba(255,255,255,.5);background:rgba(255,255,255,.12);color:#fff;border-radius:8px;cursor:pointer}" & vbLf
    ' --- 警告バナー(18章§2 meta.warnings) ---
    s = s & ".banner{margin:16px auto 0;max-width:var(--page-width);padding:10px 14px;background:var(--warn);border:1px solid var(--warn-line);border-radius:10px;font-size:12.5px}" & vbLf
    s = s & ".banner-h{color:var(--kaki);font-weight:700;margin:0 0 3px}.banner p{margin:0}" & vbLf
    ' --- 節の見出し(18章§3.0 キッカー) ---
    s = s & "main{padding-top:26px;padding-bottom:70px}" & vbLf
    s = s & ".sec{scroll-margin-top:56px;margin:0 0 30px}" & vbLf
    s = s & ".section-head{margin:0 0 13px}" & vbLf
    s = s & ".section-kicker{font-size:11px;color:var(--brand);font-weight:800;letter-spacing:.16em;margin:0 0 2px}" & vbLf
    s = s & ".sec h2{font-size:23px;line-height:1.3;margin:0;font-weight:700}" & vbLf
    s = s & ".sec h3{font-size:15px;margin:16px 0 7px;font-weight:700}" & vbLf
    s = s & ".sec h4{font-size:13px;margin:12px 0 5px;color:var(--sub)}" & vbLf
    s = s & "p{margin:0 0 8px}" & vbLf
    s = s & ".sub,.muted{color:var(--sub);font-size:12.5px}.lead{font-size:14.5px}" & vbLf
    s = s & ".note{border:1px solid var(--warn-line);background:var(--warn);padding:9px 13px;border-radius:10px;font-size:12.5px;margin:0 0 12px}" & vbLf
    ' --- 本文先頭の目次(印刷用。18章§3.6) ---
    s = s & ".tocprint{margin:0 0 24px;padding:12px 16px;background:var(--mist);border:1px solid var(--line);border-radius:10px}" & vbLf
    s = s & ".toc-h{font-size:11.5px;letter-spacing:.2em;color:var(--sub);margin:0 0 7px}" & vbLf
    s = s & ".tocprint ol{margin:0;padding-left:1.5em;columns:2;font-size:12.5px}" & vbLf
    s = s & ".tocprint li{margin:0 0 3px;break-inside:avoid}.tocprint a{color:var(--ink)}" & vbLf
    ' --- カード・表・バッジ ---
    s = s & ".card{background:var(--paper);border:1px solid var(--line);border-radius:15px;box-shadow:var(--shadow);padding:18px;margin:0 0 12px;break-inside:avoid}" & vbLf
    s = s & ".card h3{margin:0 0 6px}.card-kaki{border-left:3px solid var(--kaki)}" & vbLf
    s = s & ".card-deep{border-left:3px solid var(--deep)}.card-matsu{border-left:3px solid var(--matsu)}" & vbLf
    s = s & ".hook{color:var(--brand);font-size:13px;margin:0 0 8px}" & vbLf
    s = s & ".tblwrap{overflow:auto;border:1px solid var(--line);border-radius:12px;background:var(--paper);margin:0 0 12px}" & vbLf
    s = s & "table{border-collapse:collapse;width:100%;min-width:960px;font-size:11.5px}" & vbLf
    s = s & "th,td{padding:9px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}" & vbLf
    s = s & "th{background:var(--mist);color:var(--sub);font-weight:700;font-size:10.5px;white-space:nowrap;position:sticky;top:0;z-index:1}" & vbLf
    s = s & "tr{break-inside:avoid}td.nw,th.nw{white-space:nowrap}" & vbLf
    s = s & ".bdg{display:inline-block;font-size:10.5px;padding:2px 8px;border-radius:999px;color:#fff;white-space:nowrap;font-weight:700}" & vbLf
    s = s & ".bdg-cover{background:var(--tr-cover)}.bdg-partial{background:var(--tr-partial)}.bdg-hard{background:var(--tr-hard)}" & vbLf
    s = s & ".bdg-ok{background:var(--iq-ok)}.bdg-iqpartial{background:var(--iq-partial)}.bdg-missing{background:var(--iq-missing)}" & vbLf
    s = s & ".bdg-sub{background:var(--sub)}.bdg-brand{background:var(--brand)}" & vbLf
    s = s & ".bdg-kaki{background:var(--kaki)}.bdg-matsu{background:var(--matsu)}.bdg-deep{background:var(--deep)}" & vbLf
    s = s & ".chips{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 10px}" & vbLf
    s = s & ".chip{font-size:11.5px;padding:3px 11px;border:1px solid var(--line);background:var(--paper);border-radius:999px;color:var(--sub)}" & vbLf
    s = s & ".chip-brand{border-color:var(--brand);color:var(--brand);background:var(--soft-brand)}" & vbLf
    s = s & ".dl{display:grid;grid-template-columns:9em 1fr;gap:5px 14px;font-size:13px;margin:0 0 14px}" & vbLf
    s = s & ".dl .dt{color:var(--sub);font-size:12px;padding-top:2px}.dl .dd{margin:0}" & vbLf
    s = s & ".rowline{display:grid;grid-template-columns:3.4em 1fr 9em;gap:8px;align-items:baseline;font-size:13.5px;margin:0 0 5px}" & vbLf
    s = s & ".rank-no{color:var(--brand);font-size:12px;font-weight:800}" & vbLf
    ' --- SEC-16 ラウンド更新 ---
    s = s & ".rublock{margin:0 0 14px;break-inside:avoid}" & vbLf
    s = s & ".rubh{font-size:13.5px;font-weight:700;margin:0 0 6px;padding:4px 10px;background:var(--mist);border-left:3px solid var(--brand);border-radius:0 8px 8px 0}" & vbLf
    s = s & ".rubh-rej{border-left-color:var(--kaki)}.rubh-new{border-left-color:var(--deep)}" & vbLf
    s = s & ".strike{text-decoration:line-through}" & vbLf
    s = s & ".disc{font-size:11px;color:var(--sub);background:var(--mist);border:1px solid var(--line);border-radius:12px;padding:16px}.disc p{margin:0 0 5px}" & vbLf
    ' 18章§6: 印刷(A4縦)との両立。@page の余白はCSS変数で解決されないため
    ' リテラルで書く(§5.1・§6(1))。
    s = s & "@page{size:A4 portrait;margin:14mm 12mm;}" & vbLf
    s = s & "@media print{.no-print,.topbar{display:none}.print-only{display:block}" & vbLf
    s = s & "body{background:#fff;font-size:12px;line-height:1.7}" & vbLf
    s = s & ".hero{padding:26px 0}.hero h1{font-size:26px}" & vbLf
    s = s & ".wrap,.hero-inner,.topbar-inner{max-width:none;padding-left:0;padding-right:0}" & vbLf
    s = s & ".tblwrap{overflow:visible}table{font-size:10px;min-width:0}th{position:static}" & vbLf
    s = s & ".card,.disc{box-shadow:none}" & vbLf
    s = s & ".sec-exec{break-after:page}a{color:var(--ink);text-decoration:none}" & vbLf
    s = s & ".card,.rublock,tr,.sec{break-inside:avoid}}" & vbLf
    s = s & "@media (max-width:900px){.dl{grid-template-columns:1fr}}" & vbLf
    s = s & "@media (max-width:640px){.hero h1{font-size:24px}.sec h2{font-size:19px}" & vbLf
    s = s & ".tocprint ol{columns:1}.rowline{grid-template-columns:2.8em 1fr}}" & vbLf
    CommonCss = s
End Function

' PartsCss - 見本の部品CSS。生の色は #fff 以外書かない(18章§5.1)。rgba() の
'   半透明は下地の色を変えず濃さだけを作るので例外(§5.1の規約)。
Public Function PartsCss() As String
    Dim s As String
    ' --- SEC-02 exec: 3テーマの3カラムカード(見本の .story) ---
    s = s & ".story{display:grid;grid-template-columns:repeat(3,1fr);gap:0;border:1px solid var(--line);border-radius:15px;overflow:hidden;background:var(--paper);box-shadow:var(--shadow);margin:0 0 14px}" & vbLf
    s = s & ".story>div{padding:18px}.story>div+div{border-left:1px solid var(--line)}" & vbLf
    s = s & ".story .num{font-size:11px;font-weight:800;color:var(--brand);letter-spacing:.12em}" & vbLf
    s = s & ".story h3{font-size:17px;margin:6px 0 8px}.story p{font-size:12.5px;color:var(--sub);margin:0 0 6px}" & vbLf
    ' --- SEC-03 profile: 事実と推定の2カラム(見本の .company-map) ---
    s = s & ".company-map{display:grid;grid-template-columns:1.05fr 1fr;gap:14px}" & vbLf
    s = s & ".tag{display:inline-flex;align-items:center;font-size:10px;font-weight:800;border-radius:999px;padding:3px 8px;margin:0 4px 0 0}" & vbLf
    s = s & ".fact{background:var(--soft-blue);color:var(--navy)}" & vbLf
    s = s & ".infer{background:var(--soft-amber);color:var(--kaki)}" & vbLf
    s = s & ".verify{background:var(--soft-red);color:var(--brand2)}" & vbLf
    ' --- SEC-05 riskuniv: 10枚のカード(見本の .universe) ---
    s = s & ".universe{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:0 0 12px}" & vbLf
    s = s & ".u{padding:12px;border:1px solid var(--line);background:var(--paper);border-radius:12px;min-height:112px}" & vbLf
    s = s & ".u .n{font-size:10px;color:var(--brand);font-weight:900}" & vbLf
    s = s & ".u h4{font-size:12.5px;margin:4px 0 5px;color:var(--ink)}" & vbLf
    s = s & ".u .cnt{font-size:19px;font-weight:800}" & vbLf
    s = s & ".u .bar{height:5px;background:var(--mist);border-radius:99px;overflow:hidden;margin-top:6px}" & vbLf
    s = s & ".u .bar span{display:block;height:100%;background:var(--brand)}" & vbLf
    ' --- SEC-06 riskmap: ヒートマップ + 経営優先度 Top 8(見本の .heat-wrap) ---
    s = s & ".heat-wrap{display:grid;grid-template-columns:460px 1fr;gap:20px;align-items:start;margin:0 0 12px}" & vbLf
    s = s & ".heatmap{display:grid;grid-template-columns:34px repeat(5,1fr);gap:3px}" & vbLf
    s = s & ".heatmap .y,.heatmap .xlab{display:flex;align-items:center;justify-content:center;font-size:10px;color:var(--sub);min-height:22px}" & vbLf
    s = s & ".cell{position:relative;min-height:66px;border-radius:7px;border:1px solid var(--line);padding:4px 4px 13px}" & vbLf
    s = s & ".band{position:absolute;right:5px;bottom:2px;font-size:9.5px;color:var(--sub)}" & vbLf
    ' 帯の面色(18章§3.3 の帯番号1-5 = §5.1 の --heat-1..5)。SEC-08 の評価列
    ' (§3.8)も同じ区切りの帯を使うので、クラス定義はここに1組だけ置く。
    s = s & ".heat1{background:var(--heat-1)}.heat2{background:var(--heat-2)}" & vbLf
    s = s & ".heat3{background:var(--heat-3)}.heat4{background:var(--heat-4)}" & vbLf
    s = s & ".heat5{background:var(--heat-5)}" & vbLf
    ' --- SEC-08 coverage(18章§3.8): 評価列の帯と、確度 assumed の淡色 ---
    s = s & ".prod{display:inline-block;padding:2px 8px;border-radius:8px;white-space:nowrap}" & vbLf
    s = s & ".cov-assumed{color:var(--sub);font-style:italic}" & vbLf
    s = s & ".dot{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;margin:1px;border-radius:50%;font-size:9.5px;font-weight:800;color:#fff;background:var(--brand);cursor:help}" & vbLf
    s = s & ".dot-cover{background:var(--tr-cover)}.dot-partial{background:var(--tr-partial)}.dot-hard{background:var(--tr-hard)}" & vbLf
    s = s & ".axname{font-size:11px;color:var(--sub);margin:6px 0 0}" & vbLf
    s = s & ".priority-list{display:grid;gap:8px}" & vbLf
    s = s & ".priority{display:grid;grid-template-columns:34px 1fr 74px;gap:10px;align-items:start;padding:10px 12px;border:1px solid var(--line);border-radius:10px;background:var(--paper)}" & vbLf
    s = s & ".priority .rid{font-weight:900;color:var(--brand);font-size:12px}" & vbLf
    s = s & ".priority .riskname{font-size:12px;font-weight:700}" & vbLf
    s = s & ".priority .why{font-size:11px;color:var(--sub);margin-top:3px}" & vbLf
    s = s & ".score{font-size:11px;font-weight:800;text-align:right}" & vbLf
    ' --- SEC-07 risks: カテゴリの絞り込み(見本の .filters) ---
    s = s & ".filters{display:flex;flex-wrap:wrap;gap:7px;margin:0 0 10px}" & vbLf
    s = s & ".filters button{font-family:inherit;border:1px solid var(--line);background:var(--paper);color:var(--ink);padding:6px 11px;border-radius:999px;font-size:11px;cursor:pointer}" & vbLf
    s = s & ".filters button.active{background:var(--navy);color:#fff;border-color:var(--navy)}" & vbLf
    ' --- SEC-09 newrisk: 2列のレーダーカード(見本の .radar) ---
    s = s & ".radar{display:grid;grid-template-columns:repeat(2,1fr);gap:11px}" & vbLf
    s = s & ".radar-item{border:1px solid var(--line);border-radius:12px;padding:15px;background:var(--paper);break-inside:avoid}" & vbLf
    s = s & ".radar-top{display:flex;justify-content:space-between;gap:10px;align-items:flex-start}" & vbLf
    s = s & ".radar-item h4{font-size:13px;margin:0}" & vbLf
    s = s & ".radar-item p{font-size:11.5px;color:var(--sub);margin:7px 0}" & vbLf
    s = s & ".horizon{font-size:10px;font-weight:800;color:var(--deep);background:var(--soft-purple);border-radius:999px;padding:3px 8px;white-space:nowrap}" & vbLf
    s = s & ".bar{height:6px;background:var(--mist);border-radius:99px;overflow:hidden}" & vbLf
    s = s & ".bar span{display:block;height:100%;background:var(--brand)}" & vbLf
    ' --- SEC-17 growth: アイデア(見本の .idea。18章§3.7) ---
    s = s & ".ideas{display:grid;gap:12px}" & vbLf
    s = s & ".idea{border:1px solid var(--line);background:var(--paper);border-radius:14px;padding:16px;display:grid;grid-template-columns:44px 1fr 250px;gap:15px;align-items:start;break-inside:avoid}" & vbLf
    s = s & ".idea .rank{width:36px;height:36px;border-radius:10px;background:var(--soft-brand);color:var(--brand);font-weight:900;display:flex;align-items:center;justify-content:center}" & vbLf
    s = s & ".idea h4{font-size:14px;margin:0 0 5px}" & vbLf
    s = s & ".idea p{font-size:11.5px;color:var(--sub);margin:0 0 5px}" & vbLf
    s = s & ".idea-side{font-size:10.5px;border-left:1px solid var(--line);padding-left:14px;color:var(--sub)}" & vbLf
    s = s & ".idea-side b{color:var(--ink);display:block;margin-top:8px}" & vbLf
    s = s & ".stars{letter-spacing:1px;font-size:13px}" & vbLf
    s = s & ".stars .on{color:var(--accent)}.stars .off{color:var(--line)}" & vbLf
    s = s & ".mini{font-size:10.5px;color:var(--sub);margin-top:2px}" & vbLf
    s = s & ".pill-dif{display:inline-block;margin-top:6px;font-size:10px;font-weight:800;color:#fff;padding:3px 9px;border-radius:999px}" & vbLf
    s = s & ".dif-low{background:var(--matsu)}.dif-mid{background:var(--kaki)}.dif-high{background:var(--tr-hard)}" & vbLf
    ' --- SEC-10 story: 濃色の提案ブロックと STEP の流れ(見本の .proposal) ---
    s = s & ".proposal{background:var(--navy);color:#fff;border-radius:16px;padding:22px;box-shadow:var(--shadow);margin:0 0 14px}" & vbLf
    s = s & ".proposal h3{font-size:20px;margin:0 0 6px;color:#fff}" & vbLf
    s = s & ".proposal p{font-size:12.5px;margin:0;color:#fff;opacity:.84}" & vbLf
    s = s & ".proposal-flow{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:14px}" & vbLf
    s = s & ".pf{padding:12px;border:1px solid rgba(255,255,255,.2);border-radius:10px;background:rgba(255,255,255,.06)}" & vbLf
    s = s & ".pf .n{font-size:10px;opacity:.72}.pf b{display:block;margin:3px 0 5px;font-size:12.5px}" & vbLf
    s = s & ".pf span{font-size:10.5px;opacity:.8}" & vbLf
    ' --- SEC-13 hearing: 2段組の設問カード(見本の .questions) ---
    s = s & ".questions{columns:2;column-gap:24px}" & vbLf
    s = s & ".q{break-inside:avoid;margin:0 0 9px;padding:10px 12px;border:1px solid var(--line);border-radius:10px;background:var(--paper);font-size:11.5px}" & vbLf
    s = s & ".q b{color:var(--brand)}.q .qcat{color:var(--sub);font-size:10.5px;margin-left:6px}" & vbLf
    s = s & ".q .qwhy{color:var(--sub);font-size:10.5px;margin-top:3px}" & vbLf
    ' --- 狭い画面と印刷(18章§6(9)(4)) ---
    s = s & "@media (max-width:900px){.story,.proposal-flow{grid-template-columns:repeat(2,1fr)}" & vbLf
    s = s & ".company-map,.heat-wrap{grid-template-columns:1fr}.universe{grid-template-columns:repeat(2,1fr)}" & vbLf
    s = s & ".idea{grid-template-columns:44px 1fr}.idea-side{grid-column:2;border-left:0;border-top:1px solid var(--line);padding:10px 0 0}}" & vbLf
    s = s & "@media (max-width:640px){.story,.universe,.radar,.proposal-flow{grid-template-columns:1fr}" & vbLf
    s = s & ".story>div+div{border-left:0;border-top:1px solid var(--line)}.questions{columns:1}}" & vbLf
    s = s & "@media print{.filters{display:none}.proposal,.idea,.radar-item,.priority{box-shadow:none}" & vbLf
    s = s & ".heat-wrap{grid-template-columns:400px 1fr}.questions{columns:2}}" & vbLf
    PartsCss = s
End Function

' SEC-17 growth(18章§3.7)。`s3.growth_ideas[]` を effect の降順 -> difficulty の
'   易しい順(low -> mid -> high) -> 配列順 に並べ、順位バッジ + 本文 + 右の
'   補足カラム(★5段階・難度ピル・保険との接点)で描く。0件なら何も描かず、
'   登録表の empty:'hide' がセクションごと(目次からも)落とす。
'   `stories[]` と同じカードで描かない(11章§9-9)。
Public Function SecGrowthJs() As String
    Dim s As String
    s = s & "function renderGrowth(D,el){var s3=D.s3||{};" & vbLf
    s = s & "var gi=AR(s3.growth_ideas);if(!gi.length){return;}" & vbLf
    ' 18章§3.7 の免責は逐語(見本と同一。文言を足さない)。
    s = s & "T(el,'p','note','実現可否は保険業法、約款設計、募集スキーム、料率、データ取得可否、対象顧客の同意等の検討が必要です。ここではアイデア発散を優先しています。');" & vbLf
    s = s & "var a=gi.slice(0);" & vbLf
    s = s & "a.sort(function(x,y){var e=(y.effect||0)-(x.effect||0);" & vbLf
    s = s & "if(e!==0){return e;}" & vbLf
    s = s & "return (DIFO[S(x.difficulty)]||9)-(DIFO[S(y.difficulty)]||9);});" & vbLf
    s = s & "var box=T(el,'div','ideas');" & vbLf
    s = s & "for(var i=0;i<a.length;i++){var g=a[i];" & vbLf
    s = s & "var it=T(box,'div','idea');" & vbLf
    s = s & "T(it,'div','rank',String(i+1));" & vbLf
    s = s & "var mid=T(it,'div',null);" & vbLf
    s = s & "T(mid,'h4',null,S(g.title));" & vbLf
    s = s & "T(mid,'p',null,S(g.what));" & vbLf
    s = s & "if(NB(g.why)){T(mid,'p',null,'効く理由: '+S(g.why));}" & vbLf
    s = s & "var side=T(it,'div','idea-side');" & vbLf
    s = s & "var n=g.effect||0;var st=T(side,'div','stars');" & vbLf
    s = s & "for(var k=1;k<=5;k++){T(st,'span',(k<=n)?'on':'off','★');}" & vbLf
    ' 色と記号だけで意味を運ばない(18章§6(3))ため、数でも併記する。
    s = s & "T(side,'div','mini','効きめ '+S(n)+'/5');" & vbLf
    s = s & "T(side,'div','pill-dif dif-'+S(g.difficulty),'難度 '+LB(LDIF,g.difficulty));" & vbLf
    s = s & "if(NB(g.insurance_fit)){T(side,'b',null,'保険との接点');" & vbLf
    s = s & "T(side,'div','mini',S(g.insurance_fit));}}}" & vbLf
    SecGrowthJs = s
End Function
