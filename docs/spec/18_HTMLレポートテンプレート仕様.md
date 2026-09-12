# 18. HTMLレポートテンプレート仕様 v1.3

v1.3（W7・裁定書25「W7センターピン整合」）: 4点を改訂した。**（T-56 実装時の追認）** §4.4 の分割表で `CommonCss` の持ち主を `modHtmlTemplate7` へ移した（`modHtmlTemplate8` の連結行と `TalkCss` の追加で `modHtmlTemplate1` が25,000字規約を超えたため。関数名は変えていない）。**(1) SEC-08 の「両方0件なら非表示」を撤回**（裁定書25 S1）。`s2.risks[].insurability` があれば必ず描く。新規案件で現契約が無くても「保険カバレッジ表・未充足リスク一覧」が出る状態にした（UC案 Output3-4・見本05節）。**(2) §3.8 を新設**（同 S5）: SEC-08 の主表を**リスク単位の8列表**（見本05節と同じ粒度）とし、現契約表・gaps表をその下に置く。読むJSONパスに `insurability.gap_note`（15章 v2.6 で `line_note` から分離）を加えた。**(3) SEC-18 talk を新設**（同 S2）: `s3.talk_script` を描く。**改番はしていない**（SEC-01..SEC-17 は不変。追加は SEC-18 以降という §3 の規約どおり）。§3.0 の対応表では見本08節に SEC-10 と並置し、キッカーは節の先頭である SEC-10 が持つ。描き方は §3.9。**(4)** SEC-07 リスク一覧の読むJSONパスへ `insurability.gap_note` を追加した。§3.5 の免責フッタ（「保険料の試算は本資料の対象外です（要見積）。」）は**据え置き**である。

v1.2（裁定書21・11章v3.2 §3.8「HTMLレポートの体裁」）: 体裁の正を `docs/design/出力見本_春華堂統合提案_v0.1.html`（以下「見本」）へ移し、次の6点を改訂した。**(1)** §3のセクション表へ **SEC-17 growth（攻めの保険活用）** を新設し（読むJSONパスは `s3.growth_ideas[]`。15章§4 v2.5）、描き方を§3.7へ逐語で定めた。**(2)** §3の並びを見本の10節の流れへ**並べ替え**た（IDの改番はしていない。§4.3「並べ替えは登録配列の行順の入れ替えだけ」）。どのセクションが見本のどの節に集まるかは**§3.0の対応表**が持つ。**(3)** §3.6 の目次を**上部ナビ（`position:sticky` のアンカー帯）**へ差し替え、印刷時は帯を消して本文先頭に目次を出す形にした。**(4)** §5.1 のCSS変数の閉じた一覧を **28個→39個**へ拡張した（`--brand` / `--brand2` / `--accent` / `--navy` / `--bg` / `--shadow` / `--soft-*` 6色を追加し、旧 `--ai` は `--brand` へ改称）。**(5)** テーマを3本（`standard`＝見本の臙脂 / `mono`＝白黒印刷 / `ds`＝DS版の白地＋青帯）にした。**(6)** §4.4 の分割表へ `modHtmlTemplate7`（SEC-17 と部品CSS）を足した。§4.1「描画はJSが行う」「innerHTML系を使わない」・§5.3のエスケープ・§6の印刷規約は**変えていない**。

v1.1.1（裁定書10: W4.2 収束ウェーブ・仕様側v2.5.1と同時改訂）: 充足度（`input_quality.overall`）の日本語ラベルを19章§3の改訂（「充足度 高」→「高」。裁定書9 A-6・裁定書10 m4）へ追随させ、SEC-04 のラベル辞書（実装 `modHtmlTemplate6` の LIQO）も **高／中／低** を用いることを§3の規約へ明記した（同一enumのラベルが画面側と2系統に分岐しない。実装側の LIQO は実装班が同時更新する）。

v1.1（W3.1裁定）: 4点を改訂した。**(1) §5.3(1)のエスケープ集合**を「`</` を `<\/` へ」から「**すべての `<` を `\u003C` へ**」へ改めた（旧集合では `<!--<script>`〔`-->` を伴わない形〕を含むLLM出力でHTMLトークナイザが script data double escaped 状態へ入り、DATAブロックの正規の `</script>` が終端として働かず**生成HTMLが実ブラウザ上で白紙化**した。16章 E-47 も同時改訂）。**(2) 見出しの正を§3の表「見出し（既定）」列に確定**し、柱書の『本章が正であるもの』へ「見出し」を加え、§4.2のコード例（SEC-06/09/16）を§3のフル表記へ揃えた。**(3) §4.4の25,000字規約**の検査先を `tools/vba_lint.py` の `modHtmlTemplate*` 専用ERROR閾値と明記し、既定の割り当て表を実態（1..6）へ更新した。**(4) §5.2にテンプレ5関数の署名**（`coverFields` はvbTab区切り3値）を、§2に `meta.warnings` を明文化した。あわせて `tools/render_report.py` に「DATAリテラルに生の `<` が無い」「18章の固定文・見出しの逐語照合」「`meta.round_no=1` でSEC-16が本文からも目次からも消える」の3検査を追加した。

v1.0（ニューリスク=エマージング確定）: SEC-09「ニューリスク」の読むJSONパスを `s2.emerging_risks`（新種・新興リスク）へ差し替え、「第2ラウンド以降に表示」の文言を削除して空配列時の1行を「現時点で特筆すべきニューリスクは検出されていません」へ変更した。あわせて第2ラウンドの仮説ライフサイクル（`s2.risks[].status`）を見せる SEC-16「訪問で分かったこと（ラウンド更新）」を新設し、両者が別物であることを§3に明記した。
v1.0変更概要: 仕様書v2.4の実装前監査裁定により新設。10章FR-37の「デザインは差し替え可能なテーマとして分離」を、セクション登録表・CSS変数の閉じた一覧・1モジュールに閉じたテーマ差替として具体化し、12章§2の純文字列モジュール `modHtmlTemplate1..n` / `modHtmlTheme`（17章 T-35）の内容面の正を定義した。

**目的**: 本製品の主力出力であるHTMLリスクレポート（10章FR-37・14章§6 `modExportHtml.GenerateHtmlReport`）について、**出力物の見た目・構成・図表に対するフィードバックを1箇所の修正で吸収できる構造**を固定する。発注者・部会からの「この図を足したい」「この順番を入れ替えたい」「配色を変えたい」という要望が、そのつどHTML生成コードの改造にならないよう、変更点を (a) セクション登録表 (b) テーマのCSS変数 (c) 個々のテンプレ関数 の3種類だけに閉じ込める。

**本章が正であるもの**: セクションID・セクションの並び・**セクションの見出し（§3の表「見出し（既定）」列が正。§4.2のコード例はこれを写したものであり、食い違ったら§3が正）**・各セクションが読むJSONパス・CSS変数名の閉じた一覧とその既定値・テーマ差替の単位・テンプレモジュールの分割規約・図表追加手順・エスケープと文字コード・印刷とブラウザ表示の両立規約。
（本章が正であるものの続き）: **`modHtmlTemplate1..n` / `modHtmlTheme` の関数署名（§5.2。14章§6が本章へ委ねている）**。
**本章が正でないもの**: JSONのプロパティ名（正=15章のスキーマ）・enumの日本語ラベル（正=19章§3・15章§0）・configキーの既定値（正=13章§2.3）・エラーコードと失敗時挙動（正=16章 E-21／E-47／E-48）・タスクとDoD（正=17章 T-33／T-35）。

---

## 1. 責務分担とモジュール構成

| モジュール | 層 | 責務 | 禁止事項 |
|---|---|---|---|
| `modExportHtml` | app | 案件からJSONを解決し、匿名化復元・PII走査・DATA組立・テンプレ組立・ファイル書出・`report_path` 記録を行う | HTML本文・CSS・JSの文字列をここに書かない（12章§4。30,000字契約に抵触するため） |
| `modHtmlTemplate1..n` | app（純文字列） | HTML骨格・CSS・セクション描画スクリプト・セクション登録表を文字列として返す | Excelトークン（Worksheets/Range/MsgBox等）・`Application.Run`・案件データの参照（R4） |
| `modHtmlTheme` | app（純文字列） | テーマ名に対応するCSS変数ブロック `:root{...}` **だけ**を返す | セレクタ・レイアウト宣言・CSS変数以外の宣言を1つも書かない |

- 生成に**LLMを使わない**（10章FR-37・16章NFR-S4）。テンプレートもスクリプトも人が書いた固定文字列であり、LLM出力は**データとしてのみ**ページに入る。
- 出力は**自己完結HTML 1ファイル**。外部CSS・外部JS・外部フォント・外部画像・CDNを一切参照しない（社内PCはオフライン閲覧・メール添付での配布を想定）。図はCSSとインラインSVGだけで描く。

### 1.1 生成手順（modExportHtml。この順序で行う）

| 順 | 手順 | 失敗したら |
|---|---|---|
| ① | `modCaseStore.ResolveStepJson` で S1・S2・S3 のJSONを解決する（優先順の正は13章§2.2。S1は `s1_edited > s1_json`、S2/S3は `sN_edited > sNr_json > sN_json`） | **S1が空なら生成しない**。エラーコードは立てず（データ未整備は障害ではない）「先にStep1を実行してください」と案内して中止する。S2・S3が空でも生成は続行し、該当セクションは§3の「空のときの挙動」で処理する（S1だけでも企業理解・充足度・ヒアリング事項は出せるため） |
| ② | 匿名化の復元（16章 E-31）。`{{COMPANY}}` `{{POLICY_NO}}` を実名へ戻す。**必ずエスケープより前**に行う（復元後の実名がエスケープ対象になるようにするため） | 復元表が無い場合はプレースホルダのまま出力し、`hm_warning` に「匿名化の復元ができませんでした」を表示 |
| ③ | `modPii` 走査（14章§6・16章 E-05）。検知しても**生成はブロックせず**、`hm_warning` へ警告を出し run_log の detail に検知種別と箇所のみを記録する（本文は記録しない=NFR-S3）。出力先が利用者ローカルであり、レポートは人が確認したうえで配布するため | 走査自体の失敗は警告のみで続行 |
| ④ | DATA（§2）を1本のJSON文字列として組み立てる | 組立失敗は E0502 |
| ⑤ | `modHtmlTemplate1.BuildDocument(themeName, dataJson, coverFields)` でHTML全文を組み立てる | 組立失敗は E0502 |
| ⑥ | **UTF-8（BOMあり）で書き出す**（§5.3。実体は `modUtil.WriteUtf8File`） | 書込失敗は E0502。出力先不存在は先に16章 E-21 のフォールバックを試す |
| ⑦ | 確定パスを案件一覧の `report_path` に記録する（13章§2.1）。ファイル名は `modUtilText.SanitizeFileName` を通す（14章§6） | 記録失敗は警告のみ。生成済みファイルは残す |

- 出力先は config `html_out_dir`、テーマは config `html_theme`（13章§2.3）。この2キー以外にHTMLレポート用のconfigキーを増やさない。**セクションの取捨・並びはconfigではなく§4の登録表で行う**（設定項目を増やすとテーマ差替の単位が壊れるため）。
- 案件の `status` はレポート生成の成否で変化させない（16章 E-48）。

---

## 2. DATA（ページに埋め込む唯一のデータ構造）

ページ内に埋め込むのは次の形の**1本のJSON**だけである。`s1` / `s2` / `s3` は**15章スキーマの出力をキー名を変えずそのまま入れる**。VBA側でキー名を作り替えない（作り替えると15章とテンプレートの二重管理になるため）。

```json
{
  "meta": {
    "case_id": "C-20260901-001", "company": "浜松スイーツファクトリー株式会社",
    "industry_code": "09", "industry_name": "食料品製造業",
    "case_type": "renewal", "dossier_tier": "t2_full", "quality_mode": "deep",
    "round_no": 1, "s4_variant": "proposal",
    "generated_at": "2026/09/01 14:07:22", "app_version": "2.4.0", "theme": "standard",
    "reviewed_by": "", "reviewed_at": "", "ground_unmatched": [],
    "warnings": []
  },
  "s1": { "company_name": "...", "business_summary": "...", "...": "Schema-S1 の全キー" },
  "s2": { "risks": [], "gaps": [], "emerging_risks": [], "open_questions": [] },
  "s3": { "stories": [], "unmatched_risks": [], "do_not_propose": [] }
}
```

- `meta` の由来: `case_id` / `case_type` / `dossier_tier` / `company` / `industry_code` / `industry_name` / `round_no` / `s4_variant` は13章§2.1『案件一覧』の同名列。`quality_mode` はHOMEの `hm_quality_mode`（13章§2.10。未上書きなら config `quality_mode`）。`app_version` は config `app_version`、`theme` は config `html_theme` を `modHtmlTheme.ThemeCss` で解決したあとの実テーマ名（§5.2のフォールバック後の値）、`generated_at` は生成時刻 `yyyy/mm/dd hh:mm:ss`。
- `meta.reviewed_by` / `meta.reviewed_at` は**担当者が内容を確認・編集したか**（v1.4・裁定書37 B-06）。未確認は**両方とも空文字**（キー自体は必ず置く）。値を入れるのは `modExportHtml.GenerateHtmlReportEx` の `reviewedBy` 引数が非空のときだけで、`reviewed_at` はそのとき生成側が打つ（`modUtil.NowStamp` の `yyyy/mm/dd hh:mm:ss`）。読むのは §3.5 の免責1行目と §3 SEC-01 の表紙チップ、および §4.1 の `<noscript>`。
- `meta.ground_unmatched` は**引用の原文照合で見つけられなかったリスクの番号**の文字列配列（v1.4・裁定書37 B-03。0本以上。キー自体は必ず置き、無ければ空配列）。`s2.risks[]` は `risk_no` をそのまま、`s2.emerging_risks[]` は `risk_no` を持たないため配列の出現順に `"E1"` `"E2"` と採番した値が入る。値源は `modGround.GroundNotes`（貼付原文＝`case_data` の `input_*` 全欄＋`s1_json` と突き合わせる純関数）で、`evidence.source` が `inference` / `knowledge` のものと、貼付原文が空のときは**検査しない**（空配列になる。config `ground_check` / `ground_head_chars`＝13章§2.3）。読むのは §3 SEC-14 の「原文照合」列だけで、**SEC-04 の充足度バッジは動かさない**（別の事実なので混ぜない）。
- `meta.warnings` は**生成をブロックしない警告**の文字列配列（0本以上。キー自体は必ず置き、無ければ空配列）。載せてよいのは§1.1の②③が定める「匿名化の復元ができませんでした」（16章 E-31）と `modPii` の検知（16章 E-05(6)。**検知種別と箇所だけで本文は載せない**＝NFR-S3）に限る。ページ側は本文の前に1枚のバナーとして出す（`hm_warning` と同じ内容を、レポート単体で配布したときにも読めるようにするためのもの。W3.1で追認）。**エラーコード・スタックトレース・入力原文をここへ入れない**（レポートは成果物であり障害報告書ではない＝16章NFR-S3）。
- S3が未実行の案件では `"s3": null` とする（キー自体は必ず置く)。同様にS2未実行は `"s2": null`。`null` のときの各セクションの挙動は§3の「空のときの挙動」列が正。
- **DATAに入れないもの**: 入力貼付テキストの原文（`input_hp` 等）・run_log・err_log・ナレッジ本文・APIキーに類する一切。レポートは成果物であり、入力の原本を持ち出す口にしない（16章NFR-S3）。ただし `s1.field_insights[]`（現場メモ由来の原文パススルー。10章FR-34）はS1の出力そのものなので含む。

---

## 3. セクションID一覧と読むJSONパス

**セクションIDは `SEC-01` から `SEC-18`。既存IDの改番・再利用を禁止する。追加は `SEC-19` 以降を使う。** `slug` はHTMLの `id` 属性とJSの登録キーであり、IDと1対1で対応する（`<section id="sec-cover">`）。並び順は本表の上から下（v1.2で見本の10節の流れへ並べ替えた。§3.0の対応表がどの節にどのセクションが集まるかを持つ）。

### 3.0 見本の10節 ⇔ セクションの対応（描画のまとまり。11章§3.8.1が正）

体裁の正である見本は**10節構成**（上部ナビのアンカーと1対1）である。本章のセクションはその10節へ次のように集まる。**「集まる」は描画上のまとまりであって、登録表からIDを消すことではない**（例: SEC-11 は SEC-07 と同じ節に並ぶが、登録行も「空のときの挙動」もそのまま残る）。キッカー（`01 / Executive Summary` のような節番号＋英字ラベル）は**節の先頭セクションにだけ**付ける。

| 見本の節（アンカー / キッカー） | 集まるセクション | キッカーを持つセクション |
|---|---|---|
| （ヒーロー・上部ナビ） | SEC-01 cover | ― |
| `#sec-exec` / `01 / Executive Summary` | SEC-02 exec | SEC-02 |
| `#sec-profile` / `02 / Business Understanding` | SEC-03 profile ＋ SEC-04 sufficiency | SEC-03 |
| `#sec-riskuniv` / `03 / MECE Risk Universe` | SEC-05 riskuniv | SEC-05 |
| `#sec-riskmap` / `04 / Risk Map` | SEC-06 riskmap（＋ SEC-16 round-update を直後に置く） | SEC-06 |
| `#sec-risks` / `05 / Insurance Coverage Matrix` | SEC-07 risks ＋ SEC-08 coverage ＋ SEC-11 prevent ＋ SEC-12 limit | SEC-07 |
| `#sec-newrisk` / `06 / New Risk Radar` | SEC-09 newrisk | SEC-09 |
| `#sec-growth` / `07 / Insurance-enabled Growth` | SEC-17 growth | SEC-17 |
| `#sec-story` / `08 / Executive Proposal Story` | SEC-10 story ＋ SEC-18 talk | SEC-10 |
| `#sec-hearing` / `09 / Discovery Questions` | SEC-13 hearing | SEC-13 |
| `#sec-source` / `10 / Sources and Methodology` | SEC-14 source ＋ SEC-15 disclaimer | SEC-14 |

キッカーの文字列は `modHtmlTemplate6.LabelJs` の `KICK`（セクションID -> キッカー）が持つ（登録表の7キーを増やさないため。§4.2）。

| ID | slug | 見出し（既定） | 読むJSONパス（15章のプロパティ名） | 空のときの挙動 | 図表種別 |
|---|---|---|---|---|---|
| SEC-01 | cover | （表紙。見出しなし） | `meta.company` / `meta.case_id` / `meta.case_type` / `meta.dossier_tier` / `meta.quality_mode` / `meta.round_no` / `meta.industry_name` / `meta.generated_at` / `s1.company_name` | 常に表示（`s1` が null でも `meta` だけで描ける） | 見出し＋チップ列 |
| SEC-02 | exec | エグゼクティブサマリ | `s1.business_summary` / `s1.strategy_outlook.market_context` / `s2.risks[]`（`risk_no` `risk_name` `impact_score` `frequency_score`） / `s3.stories[]`（`story_no` `headline` `pitch` `target_risk_nos`） | 常に表示。`s3` が null のときはテーマ3本を省き「提案ストーリーは未生成です」の1行 | 文章中心（§3.1） |
| SEC-03 | profile | 企業理解 | `s1.business_summary` / `s1.main_products[]` / `s1.processes[]` / `s1.locations[]`（`name` `type` `address` `hazard_note` `notes`） / `s1.supply_chain.key_materials[]` `s1.supply_chain.notes` / `s1.customers.segments[]` `s1.customers.channels[]` / `s1.workforce_notes` / `s1.management_notes` / `s1.strategy_outlook.mvv` `aspirations[]` `market_context` | `s1` が null なら**セクションごと非表示** | 定義リスト＋拠点表 |
| SEC-04 | sufficiency | 入力の充足度と要確認事項 | `s1.input_quality.coverage[]`（`aspect` `status`） / `s1.input_quality.overall` / `s1.input_quality.advice` / `s1.missing_info[]`（`item` `why_needed`） | `s1` が null なら非表示 | 14観点バッジ＋表 |
| SEC-05 | riskuniv | リスクユニバース10分類 | `s2.risks[].category`（19章§3の日本語ラベルへ変換） / `s2.risks[].risk_no` | `s2` が null なら非表示 | 10分類の件数バー（§3.2） |
| SEC-06 | riskmap | 2軸リスクマップ（影響×頻度 5×5） | `s2.risks[]`（`risk_no` `risk_name` `impact_score` `frequency_score` `insurability.transferability`） | `s2` が null なら非表示。`risks` が0件なら「該当なし」の空マップを描く | 5×5マトリクス（§3.3） |
| SEC-16 | round-update | 訪問で分かったこと（ラウンド更新） | `meta.round_no` ／ `s2.risks[]` のうち `status` が `new`（新たに浮上した仮説）／ `confirmed`（裏が取れたリスク）／ `rejected`（否定された仮説）のもの（`risk_no` `risk_name` `scenario` `category` `status`）。**スキーマ変更はなく `status` によるフィルタのみ** | `meta.round_no` が2未満（初回ラウンド）、または3つの `status` がいずれも0件なら**セクションごと非表示**（目次からも落とす） | 3ブロック（新たに浮上した仮説／裏が取れたリスク／否定された仮説。rejected は見出しに取り消し表現を付し、`scenario` 末尾に追記された否定の理由をそのまま残す） |
| SEC-07 | risks | リスク一覧 | `s2.risks[]` の全項目（`risk_no` `category` `risk_name` `scenario` `status` `frequency` `impact` `frequency_score` `impact_score` `evidence.quote` `evidence.source` `insurability.transferability` `insurability.line_note` `insurability.gap_note` `insurability.control_note` `loss_scale_note` `check_points[]` `preventions[].measure` `preventions[].related_menu_id`） | `s2` が null なら非表示 | 表（横スクロール可） |
| SEC-08 | coverage | 保険カバレッジ表 | ①`s2.risks[]`（`risk_no` `category` `risk_name` `impact_score` `frequency_score` `insurability.transferability` `insurability.line_note` `insurability.gap_note` `insurability.control_note`） ②`s1.current_coverage[]`（`line_name` `coverage_summary` `limit_note` `special_note` `certainty`） ③`s2.gaps[]`（`gap_no` `gap_type` `target` `description` `risk_evidence` `coverage_evidence`） | **`s2.risks[].insurability` があれば必ず描く**（v1.3・裁定書25 S1。旧「両方0件なら非表示」は撤回した）。`current_coverage` が0件（新規案件）なら②を省き「新規案件のため現契約なし。以下は必要補償の見立て」の注記を出す。`gaps` が0件なら③を省く。`s2` が null のときだけセクションごと非表示 | リスク単位の表＋2枚組の表（§3.8） |
| SEC-11 | prevent | 未然防止メニュー | `s2.risks[]`（`risk_no` `risk_name` `preventions[].measure` `preventions[].related_menu_id`） | `preventions` が全リスクで0件なら非表示 | 表 |
| SEC-12 | limit | 当社にできないこと・提案を控えること | ①`s2.risks[]` のうち `insurability.transferability` が `hard`（`risk_no` `risk_name` `insurability.control_note`） ②`s3.unmatched_risks[]`（`risk_no` `risk_name` `why_unmatched`） ③`s3.do_not_propose[]`（`topic` `reason`） | 3ブロックとも0件なら「該当なし」の1行を出す（**非表示にしない**。10章FR-37「できないことを正直に書く」がこのセクションの存在理由であるため） | 3ブロック |
| SEC-09 | newrisk | ニューリスク（新種・新興リスク） | `s2.emerging_risks[]`（`risk_name` `category` `horizon` `scenario` `evidence_quote` `evidence_source` `proposal_hint`）。`category` と `horizon` は19章§3の日本語ラベルへ変換する | 0件（空配列）のときは「現時点で特筆すべきニューリスクは検出されていません」の1行を出す（**非表示にしない**。「見ていない」のではなく「見たうえで該当が無い」ことを読み手に示すため） | カード |
| SEC-17 | growth | 攻めの保険活用 | `s3.growth_ideas[]`（`title` `what` `why` `insurance_fit` `effect` `difficulty`）。`difficulty` は19章§3の日本語ラベル（低／中／高）へ変換する | `s3` が null、または `growth_ideas` が0件なら**セクションごと非表示**（目次からも落とす）。**「該当なし」の1行は出さない**（SEC-09・SEC-12 と扱いが違う。発想が出なければ出さないだけの節であるため） | 順位バッジ＋★5段階＋難度ピル（§3.7） |
| SEC-10 | story | 提案ストーリー（当社にできること） | `s3.stories[]` の全項目（`story_no` `proposal_kind` `headline` `hook_question` `target_risk_nos[]` `target_gap_nos[]` `menu_ids[]` `line_ids[]` `scheme_id` `pitch` `similar_case_id` `expected_objection` `objection_response`）。`target_risk_nos` は `s2.risks[].risk_no` を、`target_gap_nos` は `s2.gaps[].gap_no` を引いて名称に解決する | `s3` が null なら非表示 | カード3枚 |
| SEC-18 | talk | 経営層への話し方 | `s3.talk_script`（`opening` `flow[]` `closing` `taboo[]`） | `s3` が null、または `talk_script` が無いなら**セクションごと非表示**（目次からも落とす） | 吹き出し＋番号付きの流れ（§3.9） |
| SEC-13 | hearing | ヒアリング事項 | `s3.stories[].hook_question` / `s2.open_questions[]` / `s1.missing_info[]`（`item` `why_needed`） / `s2.risks[].check_points[]` | 4系統すべて0件なら非表示 | 番号付きリスト（§3.4） |
| SEC-14 | source | 出典と根拠 | `s2.risks[]`（`risk_no` `evidence.quote` `evidence.source`） | `s2` が null なら非表示 | 表 |
| SEC-15 | disclaimer | 免責とご確認事項 | `meta.company` / `meta.generated_at` / `meta.app_version` / `meta.case_id` ＋ §3.5の固定文 | **常に表示（非表示にできない唯一のセクション）** | フッタ |

- **S4は読まない**。本レポートは `GenerateHtmlReport` の契約どおり S1・S2・S3 だけから描く（14章§6）。S4の `hearing_questions` はヒアリングシート（13章§2.16・`modExportHearing`）の入力であり、S4未実行でもレポートが出せる状態を保つため本章では参照しない。
- enum値は必ず19章§3・15章§0の日本語ラベルへ変換して表示する。生の英字enumを画面に出さない。**充足度 `input_quality.overall` のラベルも19章§3の「高／中／低」を逐語で用いる**（v1.1.1・裁定書10 m4。旧「充足度 高／充足度 中／充足度 低」のラベル辞書は廃止し、同一enumのラベルが画面側〔19章§3・`modUICase.EnumPairsCsv`〕と2系統に分岐しない状態を保つ。実装のラベル辞書は `modHtmlTemplate6` の LIQO＝`{high:'高',mid:'中',low:'低'}`。**SEC-04 の総合充足度の段落は「充足度: 」を前置して `充足度: 高` の形で描画する**〔v1.1.1・裁定書10補遺P5。ラベル辞書は1字のままとし、前置は表示側 `modHtmlTemplate2` の SEC-04 描画が持つ〕）。
- 「空のときの挙動」が「非表示」のセクションは、目次（§3.6）からも同時に落とす。
- **SEC-09 と SEC-16 は別物である**: SEC-09「ニューリスク」は `s2.emerging_risks[]`＝サイバー・気候変動のような**新種・新興リスク**（ラウンドに関係なく初回から出る）、SEC-16「訪問で分かったこと」は `s2.risks[].status`＝**第2ラウンド以降の仮説ライフサイクル**（新規発見・確認済み・棄却）である。`status = "new"` のリスクは SEC-16 と SEC-07 リスク一覧の `status` 表示（バッジ）に留め、**SEC-09 には出さない**。

### 3.1 SEC-02 エグゼクティブサマリの構成（10章FR-37「3テーマ・A4 1枚相当の文字中心」）

1. **リード**: `s1.business_summary` の全文（改行はそのまま段落に変換）。続けて `s1.strategy_outlook.market_context` を1段落。
2. **最重要リスク3件**: `s2.risks[]` を `impact_score * frequency_score` の降順、同点は `impact_score` の降順、なお同点は `risk_no` の昇順で並べ、上位3件の `risk_no` と `risk_name` を1行ずつ。
3. **3テーマ**: `s3.stories[]` を `story_no` 昇順に3件。各テーマは `headline`（小見出し）／対象リスク（`target_risk_nos` を `s2.risks[].risk_name` へ解決し「・」で連結）／`pitch` の先頭200字（超過時は末尾に「…」を付す）。
4. 本セクションだけは印刷時に `break-after: page` を効かせ、A4 1枚に収める（§6）。以降のセクションは文字を減らし図表主体にする。

### 3.2 SEC-05 リスクユニバース10分類の描き方

19章§3の10分類を**常に10行**（該当0件の分類も0件と表示する）並べ、各行に件数と横バーを描く。バー長は `件数 / 全risks件数` の比。分類の並びは15章§0の変換表の記載順（strategy_market から brand_social まで）に固定する。件数0の分類を落とさないのは「見ていない領域」と「見たが該当なし」を読み手が区別できるようにするためである。

### 3.3 SEC-06 2軸リスクマップ 5×5 の描き方

- **縦軸=`impact_score`（上が5・下が1）／横軸=`frequency_score`（左が1・右が5）** の5行5列の表。軸ラベルは縦「影響度」横「発生頻度」。
- 各セルには該当する `risk_no` のバッジを昇順に並べる。バッジの色は `insurability.transferability` で決める（`cover`=`--tr-cover` / `partial`=`--tr-partial` / `hard`=`--tr-hard`）。バッジには番号のみを書き、`risk_name` は `title` 属性と、マップ直下の凡例表（`risk_no` と `risk_name` の対応）で示す。
- **セル背景の帯（heat）**は `impact_score + frequency_score` の和で決める。2から3=`--heat-1`／4から5=`--heat-2`／6から7=`--heat-3`／8から9=`--heat-4`／10=`--heat-5`。
- **色に依存させない**: 各セルの右下に帯番号（1から5）を小さく添える。白黒印刷・色覚特性・背景色印刷が無効な環境でも重篤度が読めるようにするための必須要素であり、省略しない。
- 1セルに6件以上入る場合は先頭5件のバッジ＋`+n` の表記に切り替える（セル高が崩れて5×5の形が失われるのを防ぐ）。

### 3.4 SEC-13 ヒアリング事項の生成規則（順序と件数を固定する）

次の4系統をこの順に連結し、通し番号を振る。**合計20問を上限**とし、超過分は切り捨てて末尾に「ほか{n}問（ヒアリングシートを参照）」の1行を出す。

| 順 | 系統 | 出典 | 並び | 上限 |
|---|---|---|---|---|
| 1 | 提案の切り口 | `s3.stories[].hook_question` | `story_no` 昇順 | 3問 |
| 2 | 未解決の論点 | `s2.open_questions[]` | 配列順 | 5問 |
| 3 | 不足情報 | `s1.missing_info[]`。設問文は `{item}について教えてください`、補足に `why_needed` を添える | 配列順 | 5問 |
| 4 | リスクの確認点 | `s2.risks[].check_points[]`。§3.1の順で上位5件のリスクに属するもののみ、1リスクあたり先頭2点まで | リスクの重要度順、同一リスク内は配列順 | 10問 |

同一文字列の設問が複数系統から出た場合は先に出た系統を残して後を落とす（比較は前後空白を除去した完全一致）。

### 3.5 SEC-15 免責フッタの固定文（この4行を必ず含める）

1. `本資料はAIが公開情報等から作成した営業担当者向けの分析資料です（AI生成・担当者確認前）。お客さまへ提示する前に、担当者が内容を確認・編集してください。`（16章NFR-S5の必須表記）
2. `記載のリスクは公開情報と当社担当者の見立てに基づく仮説であり、引受可否・保険料・幹事構成を確約するものではありません。`
3. `保険料の試算は本資料の対象外です（要見積）。`（10章FR-43）
4. `{meta.company} 御中 / 案件ID {meta.case_id} / 作成 {meta.generated_at} / リスク提案ナビ v{meta.app_version}`

**1行目だけは3項分岐する**（v1.4・裁定書37 B-06）。`meta.reviewed_by` が**空**なら上の既定文、**非空**なら次の1文へ差し替える（2行目以降は不変）。

> 本資料はAI支援により作成した骨子を担当者が確認・編集したものです（確認: {meta.reviewed_by} / {meta.reviewed_at}）。

- 差し替えの理由: 社内IT環境v1.1 §7.3 は顧客提示物に利用者の確認を必須と定めるが、**製品側にその担保が無いまま「人が確認・編集した」と断言していた**（裁定書37 B-06/A-01/C-5）。確認を通していない書き出しでは遵守を名乗らない。
- **`<noscript>` 側（§4.1）も同じ分岐を静的HTMLで行う**。文言が2箇所に複製される構造なので、片方だけ直る腐敗を層(a)のテスト（`modTestsPure24`）が止める。
- 表紙 SEC-01 のチップ列には `確認前` / `確認済 {meta.reviewed_by}` を出す。
- `reviewed_by` を立てる口は `modExportHtml.GenerateHtmlReportEx(caseId, outPath, reviewedBy)`。旧 `GenerateHtmlReport(caseId, outPath)` は `reviewedBy=""`（＝確認前）で委譲する。**編集の有無から自動判定しない**（「見たが直さなかった」を落とすため。裁定書37 C-5）。

### 3.6 上部ナビと目次（v1.2で改訂）

表示対象となったセクションの見出しをページ内リンク（`#sec-<slug>`）で並べる。並びは登録表から自動生成し、手で持たない。出し先は次の2つで、**中身は同じ1本の一覧から作る**（2箇所に並びを持たない）。

| 出し先 | id | 見え方 | 印刷 |
|---|---|---|---|
| 上部ナビ | `toc` | ページ最上部に `position:sticky; top:0` で貼り付く帯（見本の `.topbar`）。ヒーローより前に置く | **消す**（`@media print{.topbar{display:none}}`） |
| 本文先頭の目次 | `tocprint` | 画面では非表示（`.print-only`） | **出す**（紙でも構成が追えるようにする。ページ番号は付けない） |

**見出しが空（SEC-01 cover）のセクションは、どちらにも並べない。** 「空のときの挙動」が「非表示」で落ちたセクションも同時に落とす。

### 3.7 SEC-17 攻めの保険活用の描き方（v1.2で新設。11章§3.8.2b が正）

`s3.growth_ideas[]`（15章§4 Schema-S3。4〜8件）を描く。**`s3.stories[]` と同じカードで描かない・同じ節に置かない**（`stories` は目の前のリスクへの打ち手、`growth_ideas` は事業機会。混ぜると「提案3本」の意味が壊れる。11章§9-9）。

| 項目 | 規約 |
|---|---|
| 並び | `effect` の降順 → 同点は `difficulty` の易しい順（`low` → `mid` → `high`） → なお同点は配列順 |
| 順位 | 並べたあとの順位（1から）を44px角の `.rank` バッジに出す |
| `effect` | **★を5つ並べて塗り分ける**（`effect` 個を `--accent`、残りを `--line`）。★の数だけでなく `効きめ {n}/5` の文字も併記する（色と記号だけで意味を運ばない。§6③） |
| `difficulty` | 19章§3の日本語ラベル（低／中／高）のピル。`low`=`--matsu` / `mid`=`--kaki` / `high`=`--tr-hard` |
| `insurance_fit` | 右側の補足カラム（`.idea-side`）に `保険との接点` の見出しを添えて置く |
| 0件 | **セクションごと非表示**（目次からも落とす）。「該当なし」の1行は出さない |
| 免責 | 節の先頭に次の1文を**逐語で**置く（見本と同一。`.note` で囲む） |

1. `実現可否は保険業法、約款設計、募集スキーム、料率、データ取得可否、対象顧客の同意等の検討が必要です。ここではアイデア発散を優先しています。`

`menu_ids` / `line_ids` は**持たない**（15章§4。実在しないメニューIDを引く経路を作らない）ため、本節にID列は無い。

### 3.8 SEC-08 保険カバレッジ表の描き方（v1.3で新設。裁定書25 S5。見本05節が正）

見本05節（`Insurance Coverage Matrix`）の粒度は**契約単位ではなくリスク単位**である。契約が1本も無い新規先でも「このリスクは、ふつうどの保険で、どこまで移せて、何を確かめる必要があるか」が並ぶことが、この節の価値である（髙橋FB①「事業内容×既存の保険」のリスクマップ）。したがって本セクションは**3枚の表**を上から順に描く。

**(1) リスク単位の表（主表。`s2.risks[]` を全件）**

| 列 | 値 | 規約 |
|---|---|---|
| No | `risk_no` | 昇順。SEC-06 リスクマップのバッジ番号と同じ番号である |
| 分類 | `category` | 19章§3の日本語ラベルへ変換 |
| リスク | `risk_name` | |
| 評価 | `impact_score` / `frequency_score` | `影響{i} × 頻度{f}` と積（`i*f`）を併記する。積の帯は SEC-06 の heat と同じ区切り（2-3 / 4-5 / 6-7 / 8-9 / 10）を使う |
| 想定される既存商品 | `insurability.line_note` | 空文字なら `-` |
| カバー可能性 | `insurability.transferability` | **3値バッジ**（`cover`=`--tr-cover` / `partial`=`--tr-partial` / `hard`=`--tr-hard`）。ラベルは19章§3の日本語（比較的移転しやすい／条件付き・部分的／保険化困難）。**色だけで意味を運ばない**ため、バッジには必ず日本語ラベルの文字を入れる（§6③） |
| ギャップ・確認点 | `insurability.gap_note` | 空文字なら `-` |
| 管理策 | `insurability.control_note` | 空文字なら `-` |

- 並びは `risk_no` 昇順（SEC-07 リスク一覧と同じ）。**分類ごとにグループ化しない**（見本と同じく1枚の連続した表とし、`category` 列で読み分ける）。
- 横幅が足りない場合は**表を `overflow-x: auto` の枠に入れる**。列を落とさない（8列すべてが揃って初めて「事業内容×既存の保険」の突合になる）。

**(2) 現契約の表（`s1.current_coverage[]`）**: 列は 種目名（`line_name`）／補償内容（`coverage_summary`）／限度額（`limit_note`）／主要特約・免責（`special_note`）／**確度（`certainty`。19章§3の日本語ラベル。`assumed` の行は行全体を淡色にし、表の直前に「確認前の見立てを含みます」の1行を出す）**。0件（新規案件）のときは表ごと省き、代わりに `新規案件のため現契約なし。以下は必要補償の見立て` の注記を1行出す。

**(3) 未充足リスク一覧の表（`s2.gaps[]`）**: 列は No（`gap_no`）／種別（`gap_type`。19章§3の日本語ラベル）／対象（`target`）／説明（`description`）／リスク側の根拠（`risk_evidence`）／契約側の根拠（`coverage_evidence`）。0件なら表ごと省く（「該当なし」の1行も出さない。(1)が本節の本体であるため）。

**空のときの挙動（v1.3で変更）**: `s2` が null のときだけセクションごと非表示にする。**`current_coverage` と `gaps` が両方0件でも (1) を描く**（旧規定「両方0件なら非表示」は撤回。裁定書25 S1）。

### 3.9 SEC-18 経営層への話し方の描き方（v1.3で新設。裁定書25 S2。見本08節が正）

`s3.talk_script` を描く。見本08節は「統合ストーリー（大見出し＋リード）＋ STEP1-4 の4枚」であり、本セクションはその**話す順番**の側を受け持つ。SEC-10 提案ストーリーと**同じ節（見本08節）に並べ、SEC-10 の直後に置く**（キッカーは節の先頭である SEC-10 が持つ。§3.0）。

| 部分 | 値 | 規約 |
|---|---|---|
| 切り出しの一言 | `opening` | 節の先頭に大きめの1文で置く（見本の統合ストーリー大見出しの位置）。80字以内である前提で折り返す |
| 話す順番 | `flow[]` | **配列順のまま**番号付きカードで横（印刷時は縦）に並べる。並べ替え・要約をしない。3～5枚 |
| 締めの一言 | `closing` | カードの下に1文 |
| 触れない事 | `taboo[]` | 最後に注意の枠（`.note`）で箇条書き。**0件なら枠ごと出さない**（「該当なし」を出すと、制約が無い先で不必要に身構えさせるため） |

- 見出しは `経営層への話し方`（§3の表が正）。**画面に「トークスクリプト」「talk_script」の語を出さない**（利用者向け文の規約。11章§7）。
- 節の先頭に次の1文を**逐語で**置く（`.note` で囲む）。

1. `そのまま読み上げる原稿ではありません。話す順番の下書きとしてお使いください。`

- `taboo[]` の各項目は営業が顧客の前で見る文である。**その根拠（`field_insights` の原文）は SEC-18 には出さない**（現場メモの生の言い回しがレポートに載ると顧客同席の場で事故になる）。根拠を追いたい場合は画面側（13章§2.12 の S1 シート）で確認する。

---

## 4. セクション登録表と、図表テンプレートを1つ追加する手順

### 4.1 描画の場所

DATAはページ内のJavaScriptが `JSON.parse` で受け取り、**セクションの描画はブラウザ側のJSが行う**（10章FR-37「データはJS配列としてテンプレートに埋め込む」）。VBAは値ごとのHTML断片を組み立てない。この分担により、図表の追加・並べ替え・見た目の変更が**VBAのロジックに触れずに済む**。

- VBAが静的HTMLとして書き出すのは次の3つだけで、いずれも `modUtilText.HtmlSafe` を通す: (a) `<title>` (b) SEC-01 表紙の会社名・案件ID・生成日時 (c) `<noscript>` の案内文。
- **JS側の描画は `document.createElement` と `textContent` への代入のみで行う。`innerHTML` / `insertAdjacentHTML` / `document.write` / `outerHTML` への代入を禁止する。** 属性は `setAttribute` で与える。これによりDATA由来の文字列がマークアップとして解釈される経路が構造的に存在しなくなる（17章 T-46 の出荷前検問で、テンプレ文字列中にこれらの語が出現しないことを grep 検査する）。
- `<noscript>` には「このレポートの表示にはJavaScriptが必要です。ファイルをローカルに保存してブラウザで開いてください」と、SEC-15の免責文（§3.5の1行目。**`meta.reviewed_by` による3項分岐も同じ**。v1.4・裁定書37 B-06）を静的HTMLで書く。スクリプトが動かない環境でもAI利用の明示だけは必ず読めるようにする。

### 4.2 セクション登録表の形式

登録表は `modHtmlTemplate1.SectionsJs() As String` **1関数の中だけ**にある。この関数はセクションごとに**2行**を持つ。

```
' ---- modHtmlTemplate1.SectionsJs() の中身（返す文字列の形） ----
' (a) 描画関数の連結行: セクションの描画スクリプト本体を取り込む
s = s & modHtmlTemplate2.SecCoverJs()
s = s & modHtmlTemplate2.SecExecJs()
...
' (b) 登録配列: 描画の順序と条件を宣言する。ここに並んだ順に描画される
s = s & "var SECTIONS=[" & vbLf
s = s & " {id:'SEC-01',slug:'cover',    title:'',                  need:['meta'],   empty:'always',render:renderCover},"      & vbLf
s = s & " {id:'SEC-02',slug:'exec',     title:'エグゼクティブサマリ',need:['s1'],    empty:'always',render:renderExec},"       & vbLf
s = s & " {id:'SEC-06',slug:'riskmap',  title:'2軸リスクマップ（影響×頻度 5×5）',need:['s2'],empty:'hide',render:renderRiskMap}," & vbLf
s = s & " {id:'SEC-09',slug:'newrisk',  title:'ニューリスク（新種・新興リスク）',need:['s2'],empty:'note',note:'現時点で特筆すべきニューリスクは検出されていません',render:renderNewRisk}," & vbLf
s = s & " {id:'SEC-16',slug:'round-update',title:'訪問で分かったこと（ラウンド更新）',need:['s2'],empty:'hide',render:renderRoundUpdate}," & vbLf
s = s & " {id:'SEC-17',slug:'growth',   title:'攻めの保険活用',      need:['s3'],   empty:'hide', render:renderGrowth},"    & vbLf
s = s & "];" & vbLf
```

登録行のキーは**この7つに固定**する。増やさない。

| キー | 型 | 必須 | 意味 |
|---|---|---|---|
| `id` | 文字列 | ○ | `SEC-nn`。§3の表と一致させる。改番禁止 |
| `slug` | 文字列 | ○ | `<section id="sec-<slug>">` と目次アンカーになる。英小文字とハイフンのみ |
| `title` | 文字列 | ○ | 既定の見出し。**§3の表の「見出し（既定）」列を逐語で写す**（括弧つきのフル表記まで含めて一致させる。上のコード例もその写しであり、食い違ったら§3が正）。空文字は見出しを出さない（SEC-01のみ） |
| `need` | 配列 | ○ | 描画に必要なDATAのトップキー（`meta` / `s1` / `s2` / `s3`） |
| `empty` | 文字列 | ○ | **`need` のいずれかが `null`、または描画対象が0件**のときの挙動。`always`=それでも描く（§3で常時表示と決めたもの） / `hide`=描かずに目次からも落とす / `note`=見出しと `note` の1行だけを描く |
| `note` | 文字列 | `empty:'note'` のときのみ | 0件時に出す1行の本文 |
| `render` | 関数名 | ○ | (a)の連結行で取り込んだ描画関数。引数は `(DATA, sectionEl)` の2つに固定し、戻り値を持たない |

キッカー（§3.0）は登録行のキーにしない。`modHtmlTemplate6.LabelJs` の `KICK` がセクションIDから引く（**7キーを増やさない**という本節の規約を守るため）。

`empty` の値は§3の表の「空のときの挙動」列と、`title` の値は同表の「見出し（既定）」列と1対1で対応させる（本表・コード例と§3が食い違ったら§3が正）。目次（§3.6）は登録表の `title` をそのまま並べるため、この一致が崩れると本文と目次の両方が同時に漂流する。

### 4.3 図表テンプレートを1つ追加する手順（変更は2箇所で完結する）

例: 「拠点ハザードの一覧図を足したい」という要望を受けた場合。

1. **テンプレ関数を1本追加する（1箇所目）**: 空きのあるテンプレモジュール（§4.4の分割規約に従い、超過していれば新しい `modHtmlTemplateN`）に `Public Function SecHazardJs() As String` を追加し、`function renderHazard(DATA, el){...}` を返す。既存の描画関数・CSS・他セクションには一切触らない。
2. **登録表に2行足す（2箇所目）**: `modHtmlTemplate1.SectionsJs()` に、(a) `s = s & modHtmlTemplateN.SecHazardJs()` の連結行と (b) `{id:'SEC-17',slug:'hazard',title:'拠点ハザード',need:['s1'],empty:'hide',render:renderHazard},` の登録行を、出したい位置に挿入する。
3. **§3の表に1行足す**（本章の更新）。IDは次の空き番（現在は `SEC-18`）。既存IDは動かさない。§3.0の対応表にも、その節のどこへ入るかを1行足す。

`modExportHtml`・`modHtmlTheme`・他のテンプレ関数・CSSは変更しない。**新しい配色が必要な場合でも新しいCSS変数を足さず、§5.1の閉じた一覧から選ぶ**（一覧を増やすとテーマ側の全定義に追随が必要になり、テーマ差替が1モジュールで閉じなくなるため）。一覧の拡張が本当に必要なときは本章§5.1の改訂として扱い、`modHtmlTheme` の全テーマを同時に更新する。

セクションの**削除**は登録行の削除（および§3の表からの削除）だけで行い、テンプレ関数は残してよい（呼ばれなくなるだけ）。セクションの**並べ替え**は登録配列の行順の入れ替えだけで行う。

### 4.4 テンプレモジュールの分割規約（1モジュール30,000字契約）

- 1モジュールが**25,000字**を超えたら次番のモジュールへ切り出す（30,000字の契約に対して余白を持たせる）。**この閾値は `tools/vba_lint.py` が `modHtmlTemplate*` 専用のERROR（`TEMPLATE_MAX_CHARS = 25000`）として機械強制する**（17章 T-35 のDoDが本章の閾値を明記したうえで検査をlintへ委ねる。W3では28,000字のWARN帯に届かず `modHtmlTemplate1` の25,358字が9ゲート全緑のまま素通りしたため、v1.1で委譲先を明示した）。
- 分割は**関数単位**で行い、関数名は変えずに移動だけする。同名の `Public Function` を2つ以上のモジュールに置かない。切り出しに伴う `Private` → `Public` の変更は関数名の変更ではないので可。
- 既定の割り当て（v1.1で実態へ更新。v1.3で `modHtmlTemplate8` を追加。1..8）:

| モジュール | 持つもの |
|---|---|
| `modHtmlTemplate1` | `BuildDocument`（全体組立）／`HeadHtml`（`<meta charset>`・`<title>`・共通CSSの呼び口・テーマCSSの差込口）／`BodyShellHtml`（骨格と `<noscript>`）／**`SectionsJs`（§4.2のセクション登録表。編集が最も多い1関数）**／`RuntimeJs`（目次生成・登録配列の走査・`need`/`empty` の判定・共通の描画ヘルパ） |
| `modHtmlTemplate2` | SEC-01 cover ／ SEC-02 exec ／ SEC-03 profile ／ SEC-04 sufficiency |
| `modHtmlTemplate3` | SEC-05 riskuniv ／ SEC-06 riskmap ／ SEC-07 risks ／ SEC-08 coverage |
| `modHtmlTemplate4` | SEC-09 newrisk ／ SEC-16 round-update ／ SEC-10 story |
| `modHtmlTemplate5` | SEC-11 prevent ／ SEC-12 limit ／ SEC-13 hearing ／ SEC-14 source ／ SEC-15 disclaimer |
| `modHtmlTemplate6` | `LabelJs`（19章§3・15章§0のenum変換表を返す。`RuntimeJs` から呼ぶ下請け。19章の改訂でしか動かない表を、編集が最も多い `SectionsJs` と同じモジュールに置かないための切り出し） |
| `modHtmlTemplate7` | SEC-17 growth（§3.7）／**`CommonCss`**（骨格の共通CSS。v1.3・T-56 で `modHtmlTemplate1` が25,000字を超えたため、§4.4の分割規約どおり**関数名を変えずに**移した。呼び口は `HeadHtml` の1本のまま）／`PartsCss`（見本の部品CSS＝ヒーロー・上部ナビ・カード・表・ユニバース・ヒートマップ・レーダー・アイデア・提案ブロック・設問カード。`HeadHtml` が `CommonCss` の直後に連結する。v1.2で `modHtmlTemplate1` が25,000字を超える見込みになったため分けた。§4.4の「共通CSSは HeadHtml に一元化」は**呼び口が1本であること**を意味しており、字数規約で切り出した下請けは同じ一元化の中にある） |

| `modHtmlTemplate8` | SEC-18 talk（§3.9）／`TalkCss`（吹き出しと番号付きカードの部品CSS。`HeadHtml` が `PartsCss` の直後に連結する。v1.3で新設。`modHtmlTemplate7` の字数余白ではなく新モジュールで受けるのは§4.4の25,000字規約による） |

- 上の表は**現時点の実態**であり、25,000字規約に従って切り出した結果はここへ反映する（表と実装がずれたまま放置しない）。セクションの担当モジュールは§4.3の手順1が「空きのあるテンプレモジュール」と定めるとおり流動的で、正は登録表(§4.2)の(a)連結行である。

- 共通CSSは `modHtmlTemplate1.HeadHtml` に一元化し（実体は `modHtmlTemplate7.CommonCss` ＋ `modHtmlTemplate7.PartsCss` ＋ `modHtmlTemplate8.TalkCss` の3本を `HeadHtml` が連結する。25,000字規約による分割であって、CSSの持ち主が増えたわけではない）、セクション別のテンプレ関数に `<style>` を書かない（CSSが散ると見た目のフィードバックを1箇所で吸収できなくなる）。セクション固有のスタイルはクラス名を `sec-<slug>-*` の接頭辞で共通CSSに置く。
- 文字列の組み立ては15章と同じ `s = s & "..." & vbLf` 方式とする（`Const` は1論理行1,023字・行継続25本の制約に当たるため使わない。14章§7と同じ理由）。

---

## 5. テーマとCSS変数、エスケープ、文字コード

### 5.1 テーマが定義してよいCSS変数の閉じた一覧（39個。これ以外を定義しない・これ以外を参照しない）

`modHtmlTheme.ThemeCss(themeName)` が返すのは `:root{ ... }` **1ブロックだけ**であり、その中身は下表の39変数の宣言だけである。既定値（`standard`）は**体裁の正である見本**（`docs/design/出力見本_春華堂統合提案_v0.1.html` の `:root`）を出発点にしている。v1.2で28→39へ拡張し、旧 `--ai`（青の主色）は **`--brand`（臙脂）** へ改称した（11章§3.8.2 #1・#2）。

**寸法・書体（6）**

| 変数 | standard の既定値 | 用途 |
|---|---|---|
| `--page-width` | `1180px` | 本文カラムの最大幅（見本の `max-width:1180px`） |
| `--page-pad` | `22px` | 本文カラムの左右余白 |
| `--font-sans` | `-apple-system,BlinkMacSystemFont,"Segoe UI","Hiragino Kaku Gothic ProN","Yu Gothic",Meiryo,sans-serif` | 本文と見出しの両方。**Webフォントを読み込まない**（外部参照禁止）。端末に無い書体は後続へフォールバックし、最後は総称名で必ず解決する |
| `--font-serif` | `"Shippori Mincho","Yu Mincho","Hiragino Mincho ProN",serif` | 明朝見出しを使うテーマ用に**残すが、standard では使わない**（見本は見出しもゴシック。11章§3.8.2 #3） |
| `--font-size` | `14px` | 本文の基準サイズ |
| `--line-height` | `1.7` | 本文の行間 |

**地色と文字（6）**

| 変数 | standard の既定値 | 用途 |
|---|---|---|
| `--bg` | `#F5F6F8` | ページの外地色（カードが浮いて見える下地） |
| `--paper` | `#FFFFFF` | カード・表の地色 |
| `--ink` | `#1D2433` | 本文の文字色 |
| `--sub` | `#667085` | 補助文・注記の文字色 |
| `--mist` | `#F8F9FB` | 表ヘッダ・囲みの薄い面色 |
| `--line` | `#E6E8EC` | 罫線・区切り線 |

**主色と強調（7）**

| 変数 | standard の既定値 | 用途 |
|---|---|---|
| `--brand` | `#A5312F` | **主色**（臙脂）。キッカー・見出しの罫・順位バッジ・リスクマップの点 |
| `--brand2` | `#6F1E1E` | 主色の濃い側。ヒーローのグラデーションの起点 |
| `--accent` | `#B88A44` | 金。★（`effect`）と強調の細部 |
| `--navy` | `#27364A` | 濃紺。提案ブロックの地・絞り込みボタンの選択状態 |
| `--kaki` | `#B4552D` | 注意・できないこと・不足の強調 |
| `--matsu` | `#2F7A54` | 肯定・提案・充足の強調 |
| `--deep` | `#6F42C1` | ニューリスクの時間軸ピルなど、補足の系統色 |

**面色と影（7）**

| 変数 | standard の既定値 | 用途 |
|---|---|---|
| `--soft-brand` | `#F7EEED` | 主色系の淡い面（上部ナビのボタン・順位バッジの地） |
| `--soft-red` | `#FFF1F0` | 否定・危険側のタグの地 |
| `--soft-amber` | `#FFF7E8` | 推定・条件付きのタグの地 |
| `--soft-green` | `#EDF8F2` | 事実・移転しやすいタグの地 |
| `--soft-blue` | `#EEF4FF` | 参照・出典タグの地 |
| `--soft-purple` | `#F5F1FF` | 時間軸ピルの地 |
| `--shadow` | `0 12px 32px rgba(16,24,40,.07)` | カードの影（**色ではなく `box-shadow` の値**。印刷時は共通CSSが `none` へ落とす） |

**注意面（2）**

| 変数 | standard の既定値 | 用途 |
|---|---|---|
| `--warn` | `#FBF3E6` | 注意ボックスの背景 |
| `--warn-line` | `#E8D5B5` | 注意ボックスの罫線 |

**リスクマップの帯（5。§3.3の帯番号1から5に対応）**

| 変数 | standard の既定値 |
|---|---|
| `--heat-1` | `#E9F5EE` |
| `--heat-2` | `#F2F7E9` |
| `--heat-3` | `#FFF6DB` |
| `--heat-4` | `#FFE8D3` |
| `--heat-5` | `#FFD9D7` |

**移転可能性の3値（3。`insurability.transferability` に対応）**

| 変数 | standard の既定値 | 対応enum |
|---|---|---|
| `--tr-cover` | `#2F7A54` | `cover`（比較的移転しやすい） |
| `--tr-partial` | `#B08A2E` | `partial`（条件付き・部分的） |
| `--tr-hard` | `#B42318` | `hard`（保険化困難） |

**入力充足度の3値（3。`input_quality.coverage[].status` に対応）**

| 変数 | standard の既定値 | 対応enum |
|---|---|---|
| `--iq-ok` | `#2F7A54` | `ok` |
| `--iq-partial` | `#B08A2E` | `partial` |
| `--iq-missing` | `#A9B4C0` | `missing` |

**規約**

- 共通CSS（`modHtmlTemplate1.HeadHtml` と `modHtmlTemplate7.PartsCss`）は**色・書体・本文幅をリテラルで書かない**。必ず `var(--xxx)` を通す。唯一の例外は、色付きバッジ・見出し帯の上に載せる文字色 `#fff` であり、これ以外の生の16進色を書かない。**`rgba()` による半透明の重ね（白の被膜・影・枠の透過）は、下地の色を変えずに濃さだけを作るものなので例外とし、テーマ変数にしない**（テーマを増やしても半透明の度合いは変えたくないため）。
- **`.heat1`〜`.heat5` のクラスは `modHtmlTemplate7.PartsCss` が定義する**（v1.3・T-57 追認）。中身は `background:var(--heat-1)`〜`var(--heat-5)` の5行だけで、**SEC-06（リスクマップ）と SEC-08（付保の見立て）が共用する**（§3.6のセル背景の帯と、リスク単位8列表の帯は同じ濃さの尺度＝`impact_score + frequency_score` の和。§3.6）。**節ごとに別クラスを作らない**（同じ意味の色を2箇所で定義すると濃さが割れる）。
- `@page` の余白はCSS変数で解決されないため、テーマ変数にせず共通CSSにリテラルで書く（§6）。テーマから紙面余白は変えられない、と割り切る。
- 機械検査（17章 T-35 のDoD）: (1) `ThemeCss` の戻り値が `:root{` で始まり `}` で終わり、内側が `--` で始まる宣言のみであること (2) 全テーマが上表39変数を**過不足なく**定義していること (3) 共通CSS中に `var(--` を伴わない16進色指定が `#fff` 以外に出現しないこと。

### 5.2 テーマ差替の単位とテンプレ5関数の署名

**テンプレ5関数の署名（W3.1で明文化。14章§6は「本章は宣言を持たない」としてここへ委ねている）**: §4.4の分割表は関数名と持ち物だけを定めて引数を規定しておらず、実装のコメントだけが唯一の根拠になっていた。テンプレ関数は層(a)のテストと `tools/render_report.py` が直接叩く境界なので、ここで契約として固定する。

```vb
' === app: modHtmlTemplate1（純文字列・R4）===
Public Function BuildDocument(ByVal themeName As String, ByVal dataJson As String, _
                              ByVal coverFields As String) As String
' HTML全文。themeName=**解決済み**のテーマ名（未知名のフォールバックは呼出側で済ませる）、
' dataJson=§2のDATA（生のJSON。`JsStringSafe` を通すのは本関数の中＝§5.3(1)）、
' coverFields=下記の3値。
Public Function HeadHtml(ByVal themeName As String, ByVal coverFields As String) As String
Public Function BodyShellHtml(ByVal coverFields As String) As String
Public Function SectionsJs() As String     ' §4.2の登録表。引数を取らない
Public Function RuntimeJs() As String      ' 目次・走査・描画ヘルパ。引数を取らない
```

- **`coverFields` の書式**: §4.1(b)の3値（会社名／案件ID／生成日時）を**タブ（`vbTab`）区切りの1本の文字列**で渡す。`[0]`=会社名 `[1]`=案件ID `[2]`=生成日時。タブが区切りとして安全なのは、外部由来テキストが `modUtilText.SanitizeInput`（16章 E-04）で制御文字を落としてから案件データに入るため。3値をJSONで渡さないのは、テンプレ側にJSONパーサを持たせない（＝純文字列モジュールに留める）ため。**この3値はテンプレ側で `HtmlSafe` を通す**（§4.1・§5.3(2)）。
- 引数を増やすときは本節を先に改訂する（実装のコメントを根拠にしない）。

- **差し替えの単位は `modHtmlTheme` の1モジュールのみ**。テーマを増やす作業は「`ThemeCss` の `Select Case` に分岐を1本足し、39変数を書く」で完結し、`modExportHtml` にも `modHtmlTemplate1..n` にも触れない。
- 契約:

```vb
' === app: modHtmlTheme（純文字列・R4）===
Public Function ThemeNames() As String     ' ";"区切り。先頭が既定テーマ。例: "standard;mono;ds"
Public Function ThemeCss(ByVal themeName As String) As String
' 戻り値は ":root{ ... }" の1ブロックのみ。§5.1の39変数を過不足なく宣言する。
' 未知のテーマ名は standard へフォールバックし、run_log の detail に
' theme_fallback=<要求されたテーマ名> を記録する（黙って既定に戻さない）
```

- テーマは3本（v1.2）。`standard`（§5.1の既定値＝見本の臙脂。カラー画面・カラー印刷向け）／`mono`（白黒印刷・FAX配布向け。主色・強調色を `--ink` と `--sub` の濃淡に、heat 5段を白から薄灰の5段に置き換える）／`ds`（DS版 `docs/design/出力見本_春華堂統合提案_DS版_v0.1.html` の白地＋青の帯。`--brand` を紺青へ振り替えたもの）。`mono` でも§3.3の帯番号と `効きめ {n}/5` の併記があるため、重篤度も効きめも色なしで読める。
- config `html_theme`（13章§2.3・既定 `standard`）の値をそのまま `ThemeCss` に渡す。テーマ名の一覧をconfigに書かない（`ThemeNames` が唯一の一覧）。

### 5.3 エスケープと文字コード（16章 E-47・NFR-S7の③が正。本節はその適用手順）

**(1) DATAの埋め込み**: 次の1行の形でのみ埋める。JSONをJSのオブジェクトリテラルとして直書きすることを禁止する。

```
<script>var DATA=JSON.parse("<<JsStringSafe(dataJson) の結果>>");</script>
```

`modUtilText.JsStringSafe` の適用順は次のとおりで、**この順序を守る**（逆順にすると二重エスケープになる）。

1. `\` を `\\` に、`"` を `\"` に置換する（JSON本文をJSの二重引用符リテラルへ入れるため必須）
2. **`<` を1文字残らず `\u003C` に置換する**（旧規約「`</` を `<\/` へ」はこれに包含されるため v1.1 で置き換えた。理由は直後の注記）
3. 行区切り文字 U+2028 を `\u2028`、段落区切り文字 U+2029 を `\u2029` に置換する
4. その他の制御文字（U+0000 から U+001F）を `\u00XX` に置換する

**なぜ `</` ではなく `<` の全部なのか（v1.1改訂の理由）**: HTMLトークナイザは `<script>` の中身を **script data** 状態で読むが、そこに `<!--` が現れると **script data escaped** へ、続けて `<script` が現れると **script data double escaped** へ遷移する。double escaped 状態では `</script>` が終端として働かず、`-->` が来るまで復帰しない。したがってLLM出力の1フィールドに `<!--<script>`（`-->` を伴わない形）が入るだけで、DATAブロックの正規の `</script>` が食われ、後続のランタイムJSごとスクリプトの中身として飲み込まれ、**ページが1セクションも描かれない真っ白な状態になる**（コード実行は成立しないが、主力出力〔10章FR-37〕が無言で白紙になる）。`</` だけを狙うエスケープではこの経路を塞げない。`<` を1文字残らず `\u003C` へ落とせば `</script>` も `<!--` も `<script` も**構造上生成されえない**（`JSON.parse` が `\u003C` を `<` へ戻すので画面に出る文字は元のまま＝情報は落ちない）。

**この規約の受入条件**は「生成HTMLの `var DATA=JSON.parse("…")` の**文字列リテラル内に生の `<` が1文字も無いこと**」とする。トークナイザの状態を数え上げる検査ではなく1文字の有無で判定できることが本改訂の価値であり、17章 T-35 のDoDで `tools/render_report.py` が機械検査する。

**(2) HTML本文への差し込み**: §4.1の3箇所（`<title>` ・表紙の会社名/案件ID/生成日時・`<noscript>`）は `modUtilText.HtmlSafe`（`& < > " '` のエンティティ化）を通す。JS側の描画は `textContent` と `setAttribute` のみを使うためエスケープ不要であり、逆に `innerHTML` 系を使わないことがエスケープ規約そのものである（§4.1）。

**(3) 文字コード**: 書き出しは **UTF-8・BOMあり**で行う。実体は `modUtil.WriteUtf8File`（`Open For Binary` ＋ 純関数 `modUtilText.Utf8Bytes` の自前エンコード。既存ファイルは消してから作り直す）。**v3.4・裁定書27 W9-B2 で `ADODB.Stream` を撤去した**: このCOM生成は社内AVのAMSIがマクロ型マルウェアの特徴として重く見る形であり（2026-09-02 実測）、機能を保ったまま形だけを配布物から消した。符号化の正しさ（BOM・ASCII・2/3バイト・サロゲートペアの4バイト・CP932外文字・孤立サロゲート→U+FFFD）は層(a)が**手計算のバイト列**で固定し、`Open For Binary` が実際に書いたバイトは層(b)（`T47-W9-03`）が確かめる。VBAの `Open ... For Output` / `Print #` はCP932で書き、非CP932文字（絵文字・環境依存字・一部の丸数字）が `?` へ落ちるため**使わない**。`<head>` の**最初の要素**として `<meta charset="utf-8">` を置く（BOMを見ないブラウザ設定でも文字化けしないようにするための二重化）。

**(4) 数式インジェクション**: HTMLでは先頭の `=` `+` `-` `@` は無害なため `SetCellSafe` 相当の無害化は行わない。HTML経路のガードは (1)(2) と `innerHTML` 禁止で構成する（16章 E-46 の但し書きと同じ扱い）。

---

## 6. 印刷（A4縦）とブラウザ表示の両立規約

**前提**: 同一の1ファイルが「ブラウザで見る資料」と「印刷して持参する資料」の両方になる。片方のためにもう片方を壊さない。

| # | 規約 | 実装 |
|---|---|---|
| ① | 用紙はA4縦・余白は上下14mm/左右12mm | 共通CSSに `@page{size:A4 portrait;margin:14mm 12mm;}` をリテラルで書く（CSS変数は `@page` で解決されないため。§5.1） |
| ② | 画面は中央1カラム、印刷は紙幅いっぱい | 本文ラッパを `max-width:var(--page-width);margin:0 auto;padding:0 var(--page-pad)` とし、`@media print` で `max-width:none;padding:0` に戻す |
| ③ | 背景色印刷が無効でも読める | `body{-webkit-print-color-adjust:exact;print-color-adjust:exact}` を指定するが、**これに依存しない**。リスクマップは帯番号（§3.3）、移転可能性・充足度は色に加えて文字ラベルを併記する。色だけで意味を運ぶ表現を作らない |
| ④ | 図表を紙面で分断しない | カード・表の行・リスクマップ全体・提案カードに `break-inside:avoid` を指定する |
| ⑤ | エグゼクティブサマリはA4 1枚 | SEC-02 に `break-after:page` を指定する（10章FR-37の紙面設計） |
| ⑥ | 広い表は画面で横スクロール、印刷で全幅 | 表は `.tblwrap{overflow-x:auto}` で包み、`@media print` で `overflow:visible` に戻し表の文字を `11px` へ落とす。**印刷時に横スクロールで隠れた列が消えないこと**が受入条件 |
| ⑦ | 印刷に不要な操作要素を消す | 画面のみの要素（先頭の「印刷する」ボタン・目次の折りたたみ）に `.no-print` を付け、`@media print{.no-print{display:none}}` |
| ⑧ | 目次は印刷にも出す | §3.6。紙でも構成が追えるようにする。ページ番号は付けない（CSSだけでは本文中に採番できないため、無理に作らない） |
| ⑨ | 狭い画面でも読める | `@media (max-width:640px)` で見出しを縮小し、2カラムのグリッドを1カラムへ。リスクマップは5×5の形を保ったままセル内バッジを小さくする（1カラム化しない。形が意味だから） |
| ⑩ | 印刷の基準文字サイズ | `@media print{body{font-size:12px;line-height:1.7}}`。画面の `--font-size` とは独立に固定する |
| ⑪ | リンクは同一ファイル内アンカーのみ | 目次の `#sec-<slug>` 以外の `href` を出さない。外部URLを踏ませない（社内配布物としての安全側） |

**受入確認（17章 T-33／T-41 のDoDと対応）**: mockデータ（15章§8）から生成したHTMLを、(a) ブラウザで開いて全セクションが表示される (b) 印刷プレビューでA4縦・エグゼクティブサマリが1枚・リスクマップとカバレッジ表が分断されない (c) `</script>`・`<img onerror=`・`&`・改行・絵文字を含むmock S2を入力しても記号がそのまま文字として表示され、レイアウトが壊れない (d) `html_theme` を `standard` から `mono` へ替えるだけで配色が変わり、構成は変わらない、の4点を確認する。
