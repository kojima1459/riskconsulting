# 18. HTMLレポートテンプレート仕様 v1.0

v1.0変更概要: 仕様書v2.4の実装前監査裁定により新設。10章FR-37の「デザインは差し替え可能なテーマとして分離」を、セクション登録表・CSS変数の閉じた一覧・1モジュールに閉じたテーマ差替として具体化し、12章§2の純文字列モジュール `modHtmlTemplate1..n` / `modHtmlTheme`（17章 T-35）の内容面の正を定義した。

**目的**: 本製品の主力出力であるHTMLリスクレポート（10章FR-37・14章§6 `modExportHtml.GenerateHtmlReport`）について、**出力物の見た目・構成・図表に対するフィードバックを1箇所の修正で吸収できる構造**を固定する。発注者・部会からの「この図を足したい」「この順番を入れ替えたい」「配色を変えたい」という要望が、そのつどHTML生成コードの改造にならないよう、変更点を (a) セクション登録表 (b) テーマのCSS変数 (c) 個々のテンプレ関数 の3種類だけに閉じ込める。

**本章が正であるもの**: セクションID・セクションの並び・各セクションが読むJSONパス・CSS変数名の閉じた一覧とその既定値・テーマ差替の単位・テンプレモジュールの分割規約・図表追加手順・エスケープと文字コード・印刷とブラウザ表示の両立規約。
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
| ⑥ | `ADODB.Stream`（Charset="utf-8"・BOMあり）で書き出す（§5.3） | 書込失敗は E0502。出力先不存在は先に16章 E-21 のフォールバックを試す |
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
    "generated_at": "2026/09/01 14:07:22", "app_version": "2.4.0", "theme": "standard"
  },
  "s1": { "company_name": "...", "business_summary": "...", "...": "Schema-S1 の全キー" },
  "s2": { "risks": [], "gaps": [], "open_questions": [] },
  "s3": { "stories": [], "unmatched_risks": [], "do_not_propose": [] }
}
```

- `meta` の由来: `case_id` / `case_type` / `dossier_tier` / `company` / `industry_code` / `industry_name` / `round_no` / `s4_variant` は13章§2.1『案件一覧』の同名列。`quality_mode` はHOMEの `hm_quality_mode`（13章§2.10。未上書きなら config `quality_mode`）。`app_version` は config `app_version`、`theme` は config `html_theme` を `modHtmlTheme.ThemeCss` で解決したあとの実テーマ名（§5.2のフォールバック後の値）、`generated_at` は生成時刻 `yyyy/mm/dd hh:mm:ss`。
- S3が未実行の案件では `"s3": null` とする（キー自体は必ず置く)。同様にS2未実行は `"s2": null`。`null` のときの各セクションの挙動は§3の「空のときの挙動」列が正。
- **DATAに入れないもの**: 入力貼付テキストの原文（`input_hp` 等）・run_log・err_log・ナレッジ本文・APIキーに類する一切。レポートは成果物であり、入力の原本を持ち出す口にしない（16章NFR-S3）。ただし `s1.field_insights[]`（現場メモ由来の原文パススルー。10章FR-34）はS1の出力そのものなので含む。

---

## 3. セクションID一覧と読むJSONパス

**セクションIDは `SEC-01` から `SEC-15`。既存IDの改番・再利用を禁止する。追加は `SEC-16` 以降を使う。** `slug` はHTMLの `id` 属性とJSの登録キーであり、IDと1対1で対応する（`<section id="sec-cover">`）。並び順は本表の上から下（10章FR-37の紙面順）。

| ID | slug | 見出し（既定） | 読むJSONパス（15章のプロパティ名） | 空のときの挙動 | 図表種別 |
|---|---|---|---|---|---|
| SEC-01 | cover | （表紙。見出しなし） | `meta.company` / `meta.case_id` / `meta.case_type` / `meta.dossier_tier` / `meta.quality_mode` / `meta.round_no` / `meta.industry_name` / `meta.generated_at` / `s1.company_name` | 常に表示（`s1` が null でも `meta` だけで描ける） | 見出し＋チップ列 |
| SEC-02 | exec | エグゼクティブサマリ | `s1.business_summary` / `s1.strategy_outlook.market_context` / `s2.risks[]`（`risk_no` `risk_name` `impact_score` `frequency_score`） / `s3.stories[]`（`story_no` `headline` `pitch` `target_risk_nos`） | 常に表示。`s3` が null のときはテーマ3本を省き「提案ストーリーは未生成です」の1行 | 文章中心（§3.1） |
| SEC-03 | profile | 企業理解 | `s1.business_summary` / `s1.main_products[]` / `s1.processes[]` / `s1.locations[]`（`name` `type` `address` `hazard_note` `notes`） / `s1.supply_chain.key_materials[]` `s1.supply_chain.notes` / `s1.customers.segments[]` `s1.customers.channels[]` / `s1.workforce_notes` / `s1.management_notes` / `s1.strategy_outlook.mvv` `aspirations[]` `market_context` | `s1` が null なら**セクションごと非表示** | 定義リスト＋拠点表 |
| SEC-04 | sufficiency | 入力の充足度と要確認事項 | `s1.input_quality.coverage[]`（`aspect` `status`） / `s1.input_quality.overall` / `s1.input_quality.advice` / `s1.missing_info[]`（`item` `why_needed`） | `s1` が null なら非表示 | 14観点バッジ＋表 |
| SEC-05 | riskuniv | リスクユニバース10分類 | `s2.risks[].category`（19章§3の日本語ラベルへ変換） / `s2.risks[].risk_no` | `s2` が null なら非表示 | 10分類の件数バー（§3.2） |
| SEC-06 | riskmap | 2軸リスクマップ（影響×頻度 5×5） | `s2.risks[]`（`risk_no` `risk_name` `impact_score` `frequency_score` `insurability.transferability`） | `s2` が null なら非表示。`risks` が0件なら「該当なし」の空マップを描く | 5×5マトリクス（§3.3） |
| SEC-07 | risks | リスク一覧 | `s2.risks[]` の全項目（`risk_no` `category` `risk_name` `scenario` `status` `frequency` `impact` `frequency_score` `impact_score` `evidence.quote` `evidence.source` `insurability.transferability` `insurability.line_note` `insurability.control_note` `loss_scale_note` `check_points[]` `preventions[].measure` `preventions[].related_menu_id`） | `s2` が null なら非表示 | 表（横スクロール可） |
| SEC-08 | coverage | 保険カバレッジ表 | `s1.current_coverage[]`（`line_name` `coverage_summary` `limit_note` `special_note`） / `s2.gaps[]`（`gap_no` `gap_type` `target` `description` `risk_evidence` `coverage_evidence`） / `s2.risks[].insurability.transferability` `line_note` `control_note` | `current_coverage` が0件（新規案件）なら「新規案件のため現契約なし。以下は必要補償の見立て」の注記を出して `gaps` 側の表のみ描く。両方0件なら非表示 | 2枚組の表 |
| SEC-09 | newrisk | ニューリスク（新たに立った仮説） | `s2.risks[]` のうち `status` が `new` のもの（`risk_no` `risk_name` `scenario` `category`） | 0件のときは「初回ラウンドでは該当なし（第2ラウンド以降に表示されます）」の1行を出す（**非表示にしない**。FR-35のラウンド設計を読み手に示すため） | カード |
| SEC-10 | story | 提案ストーリー（当社にできること） | `s3.stories[]` の全項目（`story_no` `proposal_kind` `headline` `hook_question` `target_risk_nos[]` `target_gap_nos[]` `menu_ids[]` `line_ids[]` `scheme_id` `pitch` `similar_case_id` `expected_objection` `objection_response`）。`target_risk_nos` は `s2.risks[].risk_no` を、`target_gap_nos` は `s2.gaps[].gap_no` を引いて名称に解決する | `s3` が null なら非表示 | カード3枚 |
| SEC-11 | prevent | 未然防止メニュー | `s2.risks[]`（`risk_no` `risk_name` `preventions[].measure` `preventions[].related_menu_id`） | `preventions` が全リスクで0件なら非表示 | 表 |
| SEC-12 | limit | 当社にできないこと・提案を控えること | ①`s2.risks[]` のうち `insurability.transferability` が `hard`（`risk_no` `risk_name` `insurability.control_note`） ②`s3.unmatched_risks[]`（`risk_no` `risk_name` `why_unmatched`） ③`s3.do_not_propose[]`（`topic` `reason`） | 3ブロックとも0件なら「該当なし」の1行を出す（**非表示にしない**。10章FR-37「できないことを正直に書く」がこのセクションの存在理由であるため） | 3ブロック |
| SEC-13 | hearing | ヒアリング事項 | `s3.stories[].hook_question` / `s2.open_questions[]` / `s1.missing_info[]`（`item` `why_needed`） / `s2.risks[].check_points[]` | 4系統すべて0件なら非表示 | 番号付きリスト（§3.4） |
| SEC-14 | source | 出典と根拠 | `s2.risks[]`（`risk_no` `evidence.quote` `evidence.source`） | `s2` が null なら非表示 | 表 |
| SEC-15 | disclaimer | 免責とご確認事項 | `meta.company` / `meta.generated_at` / `meta.app_version` / `meta.case_id` ＋ §3.5の固定文 | **常に表示（非表示にできない唯一のセクション）** | フッタ |

- **S4は読まない**。本レポートは `GenerateHtmlReport` の契約どおり S1・S2・S3 だけから描く（14章§6）。S4の `hearing_questions` はヒアリングシート（13章§2.16・`modExportHearing`）の入力であり、S4未実行でもレポートが出せる状態を保つため本章では参照しない。
- enum値は必ず19章§3・15章§0の日本語ラベルへ変換して表示する。生の英字enumを画面に出さない。
- 「空のときの挙動」が「非表示」のセクションは、目次（§3.6）からも同時に落とす。

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

1. `本資料はAI支援により作成した骨子を人が確認・編集したものです。`（16章NFR-S5の必須表記）
2. `記載のリスクは公開情報と当社担当者の見立てに基づく仮説であり、引受可否・保険料・幹事構成を確約するものではありません。`
3. `保険料の試算は本資料の対象外です（要見積）。`（10章FR-43）
4. `{meta.company} 御中 / 案件ID {meta.case_id} / 作成 {meta.generated_at} / リスク提案ナビ v{meta.app_version}`

### 3.6 目次

SEC-01 の直後に、表示対象となったセクションの見出しをページ内リンク（`#sec-<slug>`）で並べる。目次は登録表から自動生成し、手で並びを持たない。印刷時は目次を出す（紙でも構成が追えるようにする）。

---

## 4. セクション登録表と、図表テンプレートを1つ追加する手順

### 4.1 描画の場所

DATAはページ内のJavaScriptが `JSON.parse` で受け取り、**セクションの描画はブラウザ側のJSが行う**（10章FR-37「データはJS配列としてテンプレートに埋め込む」）。VBAは値ごとのHTML断片を組み立てない。この分担により、図表の追加・並べ替え・見た目の変更が**VBAのロジックに触れずに済む**。

- VBAが静的HTMLとして書き出すのは次の3つだけで、いずれも `modUtilText.HtmlSafe` を通す: (a) `<title>` (b) SEC-01 表紙の会社名・案件ID・生成日時 (c) `<noscript>` の案内文。
- **JS側の描画は `document.createElement` と `textContent` への代入のみで行う。`innerHTML` / `insertAdjacentHTML` / `document.write` / `outerHTML` への代入を禁止する。** 属性は `setAttribute` で与える。これによりDATA由来の文字列がマークアップとして解釈される経路が構造的に存在しなくなる（17章 T-46 の出荷前検問で、テンプレ文字列中にこれらの語が出現しないことを grep 検査する）。
- `<noscript>` には「このレポートの表示にはJavaScriptが必要です。ファイルをローカルに保存してブラウザで開いてください」と、SEC-15の免責文（§3.5の1行目）を静的HTMLで書く。スクリプトが動かない環境でもAI利用の明示だけは必ず読めるようにする。

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
s = s & " {id:'SEC-06',slug:'riskmap',  title:'2軸リスクマップ',    need:['s2'],     empty:'hide',  render:renderRiskMap},"    & vbLf
s = s & " {id:'SEC-09',slug:'newrisk',  title:'ニューリスク',       need:['s2'],     empty:'note',  note:'初回ラウンドでは該当なし（第2ラウンド以降に表示されます）',render:renderNewRisk}," & vbLf
s = s & "];" & vbLf
```

登録行のキーは**この7つに固定**する。増やさない。

| キー | 型 | 必須 | 意味 |
|---|---|---|---|
| `id` | 文字列 | ○ | `SEC-nn`。§3の表と一致させる。改番禁止 |
| `slug` | 文字列 | ○ | `<section id="sec-<slug>">` と目次アンカーになる。英小文字とハイフンのみ |
| `title` | 文字列 | ○ | 既定の見出し。空文字は見出しを出さない（SEC-01のみ） |
| `need` | 配列 | ○ | 描画に必要なDATAのトップキー（`meta` / `s1` / `s2` / `s3`） |
| `empty` | 文字列 | ○ | **`need` のいずれかが `null`、または描画対象が0件**のときの挙動。`always`=それでも描く（§3で常時表示と決めたもの） / `hide`=描かずに目次からも落とす / `note`=見出しと `note` の1行だけを描く |
| `note` | 文字列 | `empty:'note'` のときのみ | 0件時に出す1行の本文 |
| `render` | 関数名 | ○ | (a)の連結行で取り込んだ描画関数。引数は `(DATA, sectionEl)` の2つに固定し、戻り値を持たない |

`empty` の値は§3の表の「空のときの挙動」列と1対1で対応させる（本表と§3が食い違ったら§3が正）。

### 4.3 図表テンプレートを1つ追加する手順（変更は2箇所で完結する）

例: 「拠点ハザードの一覧図を足したい」という要望を受けた場合。

1. **テンプレ関数を1本追加する（1箇所目）**: 空きのあるテンプレモジュール（§4.4の分割規約に従い、超過していれば新しい `modHtmlTemplateN`）に `Public Function SecHazardJs() As String` を追加し、`function renderHazard(DATA, el){...}` を返す。既存の描画関数・CSS・他セクションには一切触らない。
2. **登録表に2行足す（2箇所目）**: `modHtmlTemplate1.SectionsJs()` に、(a) `s = s & modHtmlTemplateN.SecHazardJs()` の連結行と (b) `{id:'SEC-16',slug:'hazard',title:'拠点ハザード',need:['s1'],empty:'hide',render:renderHazard},` の登録行を、出したい位置に挿入する。
3. **§3の表に1行足す**（本章の更新）。IDは `SEC-16`。既存IDは動かさない。

`modExportHtml`・`modHtmlTheme`・他のテンプレ関数・CSSは変更しない。**新しい配色が必要な場合でも新しいCSS変数を足さず、§5.1の閉じた一覧から選ぶ**（一覧を増やすとテーマ側の全定義に追随が必要になり、テーマ差替が1モジュールで閉じなくなるため）。一覧の拡張が本当に必要なときは本章§5.1の改訂として扱い、`modHtmlTheme` の全テーマを同時に更新する。

セクションの**削除**は登録行の削除（および§3の表からの削除）だけで行い、テンプレ関数は残してよい（呼ばれなくなるだけ）。セクションの**並べ替え**は登録配列の行順の入れ替えだけで行う。

### 4.4 テンプレモジュールの分割規約（1モジュール30,000字契約）

- 1モジュールが**25,000字**を超えたら次番のモジュールへ切り出す（30,000字の契約に対して余白を持たせる。17章 T-35 のDoDで文字数を検査する）。
- 分割は**関数単位**で行い、関数名は変えずに移動だけする。同名の `Public Function` を2つ以上のモジュールに置かない。
- 既定の割り当て:

| モジュール | 持つもの |
|---|---|
| `modHtmlTemplate1` | `BuildDocument`（全体組立）／`HeadHtml`（`<meta charset>`・`<title>`・共通CSS・テーマCSSの差込口）／`BodyShellHtml`（骨格と `<noscript>`）／**`SectionsJs`（§4.2のセクション登録表。編集が最も多い1関数）**／`RuntimeJs`（目次生成・登録配列の走査・`need`/`empty` の判定・共通の描画ヘルパ） |
| `modHtmlTemplate2` | SEC-01 cover ／ SEC-02 exec ／ SEC-03 profile ／ SEC-04 sufficiency |
| `modHtmlTemplate3` | SEC-05 riskuniv ／ SEC-06 riskmap ／ SEC-07 risks ／ SEC-08 coverage |
| `modHtmlTemplate4` | SEC-09 newrisk ／ SEC-10 story ／ SEC-11 prevent ／ SEC-12 limit ／ SEC-13 hearing ／ SEC-14 source ／ SEC-15 disclaimer |

- 共通CSSは `modHtmlTemplate1.HeadHtml` に一元化し、セクション別のテンプレ関数に `<style>` を書かない（CSSが散ると見た目のフィードバックを1箇所で吸収できなくなる）。セクション固有のスタイルはクラス名を `sec-<slug>-*` の接頭辞で共通CSSに置く。
- 文字列の組み立ては15章と同じ `s = s & "..." & vbLf` 方式とする（`Const` は1論理行1,023字・行継続25本の制約に当たるため使わない。14章§7と同じ理由）。

---

## 5. テーマとCSS変数、エスケープ、文字コード

### 5.1 テーマが定義してよいCSS変数の閉じた一覧（28個。これ以外を定義しない・これ以外を参照しない）

`modHtmlTheme.ThemeCss(themeName)` が返すのは `:root{ ... }` **1ブロックだけ**であり、その中身は下表の28変数の宣言だけである。既定値はライト（白地）印刷前提であり、`docs/demo/確認用v2/` および `docs/presentation/` の既存HTMLの `:root` を出発点にしている。

**寸法・書体（6）**

| 変数 | standard の既定値 | 用途 |
|---|---|---|
| `--page-width` | `900px` | 本文カラムの最大幅 |
| `--page-pad` | `24px` | 本文カラムの左右余白 |
| `--font-sans` | `"Noto Sans JP","Yu Gothic","Hiragino Kaku Gothic ProN","Meiryo",sans-serif` | 本文。**Webフォントを読み込まない**（外部参照禁止）。端末に無い書体は後続へフォールバックし、最後は総称名で必ず解決する |
| `--font-serif` | `"Shippori Mincho","Yu Mincho","Hiragino Mincho ProN",serif` | 見出し（h1・h2） |
| `--font-size` | `14.5px` | 本文の基準サイズ |
| `--line-height` | `1.85` | 本文の行間 |

**地色と文字（5）**

| 変数 | standard の既定値 | 用途 |
|---|---|---|
| `--paper` | `#FFFFFF` | 紙面の地色 |
| `--ink` | `#24303E` | 本文の文字色 |
| `--sub` | `#5A6B7E` | 補助文・注記の文字色 |
| `--mist` | `#EFF3F8` | 表ヘッダ・囲みの薄い面色 |
| `--line` | `#D8E0E9` | 罫線・区切り線 |

**強調（4）**

| 変数 | standard の既定値 | 用途 |
|---|---|---|
| `--ai` | `#2B5B8F` | 主色。見出し・キッカー・ヘッダ下線 |
| `--kaki` | `#B4552D` | 注意・できないこと・不足の強調 |
| `--matsu` | `#2F7A54` | 肯定・提案・充足の強調 |
| `--deep` | `#7A5C9E` | 入念モード由来の記述・補足ボックス |

**注意面（2）**

| 変数 | standard の既定値 | 用途 |
|---|---|---|
| `--warn` | `#FBF3E6` | 注意ボックスの背景 |
| `--warn-line` | `#E8D5B5` | 注意ボックスの罫線 |

**リスクマップの帯（5。§3.3の帯番号1から5に対応）**

| 変数 | standard の既定値 |
|---|---|
| `--heat-1` | `#E8F5EE` |
| `--heat-2` | `#F2F7E9` |
| `--heat-3` | `#FFF4D6` |
| `--heat-4` | `#FFE9D3` |
| `--heat-5` | `#FBDCD9` |

**移転可能性の3値（3。`insurability.transferability` に対応）**

| 変数 | standard の既定値 | 対応enum |
|---|---|---|
| `--tr-cover` | `#2F7A54` | `cover`（比較的移転しやすい） |
| `--tr-partial` | `#B08A2E` | `partial`（条件付き・部分的） |
| `--tr-hard` | `#B4552D` | `hard`（保険化困難） |

**入力充足度の3値（3。`input_quality.coverage[].status` に対応）**

| 変数 | standard の既定値 | 対応enum |
|---|---|---|
| `--iq-ok` | `#2F7A54` | `ok` |
| `--iq-partial` | `#B08A2E` | `partial` |
| `--iq-missing` | `#A9B4C0` | `missing` |

**規約**

- 共通CSS（`modHtmlTemplate1.HeadHtml`）は**色・書体・本文幅をリテラルで書かない**。必ず `var(--xxx)` を通す。唯一の例外は、色付きバッジ・見出し帯の上に載せる文字色 `#fff` であり、これ以外の生の色指定を書かない。
- `@page` の余白はCSS変数で解決されないため、テーマ変数にせず共通CSSにリテラルで書く（§6）。テーマから紙面余白は変えられない、と割り切る。
- 機械検査（17章 T-35 のDoD）: (1) `ThemeCss` の戻り値が `:root{` で始まり `}` で終わり、内側が `--` で始まる宣言のみであること (2) 全テーマが上表28変数を**過不足なく**定義していること (3) 共通CSS中に `var(--` を伴わない色指定（`#` に続く16進6桁・3桁）が `#fff` 以外に出現しないこと。

### 5.2 テーマ差替の単位

- **差し替えの単位は `modHtmlTheme` の1モジュールのみ**。テーマを増やす作業は「`ThemeCss` の `Select Case` に分岐を1本足し、28変数を書く」で完結し、`modExportHtml` にも `modHtmlTemplate1..n` にも触れない。
- 契約:

```vb
' === app: modHtmlTheme（純文字列・R4）===
Public Function ThemeNames() As String     ' ";"区切り。先頭が既定テーマ。例: "standard;mono"
Public Function ThemeCss(ByVal themeName As String) As String
' 戻り値は ":root{ ... }" の1ブロックのみ。§5.1の28変数を過不足なく宣言する。
' 未知のテーマ名は standard へフォールバックし、run_log の detail に
' theme_fallback=<要求されたテーマ名> を記録する（黙って既定に戻さない）
```

- 初期テーマは2本。`standard`（§5.1の既定値。カラー画面・カラー印刷向け）と `mono`（白黒印刷・FAX配布向け。強調4色を `--ink` と `--sub` の濃淡に、heat 5段を白から薄灰の5段に置き換える）。`mono` でも§3.3の帯番号があるため重篤度は読める。
- config `html_theme`（13章§2.3・既定 `standard`）の値をそのまま `ThemeCss` に渡す。テーマ名の一覧をconfigに書かない（`ThemeNames` が唯一の一覧）。

### 5.3 エスケープと文字コード（16章 E-47・NFR-S7の③が正。本節はその適用手順）

**(1) DATAの埋め込み**: 次の1行の形でのみ埋める。JSONをJSのオブジェクトリテラルとして直書きすることを禁止する。

```
<script>var DATA=JSON.parse("<<JsStringSafe(dataJson) の結果>>");</script>
```

`modUtilText.JsStringSafe` の適用順は次のとおりで、**この順序を守る**（逆順にすると二重エスケープになる）。

1. `\` を `\\` に、`"` を `\"` に置換する（JSON本文をJSの二重引用符リテラルへ入れるため必須）
2. `</` を `<\/` に置換する（本文中の `</script>` でスクリプトブロックが閉じるのを防ぐ）
3. 行区切り文字 U+2028 を `\u2028`、段落区切り文字 U+2029 を `\u2029` に置換する
4. その他の制御文字（U+0000 から U+001F）を `\u00XX` に置換する

**(2) HTML本文への差し込み**: §4.1の3箇所（`<title>` ・表紙の会社名/案件ID/生成日時・`<noscript>`）は `modUtilText.HtmlSafe`（`& < > " '` のエンティティ化）を通す。JS側の描画は `textContent` と `setAttribute` のみを使うためエスケープ不要であり、逆に `innerHTML` 系を使わないことがエスケープ規約そのものである（§4.1）。

**(3) 文字コード**: 書き出しは `ADODB.Stream`（`Charset = "utf-8"`・**BOMあり**・`SaveToFile` は上書き指定）で行う。VBAの `Open ... For Output` / `Print #` はCP932で書き、非CP932文字（絵文字・環境依存字・一部の丸数字）が `?` へ落ちるため**使わない**。`<head>` の**最初の要素**として `<meta charset="utf-8">` を置く（BOMを見ないブラウザ設定でも文字化けしないようにするための二重化）。

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
