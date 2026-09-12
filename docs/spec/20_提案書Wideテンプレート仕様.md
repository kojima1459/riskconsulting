# 20. 提案書Wideテンプレート仕様 v1.1

**v1.1（W15 Round 2・裁定書39）の変更**: §11「前提ブラウザと社内での確認」を新設（R2-07。提案書は **Chromium 版 Edge** で開く。画面の中の表示枠＝IE11 相当では提案書の CSS が効かない）。§1.1 に**必須キー検査**（R2-12）、§3 に**機械置換の記録**（R2-11）、§4.1 に**描画エラーの印**（R2-12）を足した。対訳表の機械置換は**最長一致・一般語は置換しない・冪等**の3規約になった（R2-02。`docs/design/提案書_wide/対訳表_社内語から顧客語.md` §6）。

**新設**（W15・裁定書38 §1 班C）。正の出どころは `docs/29_髙橋構想v1.0レビューと提案書フォーマット裁定.md` §5（裁定: 顧客向け提案書は Wide 22枚見本に合わせる／AI は提案書JSONだけを出し HTML は VBA が固定テンプレで組む／呼び出しは1回＋修復1回）と、見本一式 `docs/design/提案書_wide/`（`shunkado_proposal_wide_v0.1.src.html`＝22枚の実物、`【LOG】sample_wide_worklog_v0.1.md`＝作業ログ、`talk_script.md`＝発表者ノート、`プロンプト-デザイン.txt`＝設計原則、`対訳表_社内語から顧客語.md`＝用語対訳表）である。

**18章との関係**: 18章（HTMLレポート）は**営業担当者向け**の分析資料、本章（提案書 Wide）は**お客さまへそのままお渡しする**資料である。作法（DATAは1本のJSON／描画はJSの `createElement`＋`textContent` だけ／自己完結1ファイル／`<` は `<`）は18章§4.1・§5.3をそのまま適用し、**違うところだけ**を本章が定める。違うのは次の5点である。

| # | 18章（レポート） | 本章（提案書） |
|---|---|---|
| 1 | 読み手は営業担当者 | 読み手はお客さまの経営層 |
| 2 | セクションは `empty:'hide'` で消える | **枚数は22枚で固定**。データが無い枚も1行を置いて必ず描く（頁が抜けると構成が崩れる） |
| 3 | A4縦・1カラムの文書 | **16:9のスライド22枚**。A4横で1枚=1ページ |
| 4 | AI利用の明示を必ず出す | **AI表示は出さない**。そのかわり**担当者の確認が無いと出力できない**（§8） |
| 5 | 内部の値（tier・mode）も `meta` に持つ | 内部の値を**構造的に持たない**（§3） |

---

## 1. 目的と責務分担

| モジュール | 層 | 責務 | 禁止事項 |
|---|---|---|---|
| `modPipeline5` | app | S5 の実行（組立→1呼び出し→検証→修復1回→保存） | プロンプト本文・スキーマ本文を持たない |
| `modPromptsS5` | app（純文字列） | 15章§5.6 の system / user テンプレートの**写し** | 置換・分岐を書かない（組立層は `modPipeline5.AsmS5User`） |
| `modSchemas2` | app（純文字列） | Schema-S5 | 同上 |
| `modValidate4` | app | CheckS5（15章§5.6 の13件）・対訳表・機械置換 | 案件データ・configに触れない |
| `modExportProposal` | app | S2・S5 の解決、匿名化復元、PII走査、DATA組立、テンプレ組立、ファイル書出 | **HTML本文・CSS・JSの文字列をここに書かない** |
| `modProposalHtml1..4` | app（純文字列） | HTML骨格・CSS・スライド描画スクリプト・スライド登録表 | Excelトークン・案件データの参照（R4） |

- 生成に**LLMを使わない**。LLM が返すのは提案書JSON（Schema-S5）だけであり、HTML・CSS・JS は人が書いた固定文字列である（docs/29 §5.1。髙橋さん作業ログ §10.2 も同じ結論）。
- 出力は**自己完結HTML 1ファイル**。外部CSS・外部JS・外部フォント・外部画像・CDNを一切参照しない。見本が持っていた画像6点（header_bar / title_box / company_name.svg / logo_back / tagline / watermark。data URI 合計約300KB）は**1点も持たず**、ヘッダー帯は `linear-gradient`、タイトル枠は CSS の枠で再現し、ロゴと透かしは**外す**（docs/29 §5.2 の軽量化裁定 Q-4 に対する当方の決着＝画像合計 0KB）。
- 書体は社内PCに必ずある3書体だけを指定する: **メイリオ / 游ゴシック / MS Pゴシック**（`--font-sans`。§5.1）。Webフォントを読まない。

### 1.1 生成手順（`modExportProposal.GenerateProposalHtml`。この順序で行う）

| 順 | 手順 | 失敗したら |
|---|---|---|
| ① | **`reviewedBy` が空なら生成しない**。`内容を確認してから出力してください。` を返して中止する | エラーコードは立てない（未確認は障害ではない）。§8 |
| ② | `modCaseStore.ResolveStepJson(caseId, 2)` で S2、`s5_edited > s5_json` で S5 を解決する | **S5 が空なら生成しない**。「先にお客さま向け提案書の作成を実行してください。」と案内して中止 |
| ②' | **S5 の必須キー検査**（`modExportProposal.MissingProposalKeys`。Schema-S5 の最外 `required` 16キーが最外オブジェクト直下にあるか）| **1つでも欠けていたら生成しない**。E0502 を立て、欠落キー名を添えて中止する。`s5_edited` を人が直した場合はスキーマ強制が掛からないため、ここが唯一の関門である（W15 Round2 R2-12。欠けたまま描くと描画関数が例外を投げ、その枚が §4.1 の「次回のお打ち合わせで更新します」に化けて**壊れたページが意図した保留に見える**） |
| ③ | 匿名化の復元（16章 E-31）。`{{COMPANY}}` を実名へ戻す。**必ずエスケープより前** | 復元表が無くてもプレースホルダのまま出す（顧客向け資料に内部警告は出さない） |
| ④ | `modPii` 走査。検知しても**生成はブロックせず**、run_log/err_log に検知種別と箇所だけを記録する（本文は記録しない＝NFR-S3） | 走査自体の失敗は記録のみで続行 |
| ⑤ | DATA（§3）を1本のJSON文字列として組み立てる | 組立失敗は E0502 |
| ⑥ | `modProposalHtml1.BuildProposalDocument(dataJson, coverFields)` でHTML全文を組む | 組立失敗は E0502 |
| ⑦ | **UTF-8（BOMあり）で書き出す**（`modUtil.WriteUtf8File`）。ファイル名は§9.2 | 書込失敗は E0502 |

- 案件の `status` はこの出力で変化させない（16章 E-48）。`report_path` も書き換えない（あれはHTMLレポートの列である）。

---

## 2. 呼び出しと入力（S5）

- プロンプト・スキーマ・検証の**正は15章§5.6**であり、本章はそれを繰り返さない。
- 呼び出しは**1回（フラッグシップ）＋修復1回**（docs/29 §5.1）。実装は `modPipeline5.RunStep5`。
- 入力は S1＋S2＋S3（18章 DATA と同じ解決順）と、髙橋さん作業ログ §6.3 の「promptに書くこと」9項目、用語対訳表（`docs/design/提案書_wide/対訳表_社内語から顧客語.md`）。いずれも15章§5.6 の system 本文に入っている。
- **件数・スコア・順位・IDといった数値は AI に数えさせない**。VBA が S2/S3 から数えて `{{statsText}}` として渡し（`modExportProposal.StatsText`）、ページには DATA の `stats` を描く。Schema-S5 に `stats` が無いのはこのためである。

---

## 3. DATA（ページに埋め込む唯一のデータ構造）

```json
{
  "meta": {
    "company": "株式会社浜松スイーツファクトリー",
    "title": "経営リスク分析と保険活用のご提案",
    "subtitle": "...", "date": "2026/09/12",
    "app_version": "2.4.0", "reviewed_by": "浜松支店 山田",
    "reviewed_at": "2026/09/12 10:05:00"
  },
  "stats": {"risk_total": 8, "easy": 3, "design": 3, "non_ins": 2,
            "ideas": 8, "ideas_priority": 4},
  "categories": [{"label": "戦略・市場", "count": 1}],
  "risks": [{"risk_no": 3, "category_label": "施設・自然災害・事業継続",
             "risk_name": "...", "impact": 4, "frequency": 2,
             "rank_label": "中", "class_label": "保険で備えやすい",
             "line_note": "...", "check_note": "...", "control_note": "..."}],
  "p": { "...": "Schema-S5 の出力をキー名を変えずそのまま" }
}
```

- `meta` は `modExportProposal.BuildProposalMetaJson` が組む。`title` / `subtitle` は S5 の同名キー、`date` は生成時刻の日付部、`reviewed_by` / `reviewed_at` は §8 の確認者。
- `stats` は **VBA が数えた実数**。`risk_total` / `easy` / `design` / `non_ins` は S2 の `risks[].insurability.transferability`（`cover` / `partial` / `hard`）から、`ideas` / `ideas_priority` は S5 の `ideas[]` と `ideas[].priority` から数える。**AI の書いた数字を使わない**。
- `risks` は S2 の `risks[]` を `impact_score × frequency_score` の降順（同点は `risk_no` の昇順）に並べ、**顧客向けに選んだ10列だけ**を写したもの。`rank_label` は積が 16以上=最優先 / 12以上=高 / 6以上=中 / それ未満=低。`class_label` は `cover`=保険で備えやすい / `partial`=補償条件の設計が必要 / `hard`=保険以外の対策が中心。`line_note` / `check_note`（`gap_note`）/ `control_note` は S2 の `insurability` の同名キー。
- **S2 由来の文字列は、DATAへ入れる前に `modValidate4.SoftenTaboo` で対訳表の機械置換を通す**。S2 は営業向けの語彙で書かれており、LLM の言い換えを経ずにお客さまの目に触れる唯一の経路であるため。置換は次の3規約に従う（W15 Round2 R2-02）:
  1. **最長一致**。対訳表を社内語の長さの降順に並べ、本文を左から1回走査してその位置で最も長く一致する語だけを置き換える。宣言順の `Replace` は「未付保」を「未＋保険のご加入」に、「付保ギャップ」を「保険のご加入ギャップ」に、「リスク移転可能性」を「リスク保険で備える可能性」に、「座組パターン」を「ご提案の構成パターン」に壊す。
  2. **一般語は置換しない**。「移転」「保有」「抜け」は社内語であると同時に日常語であり（本社を移転する／現金を保有する）、機械が潰すと「本社を保険で備えるする」になる。対訳表 §6 の印を持つ語は置換対象から外し、**V-S5-12 の警告にだけ出す**。置換後も V-S5-12 が一般語だけで残るときは `modPipeline5.SoftenOrFail` が run_log に記録して**続行する**（語1つで顧客向け提案書を止めない。docs/29 §5.3）。
  3. **冪等**。置換結果をもう一度 `SoftenTaboo` に通しても1件も変わらない（顧客語が別の社内語を含む対があるため、変化が無くなるまで通す）。
- **機械置換したことは必ず記録する**（docs/29 §5.3「黙って直さない」。W15 Round2 R2-11）。`modExportProposal.BuildProposalDataEx` が置換した箇所数を返し、`GenerateProposalHtml` が usage_log へ `proposal_taboo_softened` / `n=<件数>` を**1行**書く（0件でも書く＝「1件も直していない」ことを示す）。
- `categories` は19章§3の10分類を**顧客語**に置き換えたラベル（`施設・自然災害・BCP` → `施設・自然災害・事業継続` 等。値源は `modExportProposal` の `EP_CAT_LABELS`）と件数。
- **DATAに入れないもの（構造的に持たない）**: `dossier_tier` / `quality_mode` / `case_type` / `round_no` / `s4_variant` / `status` / `report_path` などの内部パラメータ、S3 の `talk_script`（話法と `taboo`＝18章 SEC-18 相当）、S1 の `field_insights`（現場メモの原文）、`warnings`（当方の内部警告）、入力貼付テキストの原文、run_log・err_log、ナレッジ本文。**「入れてから隠す」のではなく、置く枝を持たない**。

---

## 4. スライド登録表（22枚）

### 4.1 マスター4種

| master | 使う枚 | 器（`modProposalHtml1.shell` が描く） |
|---|---|---|
| `cover` | 1 | 本文だけ（宛名・題・副題・日付・担当は描画関数が置く） |
| `section` | 3 / 7 / 10 / 12 / 15 | セクションラベル（`label`）＋題（`title`）＋罫線 |
| `content` | 2 / 4 / 5 / 6 / 8 / 9 / 11 / 13 / 14 / 16 / 17 / 18 / 19 / 20 / 21 | ヘッダー帯＋帯の中の題（`title`）＋本文領域 |
| `back` | 22 | 本文だけ |

- `content` と `section` には**フッター罫と `N / 22` のページ番号**が付く。`cover` と `back` には付けない。
- **データが無い枚の扱い**: `content` の本文が1つも描かれなかったときだけ、`この項目は次回のお打ち合わせで更新します。` の1行を置く（値源は `modProposalHtml1` の `TODO()` 1箇所）。**枚を消さない**。
- **描画が失敗した枚の扱い**（W15 Round2 R2-12）: 描画関数が例外を投げた枚も同じ1行に置き換わるが、それだけでは「わざと保留した項目」と見分けが付かない。`modProposalHtml1` の `run()` は `catch` で **`section` に `data-render-error="1"` を付ける**。この印は**既定では見た目を変えない**（お客さまの画面は1ピクセルも変わらない）。社内で確認するときだけ URL に `?debug=1` を付けると `body.debug` が付き、その枚に赤枠（`outline`）が出る（§11）。**上流の防壁は §1.1②' の必須キー検査**であり、この印は最後の保険である。

### 4.2 登録表（本表が正）

| No | slug | master | 見出し（既定） | 読むJSONパス | 空のときの扱い |
|---|---|---|---|---|---|
| 1 | cover | cover | （表紙。見出しなし） | `meta.company` / `p.title` / `p.subtitle` / `meta.date` / `meta.reviewed_by` | 常に描く（`meta` だけで描ける） |
| 2 | summary | content | ご提案の要旨 | `stats` / `p.themes[]`（`name` `headline` `body`） / `p.structure[]` | 指標帯だけでも描く |
| 3 | sec1 | section | 貴社の事業構造とリスクの全体像 | （登録表の `label`＝`Section 1` と `title`） | 常に描く |
| 4 | business | content | 事業構造：リスクを生む5つの事業領域 | `p.business.facts[]` / `p.business.areas[]` / `p.business.factors[]` | 1行を置く |
| 5 | universe | content | リスクの全体像：10のカテゴリーで捉える | `p.categories_note`（`heavy` `meaning`） / `categories[]` | 1行を置く |
| 6 | riskmap | content | リスクマップ：リスクの優先順位 | `p.headline.riskmap` / `risks[]`（`risk_no` `risk_name` `impact` `frequency`） | 5×5の空マップを描く |
| 7 | sec2 | section | 既存の保険で備えられるリスクと、その条件 | （`Section 2`） | 常に描く |
| 8 | classes | content | 保険での備えやすさ：3つの分類 | `p.headline.classes` / `stats` / `risks[].class_label` | 指標帯だけでも描く |
| 9 | priority | content | 優先リスク：対応する保険と確認事項 | `p.headline.priority` / `risks[]` の上位8件（`rank_label` `line_note` `check_note` `control_note`） | 1行を置く |
| 10 | sec3 | section | 保険だけでは備えにくいリスク | （`Section 3`） | 常に描く |
| 11 | hard | content | 保険だけでは備えにくいリスク | `p.headline.hard` / `p.hard_risks[]`（`risk_no` `horizon` `background` `approach`） | 1行を置く |
| 12 | sec4 | section | 成長を後押しする保険の活用 | （`Section 4`） | 常に描く |
| 13 | ideas | content | 成長支援：保険活用アイデアと優先候補 | `p.headline.ideas` / `p.ideas[]`（`title` `aim` `effect` `difficulty` `priority`） | 1行を置く |
| 14 | four | content | 優先4案：狙い・仕組み・想定する保険 | `p.headline.four` / `p.four[]`（`title` `aim` `mechanism` `insurance`） | 1行を置く |
| 15 | sec5 | section | ご提案の全体像と進め方 | （`Section 5`） | 常に描く |
| 16 | themes | content | 提案テーマ：経営課題ごとに補償と指標を束ねる | `p.headline.themes` / `p.theme_table[]`（`theme` `issues` `insurance` `kpi`） | 1行を置く |
| 17 | steps | content | 進め方：4つのステップ | `p.headline.steps` / `p.steps[]`（`title` `who` `desc` `ref`） | 1行を置く |
| 18 | decide | content | 本日ご判断いただきたいこと | `p.headline.decide` / `p.decisions[]`（`title` `options` `note`） | 1行を置く |
| 19 | share | content | 次のステップ：ご共有いただきたい事項 | `p.headline.share` / `p.share_items[]`（`text` `group` `priority`） | 1行を置く |
| 20 | appendix | content | 巻末資料：リスク一覧 | `p.headline.appendix` / `risks[]` の全件 | 1行を置く |
| 21 | premise | content | 本資料の前提と参照した公開情報 | `p.premise` / §8の免責固定文 / `meta.date` / `meta.reviewed_by` | 免責だけでも描く |
| 22 | back | back | （裏表紙。見出しなし） | §8の免責固定文 | 常に描く |

**登録行のキーはこの6つに固定する。増やさない。**

| キー | 型 | 意味 |
|---|---|---|
| `no` | 数値 | 1..22。ページ番号であり `p.notes[].slide_no` と対応する |
| `slug` | 文字列 | `<section id="sl-<slug>">` になる。英小文字とハイフンのみ |
| `master` | 文字列 | `cover` / `section` / `content` / `back` の4値 |
| `title` | 文字列 | 本表の「見出し（既定）」列を**逐語で写す**。空文字は見出しを出さない |
| `label` | 文字列 | セクション扉の `Section N`。`master='section'` のときだけ意味を持つ |
| `render` | 関数名 | 描画関数。引数は `(DATA, bodyEl, entry)` に固定し戻り値を持たない |

実体は `modProposalHtml1.SlidesJs()` **1関数**の中にあり、(a) 描画関数の連結行 と (b) 登録配列 の2行1組で1枚を表す（18章§4.2 と同じ形）。

### 4.3 スライドを1枚差し替える手順（変更は2箇所で完結する）

1. `modProposalHtml3` または `modProposalHtml4`（§4.4の字数規約に従い、超過していれば新しい `modProposalHtmlN`）に `Public Function SlXxxJs() As String` を追加し、`function renderXxx(DATA,el,d){...}` を返す。
2. `modProposalHtml1.SlidesJs()` の (a) 連結行と (b) 登録行を差し替える。
3. §4.2 の表を直す（**本表が正**。実装と食い違ったら本表に合わせる）。

### 4.4 モジュール分割規約

- 18章§4.4 の**25,000字**規約に倣う（1モジュール30,000字の契約に対して余白を持たせる）。`tools/vba_lint.py` の `TEMPLATE_MAX_CHARS` は `modHtmlTemplate*` 専用のため `modProposalHtml*` には掛からない。**本章の規約として守る**（超過に気づく口は `vba_lint` の28,000字WARN帯である）。
- 既定の割り当て:

| モジュール | 持つもの |
|---|---|
| `modProposalHtml1` | `BuildProposalDocument` / `HeadHtml` / `BodyShellHtml` / **`SlidesJs`（登録表）** / `RuntimeJs`（ヘルパ・走査・マスターの器・発表者ノート・画面の縮小率） |
| `modProposalHtml2` | `RootCss` / `FormatCss` / `ContentCss` |
| `modProposalHtml3` | 1 cover / 2 summary / 3・7・10・12・15 section / 4 business / 5 universe / 6 riskmap / 8 classes / 9 priority / 11 hard |
| `modProposalHtml4` | 13 ideas / 14 four / 16 themes / 17 steps / 18 decide / 19 share / 20 appendix / 21 premise / 22 back ／ `DisclaimerText`（§8の固定文） |

---

## 5. CSSとコンポーネント

### 5.1 テンプレが定義してよいCSS変数の閉じた一覧（19個。これ以外を定義しない・これ以外を参照しない）

`--slide-w` / `--slide-h` / `--font-sans` / `--font-base` / `--line-height` / `--ink` / `--sub` / `--paper` / `--line` / `--brand` / `--brand-deep` / `--accent` / `--soft` / `--heat-1` / `--heat-2` / `--heat-3` / `--heat-4` / `--note-bg` / `--pad-x`

- 実体は `modProposalHtml2.RootCss`（`:root{ ... }` の1ブロックだけを返す）。
- `:root` ブロックの**外**に生の色指定を書かない（例外は `#fff` と `rgba()` の半透明だけ）。`tools/render_proposal.py` が機械で見る。
- 提案書に**テーマ差替は無い**（18章の `html_theme` は本章に掛からない）。お客さまへ出す資料の見た目を設定で分岐させない。

### 5.2 本文コンポーネント

`hm`（結論の一文。朱の強調は `em` 1箇所）/ `sub` / `lead` / `muted` / `cols`＋`col` / `card` / `stats`＋`stat` / `table.t`（`tr.pri` は太字）/ `pill` / `bars`＋`bar` / `map`（5×5。`b1`..`b4` の4段）/ `steps`＋`step` / `grid2` / `kv` / `stars`。実体は `modProposalHtml2.ContentCss`。

- **色だけで意味を運ばない**（18章§6③と同じ規律）。リスクマップのセルには番号と名称を文字で置き、優先の印は「優先」の文字で示す。

---

## 6. 印刷（A4横）と画面表示の両立

| # | 規約 | 実装 |
|---|---|---|
| ① | **1枚＝1ページ** | `.frame{page-break-after:always;break-after:page}`（最後の1枚だけ `auto`） |
| ② | 用紙はA4横・余白なし | `@page{size:A4 landscape;margin:0}` をリテラルで書く（CSS変数は `@page` で解決されないため） |
| ③ | 1280×720 のスライドをA4横に収める | 印刷時 `.frame{width:297mm;height:210mm}` とし、`.frame > section.slide{transform:scale(0.877)}`（297mm ≒ 1122.5px ÷ 1280px） |
| ④ | 画面では幅に合わせて縮小する | `.frame{width:min(var(--slide-w),calc(100vw - 48px));aspect-ratio:16/9}` とし、JSが `--s` を `clientWidth/1280` で設定する（`style.setProperty` のみ。`innerHTML` を使わない） |
| ⑤ | 背景色印刷が無効でも読める | `print-color-adjust:exact` を指定するが依存しない。ヘッダー帯が出なくても題は黒文字で読める |
| ⑥ | 発表者ノートは印刷しない | `@media print{aside.note{display:none!important}}` |

---

## 7. 発表者ノート

- `p.notes[]`（22件。`slide_no` / `read` / `ask` / `probe[]` / `follow[]`）を、各スライドの直後に `<aside class="note">` として置く。
- **既定では画面にも出さない**。URLに `?notes=1` を付けたときだけ `body.notes` が付いて表示される。印刷では常に消える（⑥）。
- お客さまへお渡しするファイルにノートが含まれることは避けられない（自己完結1ファイルであるため）。**`?notes=1` を知らなければ見えない**という水準の扱いであり、機微な内容をノートへ書かない運用とあわせる（15章§5.6 の system が「発表者ノートも顧客が読んでよい言葉で書く」ことを求めているのはこのためである）。

---

## 8. 免責と、顧客向けだから入れないもの

1. **免責は次の1文だけ**（値源は `modProposalHtml4.DisclaimerText` の1箇所。21枚目と22枚目が同じ関数から取る）:

   `本資料は、引受・保険料・契約条件を確約するものではありません。`

2. **AI利用の明示は入れない**（社内IT・AI環境 v1.1 §7.3）。理由は「隠す」ためではなく、本資料が**担当者の確認を経た顧客提示物**だからである。その担保が次の3。
3. **`reviewedBy` が空なら生成しない**（§1.1①）。画面（区画④）でもチェックと氏名がそろわないと送信しないが、**判断の正は `modExportProposal` の1箇所**であり、画面を書き換えても抜けられない。
4. **18章 SEC-18 相当（経営層への話し方＝`s3.talk_script` の `flow` と `taboo`）は入れない**。あれは当社の商談の作法であってお客さまへお見せするものではない。`s1.field_insights`（現場メモの原文）も同じ理由で入れない。DATAに置く枝そのものを持たない（§3）。

---

## 9. エスケープ・文字コード・ファイル名

### 9.1 エスケープと文字コード

18章§5.3 をそのまま適用する。**(1)** DATAは `<script>var DATA=JSON.parse("<<JsStringSafe(dataJson)>>");</script>` の形でのみ埋め、`<` を1文字残らず `<` にする。**(2)** HTML本文への静的な差し込みは `<title>` と `<noscript>` の2箇所だけで、`modUtilText.HtmlSafe` を通す。**(3)** 書き出しは UTF-8（BOMあり）。`<head>` の最初の要素として `<meta charset="utf-8">` を置く。

### 9.2 ファイル名（W12-c）

`提案書_<Sanitize(company)>_<yyyymmdd>_v<app_version>.html`

- 組立の実体は `modUtilPath.BuildVersionedFileName`（**HTMLレポートと共有の1本**。レポートは先頭語が `レポート`）。会社名は `modUtilText.SanitizeFileName` を通す。
- 最終パスが240字を超えるときは会社名部を `Fnv1a64Hex`（16桁）へ置換する（13章§2.8 手順5と同じ）。
- 同名が既にあれば `_2` `_3` の連番を探して**新規ファイルとして作る**（過去の出力を消さない。13章§2.8）。

---

## 10. 検問

- `tools/render_proposal.py`（`render_report.py` と同型。gate.py への登録は司令塔が行う）。mock の S5 JSON（`MK-S5`）と S2 JSON（`MK-S2-RNW`）から実物のHTMLを組み、次を機械で見る。
  1. `<meta charset="utf-8">` が先頭付近にある
  2. 登録表に22行あり、`no` / `slug` / `master` / `title` が §4.2 の表と**逐語一致**する。描画関数も定義されている
  3. DOMスタブでページのJSを実際に走らせ、`sl-<slug>` が**22枚**生成され、`data-no` が 1..22 になる
  4. `innerHTML` / `insertAdjacentHTML` / `document.write` / `outerHTML` が1つも出ない
  5. DATAの文字列リテラル内に生の `<` が1文字も無い
  6. §8の免責固定文と §4.1 の「1行」の文言が逐語で入っている
  7. **内部の値がDATAに無い**（`dossier_tier` / `quality_mode` / `case_type` / `round_no` / `s4_variant` / `talk_script` / `field_insights` / `taboo`）
  8. `@page{size:A4 landscape` がある（§6②）
  9. CSS変数が §5.1 の19個と過不足なく一致し、`:root` の外に生の色指定が無い
  10. 外部URL（`http://` / `https://` / `//cdn`）が1つも無い
  11. 対訳表（`docs/design/提案書_wide/対訳表_社内語から顧客語.md`）の全語が15章§5.6 の system 本文に現れる
- 純テスト `src/test/modTestsPure27.bas`（層(a)）。CheckS5 の13件・登録表の枚数・空データで落ちないこと・内部値がDATAに入らないこと・`reviewedBy` 空で生成しないこと・ファイル名規則。W15 Round 2 で次の16本を足した（裁定書39）: 対訳表の最長一致4本（壊す班が実測した「未保険のご加入」「保険のご加入ギャップ」「リスク保険で備える可能性」「ご提案の構成パターン」を逐語で禁じる）・一般語を置換しないこと2本・一般語でも V-S5-12 は出ること・冪等・英字の語境界（CBI を BI で刻まない）・機械置換の件数を呼出側が受け取ること・描画例外の印・`?debug=1` の赤枠・必須キー検査2本・空白類だけの確認者名を認めないこと2本。
- 目視: ヘッドレス Chromium でPDF化して22枚を確認する（髙橋さん作業ログ §6.4 の C 項目。docs/24 §7 の観察項目）。

---

## 11. 前提ブラウザと社内での確認（W15 Round 2・裁定書39 R2-07）

### 11.1 提案書は **Chromium 版 Edge** で開く

- 提案書HTMLを表示・印刷・PDF化するときは、**エクスプローラーからファイルをダブルクリックして Chromium 版 Microsoft Edge（会社PCの既定ブラウザ）で開く**。
- **画面の中の表示枠（`frmNaviHtml` の WebBrowser）で提案書を開かない**。あれは IE11 相当であり、本章が使う次の3つをいずれも解さないため、22枚が崩れる。
  1. CSS カスタムプロパティ（`var(--slide-w)` ほか §5.1 の19変数）→ `.frame` と `section.slide` の幅・高さ・色が全滅する
  2. `display:grid`（§5.2 の `map` 5×5 と `grid2`）→ ヒートマップが縦1列に潰れる
  3. `min()` と `aspect-ratio`（§6④）／JS の `style.setProperty('--s', …)` → 画面の縮小率が効かない
- これは**欠陥ではなく運用の前提**である。HTML画面（`ui/`）は IE11 相当で動くように書いてあり（ES5・`innerHTML` 禁止・CSS 変数と grid 禁止）、提案書とHTMLレポートだけが**ブラウザで開く成果物**として Chromium を前提にする。区画④の[ブラウザで開く]も既定ブラウザへ渡す。
- 実機手順の正は `docs/24` §7（v7.6 で「Microsoft Edge で開く」に統一済み）。**IE11 での表示は要件ではない**ので、IE用フォールバックCSS（実値での二重指定）は書かない（§5.1 の「19変数以外を定義しない」と両立しないため）。

### 11.2 `?debug=1`（社内だけで使う確認用スイッチ）

| URL | 効果 | 誰が使うか |
|---|---|---|
| （何も付けない） | 22枚をそのまま表示する。**お客さまに渡す状態** | お客さま・営業担当者 |
| `?notes=1` | 発表者ノート（§7）を画面に出す。印刷では常に消える | 商談前の営業担当者 |
| `?debug=1` | **描画に失敗した枚（`data-render-error="1"`）に赤枠を出す**（§4.1） | 開発担当者・不具合報告のとき |

- どちらのスイッチも**ファイルの中身は同じ**であり、付けなければ何も起きない。`?debug=1` は `body.debug` を付けるだけで、色は `rgba()` で書く（§5.1 により `:root` の外に生の色指定を書けないため）。
- 赤枠が出た枚があれば、それは「次回のお打ち合わせで更新します」の**保留ではなく不具合**である。区画④の[報告メールを作成]で err_log とともに送る。
