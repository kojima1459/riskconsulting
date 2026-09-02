# tools/ - 開発機用ツールチェーン

「リスク提案ナビ」の開発で毎回回す検査群。**配布物には一切含まれない**(開発機のみ)。

## 環境要件

| もの | 要件 | 使うツール | 備考 |
|---|---|---|---|
| Python | **3.9 以上**(検証環境は 3.11) | 全ツール | 標準ライブラリのみで動く(下記の例外を除く) |
| LibreOffice(`soffice`) | 7.x 以降 | `run_lo_tests.py` | ヘッドレス実行。`/usr/bin/soffice` か PATH 上にあること |
| `openpyxl` | 3.x | `build/build_rpn.py` / `sheet_check.py` | ブックの組み立てと自己検証 / 成果物ブックの読み出し |
| `olefile` | 0.4 以降 | `build/build_rpn.py` | テンプレート由来 `vbaProject.bin` への自己インストーラ外科パッチに**必須**(テンプレート無しの縮退ビルドなら無くても通る) |
| `git` | 任意のバージョン | `ship_check.py` | `git ls-files` / `git check-ignore` を使う |
| `coreutils` の `timeout` | - | `run_lo_tests.py` | soffice のハング検知とプロセス後始末を委譲している |

Windows実機テスト(層(b))の要件は `wintest/README.md` を参照。

## いつ何を回すか(17章§1・§5-1)

```
vba_lint.py 緑
  -> run_lo_tests.py 緑(全モジュール構文コンパイル + 純ロジック実行)
  -> modTestsPure(FAIL 0 / SKIP 0 / 実行本数 = tests_expected)
  -> prompt_diff.py --strict / validate_check.py(17章§4-2の一致検査)
  -> build/build_rpn.py でビルド
  -> sheet_check.py(13章 <-> シート台帳 <-> 成果物ブックの照合。T-02/T-03)
  -> ship_check.py(T-46 の(1)(2)(3))
  -> wintest 実機(T-46 の(4)(5))
  -> app_version 更新 -> 15分スモーク -> 署名 -> 共有フォルダ配置
```

- **検問はまとめて `python3 tools/gate.py` で回すのが標準**(全16ゲート・緑は1行/赤だけ末尾ログ+全文ログパス。テストが数千本規模になっても出力が肥大しない非対称出力)。`--only lint,lo-pure` で絞り、`--tail N` で失敗時表示量を調整、`--list` で一覧。個別ツールの直接実行はデバッグ時のみ。
- `vba_lint.py` と `run_lo_tests.py` は**コミット条件**。HTMLテンプレ系(modHtmlTemplate*/modHtmlTheme/modExportHtml)へ触れたコミットは `render_report.py`(standard・--faithful の両方)も**コミット条件**に含める(18章固定文の逐語照合・DATAリテラル検査・SEC-16非描画検査はここが唯一の検問。層(a)のG90/G91は空白畳み照合のため見出し内空白の漂流には盲)。
  これらが緑でも**実機 wintest(層b)を飛ばしてよい理由にはならない**(17章§1)。
- 検問を1つでも飛ばした版は配布しない。

## 各ツール

### `vba_lint.py` - 静的Lint(層(c))

```bash
python3 tools/vba_lint.py                    # <repo>/src 配下を検査
python3 tools/vba_lint.py --path <dir>       # 対象を変える(検査自体のテスト用)
python3 tools/vba_lint.py --dump-argcount-skips
# exit code: 0 = ERROR 0件 / 1 = ERROR 1件以上
```

検査するもの(主なもの):

- **仕様側プロンプト本文のCP932検査**(裁定書6 A-2): `.bas` に加えて
  `docs/spec/15_プロンプトとJSONスキーマ.md` と `docs/08_ドシエ収集プロンプト集.md` の
  **コードフェンス内**を同じ基準(`_CP932_DENY` の6字＋cp932コーデック)で検査する。
  実装は15章から一字一句写すので、仕様側が汚れると下流の `.bas` が全部赤くなる。
  発生源で止めるための検問(`check_docs_prompt_cp932`)。
- **契約**: 14章§6 / 15章§10.2 / 18章§4.4・§5.2 の公開関数が実装されているか
  (`CONTRACT`)。ファイルがまだ無いモジュールは SKIP でエラーにしない。
  起動時に「`MODULE_REGISTRY` の各モジュールに `CONTRACT` 定義があるか」を自己検査し、
  未定義があれば1件の **WARN** で一覧化する(公開契約の整備漏れの検出。裁定書4 項目14)
- **R1** 依存方向 ui -> app -> core の一方向。製品コードからテスト層を参照しない
- **R3** `Application.Run` は `modGatewayRPN` のみ
- **R4** Excelトークンは ui層 と `R4_EXCEL_ALLOWED_MODULES` のみ
- **30,000字契約**(28,000字で警告)。**`modHtmlTemplate*` だけは25,000字でERROR**
  (18章§4.4の分割規約。17章 T-35 のDoDが本ツールへ検査を委ねている。28,000字の
  WARN帯に届かない超過を見逃さないための専用閾値)
- **CP932安全**: VBEはソースをCP932で保持するため、CP932外文字は実行時に "?" 化ける
- **NFR-S7 ①③**(17章 T-46①を毎コミット走らせる): セルへの書込は `SetCellSafe`、
  HTML連結は `HtmlSafe` / `JsStringSafe` を通す。定数・ヘッダの書込は行末に
  `' SAFE:const`、エスケープ不要が確実なHTML行は `' SAFE:html` を付けて明示する
- **12章§4**: core層に製品名・シート名を書かない
- **実機Excel特有の地雷**(PoCが実機事故で学んだもの。LibreOffice では再現しない):
  MS-VBAL の reserved-name / special-form(`Scale` 等)、VBA予約語との衝突、
  `Dim a, b As T` の型落ち、`As Integer`、負範囲 `ReDim`、モジュールレベル宣言の
  位置、切り出し漏れ(未定義参照)、エラーハンドラの `Resume` 作法、
  `.Find(` の `LookIn` 省略、`Split()` の具体配列型引数への直渡し、
  修飾呼び出しの引数数不一致
- **16進リテラルの Long 接尾辞**(裁定書8 A-5): 値が `&H8000` 以上の16進リテラルには
  接尾辞 `&` を必須とする。VBAは16進の型を桁数で決めるため、1～4桁は符号付き
  Integer になり `&H9FFF` が **-24577** へ化ける。W2b ではこれで `modPii` の漢字域
  判定が常に False になり漢字姓の検知が全滅した(PII走査の素通り=16章E-05の骨抜き)。
  LibreOffice も同じ型付けをするので、**当該の値を通るテストがあれば** `run_lo_tests`
  でも落ちる(実測: `modPii` の接尾辞を外すと純ロジック3本がFAIL)。ただしそれは
  テストの網掛かり次第の**間接検出**で、網の無い定数・新規コードは素通りする。
  本検査はその依存を断って**静的に確実へ**変えるために置いている。
  例外は **`ChrW()` の唯一の引数**に置かれた素のリテラルのみで、この位置は
  -32768～65535 を受けるため値が保たれる。ただし黙って通さず **WARN** で表に出す

### `enum_check.py` - 19章§3とenum変換表の一致検査(17章§4-2の2本目・T-31)

```bash
python3 tools/enum_check.py             # 照合
python3 tools/enum_check.py --dump      # 19章から期待されるCSV本文を出力
python3 tools/enum_check.py --dump-bas  # 上を .bas の連結文へ整形して出力
# exit code: 0 = 一致 / 1 = 不一致
```

11章§5 は「日本語ラベル⇔enumの変換は modUICase の共通変換表(19章と一致必須)のみで
行う」と定めている。その**一致必須**を目視から機械へ移したもの。

- **分類の完全性**: 19章§3の全行を、変換表に載せる行(`REQUIRED`)と載せない行
  (`EXCLUDED`＋理由)へ**漏れなく**分類できることを検査する。19章にenum行が増えたのに
  分類表へ足していなければ落ちる(新しいenumが黙って変換表から漏れるのを防ぐ)。
- **中身の一致**: 載せる行は `modUICase.EnumPairsCsv()` と機械値・日本語ラベル・
  **並び順**まで一致すること。1グループ内での機械値・ラベルの重複(逆引きが一意に
  決まらない)と、CSV区切り `,` の混入も検査する。
- **静的評価**: `prompt_diff.py` と同じく `s = s & "..." & vbLf` 方式の関数本体を
  静的に評価する。制御構文・未対応の項があれば不一致に数えるので、**表を関数の中で
  組み立て直して検査を骨抜きにできない**。
- 実装を直すときは 19章§3 を直してから `--dump-bas` の出力で `EnumPairsCsv()` を
  差し替えること(手で写さない)。

### `caption_check.py` - HOMEのボタン名(キャプション)の4系統逐語照合(W5.2.1裁定)

```bash
python3 tools/caption_check.py
# exit code: 0 = OK / 1 = NG
```

HOMEのボタン名は**同じ文字列が4か所に別々に書かれている**。W5/W5.2 では旧名
([① 案件を作る] 等)が文書側に残り、利用者からは「押せと書いてあるボタンが画面に無い」
状態になった。人手の目視でしか守れていなかった一致を機械化する。

- **値源は実装**: `src/ui/modUIHome.bas` の `UH_ROW_MAIN1`～`UH_ROW_SUB4`
  (「図形名;キャプション;OnAction;幅pt」の配置表)を静的に読み、主要動線5本＋
  くわしい操作14本＝19本を取り出す。
- **完全一致を求める2系統**: `build/build_rpn.py` の `GUIDE_HOME_BUTTONS`
  (操作ガイド③のボタン早見表)と、13章§2.10 の配置表の `[～]` 表記。
- **部分集合でよい1系統**: `src/ui/modUIGuide.bas` の初回ツアー文言
  (`TitleOf`/`BodyOf`)の `[～]` 表記(ツアーは全ボタンには触れないため)。
- **集約表記の許容**: 早見表だけは [S1][S2][S3][S4] を「S1 / S2 / S3 / S4」の1行へ
  まとめてよい。許すのはこの1件だけで、`AGGREGATES` に明記してある。
- **fail-closed**: 対象ファイル不在・定数/表/関数のアンカー不在・抽出0件・
  実装側キャプションの重複は、いずれも**NG**にする(照合対象が無いので緑、を作らない)。

### `t48_check.py` - ブック内テスト実行(T-48)の合否判定4条件の検査(裁定書16 F1)

```bash
python3 tools/t48_check.py          # 検査
python3 tools/t48_check.py --dump   # 抽出した判定式と役の割当を表示
# exit code: 0 = OK / 1 = NG
```

社内PCで回せる唯一の検問が `modTestsRunnerUi.RunAllTestsFromBook` であり、その合否は
**ps1と同一の4条件**(FAIL 0件 / SKIP 0件 / 純層の実行本数=期待本数 / 層(b) 1本以上)で
決まる。この判定式を守る機械検査が無かったため、判定を1条件外しても全ゲート緑のまま
配布できてしまった。本ツールがその穴を塞ぐ。

- **役の割当**: `src/test/modTestsRunnerUi.bas` を静的に読み、And で連なる `If` の各項を
  変数の**代入元まで遡って**解決し、`FailCount`/`SkipCount`/`ExecutedCount`(引き算なし=
  純層 / 差し引き=層(b))/`ExpectedCount` の4役へ過不足なく割り当てる。1条件でも消えれば
  役が欠け、条件を足せば未分類の項が出て落ちる(`Or` への緩和・`>=` への緩和も落ちる)。
- **fail-open検出**: 真枝に「全PASS」、偽枝(`Else`)に「NG」の表示が在ることも見る。
- **規定文との突合**: 17章 T-48 の行に4条件の語がそのまま在ることを検査する(仕様の
  規定文を薄めてから実装を薄める順序の骨抜きを止める)。
- **fail-closed**: 対象モジュール不在・判定関数不在・条件式の抽出0件・T-48行の不在は
  いずれも**NG**にする(検査対象が見つからないので緑、を作らない)。

### `run_lo_tests.py` - LibreOffice実行テスト(層(c))

```bash
python3 tools/run_lo_tests.py                 # モード1+2
python3 tools/run_lo_tests.py --mode pure     # 純ロジック実行のみ
python3 tools/run_lo_tests.py --mode compile  # 全モジュール構文チェックのみ
# exit code: 0 = 全PASS / 1 = いずれか失敗
```

- モード1は `wintest/tests_expected.txt` を読んで `modTestRunner.SetExpectedCount`
  へ渡し、**実行本数の完全一致**を検査する。
- モード1に注入するモジュールは `PURE_ALLOWLIST`。**テストが新しいモジュールを
  叩くようになったら必ず足すこと**(未注入だと実行時エラー12でそのテスト群が
  丸ごと走らないまま「全部PASS」に見える)。
- SKIP の天井は `EXPECTED_SKIP_MAX`(初期値 0)。上げるときは
  「LOでは原理的に実行できない」理由をテスト側のコメントに書いてから上げる。

### `prompt_diff.py` - 15章とプロンプト実装の一致検査(T-23)

```bash
python3 tools/prompt_diff.py            # 未実装関数はスキップ扱い
python3 tools/prompt_diff.py --strict   # 未実装も差分に数える(T-23の最終確認)
# exit code: 差分件数(0 = 一致)
```

15章§10.1 の抽出規約でプロンプト本文を取り出し、15章§10.2 の節⇔関数名対応表を
突合キーにして `src/app/modPrompts*.bas` / `modSchemas*.bas` の関数の戻り値と
比較する。関数は `s = s & "..." & vbLf` 方式(Const禁止・14章§7)で組み立てる
前提で評価する。制御構文や未対応の項があると「評価できません」として差分に数える。

- **`--strict` の対象0件の扱い**(裁定書4 項目13): 突合先ディレクトリが無い、または
  対象 `.bas` が1本も無い(=全関数未実装)ときも、`--strict` では**不合格(exit 1)**に
  倒す。`--strict` を付けない既定はこの状態を「T-23で実装予定のスキップ」として
  **exit 0** で通す(W0～W1では実装前が正常なため)。この非対称が無いと「対象0件」を
  無条件に合格扱いにして `--strict` が骨抜きになる。
- **スキーマJSON検査**(裁定書6 項目13。17章 T-23 DoD「schemaはJSONとしてパース可能」の
  機械化): 実装済みの `Schema*` の抽出本文を `json.loads` でパースし、さらに 14章§3 が
  direct経路の必須要件とする **strict要件**(すべてのオブジェクトで `properties` と
  `required` が一致し `additionalProperties: false` が付いている)を再帰検査する。
  違反は差分と同じく exit code に算入する(`--strict` の合格条件に含まれる)。
  15章とコードを**同じように**壊した変異は本文diffでは検出できないため、この検査が
  最後の砦になる(実測: `additionalProperties` を両側 true にした変異を検出)。
- **突合対象の31関数はすべて無引数**(14章§6の二層分離・裁定書6 B)。実値の埋め込みは
  `modPromptsOps` の `Fill` / `Asm*`(組立層)が行い、突合対象外である。

### `validate_check.py` - 15章§11の検証ルール表と modValidate/modTestsPure の照合(T-22)

```bash
python3 tools/validate_check.py                  # (a)(b)(c)(d) すべて
python3 tools/validate_check.py --no-tests       # (b)を省く(modValidate 実装先行時)
python3 tools/validate_check.py --no-templates   # (d)を省く
# exit code: 0 = 全一致 / 1 = 不一致あり
```

17章§4-2「15章§11の検証ルール表 ⇔ modValidate の照合」の実体。15章§11の表を
**唯一の正**としてケースIDを展開し(範囲記法 `V-S1-01 ～ V-S1-11(11件)` を開く)、

- **(a) 実装**: 判定が「不合格」「警告」のケースは `"[ケースID] ` で始まる文字列
  リテラルが `src/app/modValidate*.bas` にあること。判定「合格」の
  `V-S2C-05` / `V-S3C-05` は**エラー文を持たない**ので、逆にそのリテラルが
  **あってはならない**(ただしIDはコメント等に現れていること)
- **(b) テスト**: `src/test/modTestsPure*.bas` の Check系呼び出しの第1引数
  (テスト名)にケースIDが含まれ、**1ケース1本**であること。1つのテスト名に
  2つ以上のケースIDが同居していたら「1テスト1ケース」違反として落とす
- **(c) 件数**: §11の各行の「(n件)」/ 合計行の「合計64件(不合格52/警告10/合格2)」/
  各Check節の表のID集合 / ケースIDを含むテスト名の総数 が全部一致すること。
  あわせて**判定3値の網羅**(各IDが不合格・警告・合格のどれか1列に必ず載る)も見る
- **(d) エラー文テンプレの一字一句**: 各Check節の表の `エラー文テンプレ` を
  `{...}` で切った固定部が実装ソースにそのまま現れること(文言の漂流検出)。
  照合範囲は**当該ケースIDの近傍だけ**(裁定書7 C-11。当該IDが現れた行から
  次にケースIDが現れる行の直前まで。行継続 `_` の続きも同じ区間に入る)。
  全ソースを1本に連結した素の部分文字列検索では、別ケースが偶然もつ同一文言が
  肩代わりして漂流を見逃す(実測: `V-S3-01` の「件です(3件固定)」を「(3件確定)」へ
  壊しても `V-S3C-01` の同一文言が吸収して緑のままだった)
- **(e) 自己整合**(裁定書7 C-12): (a)(b)(d) の各ループが**64件すべてを回ったこと**を
  件数で確認し、(d)の照合片が0件なら落とす。さらに「エラー文テンプレを持つケース数」
  =「§11で不合格・警告のケース数」を突き合わせる。照合器自身に
  `if cid == "V-S2-06": continue` のような骨抜きを入れると、この検問が落とす
  (実測: continueスキップ+当該リテラル破壊で ERROR 3件・exit 1)

`--no-tests` は modValidate だけ先に出来ている段階のための逃げ道であり、
**出荷前の検問では必ず外して回す**(§4-2はテスト側の存在まで求めている)。

### `sheet_check.py` - 13章とシート実体の照合(T-02 / T-03)

```bash
python3 tools/sheet_check.py                       # dist/ の dev -> prod の順に自動選択
python3 tools/sheet_check.py --book dist/リスク提案ナビ.xlsm
python3 tools/sheet_check.py --no-book             # 台帳(JSON)と13章だけを突合
# exit code: 0 = 全一致 / 1 = 不一致あり
```

**13章のMarkdown表を直接パースして**、`build/sheets_main.json`(台帳)と
ビルド済みブックの両方を突き合わせる。**13章が正・台帳が従**であり、ずれたら
台帳を直す。13章そのものの誤りを疑ったときは直さずに報告する(17章§1)。

照合手順は19章§5に従う:

| 対象 | 突合の相手 |
|---|---|
| シート集合 | 13章§2の節見出し＋§2.4のログ行＋§2.9のガードシート規定 = 全18枚 |
| シート型 | 13章§2.9の型表(帳票型=HOME/案件入力・テーブル型=6枚)と台帳の `role` |
| 列物理名と物理順 | 1シート1テーブル=ブックの1行目 / ブロック縦積み=**ブロック名の名前付きレンジ(アンカー)から特定したヘッダ行** |
| 名前付きレンジ | 帳票型(`hm_` / `ci_`)と見出し用(`hs_` / `sp_` の計8本) |
| configキー | 13章§2.3の name 列とその順序。既定値欄が機械分割できる行は値も |

- 参考として19章§4(本体シート一覧・configキー一覧)との差分を **WARN** で出す。
  19章§5は13章と19章の同時更新を求めるため、WARNが出たら追随漏れを疑う
  (このツールは docs/ を書き換えない)。
- `build/sheets_main.json` で `build_infrastructure: true` としたシート
  (`vba_src`。PoC由来の配布機構で13章には存在しない)は、黙って無視せず
  「仕様外のビルド機構」として明示的に除外し、その事実を出力する。

### `ship_check.py` - 出荷前検問(T-46 の(1)(2)(3))

```bash
python3 tools/ship_check.py
# exit code: 0 = (1)(2)(3) すべてPASS / 1 = いずれか失格
```

- (1) は `vba_lint.py` の同名ルールを import して共有する(二重実装しない)
- (2) のキー検出は**ファイル名と行番号だけを出力し、キー値は出力しない**
  (検問ログ自体が漏洩経路にならないようにする)
- (4)(5) は Windows 実機が要るため本スクリプトでは**未実施**として申告する。
  合格扱いにはしない

## 補足: `build/template_skeleton.xlsm` について

`build/build_rpn.py` はテンプレートの**本物の `vbaProject.bin`** を成果物へ引き継ぎ、
その `ThisWorkbook` ストリームだけを自己インストーラへ外科パッチする(`dir` の
`MOFFSET=0`・`_VBA_PROJECT` の無害化を含む。12章§2 の自己インストール機構)。
低レベルの OVBA 圧縮/解凍・CFB リーダーは `build/ovba.py`(自己完結・実証済み)。

このテンプレートは成果物ではなく**ビルド入力**なので、`.gitignore` の
`!build/template_skeleton.xlsm` 例外で **tracked** にして配る(自己インストール機構を
再現可能にするため。裁定書4 項目12)。中身はキー走査クリーンで、`ship_check.py` の②が
全パート展開して毎リリース検査し、③は `TRACKED_XLSM_ALLOWED` の許可枠として扱う
(成果物 `.xlsm` の tracked は引き続き禁止)。`olefile` はこの外科パッチに必須。

### `render_report.py` - サンプルHTMLレポートの生成(T-33 / T-35 の受入確認)

```bash
python3 tools/render_report.py                 # dist/サンプルレポート.html
python3 tools/render_report.py --theme mono    # dist/サンプルレポート_mono.html
python3 tools/render_report.py --faithful      # 素材合成なし(素のmock・round 1)
# exit code: 0 = 生成+検査OK / 1 = 生成できたが検査NG / 2 = 生成できず
```

LibreOffice へ純文字列モジュール一式(`modHtmlTheme` / `modHtmlTemplate1..6` /
`modExportHtml`)を読み込ませ、mock素材(`modMockLlm.ResponseById`)を入力に
**純組立関数** `modExportHtml.BuildMetaJson` / `BuildReportHtml` を実行して、
人がブラウザで開ける実物を `dist/` に出す。実行機構(雛形プロファイル・.xba変換・
`Public Type` 注入・timeout)は `run_lo_tests.py` を import して流用している。

- 書出だけPython側なのは、製品の `ADODB.Stream`(utf-8・BOMあり。18章§5.3(3))が
  Windows専用でLinuxに無いため。**同じバイト配置**(utf-8-sig)で書く。組立ロジックは
  1行もツール側に持たない。
- 検査: `<meta charset="utf-8">` / 18章§3の全16セクションが登録表にあること /
  `node` があれば最小DOMスタブでページのJSを実際に走らせ `sec-<slug>` が
  16本生成されること / `innerHTML` 系が1つも無いこと(18章§4.1) /
  §5.1の28変数を過不足なく定義し共通CSSに `#fff` 以外の生の色が無いこと。
  **W3.1で3本追加**:
  1. **DATAリテラルに生の `<` が1文字も無い**(18章§5.3(1) v1.1)。`</` だけを
     逃がす旧規約では `<!--<script>`(`-->` を伴わない形)を含むLLM出力で
     HTMLトークナイザが script data double escaped 状態へ入り、DATAブロックの
     正規の `</script>` が終端として働かず**ページが白紙化**する。あわせて
     `<script>` の対が文書全体でちょうど2組であることも見る。
  2. **18章の固定文・見出しの逐語照合**(§3の見出し16件・§3.5の免責4行・
     §3.4のヒアリング2文・SEC-08注記・SEC-09/SEC-12のnote)。**期待値は18章
     Markdownからパースする**のでツール側に写経が無い(二重管理にしない)。
     W3で見つかった「です。」の付加・半角括弧化のような漂流をここで止める。
  3. **DOMスタブのパスB**: `DATA.meta.round_no` を1に落として同じページを
     もう一度描き、SEC-16 が**本文からも目次からも**消えることを見る
     (18章§3「round_no が2未満ならセクションごと非表示。目次からも落とす」。
     この規定の回帰網はここだけで、`--faithful` の素材は status が全件
     proposed のため後段のガードが先に効いて空振りする)。
- `dist/` は `.gitignore` 済み(16章NFR-S2)。サンプルはコミットせず、必要なときに
  このコマンドで再生成する。
