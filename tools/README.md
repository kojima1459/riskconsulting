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
  -> build/build_rpn.py でビルド
  -> sheet_check.py(13章 <-> シート台帳 <-> 成果物ブックの照合。T-02/T-03)
  -> ship_check.py(T-46 の(1)(2)(3))
  -> wintest 実機(T-46 の(4)(5))
  -> app_version 更新 -> 15分スモーク -> 署名 -> 共有フォルダ配置
```

- `vba_lint.py` と `run_lo_tests.py` は**コミット条件**。
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
- **30,000字契約**(28,000字で警告)
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
