# src/core - 基盤層

製品固有の語彙（製品名・シート名・ドメインenum・ドメイン型）を持たない汎用基盤（ゲートウェイ／config／ログ／JSON／文字列ユーティリティ）。依存方向は ui → app → core の一方向で、core は app / ui を参照しない（12章§4）。

## W6第1弾(T-49)で加えた純関数モジュール

| モジュール | 責務 |
|---|---|
| `modUIGeom` | 画面の幾何（帯・ボタンの並び・カードの高さと表示時間）の純関数。Excelを1つも触らないので層(a)からテストできる（11章§8.6の流用表） |
| `modNavText` | 貼付テキストの純変換（`StripDrFooter` / `PreviewLines` / `SplitFieldNotes` / `JoinFieldNotes`。11章§7.2(a)） |

どちらも**製品固有の語彙（製品名・シート名）を持たない**という core の約束を守る。フッター語の一覧の正は 15章§2.0 であり、登記されるまでは `modNavText` の `NT_FOOTER_WORDS` が実体を持つ（司令塔への申し送り事項）。

## W10.1(T-60)で加えたモジュール

| モジュール | 責務 |
|---|---|
| `modUtilPath` | パスの連結（`JoinPathWith`（純）/ `JoinPath`）・分解（`FileNameOf`（純））・一時フォルダ（`TempDir`）。**区切り文字（`\` / `/`）を知っているのはここだけ**という状態を作る（12章§4。決め打ち連結と `Environ$("TEMP")`/`("TMP")`/`("TMPDIR")` は `tools/vba_lint.py` が ERROR で止める）。`modUtil` が 30,000字契約で満杯のため分割した |

`modUtilPath` は **Excelトークンを1つも持たない**（R4の例外を作っていない）。本体と同じフォルダの値も `modUtil.BookDirHint()`（`modBoot` が起動手順(2)で `modUtil.SetBookDir` へ預けたもの）から借りる。Mac版Excelかどうかの判定（`Application.OperatingSystem`）は層(b)のSKIP判定にしか使わないため、core ではなく**テスト層の `modTestsExcel3.IsMacExcel`** が持つ（司令塔の裁定 2026-09-05）。
