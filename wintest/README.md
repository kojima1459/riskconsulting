# Windows実機テスト環境(層(b)の実行環境)

> **この文書は誰向けか**: 「リスク提案ナビ」の実機テストを、自分の手ではなく
> **Windows PC上のClaude Codeに自動でやらせたい**開発オーナー向け。

## なぜ実機が要るのか(17章§1)

テストは3層に分かれており、**出荷条件は層(b)だけ**である。

| 層 | 実体 | どこで走る | 位置づけ |
|---|---|---|---|
| (a) 純VBA | `modTestsPure1..n`(実行と集計は `modTestRunner.RunAllPureTests`) | どこでも | 中身の検証 |
| (b) Excel固有 | `modTestsExcel`(**T-47で作る**) | **wintest 実機のみ** | **出荷条件** |
| (c) 実機前静的 | `tools/vba_lint.py` + `tools/run_lo_tests.py` | Linux | **コミット条件**。実機の代用ではない |

LibreOffice の合格を出荷条件にしない。姉妹PJ(PoC「マイ本棚AI」)は
LO緑のまま出荷して34ラウンドの手戻りを出した。(c)はタイプミス1つを実機まで
持ち越さないための門であって、実機の代わりではない。

## 仕組み(3行)

1. WindowsPCにデスクトップ版Excelと [Claude Code](https://claude.com/claude-code) を入れる
2. そのPC上のClaudeがPowerShell経由でExcelを直接操縦(COM自動化)し、
   ビルド -> 自己インストール -> マクロ実行 -> 結果検証を全自動で回す
3. 本物の「リボンちゃん」が無い分は、確定シグネチャと同一のスタブ
   **「ニセリボンちゃん」**(`mock_ribbon/modMockRibbon.bas`)で配管を検証する

## 買い物リスト

| もの | 目安 | 備考 |
|---|---|---|
| Windows PC | 中古2万円台から | Windows 10/11 64bit・メモリ8GB以上。会社PCと同じWindows 11推奨 |
| Excel | Microsoft 365 Personal または Office Home & Business | **デスクトップ版必須**。Web版・モバイル版はVBAが動かないので不可 |

Mac は不可(Mac版ExcelはVBAの互換性が別物で、本製品はWindows専用設計)。

## 初期セットアップ(1回だけ)

1. ExcelをインストールしてMicrosoftアカウントでライセンス認証する
2. Git for Windows と Python 3(ビルド用)を入れる
3. Claude Code を入れる
4. このリポジトリを clone する
5. **ニセリボンちゃんを作る**(T-14b で使う。5分・1回だけ):
   1. Excelで空のブックを開き `Alt+F11` でVBEを開く
   2. 「ファイル -> ファイルのインポート」で `wintest/mock_ribbon/modMockRibbon.bas` を取り込む
   3. `F12`(名前を付けて保存)-> ファイルの種類「**Excelアドイン(*.xlam)**」->
      ファイル名 **`リボンちゃん(検証用).xlam`** で保存
      (※ **「リボンちゃん」という文字を必ず含めること**。本製品のアドイン検出は
      config `ribbon_addin_name`(既定「リボンちゃん」)の**部分一致**のため。14章§2)
   4. 保存先は既定のAddInsフォルダ(`%APPDATA%\Microsoft\AddIns`)のままでよい

## 実行

```powershell
# 1) ビルド(開発版 = config mock_llm=TRUE)
python build\build_rpn.py --dev

# 2) 実機テスト(自己インストール + 層(a) 純ロジックテストを実機VBAで実行)
powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1

# 3) 層(b) Excel固有テストも回す(T-47 で modTestsExcel を作ってから)
powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1 `
  -ExcelLayerEntry "modTestsExcel.RunAllExcelTests"

# 4) 本番版 + ニセリボンちゃんで、リボン依存の配管まで検証(T-14b)
python build\build_rpn.py --prod
powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1 -Target prod `
  -MockAddinPath "$env:APPDATA\Microsoft\AddIns\リボンちゃん(検証用).xlam"
```

### 合格条件(17章§4-1・T-46⑤)

`run_excel_tests.ps1` は次の**3つすべて**が成立したときだけ exit 0 を返す。

- `modTestRunner.Failures()` が **0件**
- レポートの **SKIP が0件**([SKIP]を貼ればテストはPASSにもFAILにも現れないため、
  件数そのものを見張る)
- **実行本数 = `wintest/tests_expected.txt` の値**(0件実行の「全緑」を成立させない)

`tests_expected.txt` は1行目に10進整数のみを書く。**テストを増減したコミットで
このファイルを更新するのは実装者の義務**であり、更新漏れは即FAILとして現れる。

## T-14b(ニセリボンちゃん配管検証)で確認する4点

`modGatewayRPN.CallStep` が ribbon 経路を通ったうえで、

1. **12引数が宣言順どおり型どおりに渡る**
   (Text, roleSystem, Temperature, MaxTokens, Wait, optModel, prevU, prevA,
   toolN, reasoning_effort, verbosity)
2. **アドイン名の部分一致**でリボンを検出する
3. **`LimitCheck` が True/False いずれでも仕様どおり解釈される**
   (**True でも起動を止めない**。12章§2.1 手順(7))
4. 送信した **toolN の内容が応答へ反映される**

ニセリボンちゃんは受領した引数を応答本文へそのまま書き戻すので、1と4はログの
突き合わせだけで確認できる。

## T-46(出荷前検問)との関係

`tools/ship_check.py` が ①②③ を機械実行する。**④(wintest全PASS)と
⑤(modTestsPure の本数条件)はここで確認する**。5項目のうち1つでも落ちたら
出荷しない。

## 既知の制約(W4.1時点)

- **自己インストーラは移植済み**: `build/template_skeleton.xlsm` と `build/ovba.py` は
  リポジトリに実在し、`build/build_rpn.py` が Stage 4 で `vbaProject.bin` の外科パッチ
  (ThisWorkbookストリームへの自己インストーラ注入)を当て、注入結果を自己検証している。
  W0時点の「手順2以降は実機で動かない」「VBEで手動インポートする暫定手順」は**解消済み**であり、
  上の[実行]のコマンドをそのまま流せる。
- **層(b)の実体が未実装**: `modTestsExcel`(17章 T-47)がまだ無いため、`run_excel_tests.ps1` に
  `-ExcelLayerEntry` を渡さない実行はログに
  「層(b)未指定: -ExcelLayerEntry を渡していないため modTestsExcel は実行していません(T-47)」
  の1行を出す。**この行が出た実行は T-46(4) を満たしていない**(合格扱いにしない)。
  現時点で実機確認できるのは層(a)(`modTestsPure`)の分だけである。
- **Linux側の検問はまとめて回すのが標準**: `python3 tools/gate.py`(全15ゲート一括)。
  個別ツールの直接実行はデバッグ時のみ(17章§5-1)。

## セキュリティ上の注意

- ニセリボンちゃんはAPIキーもエンドポイントも一切含まない(完全スタンドアロン)
- **会社PCにニセリボンちゃんを入れないこと**(本物と名前が部分一致するため誤検出の元)
- 個人PCに会社の実データ(顧客情報を含む資料)を持ち込まない。テストはダミーで行う
