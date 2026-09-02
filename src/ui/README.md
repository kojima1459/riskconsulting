# src/ui - 画面層

起動シーケンス（modBoot）と各シートの描画・読取・進捗／ui_lock を持つ最上位層。app / core を参照してよく、app・core から ui を参照してはならない（依存方向 ui → app → core・12章§4）。Excelトークン（Worksheets / Range( / Application. / ThisWorkbook / MsgBox / ActiveSheet）を書いてよいのは原則この層だけ（R4）。

## モジュール構成(12章§2・30,000字契約による分割)

| モジュール | 責務 | タスク |
|---|---|---|
| `modBoot` | 起動シーケンス(12章§2.1の7手順)・ガードシート制御 | T-16 |
| `modUISheet` | ui層のシート操作プリミティブ(名前付きレンジ・ブロックアンカー・図形ボタン＋OnAction配線・クリップボード) | T-30 |
| `modUIProgress` | `SetStage` / `TryEnterUiLock` / `ExitUiLock` / `ParkFocus`(14章§6・16章 E-11/E-50/E-51) | T-30 |
| `modUIHome` | HOMEの描画・状態表示・図形ボタンの配置表・赤帯 | T-30 |
| `modUIHome2` | 上の分割先（30,000字契約・17章§7 Z-13）。HOMEとS1～S4のOnActionハンドラ群（プレイ起動・出力・企業ファイル・画面遷移） | T-30 |
| `modUICase` | enum変換表(19章§3と一致必須)・入力規則の隠しレンジ・匿名化(E-31)の置換/復元・S1～S4の入口 | T-31 |
| `modUICaseFmt` | 13章§2.2 セル格納規約の変換(セル⇔JSON値)と 13章§2.12-§2.15 の列定義。純関数 | T-31 |
| `modUICase2` | S1～S4シートの描画と逆シリアライズ(`SerializeSheet` の本体) | T-31 |
| `modUICase3` | 案件入力(貼付欄9欄・続き欄・字数カウンタ・追加収集・E-01～05/E-31) | T-31 |
| `modUICase4` | フィードバック・判断台帳 | T-31 |
| `modUIInbox` | 受信箱(投函・一括診断・診断表示・判定入力 E-41) | T-32 |
| `modUISparring` | 壁打ち(履歴表示・発話入出力・受信箱へ送信) | T-34 |
| `modUIGuide` | 初回ガイドツアー(カード3枚・Shape接頭辞 gt_)と操作ガイドの図形ボタン | T-30 |
| `modUIToast` | トースト(図形カード ts_ ＋ `Application.OnTime` の自動消去)・次の一手の文言・`hm_warning` の1行組み立て(裁定書17 H2/H4) | T-30 |

## この層の約束(11章§5)

- **ボタンは図形＋OnAction**。配線は `modUISheet.EnsureButton` の1本に集約し、
  OnActionで呼ばれるPublicハンドラは**先頭で `modUIProgress.TryEnterUiLock`** を通す
  (16章 E-11。`tools/vba_lint.py` の `check_onaction_handler_guard` が機械検査する)。
- **UserFormは使わない**。シート保護(`Worksheet.Protect`)も使わない(16章 E-51)。
- **絵文字リテラルを書かない**。必要なときは `ChrW` で組み立てる(VBEのCP932保持で
  `?` へ化けるため)。
- 日本語ラベル⇔enumの変換は `modUICase` の変換表**だけ**で行う
  (19章§3との一致は `tools/enum_check.py` が機械照合する)。
