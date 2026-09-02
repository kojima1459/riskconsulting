# src/ui - 画面層

起動シーケンス（modBoot）と各シートの描画・読取・進捗／ui_lock を持つ最上位層。app / core を参照してよく、app・core から ui を参照してはならない（依存方向 ui → app → core・12章§4）。Excelトークン（Worksheets / Range( / Application. / ThisWorkbook / MsgBox / ActiveSheet）を書いてよいのは原則この層だけ（R4）。

## モジュール構成(12章§2・30,000字契約による分割)

| モジュール | 責務 | タスク |
|---|---|---|
| `modBoot` | 起動シーケンス(12章§2.1の7手順)・ガードシート制御 | T-16 |
| `modUISheet` | ui層のシート操作プリミティブ(名前付きレンジ・ブロックアンカー・図形ボタン＋OnAction配線・クリップボード) | T-30 |
| `modUIProgress` | `SetStage` / `TryEnterUiLock` / `ExitUiLock` / `ParkFocus`(14章§6・16章 E-11/E-50/E-51) | T-30 |
| `modUIHome` | ナビの状態表示・赤帯（v3.2でHOME廃止。hm_* は名前を変えずナビのセルへ付け替え。11章§8.1） | T-30 |
| `modUIHome2` | 上の分割先（30,000字契約・17章§7 Z-13）。ナビとS1～S4のOnActionハンドラ群（プレイ起動・出力・企業ファイル・画面遷移） | T-30 |
| `modUINav` | ナビの状態（STEPの自動決定＝11章§3.1.1の優先順位10行）・図形ボタンの配置表・OnActionハンドラ（NavPrev/NavNext/ShowDrafts/BackToNav） | T-49 |
| `modUINavDraw` | ナビの描画（コーチ帯・4区画のパネル・強調枠 nv_focus・待ちカード nv_wait・区画②の状態行とプレビュー）。状態は持たず描くだけ | T-49 |
| `modUIResearch` | 区画①の調べる文8本の組立と[コピー]。雛形は使い方タブの非表示行が唯一の値源（11章§3.2） | T-49 |
| `modUICase6` | 区画②の保管＋プレビュー貼付（[ここに貼る][中身を見る][消す]・case_dataへ直接保存・現場メモの分解）。11章§3.3・§7.2(a) | T-49 |
| `modUICase` | enum変換表(19章§3と一致必須)・入力規則の隠しレンジ・匿名化(E-31)の置換/復元・S1～S4の入口 | T-31 |
| `modUICaseFmt` | 13章§2.2 セル格納規約の変換(セル⇔JSON値)と 13章§2.12-§2.15 の列定義。純関数 | T-31 |
| `modUICase2` | S1～S4シートの描画と逆シリアライズ(`SerializeSheet` の本体) | T-31 |
| `modUICase3` | 旧案件入力(貼付欄9欄・続き欄・字数カウンタ・追加収集・E-01～05/E-31)。**v3.2で区画②は modUICase6 が担うため呼ばれない経路が残る**（撤去は第2弾・Z-5） | T-31 |
| `modUICase4` | フィードバック・判断台帳 | T-31 |
| `modUIInbox` | 受信箱(投函・一括診断・診断表示・判定入力 E-41) | T-32 |
| `modUISparring` | 壁打ち(履歴表示・発話入出力・受信箱へ送信) | T-34 |
| `modUIGuide` | 初回ガイドツアー(カード4枚・Shape接頭辞 gt_)と使い方タブの図形ボタン([テストを実行][ツアーをもう一度見る][記録を見る][表示する]5本) | T-30 |
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
