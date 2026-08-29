# 14. API設計（LLM呼び出し仕様と内部インターフェース契約）v2.4

> v2.4.1（裁定書5: W1整合）: §6を唯一の命名権者として確定し、`modGatewayRPN.DecideOk`（帯域外成否の唯一の判定点）・`modUtilText.SanitizeForCell`・`modLog.TruncDetail` / `ShouldRotate`・`modGatewayDirect.CallDirect` / `BackoffMs` / `RetryBudgetFor` / `ParseKeyLine` / `IsOSeriesModel` / `BuildRequestBody`・`modMockLlm.MockResponse` / `ResponseById` / `FaultResponse` を追記。`SanitizeInput` の宣言を実装の Optional ByRef 2本（16章E-04の件数記録）へ追随させ、`GetStr`（エスケープを解いて返す）・`GetArrayItems`（キー不在は空Collection）の未確定点を明文化。§5に括弧の対応判定が文字列リテラル内を数えない旨、§2に `LooksLikeLimitError` の判定（先頭 `#LIMIT:` または「利用上限に達しました」を含む）を固定した。

> v2.4: 実装前監査72件の裁定を反映。§6の関数契約を15章のプロンプト本文が要求する引数へ全面改訂（BuildS1User/S2User/S3User/S4System・CallChatの帯域外成否）、NormalizeLlmJson・MenusSummaryFor・MechsText・LastInjectedIds・SanitizeFileName・TryEnterUiLock・FreezeRoundを新設、§7を定数から純関数へ、§2/§3のハードコードをconfig駆動へ、§5にExtractJsonBlock入力パターン表を新設。（検証指摘の修正）BuildS3Userの引数順を15章の貼付ブロック出現順（menus, lines, schemes, cases）へ修正、JsStringSafeの適用順を明示、modUIProgressにParkFocusを追加、§4のmock本数を「7 step種・11応答」へ、modHtmlTemplate/modHtmlThemeの関数契約の正が18章であることを明記。
> v2.4（W0実装報告の裁定）: §6にtest層の入口 `modTestsExcel.RunAllExcelTests()`（層(b)実機E2Eスモーク。17章 T-47）を追加した。

## 1. LLM経路と分岐

唯一の窓口は `modGatewayRPN.CallStep()`。呼び出し元（modPipeline / modPlayOps）は経路を意識しない。

- 分岐規則（PoC互換）: `mock_llm=TRUE` → mock。それ以外は `llm_transport`（ribbon/direct/mock）
- ribbon選択時にリボン未検出 → **E0201で停止**（directへ自動フォールバックしない。一般利用者PCでの誤課金・誤失敗防止）
- 全呼び出しは run_log に1行記録（step, play, transport, model, latency, validate_result, injected_kb_ids）

### Stepレジストリ（呼び出し単位の正）

| step | play | スキーマ | 呼び出し元 |
|---|---|---|---|
| s1/s2/s3/s4 | PL-01/02 | Schema-S1/S2/S3/S4 | modPipeline |
| s2c/s3c（批判）・s2r/s3r（改訂） | PL-01/02のdeep時 | Schema-S2C/S3C・改訂はS2/S3と同一 | modPipeline |
| pf | PL-03 | Schema-PF | modPlayOps |
| sp | PL-04（壁打ち） | なし（自由対話。CallChat経由） | modSparring |
| wt | PL-05 | Schema-WT | modPlayOps(1.5) |
| fg | PL-06 | Schema-FG | modPlayOps(1.5) |

## 2. ribbon経路（主経路・本番）

確定台帳（PoC `docs/dev/RIBBON_API_CONFIRMED.md`）準拠の12引数呼び出し:

```vb
' 引数値はすべて modConfig 参照（10章 NFR-M3。リテラル直書き禁止）
Dim temp As Double, maxTok As Long, waitSec As Long, model As String, toolN As String
temp    = modConfig.GetDouble("temperature", 0.3)
maxTok  = modConfig.GetLong(stepName & "_max_tokens", modConfig.GetLong("llm_max_tokens", 0))
waitSec = modConfig.GetLong("llm_wait_sec", 1200)
model   = modConfig.GetStr("recommended_model", "")
toolN   = modConfig.GetStr("app_tool_prefix", "リスク提案ナビ:") & stepName
result = Application.Run("ChatGPT", _
    userPrompt, systemPrompt, temp, maxTok, waitSec, model, _
    "", "", toolN, effort, verbosity)
'  1 Text  2 roleSystem  3 Temp(Double, config temperature)
'  4 MaxTokens(Long, 0=リボン側既定(台帳では4096)。長文Stepは config llm_max_tokens
'    または Step別上書き s1_max_tokens 等で明示指定する)
'  5 Wait(config llm_wait_sec)  6 model(config recommended_model)
'  7 prevU  8 prevA(各Step独立のため常に"")
'  9 toolN(管理側ログ識別・必須。config app_tool_prefix + stepName。製品名の直書き禁止)
' 10 effort  11 verbosity(GPT-5系)
```

- 戻り値はAzure OpenAI応答の素通し。response_format指定不可 → JSONは15章のプロンプト強制＋§5の防衛線で担保
- アドイン検出: `Application.AddIns` ループ（`ribbon_addin_name` 部分一致＋Installed、セッションキャッシュ）
- 起動時 `LimitCheck()`（True=続行不可→案内し、実行時に再案内。config limit_check で無効化可）。起動シーケンス上の位置は12章§2.1のmodBoot手順⑦（Trueでも起動は止めない）
- 温度・MaxTokensはGPT-5系では無視され effort/verbosity が効く（V2実運用で確認済み）。`reasoning_tuning` エスケープハッチはPoC同様に維持
- **呼出中の画面**: リボン呼出はVBAを同期ブロックするため、呼出の前に `modUIProgress.SetStage` でStep名・開始時刻・最大待ち時間・「画面が白くなっても処理は続いている」旨を確定表示し、config `keep_window_alive`（既定TRUE）で画面ゴースト化を抑止する（16章 E-50）
- エラー: 空応答=E0202／上限系文字列（LooksLikeLimitError移植）=E0204／アドイン無し=E0201。戻り値 `"#ERR:E02xx:説明"`（例外は投げない）。**ただし成否判定は§6の帯域外フラグで行い、文字列プレフィクスを判定に使わない**
- **`LooksLikeLimitError` の判定を固定する**: 「**先頭が `#LIMIT:`**（前後の空白は無視）**または本文に「利用上限に達しました」を含む**」の2条件のいずれかを満たすときだけ True。語彙の供給元は15章§8.2の `limit` 応答実体1箇所であり、mock と gateway が同じ文字列を見る（語彙を2箇所に書かない）。「上限」「回数」「limit」のような**部分語での曖昧判定はしない**。約款や提案本文はこれらの語を普通に含み、長文の正当な応答が E0204 に誤爆して全Stepが止まる事故が実機で起きているため。なお `#LIMIT:` は**人間向けの表示文字列であって成否の判定材料ではない**（成否は§6の帯域外フラグ。ここでの判定は「ok=False にしたうえでどのエラーコードを載せるか」の分類にのみ使う）

## 3. direct経路（開発・検証用）

`MSXML2.ServerXMLHTTP.6.0`、`setTimeouts 5000,10000,{t},{t}`（t=direct_http_timeout_ms）。

```
POST {direct_api_base}/chat/completions
Authorization: Bearer {keyファイル1行目}   ※ブック・config・ログ・リポジトリ・配布物のいずれにも保存禁止（キーはローテ不能の借用1本＝漏れたら作り直せない。16章NFR-S2）
{
  "model": "{direct_model}", "temperature": {config temperature},
  "max_tokens": {config llm_max_tokens ※0のときはキー自体を送らない},
  "messages": [{"role":"system","content":sys},{"role":"user","content":usr}],
  "response_format": {"type":"json_schema",
    "json_schema":{"name":"{step}","strict":true,"schema":{...modSchemasの該当スキーマ...}}}
}
```

- **strict:true必須**。15章の全スキーマはstrict要件（全プロパティrequired・additionalProperties:false・enum統制）を満たす
- 応答は `choices[0].message.content` を modJsonLite で抽出。`refusal`非空 / `finish_reason≠"stop"` は E0206
- リトライ: 429/500/502/503=指数バックオフ最大3回（2s/4s/8s）。408/タイムアウト=1回。その他4xx=リトライなし
- o系モデル名（先頭"o"）では temperature を送らない。temperature / max_tokens はリテラルで書かず必ず config から読む（NFR-M3。ribbon経路の§2と同じ値源）
- キー不存在=E0205「direct経路は開発者専用です」
- コスト目安: 1案件=4呼び出し・入力≈25k tok・出力≈8k tok → 数十円/案件。PoC全体で数千円以内
- **コスト前提（発注者確認 2026-08-28）: API利用コストは設計制約としない**。トークン節約のための品質妥協（入力の間引き・批判パスの省略・リトライ回数の切詰め）は行わない。有報級の長文（5万字≒25k tok強）を1呼び出しに載せる設計も可。usage ログ（modLog）は引き続き全呼び出しで記録する（コスト管理でなく挙動監視のため）
- **長文入力の根拠（2026-08-28実機検証済み）**: リボン側ラッパーは機能制限なし・Azure OpenAI応答素通し（PoC台帳 RIBBON_API_CONFIRMED.md）。上限はモデルのコンテキスト長のみ。**実機テスト合格**: 三菱電機・有報「事業等のリスク」章全文を1呼び出しで構造化、最終項目まで完走・切り捨てなし。Wait（config `llm_wait_sec`）と MaxTokens（config `llm_max_tokens`・Step別上書き可）は長文時に引数で拡張する前提で設計する
- **長文出力の既知欠陥（尾部劣化）**: 上記テストで**末尾項目の重複出力＋重複側への他項目引用の誤混入**を観測。**16章 E-49 として登録済み**。対応の実体は§5防衛線の(2.5)＝`modValidate.NormalizeLlmJson(stepName, json, removedCount)` であり、**全step・両経路で必須**（S1だけの対策にしない）。配列要素を `modUtilText.NormalizeForHash` 正規化後の fnv1a64 で重複排除し（PoC modPack の fnv 重複排除を転用。docs/08 実機確認1参照）、除去件数を run_log の detail に記録して E0303（重複除去実施・警告）を残す。黙って畳んで済ませない

## 4. mock経路（mockは2本立て。同名にしない）

**(a) 本体内のmockトランスポート = `modMockLlm`**（`src/test/` 配置・12章§2 test層）。`modGatewayRPN` のmock分岐が呼ぶ唯一の相手。step別に**決定的**なサンプルJSON（15章§8「浜松スイーツファクトリー」一式）を返す。乱数・現在時刻不使用。同一入力（正規化ハッシュ）→同一応答。

- 応答は**7 step種・11応答**（15章§8.1の表が正）: S1（new/renewal）／S2（new/renewal）／S3／S4／PF／S2C（issues非空版・0件版）／S3C（lands=true版・issues 0件版）。各応答は**自分のバリアント文脈**で modValidate に合格すること（15章§8.1受入条件1）
- **障害注入**: config `mock_fault`（既定=空。空なら正常応答）
  | 値 | 返すもの | 検査したい挙動 |
  |---|---|---|
  | `broken_json` | 閉じ括弧が欠けた途中切れJSON（毎回） | §5(1)が `""` を返し修復リトライ→なお破損のためE0302で当該Step失敗 |
  | `broken_json_once` | 初回呼出のみ途中切れJSON・修復呼出には正常応答（唯一の状態保持例外。modMockLlm.ResetFaultOnce でリセットし、fault値の変更でも解除） | 修復リトライの成功系（validate_result=repaired）の検証 |
  | `enum_violation` | enum外の値を含む正常JSON | §5(3)の不合格→修復リトライ |
  | `ghost_id` | 実在しない `M-9999` を s3 の menu_ids に含む（15章§8.2と一致） | ID実在チェック（ホワイトリスト照合）の発火 |
  | `count_violation` | risks 件数が規定範囲外 | 件数検証の発火 |
  | `empty` | 空文字 | E0202（空応答） |
  | `limit` | 上限系エラー文字列 | E0204と「往復数を減らして再開」案内（E-44） |
  | `fake_err` | 本文が `#ERR:E0201:...` で始まる**正常なJSON** | §6の帯域外規約。ok=True のまま素通しし、エラーUIに昇格させない |

**(b) `wintest/mock_ribbon/modMockRibbon.bas` = ニセリボンちゃん**（本体ブックの**外**に置くスタブアドイン `.xlam`）。`ChatGPT` / `LimitCheck` を公開し、**`Application.Run` の12引数配管そのもの**（引数の型と順序・アドイン部分一致検出・LimitCheckの真偽解釈・toolN の中身）を実機で検証するためのもの。PoCの確定シグネチャ（Text, roleSystem, Temperature, MaxTokens, Wait, optModel, prevU, prevA, toolN, reasoning_effort, verbosity）を維持し、受領した toolN を応答に反映する。

**(a)と(b)は別物**であり、役割を1モジュールに兼ねさせない。(a)だけでは12引数配管が一度も動作検証されないまま実機1発勝負になるため、両方を用意する（12章§2・17章 T-14）。

## 5. JSON防衛線（両経路共通・全step）

```
raw → (1) modJsonLite.ExtractJsonBlock（説明文・コードフェンス除去・全角記号正規化・最外{}）
    → (2) modJsonLite で対象スキーマに必要なキーのみ抽出（汎用パーサは作らない）
    → (2.5) modValidate.NormalizeLlmJson（配列要素の重複排除＝尾部劣化対策。16章 E-49／E0303）
    → (3) modValidate.Check{S1..S4|S2C|S3C|PF|WT|FG}（型・enum・件数・ID実在・整合）
    → NG: (4) 修復リトライ（json_repair_retry=1回。15章§7のサフィックスを user末尾に追記して同stepを再呼び出し）
    → なおNG: E0302でstep失敗。生応答は case_data の `sN_json_failed`（PFは受信箱の pf_json）へ保存し、
              検証合格済みの `sN_json` を上書きしない（13章§2.2）
```
- (1)〜(2.5)を通すのは **modPipeline / modPlayOps の責務**（modGatewayRPNは素通しのまま返す）。(2.5)は全step・両経路で必ず通す
- directはstrictで(1)(2)がほぼ素通しになるが、**(3)の業務検証（ID実在・件数・整合）は両経路で必須**（スキーマでは表現できない）
- validate_result（ok/repaired/failed）を run_log に記録し、月次でプロンプト改善のシグナルにする

### ExtractJsonBlock 入力パターン表（実体の固定。テストはこの7本を最低ケースとする）

戻り値 `""` は「抽出失敗」を意味し、上表(4)の修復リトライへ進む。壊れたJSONを推測で補完しない（欠落を捏造しないため）。

| # | 入力パターン | 実体（LLM応答の例） | ExtractJsonBlockの動作 | 戻り値 |
|---|---|---|---|---|
| ① | 前後に説明文 | `承知しました。{"a":1} 以上です。` | 最初の `{` から対応が閉じる `}` までを切り出す | 正常抽出 |
| ② | コードフェンス（3連バッククォート+json） | フェンス行で囲まれたJSON | フェンス行を除去してから①と同じ処理 | 正常抽出 |
| ③ | フェンス閉じ忘れ | 開始フェンス行のみ有り・終端フェンス無し | 開始フェンス行を除去し、以降で括弧が閉じていれば抽出 | 正常抽出 |
| ④ | 末尾途切れ（閉じ括弧欠落） | `{"a":1,"b":[{"c":2}` | 括弧の対応が最後まで取れない。**補完しない** | `""` |
| ⑤ | JSON2連結 | `{"a":1}{"a":2}` | 最初に対応が閉じた1本目のみを返し、2本目は破棄。破棄したことを run_log detail に `extra_json=1` と記録 | 正常抽出（1本目） |
| ⑥ | 全角波括弧・全角引用符 | `｛"a"：1｝` | 前処理P0（下記）で構造記号を半角化してから① | 正常抽出 |
| ⑦ | 文字列値内の生改行とエスケープ引用符 | 値の途中に生のCR/LFがあり、値の中に `\"` を含む | 抽出は括弧の対応のみを見るので成功。値の整形は(2)で行う（下記） | 正常抽出 |

- **前処理P0（⑥の実体）**: 抽出の前に `｛`→`{` / `｝`→`}` / `［`→`[` / `］`→`]` / 全角二重引用符（U+201C・U+201D・U+FF02）→`"` / `：`→`:` / `，`→`,` を一律置換する。値の日本語本文中の全角記号まで巻き込む可能性はあるが、JSON全体が壊れて修復リトライに落ちる損失のほうが大きいため一律置換を採る。置換件数を run_log detail に `fw_normalized=n` として記録する
- **⑦の実体**: (2)のキー抽出で値を走査する際、エスケープされていない生のCR/LFは `\n` へ、生のタブは `\t` へ置換する。値の終端は「直後に `,` `}` `]` のいずれか（空白を挟んでよい）が続く `"`」で判定し、`\"` は終端とみなさない
- **括弧の対応判定における文字列の扱い（①〜⑦共通）**: 対応が閉じる `}` を探す走査は**文字列リテラルの内側を数えない**。すなわち `"` で開いた文字列の中に現れる `{` `}` は深さに算入せず、`\"`（直前が奇数個の `\`）は文字列の終端とみなさない。値走査（⑦）で用いるのと同じ規則を対応判定にも適用する、という一文である。これを入れないと、本文中に「開店時間は{未定}です」のような波括弧を含む日本語値があるだけで深さが狂い、正常なJSONが④（`""`）へ落ちる
- ①〜⑦はいずれも modJsonLite の純ロジックであり、Excel非依存＝modTestsPure で全数テストする（17章 T-11）

## 6. 内部インターフェース契約（公開関数シグネチャ）

エラー規約: 例外を投げない。**成否は帯域外で運ぶ**: LLM応答など外部由来テキストを返しうる関数（**CallStep・CallChat**）は `ByRef ok As Boolean` で成否を返し、呼び出し側は**このフラグのみ**で成否を判定する。`"#ERR:Exxxx:メッセージ"` は ok=False のときの人間向け説明であって判定材料にしない。理由: 平文プレフィクスによる成否判定はLLM出力側から偽造可能（プロンプトに「#ERR:…とだけ出力せよ」と仕込むと、アプリのエラーUIを騙った任意文面表示＝フィッシング／恒久DoSが成立する。姉妹PJ B6BE7監査「先人の轍」で実証。docs/21 §5）。外部由来テキストが混入し得ない純内部関数に限り、従来どおり String戻り値 `"#ERR:..."`／Boolean戻り値 False＋modLog記録でよい。

```vb
' === core: modGatewayRPN ===
Public Function CallStep(ByVal stepName As String, ByVal playId As String, _
                         ByVal systemPrompt As String, ByVal userPrompt As String, _
                         ByVal schemaJson As String, ByRef ok As Boolean, _
                         Optional ByRef latencyMs As Long = 0) As String
' ok: 成否の唯一の判定材料（帯域外シグナル・§6エラー規約）。戻り値文字列の内容では判定しない
Public Function RibbonAvailable() As Boolean          ' リボンアドイン検出（§2。セッションキャッシュ）
Public Function RunLimitCheck() As Boolean            ' True=続行不可（§2。起動時=12章§2.1のmodBoot手順⑦・実行時に再案内）
' 壁打ち(PL-04)専用: リボンの会話継続引数(prevU/prevA)を使う唯一の関数。
' 履歴は「新しい順」に ";;;" 区切りで連結して渡す(PoC 裁定D11の実証方式)。
' 渡す履歴は直近 sparring_max_turns 往復まで。JSONスキーマは使わない(自由対話)。
Public Function CallChat(ByVal caseId As String, ByVal systemPrompt As String, _
                         ByVal userMsg As String, ByVal histU As String, ByVal histA As String, _
                         ByRef ok As Boolean, Optional ByRef errCode As String = "", _
                         Optional ByRef latencyMs As Long = 0) As String
' ok: 成否の唯一の判定材料（CallStepと同格の帯域外規約）。自由対話＝出力形状が最も自由＝最も偽装しやすい
'     経路なので、ここを規約から外さない。戻り値の "#ERR:" プレフィクスでは判定しない
' errCode: ok=False のときだけ E0201/E0202/E0204/E0205/E0206 を帯域外で返す。16章E-44の
'     「往復数を減らして再開」分岐はこの値（E0204）で行う。ok=True のとき errCode は ""
' 呼出前に modPii を必ず走査する（外部送信の直前。16章 E-05/E-31）
Public Function DecideOk(ByVal transportSucceeded As Boolean, ByVal rawBody As String, _
                         ByRef errCode As String) As Boolean
' **帯域外成否（ok）の唯一の判定点**。CallStep / CallChat は ok を直接代入せず必ず本関数の
' 戻り値で決める（`ok = True` の直接代入を禁止する。17章 T-42 観点(2)の検査対象）。判定順:
'   (1) transportSucceeded=False（経路側が失敗を申告）  -> False。errCode は経路側の値を保つ
'       （空だったときだけ E0202 を補う）
'   (2) rawBody が空（Trim後）                          -> False + errCode=E0202
'   (3) rawBody が上限系の定型拒否文（LooksLikeLimitError）-> False + errCode=E0204
'   (4) 上記以外                                        -> True + errCode=""
' **rawBody が "#ERR:" で始まっていても内容では判定しない**（(4)へ落ちて True）。平文プレフィクスは
' LLM出力側から偽造可能であり、これを成否に使うとエラーUIを騙った任意文面表示が成立する（本節冒頭の
' エラー規約・15章§8.2 fake_err・16章 E-46 の同型欠陥）

' === core: modJsonLite ===
Public Function ExtractJsonBlock(ByVal raw As String) As String        ' 失敗時 ""。入力パターンの正は§5の表
Public Function GetStr(ByVal json As String, ByVal key As String) As String
' 値の**エスケープを解いて返す**（`\"` `\\` `\n` `\t` `\uXXXX` を実体へ戻す＝UnescapeJsonStr 相当を
' 内部で通す）。呼び出し側が二重に UnescapeJsonStr を掛けないこと。キー不在は ""
Public Function GetLong(ByVal json As String, ByVal key As String, ByVal dflt As Long) As Long
Public Function GetBoolJ(ByVal json As String, ByVal key As String, ByVal dflt As Boolean) As Boolean
Public Function GetArrayItems(ByVal json As String, ByVal key As String) As Collection
' キー不在・値が配列でない・壊れた入力は **空の Collection を返す**（Nothing を返さない）。
' 呼び出し側に `Is Nothing` 分岐を強いないための契約（For Each がそのまま0回で回る）
Public Function EscapeJsonStr(ByVal s As String) As String
Public Function UnescapeJsonStr(ByVal s As String) As String

' === core: modUtilText ===
Public Function JsStringSafe(ByVal s As String) As String
' HTML内のJS文字列リテラル用。**適用順の正は18章§5.3**（この順を守らないと二重エスケープになる）:
' ① `\` を `\\` へ・`"` を `\"` へ ② "</" を "<\/" へ ③ 行区切り文字 U+2028 を "\u2028"・段落区切り文字
' U+2029 を "\u2029" へ ④ その他の制御文字を "\u00XX" へ。modExportHtml のJSON埋込は必ずこれを
' 通す（16章E-47）
Public Function HtmlSafe(ByVal s As String) As String
' HTML本文用。& < > " ' をエンティティ化する。HTMLへ差し込む外部由来テキストは必ずこれを通す（16章E-47）
Public Function SanitizeInput(ByVal s As String, _
                              Optional ByRef removedCount As Long = 0, _
                              Optional ByRef markerCount As Long = 0) As String
' 外部由来テキストの無害化（16章E-04）: 制御文字・私用領域文字を除去し、あわせて本文中の
' "■■■" を "[境界記号]" へ置換する（15章のデータ境界記号の偽装防止。E-43の多層防御の1枚）
' removedCount / markerCount は**件数だけ**を帯域外で返す省略可能な出口。16章E-04が
' 「除去・置換の件数はrun_logのdetailに件数のみ記録し本文は残さない」を求めるため必須（本文は返さない
' ＝NFR-S3）。省略した呼び出し（`SanitizeInput(s)`）も従来どおり動く
Public Function SanitizeFileName(ByVal rawName As String) As String
' ファイル名の安全化。禁止文字 \ / : * ? " < > | と制御文字を "_" へ／前後空白と末尾ピリオドを除去／
' 32字で切詰め＋case_id由来8桁を付与して衝突回避／最終パスが240字を超える場合は company 部を
' Fnv1a64Hex 16桁へ置換。企業ドシエファイル名・report_path・ppt_path に必ず適用する
Public Function NormalizeForHash(ByVal s As String) As String   ' 重複判定用の正規化（(2.5)のfnv1a64の前段）
Public Function Fnv1a64Hex(ByVal s As String) As String         ' 16桁の16進文字列
Public Function SanitizeForCell(ByVal s As String) As String
' SetCellSafe がセルへ書く直前に通す**純変換部**（Excel非依存。CSV書出＝16章NFR-S7(2)も同じ関数を通す）。
' ガード3点の**適用順は16章NFR-S7(1)が正**（NUL除去 -> 32,000字切詰め -> 先頭式記号の ' 前置）

' === core: modLog ===
Public Function TruncDetail(ByVal s As String) As String
' run_log / err_log / usage_log の detail 列の切詰め（**最大400字**。13章§2.4・16章NFR-S3）。
' 本文を残さない規約の実体であり、記録側は必ずこれを通す
Public Function ShouldRotate(ByVal rowCount As Long, ByVal maxRows As Long) As Boolean
' ログのローテ判定（config `log_max_rows`）。**rowCount >= maxRows で True**（閾値ちょうどで回す）。
' maxRows <= 0 はローテ無効で常に False

' === core: modGatewayDirect ===
Public Function CallDirect(ByVal stepName As String, ByVal systemPrompt As String, _
                           ByVal userPrompt As String, ByVal schemaJson As String, _
                           ByRef modelUsed As String, ByRef errCode As String, _
                           ByRef errMsg As String) As String   ' direct経路の唯一の入口（§3）
Public Function BackoffMs(ByVal attemptNo As Long) As Long
' 429/5xx の指数バックオフ間隔（ms）。**attemptNo は1始まり**で 1=2000 / 2=4000 / 3=8000、
' 範囲外は0（§3の「2s/4s/8s」の実体）
Public Function RetryBudgetFor(ByVal httpStatus As Long) As Long
' そのHTTPステータスで**許される再試行回数**（初回の呼び出しを含まない）。§3の表の実体:
' 429=3 / 500・502・503=3 / 408=1（タイムアウトも同じ扱い）/ その他4xx=0。
' ShouldRetryDirect 等の判定はこの関数を唯一の値源とする（回数表を2箇所に書かない）
Public Function ParseKeyLine(ByVal fileText As String) As String
' キーファイルの生テキストから**1行目だけ**を取り出す（16章NFR-S2）。LF/CRLFの双方に対応し、
' UTF-8 BOM を落とし、**前後の空白をTrimする**。値はセル・config・ログ・配布物へ一切出さない
Public Function IsOSeriesModel(ByVal modelName As String) As Boolean   ' 先頭が "o" のモデル（§3。temperatureを送らない）
Public Function BuildRequestBody(ByVal stepName As String, ByVal model As String, _
                                 ByVal systemPrompt As String, ByVal userPrompt As String, _
                                 ByVal schemaJson As String, ByVal temperature As Double, _
                                 ByVal sendTemperature As Boolean, ByVal maxTokens As Long) As String
' §3のリクエストボディ。sendTemperature は `Not IsOSeriesModel(model)` を呼び出し側から渡す
' （1判断1箇所。この関数は model 名を見て自ら判定しない）。maxTokens=0 はキーごと送らない。
' schemaJson が空のときは response_format を出力しない

' === app: modValidate ===  戻り値 ""=合格 / 非空=エラー列挙（修復プロンプト用の日本語）
Public Function NormalizeLlmJson(ByVal stepName As String, ByVal json As String, _
                                 ByRef removedCount As Long) As String
' 尾部劣化（同名項目の重複出力）の正規化。§5防衛線(2.5)。16章E-49。
' 配列要素を NormalizeForHash 後の fnv1a64 で重複排除し、除去件数を removedCount に返す。
' removedCount>0 は E0303（警告）として run_log detail に記録する（黙殺しない）
Public Function CheckS1(ByVal json As String, ByVal caseType As String) As String
Public Function CheckS2(ByVal json As String, ByVal caseType As String) As String
Public Function CheckS3(ByVal json As String, ByVal s2Json As String) As String  ' ID実在はmodKnowledge参照
Public Function CheckS4(ByVal json As String) As String
Public Function CheckS2C(ByVal json As String) As String          ' 入念モードの批判JSON（15章§4.5）
Public Function CheckS3C(ByVal json As String) As String          ' 入念モードの批判JSON（15章§4.6）
Public Function CheckPF(ByVal json As String) As String
Public Function CheckWT(ByVal json As String) As String          ' Phase1.5
Public Function CheckFG(ByVal json As String) As String          ' Phase1.5

' === app: modKnowledge ===
Public Function LoadKnowledge() As Boolean            ' 起動時/再読込。スナップショット保存込み
Public Function RiskLibFor(ByVal industryCode As String) As String     ' 整形済注入テキスト（S2用。書式の正は15章§3）
Public Function MenusSummaryFor(ByVal industryCode As String) As String
' S2用のメニュー「要約」。preventions.related_menu_id が選べる候補一覧を与える唯一の口。
' これを注入せずに CheckS2 の「related_menu_id は "" または実在」を課すと候補ゼロで実在を要求する
' 矛盾になるため必須。整形書式の正は15章§3
Public Function MenusFor(ByVal industryCode As String) As String        ' S3用のメニュー一覧（実在するサービス。S2の要約とは別テキスト。書式の正は15章§4）
Public Function LinesText() As String
Public Function CasesFor(ByVal industryCode As String) As String
Public Function SchemesFor(ByVal industryCode As String) As String     ' status∈{proven,adopted}のみ
Public Function PatternsText() As String                                ' P1-P15全件（PF/FG用）
Public Function RulesText() As String                                   ' 判断基準（PF/FG用）
Public Function ResearchingText() As String                             ' 研究テーマ一覧（PF用）
Public Function MechsText() As String
' 機構ライブラリ抜粋 kb_mech_rows 件（PL-04壁打ちのsystemに注入）。機構シートはPhase1.5だが
' 壁打ちはPhase1のため、未装填・0行のときは "(登録なし)" を返して続行する（16章E-09）
Public Function LastInjectedIds() As String
' 直近の ResetInjectedIds 以降に各注入関数が実際に使ったナレッジIDの累積（";"区切り）。
' run_log.injected_kb_ids（13章§2.4・10章FR-10）はこの値で埋める
Public Sub ResetInjectedIds()   ' Step開始時に modPipeline / modPlayOps が呼び、累積を初期化する
Public Function MenuIdExists(ByVal id As String) As Boolean
Public Function LineIdExists(ByVal id As String) As Boolean
Public Function SchemeIdExists(ByVal id As String) As Boolean
Public Function CaseLibIdExists(ByVal id As String) As Boolean
Public Function PatternIdExists(ByVal id As String) As Boolean
Public Sub AppendServiceGap(ByVal caseId As String, ByVal industryCode As String, ByVal riskDesc As String)

' === app: modPromptsCore / modPromptsBlocks / modPromptsOps / modSchemas ===
' 本文は15章と一字一句一致（T-23がdiffゼロを検査）。**Const は使わず、`s = s & "..." & vbLf` 方式の
' 純関数で組み立てて返す**（VBAの Const は1行1023字・行継続25本の制約に当たり、3,700字級の
' スキーマ本体を1宣言で書けないため）。一致検査の正規化は「改行=vbLf・末尾改行なし」。
' 引数の順序は15章の貼付ブロックの出現順に一致させる（15章の全 {{プレースホルダ}} に対応する引数が
' 存在することが受入条件。19章§5の文書間整合チェックリスト）。
Public Function BuildS1System() As String                                        ' 15章§2 system
Public Function BuildS1User(ByVal ctx As TCaseCtx, ByVal hpTxt As String, ByVal yuhoTxt As String, _
                            ByVal memoTxt As String, ByVal contractTxt As String, ByVal prevRenewalTxt As String, _
                            ByVal dossierTxt As String, ByVal fieldNotes As String, ByVal coverageNote As String, _
                            ByVal hearingAnswers As String) As String
' 15章§2 S1 userの9貼付ブロックを順に埋める（hp→yuho→memo→contract→prevRenewal→dossier→
' fieldNotes→coverageNote→hearingAnswers）。coverageNote=13章 input_coverage_note（付保の見立て。
' 伝聞情報だが insurance_ctx 観点の充足度評価に算入するため省略不可）
' 入念モード(quality_mode=deep)用（15章§4.5〜4.7）:
Public Function BuildS2CriticSystem() As String                                  ' 15章§4.5 批判system
Public Function BuildS2CriticUser(ByVal s1Json As String, ByVal s2Json As String, ByVal riskLib As String) As String
Public Function BuildS3CriticSystem() As String                                  ' 15章§4.6 批判system
Public Function BuildS3CriticUser(ByVal ctx As TCaseCtx, ByVal s1Summary As String, _
                                  ByVal s2Json As String, ByVal s3Json As String) As String
Public Function ReviseSuffix(ByVal critiqueDigest As String) As String           ' 15章§4.7 改訂サフィックス
Public Function BuildSparringSystem(ByVal dossierSummary As String, ByVal s1s2s3Json As String, _
                                    ByVal schemes As String, ByVal patterns As String, _
                                    ByVal mechs As String, ByVal rules As String) As String
' PL-04壁打ちのsystem（15章§6.5）。mechs は modKnowledge.MechsText()（Phase1は "(登録なし)"）
Public Function BuildS2System() As String                                        ' 15章§3 system
Public Function BuildS2User(ByVal ctx As TCaseCtx, ByVal s1Json As String, ByVal riskLib As String, _
                            ByVal menus As String, ByVal prevS2Json As String, _
                            ByVal hearingAnswers As String) As String
' 15章§3 S2 user。menus=MenusSummaryFor()（preventions.related_menu_id の候補一覧。空で渡すと
' CheckS2の「実在ID」検査が構造的に落ちる）。prevS2Json=case_data の s2_prev_json、
' hearingAnswers=input_hearing_answers（いずれも初回ラウンドは "なし"）。FR-35のstatusライフサイクル
' （confirmed/rejected/new）はこの2引数がなければ成立しない
Public Function BuildS3System() As String                                        ' 15章§4 system
Public Function BuildS3User(ByVal ctx As TCaseCtx, ByVal s1Summary As String, ByVal s2Json As String, _
                            ByVal menus As String, ByVal lines As String, ByVal schemes As String, _
                            ByVal cases As String) As String
' 15章§4 S3 user。s1Summary=S1出力の business_summary / strategy_outlook / current_coverage /
' field_insights だけを抜き出した要約JSON（S3 systemルール7が field_insights の参照を命じており、
' S2 JSONには含まれないため必須）。menus=MenusFor()（S3は実在サービスの一覧。S2の要約とは別テキスト）
Public Function BuildS4System(ByVal variant As String, ByVal tier As String) As String
' 15章§5。variant=proposal / alliance（案件一覧 s4_variant）で BLOCK_S4_PROPOSAL / BLOCK_S4_ALLIANCE を
' 差し替える。tier=t1_quick / t2_full（t1=5枚固定 / t2=5〜config ppt_max_slides_t2 枚）。
' 枚数の実値は modPrompts* から config 参照でよい（案件単位の値である variant / tier は引数で受ける）
Public Function BuildS4User(ByVal ctx As TCaseCtx, ByVal s1Json As String, ByVal s2Json As String, ByVal s3Json As String) As String
Public Function BuildPFSystem() As String                                        ' 15章§6 PL-03
Public Function BuildPFUser(ByVal theme As String, ByVal body As String, ByVal rules As String, _
                            ByVal menusSummary As String, ByVal schemes As String, ByVal patterns As String, _
                            ByVal researching As String) As String
Public Function RepairSuffix(ByVal validationErrors As String) As String         ' 15章§7 修復サフィックス
Public Function SchemaS1() As String   ' 同様に SchemaS2 / S3 / S4 / S2C / S3C / PF / WT / FG（§7の表が正）
' TCaseCtx（**app層 modAppTypes**。ドメイン型なのでcore層 modTypes から移設。12章§2・§4）:
'   case_type, dossier_tier, channel, kanji, bid, reins, other_insurers, company, industry_code, industry_name

' === app: modPipeline / modPlayOps ===
Public Function RunAll(ByVal caseId As String) As Boolean
Public Function RunStep(ByVal caseId As String, ByVal stepNo As Long) As Boolean
    ' quality_mode=deep のとき、S2/S3は 生成→批判(CheckS2C/S3C合格の批判JSON)→
    ' 指摘ありなら改訂(ReviseSuffix)の3呼び出しで実行。批判・改訂はrun_logに
    ' step="s2c"/"s3c"/"s2r"/"s3r" として記録。deep_transport指定時は批判・改訂のみ経路変更
Public Function RunPreflight(ByVal inboxId As String) As Boolean

' === app: modCaseStore / modInboxStore / modJudgeStore ===
Public Function NewCase(ByVal company As String, ByVal industryCode As String, ByVal caseType As String) As String
Public Function SaveData(ByVal caseId As String, ByVal dataKey As String, ByVal content As String) As Boolean
Public Function LoadData(ByVal caseId As String, ByVal dataKey As String) As String
Public Function ResolveStepJson(ByVal caseId As String, ByVal stepNo As Long) As String
' 下流Stepが参照すべきJSONを一元解決する（優先順の正は13章§2.2）。
' N=2,3 は sN_edited > sNr_json > sN_json、S1/S4 は sN_edited > sN_json。呼び出し側で個別に分岐しない
Public Function SetStatus(ByVal caseId As String, ByVal status As String) As Boolean
Public Sub InvalidateDownstream(ByVal caseId As String, ByVal fromStepNo As Long)
Public Function RepairStates() As Long                 ' 起動時整合修復（16章E-12・12章§2.1のmodBoot手順③）。戻り=修復件数
Public Function FreezeRound(ByVal caseId As String) As Long
' ラウンド確定（FR-35マルチラウンド）。s2（edited優先で解決した1本）を data_key `s2_prev_json` へ
' 退避し、案件一覧の round_no を +1 して新しい round_no を返す。次ラウンドのS2は
' BuildS2User の prevS2Json にこの退避分を渡す
Public Function NewInboxItem(ByVal sourceKind As String, ByVal theme As String, ByVal body As String) As String
Public Function SetInboxJudgement(ByVal inboxId As String, ByVal status As String, _
                                  ByVal dropType As String, ByVal reviveTag As String, ByVal reviveDue As Date) As Boolean
Public Function NewJudgement(ByVal rec As TJudgement) As String

' === ui: modUIProgress ===
Public Sub SetStage(ByVal stepName As String, ByVal maxWaitSec As Long)
' LLM呼出の**前**に、Step名・開始時刻・最大待ち時間（maxWaitSec=config llm_wait_sec）・
' 「画面が白くなっても処理は続いている」旨を確定表示する（16章E-50(a)）。
' ScreenUpdating=False にした場合は必ずエラーハンドラで復帰させる（E-50(d)）
Public Function TryEnterUiLock(ByVal stepName As String) As Boolean
' 多重実行ガード（16章E-11）。取得できたら True。ui_lock は**modUIProgressのモジュール変数**
' （保持者Step名＋取得時のTimer値）で保持し、シートには書かない＝プロセス終了で自然消滅する（E-51）。
' 保持時間が `llm_wait_sec + 120` 秒を超えたロックは失効とみなして自動解除し、E0602 を記録して取得を許す
Public Sub ExitUiLock()   ' 正常終了・異常終了のどちらでも必ずエラーハンドラから呼ぶ
Public Sub ParkFocus()
' 全アクション完了時のフォーカス退避。フォーカスを編集不可の待避セルへ戻し、セル編集モードでVBAが
' 止まるのを防ぐ（11章§5・16章E-51）。実行開始時にも通してから処理へ入る

' === app: modExportHtml / modExportPpt / modExportHearing ===
Public Function GenerateHtmlReport(ByVal caseId As String, ByRef outPath As String) As String
    ' ""=成功 / 非空=失敗理由（コードは E0502。16章E-48）。S1+S2+S3のJSONを固定HTMLテンプレート
    ' (高橋PLプロト準拠・10章FR-37)に流し込み、自己完結HTML 1ファイルを出力(LLM不使用)。
    ' v2.3で主力出力(旧GenerateReportを置換)
    ' テンプレ本体は modHtmlTemplate1..n（純文字列・R4・12章§2 app層）、テーマCSSは modHtmlTheme
    ' （config `html_theme`。既定 standard）。出力先は config `html_out_dir`
    ' ファイル名は modUtilText.SanitizeFileName を通し、確定パスを案件一覧 report_path に記録する
    ' **文字コード**: 書き出しは `ADODB.Stream`（Charset="utf-8"・BOMあり）。VBAの Open/Print # は
    ' CP932で書かれ非CP932文字が "?" 化するため使わない。テンプレ先頭に <meta charset="utf-8"> を必ず含める
    ' **埋め込み**: JSONは「1本のJS文字列リテラル＋JSON.parse」形式で埋め、modUtilText.JsStringSafe を
    ' 必ず通す。素のJSリテラル直書きは禁止。HTML本文に差し込む値は HtmlSafe を通す（16章E-47）
Public Function GeneratePpt(ByVal caseId As String, ByVal s4Json As String, _
                            ByVal variant As String, ByRef outPath As String) As String ' ""=成功。variant=proposal/alliance。Phase 1.5
Public Function BuildHearingSheet(ByVal caseId As String) As Boolean

' === test: modMockLlm（本体内mockトランスポート。§4(a)・12章§2 test層） ===
Public Function MockResponse(ByVal stepName As String, ByVal variantName As String, _
                             ByVal fault As String) As String
' modGatewayRPN のmock分岐が呼ぶ**ゲートウェイ入口**（12章§2の「唯一の相手」の実体）。
' variantName は modGatewayRPN.ResolveMockVariant の戻り値（new/renewal/hit/clean/common）、
' fault は config `mock_fault`。内部は下の2本へ振り分けるだけで、独自の応答本文を持たない
Public Function ResponseById(ByVal mockId As String) As String
' 15章§8.1の表の mock ID（MK-S1-NEW / MK-S1-RNW / MK-S2-NEW / MK-S2-RNW / MK-S3 / MK-S4 /
' MK-PF / MK-S2C-HIT / MK-S2C-CLEAN / MK-S3C-HIT / MK-S3C-CLEAN の**11 ID**）で正常応答を返す。
' 表に無いIDは ""。決定的（乱数・現在時刻・呼び出し回数に依存しない）
Public Function FaultResponse(ByVal faultKind As String, ByVal stepName As String) As String
' 15章§8.2の**8値**（broken_json / broken_json_once / enum_violation / ghost_id / count_violation / empty /
' limit / fake_err）に対応する障害注入応答。**状態レス**＝同じ引数なら常に同じ応答を返す
' （「最初の1回だけ」型の内部カウンタを持たない）。faultKind が空のときは "" を返す。
' 8値以外の未知の値は正常応答へフォールバックする（config入力ミスでE2Eを暴走させない）

Public Sub ResetFaultOnce()                                  ' broken_json_once の状態リセット(15章§8.2。テスト・T-24のシナリオ冒頭で呼ぶ)
' === test: modTestsExcel ===
Public Sub RunAllExcelTests()   ' 層(b)=Excel固有E2Eスモークの入口(12章§2 test層・17章 T-47)。
                                ' wintest実機(実Excel・COM経由)からのみ呼ぶ。層(a)の入口は
                                ' modTestRunner.RunAllPureTests(17章§4-1のランナー要件)
```

- **`modHtmlTemplate1..n` / `modHtmlTheme` の関数契約（`BuildDocument` / `HeadHtml` / `BodyShellHtml` / `SectionsJs` / `RuntimeJs` / `ThemeCss` / `ThemeNames` 等）は18章§4.4・§5.2が正**（本章は宣言を持たない。追加・分割の規約も18章に従う）
- **`modValidate` の CheckS2C / CheckS3C**、**`modSchemas` の SchemaS2C / SchemaS3C** は入念モード用の追加分（15章§4.5〜4.6・§7の表）
- 呼出前の走査: 外部へ送るテキスト（CallStep / CallChat の systemPrompt・userPrompt、企業ドシエファイルの書出、HTMLレポート出力）は送信・保存の直前に `modPii` を通す（16章 E-05／E-31。走査結果は run_log と dossier_meta に記録）

## 7. スキーマ・レジストリ（modSchemas。本文は15章）

**Const は使わない。すべて引数なしの純関数で返す**（`Public Function SchemaS1() As String` を `s = s & "..." & vbLf` で組み立てる）。理由: VBAの Const は1論理行1023字・行継続 `_` 25本までで、3,700字級のスキーマ本体を1宣言に収められず、1行追加した瞬間に壊れるため。旧表記の `SCHEMA_S1` 等は使わず関数名に統一する（19章§4のスキーマ・レジストリもこの関数名で読む）。

| 関数名 | 対応step | strict検証済み観点 |
|---|---|---|
| `SchemaS1()` | s1 | current_coverage は常に必須（newは空配列） |
| `SchemaS2()` | s2 | gaps は常に必須（newは空配列）。リスクユニバース10分類/頻度/影響/1〜5スコア/移転可能性/status/出所enum（v2.3）。**emerging_risks（ニューリスク0〜3件・空配列可・category/horizon/出所enum）を含む（v2.4）** |
| `SchemaS3()` | s3 | proposal_kind enum。scheme_id は "" 許容 |
| `SchemaS4()` | s4 | slides配列・hearing_questions |
| `SchemaS2C()` | s2c | 入念モードの批判JSON。issue_type 6値のenum。issues は0件（指摘なし）を許容 |
| `SchemaS3C()` | s3c | 入念モードの批判JSON。executive_reactions はちょうど3件。issue_type 6値のenum |
| `SchemaPF()` | pf | 5問判定・文法4値・予測類型・組み替え案 |
| `SchemaWT()` | wt | 分類enum・pattern_id |
| `SchemaFG()` | fg | 格付enum・文法4bool |

- s2r / s3r（入念モードの改訂）はスキーマを新設せず `SchemaS2()` / `SchemaS3()` を再利用する（§1 Stepレジストリ）
- 1モジュール30,000字契約に収まらない場合は modSchemas を `modSchemas1..n` へ分割してよい（関数名は変えない）
