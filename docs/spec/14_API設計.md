# 14. API設計（LLM呼び出し仕様と内部インターフェース契約）

## 1. LLM経路の全体方針

唯一の窓口は `modGatewayRPN.CallStep()`。呼び出し元（modPipeline）は経路を意識しない。

```
modPipeline ──> modGatewayRPN.CallStep(stepName, systemPrompt, userPrompt, schemaJson) ──┬─ ribbon（既定）
                                                                                         ├─ direct
                                                                                         └─ mock
```

分岐規則（PoC互換）: `mock_llm=TRUE` → mock。それ以外は `llm_transport`（ribbon/direct/mock）。ribbon選択時にリボン未検出なら **E0201 を返して停止**（directへの自動フォールバックはしない。キーの無い一般利用者PCで誤ってdirectに落ちて失敗する事故を防ぐ）。

## 2. ribbon経路（主経路・本番）

社内AIリボン「リボンちゃん」の確定API（PoC `docs/dev/RIBBON_API_CONFIRMED.md` が一次情報）に従う。

```vb
result = Application.Run("ChatGPT", _
    userPrompt,               ' 1 Text
    systemPrompt,             ' 2 roleSystem ← Stepのsystemプロンプトを渡す
    0.3,                      ' 3 Temperature (Double。GPT-5系では無視される)
    0,                        ' 4 MaxTokens (Long。0=既定)
    waitSec,                  ' 5 Wait (config llm_wait_sec 既定1200)
    model,                    ' 6 config recommended_model 既定 gpt-5.5
    "",                       ' 7 prevU（会話継続。RPNは各Step独立のため常に""）
    "",                       ' 8 prevA（同上）
    "リスク提案ナビ:" & stepName, ' 9 toolN（管理側ログでのツール識別。必須）
    effort,                   '10 config reasoning_effort（GPT-5系）
    verbosity)                '11 config reasoning_verbosity
```

- 戻り値: Azure OpenAIの応答テキストがそのまま返る（ラッパー側の加工なし。上限エラー等もそのまま文字列）
- **response_format は指定できない** → JSONはプロンプトで強制し（15章）、後段の `modJsonLite` + `modValidate` + 修復リトライで担保（§5）
- アドイン検出: `Application.AddIns` ループで `InStr(name, config ribbon_addin_name)>0 And Installed`（セッションキャッシュ）
- 起動時に `LimitCheck()`（True=続行不可）を1回呼ぶ（config limit_check で無効化可）。続行不可時は案内を出し、実行ボタン押下時に再案内
- エラー分類: 空応答=E0202 / 上限系文字列（PoC `LooksLikeLimitError` 移植）=E0204 / アドイン無し=E0201。戻り値は `"#ERR:E02xx:説明"` 形式（例外は投げない。PoC規約）

## 3. direct経路（開発・検証用）

OpenAI API を `MSXML2.ServerXMLHTTP.6.0` で直叩き。PoC modGatewayDirect の骨格（setTimeouts・ステータス表示・構造化ログ）を流用し、embeddings→chat/completions に書き換える。

### リクエスト

```
POST {direct_api_base}/chat/completions
Content-Type: application/json
Authorization: Bearer {key}          ← config direct_key_path のファイルから読む（§6）
```

```json
{
  "model": "gpt-4.1",
  "temperature": 0.3,
  "messages": [
    {"role": "system", "content": "<systemPrompt>"},
    {"role": "user",   "content": "<userPrompt>"}
  ],
  "response_format": {
    "type": "json_schema",
    "json_schema": {"name": "<step名>", "strict": true, "schema": { ...15章のSchema... }}
  }
}
```

- **Structured Outputs (strict:true)** を必ず使う。15章のスキーマはstrict要件（全プロパティrequired・additionalProperties:false）を満たして定義してある
- レスポンスは `choices[0].message.content` を modJsonLite で抽出（`"content"` キーの文字列値抽出。エスケープ解除は modJsonLite.UnescapeJsonStr）
- `refusal` 非空 / `finish_reason != "stop"` はエラー扱い（E0206）
- タイムアウト: `setTimeouts 5000, 10000, {direct_http_timeout_ms}, {direct_http_timeout_ms}`
- リトライ: HTTP 429/500/502/503 は指数バックオフで最大3回（2s→4s→8s、`Application.Wait`）。408/タイムアウトは1回。4xx（429除く）はリトライしない
- o系モデル指定時は temperature を送らない（パラメタ非対応）。モデル名先頭 "o" で分岐する簡易判定を modGatewayDirect 内に置く

### コスト試算（direct経路・参考値）
1案件 = 4呼び出し、入力合計 ≈ 25k tokens・出力合計 ≈ 8k tokens。gpt-4.1相当の単価で **1案件あたり数十円のオーダー**。PoC規模（10社×試行3回）で数千円以内。ribbon経路は社内利用枠のため追加費用なし。

## 4. mock経路（社外開発・デモ用）

- `wintest/mock_ribbon/modMockRibbon.bas` 方式: 各Stepに対して**決定的**なサンプルJSON（春華堂を模した架空企業データ）を返す
- 決定性: 同じ入力（正規化後ハッシュ）→同じ応答。乱数不使用（PoC規約）
- 用途: 画面フロー確認・PPT生成テスト・ユニットテスト・自宅Macでの開発

## 5. JSON信頼性の担保（両経路共通の防衛線）

```
LLM応答文字列
 → (1) modJsonLite.ExtractJsonBlock   前後の説明文・```json フェンスを剥がし最外{ }を取り出す
 → (2) modJsonLite.ParseObject        軽量パース（対象スキーマに必要なキーのみ抽出。汎用パーサは作らない）
 → (3) modValidate.CheckS{n}          必須キー・型・enum・件数・ID実在（S3のmenu_id/line_id）
 → NG時: (4) 修復リトライ（config json_repair_retry=1 回）
          「あなたの直前の出力は次の検証エラーで不合格でした: <エラー列挙>。
           同じ内容を、指示したJSON形式のみで（説明文なしで）再出力してください。」
          を追記して同Stepを再呼び出し
 → なおNG: E03xx でStep失敗。案件状態=error、S{n}シートにエラー内容と「入力を短くする/再実行」の対処を表示
```

- direct経路はstrictにより(1)(2)はほぼ素通しになるが、**(3)の業務検証（menu_id実在等）はスキーマでは表現できないため両経路で必須**
- 検証エラーは err_log(E03xx) と run_log(validate_result) に記録。修復で直った場合も `repaired` と記録し、プロンプト改善のシグナルにする

## 6. 認証・キー管理（direct経路のみ）

- キーは `config direct_key_path`（既定 `%APPDATA%\RPN\api_key.txt`、環境変数展開対応）のテキストファイル1行目から読む
- **ブック内（configシート含む）にキーを保存することを禁止**。modConfig にキーを持たせない（modGatewayDirect がファイルを直接読む）
- ファイル不存在時: E0205「APIキーファイルがありません（direct経路は開発者専用です）」を表示し停止
- ログ・画面にキーを一切出力しない（err_log の detail にリクエストヘッダを含めない）

## 7. 内部インターフェース契約（実装対象の公開関数シグネチャ）

エラー規約: core/app層の関数は**例外を投げず**、String戻り値は `"#ERR:Exxxx:メッセージ"`、Boolean戻り値は False＋modLog記録（PoC規約踏襲）。

```vb
' === modGatewayRPN ===
Public Function CallStep(ByVal stepName As String, ByVal systemPrompt As String, _
                         ByVal userPrompt As String, ByVal schemaJson As String, _
                         Optional ByRef latencyMs As Long = 0) As String
    ' 経路分岐・呼出・#ERR規約。schemaJsonはdirect経路のresponse_formatにのみ使用
Public Function RibbonAvailable() As Boolean
Public Function RunLimitCheck() As Boolean   ' True=続行不可

' === modJsonLite ===
Public Function ExtractJsonBlock(ByVal raw As String) As String          ' 失敗時 ""
Public Function GetStr(ByVal json As String, ByVal key As String) As String
Public Function GetLong(ByVal json As String, ByVal key As String, ByVal dflt As Long) As Long
Public Function GetArrayItems(ByVal json As String, ByVal key As String) As Collection ' 各要素のJSON文字列
Public Function EscapeJsonStr(ByVal s As String) As String
Public Function UnescapeJsonStr(ByVal s As String) As String

' === modValidate ===  戻り値: "" =合格 / 非空=エラー列挙（修復プロンプトへ渡す日本語文）
Public Function CheckS1(ByVal json As String) As String
Public Function CheckS2(ByVal json As String) As String
Public Function CheckS3(ByVal json As String, ByVal kb As Object) As String ' kb=modKnowledgeのID集合
Public Function CheckS4(ByVal json As String) As String

' === modKnowledge ===
Public Function LoadKnowledge() As Boolean                    ' 起動時/再読込。内部キャッシュへ
Public Function RiskLibFor(ByVal industryCode As String) As String   ' 注入用テキスト整形済（上限kb_risk_rows）
Public Function MenusFor(ByVal industryCode As String) As String     ' 同（上限kb_menu_rows・is_active=TRUEのみ）
Public Function CasesFor(ByVal industryCode As String) As String     ' 同（上限kb_case_rows）
Public Function MenuIdExists(ByVal menuId As String) As Boolean
Public Function LineIdExists(ByVal lineId As String) As Boolean
Public Sub AppendServiceGap(ByVal caseId As String, ByVal industryCode As String, ByVal riskDesc As String)

' === modPromptsRPN ===（純文字列・R4。configアクセスのみ許可）
Public Function BuildS1System() As String
Public Function BuildS1User(ByVal company As String, ByVal industryName As String, _
                            ByVal hpText As String, ByVal yuhoText As String, ByVal memoText As String) As String
Public Function BuildS2System() As String
Public Function BuildS2User(ByVal s1Json As String, ByVal riskLibText As String) As String
Public Function BuildS3System() As String
Public Function BuildS3User(ByVal s2Json As String, ByVal menusText As String, _
                            ByVal linesText As String, ByVal casesText As String) As String
Public Function BuildS4System() As String
Public Function BuildS4User(ByVal s1Json As String, ByVal s2Json As String, ByVal s3Json As String) As String
Public Function BuildRepairSuffix(ByVal validationErrors As String) As String
Public Function SchemaS1() As String   ' 15章のスキーマJSON文字列（direct経路用・定数分割格納）
Public Function SchemaS2() As String
Public Function SchemaS3() As String
Public Function SchemaS4() As String

' === modPipeline ===
Public Function RunAll(ByVal caseId As String) As Boolean
Public Function RunStep(ByVal caseId As String, ByVal stepNo As Long) As Boolean
    ' 前提状態チェック→プロンプト組立→CallStep→検証→保存→描画→状態更新→下流無効化

' === modCaseStore ===
Public Function NewCase(ByVal company As String, ByVal industryCode As String) As String ' 戻り=case_id
Public Function SaveData(ByVal caseId As String, ByVal dataKey As String, ByVal content As String) As Boolean
Public Function LoadData(ByVal caseId As String, ByVal dataKey As String) As String  ' 分割の透過結合
Public Function SetStatus(ByVal caseId As String, ByVal status As String) As Boolean
Public Sub InvalidateDownstream(ByVal caseId As String, ByVal fromStepNo As Long)

' === modExportPpt ===
Public Function GeneratePpt(ByVal caseId As String, ByVal s4Json As String, _
                            ByRef outPath As String) As String  ' ""=成功

' === modExportHearing ===
Public Function BuildHearingSheet(ByVal caseId As String) As Boolean
```
