# 14. API設計（LLM呼び出し仕様と内部インターフェース契約）v2.0

## 1. LLM経路と分岐

唯一の窓口は `modGatewayRPN.CallStep()`。呼び出し元（modPipeline / modPlayOps）は経路を意識しない。

- 分岐規則（PoC互換）: `mock_llm=TRUE` → mock。それ以外は `llm_transport`（ribbon/direct/mock）
- ribbon選択時にリボン未検出 → **E0201で停止**（directへ自動フォールバックしない。一般利用者PCでの誤課金・誤失敗防止）
- 全呼び出しは run_log に1行記録（step, play, transport, model, latency, validate_result, injected_kb_ids）

### Stepレジストリ（呼び出し単位の正）

| step | play | スキーマ | 呼び出し元 |
|---|---|---|---|
| s1/s2/s3/s4 | PL-01/02 | Schema-S1/S2/S3/S4 | modPipeline |
| pf | PL-03 | Schema-PF | modPlayOps |
| wt | PL-05 | Schema-WT | modPlayOps(1.5) |
| fg | PL-06 | Schema-FG | modPlayOps(1.5) |

## 2. ribbon経路（主経路・本番）

確定台帳（PoC `docs/dev/RIBBON_API_CONFIRMED.md`）準拠の12引数呼び出し:

```vb
result = Application.Run("ChatGPT", _
    userPrompt, systemPrompt, 0.3, 0, waitSec, model, _
    "", "", "リスク提案ナビ:" & stepName, effort, verbosity)
'  1 Text  2 roleSystem  3 Temp(Double)  4 MaxTokens(Long,0=既定)  5 Wait
'  6 model(config recommended_model)  7 prevU  8 prevA(各Step独立のため常に"")
'  9 toolN(管理側ログ識別・必須)  10 effort  11 verbosity(GPT-5系)
```

- 戻り値はAzure OpenAI応答の素通し。response_format指定不可 → JSONは15章のプロンプト強制＋§5の防衛線で担保
- アドイン検出: `Application.AddIns` ループ（`ribbon_addin_name` 部分一致＋Installed、セッションキャッシュ）
- 起動時 `LimitCheck()`（True=続行不可→案内し、実行時に再案内。config limit_check で無効化可）
- 温度・MaxTokensはGPT-5系では無視され effort/verbosity が効く（V2実運用で確認済み）。`reasoning_tuning` エスケープハッチはPoC同様に維持
- エラー: 空応答=E0202／上限系文字列（LooksLikeLimitError移植）=E0204／アドイン無し=E0201。戻り値 `"#ERR:E02xx:説明"`（例外は投げない）

## 3. direct経路（開発・検証用）

`MSXML2.ServerXMLHTTP.6.0`、`setTimeouts 5000,10000,{t},{t}`（t=direct_http_timeout_ms）。

```
POST {direct_api_base}/chat/completions
Authorization: Bearer {keyファイル1行目}   ※ブック・config・ログに保存禁止
{
  "model": "{direct_model}", "temperature": 0.3,
  "messages": [{"role":"system","content":sys},{"role":"user","content":usr}],
  "response_format": {"type":"json_schema",
    "json_schema":{"name":"{step}","strict":true,"schema":{...modSchemasの該当スキーマ...}}}
}
```

- **strict:true必須**。15章の全スキーマはstrict要件（全プロパティrequired・additionalProperties:false・enum統制）を満たす
- 応答は `choices[0].message.content` を modJsonLite で抽出。`refusal`非空 / `finish_reason≠"stop"` は E0206
- リトライ: 429/500/502/503=指数バックオフ最大3回（2s/4s/8s）。408/タイムアウト=1回。その他4xx=リトライなし
- o系モデル名（先頭"o"）では temperature を送らない
- キー不存在=E0205「direct経路は開発者専用です」
- コスト目安: 1案件=4呼び出し・入力≈25k tok・出力≈8k tok → 数十円/案件。PoC全体で数千円以内

## 4. mock経路

modMockRibbon が step別に**決定的**なサンプルJSON（15章§8「浜松スイーツファクトリー」一式＋PF応答）を返す。乱数・現在時刻不使用。同一入力（正規化ハッシュ）→同一応答。

## 5. JSON防衛線（両経路共通・全step）

```
raw → (1) modJsonLite.ExtractJsonBlock（説明文・```フェンス除去・最外{}）
    → (2) modJsonLite で対象スキーマに必要なキーのみ抽出（汎用パーサは作らない）
    → (3) modValidate.Check{S1..S4|PF|WT|FG}（型・enum・件数・ID実在・整合）
    → NG: (4) 修復リトライ（json_repair_retry=1回。15章§7のサフィックスを user末尾に追記して同stepを再呼び出し）
    → なおNG: E03xxでstep失敗。生応答は case_data/pf_json に保存済み
```
- directはstrictで(1)(2)がほぼ素通しになるが、**(3)の業務検証（ID実在・件数・整合）は両経路で必須**（スキーマでは表現できない）
- validate_result（ok/repaired/failed）を run_log に記録し、月次でプロンプト改善のシグナルにする

## 6. 内部インターフェース契約（公開関数シグネチャ）

エラー規約: 例外を投げない。String戻り値は `"#ERR:Exxxx:メッセージ"`、Boolean戻り値は False＋modLog記録。

```vb
' === core: modGatewayRPN ===
Public Function CallStep(ByVal stepName As String, ByVal playId As String, _
                         ByVal systemPrompt As String, ByVal userPrompt As String, _
                         ByVal schemaJson As String, Optional ByRef latencyMs As Long = 0) As String
Public Function RibbonAvailable() As Boolean
Public Function RunLimitCheck() As Boolean            ' True=続行不可
' 壁打ち(PL-04)専用: リボンの会話継続引数(prevU/prevA)を使う唯一の関数。
' 履歴は「新しい順」に ";;;" 区切りで連結して渡す(PoC 裁定D11の実証方式)。
' 渡す履歴は直近 sparring_max_turns 往復まで。JSONスキーマは使わない(自由対話)。
Public Function CallChat(ByVal caseId As String, ByVal systemPrompt As String, _
                         ByVal userMsg As String, ByVal histU As String, ByVal histA As String, _
                         Optional ByRef latencyMs As Long = 0) As String

' === core: modJsonLite ===
Public Function ExtractJsonBlock(ByVal raw As String) As String        ' 失敗時 ""
Public Function GetStr(ByVal json As String, ByVal key As String) As String
Public Function GetLong(ByVal json As String, ByVal key As String, ByVal dflt As Long) As Long
Public Function GetBoolJ(ByVal json As String, ByVal key As String, ByVal dflt As Boolean) As Boolean
Public Function GetArrayItems(ByVal json As String, ByVal key As String) As Collection
Public Function EscapeJsonStr(ByVal s As String) As String
Public Function UnescapeJsonStr(ByVal s As String) As String

' === app: modValidate ===  戻り値 ""=合格 / 非空=エラー列挙（修復プロンプト用の日本語）
Public Function CheckS1(ByVal json As String, ByVal caseType As String) As String
Public Function CheckS2(ByVal json As String, ByVal caseType As String) As String
Public Function CheckS3(ByVal json As String, ByVal s2Json As String) As String  ' ID実在はmodKnowledge参照
Public Function CheckS4(ByVal json As String) As String
Public Function CheckPF(ByVal json As String) As String
Public Function CheckWT(ByVal json As String) As String          ' Phase1.5
Public Function CheckFG(ByVal json As String) As String          ' Phase1.5

' === app: modKnowledge ===
Public Function LoadKnowledge() As Boolean            ' 起動時/再読込。スナップショット保存込み
Public Function RiskLibFor(ByVal industryCode As String) As String     ' 整形済注入テキスト
Public Function MenusFor(ByVal industryCode As String) As String
Public Function LinesText() As String
Public Function CasesFor(ByVal industryCode As String) As String
Public Function SchemesFor(ByVal industryCode As String) As String     ' status∈{proven,adopted}のみ
Public Function PatternsText() As String                                ' P1-P15全件（PF/FG用）
Public Function RulesText() As String                                   ' 判断基準（PF/FG用）
Public Function ResearchingText() As String                             ' 研究テーマ一覧（PF用）
Public Function MenuIdExists(ByVal id As String) As Boolean
Public Function LineIdExists(ByVal id As String) As Boolean
Public Function SchemeIdExists(ByVal id As String) As Boolean
Public Function CaseLibIdExists(ByVal id As String) As Boolean
Public Function PatternIdExists(ByVal id As String) As Boolean
Public Sub AppendServiceGap(ByVal caseId As String, ByVal industryCode As String, ByVal riskDesc As String)

' === app: modPromptsCore / modPromptsBlocks / modPromptsOps / modSchemas ===
' 15章と一字一句一致。シグネチャ:
Public Function BuildS1System() As String
Public Function BuildS1User(ByVal ctx As TCaseCtx, ByVal hp As String, ByVal yuho As String, _
                            ByVal memo As String, ByVal contractTxt As String, ByVal prevRenewal As String, _
                            ByVal dossierTxt As String) As String
Public Function BuildSparringSystem(ByVal dossierSummary As String, ByVal s1s2s3Json As String, _
                                    ByVal schemes As String, ByVal patterns As String, _
                                    ByVal mechs As String, ByVal rules As String) As String  ' PL-04
Public Function BuildS2System() As String
Public Function BuildS2User(ByVal ctx As TCaseCtx, ByVal s1Json As String, ByVal riskLib As String) As String
Public Function BuildS3System() As String
Public Function BuildS3User(ByVal ctx As TCaseCtx, ByVal s2Json As String, ByVal menus As String, _
                            ByVal lines As String, ByVal cases As String, ByVal schemes As String) As String
Public Function BuildS4System() As String
Public Function BuildS4User(ByVal ctx As TCaseCtx, ByVal s1Json As String, ByVal s2Json As String, ByVal s3Json As String) As String
Public Function BuildPFSystem() As String
Public Function BuildPFUser(ByVal theme As String, ByVal body As String, ByVal rules As String, _
                            ByVal menusSummary As String, ByVal schemes As String, ByVal patterns As String, _
                            ByVal researching As String) As String
Public Function RepairSuffix(ByVal validationErrors As String) As String
Public Function SchemaS1() As String   ' 同様に S2/S3/S4/PF/WT/FG
' TCaseCtx（modTypes）: case_type, dossier_tier, channel, kanji, bid, reins, other_insurers, company, industry_code, industry_name

' === app: modPipeline / modPlayOps ===
Public Function RunAll(ByVal caseId As String) As Boolean
Public Function RunStep(ByVal caseId As String, ByVal stepNo As Long) As Boolean
Public Function RunPreflight(ByVal inboxId As String) As Boolean

' === app: modCaseStore / modInboxStore / modJudgeStore ===
Public Function NewCase(ByVal company As String, ByVal industryCode As String, ByVal caseType As String) As String
Public Function SaveData(ByVal caseId As String, ByVal dataKey As String, ByVal content As String) As Boolean
Public Function LoadData(ByVal caseId As String, ByVal dataKey As String) As String
Public Function SetStatus(ByVal caseId As String, ByVal status As String) As Boolean
Public Sub InvalidateDownstream(ByVal caseId As String, ByVal fromStepNo As Long)
Public Function RepairStates() As Long                 ' 起動時整合修復（E-12）。戻り=修復件数
Public Function NewInboxItem(ByVal sourceKind As String, ByVal theme As String, ByVal body As String) As String
Public Function SetInboxJudgement(ByVal inboxId As String, ByVal status As String, _
                                  ByVal dropType As String, ByVal reviveTag As String, ByVal reviveDue As Date) As Boolean
Public Function NewJudgement(ByVal rec As TJudgement) As String

' === app: modExportPpt / modExportHearing ===
Public Function GeneratePpt(ByVal caseId As String, ByVal s4Json As String, ByRef outPath As String) As String ' ""=成功
Public Function BuildHearingSheet(ByVal caseId As String) As Boolean
```

## 7. スキーマ・レジストリ（modSchemas。本文は15章）

| 定数名 | 対応step | strict検証済み観点 |
|---|---|---|
| SCHEMA_S1 | s1 | current_coverage は常に必須（newは空配列） |
| SCHEMA_S2 | s2 | gaps は常に必須（newは空配列）。6カテゴリ/頻度/影響/出所enum |
| SCHEMA_S3 | s3 | proposal_kind enum。scheme_id は "" 許容 |
| SCHEMA_S4 | s4 | slides配列・hearing_questions |
| SCHEMA_PF | pf | 5問判定・文法4値・予測類型・組み替え案 |
| SCHEMA_WT | wt | 分類enum・pattern_id |
| SCHEMA_FG | fg | 格付enum・文法4bool |
