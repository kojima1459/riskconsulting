# 14. API設計（LLM呼び出し仕様と内部インターフェース契約）v2.5.4

> v2.5.6（裁定書17 裁定A: W5.2ホットフィックス）: §6の登記表へ本裁定で許可した新設名を登記した。**`modUIToast`**（`ShowToast` / `ShowNext` / `HideToast` / `CancelToast` / `WarnLine`）・Shape接頭辞 **`ts_`**・**`modBoot.KbAutoNote`**（ナレッジブック自動発見の結果の読み出し口。探索本体は modBoot の Private `ResolveKbPath`）・`modUISheet.EnsureButtonEx` の `kind="primary"` を高さ30ptとする寸法規約・**`modUIToast.ShowResearchPrompts`** と名前付きレンジ **`gd_ch7_head`**・HOMEの主要動線5本の新キャプションと図形名 `btn_hm_step1`～`btn_hm_step5`（司令塔追補）。**本波で許可した新設はこの範囲のみ**（公開シグネチャの変更は無い）。

> v2.5.5（裁定書14: W5 社内環境対応＋UI/UX改善）: §6の登記表へ本裁定で許可した新設名を登記した。**`modUIGuide`**（初回ガイドツアー。`StartTourIfFirstRun` / `RestartTour` / `OnTourNext` / `OnTourSkip` / `ClearTour` / `EnsureGuideButtons`）・**`modTestsRunnerUi.RunAllTestsFromBook`**（ブック内テスト実行＝17章 T-48）・**`modTestRunner.PassCount` / `FailCount` / `SkipCount` / `ExecutedCount`**（計数の読み出し口）・**`modUISheet.EnsureButtonEx` / `CellLeft`**（種別つき図形ボタンと幾何計算の読み口）・config キー **`guide_tour_done`**・Shape接頭辞 **`gt_`**・名前付きレンジ **`gd_btn_tour`**・HOMEの主要動線4本の新キャプション（[① 案件を作る]／[② 一括実行]／[③ レポートを出す]／[④ ヒアリングシート]）。**本波で許可した新設はこの範囲のみ**。
>
> v2.5.4（裁定書13: W4.5 クローズパッチ）: §6の登記表へ **`modUICase4.ClearCaseInput`**（W1。案件入力の全クリア。HOMEの[＋新規案件]が新規モードのマーカーを書く前に呼ぶ唯一の口。30,000字契約により `modUICase3` へ置けないため `CopyResearchRow` と同じ移設先へ置く）を登記した。**本波で許可した新設はこの1件のみ**。

> v2.5.3（裁定書12: W4.4 最終収束パッチ）: §6の登記表へ **`modUICase3.U3_NEW_MARK`**（V1。案件入力の新規モードの固定マーカー `(新規)`。`modInboxStore.IB_DRAFT_MARK` と同作法の公開定数）を登記した。**本波での公開関数の新設・移設はゼロ**。あわせて `AnswerMemoCount` の別案件確認文言を実物へ逐語化した（V9）。

> v2.5.2（裁定書11: W4.3 最終パッチ）: §6の登記表へ新設・移設3件を登記した。**`modUICase.RebindFlatValidation`**（Q3(a)）・**`modUICase4.CopyResearchRow`**（Q1。30,000字契約により `modUICase3` から移設）・**`modTestsExcel2.RunExcelTests2`**（Q9/Q1。層(b)の分割先。wintest からの入口は `RunAllExcelTests` の1本のまま）。**本裁定で許可した新設・移設はこの3件のみ**。

> v2.5.1（裁定書10: W4.2 収束ウェーブ）: §6へ **`modPipeline2.ResetDeepOutcome`**（N9。deep outcome の明示リセット口。M1により `RunStep` 冒頭の自動リセットを廃止し、`LastDeepOutcome` は「その実行で最後に立った非空 outcome」を返す契約へ改めた）を新設し、**`modUICase5` の Public 4本**（`SerializeBody` / `ColIndexes` / `ColCount` / `RoomOf`。W4.1の30,000字契約分割の追認）を「ui層内部ヘルパ」として登記した。N6（`ib_body_draft`）は**廃止**し「受信箱の投函下書き行」方式（13章§2.6・C1）へ差し替え。あわせて `AnswerMemoCount` を「常に数える」契約へ（M5。別案件判定は呼出側が `hs_case_id` で行う）、`LoadData` へ仮seq帯の残留検査（M6）を追記した。

> v2.5（裁定書9: W4.1 最終修正ウェーブ）: §6へ新設3本を宣言した。**`modPipeline2.LastDeepOutcome`**（N1。E-35/E-36 の警告を ui層へ渡す唯一の口）・**`modCaseStore.PromoteTier`**（N2。案件一覧 `dossier_tier` の唯一の書込口。v2.4.7 が「本節の裁定事項」と書いた未解決(a)の解消）・**`modExportHearing.AnswerMemoCount`**（N8。手書き回答の上書き確認の要否判定）。あわせて `modCompanyFile2.DossierSaveAndClose` を Boolean へ改め（N3）、`SetStatus` が遷移検査を通さない設計を明記し、`exported` / `feedback_done` の結線先・`AppendServiceGap` の呼出点・`MenuIdExists` 系5本の二次照合の呼出点・企業ドシエのファイル名8桁を company 由来とする例外（13章§2.8）を注記した。**本裁定で許可した新設名は N1～N8 の8件のみ**であり、これ以外の公開関数・名前付きレンジを新設しない。あわせて §6 の末尾へ**名前付きレンジ・図形ボタン・入力列の登記表**（N4～N7）を新設し、本章が関数名だけでなく名前全体の唯一の正であることを表に固定した。

> v2.4.8（裁定書9-1/9-2: W2c検証 MAJOR の解消）: 規約が実行制御・シートI/Oの中に閉じ込められていた2点を**純核として宣言**した。(1) **`modInboxStore.InterestSummaryOf(themeLines)`** - 10章FR-17 の関心度集計（件数集計・件数降順・**2件以上集まったテーマだけ**・上限は既定3件）の唯一の値源。シートI/Oの `InterestText` は theme 列を1件1行で集めて渡すだけになり、上限件数を引数で受けなくなったので **`InterestText()` は引数なし**へ改めた（`maxItems` の呼び出し実績は無い）。(2) **`modPipeline2.AdoptRevisionOf(outcome, originalJson, revisedJson)`** - 16章E-36 の「改訂を破棄して改訂前を採用」の唯一の選択点。`RunPipe` の確定JSON選択と `sNr_json` の保存可否はこの戻り値を経由し、不合格の改訂版が `sNr_json` へ入る経路を構造として持たない。あわせて §6 が「Private へ戻すことは契約違反」と書く純核（`modPipeline2` 8本 ＋ `modInboxStore` / `modPlayOps` / `modJudgeStore` の宣言済み純核）を `vba_lint.py` の CONTRACT `required` へ同期した（裁定書9-3）。

> v2.4.7（裁定書8 B-9: T-27 実装時の命名）: §6へ **`modSparring`**（PL-04 壁打ち）の節を新設し、実行制御4本（`ResumeSparring` / `SendSparring` / `SendToInbox` / `HistoryOf`）と純核3本（`CanContinueSparring` / `TrimHistoryOf` / `HistoryJoinOf`）を宣言した。あわせて `sparring_u` / `sparring_a`（13章§2.2）の**保存形式**（1発話＝1行 `seq <TAB> spoke_at <TAB> 本文`）を本節で確定した--§2.2 の列定義（`seq`＝32,000字の分割連番）と data_key 注記（「発話単位seqで保存」）の食い違いを、`SaveData` の契約（data_key 単位で全行を置換）を変えずに吸収するためである。**未解決2件**を本文中に明記した: (a) 13章§2.1/§2.17 が求める `dossier_tier` の t3_sparring への**自動昇格の書込口**が本節に無い（`ResumeSparring` は昇格せず usage_log に事実を残す）、(b) 15章§6.5 の `{{schemes}}` は「全status」だが `modKnowledge.SchemesFor` は S3用の proven/adopted 絞込しか持たない（狭い側で注入し run_log へ事実を残す）。
>
> v2.4.6（裁定書8 B-8: T-26 実装時の命名）: 裁定書8が予約していた **`modJudgeStore.NewJudgement`**（`TJudgement`受取・`judge_id`返却）に加え、読取口 **`ReadJudgement`**・事後結果の更新口 **`SetJudgementResult`**・純ロジック4本（`BuildJudgeId` / `IsValidJudgeId` / `IsValidDecision` / `IsValidJudgeResult`）を宣言した。あわせて `modAppTypes` へ判断台帳9列（`judge_id`/`judged_at`を除く）の入れ物 **`TJudgement`** を新設した。判断台帳は13章§4のとおり削除しないため、本モジュールに Delete 相当の公開関数は無い。
>
> v2.4.5（裁定書8 B-7/B-10: T-25・T-28 実装時の命名）: 裁定書8 A-1 が「T-28 の実装時に本節へ足す」と予告していた **`modPipeline2` の判定核7本**（`CritiqueStepOf` / `ReviseStepOf` / `NeedsRevision` / `CritiqueDigest` / `DeepOutcomeOf` / `DeepWarningOf` / `DeepRouteOf`）を宣言した。あわせて T-25 の **`modInboxStore`**（シートI/O4本 `SavePfResult` / `ReadInboxItem` / `UndiagnosedIds` / `InterestText` ＋ 純ロジック6本 `BuildInboxId` / `IsValidInboxId` / `CanInboxTransition` / `JudgementError` / `InterestKeyOf` / `FmtInterestLine`）と **`modPlayOps`**（`RunPreflightAll` ＋ 判定核5本 `PfSurvivalOf` / `PfPredTypesOf` / `PfRefIds` / `PfFailCodeOf` / `CaseIdOfPfLine`）を宣言し、`RunPreflight` に契約（E-40の「失敗分は undiagnosed のまま」）を明記した。**未解決2件**は本文中に明記した: (a) `SetInboxJudgement` に `merged_into` を渡す引数が無く `merged` を fail-closed で拒否している、(b) `CallStep` に呼び出し単位の経路上書き口が無く config `deep_transport` が未結線（`DeepRouteOf` は解決だけを行い、指定がある間は usage_log に事実を残す）。

> v2.4.5b（W3.1）: §6へ **`modCaseStore.SetReportPath`**（18章§1.1⑦の `report_path` 書込口）、**`modExportHtml.BuildMetaJson` / `BuildReportHtml`**（18章§2・§1.1④⑤の純組立2本。層(a)と `tools/render_report.py` が叩く実契約）、**`modUICase.EnumPairsCsv` / `EnumJa` / `EnumEn`**（11章§5の共通変換表。`tools/enum_check.py` の照合先が実装名に依存したままだったのを、命名権を本節へ戻して解消）を宣言した。あわせて `modUtilText.JsStringSafe` の適用順②を18章§5.3(1) v1.1（すべての `<` を `\u003C` へ）へ追随させた。
>
> v2.4.4（裁定書8 A: W2c構造裁定）: §6へ **`modCaseStore.SetStepOutcome`**（16章E-06が要求する `last_ok_step` / `failed_step` の書込口。modPipeline の成功・失敗経路から結線し、usage_log への退避は廃止）と **`modPipeline2.RunDeep`**（入念モードの批判・改訂パイプ＝T-28 の入口。modPipeline からの委譲は1行フック）を宣言し、`RunAll` / `RunStep` に **`Optional ByVal qualityOverride As String`**（quality_mode の案件単位の上書き。ui層が `hm_quality_mode` を読んで実行時に渡す。案件一覧には保存しない）を追加した。あわせて 30,000字契約による分割先 **`modCaseStore2`**（案件2枚の下位シートI/O 11本。公開契約面には載せない）を追認した。

> v2.4.3（裁定書7: W2b整合）: §6の `modValidate` を**引数渡し設計**へ正式改訂した（Check系の末尾 Optional＝実在ID一覧テキストとJSONの外側の文脈を宣言に昇格。「ID実在はmodKnowledge参照」は**呼出側＝modPipelineがmodKnowledgeから取得して渡す**の意であると明記）。ID実在検査を**fail-closed**（一覧未提供は当該ケースIDで不合格・16章E-07のKPI「すり抜け0件」と整合）とし、`CheckS4` のティア不明時は V-S4-01/02 の両方を当てる規約を確定。**LibreOffice制約**（Optional String に `= ""` を書かない）を注記。命名権の一括裁定として `modValidate2` の *Core 5本 / `modPipeline` の判定核16本 / `modPii` 5本 / `modCompanyFile` 4本（`modCompanyFile2` の下位I/Oは公開契約面に載せない）を宣言し、案件一覧の読取専用API **`modCaseRead.ReadCaseCtx`** を新設した（modPipeline.LoadCtx と modCompanyFile.ExportCompanyFile の死に経路を解消）。

> v2.4.2（裁定書6: W2a整合）: §6を**二層**へ改訂した。(1) `Build*System` / `Build*User` / `ReviseSuffix` / `RepairSuffix` / `Block*` / `Schema*` は**引数なしのテンプレート関数**（`{{...}}` を素のまま返す＝`prompt_diff.py` の突合対象31関数）、(2) `modPromptsOps` に**組立層** `Fill` と `Asm*`（`AsmS1User` / `AsmS2User` / `AsmS3User` / `AsmS4System` / `AsmS4User` / `AsmS2CriticUser` / `AsmS3CriticUser` / `AsmSparringSystem` / `AsmPFUser`）を新設し、条件ブロック（renewal）・S4バリアント差替・想定外variantのproposalフォールバック（`fallbackNote` で帯域外に返し記録は modPipeline）をその責務とした。あわせて **`modKnowledgeFmt`**（整形の純関数10本＋`TrimKbLine`＋`TrimPlan`）を新設、`modKnowledge` の各注入関数へ `Optional maxRows`（15章§0.7の半減の口）を追加、`modCaseStore` の純ロジック4本（`BuildCaseId` / `IsValidCaseId` / `CanTransition` / `ResolveDataKey`）を公開、**`modUtil` の節を新設**して10関数を契約化した（`BufText` の区切りは vbLf）。

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
- コスト目安: 1案件=4呼び出し・入力 約25k tok・出力 約8k tok → 数十円/案件。PoC全体で数千円以内
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
- (1)～(2.5)を通すのは **modPipeline / modPlayOps の責務**（modGatewayRPNは素通しのまま返す）。(2.5)は全step・両経路で必ず通す
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
- **括弧の対応判定における文字列の扱い（①～⑦共通）**: 対応が閉じる `}` を探す走査は**文字列リテラルの内側を数えない**。すなわち `"` で開いた文字列の中に現れる `{` `}` は深さに算入せず、`\"`（直前が奇数個の `\`）は文字列の終端とみなさない。値走査（⑦）で用いるのと同じ規則を対応判定にも適用する、という一文である。これを入れないと、本文中に「開店時間は{未定}です」のような波括弧を含む日本語値があるだけで深さが狂い、正常なJSONが④（`""`）へ落ちる
- ①～⑦はいずれも modJsonLite の純ロジックであり、Excel非依存＝modTestsPure で全数テストする（17章 T-11）

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

' === core: modUtil（文字列以外の道具。裁定書6 項目8で契約化）===
Public Function SplitForCells(ByVal s As String, ByVal chunkLen As Long) As String()
' 1セル上限を超える本文の分割（13章§2.2 case_data・16章 E-22）。**空文字は0要素**を返す
' （「1件の空断片」にしない）。chunkLen <= 0 は既定の 32,000 字を使う。切れ目でサロゲート
' ペアを割らないため、各断片は chunkLen ちょうどではなく chunkLen-1 になることがある
Public Function JoinCellChunks(ByRef parts() As String) As String
' 断片の**単純連結**（区切りを入れない）。SplitForCells との往復で1字も欠けないことが契約。
' 未初期化配列は ""
Public Function SplitKeepNonEmpty(ByVal s As String, ByVal sep As String) As String()
' 区切って空要素を捨てる。**各要素は Trim する**（13章§2.2 のセル格納規約「`; ` で分割し
' 前後空白を除去」）。0件・sep が空のときは0要素
Public Function AppendIdList(ByVal listText As String, ByVal idText As String) As String
' `;` 区切りのID列へ1件足す。**既にある同一IDは追加しない**（大小文字を区別しない比較）。
' 空IDも追加しない。既存の順序は保つ（run_log.injected_kb_ids・13章§2.4）
Public Function ClampLong(ByVal v As Long, ByVal minV As Long, ByVal maxV As Long) As Long
' 値を [minV, maxV] へ収める。**minV > maxV の指定は minV を優先**する
Public Function SafeLeft(ByVal s As String, ByVal n As Long) As String
' 先頭n字で切る。**n <= 0 は ""**。末尾に単独の高位サロゲートを残さない（残すとExcelの
' 保存時に化け、開き直すまで気付けない）ため、切り口が高位サロゲートなら1字余分に落とす
Public Sub BufInit(ByRef buf() As String, ByRef itemCount As Long)
Public Sub BufAdd(ByRef buf() As String, ByRef itemCount As Long, ByVal s As String)
Public Function BufText(ByRef buf() As String, ByVal itemCount As Long) As String
' 行バッファ（16章 E-26 の32bitメモリ対策）。**BufText の区切りは vbLf**＝「1行1件の注入
' テキストを順に積む」道具である（15章§6.1）。区切り無しで継ぎたい場合は JoinCellChunks を
' 使う（BufText に区切りの分岐を持たせない）。itemCount <= 0 は ""
Public Function FindHeaderCol(ByVal headerRow As Variant, ByVal headerName As String) As Long
' 見出し行から列名の位置を引く（13章§6「列名ベース・列番号ハードコード禁止」の実体）。
' 受けるのは **Range.Value 由来の2次元配列（1行ぶん）** または Array() の1次元配列。
' 戻り値は**添字ではなく1始まりの列位置**（`Cells(r, col)` へそのまま渡せる）。不在は 0。
' 比較は前後空白を無視した大小文字非依存
Public Function ElapsedMsSince(ByVal t0 As Double) As Double   ' Timer基準の経過ms（日跨ぎ補正つき）
Public Function NowStamp() As String                           ' "yyyy-mm-dd hh:nn:ss"（正は modUtilText.IsoDateTime）

' === core: modUtilText ===
Public Function JsStringSafe(ByVal s As String) As String
' HTML内のJS文字列リテラル用。**適用順の正は18章§5.3**（この順を守らないと二重エスケープになる）:
' ① `\` を `\\` へ・`"` を `\"` へ ② **すべての "<" を "\u003C" へ**（`</` 限定では `<!--<script>` で白紙化する。18章§5.3(1) v1.1）③ 行区切り文字 U+2028 を "\u2028"・段落区切り文字
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
' **JSONの外側の文脈は引数で受ける**（裁定書7 A-1 が正式に確定した設計）。modValidate は
'   12章§2・17章§4-1 層(a) の純関数であり、シート・config・ログ・modKnowledge に触れない。
'   したがって**実在ID一覧テキスト（15章§3/§4/§6.1 の1行書式。行頭 `[ID] `）は、
'   呼び出し側＝modPipeline / modPlayOps が modKnowledge から取得して渡す**。
'   （v2.4までの CheckS3 の注記「ID実在はmodKnowledge参照」はこの引数渡しの意味であり、
'    modValidate から modKnowledge を直接呼ぶことではない。直接呼ぶと層(a)から試験できない）
' **fail-closed**（裁定書7 A-2・16章E-07のKPI「S3実在チェックのすり抜け0件」）: 検査対象の
'   キーが非空のIDを持つのに対応する一覧テキストが空（＝未提供）なら、当該ケースIDで
'   **不合格**とし `[ケースID] ID実在検査が実行できません（ID一覧未提供）` を返す。
'   値が "" のIDは従来どおり検査対象外。15章§6.1 の "(登録なし)" は空ではないので通常の
'   実在検査が走る。**引数の渡し忘れで検査が黙って消えることを禁じる**のがこの規約である
'   （対象ケース: V-S2-06 / V-S3-03..06 / V-PF-03 / V-S2C-03。15章§11の該当行にも明記）
' **LibreOffice制約**: `Optional ... As String = ""` と書かない（空文字の既定値を束縛できず、
'   省略呼び出しが実行時エラー13になり層(c)のテスト群が丸ごと止まる）。既定値を書かない
'   Optional String はVBAでも省略時に長さ0の文字列になるため意味は同一
Public Function NormalizeLlmJson(ByVal stepName As String, ByVal json As String, _
                                 ByRef removedCount As Long) As String
' 尾部劣化（同名項目の重複出力）の正規化。§5防衛線(2.5)。16章E-49。
' 配列要素を NormalizeForHash 後の fnv1a64 で重複排除し、除去件数を removedCount に返す。
' removedCount>0 は E0303（警告）として run_log detail に記録する（黙殺しない）
Public Function CheckS1(ByVal json As String, ByVal caseType As String, _
                        Optional ByVal fieldNoteProvided As Boolean = True) As String
' fieldNoteProvided: 現場メモの有無（V-S1-11。s1Json から判別できないので modPipeline が渡す）。
'   **既定は True**＝不明なら黙らず出す（V-S1-11 の判定は「警告」。15章§0 原則10）
Public Function CheckS2(ByVal json As String, ByVal caseType As String, _
                        Optional ByVal menusText As String, _
                        Optional ByVal prevS2Json As String) As String
' menusText: S2へ注入した menusSummary（MenusSummaryFor の戻り値。V-S2-06 の実在判定）。
' prevS2Json: 前ラウンドのS2 JSON（case_data `s2_prev_json`）。"" または "なし" は初回実行
'   （V-S2-09 は初回のみ / V-S2-10 は第2ラウンド以降のみ）
Public Function CheckS3(ByVal json As String, ByVal s2Json As String, _
                        Optional ByVal menusText As String, _
                        Optional ByVal linesText As String, _
                        Optional ByVal schemesText As String, _
                        Optional ByVal casesText As String, _
                        Optional ByVal caseType As String) As String
' menusText/linesText/schemesText/casesText: S3へ注入した一覧（MenusFor / LinesText /
'   SchemesFor / CasesFor の戻り値。V-S3-03..06 の実在判定）。caseType は V-S3-13 用
'   （省略時は s2Json の gaps 非空から更新案件を導く。V-S2-12 が「新規で gaps 1件以上」を
'    不合格にしているため確定できる）
Public Function CheckS4(ByVal json As String, _
                        Optional ByVal dossierTier As String, _
                        Optional ByVal maxSlidesT2 As Long = 10) As String
' dossierTier: 13章§2.1 の dossier_tier（s4Json に無いので呼び出し側が渡す）。**省略・空は
'   「ティア不明」として V-S4-01 と V-S4-02 の両方を当てる**（裁定書7 A-4。片方のティアを
'   黙って既定に据えると、もう一方の違反〈t1_quick の5枚固定〉が構造的にすり抜ける。
'   5枚は両ティアで合法なので正常な応答はティア不明でも0件）。maxSlidesT2=config `ppt_max_slides_t2`
Public Function CheckS2C(ByVal json As String, Optional ByVal s2Json As String) As String
' 入念モードの批判JSON（15章§4.5）。s2Json=審査対象のS2（V-S2C-03 の番号実在判定。
'   未提供は fail-closed で V-S2C-03 を不合格にする）
Public Function CheckS3C(ByVal json As String) As String          ' 入念モードの批判JSON（15章§4.6）
Public Function CheckPF(ByVal json As String, Optional ByVal refIdsText As String) As String
' refIdsText: PFへ注入した一覧（menusSummary + patternsText + rulesText + researchingText を
'   連結したもの。V-PF-03 の実在判定）。PL-03を結線する modPlayOps は必ず渡すこと
Public Function CheckWT(ByVal json As String) As String          ' Phase1.5
Public Function CheckFG(ByVal json As String) As String          ' Phase1.5

' === app: modValidate2（30,000字契約による modValidate の分割先。裁定書7 B-5）===
' **分割の継ぎ目**であり、呼んでよいのは modValidate だけ（依存は modValidate -> modValidate2 の
'   一方向。逆参照はしない）。公開名の `*Core` 接尾辞は「14章§6の公開名の本体」の意で、
'   引数はいずれも modValidate 側の入口と同じ意味を持つ（Optional は入口側にだけ置く）。
Public Function NormalizeCore(ByVal stepName As String, ByVal json As String, _
                              ByRef removedCount As Long) As String
Public Function CheckS4Core(ByVal json As String, ByVal dossierTier As String, _
                            ByVal maxSlidesT2 As Long) As String
Public Function CheckPFCore(ByVal json As String, ByVal refIdsText As String) As String
Public Function CheckS2CCore(ByVal json As String, ByVal s2Json As String) As String
Public Function CheckS3CCore(ByVal json As String) As String

' === app: modKnowledge ===
' 各注入関数の `Optional ByVal maxRows As Long = 0` は**15章§0.7の段階的な半減を外から
' 掛けるための口**（0=config既定＝`kb_*_rows`。正の値を渡すとその行数で打ち切る）。
' modPipeline は `modKnowledgeFmt.TrimPlan` が返した行数をここへ渡して再取得する。
' 整形（1行の書式・0行の既定文言・空項目の省略）そのものは **modKnowledgeFmt** の純関数が
' 行い、本モジュールは「読む・絞る・注入IDを積む」だけを担う（12章§2）。
Public Function LoadKnowledge() As Boolean            ' 起動時/再読込。スナップショット保存込み
' **退避を空で上書きしない**（裁定書9 B11）: ナレッジブックの読取（`ReadKbSheets`）が
'   データ行のあるシートを**1枚も返さなかったときはスナップショットに触らない**（`SaveSnapshot`
'   を呼ばず E0401 を記録して False を返す）。0行読込で前回の正常な退避を消すと、次に接続
'   できない起動で16章 E-08 の退路（`RestoreSnapshot`）が失われる
Public Function RiskLibFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
' 整形済注入テキスト（S2用。書式の正は15章§3）
Public Function MenusSummaryFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
' S2用のメニュー「要約」。preventions.related_menu_id が選べる候補一覧を与える唯一の口。
' これを注入せずに CheckS2 の「related_menu_id は "" または実在」を課すと候補ゼロで実在を要求する
' 矛盾になるため必須。整形書式の正は15章§3
Public Function MenusFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
' S3用のメニュー一覧（実在するサービス。S2の要約とは別テキスト。書式の正は15章§4）
Public Function LinesText(Optional ByVal maxRows As Long = 0) As String
Public Function CasesFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
Public Function SchemesFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
' status∈{proven,adopted}のみ
Public Function PatternsText(Optional ByVal maxRows As Long = 0) As String   ' P1-P15全件（PF/FG用）
Public Function RulesText(Optional ByVal maxRows As Long = 0) As String      ' 判断基準（PF/FG用）
Public Function ResearchingText(Optional ByVal maxRows As Long = 0) As String ' 研究テーマ一覧（PF用）
Public Function MechsText(Optional ByVal maxRows As Long = 0) As String
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
' **二次照合の呼出点**（裁定書9 B14。16章NFR-S7 設計原則(2)「取得側でも再チェック」）: 上の5本は
'   `modUICase2.ValidateEdited` の**後段**で、確定しようとしているJSONに現れるIDをナレッジブック
'   本体と突き合わせるために呼ぶ。一次防御（`modValidate` の注入テキスト照合＝ListGiven の
'   fail-closed）は**そのまま残す**。二次側の不一致は保存をブロックせず `hm_warning` へ出す
'   （一次が通った以上そのIDは提示済み候補であり、差分は注入の切詰め〈15章§0.7〉や
'   ナレッジブック側の更新で生じうるため。宣言だけして呼ばない状態を残さない）
Public Sub AppendServiceGap(ByVal caseId As String, ByVal industryCode As String, ByVal riskDesc As String)
' 10章 FR-13（Must）・12章§2.1 手順6・13章§3.9 の「該当メニュー・型なし＝新サービス候補として
'   自動記録」の実体。**呼出点は `modPipeline.ExecStep` の stepNo=3 成功直後**に固定する
'   （裁定書9 B10）。確定した `s3_json` の `unmatched_risks` を1件ずつ回し、
'   `AppendServiceGap caseId, ctx.industry_code, risk_name & " / " & why_unmatched` を呼ぶ。
'   11章のS3ワイヤーが「新サービス候補として記録済み」と表示する以上、宣言だけして
'   1行も書かない状態を残さない。書込失敗時の退避キューの扱いは下の注記による
' **退避キューを空で消さない**（裁定書9 B19。16章 E-13）: ナレッジブックへの書き戻しは
'   列名一致で1行ずつ書き、**書けた件数を数える**。0件しか書けなかったときは退避キューの
'   `Cells.Clear` を実行しない（見出し不在・列名変更で1件も書けていないのにキューだけ
'   空になる経路を残さない）

' === app: modKnowledgeFmt（ナレッジ整形の純関数。Excel非依存＝層(a)から直接叩ける） ===
' modKnowledge から「整形」だけを切り出したモジュール（12章§2）。シート・config・ログに
' 一切触れない純文字列関数だけを置く。**15章の1行書式の唯一の実装**であり、書式を変えると
' ここのテストが落ちる（W2aでは整形が Private のままだったため、区切り記号を ` | ` から
' ` / ` へ壊してもどのゲートも気づかなかった）。
' 共通の引数 `rows`: **Range.Value 由来の2次元 Variant 配列**（1行目=見出し行＝列名、
'   2行目以降=データ行。添字は 1..n / 1..cols）。列は**列名で引く**（13章冒頭。列番号の
'   ハードコード禁止＝`modUtil.FindHeaderCol`）。渡された全データ行を整形する（業種の絞込・
'   is_active・status・行数上限は modKnowledge が適用済みで渡す）。配列でない／データ行0件の
'   ときは各仕様の既定文言を返す。
' 共通の戻り値: 15章の書式の**複数行文字列**（1行1件・行区切りは vbLf・末尾に改行を付けない）。
'   行頭は `[ID] `、項目区切りは ` | `、項目内の複数値は `;`、**値が空の項目は項目ごと省略**する
'   （`market_note` の空欄省略もこの規約の一適用。15章§4・§6.1）。
Public Function FmtRiskLib(ByVal rows As Variant) As String     ' 15章§3。0行は "(この業種の登録知識はまだありません)"
Public Function FmtMenus(ByVal rows As Variant) As String       ' 15章§4 menusText（概要つき）。0行は "(登録なし)"
Public Function FmtMenusSummary(ByVal rows As Variant) As String ' 15章§3・§6.1 menusSummary。0行は "(登録なし)"
Public Function FmtLines(ByVal rows As Variant) As String       ' 15章§4 linesText（market_note 空欄は省略）。0行は "(登録なし)"
Public Function FmtCases(ByVal rows As Variant) As String       ' 15章§4 casesText。0行は "なし"（S3 userの見出し規約）
Public Function FmtSchemes(ByVal rows As Variant) As String     ' 15章§4 schemesText。0行は "なし"
Public Function FmtPatterns(ByVal rows As Variant) As String    ' 15章§6.1 patternsText。0行は "(登録なし)"
Public Function FmtRules(ByVal rows As Variant) As String       ' 15章§6.1 rulesText。0行は "(登録なし)"
Public Function FmtResearching(ByVal rows As Variant) As String ' 15章§6.1 researchingText。0行は "(登録なし)"
Public Function FmtMechs(ByVal rows As Variant) As String       ' 15章§6.1 mechs。0行は "(登録なし)"（16章E-09）
Public Function TrimKbLine(ByVal s As String) As String  ' 適用点はmodKnowledgeFmt内部のRowLine最終段(全整形行が必ず通る)
' 15章§0.7 の最終段「各行を先頭400字で切り『…』を付す」の実体。**400字以内はそのまま返す**
' （何も足さない）。超える場合は先頭400字（`modUtil.SafeLeft` と同じサロゲート安全な切り方）へ
' `…`（U+2026。CP932内）を付けて返すので、戻り値は最大401字になる。
Public Function TrimPlan(ByRef counts() As Long, ByVal budgetChars As Long) As Long()
' 15章§0.7「ナレッジ側の切詰め」を**計画するだけ**の純関数（実際に削るのは modKnowledge の
' maxRows）。切詰め順は 成功事例→型→メニュー→種目→リスクライブラリ で固定。
'   counts: 10要素（0始まり）。前半 counts(0..4)＝各対象の**現在の行数**、
'           後半 counts(5..9)＝同じ並びの**現在の文字数**。並びは上の切詰め順。
'           要素が5個以下のときは文字数を0とみなす（＝切詰め不要と判断する）。
'   budgetChars: ナレッジ注入に許される**合計文字数**（15章§0.7「上限の3割」）。0以下は
'           上限なしとして扱い、現在の行数をそのまま返す。
' 戻り値: 5要素（0始まり）の「注入してよい行数」。**1～5を順に1段ずつ**適用し、そのつど
'   総量を再計算して budgetChars 以下になった時点で止める（§0.7の本文どおり。1対象あたり
'   半減は1回まで）。半減は端数切上げ、**下限は 0 / 0 / 5 / 5 / 5 行**（メニュー・種目・
'   リスクライブラリを5行未満にすると S3のID実在制約と§0.5第2層が崩れるため）。
'   文字数は行数に比例すると見積もる（削った行の実長は事前に測れないため）。5段すべてを
'   適用してなお超過する場合も戻り値は下限どおりで、次の手当ては `TrimKbLine` の行内切詰め。

' === app: modPromptsCore / modPromptsBlocks / modPromptsOps / modSchemas ===
' **テンプレート層と組立層の二層に分ける**（W2aで「Build* が受け取った引数を1つも使わず
' プレースホルダを素のまま返す」欠陥が出たため、責務を名前で分離した）:
'   (1) **テンプレート層 = 引数なしの純関数**。15章の本文を `{{プレースホルダ}}` を含んだ
'       **素のまま**返す。`tools/prompt_diff.py` の突合対象は**この層の31関数だけ**であり、
'       15章§10.2の対応表はこの層のまま不変。**31関数はすべて無引数**である（prompt_diff の
'       評価器は文字列リテラルと vbLf 等の組込定数の連結しか評価できず、引数参照は評価不能=
'       差分になる。使えない引数は持たせない）
'   (2) **組立層 = modPromptsOps の `Fill` と `Asm*`**。テンプレートへ実値を埋め、条件ブロックの
'       挿入とS4バリアントの差替を行う。**15章の本文は1文字も持たない**（本文を2箇所に書かない）
' 本文は15章と一字一句一致（T-23がdiffゼロを検査）。**Const は使わず、`s = s & "..." & vbLf` 方式の
' 純関数で組み立てて返す**（VBAの Const は1行1023字・行継続25本の制約に当たり、3,700字級の
' スキーマ本体を1宣言で書けないため）。一致検査の正規化は「改行=vbLf・末尾改行なし」。
' 15章の全 {{プレースホルダ}} に対応する引数は (2) の Asm* が持つ（19章§5の文書間整合チェックリスト）。
' --- (1) テンプレート層（prompt_diff の突合対象31関数。全て無引数）---
Public Function BuildS1System() As String                                        ' 15章§2 system
Public Function BuildS1User() As String                                          ' 15章§2 user
Public Function BuildS2System() As String                                        ' 15章§3 system
Public Function BuildS2User() As String                                          ' 15章§3 user
Public Function BuildS3System() As String                                        ' 15章§4 system
Public Function BuildS3User() As String                                          ' 15章§4 user
Public Function BuildS2CriticSystem() As String                                  ' 15章§4.5 批判system
Public Function BuildS2CriticUser() As String                                    ' 15章§4.5 批判user
Public Function BuildS3CriticSystem() As String                                  ' 15章§4.6 批判system
Public Function BuildS3CriticUser() As String                                    ' 15章§4.6 批判user
Public Function ReviseSuffix() As String                                         ' 15章§4.7 改訂サフィックス
Public Function BuildS4System() As String                                        ' 15章§5 system
Public Function BuildS4User() As String                                          ' 15章§5 user
Public Function BuildPFSystem() As String                                        ' 15章§6 PL-03 system
Public Function BuildPFUser() As String                                          ' 15章§6 PL-03 user
Public Function BuildSparringSystem() As String                                  ' 15章§6.5 壁打ちsystem
Public Function RepairSuffix() As String                                         ' 15章§7 修復サフィックス
Public Function SchemaS1() As String   ' 同様に SchemaS2 / S3 / S4 / S2C / S3C / PF / WT / FG（§7の表が正）
' `Block*` 7本（BlockCtx / BlockRenewalS1..S3 / BlockGuard / BlockS4Proposal / BlockS4Alliance）は
'   modPromptsBlocks の Public 関数（modPromptsOps が参照するため Private 不可）だが、本節の公開契約面には載せない（正は15章§10.2）。同じく無引数。
' `ReviseSuffix` / `RepairSuffix` は user 末尾へ連結する1ブロックであり専用の Asm* を置かない。
'   `{{critiqueDigest}}` / `{{validationErrors}}` の埋め込みは呼び出し側が `Fill` で行う。
' --- (2) 組立層（modPromptsOps。純関数=Excelトークン禁止・15章の本文を持たない）---
Public Function Fill(ByVal tpl As String, ByRef names() As String, ByRef vals() As String, _
                     Optional ByRef unresolved As Long = 0) As String
' テンプレート中の `{{name}}` を対応する値へ**全置換**する唯一の口。names(i) と vals(i) は同じ
' 添字で対応させる（要素数が食い違う場合は短いほうまでを処理する）。names の要素には
' `{{` `}}` を含まない**識別子だけ**を渡す。値の中に `{{...}}` が含まれていても**再帰置換はしない**
' （置換は names の順に1巡だけ行い、置換後の文字列を再走査しない。外部由来テキストが
' プレースホルダを名乗って別の値を奪うのを防ぐ＝16章E-04と同じ考え方）。
' unresolved には**置換後になお残っている `{{` の個数**を返す（0が正常。呼び出し側は
' run_log の detail に記録する）。省略可能な出口なので `Fill(tpl, n, v)` でも動く。
Public Function AsmS1User(ByVal ctx As TCaseCtx, ByVal hpTxt As String, ByVal yuhoTxt As String, _
                          ByVal memoTxt As String, ByVal contractTxt As String, ByVal prevRenewalTxt As String, _
                          ByVal dossierTxt As String, ByVal fieldNotes As String, ByVal coverageNote As String, _
                          ByVal hearingAnswers As String) As String
' 15章§2 S1 userの9貼付ブロックを順に埋める（hp→yuho→memo→contract→prevRenewal→dossier→
' fieldNotes→coverageNote→hearingAnswers）。coverageNote=13章 input_coverage_note（付保の見立て。
' 伝聞情報だが insurance_ctx 観点の充足度評価に算入するため省略不可）。
' **条件ブロック**: ctx.case_type="renewal" のとき `{{BLOCK_RENEWAL_S1}}` の行を BlockRenewalS1() の
' 本文へ差し替え、それ以外では**その行ごと削除する**（空行を残さない。15章§10.1(d)）
Public Function AsmS2User(ByVal ctx As TCaseCtx, ByVal s1Json As String, ByVal riskLib As String, _
                          ByVal menus As String, ByVal prevS2Json As String, _
                          ByVal hearingAnswers As String) As String
' 15章§3 S2 user。menus=MenusSummaryFor()（preventions.related_menu_id の候補一覧。空で渡すと
' CheckS2の「実在ID」検査が構造的に落ちる）。prevS2Json=case_data の s2_prev_json、
' hearingAnswers=input_hearing_answers（いずれも初回ラウンドは "なし"）。FR-35のstatusライフサイクル
' （confirmed/rejected/new）はこの2引数がなければ成立しない。
' `{{BLOCK_CTX}}` は BlockCtx() を ctx で埋めたものへ、`{{BLOCK_RENEWAL_S2}}` は AsmS1User と同じ規約
Public Function AsmS3User(ByVal ctx As TCaseCtx, ByVal s1Summary As String, ByVal s2Json As String, _
                          ByVal menus As String, ByVal lines As String, ByVal schemes As String, _
                          ByVal cases As String) As String
' 15章§4 S3 user。s1Summary=S1出力の business_summary / strategy_outlook / current_coverage /
' field_insights だけを抜き出した要約JSON（S3 systemルール7が field_insights の参照を命じており、
' S2 JSONには含まれないため必須）。menus=MenusFor()（S3は実在サービスの一覧。S2の要約とは別テキスト）。
' `{{BLOCK_RENEWAL_S3}}` は AsmS1User と同じ規約
Public Function AsmS4System(ByVal variantName As String, ByVal tier As String, _
                            Optional ByRef fallbackNote As String = "") As String
' 15章§5。**S4バリアント差替の唯一の担い手**。variantName="proposal" なら BlockS4Proposal()、
' "alliance" なら BlockS4Alliance() を `{{BLOCK_S4_VARIANT}}` の位置へ差し込む。
' **それ以外の値（空文字を含む）は proposal として扱い**、fallbackNote へ
' `s4_variant_fallback:{value}` を返す（modPrompts* は run_log へ書けないため、記録は
' 呼び出し側=modPipeline が行う。黙って既定に落とさない）。
' tier=t1_quick / t2_full / t3_sparring（t3_sparring は t2_full と同じ扱い）。
' `{{pptMaxSlidesT2}}` は modConfig の `ppt_max_slides_t2`（既定10）を展開する
Public Function AsmS4User(ByVal ctx As TCaseCtx, ByVal s1Json As String, ByVal s2Json As String, _
                          ByVal s3Json As String) As String
' 15章§5 user。`{{slideCountHint}}` は ctx.dossier_tier から決める（t1_quick は "5"、
' それ以外は "5～" & ppt_max_slides_t2）。`{{company}}` は ctx.company
Public Function AsmS2CriticUser(ByVal s1Json As String, ByVal s2Json As String, _
                                ByVal riskLib As String) As String               ' 15章§4.5 批判user
Public Function AsmS3CriticUser(ByVal ctx As TCaseCtx, ByVal s1Summary As String, _
                                ByVal s2Json As String, ByVal s3Json As String) As String  ' 15章§4.6 批判user
Public Function AsmSparringSystem(ByVal dossierSummary As String, ByVal s1s2s3Json As String, _
                                  ByVal schemes As String, ByVal patterns As String, _
                                  ByVal mechs As String, ByVal rules As String) As String
' PL-04壁打ちのsystem（15章§6.5）。mechs は modKnowledge.MechsText()（Phase1は "(登録なし)"）
Public Function AsmPFUser(ByVal theme As String, ByVal body As String, ByVal rules As String, _
                          ByVal menusSummary As String, ByVal schemes As String, ByVal patterns As String, _
                          ByVal researching As String) As String                 ' 15章§6 PL-03 user
' Asm* 共通: **enum→日本語ラベルの変換は行わない**（変換表の正は19章§3・実体は modUICase であり、
'   同じ表を2箇所に書かないため）。`{{case_typeの日本語}}` 等には ctx の値をそのまま埋めるので、
'   ラベル済みの ctx を渡すのは呼び出し側（modPipeline）の責務である。
' TCaseCtx（**app層 modAppTypes**。ドメイン型なのでcore層 modTypes から移設。12章§2・§4）:
'   case_type, dossier_tier, channel, kanji, bid, reins, other_insurers, company, industry_code, industry_name

' === app: modPipeline / modPipeline2 / modPlayOps ===
Public Function RunAll(ByVal caseId As String, Optional ByVal qualityOverride As String) As Boolean
Public Function RunStep(ByVal caseId As String, ByVal stepNo As Long, _
                        Optional ByVal qualityOverride As String) As Boolean
    ' quality_mode=deep のとき、S2/S3は 生成→批判(CheckS2C/S3C合格の批判JSON)→
    ' 指摘ありなら改訂(ReviseSuffix)の3呼び出しで実行。批判・改訂はrun_logに
    ' step="s2c"/"s3c"/"s2r"/"s3r" として記録。deep_transport指定時は批判・改訂のみ経路変更
' qualityOverride(裁定書8 A-3): quality_mode の**案件単位の上書き**(13章§2.3・§2.10)。
'   ui層が HOME の `hm_quality_mode` を読んで実行時に渡す値であり、**案件一覧には
'   保存しない**(実行時の指定であって案件の属性ではない)。空は config・ティア連動の
'   ままで、解決は純核 `ResolveQualityMode` が唯一の値源。RunAll は素通しで各Stepへ
'   渡す(4Stepを同じ品質モードで走らせる)。
'   **LibreOffice制約**: Optional String に既定値リテラル(`= ""`)は書かない。
'   空判定は IsMissing ではなく LenB(Trim$(...)) で行う。
Public Function RunPreflight(ByVal inboxId As String) As Boolean
' 受信箱1件のプリフライト診断（PL-03・step=pf。15章§6）。True=検証合格まで到達し
'   `pf_json` / `pf_survival` / `pf_pred_types` を保存できた。False は**当該行を
'   undiagnosed のまま残す**（16章 E-40 の「失敗分は undiagnosed のまま」を1件単位でも守る）。
'   受信箱のI/Oは modInboxStore が唯一の口（modPlayOps は R4 でシートに触れない）
Public Function RunPreflightAll() As Long
' 未診断の投函の一括診断（17章 T-25・16章 E-40。裁定書8 B-7）。途中で止めず最後まで回し、
'   **診断できた分は保存・失敗した分は undiagnosed のまま残す**。1件でも落ちたら E0701 を
'   件数つきで記録する。戻り値=診断できた件数。Step間で DoEvents を挟む（E-50(c)）
' --- modPlayOps の判定核5本（裁定書8 B-7。シート・ログ・LLMに触れない純関数） ---
' 規約（診断結果の要約列・ID一覧の組立・不合格の内訳）を実行制御へ閉じ込めない。**Private へ
'   戻すことは契約違反**（vba_lint の CONTRACT required が検出する。裁定書9-3）。
Public Function PfSurvivalOf(ByVal pfJson As String) As String
' 診断JSONから受信箱の要約列 `pf_survival` を取り出す（13章§2.6）。enum（high/mid/low）
'   以外は ""（未定義値を列へ書かない）
Public Function PfPredTypesOf(ByVal pfJson As String) As String
' 同 `pf_pred_types`（";"区切り T1～T10）。enum外は落とし重複は1件へ寄せる。T0 は PF の
'   `predicted_drop_types` には現れない（19章§3）
Public Function PfRefIds(ByVal rulesText As String, ByVal menusSummary As String, _
                         ByVal schemesText As String, ByVal patternsText As String, _
                         ByVal researchingText As String) As String
' `CheckPF` の V-PF-03（ref_id 実在検査）へ渡す一覧テキスト。15章§6 が PF へ注入する5種を
'   1行書式（15章§6.1）のまま vbLf で連結する。空の注入は行ごと落とす。**一覧を渡さない
'   呼び出しは fail-closed で不合格**なので、PFを結線する側は必ずこの戻り値を渡す
Public Function PfFailCodeOf(ByVal errText As String) As String
' PFの不合格の内訳をコードへ。**ID幻覚（V-PF-03）を含めば E0301（16章E-07）、他は E0302**
Public Function CaseIdOfPfLine(ByVal lineText As String) As String
' 検証エラー1行の先頭 `[ケースID] ` からケースIDを取り出す（15章§0 原則10）。不一致は ""
' --- modPipeline の判定核16本（裁定書7 B-6。シート・ログ・LLMに触れない純関数）---
' 規約そのもの（打切り計画・修復要否・結果の分類・失敗コードの切分け・deep分岐）を実行制御の
'   中に閉じ込めると層(a)から誰も検査できない（W2aの modKnowledge 整形と同じ轍）。以下は
'   modTestsPure から直接叩く前提で公開し、run_lo_tests の PURE_ALLOWLIST にも modPipeline を
'   登録する。**Private へ戻すことは契約違反**（vba_lint の CONTRACT required が検出する）。
Public Function TrimInputPlan(ByRef lens() As Long, ByVal budgetChars As Long) As Long()
' 16章E-03(2)の打切りを**計画するだけ**（実際に削るのは呼び出し側）。lens(0..4)=切る順
'   （追加ドシエ/前回更新メモ/有報/営業メモ/HP）の現在字数、lens(5)=打切らない4欄の合計。
'   budgetChars<=0 は上限なし。戻り値=5要素の「残してよい字数」。1欄ずつ削って再計測し、
'   上限を下回った時点で止める。**4欄だけで超過する場合は自動では削らない**（E-03(3)(4)）
Public Function ProtectedOverBudget(ByVal keepChars As Long, ByVal budgetChars As Long) As Boolean
' 打切らない4欄だけで上限超過か（E-03(4) の E0102 警告の唯一の条件）
Public Function TruncField(ByVal s As String, ByVal allowedChars As Long) As String
' 切詰めた欄に注記「（一部省略）」を付す（E-03(6)）。allowedChars<=0 は注記だけを返す
Public Function BudgetOf(ByVal limitChars As Long, ByVal pct As Long) As Long
' 15章§0.7 の予算配分（貼付=上限の7割 / ナレッジ=3割）。pct は 7 または 3。16章E-03の
'   「上限」の解釈もこの配分が正（E-03 に同旨を追記済み。裁定書7 B-9）
Public Function NeedsRepair(ByVal errText As String, ByVal retryBudget As Long) As Boolean
' 修復要否（§5防衛線(4)）。検証エラーがあり再試行枠（config `json_repair_retry`）が残るときだけ
Public Function ClassifyResult(ByVal firstOk As Boolean, ByVal repairTried As Boolean, _
                               ByVal repairOk As Boolean) As String
' run_log の validate_result（13章§2.4）の3値。初回合格=ok / 修復後合格=repaired / 他=failed
Public Function FailCodeOf(ByVal errText As String) As String
' 不合格の内訳をエラーコードへ。**ID幻覚（16章E-07）を含めば E0301、それ以外は E0302**。
'   判定は検証エラー行の先頭のケースID（V-S2-06 / V-S3-03..06）で行う
Public Function CaseIdOfLine(ByVal lineText As String) As String
' 検証エラー1行の先頭 `[ケースID] ` からケースIDを取り出す（15章§0 原則10の書式）。不一致は ""
Public Function DeepEnabled(ByVal qualityMode As String, ByVal stepNo As Long) As Boolean
' 入念モードの分岐（S2/S3のみ。15章§4.5-4.7）
Public Function ResolveQualityMode(ByVal cfgMode As String, ByVal tier As String) As String
' config `quality_mode`（13章§2.3）の解決。空はティア連動（t1_quick=standard / 他=deep）
Public Function StepNameOf(ByVal stepNo As Long) As String        ' 19章§4 の step 値（s1..s4。範囲外は ""）
Public Function StatusForStep(ByVal stepNo As Long) As String     ' Step成功時の遷移先（13章§2.1 sN_done）
Public Function PlayIdOf(ByVal caseType As String) As String      ' renewal=PL-02 / それ以外=PL-01（§1）
Public Function KbRowCount(ByVal s As String) As Long             ' 注入テキストの行数（0行の既定文言は0）
Public Function UsesSlot(ByVal stepNo As Long, ByVal slot As Long) As Boolean
' 12章§3のStep別ナレッジ枠（S2=リスクライブラリ+メニュー要約 / S3=メニュー/種目/型/事例）
Public Function S1SummaryOf(ByVal s1Json As String) As String
' 15章 S3/S3C の {{s1SummaryJson}}（business_summary / strategy_outlook / current_coverage /
'   field_insights の4キーだけのJSON。他のキーを含めない）

' === app: modPipeline2（入念モードの批判・改訂パイプ。T-28。裁定書8 A-1）===
' 30,000字契約による modPipeline の分割先。**分割の継ぎ目**であり、呼んでよいのは
'   modPipeline だけ(依存は modPipeline -> modPipeline2 の一方向)。R4でシートには
'   触れない(案件データは modCaseStore / modCaseRead 経由)。
Public Function RunDeep(ByVal caseId As String, ByVal stepNo As Long) As Boolean
' 入念モードの入口。modPipeline は sN_json を確定した直後、`DeepEnabled` が True の
'   ときだけ**1行で委譲**する(分岐判定の純核は modPipeline の16本のまま)。stepNo は
'   2 または 3(それ以外は False で何もしない)。15章§4.5-4.7 の 生成→批判→改訂 を
'   回し、s2c/s3c/s2r/s3r を modCaseStore 経由で保存して run_log へ記録する。
'   戻り値 True=パイプを完走した(改訂の採否は問わない)。**戻り値で本体Stepの成否を
'   左右しない**のが契約(16章 E-35 批判不合格=生成版を確定して警告 / E-36 改訂不合格=
'   改訂を破棄して改訂前を採用。どちらも本体Stepは成功のままで failed_step を立てない)。
'   呼び出し側はこの値を握りつぶしてよい。
' --- modPipeline2 の判定核8本（裁定書8 B-10 の7本 ＋ 裁定書9-2 の `AdoptRevisionOf`。
'     A-1 が「T-28 の実装時に本節へ足す」と予告していた「パイプ内部の関数構成」。
'     シート・ログ・LLMに触れない純関数） ---
' 規約そのもの（改訂へ進む条件・批判の日本語整形・E-35/E-36 の結末の分類と確定JSONの
'   選択・HOMEへ出す文言・deep_transport の解決）を実行制御の中に閉じ込めると層(a)から
'   誰も検査できない。
'   modTestsPure から直接叩く前提で公開し、run_lo_tests の PURE_ALLOWLIST にも
'   modPipeline2 を登録する。**Private へ戻すことは契約違反**（裁定書9-3 で
'   `AdoptRevisionOf` を含む8本を vba_lint の CONTRACT required へ載せ、機械で検出する）。
Public Function CritiqueStepOf(ByVal stepNo As Long) As String   ' 2→s2c / 3→s3c / 他は ""（19章§4）
Public Function ReviseStepOf(ByVal stepNo As Long) As String     ' 2→s2r / 3→s3r / 他は ""（同）
Public Function NeedsRevision(ByVal critiqueJson As String, ByVal stepNo As Long) As Boolean
' 改訂パスへ進むか。issues が1件でもあれば True。加えて S2C は `additional_risks` が
'   非空なら、S3C は `lands=false` の反応があれば True（V-S2C-05 / V-S3C-05 のスキップ
'   条件の裏返し）。範囲外の stepNo・空JSONは False
Public Function CritiqueDigest(ByVal critiqueJson As String, ByVal stepNo As Long) As String
' 15章§4.7 の `{{critiqueDigest}}`。批判JSONを日本語の箇条書き（行区切り vbLf）へ整形する。
'   S2 は issues → additional_risks、S3 は lands=false の反応 → issues の順
Public Function DeepOutcomeOf(ByVal critiqueOk As Boolean, ByVal revisionTried As Boolean, _
                              ByVal revisionOk As Boolean) As String
' パイプの結末の唯一の分類点。`critique_skipped`（E-35）/ `revision_skipped`（指摘0件）/
'   `revised` / `revision_discarded`（E-36）の4値。**どの値でも本体Stepは成功のまま**
Public Function AdoptRevisionOf(ByVal outcome As String, ByVal originalJson As String, _
                                ByVal revisedJson As String) As String
' **確定として採用するJSONの唯一の選択点**（16章 E-36。裁定書9-2）。`outcome` は
'   `DeepOutcomeOf` の4値で、`revised` のときだけ改訂版を、それ以外（`critique_skipped` /
'   `revision_skipped` / `revision_discarded`）は改訂前＝検証合格済みの生成版を返す。
'   **2本を継ぎ合わせない**（どちらか1本をそのまま返す）。`RunPipe` は確定JSONの選択も
'   `sNr_json` の保存可否もこの戻り値を経由し、**不合格の改訂版が `sNr_json` へ入る経路を
'   構造として持たない**（E-36 の意思決定を実行制御側の If で書き直さない）
Public Function DeepWarningOf(ByVal outcome As String) As String
' HOME の `hm_warning` へ出す文言（16章 E-35/E-36 の逐語）。警告の要らない結末は ""。
'   app層から ui層 は呼べない（R1）ので、値だけを供給して ui層（modUIHome）が読む
Public Function DeepRouteOf(ByVal cfgDeepTransport As String) As String
' config `deep_transport`（13章§2.3）の解決。`direct` のときだけ "direct"、他は ""。
'   **経路の切替そのものは未結線**（本節の `CallStep` に呼び出し単位で経路を上書きする
'   口が無い）。RunDeep は指定がある間その事実を usage_log に残す（黙って無視しない）
Public Function LastDeepOutcome() As String
' 直近の実行が回した入念パイプの結末（裁定書9 N1・B9。契約はv2.5.1・裁定書10 M1で改訂）。
'   値は `DeepOutcomeOf` の4値のうち**警告を伴う2値**（`critique_skipped` / `revision_discarded`）
'   と、警告が要らない場合の `""` の3通り。**`RunStep` は本値をリセットしない**。リセットは
'   下の `ResetDeepOutcome`（N9）のみが行い、ui層（`modUIHome.RunStepUi` / `HomeRunAll`）が
'   **実行開始前に1回**呼ぶ。したがって本関数は「その実行で最後に立った非空 outcome」を返す
'   （一括実行では Step2/3 で立った outcome が Step4 の実行後も残る＝E-35/E-36 警告が
'   `RunAll` でも消えない。旧契約の「RunStep 開始時リセット」は一括実行で警告を握り潰すため廃止）。
'   ui層は成功分岐でこの値を読み、非空なら `DeepWarningOf` の文言を `hm_warning` へ出す
'   （16章 E-35/E-36 の逐語表示の唯一の経路）。
'   **`ShowWarning vbNullString` によるクリアより後で書く**（冒頭のクリアに消されない順序）。
'   モジュール変数による状態保持は本関数を**例外として許可する**（`broken_json_once` に次ぐ
'   2例目。理由: `RunStep` の Boolean 戻り値の契約を変えずに E-35/E-36 を ui へ渡す口が他に無い）
Public Sub ResetDeepOutcome()
' deep outcome の**明示リセット口**（v2.5.1・裁定書10 N9/M1）。`LastDeepOutcome` の内部状態を
'   "" へ戻す。呼ぶのは ui層（`modUIHome.RunStepUi` / `HomeRunAll`）の**実行開始前に1回**だけ。
'   `RunStep` / `RunAll` 自身は呼ばない（実行の途中で立った警告を実行自身が消さない）

' === app: modSparring（PL-04 壁打ち。T-27。裁定書8 B-9）===
' 自由対話（スキーマなし）。呼び出しは `CallChat` の1本だけで、成否は `ByRef ok`
'   （＝`DecideOk`）でしか判定しない。JSON防衛線（§5）は通さない。run_log は
'   `CallChat` が自分で1行書くため本モジュールは書かない（1呼び出し＝2行にしない）。
' R4: Excelトークン許可の13本（12章§4）に**入れない**。案件一覧は `modCaseRead`、
'   case_data は `modCaseStore`、受信箱は `modInboxStore` を通す。
Public Function ResumeSparring(ByVal caseId As String, ByRef contextNote As String) As Long
' 「壁打ちを開始/再開」（11章 壁打ちワイヤー・15章§6.5）。戻り値＝保存済みの発話数
'   （0＝履歴なし＝新規開始）。**-1＝案件一覧を読めない**（呼び出し側は fail-closed で
'   開始させない）。contextNote＝13章§2.17 `sp_context_note` の表示文字列
'   （例「ドシエ+S1-S3+型/機構 注入済」）。
'   **解消（裁定書9 A-1）**: 13章§2.1/§2.17 が求める `dossier_tier` の t3_sparring への
'   自動昇格は **`modCaseStore.PromoteTier`**（N2）を唯一の書込口として実行する。本関数は
'   履歴の解決に先立って `PromoteTier caseId, "t3_sparring"` を呼び、戻り値 False（enum不合格
'   ・行不在）のときは昇格せずに usage_log へ事実を残して続行する（昇格の失敗で壁打ちの
'   開始そのものを止めない）。usage_log は「昇格しなかった事実」ではなく**昇格の実行**の記録
'   へ変わる（従来の `sparring_tier_not_promoted` は残さない）。
Public Function SendSparring(ByVal caseId As String, ByVal utterance As String, _
                             ByRef replyText As String, ByRef errCode As String) As Boolean
' 発話1本の送信（15章§6.5）。True＝応答を受け取り、発話と応答を case_data へ保存できた。
'   errCode: E0103＝送信前のPII検知で遮断（16章 E-05(3)。`CallChat` の**前**に `modPii` を
'   通すのは本関数の責務）/ E0101＝前提不足 / E0604＝履歴の保存失敗 / E02xx＝`CallChat` が
'   帯域外で返した経路失敗をそのまま透す。16章 E-44 の「往復数を減らして再開」の案内は
'   E0204 のときに usage_log へ残す。
Public Function SendToInbox(ByVal caseId As String, ByVal roleKind As String, _
                            ByVal seqNo As Long) As String
' 選択した発話を受信箱へ登録し inbox_id を返す（15章§6.5「受信箱へ」。source_kind=
'   `field_voice` / theme=案件ID＋発話の要約）。失敗は ""。二重送信の抑止は 13章§2.17 の
'   `inbox_id` 列（壁打ちシート側）が鍵なので本関数は持たない。本文にPIIを検知したら
'   登録しない（16章 E-05(2)）
Public Function HistoryOf(ByVal caseId As String, ByVal roleKind As String) As String
' 保存済み履歴を**保存形式のまま**返す（ui が seq / spoke_at / 本文へ分解して 13章§2.17 の
'   `sparring_log` を描く）。roleKind は 13章§2.17 の enum（user / ai）。表に無い値は ""。
' **保存形式**（13章§2.2 の `sparring_u` / `sparring_a` の中身。読み書きの唯一点は本モジュール）:
'   1発話＝1行＝`seq <TAB> spoke_at <TAB> 本文`、行区切りは vbLf、本文は
'   `modJsonLite.EscapeJsonStr` で `\n` `\t` `\\` を畳む。seq は 13章§2.17 と同じ**発話単位の
'   通し連番**で user と ai が1本の番号列を共有する（発話=n / その応答=n+1）。
'   case_data の `seq` 列は §2.2 の定義どおり**分割連番**のままである（`SaveData` は data_key
'   単位で全行を置換する契約であり、1発話＝1物理行を持たせる口が無い）。§2.2 の列定義と
'   data_key 注記「発話単位seqで保存」の食い違いはこの行形式が吸収する。
' --- modSparring の純核3本（裁定書8 B-9。シート・ログ・LLMに触れない純関数） ---
' 送信可否・履歴上限・線上形式への変換は「規約そのもの」であり、実行制御の中に閉じ込めると
'   層(a)から誰も検査できない。modTestsPure から直接叩く前提で公開し、run_lo_tests の
'   PURE_ALLOWLIST にも modSparring を登録する。**Private へ戻すことは契約違反**。
Public Function CanContinueSparring(ByVal caseIdText As String, ByVal utterance As String, _
                                    ByVal hasPii As Boolean) As Boolean
' 発話を1本送ってよいかの**唯一の判定点**（fail-closed）。(1) caseIdText が 13章§1 の案件ID
'   書式（判定は `modCaseStore.IsValidCaseId`）(2) 発話が空白・改行だけでない (3) hasPii=False
'   （16章 E-05(3) は壁打ちの発話送信前の検知で**送信をブロック**する）。走査そのものは
'   `modPii` が唯一の実装なので、ここは結果の真偽だけを受け取る（検知規則を2箇所に書かない）
Public Function TrimHistoryOf(ByVal storedText As String, ByVal maxTurns As Long) As String
' 保存形式の履歴を**直近 maxTurns 発話**へ切り詰める（古い順のまま返す）。16章 E-44 の
'   「渡す履歴を直近 `sparring_max_turns` 往復に制限」を保存形式の側で行う唯一の点で、
'   全履歴は case_data に残る。maxTurns<=0 は全件。線上形式（";;;"連結）側の最終防衛は
'   `modGatewayRPN.TrimHistoryPairs` が別に持つ（形式が違うので同じ実装は使えないが、
'   **件数の値はどちらも config `sparring_max_turns` の1箇所**から来る）
Public Function HistoryJoinOf(ByVal storedText As String, ByVal maxTurns As Long) As String
' 保存形式から `CallChat` の histU / histA を組む唯一の点。直近 maxTurns 発話を**新しい順**に
'   `modGatewayRPN.GW_HIST_SEP`（";;;"）で連結する（切詰めは `TrimHistoryOf` に委ねる）。
'   本文はエスケープを解いて原文へ戻し、本文中に区切りが現れたら ";" へ潰す（線上形式だけの
'   非可逆処理。case_data 側の原文は書き換えない）

' === app: modCaseRead（案件一覧の読取専用API。裁定書7 B-7）===
Public Function ReadCaseCtx(ByVal caseId As String, ByRef ctx As TCaseCtx, _
                            ByRef roundNo As Long, ByRef qualityMode As String, _
                            ByRef s4Variant As String, ByRef dossierTier As String) As Boolean
' 13章§2.1『案件一覧』の1行から「Stepを実行するのに要る文脈」を1回の読取で取り出す**唯一の口**。
'   R4でシートに触れない modPipeline / modCompanyFile は必ずこれを通す（読取APIが無いまま
'   既定値で走らせると、企業名が空・case_type=new 固定のプロンプトでAI利用枠を消費し、
'   誤った sN_json を確定してしまう）。**書込は一切持たない**（状態遷移・採番・case_data は
'   modCaseStore が唯一の口。R4許可は「案件一覧の読取だけ」に限る）。
'   ctx=13章§2.1の10列（TCaseCtx）/ roundNo=round_no（空・不正は1）/ s4Variant=s4_variant
'   （空は proposal）/ dossierTier=dossier_tier（空は t1_quick。ctx.dossier_tier と同値）/
'   qualityMode=config `quality_mode`（案件一覧に列は無い。空はティア連動＝ResolveQualityMode が解決）。
'   戻り値 False=シート・見出し・当該行が無い（呼び出し側は fail-closed で中止する）
' `last_ok_step` / `failed_step` の**書込**口（16章E-06）は modCaseStore.SetStepOutcome
'   （本節の modCaseStore の項）。裁定書8 A-2 で新設し、modPipeline の成功経路・失敗経路
'   から結線済み（usage_log への退避は廃止した）

' === app: modPii（PII走査の本体。16章E-05・12章§2/§4。裁定書7 B-5）===
' 走査の実施点は16章E-05の一覧（modUICase / modUIInbox / modSparring / modJudgeStore /
'   modExportHtml / modCompanyFile）が正。本モジュールは**判定の唯一の実装**であり、
'   検知規則（@付き・電話番号・敬称つき人名）を2箇所に書かない。
Public Function HasPii(ByVal sText As String) As Boolean          ' 1件でも検知したか
Public Function DetectionCount(ByVal sText As String) As Long     ' 検知件数
Public Function KindsOf(ByVal sText As String) As String          ' 検知種別（";"区切り。例 "mail;phone"）
Public Function ScanReport(ByVal sText As String, ByVal whereNote As String) As String
' 走査結果の1行表現（"箇所名|種別@文字位置" の列挙）。**本文を含めない**（16章NFR-S3。
'   err_log / dossier_meta へそのまま記録できることが契約）
Public Function MaskText(ByVal sText As String) As String
' 検知箇所を `{{PERSON}}` 等へ置換した伏字案（16章E-05(5) の customer_quote の差し替え案）。
'   置換した文面を返すだけで、保存の採否は呼び出し側（利用者の選択）が決める

' === app: modCompanyFile（企業ドシエファイル。13章§2.8・FR-45。裁定書7 B-5）===
Public Function CompanyFilePath(ByVal company As String, ByVal caseId As String, _
                                ByVal dirPath As String) As String
' 13章§2.8のファイル名規則で決まる絶対パス（生の company を使わず SanitizeFileName 系を通す）
' **8桁は company 由来**（裁定書9 B7。13章§2.8 手順4の但し書き）: 企業ドシエファイルに限り
'   `Left$(Fnv1a64Hex(NormalizeForHash(company)), 8)` を使う（`ExportCompanyFile` が
'   `dossier_meta.company_id` を作るのと同じ式＝値源は1つ）。case_id 由来では同じ会社の2件目の
'   案件が必ず別ファイルになり、FR-45「1社1ファイル・追記して育てる」が成立しないため。
'   HTMLレポート・PPTは従来どおり case_id 由来のままとする（1案件1出力であり衝突回避が目的）
Public Function ScanCaseForPii(ByVal caseId As String) As String
' 書き出す予定の中身をまとめて modPii へ通す（16章E-05(7)）。""=検知なし
Public Function ExportCompanyFile(ByVal caseId As String, ByVal dirPath As String, _
                                  ByVal confirmedAt As String) As String
' 現ラウンドを追記書き出し（HOMEの[保存]）。戻り値=書き出した絶対パス。失敗・中止は ""。
'   company / industry_code / dossier_tier / round_no は modCaseRead.ReadCaseCtx で読む
'   （読めなければ書き出さない）。confirmedAt=PII検知に対し利用者が「確認した」を選んだ日時。
'   検知があるのに confirmedAt が空なら**書き出さない**（16章E-05(7)）
'   **保存の失敗を成功として返さない**（裁定書9 B8・N3）: 下位の `modCompanyFile2.DossierSaveAndClose`
'   は **`Public Function ... As Boolean`** へ改め、`SaveAs` の失敗を呼び出し側へ返す（共有フォルダの
'   読取専用・他者ロック・パス長超過で現実に起きる）。False のときは `VerifyRoundTrip` へ進まず
'   **`ExportCompanyFile = ""`** を返し、ui層は「保存できませんでした」を表示する。
'   `VerifyRoundTrip` は**対象ブックが閉じていることを確認してから**開き直す（同一プロセスで
'   開いたままのブックを読むと、ディスクではなくメモリ上の未保存内容と突合して合格してしまう）
Public Function ImportCompanyFile(ByVal filePath As String, ByVal caseId As String) As Boolean
' 最新ラウンドの s1/s2 と notes を案件へ復元（HOMEの[開く]）。**その枠が空のときだけ書く**
' `modCompanyFile2` は 30,000字契約による分割先（ブック・シートの下位I/O 14本）。
'   **本節の公開契約面には載せない**（modCompanyFile の下位実装であり、呼んでよいのは
'   modCompanyFile だけ。vba_lint の CONTRACT は required=[] で登録する）。
'   ただし **`DossierSaveAndClose` の戻り値の型だけは本節の裁定事項**であり、裁定書9 N3 で
'   `Public Sub` から **`Public Function ... As Boolean`**（True=SaveAs とClose が成功）へ
'   変更した。下位実装であっても「失敗を握り潰さない」ことは§6冒頭のエラー規約そのものである

' === app: modCaseStore / modInboxStore / modJudgeStore ===
Public Function NewCase(ByVal company As String, ByVal industryCode As String, ByVal caseType As String) As String
Public Function SaveData(ByVal caseId As String, ByVal dataKey As String, ByVal content As String) As Boolean
' 13章§2.2 の縦持ち保存（32,000字分割）。**2相書込**とする（裁定書9 B13）: 新しい断片を先に
'   書き切ってから旧行を消す。途中で失敗したら旧行を1行も消していない状態で False を返す
'   （「削除してから書く」順序を禁じる。VBAの行削除は Undo できず、途中失敗で旧内容が
'   消えたまま短いJSONだけが残る経路を構造として持たない）。seq 帯の取り方は実装裁量だが、
'   **途中失敗で旧データが残ること**を層(a)のテストで実証する
Public Function LoadData(ByVal caseId As String, ByVal dataKey As String) As String
' 保存された断片を seq 昇順に連結して返す。**完全性を検査する**（裁定書9 B13）: seq 1..maxSeq が
'   1つでも欠けていたら詰めて返さず、E0604 を記録して "" を返す（切れたJSONを正常値として
'   返さない＝fail-closed）。当該 data_key が1行も無い場合は従来どおり "" （欠損ではない）。
'   **仮seq帯の残留検査**（v2.5.1・裁定書10 M6）: 「本seq帯が空だが仮seq帯（2相書込の作業帯）に
'   行が残っている」状態を検査し、該当時は E0604 を記録して "" を返す。相2（旧行削除）成功後・
'   相3（seq帯確定）前に落ちたケースを「データが無いこと」と区別できない静かな消失にしない
Public Function ResolveStepJson(ByVal caseId As String, ByVal stepNo As Long) As String
' 下流Stepが参照すべきJSONを一元解決する（優先順の正は13章§2.2）。
' N=2,3 は sN_edited > sNr_json > sN_json、S1/S4 は sN_edited > sN_json。呼び出し側で個別に分岐しない
Public Function SetStatus(ByVal caseId As String, ByVal status As String) As Boolean
' 案件一覧 `status`（13章§2.1の8値）の唯一の書込口。**`CanTransition` を通さない**（裁定書9 B16(a)）。
'   理由: 起動時の状態修復（`RepairStates` / `ApplyRepairedState`。16章 E-12）と失敗時の `error`
'   書込は、遷移表に無い並びで書く必要がある修復系であり、ここで遷移検査を掛けると修復自身が
'   弾かれる。本関数が課すのは enum 検査だけであり、**遷移の整合は `RepairStates` が担う**
'   （`CanTransition` は11章§4の遷移表の宣言であり、修復の期待値を層(a)で検査するための純核）。
'   `exported` / `feedback_done` の書込点は次の2つに固定する（裁定書9 B16(b)）:
'     `exported`      = `modUIHome.HomeExportHtml` の**成功分岐**（HTMLレポートの生成成功時）
'     `feedback_done` = `modUICase4.FeedbackSave` の**成功分岐**（フィードバック保存の成功時）
'   これにより `CS_STATUSES_ABOVE_S4` の降格抑止（s4_done より上の状態から巻き戻さない）が
'   到達可能になる。出力・記録の失敗は状態を動かさない（16章 E-48）
Public Function PromoteTier(ByVal caseId As String, ByVal tierText As String) As Boolean
' 案件一覧 `dossier_tier` の**唯一の書込口**（裁定書9 N2・A-1。13章§2.1/§2.17 の t3_sparring
'   自動昇格の実体）。`tierText` は19章§3の enum（t1_quick / t2_full / t3_sparring）のみ受け付け、
'   不一致は**1列も書かず** False（E0101 を記録）。案件行が無いときも False。
'   **status は動かさない**（ティアは案件の属性であって状態ではない）。案件入力の画面から
'   ui層が `ci_dossier_tier` を書く経路（13章§2.1 の属性列の書込経路）とは別に、app層から
'   昇格する必要があるためここに置く。**方向（昇格か降格か）は本関数では判定しない**
'   （enum に合致する値をそのまま書く）。app層からの呼び出しは `modSparring.ResumeSparring`
'   の t3_sparring 昇格の1点だけであり、そこ以外から呼ばない
Public Function SetStepOutcome(ByVal caseId As String, ByVal lastOkStep As Long, _
                               ByVal failedStep As String) As Boolean
' 16章 E-06 が要求する案件一覧の `last_ok_step` / `failed_step` の【書込口】（裁定書8 A-2で
'   新設）。modPipeline の成功経路が (stepNo, "")、失敗経路が (-1, "sN") で呼ぶ。
'   lastOkStep: 0～4 を書く。**負値は「更新しない」**＝E-06 の「失敗時は last_ok_step を
'     更新しない」を、呼び出し側の分岐ではなく引数で表す。4を超える値は E0101 で拒否。
'   failedStep: "" は失敗の記憶を消す（13章§2.1「Step成功時に空へ戻す」）。非空は enum
'     s1 / s2 / s3 / s4 / s2c / s3c のみ受け付け、表に無い値は E0101 で拒否して1列も
'     書かない。**status は動かさない**（状態遷移の唯一の口は SetStatus）
Public Function SetReportPath(ByVal caseId As String, ByVal pathText As String) As Boolean
' 18章§1.1⑦ が要求する案件一覧 `report_path` の【書込口】（W3.1で宣言）。`modExportHtml` は
'   12章R4によりシートに触れないため、確定パスの記録は必ずここを通す。
'   **status は動かさない**（出力の成否は状態遷移に影響しない＝16章 E-48）。書込に失敗しても
'   生成済みのHTMLファイルは残るので、呼び出し側は警告に留めて生成を成功として扱う
'   （戻り値 False は「記録できなかった」であって「出力できなかった」ではない）。
'   pathText は `modUtilText.SanitizeFileName` を通したあとの**確定フルパス**
Public Sub InvalidateDownstream(ByVal caseId As String, ByVal fromStepNo As Long)
' 16章 E-10 の下流無効化。**`last_ok_step` は下げる方向にしか動かさない**（裁定書9 B20）:
'   書き込む値は `Min(keepStep, 現在の last_ok_step)` とする（`ApplyRepairedState` の冒頭で
'   丸める）。無効化の呼び出しで案件が昇格しうる経路を残さない
Public Function RepairStates() As Long                 ' 起動時整合修復（16章E-12・12章§2.1のmodBoot手順③）。戻り=修復件数
Public Function FreezeRound(ByVal caseId As String) As Long
' ラウンド確定（FR-35マルチラウンド）。s2（edited優先で解決した1本）を data_key `s2_prev_json` へ
' 退避し、案件一覧の round_no を +1 して新しい round_no を返す。次ラウンドのS2は
' BuildS2User の prevS2Json にこの退避分を渡す
Public Function NewInboxItem(ByVal sourceKind As String, ByVal theme As String, ByVal body As String) As String
' 投函を1件起票して inbox_id（13章§1 `I-YYYYMM-NNN`）を返す。当月の使用済み最大連番の次から
'   採り、衝突は E0605 を記録して次番号へ（999で枯渇）。失敗は ""。`posted_by_group` は本節の
'   シグネチャが受け取らないため空のまま起こす（`NewCase` の channel 等と同じ扱い）。
'   body の32,000字打切りと先頭式記号の無害化は `SetCellSafe` が行う（13章§2.6・NFR-S7①）
Public Function SetInboxJudgement(ByVal inboxId As String, ByVal status As String, _
                                  ByVal dropType As String, ByVal reviveTag As String, ByVal reviveDue As Date) As Boolean
' 判定（統制語彙）を記録する。status は `adopted` / `conditional_hold` / `rejected` のみ。
'   16章 E-41 の必須検査（`JudgementError`）に掛かったら**1列も書かない**（保存ブロック）。
'   遷移可否は `CanInboxTransition` が唯一の判定点。判定に対応しない列（却下でない
'   `drop_type` 等）は空へ戻す（前の判定の語彙を行に残さない）。
'   **`merged` は受け付けない**: 13章§2.6 は merged に `merged_into` を必須とするが、本
'   シグネチャは統合先を受け取る引数を持たない。空の `merged_into` を書くと13章の必須を
'   満たさない行ができるため fail-closed で拒否し E0101 を記録する（引数の追加は本節の裁定事項）
'   **判定の入力口（裁定書9 B2・N5）**: ui層は `status` 列ではなく受信箱シートの入力列
'   **`judge_to`**（13章§2.6）を読み、その値を第2引数 `status` へ渡す。`status` 列は判定の
'   **結果**を表す列であり利用者に触らせない（列見出しの header_note に明記）。本関数の
'   `CanInboxTransition` は「現在の `status`（diagnosed）→ `judge_to` の値」を検査するため、
'   from と to が同一セル由来になって自己遷移で必ず False になる経路が構造として消える。
'   成功時に ui層は `status` の書き換え結果を再描画し、**`judge_to` を空へ戻す**（判定済みの
'   行に入力値を残さない）。本関数のシグネチャは変更しない
Public Function SavePfResult(ByVal inboxId As String, ByVal pfJson As String, _
                             ByVal survival As String, ByVal predTypes As String) As Boolean
' プリフライト診断の結果を格納する（13章§2.6）。あわせて status を undiagnosed →
'   diagnosed へ進める（可否は `CanInboxTransition`）。判定済みの行は診断結果だけを
'   上書きし status は動かさない（巻き戻さない）
Public Function ReadInboxItem(ByVal inboxId As String, ByRef theme As String, _
                              ByRef body As String, ByRef statusText As String) As Boolean
' 1件の投函を読む【唯一の口】。modPlayOps は R4 でシートに触れないためここを通す。
'   戻り値 False=シート・見出し・当該行が無い（呼び出し側は fail-closed で中止する）
Public Function UndiagnosedIds() As String
' 未診断（undiagnosed）の inbox_id を投函順に ";" 区切りで返す（一括診断の対象一覧）
Public Function InterestText() As String
' 関心度の集計表示（10章 FR-17・11章 受信箱ワイヤー「関心度: 熊対策12件 雹災5件」）。
'   受信箱の `theme` 列を投函順に**1件1行**で集め、集計そのものは純核
'   `InterestSummaryOf` に渡すだけ（件数集計・降順・2件以上・上限件数の規約を1つも
'   持たない）。上限件数は純核が持つため**引数で受けない**（v2.4.8 で `maxItems` を廃止）。
'   本体シートの読取だけで完結しナレッジブックへは書かない（12章§4）
' --- modInboxStore の純ロジック（Excel非依存。層(a)から直接叩く。裁定書8 B-7）---
' 規約（採番・ID書式・状態遷移・E-41の必須検査・FR-17の集計）をシートI/Oの中へ閉じ込めると
'   層(a)から誰も検査できない。以下7本は modTestsPure から直接叩く前提で公開し、
'   run_lo_tests の PURE_ALLOWLIST にも modInboxStore を登録する。**Private へ戻すことは
'   契約違反**（vba_lint の CONTRACT required が検出する。裁定書9-3）。
Public Function BuildInboxId(ByVal monthText As String, ByVal seq As Long) As String
' 13章§1 の `I-YYYYMM-NNN`。monthText は yyyymm の6桁ちょうど（数字のみ）、seq は 1..999。
' 桁違い・範囲外は ""。引数名が `monthText` なのは **`Month` がVBAの組込関数**で
' `vba_lint.py` の予約語検査がERRORにするため（`BuildCaseId` の `dayText` と同じ理由）
Public Function IsValidInboxId(ByVal id As String) As Boolean
' `I-` + 数字6桁 + `-` + 数字3桁（連番は 001..999）ちょうどの形か。前後空白は許さない
Public Function CanInboxTransition(ByVal fromStatus As String, ByVal toStatus As String) As Boolean
' 11章§4 の受信箱ステータス遷移表。許すのは (1) undiagnosed→diagnosed (2) diagnosed→
'   adopted / conditional_hold / rejected / merged の2種だけ。自己遷移・判定済みからの
'   再判定・診断を飛ばした判定は False。13章§2.6 の enum に無い値はどちらの側でも False
Public Function JudgementError(ByVal statusText As String, ByVal dropType As String, _
                               ByVal reviveTag As String, ByVal hasDue As Boolean) As String
' 16章 E-41（統制語彙の必須化）の唯一の判定。保存してよければ ""、止めるなら理由の1行。
'   rejected は `drop_type`（T0～T10）必須、conditional_hold は `revive_tag`（5値）と
'   見直し期日の両方が必須。`hasDue` は Date 型を純関数へ持ち込まないための真偽（呼び出し側が落とす）
Public Function InterestKeyOf(ByVal themeText As String) As String
' 関心度集計のテーマキー（FR-17）。改行・空白の揺れと大小文字だけを吸収する（意味の同一視は
'   しない＝人が読める粒度で数える）。空テーマは ""
Public Function FmtInterestLine(ByVal themeText As String, ByVal itemCount As Long) As String
' 関心度1件の表示（`熊対策12件`）。件数0以下・テーマ空は ""
Public Function InterestSummaryOf(ByVal themeLines As String) As String
' 10章 FR-17 の関心度集計の**唯一の値源**（裁定書9-1）。`themeLines` は投函1件につき1行
'   （vbLf区切り）のテーマ文字列。規約はこの1本にしか無い: (1)同一テーマの重複投函を
'   棄却せず件数へ寄せる (2)件数の多い順に並べる (3)**2件以上集まったテーマだけ**を
'   載せる（1件は「重複投函」ではない） (4)上限は既定3件 (5)区切りは半角空白。
'   同一視の粒度は `InterestKeyOf`、1件ぶんの表示は `FmtInterestLine` が唯一の値源。
'   空・全行が空テーマなら ""（空の集計をでっち上げない）
Public Function NewJudgement(ByVal rec As TJudgement) As String
' UW判断を1件起票して judge_id（13章§1 `J-YYYYMM-NNN`）を返す（裁定書8 B-8）。当月の
'   使用済み最大連番の次から採り、衝突は E0605 を記録して次番号へ（999で枯渇）。失敗は ""。
'   `TJudgement`（modAppTypes）は13章§2.7の11列から `judge_id`/`judged_at` を除いた9列。
'   必須（line_id/situation/decision/key_reason/recorded_by）が1つでも空なら1列も書かない。
'   `decision` は19章§3のenum（raise/close/restrict/keep/improve）のみ受け付ける。`result`
'   は任意だが非空なら13章§2.7のenum（won/lost/pending）のみ受け付ける。
'   **16章E-05(4)**: 保存直前に `situation`/`key_reason` を `modPii.HasPii` へ通し、検知したら
'   1列も書かず `modPii.ScanReport` の返り値（本文を含まない）を E0103 の detail に記録する
'   （伏字差し替えの例外は無い＝(1)～(4)と同じブロック仕様）
Public Function ReadJudgement(ByVal judgeId As String, ByRef rec As TJudgement) As Boolean
' 1件の判断を読む唯一の口（17章T-26 DoDの「入力→保存→再表示一致」）。戻り値 False=シート・
'   見出し・当該行が無い（rec は全列空へ戻す。呼び出し側は fail-closed で中止する）
Public Function SetJudgementResult(ByVal judgeId As String, ByVal resultText As String, _
                                   ByVal postLoss As String) As Boolean
' 事後結果（result/post_loss）だけを更新する（FR-22）。situation/decision/key_reason 等の
'   起票内容は書き換えない。resultText が非空なら13章§2.7のenumのみ受け付け、不一致は保存
'   ブロック。""を渡すと result 列を空へ戻す（pendingの取消）。result/post_loss は16章E-05(4)
'   の走査対象外（situation/key_reasonの2列に限る）なので modPii は通さない
' --- modJudgeStore の純ロジック（Excel非依存。層(a)から直接叩く。裁定書8 B-8）---
' **Private へ戻すことは契約違反**（vba_lint の CONTRACT required が検出する。裁定書9-3）。
Public Function BuildJudgeId(ByVal monthText As String, ByVal seq As Long) As String
' 13章§1 の `J-YYYYMM-NNN`。monthText は yyyymm の6桁ちょうど（数字のみ）、seq は 1..999。
'   桁違い・範囲外は ""。`modInboxStore.BuildInboxId` と同型
Public Function IsValidJudgeId(ByVal id As String) As Boolean
' `J-` + 数字6桁 + `-` + 数字3桁（連番は 001..999）ちょうどの形か。前後空白は許さない。
'   19章§4の注記どおり判断基準ID（`J-NN`）とは桁数が異なるため字数で区別できる
Public Function IsValidDecision(ByVal decisionText As String) As Boolean
' 19章§3の decision enum（raise/close/restrict/keep/improve）に一致するか。空は False
Public Function IsValidJudgeResult(ByVal resultText As String) As Boolean
' 13章§2.7の result enum（won/lost/pending）に一致するか。空は False（resultは任意列なので、
'   空を許すかどうかの判断は呼び出し側が「空なら検査自体をスキップする」形で行う）
' --- modCaseStore の純ロジック（Excel非依存。層(a)から直接叩く。裁定書6 項目7）---
' 採番・参照優先・状態遷移は「規約そのもの」であり、シートI/Oの中に閉じ込めると誰も検査
' できない（W2aでは参照優先の並びを入れ替えてもテストが1本も落ちなかった）。以下4本は
' シートを1行も触らないので、シート側の関数は必ずこれらを通す（答えを2箇所に書かない）。
Public Function BuildCaseId(ByVal dayText As String, ByVal seq As Long) As String
' 13章§1 の `C-YYYYMMDD-NNN`。dayText は yyyymmdd の8桁ちょうど（数字のみ）、seq は 1..999。
' 桁違い・範囲外は "" を返す（呼び出し側が枯渇・不正として扱う。例外は投げない）。
' 第1引数名が `dayText` なのは **`datePart` がVBAの組込関数 `DatePart` と衝突する**ため
' （`vba_lint.py` の予約語検査がERRORにする。VBA制約が命名に優先する）
Public Function IsValidCaseId(ByVal id As String) As Boolean
' `C-` + 数字8桁 + `-` + 数字3桁（連番は 001..999。000 は不正）ちょうどの形か。前後空白は許さない
Public Function CanTransition(ByVal fromStatus As String, ByVal toStatus As String) As Boolean
' 11章§4 の案件ステータス遷移表 + 16章 E-12。許すのは次の3種だけ:
'   (1) 正順の1段進み: draft→s1_done→s2_done→s3_done→s4_done→exported→feedback_done
'   (2) **任意の状態から error へ**（失敗はどこでも起こりうる）
'   (3) error からの復帰: `failed_step` の再実行で戻る先＝draft..feedback_done の**いずれか**
'       （完了済みStepの結果を保持したまま再導出するため。10章NFR-R3・16章 E-12(1)）
' 同じ状態への遷移（自己遷移）と、段飛ばし・巻き戻し（error 経由を除く）は False。
' 8値のenum（13章§2.1）に無い値はどちらの側でも False
Public Function ResolveDataKey(ByVal stepNo As Long, ByVal hasEdited As Boolean, _
                               ByVal hasRevised As Boolean, ByVal hasJson As Boolean) As String
' 13章§2.2 の参照優先の**純核**。その3つの有無から「実際に読むべき data_key 名」を1つ返す。
'   N=2,3: sN_edited > sNr_json > sN_json ／ N=1,4: sN_edited > sN_json（改訂は無いので
'   hasRevised は無視する）。どれも無ければ ""。範囲外の stepNo も ""。
' `ResolveStepJson` は必ずこの関数の答えに従う（分岐を2箇所に書かない）
' `modCaseStore2` は 30,000字契約による分割先（案件一覧 と case_data の下位シートI/O。
'   シートを取る・最終行・矩形読み・行削除・列名で1セル書く・セル値をLongへ、の11本）。
'   **本節の公開契約面には載せない**（modCompanyFile2 と同じく modCaseStore の下位実装で
'   あり、呼んでよいのは modCaseStore だけ。vba_lint の CONTRACT は required=[] で登録する）

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

' === ui: modUICase（enum変換表。11章§5「日本語ラベル⇔enumの変換は modUICase の共通変換表
'      （19章と一致必須）のみで行う」の実体。W3.1で宣言＝命名権を本節へ戻した） ===
Public Function EnumPairsCsv() As String
' 19章§3の変換表そのもの（`グループ,機械値,日本語` を vbLf 区切りで返す）。**この1本だけが
'   値の出どころ**であり、`EnumJa` / `EnumEn` / `EnumLabels` / 入力規則の隠しレンジはすべて
'   この戻り値を走査して答える。並び順も19章§3の記載順であること。
'   **`tools/enum_check.py`（17章§4-2の一致検査）が静的評価する照合先が本関数**なので、
'   Private化・改名すると検査が無言で対象を失う。行を足す・直すときは19章§3を直してから
'   `python3 tools/enum_check.py --dump-bas` の出力で差し替える（手で写さない）
Public Function EnumJa(ByVal groupName As String, ByVal enumValue As String) As String
' 機械値 → 日本語ラベル。表に無い組み合わせは `""`
Public Function EnumEn(ByVal groupName As String, ByVal labelText As String) As String
' 日本語ラベル → 機械値。表に無いラベルは `""` を返し、**推測で近いものを返さない**
'   （13章§2.2「表に無いラベルは検証不合格」。呼び出し側は `""` を受けたら原文をそのまま
'   JSONへ載せ `modValidate` に弾かせる＝黙って直さない）

' === ui: modUICase5（ui層内部ヘルパ。W4.1分割裁定＝30,000字契約による modUICase2 の分割先。
'      v2.5.1・裁定書10 §1（M2解消）で本節へ登記） ===
' 呼んでよいのは modUICase2 だけ（ui層内部の下位ヘルパであり、公開契約面の入口ではない）。
'   変換表（19章§3）と列定義（13章§2.12-§2.15）は持たない（modUICase / modUICaseFmt が唯一持つ）。
Public Function SerializeBody(ByVal stepNo As Long) As String
' シート → JSON の逆シリアライズ本体（13章§2.2 規約1/2/3/5）。組めなければ ""。
'   例外の捕捉と E0302 の記録は呼び出し側 modUICase2.SerializeStep が持つ（捏造しない口を1本に保つ）
Public Function ColIndexes(ByVal ws As Object, ByVal headerRow As Long, _
                           ByVal colSpec As String) As Variant
' ブロックの列引き当て（ヘッダ行から colSpec の各列番号を解決）。描画側（modUICase2）と
'   読取側（本モジュール）が**同じ1本**を呼ぶことで、列の引き当て方が2箇所へ分かれない
Public Function ColCount(ByVal colSpec As String) As Long   ' colSpec の列数
Public Function RoomOf(ByVal anchorName As String) As Long
' ブロックの部屋数。13章§2.9「行番号を仮定しない」に従い、次のブロックのアンカー行から
'   動的に決める（確保行数を定数で持たない）

' === app: modExportHtml / modExportPpt / modExportHearing ===
Public Function GenerateHtmlReport(ByVal caseId As String, ByRef outPath As String) As String
    ' ""=成功 / 非空=失敗理由（コードは E0502。16章E-48）。S1+S2+S3のJSONを固定HTMLテンプレート
    ' (高橋PLプロト準拠・10章FR-37)に流し込み、自己完結HTML 1ファイルを出力(LLM不使用)。
    ' v2.3で主力出力(旧GenerateReportを置換)
    ' テンプレ本体は modHtmlTemplate1..n（純文字列・R4・12章§2 app層）、テーマCSSは modHtmlTheme
    ' （config `html_theme`。既定 standard）。出力先は config `html_out_dir`
    ' ファイル名は modUtilText.SanitizeFileName を通し、確定パスを案件一覧 report_path に記録する
    ' **上書きしない**（裁定書9 B4。13章§2.8）: 名前の末尾に `_<yyyymmdd>`（`IsoDateCompact`）を
    ' 付け、同名が既に存在する場合は `_2` `_3` と連番を探して**新規ファイルとして作る**。
    ' 既存ファイルへ `adSaveCreateOverWrite` で書かない（利用者が手で注記を入れた前回HTMLを消さない）
    ' **PII走査の失敗を黙らせない**（裁定書9 B17。18章§1.1(3)）: 走査は `On Error GoTo` ラベル方式で
    ' 包み、失敗しても生成は続行するが `meta.warnings` へ「個人情報の走査に失敗しました。配布前に
    ' 本文をご確認ください」を積む（プロシージャ冒頭の `On Error Resume Next` で3本の走査を
    ' まとめて覆い、警告も E0103 も出ないまま合格に見える経路を残さない）
    ' **文字コード**: 書き出しは `ADODB.Stream`（Charset="utf-8"・BOMあり）。VBAの Open/Print # は
    ' CP932で書かれ非CP932文字が "?" 化するため使わない。テンプレ先頭に <meta charset="utf-8"> を必ず含める
    ' **埋め込み**: JSONは「1本のJS文字列リテラル＋JSON.parse」形式で埋め、modUtilText.JsStringSafe を
    ' 必ず通す。素のJSリテラル直書きは禁止。HTML本文に差し込む値は HtmlSafe を通す（16章E-47）
Public Function BuildMetaJson(ByVal caseId As String, ByVal company As String, _
                              ByVal industryCode As String, ByVal industryName As String, _
                              ByVal caseType As String, ByVal dossierTier As String, _
                              ByVal qualityMode As String, ByVal roundNo As Long, _
                              ByVal s4Variant As String, ByVal generatedAt As String, _
                              ByVal appVersion As String, ByVal themeName As String, _
                              ByVal warnText As String) As String
' 18章§2 の `meta` オブジェクトを1本のJSON文字列として組み立てる**純関数**（W3.1で宣言。
'   Excel・configに触れず、値はすべて引数で受け取る＝層(a)から叩ける）。値の由来は18章§2の
'   とおり（案件一覧の同名列／`hm_quality_mode`／config `app_version`・`html_theme`）。
'   `warnText` は18章§2の `meta.warnings`（16章 E-05 のPII検知・E-31 の復元失敗など、
'   生成をブロックしない警告）。**vbLf区切りの0本以上**を受け取り、空なら空配列を書く
Public Function BuildReportHtml(ByVal metaJson As String, ByVal s1Json As String, _
                                ByVal s2Json As String, ByVal s3Json As String, _
                                ByVal themeName As String) As String
' 18章§1.1④⑤ を1本にした**純関数**（W3.1で宣言）。`meta`/`s1`/`s2`/`s3` を§2のDATAへ
'   組み、`modHtmlTemplate1.BuildDocument` へ渡してHTML全文を返す。未実行のStepは `null`
'   （キー自体は必ず置く＝§2）。`""` を返したら組立失敗＝E0502（16章 E-48）。
'   ファイル書出・`report_path` 記録・PII走査は含まない（`GenerateHtmlReport` の責務）
Public Function GeneratePpt(ByVal caseId As String, ByVal s4Json As String, _
                            ByVal variant As String, ByRef outPath As String) As String ' ""=成功。variant=proposal/alliance。Phase 1.5
Public Function BuildHearingSheet(ByVal caseId As String) As Boolean
' 13章§2.16 のヒアリングシートを S4 から再生成する（LLM不使用）。**既存行を消してから書く**ため、
'   訪問後に `answer_memo` へ書き込まれた手書き回答は失われる。VBAの書込は Undo できないので、
'   呼び出し側（ui層）が下の `AnswerMemoCount` で事前に数え、1行以上なら確認を挟む
Public Function AnswerMemoCount(ByVal caseId As String) As Long
' ヒアリングシートの `answer_memo` 列の**非空行数**（裁定書9 N8・B12。契約はv2.5.1・裁定書10 M5で改訂）。
'   シートが無い・見出しが無い場合は 0（読めないことを「回答あり」と誤認しない）。
'   **`hs_case_id` が引数 caseId と一致しなくても数える**（ヒアリングシートはブックに1枚しか
'   なく、別案件の手書き回答こそ守るべき対象。旧契約の「当該案件のシートでない場合は 0」は
'   守るべき条件でちょうど素通りする fail-open だったため廃止）。
'   `modUIHome` はこの値が 1 以上のとき `MsgBox`（vbYesNo）で上書き確認を出し、No なら
'   `BuildHearingSheet` を呼ばずに中止する（16章 E-10 の下流無効化と同じ作法）。確認文言は
'   `hs_case_id` を読み分ける: caseId と不一致かつ 1 以上なら
'   「別案件（<hs_case_id>）の手書き回答が<N>行残っています。作り直すとこの回答は消えます。」
'   （<N>=本関数の戻り値。裁定書12 V9で実物〔modUIHome〕へ逐語化）。
'   本関数は数えるだけで、シートを1セルも書き換えない

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

' === test: modTestsExcel2 ===
Public Function RunExcelTests2() As Long
' 30,000字契約(12章§2)による modTestsExcel の分割先の入口(v2.5.2・裁定書11)。
' 呼んでよいのは modTestsExcel.RunAllExcelTests のみ(wintest からの入口は
' 従来どおり RunAllExcelTests の1本)。戻り値=本モジュールが打った Check の本数で、
' 呼出側の本数自己照合(TE_EXPECTED)へ合流させる
```

- **`modHtmlTemplate1..n` / `modHtmlTheme` の関数契約（`BuildDocument` / `HeadHtml` / `BodyShellHtml` / `SectionsJs` / `RuntimeJs` / `ThemeCss` / `ThemeNames` 等）は18章§4.4・§5.2が正**（本章は宣言を持たない。追加・分割の規約も18章に従う）
- **`modValidate` の CheckS2C / CheckS3C**、**`modSchemas` の SchemaS2C / SchemaS3C** は入念モード用の追加分（15章§4.5～4.6・§7の表）
- 呼出前の走査: 外部へ送るテキスト（CallStep / CallChat の systemPrompt・userPrompt、企業ドシエファイルの書出、HTMLレポート出力）は送信・保存の直前に `modPii` を通す（16章 E-05／E-31。走査結果は run_log と dossier_meta に記録）
- **名前付きレンジ・図形ボタン・入力列の新設（v2.5・裁定書9 §1。v2.5.1・裁定書10で改訂）**: 本章§6は公開関数だけでなく**名前の唯一の正**でもある。新設を許可したのは次の表のみであり、実体の定義（配置・列順・書式）は各章が持つ。

  | # | 名前 | 種別 | 定義の正 | 用途 |
  |---|---|---|---|---|
  | N4 | `s1_case_id` / `s2_case_id` / `s3_case_id` / `s4_case_id` | 名前付きレンジ（単点セル・読取専用） | 13章§2.12 | S1～S4の案件ID表示。`DrawStep` が書き、`SaveEditedStep` が突合する（B1） |
  | N5 | `judge_to` | 受信箱シートの入力列 | 13章§2.6・19章§3 | 判定の入力口。日本語ラベル（採択／条件付き保留／却下）で選び、ui層が `EnumEn` で機械値へ変換して `SetInboxJudgement` の `status` 引数へ渡す（B2・裁定書10 M7） |
  | N6 | **廃止**（旧 `ib_body_draft`） | - | 13章§2.6 | v2.5.1・裁定書10 C1で**「受信箱の投函下書き行」方式へ差し替え**。名前付きレンジは作らない（受信箱は columns のみのフラットシートで、台帳の現機構では header_fields の名前付きレンジを作れない）。投函本文は受信箱テーブルの先頭データ行（id列マーカー「(下書き)」）から読む |
  | N7 | `btn_round_freeze` | HOMEの図形ボタン（caption「第2ラウンド開始」） | 11章§2 | `modCaseStore.FreezeRound` の起動口（A-2） |
  | N9 | `modPipeline2.ResetDeepOutcome` | 公開関数（宣言は本節の modPipeline2） | 本章§6 | deep outcome の明示リセット口。ui層が実行開始前に1回呼ぶ（裁定書10 M1） |
  | - | 業種ドロップダウンの隠しレンジ | 名前付きレンジ（既存作法の内部レンジ） | 13章§2.11 | `ci_industry_code` / `ci_industry_name` の入力規則の参照元（`RestoreDataKeyHiddenRange` と同作法であり、公開名を新設しない） |
  | - | `modUICase.RebindFlatValidation` | 公開関数（ui層内部ヘルパ） | 本章§6 | フラット表（f種別）の入力規則を1シートぶん張り直す口。受信箱の投函下書き行を行挿入で用意したときに ui層から呼ぶ（v2.5.2・裁定書11 Q3(a)。13章§2.6） |
  | - | `modUICase4.CopyResearchRow` | 公開関数（図形ボタンの OnAction） | 本章§6 | 追加収集の[コピー]。30,000字契約により `modUICase3` から移設（v2.5.2・裁定書11 Q1。移設前の名は `modUICase3.CopyResearchRow`。呼出は図形の OnAction 文字列のみ） |
  | - | `modTestsExcel2.RunExcelTests2` | 公開関数（test層の分割先の入口） | 本章§6 | 30,000字契約による `modTestsExcel` の分割先。呼んでよいのは `RunAllExcelTests` のみ（v2.5.2・裁定書11 Q9/Q1） |
  | - | `modUICase4.ClearCaseInput` | 公開関数（ui層内部ヘルパ） | 本章§6・13章§2.11 | 案件入力の全クリア（属性欄11・貼付欄17・実行後表示欄）。`modUIHome.HomeNewCase` が `ci_case_id` へ `(新規)` を書く**前**に呼ぶ。呼んでよいのは modUIHome のみ。30,000字契約により `modUICase3` に置けないため `CopyResearchRow` と同じ移設先へ置く（v2.5.4・裁定書13 W1） |
  | - | `modUICase3.U3_NEW_MARK`（値 `(新規)`） | 公開定数（モジュール間で共有する固定マーカー。`modInboxStore.IB_DRAFT_MARK` と同作法） | 本章§6・13章§2.11 | 案件入力の**新規モード**の固定マーカー。HOMEの[＋新規案件]（`modUIHome.HomeNewCase`）が `ci_case_id` へ書き、`modUICase3.CaseSave` の3値判定がこの値のときだけ採番する（v2.5.3・裁定書12 V1）。**公開関数の新設は本波では無い** |
  | - | `modUICase5` の Public 4本（`SerializeBody` / `ColIndexes` / `ColCount` / `RoomOf`） | ui層内部ヘルパ（W4.1分割裁定） | 本章§6 | 30,000字契約による `modUICase2` の分割先。呼んでよいのは modUICase2 のみ（裁定書10 §1でM2を解消） |

  | - | `modUIGuide`（`StartTourIfFirstRun` / `RestartTour` / `OnTourNext` / `OnTourSkip` / `ClearTour` / `EnsureGuideButtons`） | 公開関数（ui層。図形の OnAction と modBoot からの結線先） | 本章§6・13章§2.18 | 初回ガイドツアー（カード3枚）と`操作ガイド`の図形ボタン。起動シーケンスからの結線は `modBoot` の**1行**（`StartTourIfFirstRun`）だけ。`EnsureGuideButtons` は `modUIHome.EnsureScreens` が他の `Ensure*Buttons` と同じ並びで呼ぶ（v2.5.5・裁定書14 裁定6） |
  | - | Shape接頭辞 `gt_` | 図形名の接頭辞（HOME上のツアーのカード・ボタン） | 本章§6 | `modUIGuide` が置く図形はすべてこの接頭辞。削除は名前を配列へ集めてから行う。既存の `btn_` / `lbl_` / `btncopy_` と衝突しない（v2.5.5・裁定書14 裁定6） |
  | - | config `guide_tour_done`（既定 `0`） | configキー | 13章§2.3 | 初回ガイドツアーを見終えたか。`1`=済。`1` 以外はすべて未完了として扱う（v2.5.5・裁定書14 裁定6） |
  | - | `gd_btn_tour` | 名前付きレンジ（`操作ガイド`②のボタンアンカー） | 13章§2.18 | [ツアーをもう一度見る]の置き場所。OnAction は `modUIGuide.RestartTour`（v2.5.5・裁定書14 裁定6） |
  | - | `modTestsRunnerUi.RunAllTestsFromBook` | 公開関数（test層。図形の OnAction。引数なし） | 本章§6・17章 T-48 | ブック内テスト実行。`操作ガイド`⑤の[テストを実行]から呼び、ps1と同一の4条件を判定して `gd_test_result` へ書く。ui層からこの1本だけを参照してよい（`tools/vba_lint.py` の R1例外表に名指しで登録。v2.5.5・裁定書14 裁定5） |
  | - | `modTestRunner.PassCount` / `FailCount` / `SkipCount` / `ExecutedCount` | 公開関数（test層。読み出し専用） | 本章§6・17章§4-1 | 集計値の読み出し口。**集計の仕方は変えない**（R4の純ロジックのまま。LibreOffice実行テストへの影響なし）。`modTestRunner` は closed な公開契約なので本4本を `vba_lint.py` の required にも同期する（v2.5.5・裁定書14 裁定5） |
  | - | `modUISheet.EnsureButtonEx(ws, shapeKey, caption, anchorRow, anchorCol, widthPt, onActionName, kind)` / `modUISheet.CellLeft(ws, rowNo, colNo)` | 公開関数（ui層内部ヘルパ） | 本章§6・13章§2.10 | 種別つき図形ボタン（`kind` = `primary` / `plain` / `danger`）と、列アンカーの実測左端の読み口。`EnsureButton` は `plain` の薄い包みになり**呼出側のシグネチャは不変**。ボタン高は26ptで、アンカー行の行高をボタンが収まる高さまで広げてから置く（縦の重なりを構造的に潰す）。`CellLeft` は HOMEの横並びの幾何計算（直前のボタンの右端＋8pt より右の列だけをアンカーにする）に使う（v2.5.5・裁定書14 裁定7＋追補1） |
  | - | HOMEの主要動線4本のキャプション（[① 案件を作る] / [② 一括実行] / [③ レポートを出す] / [④ ヒアリングシート]） | 図形ボタンのキャプション | 13章§2.10 | 番号つき動線への再レイアウト。図形名（`btn_hm_newcase` / `btn_hm_runall` / `btn_hm_html` / `btn_hm_hearing`）と OnAction は不変で、キャプションだけを改めた。`操作ガイド`③の早見表（`build/build_rpn.py` の `GUIDE_HOME_BUTTONS`）と逐語一致させる（v2.5.5・裁定書14 追補1） |

  | - | **`modUIToast`**（`ShowToast(messageText, [kind])` / `ShowNext(stepNo)` / `HideToast()` / `CancelToast()` / `WarnLine(messageText, kind) As String` / `ShowResearchPrompts()`） | 公開関数（ui層。`modUIHome` からの結線先と `Application.OnTime` のコールバック） | 本章§6・11章§4 | **トースト**（アクティブシート右上の角丸カード。Shape接頭辞 `ts_`・Yu Gothic UI 11pt・`kind` = `info`（白地／緑枠）/ `warn`（黄地）/ `error`（赤地・白字）。いずれもコントラスト比4.5:1以上）。`Application.OnTime Now+6秒` で `HideToast` を予約し、予約は常に1本（`ShowToast` が張り替え前に `CancelToast` を呼ぶ）。**`HideToast` は対象図形が無ければ何もしない**（ブックを閉じたあとに残った予約が発火しても無害）。`CancelToast` はブックの終了処理からの取り消し口（現在の `ThisWorkbook` は終了イベントを持たないため未結線）。`ShowNext` は主要4ボタンの成功経路に出す「次の一手」1行の唯一の値源。`WarnLine` は `hm_warning` へ書く1行を組み立て、`kind="error"` のときだけ末尾に「（err_logタブの最後の行を開発担当へ送ってください）」を足す。`ShowResearchPrompts` は HOMEの[① 調べる指示文を出す]の OnAction（図形ボタンなので先頭で `modUIProgress.TryEnterUiLock` を通す）で、`操作ガイド`を開いて名前付きレンジ `gd_ch7_head` へ `Application.Goto` し、次の一手のトーストを出す（**見出し文字列は検索しない**）。**文言はすべて本モジュールが持つ**（`modUIHome` に残量が無いため。v2.5.6・裁定書17 H2/H4＋司令塔追補） |
  | - | `gd_ch7_head` | 名前付きレンジ（`操作ガイド`⑦の章見出し行） | 13章§2.18 | HOMEの[① 調べる指示文を出す]の飛び先。ビルド（`_make_guide`）が⑦章の見出し行へ付ける（v2.5.6・司令塔追補） |
  | - | HOMEの主要動線**5本**のキャプションと図形名（[① 調べる指示文を出す]=`btn_hm_step1` / [② 案件を作って貼る]=`btn_hm_step2` / [③ まとめて作る]=`btn_hm_step3` / [④ レポートを出す]=`btn_hm_step4` / [⑤ ヒアリングシートを出す]=`btn_hm_step5`） | 図形ボタンのキャプション・図形名 | 13章§2.10 | W6のボタン名を先取りし、利用者に見えるボタン名を2度変えない（v2.5.6・司令塔追補）。**OnAction は①以外すべて既存ハンドラのまま**（`HomeNewCase` / `HomeRunAll` / `HomeExportHtml` / `HomeBuildHearing`）。`操作ガイド`③の早見表（`GUIDE_HOME_BUTTONS`）と逐語一致させる |
  | - | Shape接頭辞 `ts_` | 図形名の接頭辞（トーストのカード） | 本章§6 | `modUIToast` が置く図形はすべてこの接頭辞。削除は `modUISheet.DropShapesByPrefix` に委ねる。既存の `btn_` / `lbl_` / `btncopy_` / `gt_` と衝突しない（v2.5.6・裁定書17 H2） |
  | - | `modBoot.KbAutoNote() As String` | 公開関数（ui層。読み出し専用） | 本章§6・13章§2.3 | 起動時の**ナレッジブック自動発見**（config `kb_path` が空／プレースホルダ `\\...\`／`Dir$` で不在のとき、`ThisWorkbook.Path & "\ナレッジブック.xlsx"` を探して `kb_path` へ書く）で実際に書いたときだけ「同じフォルダのナレッジブックを読み込みました。」を返す。**探索そのものは `modBoot` の Private 1本**（`ResolveKbPath`。起動処理を modBoot 以外へ散らさない）で、読むのは `modUIHome.KbStatusText` だけ（v2.5.6・裁定書17 H1） |
  | - | `modUISheet.EnsureButtonEx` の `kind="primary"` の高さ | 図形ボタンの寸法規約 | 13章§2.10 | 主要動線だけボタン高を**30pt**にする（他は従来どおり26pt）。高さを決める場所は `modUISheet.BtnHeightOf` の1本で、`AddShape` と行高の確保が同じ値を読む。**公開シグネチャは不変**（v2.5.6・裁定書17 H3(a)） |

  これ以外の名前（公開関数・名前付きレンジ・シート・列）を実装側で新設しない。必要が生じたら司令塔の裁定を経て本章§6へ先に登録する。

## 7. スキーマ・レジストリ（modSchemas。本文は15章）

**Const は使わない。すべて引数なしの純関数で返す**（`Public Function SchemaS1() As String` を `s = s & "..." & vbLf` で組み立てる）。理由: VBAの Const は1論理行1023字・行継続 `_` 25本までで、3,700字級のスキーマ本体を1宣言に収められず、1行追加した瞬間に壊れるため。旧表記の `SCHEMA_S1` 等は使わず関数名に統一する（19章§4のスキーマ・レジストリもこの関数名で読む）。

| 関数名 | 対応step | strict検証済み観点 |
|---|---|---|
| `SchemaS1()` | s1 | current_coverage は常に必須（newは空配列） |
| `SchemaS2()` | s2 | gaps は常に必須（newは空配列）。リスクユニバース10分類/頻度/影響/1～5スコア/移転可能性/status/出所enum（v2.3）。**emerging_risks（ニューリスク0～3件・空配列可・category/horizon/出所enum）を含む（v2.4）** |
| `SchemaS3()` | s3 | proposal_kind enum。scheme_id は "" 許容 |
| `SchemaS4()` | s4 | slides配列・hearing_questions |
| `SchemaS2C()` | s2c | 入念モードの批判JSON。issue_type 6値のenum。issues は0件（指摘なし）を許容 |
| `SchemaS3C()` | s3c | 入念モードの批判JSON。executive_reactions はちょうど3件。issue_type 6値のenum |
| `SchemaPF()` | pf | 5問判定・文法4値・予測類型・組み替え案 |
| `SchemaWT()` | wt | 分類enum・pattern_id |
| `SchemaFG()` | fg | 格付enum・文法4bool |

- s2r / s3r（入念モードの改訂）はスキーマを新設せず `SchemaS2()` / `SchemaS3()` を再利用する（§1 Stepレジストリ）
- 1モジュール30,000字契約に収まらない場合は modSchemas を `modSchemas1..n` へ分割してよい（関数名は変えない）
