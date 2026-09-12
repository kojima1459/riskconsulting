# CLAUDE.md - リスク提案ナビ(AIリスクコンサルティングPJ)の開発体制

読む順番と設計の正は `README.md` と `docs/spec/10〜19章`。本書は**誰が何をやるか**(体制)だけを書く。

## 1. 体制(司令塔と班)

- **Fable = 司令塔**。俯瞰・裁定・タスク分解・検収・検問・コミット・push・ユーザーへの説明。実装は原則しない(docs の1行修正や検問ツールの小修正は司令塔がやることもある)。
- **実装は班(サブエージェント)**。班は**勝手な修正・よしなに判断を禁止**: 裁定書の外にある判断は concerns として報告し、司令塔が裁定する。班は `git commit/push/checkout/restore/stash/merge` を**行わない**(司令塔のみ)。並行班は worktree で隔離し、司令塔がマージする。
- **難易度で振り分ける**(`.claude/skills/delegate/SKILL.md`): 難易度 **1〜4 = Gemini 3.1 Pro**(`python3 tools/gemini_worker.py prompt.txt out.txt`。既定。軽いタスクは全部。速さ優先の雑務だけ `--model gemini-3.8-flash`)／**5〜7 = Sonnet**／**8〜9 = Opus**／**10 と司令塔・PO の役割 = Fable**。迷ったら1段上へ。Gemini はリポジトリを見られないので、プロンプトに仕様書と参考ファイルを全部貼る。成果は司令塔が全文を検収し、**数値・行番号・HEAD は測り直す**。**最終検問と手直しは司令塔自身が行う**(委任しない)。鍵は環境変数 `GEMINI_API_KEY`(クラウドは claude.ai の Environments、手元PCは `~/.gemini_key`)。ファイルやチャットで渡さない。無ければ Sonnet で代替する。
- テストは**出来レース禁止**: 敵対的検証・変異注入を必須とし、司令塔も班とは別の箇所へ抜き打ち変異を1点以上入れる。
- こまめにコミット(WIP保全 → 検問 → クローズ)。push 先は指定ブランチのみ。

## 2. 検問と規約(要点。正は docs/spec/12章・17章)

- 全ゲート一括: `python3 tools/gate.py`(全ゲート一括・約4〜5分。一覧は gate.py の GATES が正。`--only` で絞れる)。
- ソース(.bas)は **UTF-8**(ビルドが CP932 化する)。1モジュール 30,000字・1物理行 1,000 バイト以下。
- 社内環境の確定事実(ターミナル/PowerShell/Python 不可、信頼できる場所は D: のみ、AV が重く見る書き方、配布は SharePoint/Teams の zip)は `docs/24` と `docs/29` を正とする。
