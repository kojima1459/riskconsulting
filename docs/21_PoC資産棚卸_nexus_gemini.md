# 21. PoC資産棚卸（nexus-agent / gemini-demo ブランチ探索・2026-08-28）

**経緯**: 発注者の指示でNotebookLMリポジトリ（github.com/kojima1459/notebook）の未探索2ブランチを全読した結果。従来の準用元（`claude/internal-notebook-lm-chatbot-B6BE7`）に加え、**転用可能な資産が大量に見つかった**。12章の移植対応表を補完する詳細台帳。ワークツリー: `/home/user/kojima1459/nexus`（product/nexus-agent）・`/home/user/kojima1459/gemini-demo`（claude/prototype-gemini-demo）。

## 1. 最重要発見トップ5

### ① gemini-demo/demo_data の16PDF＝シードナレッジの即戦力（T-90直結）

すべて当社の実業務文書。**抽出済みテキスト785チャンクが `dist/index/chunks.json` に入っている**ので、PDF処理なしでシード装填に使える。

| 分類 | 中身 |
|---|---|
| 約款3種 | 瑕疵保証責任保険・生産物回収費用保険（リコールプロテクション）・（賠責Q&A集付随） |
| 【社内限】引受ガイドライン2種 | 顧客サービス費用保険（債務履行型）・瑕疵保証責任保険（いずれも新種保険部 費用・信用G） |
| 研修資料4種 | 商品カレッジ基礎講座（再保険/収益改善/損保基礎①②。商品・CSV×DX企画部） |
| その他 | リコールプロテクションHB別冊・保証制度入門教材・賠責Q&A集（判例付き）・家主費用利益保険（孤独死対応）・契約者パンフ |

→ **ニューリスクG領域（瑕疵保証・リコール・費用信用）の判定基準・約款根拠がそのまま手に入る**。種目ナレッジ確認シートの回収を待たずに、この領域だけ先行でシード化できる。※【社内限】資料のAI投入・ナレッジ転用はOQ-11の枠組みで確認（公開情報でない）。

### ② VerifierSystemPrompt＝S2C/S3C批判パスの先行実装（15章に反映済み）

`gemini-demo/src/chatbot/modRagEngine.bas:157-173`。損保ドメインでチューニング済みの自己検証プロンプト。特に価値が高いルール:
- **条件分岐（「ただし」「〜の場合を除く」「〜に限り」）の取り違いと、否定/限定（「支払わない」「対象外」）の反転を必ず検査**
- ナレッジに無い数値・条文番号・金額・期間は「記載なし」に置換（削除でなく置換）
- 出典マーカーが本当にその主張を支えているかの照合
- 検証で除外/修正した内容を末尾に明示（無言で直さない）

→ 補償の適否を扱う本製品にとって「支払う/支払わないの反転検査」は必須観点。15章のS2C/S3Cプロンプトに採り込み済み。

### ③ structure_chunker.py＝日本語約款チャンカー（Phase 1.5+のナレッジ装填用）

`gemini-demo/build/structure_chunker.py`（257行・Python）。第X章/節/条を全角半角漢数字対応で検出、**「第X条の/に/は…」は参照であって見出しでないと判定**する誤検出回避、ただし書きを親条文と同居させる分割、表・箇条書き不分割、コンテキストヘッダ前置。約款・引受GLをナレッジ化する際の最難関部品が完成済み。Pythonなので**装填は開発機でオフライン実行**（社内PCはPython不可）。

### ④ 情シス交渉ドキュメントの型（OQ-1〜5の解決を加速）

| 文書 | 用途 |
|---|---|
| `gemini-demo/docs/it-checklist.md` | 情シスに持参する記入式確認表（API形態/プロキシ/SharePoint/マクロ環境/ログ・PII・法務/コスト） |
| `gemini-demo/docs/security.md` | 「何を防ぐ/防げない/結論」3段構成。キー難読化の限界を正直に書いた上で移行図を示す**セキュリティ部門との会話の型** |
| `gemini-demo/docs/deployment.md` | Trust Center・組織CA署名・ロールバック・監視閾値 |
| `gemini-demo/docs/01/02_ガイド` | 非エンジニア向け逐語手順・配布5点セットの型 |

### ⑤ nexus自己インストーラ配布機構＝配布課題の丸ごと解決

`nexus/build/build_mybookshelf.py` + `build/ovba.py` + `template_skeleton.xlsm`: VBAソースを`vba_src`シート（veryHidden）に文字列で埋め、起動時に`VBComponents.Add`で自己注入。**「メール/共有フォルダで.xlsm1個渡すだけ」の配布が実証済み**。＋`gemini-demo/build/BuildBootstrap.bas`（PowerShell不可PCでもVBAだけでビルド）と`make_xlsm.py`（Windows無しの純Pythonビルダ）。

## 2. その他の転用資産（12章対応表の補完）

| 資産 | 場所 | 転用先 |
|---|---|---|
| modPiiGuard（契約番号/証券番号/電話/メール等7パターン） | gemini-demo/src/chatbot/modPiiGuard.bas | modPiiの検知パターン拡充 |
| modRateLimiter（1h/1day滑り窓） | gemini-demo/src/chatbot/modRateLimiter.bas | 暴走防止の最終防衛線（コスト度外視でも事故防止として） |
| modUsageLogger/Aggregator（per-user CSV→管理者集計。**質問本文は既定で記録せずオプトイン**） | gemini-demo/src | 利用ログ設計＋法務説明の型 |
| RAG基盤（modSimilarity純VBA内積top-k・modIndexReader/Writerアトミック公開） | gemini-demo/src | Phase 2でナレッジがベクトル検索規模になった時の保険。**v1では使わない**（決定的フィルタ方針は不変） |
| modApiGatewayのプロバイダ抽象（Embed/Chatの2関数＋Select Case切替） | gemini-demo/src/shared | 設計参考（本製品はnexus/modGateway系が本命） |
| Shape-SPA UI・MS&ADデザインシステム・描画4不変条件 | nexus/src/ui/modSkin.bas ほか | 11章UIの実装時の流儀・配色 |
| EDGE_CASES.md（実バグ由来の罠カタログ6章） | nexus/docs/dev | 32bit Shape増殖/MAX_PATH/AVロック/サロゲートペア/予約語衝突等。**同じ罠を踏まないための必読文書** |
| wintest自律テストループ（AUTONOMOUS_LOOP_PROMPT.md） | nexus/wintest | 実装フェーズのAI自律実機検証の運用ノウハウ |
| FUTURE_IDEAS.md「確信度の自己申告を不採用にした理由」 | nexus/docs/dev/FUTURE_IDEAS.md:66-79 | 「偽の精度表示は過信事故を招く」— 本製品でも同じ議論が必ず出る。判断の先例 |
| 賠責Q&A集の判例付きQ&A形式 | demo_data/sample_01 | ナレッジの「刺さる形式」の見本（Q&A+判例引用） |

## 3. ⚠️ セキュリティ発見（発注者対応要）

- **`nexus/dist/MyBookshelf.xlsm`・`MyBookshelf_dev.xlsm` に Azure OpenAI の実キー（埋め込み用）が難読化つきでコミット済み**。難読化はXOR+16進の軽量方式で、鍵文字列が `build/build_mybookshelf.py:139` に平文であるため誰でも復元可能。TODO.mdでリスク認識済みだが未対処。
- 推奨: ①当該キーのローテーション ②distをgit管理から外す（.gitignore） ③`embed_transport`既定を設計書どおり`ribbon`へ（実装は`direct`既定で設計書と食い違い）
- 本製品側の対応: **direct経路のキーはブック・リポジトリに入れない**（%APPDATA%外部ファイル方式・16章）を堅持。キーの値はdocs/チャットに転記しない。
- なお gemini-demo 側に実キーの混入は無し（プレースホルダのみ）。

## 4. 探索で判明した否定情報（無かったもの）

- 「マルチモーダルチャット」「Chainlit」「ディープリサーチ」への言及: **両ブランチとも0件**（社内AIツール群の棲み分け文書は存在しない。棲み分けの一次情報は本リポジトリdocs/08とOQ-12が唯一）
- gpt-5.6の文字列も0件（PoC最新はgpt-5.5。リボン四半期更新でconfig差替の設計）
- 企業リスク分析・保険提案そのもののデモデータは無し（demo_dataは「提案の根拠側」の素材）
