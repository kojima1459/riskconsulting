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
- **条件分岐（「ただし」「～の場合を除く」「～に限り」）の取り違いと、否定/限定（「支払わない」「対象外」）の反転を必ず検査**
- ナレッジに無い数値・条文番号・金額・期間は「記載なし」に置換（削除でなく置換）
- 出典マーカーが本当にその主張を支えているかの照合
- 検証で除外/修正した内容を末尾に明示（無言で直さない）

→ 補償の適否を扱う本製品にとって「支払う/支払わないの反転検査」は必須観点。15章のS2C/S3Cプロンプトに採り込み済み。

### ③ structure_chunker.py＝日本語約款チャンカー（Phase 1.5+のナレッジ装填用）

`gemini-demo/build/structure_chunker.py`（257行・Python）。第X章/節/条を全角半角漢数字対応で検出、**「第X条の/に/は…」は参照であって見出しでないと判定**する誤検出回避、ただし書きを親条文と同居させる分割、表・箇条書き不分割、コンテキストヘッダ前置。約款・引受GLをナレッジ化する際の最難関部品が完成済み。Pythonなので**装填は開発機でオフライン実行**（社内PCはPython不可）。

### ④ 情シス交渉ドキュメントの型（OQ-1～5の解決を加速）

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
- **重要な現実（発注者確認 2026-08-28）**: このキーは**全社共用の借り物1本でローテーション不可**。「作り直す」ができない＝露出は恒久被害。よって推奨は「ローテ」ではなく**「これ以上広げない」**に一本化: ①**リポジトリを非公開（private）に保つ／すでに公開なら即最小化** ②**これ以降キーを配布物・コミットに含めない**（`dist/`を`.gitignore`・`git check-ignore`で実効確認） ③難読化埋め込み（OBF1）方式そのものを廃止（環境変数注入でも「配布物にキーが入る」限り同じ穴）。※既に配布・クローンされた分の露出は技術的に取り消せないため、運用（アクセス範囲の管理・監視）で受けるしかない。
- 本製品側の対応: **direct経路のキーはブック・リポジトリに入れない**（%APPDATA%外部ファイル方式・16章）を堅持。キーの値はdocs/チャットに転記しない。
- なお gemini-demo 側に実キーの混入は無し（プレースホルダのみ）。

## 4. 探索で判明した否定情報（無かったもの）


- 「マルチモーダルチャット」「Chainlit」「ディープリサーチ」への言及: **両ブランチとも0件**（社内AIツール群の棲み分け文書は存在しない。棲み分けの一次情報は本リポジトリdocs/08とOQ-12が唯一）
- gpt-5.6の文字列も0件（PoC最新はgpt-5.5。リボン四半期更新でconfig差替の設計）
- 企業リスク分析・保険提案そのもののデモデータは無し（demo_dataは「提案の根拠側」の素材）

## 5. セッション間照合の結果（2026-08-28・伝書鳩ラウンド）

**系譜の確定（発注者回答 2026-08-28）**: main → product/nexus-agent → claude/internal-notebook-lm-chatbot-B6BE7 と派生し、**B6BE7が最終生き残り（正史）**。nexus・gemini-demoは死に筋ブランチだが、本書§1～2の転用資産はそこにしか無いものが多く、ブランチ削除前に必要分を退避すること。

**⚠️ 実物照合による訂正（同日追記・tip 51cc7a1「R34クローズ」を直接検証）**:
- **B6BE7 tipは MyBookshelfそのもの**（wintest/・build_mybookshelf.py・dist/MyBookshelf(_dev).xlsm・開発憲法CLAUDE.md を全部含む。nexus系譜の直接の後継で、§1～2の主要資産の大半はB6BE7にも実在＝退避不要のものが多い。wintestは「未移植」でなく**既にある**）
- **🔴 キー露出は正史B6BE7上でも確定**: 最新tip `51cc7a1`「R34クローズ・配布物確定」の `dist/MyBookshelf.xlsm` の `sheet23.xml` に `OBF1:` blobがあり、同ブランチの `build_mybookshelf.py` 平文 `_OBF_KEY` で復号すると32字のキー形状文字列が復元される（値は転記しない）。「キー混入は誤報」判定は完全に覆った。R34で「配布物確定」した成果物に復元可能キーが載ったまま＝死に筋ブランチ限定ではない。**このキーはローテ不能の借り物1本（発注者確認）なので「作り直し」は不可**。取れる対策は「これ以上広げない」のみ＝**①リポジトリ非公開の維持／公開なら即最小化 ②今後キーを配布物・コミットに入れない ③OBF1埋め込み方式の廃止**。B6BE7に直接上げるべき案件（GqNZ6は権限・文脈なし）。
- **`claude/claude-md-setup-GqNZ6` は系譜外の孤立ブランチ**: B6BE7とのmerge-base空、tree直下は `.claude / CLAUDE.md / README.md / docs` のみで src/dist/build を持たない。CLAUDE.md整備専用で本体コードを見ていないため「dist/が無い」「開発憲法用語が無い」との同セッション回答は**そのブランチについては正しいが正史の実態ではない**。開発憲法・検問方式の照会先はGqNZ6ではなく **B6BE7 の `.claude/skills/final-gates/SKILL.md` と CLAUDE.md**（実在確認済み）。本製品17章が追随すべき正はこちら。
- **実機テスト地獄は移植でなく検問組込の問題**: wintest（実Excel自動テスト）はB6BE7最新tipにも同梱済み（tree直下 `wintest`）。34ラウンドで実機バグが続くのは、LibreOffice合格を出荷条件にし wintest実Excel PASS を final-gates に必須化していない疑い。B6BE7確認事項:「テスト3004はLibreOffice実行か実Excel(wintest)実行か。乖離するなら final-gates に実Excel PASS を必須化しているか」。本製品17章はテストを3層分離（(a)純VBA=modTestsPure・どこでも／(b)Excel固有=wintest実機のみ／(c)実機前静的検査=vba_lint.py）し、**「(b)が通るまで出荷しない」を検問に置く**。

**決着（2026-08-28・向こうが全面訂正）**: 姉妹PJセッションは当初「4問中3問は前提不成立（誤報）」と返したが、`product/nexus-agent` を調査範囲外にしたまま「全45コミットでゼロ」と断言していた自らの範囲ミスを認め、**キー露出の指摘は完全に正しいと訂正**した。向こうの実測確定値: config シート（`sheet6.xml`）にラベル `azure_embed_key`＝`OBF1:`＋hex64（32バイトの実在キー・`obfuscate_secret()` は空なら空を返す実装なので空ではない）、XOR鍵は平文2箇所（`build/build_mybookshelf.py:136` と `src/core/modUtil.bas:52`＝`_OBF_KEY="NexusAgentBuildObfuscationKey2026"`）、**対になる Azure エンドポイントも平文で同居**（テナント名・デプロイ名・キーが揃った即利用可能な組）。3セッション（本セッション・nexus読み・B6BE7監査）が独立に同一結論へ到達。**⚠️ ただしローテーションはできない**: このキーは全社共用の借用キー1本で、現場に再発行権限がない（発注者確認 2026-08-28）。当初「①ローテーション最優先」と書いたのは前提誤りで撤回する。**ローテ不能キーで取れる対応は「これ以上漏らさない」のみ・順序**: ①**リポジトリを非公開に保つ／公開中なら即最小化**（既存クローン・フォーク・キャッシュ分の露出は技術的に取り消せない＝運用で監視するしかない） ②**今後キーを配布物・コミットに入れない**（`dist/`除外・`git check-ignore`実効確認） ③**OBF1難読化埋め込み方式の廃止**（本製品は`%APPDATA%`外部ファイル方式＝配布物にキーを入れない・16章NFR-S2）。以下の照合表は当初の食い違いの記録として残す。

**訂正（2026-08-29・ハッシュ突合による実物確定）**: 本セッションでビルド再現とSHA-256突合を行った結果、**「B6BE7最新tipにも露出」は誤りだった**と判明。B6BE7の旧配布物(`51cc7a1`のdist)にあったhex64 blobは、H-17レビュー(2026-07-28)で本番キーから差し替え済みの**ダミー値**(`modTestsPure.bas`のテストベクタ・復号結果は"DUMMY-not-a-real-key-…"の形)であり、実キーではない。**実キーの露出が現存するのは `product/nexus-agent` ブランチのdist 2ファイル**（`sheet6.xml`と`sheet14.xml`の両方に`azure_embed_key`ラベル同居のhex64 blob・ダミーとはハッシュ不一致）**のみ**。ただしB6BE7系の**git履歴**にはH-17差し替え前のコミット（テストに実キー直書き）が残るため、「リポジトリ非公開の維持が唯一の防壁」という結論は不変。向こうのセッションへの残タスクは**product/nexus-agent側でもdistの追跡を停止すること**（B6BE7で実施済みの処置の横展開）。教訓: 「同種のblob」を同一物と思い込んだ——ハッシュ突合まで下りて初めて事実になる。

**追記（2026-08-29・②が実行された）**: B6BE7に新コミット2件を確認——`86934a5`「Stop tracking the built workbooks: they carry the API key」（キー入り配布物3点 `MyBookshelf.xlsm`/`MyBookshelf_dev.xlsm`/`MyBookshelf_配布.zip` の追跡停止）と `92daebf`「Ignore all of dist/ by default」（`.gitignore`で`dist/`全面除外）。**今後の配布物へのキー混入経路は閉じた**。ただし過去コミット履歴上のキーは残存したまま（履歴書換はしていない）なので、**①非公開の維持が引き続き唯一の防壁**である点は不変。

| 論点 | 当初の向こうの回答 | 実物（本書対象＝nexus系） | 決着 |
|---|---|---|---|
| build_mybookshelf.py / TODO.md / EDGE_CASES.md / wintest | 存在しない | **実在**（TODO.mdは `docs/dev/TODO.md`） | ブランチ相違。双方とも自ブランチについて正 |
| dist内の実キー | 全45コミット走査で検出0 | **確定**: `dist/MyBookshelf(_dev).xlsm` の `sheet6.xml`（config）にラベル `azure_embed_key`＝`OBF1:`＋hex64。平文XOR鍵で復号可能・空でない実キー・Azureエンドポイント平文同居（値は転記しない） | **向こうが訂正し完全一致**。範囲外走査による誤断言だった。ただしキーはローテ不能の借用1本＝「これ以上漏らさない」（非公開維持・今後コミットに入れない・OBF1廃止）が唯一の対応 |
| リボン呼出 | `Application.Run("ChatGPT", prompt)` 引数1個・Wait/MaxTokens不在 | 12引数（RIBBON_API_CONFIRMED.md） | ブランチで実装世代が違う。**本製品は12引数版を正**とする（14章） |
| 【社内限】引受GL | **向こうでも確認**。抽出本文が `dist/index/chunks.json`（785チャンク）と配布xlsmの `knowledge_base` シート（867行・約51万字）にコミット・配布済み | demo_data/にPDF原本 | **両ブランチで露出**。承認記録なし→**OQ-11の先例にはならない**（許可の事実は発注者しか知らない） |
| 開発憲法の正 | GqNZ6は「規約用語0件」 | 対コード規約＝`product/nexus-agent:docs/dev/CONTRIBUTING.md`（§2.5=30,000字・§3=3ゲート）。CLAUDE.md（GqNZ6）は対人作法のみ | **17章はCONTRIBUTING.md準拠**。人向け作法は別レイヤー |

### B6BE7監査「先人の轍」から本製品仕様に取り込んだもの

出典: 監査レポートArtifact「先人の轍」（claude.ai/code/artifact/a0401121-…）。取込済み:

- **平文プレフィクス成否判定の禁止**（`#LLM_ERROR:`偽造でフィッシング/恒久DoS）→ 14章§6を帯域外フラグ（`ByRef ok As Boolean`）方式に改訂。CallStepシグネチャ変更
- **数式インジェクション**（セル=実行環境。先頭`=+-@`・CSV含む）＋**書き込み口の一元化**（SetCellSafe強制・直書きgrepゼロをDoDに）→ 16章 E-46・NFR-S7新設
- **クライアント側データはセキュリティ境界にしない**／**LLM返却IDのホワイトリスト照合**（本製品はE-43のID実在検証で既に対応・取得側再チェックを明文化）→ NFR-S7
- **設計書と実装の既定値乖離・防御機構の移植漏れ・`git check-ignore`実測**などの点検習慣 → 17章の実装フェーズで「先人の轍」チェックリスト（同Artifact末尾の10コマンド）を各マイルストーンのDoDに準用する

### 向こうの実測で本製品に効く数字

- **1問あたり索引表178,559字×3呼出・キャッシュ無し**（B6BE7の実測）: 長文入力自体は問題ないが、**入力上限超過時に黙って切り詰められると末尾の資料が恒久的に検索対象外になる**という指摘は、本書§1のリボン尾部劣化の観測と同根。S1後検証の重複排除（14章）に加え、**長文注入するリスト類（ナレッジ行・メニュー一覧）は末尾に番兵項目を置き、応答に番兵が反映されているかで切詰め検知する**ことをS1～S3の後検証に追加検討（15章課題）
- リトライ実装ゼロ・レート制限の移植漏れ（v1→v2で消失）: 本製品は14章§3のリトライ規約＋modRateLimiter転用（§2表）で対応済みだが、**「旧版にあった防御機構の棚卸し」をリファクタ時の必須手順**として17章に準用
