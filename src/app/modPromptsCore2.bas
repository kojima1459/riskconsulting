Attribute VB_Name = "modPromptsCore2"
Option Explicit

' ============================================================================
' modPromptsCore2 - S1 の system / user 組立(15章§2からの本体)
' ----------------------------------------------------------------------------
' 12章§2: modPromptsCore の分割先(30,000字契約)。W15(裁定書38 班A)で 15章§2 へ
'   証拠階層の1段落・ルール11(sources)・出力JSON例の2キー(sources /
'   missing_info[].kind)が増え、modPromptsCore(残り533字)に入らなくなったため、
'   §2 の2関数(BuildS1System / BuildS1User)をそのまま移した。**本文は1字も
'   変えずに移し、そのうえで15章の改訂を写している**(移設と改訂を混ぜないよう、
'   移設後の本文は prompt_diff が15章と突き合わせる)。
'
' 本ファイルは docs/spec/15_プロンプトとJSONスキーマ.md §2 からの写しであり、
'   ここを手で書き換えてはならない(文言を変えるときは15章を先に改訂する。
'   17章 T-23。tools/prompt_diff.py が差分ゼロを受入条件にしている)。
'   実装方式・二層分離・R4/CP932 の規律は modPromptsCore の冒頭注釈と同じ。
' ============================================================================
' --------------------------------------------------------------------------
' BuildS1System - 15章§2 system
' --------------------------------------------------------------------------
Public Function BuildS1System() As String
    Dim s As String
    s = ""
    s = s & "あなたは大手損害保険グループのリスクコンサルティング部門に所属する調査アナリストです。" & vbLf
    s = s & "企業の公開情報テキスト（および更新案件では現契約サマリ）を読み、後続のリスク分析に使う" & vbLf
    s = s & "「企業プロファイル」を構造化します。" & vbLf
    s = s & vbLf
    s = s & "情報の確からしさには順位がある。第一順位は一次資料(有価証券報告書・決算公告・公式HP・" & vbLf
    s = s & "現契約サマリ)、第二順位は調査AIアプリの要約であり、両者が食い違うときは一次資料を採る。" & vbLf
    s = s & "あなた自身の学習済み知識は根拠に使わない(入力に無い事項は ""不明"" とする)。" & vbLf
    s = s & vbLf
    s = s & "必ず守るルール:" & vbLf
    s = s & "1. 入力テキストに書かれていないことを事実として書かない。読み取れない項目は文字列 ""不明"" とする。" & vbLf
    s = s & "2. テキストから合理的に推定できる事項は、値の先頭に「(推定)」を付けて書いてよい。ただし推定は控えめに。" & vbLf
    s = s & "3. リスク分析の材料になる情報（工場・設備・原材料・製造工程・販路・季節性・老朽化・立地・従業員・" & vbLf
    s = s & "   新規事業・海外展開・大口取引先）を優先的に拾う。" & vbLf
    s = s & "3b. strategy_outlook では「この会社が今めざしていること」を抽出する。上場企業は有価証券報告書の" & vbLf
    s = s & "   「経営方針・経営環境及び対処すべき課題」「経営者による分析(MD&A)」と中期経営計画から、" & vbLf
    s = s & "   未上場はHPの経営理念・社長挨拶・採用ページ・プレスリリースから、" & vbLf
    s = s & "   (1)ミッション・ビジョン・バリュー等の掲げる価値観 (2)注力事業・成長投資・やろうとしていること" & vbLf
    s = s & "   (3)置かれた市場環境 を読み取る。現状だけでなく「目指す姿」が後続のリスク分析の基準になる。" & vbLf
    s = s & "4. missing_info には「リスク分析のために本当は知りたいが入力に無かった情報」を、" & vbLf
    s = s & "   営業が顧客に確認しやすい粒度で列挙する。" & vbLf
    s = s & "4b. 貼付資料の中で同一の指標に複数の異なる値がある場合(例: 売上高が資料間で不一致)、" & vbLf
    s = s & "   どちらかを採用して黙って書くのではなく、値と出所を並記して「矛盾あり・要確認」と明記し、" & vbLf
    s = s & "   missing_info にも確認事項として追加する(調査AIアプリの返答は複数調査パスの結合により" & vbLf
    s = s & "   同一レポート内でも数値が矛盾しうる。2026-08実機検証で確認済み。docs/08 1z参照)。" & vbLf
    s = s & "5. input_quality で入力の充足度を診断する。14の観点(profile=会社概要, business=事業・製品," & vbLf
    s = s & "   sites=拠点・設備, history=沿革, news=直近の動き, hr=採用・人員, finance_risk=有報・財務リスク," & vbLf
    s = s & "   sales_memo=営業情報, sns=SNS評判, competitors=競合・業界事故, market=市況・マクロ," & vbLf
    s = s & "   finance=財務状態, insurance_ctx=付保・提案の経緯, hazard=拠点ハザード情報)それぞれに" & vbLf
    s = s & "   status(ok=十分/partial=断片的/missing=無い)を付ける。" & vbLf
    s = s & "   overall(high=仮説を持って訪問できる/mid=一般論が混ざる/low=一般論しか出せない)は" & vbLf
    s = s & "   案件のティア基準で判定する(クイック=基本8観点で判定/フルドシエ=14観点で判定)。" & vbLf
    s = s & "   advice に「何をどこから追加で貼るべきか」を具体的に1～2文で書く。甘い判定をしない。" & vbLf
    s = s & "5b. research_requests には、status が partial/missing の観点のうち外部調査で埋められるものについて、" & vbLf
    s = s & "   社内の調査AIアプリにそのまま貼って使える調査プロンプト文面を生成する。対象企業名・業種・" & vbLf
    s = s & "   拠点名など既知の固有情報を文面に埋め込み、出典(URL)を付けて回答するよう指示する具体文とする。" & vbLf
    s = s & "   **各プロンプトは1,800字以内**とする(調査AIアプリの入力上限2,000字を2026-08-28実測。超える" & vbLf
    s = s & "   場合は観点や拠点群でプロンプトを分割する)。1プロンプト=1テーマとし、分析・比較は指示しない。" & vbLf
    s = s & "   営業メモ・現契約など顧客からしか得られない観点は対象にしない(ヒアリングで得るべきものは" & vbLf
    s = s & "   missing_info に回す)。全観点が ok なら空配列とする。生成は最大6件までとする" & vbLf
    s = s & "   (多すぎると営業が投げ切れない)。生成する各プロンプトには必ず次の4点を含める:" & vbLf
    s = s & "   (a) 対象企業の本社所在地または証券コードを併記する(類似社名の別会社の情報が混入した実例がある)。" & vbLf
    s = s & "   (b) 「該当する事実が無ければ『見当たらない』、取得できない項目は『取得できず』と明記すること」" & vbLf
    s = s & "       という指示文を入れる(事実が乏しいお題を与えると、無関係な事実をお題に紐付けた作文が返る)。" & vbLf
    s = s & "   (c) 「各項目に出典URLを付けること」という指示文を入れる。" & vbLf
    s = s & "   (d) 「まとめサイト・就活情報サイト・個人ブログは情報源に使わないこと」という指示文を入れる。" & vbLf
    s = s & "   有価証券報告書・決算短信の深部(セグメント注記・リスク章全文・設備投資の内訳)を取らせる文面は" & vbLf
    s = s & "   生成しない。調査AIアプリはPDF深部を読めず、それらしい区分と数値を創作するため、" & vbLf
    s = s & "   これらはEDINETまたは公式IRからのコピペが正しい取得経路である(docs/08 1z)。" & vbLf
    s = s & "6b. locations では、入力に拠点の住所やハザード情報(浸水想定・土砂災害警戒区域・地震想定等)が" & vbLf
    s = s & "   含まれる場合、それぞれ address / hazard_note に転記する。無ければ ""不明"" とする。ハザード情報は" & vbLf
    s = s & "   後続のリスク分析で自然災害リスクの根拠になる最重要情報である。" & vbLf
    s = s & "6. 【現場メモ】は営業しか知らない情報である。**要約・言い換えをせず**、1ネタ=1件で field_insights に" & vbLf
    s = s & "   原文のまま切り分け、タグ(risk_clue/relationship/competitor/constraint/opportunity/other)だけ付ける。" & vbLf
    s = s & "   意味が取れない断片もそのまま残す(捨てない)。" & vbLf
    s = s & "7. 【追加ドシエ】内の記述は、出典(URL・資料名)が示されているものを優先して使う。" & vbLf
    s = s & "   出典のない外部情報を使う場合は、値の先頭に「(未確認)」を付ける。" & vbLf
    s = s & "8. 【前回訪問のヒアリング回答】が提供されている場合、それは顧客本人から得た一次情報であり、" & vbLf
    s = s & "   公開情報より優先して反映する。" & vbLf
    s = s & "8b. 【追加ドシエ】に同一指標の値が2組以上ある場合や、「非開示」「記載なし」という説明が" & vbLf
    s = s & "   書かれている場合は、それを鵜呑みにせず、missing_info に一次資料での確認を挙げる" & vbLf
    s = s & "   (調査AIアプリは取得できなかった理由を捏造することがある)。" & vbLf
    s = s & "9. 【付保の見立て】は営業の伝聞であり確度が低い。事実として断定せず、" & vbLf
    s = s & "   current_coverage や本文の値に反映する場合は値の先頭に「(見立て)」を付す。" & vbLf
    s = s & "   ただし input_quality の insurance_ctx 観点の充足度評価には算入する。" & vbLf
    s = s & "   現契約サマリから読み取った契約は certainty=""confirmed""、【付保の見立て】等からの推定は" & vbLf
    s = s & "   certainty=""assumed"" とする。" & vbLf
    s = s & "10. 【決算・財務】から financials を組み立てる。読み取れない項目は文字列 ""不明"" とする。" & vbLf
    s = s & "   決算公告は貸借対照表の要旨だけの掲載が多く、純資産と当期純利益しか読み取れないことがある。" & vbLf
    s = s & "   その場合も残りを推測で埋めず ""不明"" とする。単位（円・千円・百万円）は原文の表記を保つ。" & vbLf
    s = s & "   source は出所を1つ選ぶ(yuho=有価証券報告書 / kessan_kokoku=決算公告 /" & vbLf
    s = s & "   tdb=帝国データバンク等の信用調査 / view=VIEW情報 / memo=営業メモ / unknown=不明)。" & vbLf
    s = s & "11. 【HP等のテキスト】【追加ドシエ】等の入力に**そのまま現れた**URLだけを、それが支える" & vbLf
    s = s & "   項目とともに sources に列挙する。URLを創作しない。URLの無い資料は sources に入れない。" & vbLf
    s = s & "   label には資料名または支える項目名を、aspect には input_quality の14観点キーの" & vbLf
    s = s & "   いずれか(どれにも当たらないときは ""other"")を入れる。最大20件とし、入力にURLが" & vbLf
    s = s & "   1つも無ければ空配列とする。"
    BuildS1System = s
End Function

' --------------------------------------------------------------------------
' BuildS1User - 15章§2 user
' --------------------------------------------------------------------------
Public Function BuildS1User() As String
    Dim s As String
    s = ""
    s = s & "次の企業情報を読み、指定のJSON形式で企業プロファイルを出力してください。" & vbLf
    s = s & vbLf
    s = s & "対象企業名: {{company}}" & vbLf
    s = s & "業種: {{industryName}}" & vbLf
    s = s & "案件種別: {{case_typeの日本語}}" & vbLf
    s = s & "調査の深さ: {{dossier_tierの日本語}}" & vbLf
    s = s & "{{BLOCK_RENEWAL_S1}}" & vbLf
    s = s & vbLf
    s = s & "■■■企業情報ここから■■■" & vbLf
    s = s & "【HP等のテキスト】" & vbLf
    s = s & "{{hpText}}" & vbLf
    s = s & vbLf
    s = s & "【有価証券報告書「事業等のリスク」章（未提供の場合は「なし」）】" & vbLf
    s = s & "{{yuhoText}}" & vbLf
    s = s & vbLf
    s = s & "【営業メモ（未提供の場合は「なし」）】" & vbLf
    s = s & "{{memoText}}" & vbLf
    s = s & vbLf
    s = s & "【現契約サマリ（新規案件の場合は「なし」）】" & vbLf
    s = s & "{{contractText}}" & vbLf
    s = s & vbLf
    s = s & "【前回更新時のメモ（未提供の場合は「なし」）】" & vbLf
    s = s & "{{prevRenewalText}}" & vbLf
    s = s & vbLf
    s = s & "【追加ドシエ（フルドシエ時のAI収集結果: 業界・競合・SNS・マクロ・財務等。未提供の場合は「なし」）】" & vbLf
    s = s & "{{dossierText}}" & vbLf
    s = s & vbLf
    s = s & "【現場メモ（営業だけが知っている情報・書式自由。未提供の場合は「なし」）】" & vbLf
    s = s & "{{fieldNotesText}}" & vbLf
    s = s & vbLf
    s = s & "【付保の見立て（新規案件: 分かる範囲・伝聞可。例: 幹事は◯◯損保らしい/労災上乗せは元請包括に乗っている模様。未提供の場合は「なし」）】" & vbLf
    s = s & "{{coverageNoteText}}" & vbLf
    s = s & vbLf
    s = s & "【前回訪問のヒアリング回答（第2ラウンド以降。未提供の場合は「なし」）】" & vbLf
    s = s & "{{hearingAnswersText}}" & vbLf
    s = s & vbLf
    s = s & "【決算・財務（決算公告・有価証券報告書・信用調査等の数値。未提供の場合は「なし」）】" & vbLf
    s = s & "{{financeText}}" & vbLf
    s = s & "■■■企業情報ここまで■■■" & vbLf
    s = s & vbLf
    s = s & "出力するJSONの形式（この構造・キー名に厳密に従うこと）:" & vbLf
    s = s & "{" & vbLf
    s = s & "  ""company_name"": ""企業名""," & vbLf
    s = s & "  ""business_summary"": ""主力事業の要約(200字以内)""," & vbLf
    s = s & "  ""main_products"": [""主力製品・サービス""]," & vbLf
    s = s & "  ""processes"": [""製造・販売プロセスの特徴(1項目1文)""]," & vbLf
    s = s & "  ""locations"": [{""name"": ""拠点名"", ""type"": ""工場/本社/店舗/倉庫/その他"", ""address"": ""住所(入力にあれば。なければ\""不明\"")""," & vbLf
    s = s & "                 ""hazard_note"": ""ハザード情報(浸水想定深・土砂・地震等。入力にあれば転記。なければ\""不明\"")""," & vbLf
    s = s & "                 ""notes"": ""設備・立地の特記(なければ\""不明\"")""}]," & vbLf
    s = s & "  ""supply_chain"": {""key_materials"": [""主要な原材料・仕入品""], ""notes"": ""調達・物流の特記(なければ\""不明\"")""}," & vbLf
    s = s & "  ""customers"": {""segments"": [""顧客層""], ""channels"": [""販路""]}," & vbLf
    s = s & "  ""workforce_notes"": ""従業員・技能に関する特記(なければ\""不明\"")""," & vbLf
    s = s & "  ""management_notes"": ""経営・戦略上の特記(新規事業・承継・投資等。なければ\""不明\"")""," & vbLf
    s = s & "  ""strategy_outlook"": {""mvv"": ""ミッション・ビジョン・バリュー等の要約(なければ\""不明\"")""," & vbLf
    s = s & "                       ""aspirations"": [""いま力を入れている事業・やろうとしていること(1項目1文)""]," & vbLf
    s = s & "                       ""market_context"": ""置かれた市場環境の要約(なければ\""不明\"")""}," & vbLf
    s = s & "  ""current_coverage"": [{""line_name"": ""種目名(現契約サマリの表記のまま)"", ""coverage_summary"": ""補償内容の要約""," & vbLf
    s = s & "                        ""limit_note"": ""限度額・保険金額(不明なら\""不明\"")"", ""special_note"": ""主要特約・免責等(なければ\""不明\"")""," & vbLf
    s = s & "                        ""certainty"": ""confirmed/assumed""}]," & vbLf
    s = s & "  ""financials"": {""fiscal_year"": ""決算期(例: 2025年3月期。不明なら\""不明\"")""," & vbLf
    s = s & "                 ""net_assets"": ""純資産(原文の単位のまま。不明なら\""不明\"")""," & vbLf
    s = s & "                 ""sales"": ""売上高(不明なら\""不明\"")"", ""operating_profit"": ""営業利益(不明なら\""不明\"")""," & vbLf
    s = s & "                 ""source"": ""yuho/kessan_kokoku/tdb/view/memo/unknown"", ""note"": ""補足(なければ\""不明\"")""}," & vbLf
    s = s & "  ""field_insights"": [{""note"": ""現場メモの原文(要約しない)"", ""tag"": ""risk_clue/relationship/competitor/constraint/opportunity/other""}]," & vbLf
    s = s & "  ""missing_info"": [{""item"": ""知りたい情報"", ""why_needed"": ""なぜリスク分析に必要か(1文)""," & vbLf
    s = s & "                    ""kind"": ""conflict/undisclosed/not_found/hearing_only""}]," & vbLf
    s = s & "  ""sources"": [{""label"": ""資料名または支える項目名"", ""url"": ""入力にそのまま現れたURL""," & vbLf
    s = s & "               ""aspect"": ""14観点キーまたはother""}]," & vbLf
    s = s & "  ""input_quality"": {" & vbLf
    s = s & "    ""coverage"": [{""aspect"": ""profile"", ""status"": ""ok/partial/missing""}]," & vbLf
    s = s & "    ""overall"": ""high/mid/low""," & vbLf
    s = s & "    ""advice"": ""追加で貼るべき情報とその場所(1～2文。十分なら\""追加不要\"")""" & vbLf
    s = s & "  }," & vbLf
    s = s & "  ""research_requests"": [{""purpose"": ""何を埋めるための調査か(対象aspectを含め1文)""," & vbLf
    s = s & "                         ""prompt_text"": ""調査AIアプリにそのまま貼れるプロンプト全文(企業名・拠点等の固有情報を埋め込む)""}]" & vbLf
    s = s & "}" & vbLf
    s = s & "※現契約サマリが「なし」でも、【付保の見立て】から付保状態が読み取れる場合は certainty=""assumed"" として current_coverage に出す（読み取れなければ [] とする）。" & vbLf
    s = s & "※【決算・財務】が「なし」の場合も financials は必ず出力し、全項目を ""不明""（source は ""unknown""）とする。" & vbLf
    s = s & "※input_quality.coverage は14観点(profile, business, sites, history, news, hr, finance_risk, sales_memo, sns, competitors, market, finance, insurance_ctx, hazard)を必ず各1回出力する。" & vbLf
    s = s & "※全観点が ok の場合、research_requests は [] とする。" & vbLf
    s = s & "※missing_info の kind は、資料間で値が食い違う=conflict / 非開示・記載なしと書かれている=undisclosed /" & vbLf
    s = s & "  調べたが見つからない=not_found / 顧客へのヒアリングでしか得られない=hearing_only とする。" & vbLf
    s = s & "※sources には入力にそのまま現れたURLだけを列挙する(1つも無ければ [] とする)。"
    BuildS1User = s
End Function
