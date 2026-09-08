Attribute VB_Name = "modMockLlm2"
Option Explicit

' ==============================================================================
' modMockLlm2 - modMockLlm の分割モジュール(1モジュール30,000字契約対応)
' ------------------------------------------------------------------------------
' 役割:
'   modMockLlm 単体では30,000字契約(12章§2)を超えるため、定型項目が多い
'   S3以降の応答(S3・S4・PF・S2C・S3C)とcount_violation派生をこちらへ
'   切り出した。判定ロジックは一切持たず、単純な文字列返却関数の集まりに
'   徹する(連結はmodMockLlm.NormalBody / MockResponseが行う)。
'   本モジュール単独では呼ばれない(modGatewayRPNの公開契約はmodMockLlmの
'   MockResponseのみ。14章§4)。
'
' 正常応答(15章§8.1): MK-S3 / MK-S4 / MK-PF / MK-S2C-HIT / MK-S2C-CLEAN /
'   MK-S3C-HIT / MK-S3C-CLEAN の7本。いずれも case_type=new / renewal の
'   両文脈で合格する「共通」応答(受入条件1)。
' 障害注入(15章§8.2): BuildS3CountViolationJson はcount_violation
'   (stepName=s3の最初の1呼出のみ)専用の派生で、storiesを2件に減らして
'   V-S3-01を発火させる。
'
' 移植元: 新規(PoCに対応物なし。12章§2)。
' ==============================================================================

Public Function BuildS3Json() As String
    Dim s As String

    s = s & "{""stories"":[{""story_no"":1,""proposal_kind"":""upsell"",""headline"":""PL保険のリコール費用特約拡充"",""hook_question"":""もし来月商品回収が起きたら、今の契約で費用はどこまでカバーできますか"",""target_risk_nos"":[1],""target_gap_nos"":[],""menu_ids"":[""M-0012""],""line_ids"":[""L-04""],""scheme_id"":"""",""pitch"":""アレルゲン表示誤りによる自主回収リスクに対し、まず食品工場リスク診断サービスで表示チェック体制の弱点を可視化したうえで、生産物賠償責任保険にリコール費用特約を付帯し費用面の備えを固める"",""similar_case_id"":""K-0003"",""expected_objection"":""今の保険で十分だと思っている"",""objection_response"":""リコール費用特約の有無を一緒に確認しましょう""" & vbLf
    s = s & "},{""story_no"":2,""proposal_kind"":""cross_sell""" & vbLf
    s = s & ",""headline"":""サイバー保険の新規提案"",""hook_question"":""ECサイトの個人情報が漏れたら、お客様にどう説明されますか"",""target_risk_nos"":[5],""target_gap_nos"":[],""menu_ids"":[],""line_ids"":[],""scheme_id"":"""",""pitch"":""刷新した自社ECサイトの個人情報漏えいリスクに対し、サイバー保険で対応費用と賠償を担保する提案を行う"",""similar_case_id"":"""",""expected_objection"":""うちは大手ではないので狙われないと思っている"",""objection_response"":""中小企業を狙う攻撃が増えている実例を共有します""" & vbLf
    s = s & "},{""story_no"":3,""proposal_kind"":""scheme"",""headline"":""検知×補償バンドル型の水濡れ対策"",""hook_question"":""浸水で工場が止まったら、まず何を守りたいですか"",""target_risk_nos"":[3],""target_gap_nos""" & vbLf
    s = s & ":[],""menu_ids"":[""M-0012""],""line_ids"":[""L-04""],""scheme_id"":""S-0004"",""pitch"":""浸水による本社工場の操業停止リスクに対し、見守りヤモリ型の座組で検知センサーと水濡れ費用保険をセットにし、検知で防げなかった残余損害を保険で引き受ける"",""similar_case_id"":""K-0003"",""expected_objection"":""止水板だけで十分ではないか"",""objection_response"":""検知と保険をセットにすることで見えない損害まで備えられます""}],""unmatched_risks"":[{""risk_no"":6,""risk_name"":""国内菓子市場の縮小による需要減"",""why_unmatched"":""当社の実在メニュー・型では需要減そのものへの対応が難しいため""}],""do_not_propose"":[{""topic"":""役員賠償責任保険(D&O)"",""reason"":""係争中の紛争は把握しておらず経営体制の変化の兆候も無いため、今回は優先度が低く見送る""" & vbLf
    s = s & "}],""growth_ideas"":[" & vbLf
    s = s & "{""title"":""ギフト配送の遅延補償つき定期便"",""what"":""贈答品の配送遅延や破損に備える補償を、定期便のオプションとして付ける"",""why"":""ギフト需要が主力でEC直販比率を伸ばしており、配送品質が購入の決め手になっているため"",""insurance_fit"":""運送保険と費用保険を組み合わせ、当社が引受条件と支払基準の設計を担う"",""effect"":5,""difficulty"":""mid""" & vbLf
    s = s & "}," & vbLf
    s = s & "{""title"":""取引先を束ねる共同の食品事故補償"",""what"":""卸先や土産物店を束ねた共同購入型の食品事故補償の枠組みを作る"",""why"":""販売網が広く、取引先で起きた事故がブランド全体の評判に跳ね返る構造にあるため"",""insurance_fit"":""生産物賠償責任保険を団体扱いにし、当社が幹事として枠組みを運営する"",""effect"":4,""difficulty"":""low""" & vbLf
    s = s & "}," & vbLf
    s = s & "{""title"":""工場見学とイベントの中止補償"",""what"":""うなぎパイファクトリーの見学・催事の中止や延期に備える補償を企画する"",""why"":""体験型施設の集客が売上と評判を支えており、天候や感染症での中止が直接損失になるため"",""insurance_fit"":""興行中止保険の考え方を小口化し、当社が引受と募集の枠組みを設計する"",""effect"":4,""difficulty"":""mid""" & vbLf
    s = s & "}," & vbLf
    s = s & "{""title"":""原材料の価格変動に備える仕組み"",""what"":""小麦・卵など主要原材料の価格上昇に備える金融ヘッジと保険の組み合わせを検討する"",""why"":""相場変動が原価を直撃し、値上げの判断が経営の重荷になっているため"",""insurance_fit"":""保険単独の引受は難しく、まず分析メニューで感応度を可視化する段階から入る"",""effect"":3,""difficulty"":""high""}" & vbLf
    s = s & "],""talk_script"":{""opening"":""御社が掲げるEC直販の拡大について、いま一番気になっている足元のリスクからお伺いできますか。""," & vbLf
    s = s & """flow"":[""まず守る対象を、工場の設備だけでなくEC直販の売上とお客様からの信頼まで広げて捉え直します。"",""次に浜松本社工場が浸水で止まったとき、どこまで操業と出荷が止まるかを一緒に可視化します。"",""そのうえで自社で持つ範囲と保険へ移す範囲を、限度額と免責の線引きで整理します。"",""最後に、保険を守りだけでなくEC拡大を後押しする道具として使う形をご提案します。""]," & vbLf
    s = s & """closing"":""まずは工場の水災対策とEC物流の代替手段について、現場を拝見させてください。"",""taboo"":[""先代からの工場設備の更新を急がせる言い方""]}}"

    BuildS3Json = s
End Function

Public Function BuildS3CountViolationJson() As String
    Dim s As String

    s = s & "{""stories"":[{""story_no"":1,""proposal_kind"":""upsell"",""headline"":""PL保険のリコール費用特約拡充"",""hook_question"":""もし来月商品回収が起きたら、今の契約で費用はどこまでカバーできますか"",""target_risk_nos"":[1],""target_gap_nos"":[],""menu_ids"":[""M-0012""],""line_ids"":[""L-04""],""scheme_id"":"""",""pitch"":""アレルゲン表示誤りによる自主回収リスクに対し、まず食品工場リスク診断サービスで表示チェック体制の弱点を可視化したうえで、生産物賠償責任保険にリコール費用特約を付帯し費用面の備えを固める"",""similar_case_id"":""K-0003"",""expected_objection"":""今の保険で十分だと思っている"",""objection_response"":""リコール費用特約の有無を一緒に確認しましょう""" & vbLf
    s = s & "},{""story_no"":2,""proposal_kind"":""cross_sell""" & vbLf
    s = s & ",""headline"":""サイバー保険の新規提案"",""hook_question"":""ECサイトの個人情報が漏れたら、お客様にどう説明されますか"",""target_risk_nos"":[5],""target_gap_nos"":[],""menu_ids"":[],""line_ids"":[],""scheme_id"":"""",""pitch"":""刷新した自社ECサイトの個人情報漏えいリスクに対し、サイバー保険で対応費用と賠償を担保する提案を行う"",""similar_case_id"":"""",""expected_objection"":""うちは大手ではないので狙われないと思っている"",""objection_response"":""中小企業を狙う攻撃が増えている実例を共有します""}],""unmatched_risks"":[{""risk_no"":6,""risk_name"":""国内菓子市場の縮小による需要減"",""why_unmatched"":""当社の実在メニュー・型では需要減そのものへの対応が難しいため""}],""do_not_propose"":[{""topic"":""役員賠償責任保険(D&O)""" & vbLf
    s = s & ",""reason"":""係争中の紛争は把握しておらず経営体制の変化の兆候も無いため、今回は優先度が低く見送る""}]}"

    BuildS3CountViolationJson = s
End Function

Public Function BuildS4Json() As String
    Dim s As String

    s = s & "{""file_title"":""株式会社浜松スイーツファクトリー様 リスクマネジメントのご提案(骨子)"",""slides"":[{""slide_no"":1,""title"":""貴社の事業環境の理解"",""bullets"":[""浜松2工場体制で焼き菓子を一貫生産"",""EC直販比率の拡大を経営方針として推進中"",""国内菓子市場は縮小基調もギフト需要とEC市場は堅調""],""notes"":""まず御社を調べてきたことが伝わる事実整理から入ります""" & vbLf
    s = s & "},{""slide_no"":2,""title"":""潜在リスクの全体像"",""bullets"":[""アレルゲン表示誤りによる自主回収リスク"",""浸水による工場操業停止リスク"",""ECサイト個人情報漏えいリスク"",""原料調達難による生産停滞リスク""],""notes"":""影響度と発生しやすさで整理して見せます""" & vbLf
    s = s & "},{""slide_no"":3,""title"":""同業種で見られる事故・トラブルの類型"",""bullets"":[""食品製造業でのアレルゲン混入による自主回収の類型が知られている"",""水災による工場操業停止の類型が知られている""" & vbLf
    s = s & ",""ECサイトからの個人情報漏えいの類型が知られている""],""notes"":""実在の個別事故の社名・数値は出さず一般的な類型として話します""" & vbLf
    s = s & "},{""slide_no"":4,""title"":""当社がご支援できること"",""bullets"":[""PL保険とリコール費用特約の拡充提案"",""サイバー保険の新規提案"",""検知×補償バンドル型による水濡れ対策提案""],""notes"":""提案ストーリー3本をメニュー名・型を使って紹介します""" & vbLf
    s = s & "},{""slide_no"":5,""title"":""次のステップ"",""bullets"":[""詳細診断のご提案"",""伺いたい事項の予告""],""notes"":""次回訪問でヒアリングしたい項目を予告します""}],""hearing_questions"":[{""question"":""浜松本社工場の建物構造(耐火・耐震等級)について教えていただけますか"",""purpose"":""施設・自然災害リスクの評価のため""" & vbLf
    s = s & "},{""question"":""EC物流の委託先は何社に分散していますか"",""purpose"":""サプライチェーンリスクの評価のため""" & vbLf
    s = s & "},{""question"":""直近で食品衛生に関する指摘や自主回収はありましたか"",""purpose"":""製造・品質リスクの評価のため""" & vbLf
    s = s & "},{""question"":""止水板などの水災対策は導入されていますか"",""purpose"":""施設・自然災害リスクの確認のため""" & vbLf
    s = s & "},{""question"":""ECサイトの個人情報保護体制はどの程度整備されていますか"",""purpose"":""デジタル・情報リスクの確認のため""" & vbLf
    s = s & "},{""question"":""パート従業員の労災上乗せ補償の加入状況を教えてください"",""purpose"":""人材・労務リスクの確認のため""" & vbLf
    s = s & "},{""question"":""主要原材料の仕入先は複数確保されていますか"",""purpose"":""調達・供給網リスクの確認のため""" & vbLf
    s = s & "},{""question"":""現在のPL保険にリコール費用特約は付帯していますか"",""purpose"":""付保ギャップの確認のため""}]}"

    BuildS4Json = s
End Function

Public Function BuildPfJson() As String
    Dim s As String

    s = s & "{""summary"":""新しい菓子の詰め合わせ定期便サービスについての投函。既存メニューとの重複や引受可否を診断してほしいという内容"",""principle_checks"":[{""q_no"":1,""question"":""誰が被保険者か"",""answer"":""法人(菓子製造企業)を想定しているが対象範囲の記載がまだ弱い"",""ok"":false},{""q_no"":2,""question"":""どんな偶然の事故か"",""answer"":""定期便の配送遅延という確実に起こり得る事象で偶然性が弱い"",""ok"":false},{""q_no"":3,""question"":""損害は誰にいくら発生するか"",""answer"":""実損の測り方についての記載が無い"",""ok"":false},{""q_no"":4,""question"":""それを客観的・機械可読に測るトリガーは何か"",""answer"":""配送遅延の記録データを使えば測定可能と考えられる"",""ok"":true},{""q_no"":5,""question"":""加入する人は予兆を知っているか"",""answer"":""配送遅延の予兆を加入者が把握できる設計であり逆選択の懸念がある""" & vbLf
    s = s & ",""ok"":false}],""grammar_checks"":[{""key"":""a"",""label"":""既存アセットに載る"",""ok"":true,""note"":""検知パートナー連携の枠組みは既存の型ライブラリに近い""" & vbLf
    s = s & "},{""key"":""b"",""label"":""引受判断に変換されている"",""ok"":false,""note"":""トリガーの定義がまだ曖昧で引受判断に落とし込めていない""" & vbLf
    s = s & "},{""key"":""c"",""label"":""当社の支払データで損害が語れる"",""ok"":false,""note"":""配送遅延の支払データが社内に無い""" & vbLf
    s = s & "},{""key"":""d"",""label"":""保険料を払う法人・自治体が特定できる"",""ok"":true,""note"":""菓子製造企業が契約者になる想定は明確""}],""duplicates"":[{""ref_id"":""M-0012"",""relation"":""近接"",""note"":""検知×診断サービスの発想が既存メニューと近い""}],""predicted_drop_types"":[""T4"",""T9""],""rework_suggestions""" & vbLf
    s = s & ":[{""approach"":""entry_path"",""pattern_id"":""P9"",""suggestion"":""個人向けではなく卸先の菓子問屋を契約者とする経路に組み替える""" & vbLf
    s = s & "},{""approach"":""benefit_form"",""pattern_id"":"""",""suggestion"":""金銭給付ではなく代替配送の現物給付に組み替えて偶然性を明確にする""}],""survival"":""mid"",""advice_to_poster"":""着眼点は良いので、まずトリガーを配送遅延の記録データで定義し直すところから磨きましょう""}"

    BuildPfJson = s
End Function

Public Function BuildS2CHitJson() As String
    Dim s As String

    s = s & "{""verdict_summary"":""見落としは少ないが根拠の質と移転可能性の判定に改善余地がある"",""issues"":[{""target"":""overall"",""issue_type"":""missing"",""detail"":""営業の現場メモにあるEC物流の外部委託先1社依存という手がかりがリスクとして拾われていない"",""suggestion"":""サプライチェーンのカテゴリで単一委託先依存のリスクを追加してほしい""" & vbLf
    s = s & "},{""target"":""risk_no:6"",""issue_type"":""generic"",""detail"":""国内菓子市場の縮小というリスクが業種名を変えても通用する一般論になっている"",""suggestion"":""当社固有の商品構成やEC比率に結びつけて具体化してほしい""" & vbLf
    s = s & "},{""target"":""risk_no:3"",""issue_type"":""insurability_error"",""detail"":""浸水リスクをcoverと判定しているが地震保険は別立てであり移転可能性の説明が不十分"",""suggestion"":""transferabilityの根拠に種目ごとの適用範囲を明記してほしい""" & vbLf
    s = s & "}],""additional_risks"":[{""risk_name"":""EC物流の単一委託先依存によるサービス停止"",""why"":""現場メモに外部委託先1社に依存しトラブル時の代替が無いとの記述があるため""}]}"

    BuildS2CHitJson = s
End Function

Public Function BuildS2CCleanJson() As String
    Dim s As String

    s = s & "{""verdict_summary"":""見落とし・根拠の質・整合性いずれも問題は見当たらない"",""issues"":[],""additional_risks"":[]}"

    BuildS2CCleanJson = s
End Function

Public Function BuildS3CHitJson() As String
    Dim s As String

    s = s & "{""executive_reactions"":[{""story_no"":1,""reaction"":""PL保険は入っているつもりだったが、リコール費用の話は初耳だ"",""lands"":true},{""story_no"":2,""reaction"":""サイバーは大手が狙われる話でしょう、うちには関係ないのでは"",""lands"":false},{""story_no"":3,""reaction"":""止水板は入れているから、それ以上は今は要らないかな"",""lands"":true}],""issues"":[{""target"":""story_no:2"",""issue_type"":""wont_land"",""detail"":""中小企業を狙うサイバー攻撃の実例が無いと経営者に響かない"",""suggestion"":""同規模企業の被害実例を1件添えて具体性を出す""" & vbLf
    s = s & "},{""target"":""overall"",""issue_type"":""uw_concern"",""detail"":""D&Oを見送った判断は妥当だが根拠が投函上に薄い"",""suggestion"":""do_not_proposeの理由を係争有無の確認結果とともに明記する""" & vbLf
    s = s & "}]}"

    BuildS3CHitJson = s
End Function

Public Function BuildS3CCleanJson() As String
    Dim s As String

    s = s & "{""executive_reactions"":[{""story_no"":1,""reaction"":""リコール費用の話は納得感がある"",""lands"":true},{""story_no"":2,""reaction"":""うちも他人事ではないと思えた"",""lands"":true},{""story_no"":3,""reaction"":""止水板とセットなら検討したい"",""lands"":true}],""issues"":[]}"

    BuildS3CCleanJson = s
End Function

' ==============================================================================
' 実リボンの定型失敗文(15章§8.2・裁定書24 A-1)
' ------------------------------------------------------------------------------
'   社内AIアドイン「リボンちゃん」は失敗時に空文字を返さず、log.bas parseText
'   が組み立てた定型の日本語文字列を返す。mock_fault の ribbon_429 /
'   ribbon_disconnect / ribbon_content_filter がこの3実体を返し、
'   modGatewayRPN.RibbonFailureCode の先頭一致(E0204/E0202/E0207)を通す。
'   語彙をここ1箇所に置き、gateway側は先頭語だけを持つ。
' ==============================================================================
Public Function RibbonErr429Text() As String
    RibbonErr429Text = "(error:429)Too Many Requests"
End Function

Public Function RibbonDisconnectText() As String
    RibbonDisconnectText = "接続切れ"
End Function

Public Function RibbonContentFilterText() As String
    RibbonContentFilterText = "content_filterに該当しました"
End Function
