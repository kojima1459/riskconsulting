Attribute VB_Name = "modMockLlm"
Option Explicit

' ==============================================================================
' modMockLlm - 本体内mockトランスポート(14章§4(a)・15章§8)
' ------------------------------------------------------------------------------
' 役割:
'   modGatewayRPN のmock分岐(CurrentTransport=mock)が呼ぶ唯一の相手。架空企業
'   「株式会社浜松スイーツファクトリー」(業種09・菓子製造・浜松2工場・EC直販)
'   の決定的JSONを、乱数・現在時刻を使わずに返す(同一入力=同一応答)。
'   wintest/mock_ribbon/modMockRibbon.bas(ニセリボンちゃん。本体外の.xlamスタブ。
'   12引数Application.Run配管の実機検証用)とは別物であり役割を兼ねない
'   (14章§4(a)(b)・T-14b)。
'
' 公開契約(modGatewayRPNが前提とする唯一の公開口。14章§4・T-14):
'   Public Function MockResponse(stepName, variantName, fault) As String
'     stepName    : s1/s2/s2r/s3/s3r/s4/pf/s2c/s3c/sp(壁打ち)
'     variantName : modGatewayRPN.ResolveMockVariant の戻り値
'                   (new/renewal/hit/clean/common)
'     fault       : config mock_fault の値(空文字なら15章§8.1の正常応答のみ)
'   戻り値は生応答そのもの。JSON防衛線(抽出・検証)は呼び出し側の責務であり、
'   ここでは通さない(14章§5)。
'
' 正常応答(15章§8.1。7 step種・11応答):
'   MK-S1-NEW / MK-S1-RNW / MK-S2-NEW / MK-S2-RNW / MK-S3 / MK-S4 / MK-PF /
'   MK-S2C-HIT / MK-S2C-CLEAN / MK-S3C-HIT / MK-S3C-CLEAN。各応答は自分の
'   バリアント文脈(NEWはcase_type=new、RNWはrenewal。共通応答は両文脈)で
'   modValidateに合格する(受入条件1)。sp(壁打ち)はスキーマを持たないため
'   自由文を返す。
'
' 障害注入(15章§8.2。config mock_fault。既定は空=正常応答のみ):
'   broken_json    : 全step通算で最初の1呼出だけ、末尾の閉じ括弧を欠いた
'                    不完全JSONを返す(2回目以降は正常)
'   enum_violation : stepName=s2 の最初の1呼出だけ、risk_no=1のcategoryを
'                    enum外の"quality"に差し替えて返す(2回目以降・他stepは正常)
'   count_violation: stepName=s3 の最初の1呼出だけ、storiesを2件に減らして
'                    返す(2回目以降は正常)
'   ghost_id       : stepName=s3 の呼出のたび毎回、menu_idsに実在しない
'                    "M-9999"を混ぜて返す(修復リトライしても直らずE0301で
'                    停止する経路の検査用)
'   empty          : どのstepでも呼出のたび毎回、空文字列を返す
'   limit          : どのstepでも呼出のたび毎回、利用上限を示す文字列を返す
'   fake_err       : どのstepでも呼出のたび毎回、先頭行が"#ERR:E0201:..."で
'                    続く行に正常なJSON本体を持つ文字列を返す(帯域外成否規約の
'                    検査用。ok=Trueのまま素通しされ、エラーUIへ昇格しない)
'   上記7値以外(未知の値)は正常応答へフォールバックする(config入力ミスで
'   E2E全体を暴走させないため)。
'
' 「最初の1呼出のみ」型(broken_json/enum_violation/count_violation)は本体
' ブックの起動中(=このVBAプロジェクトが生きている間)だけ有効なモジュール
' レベル変数で数える。ブックを閉じる、またはVBAプロジェクトがリセットされると
' 初回に戻る。s2r/s3r(15章§4.7 改訂パス)はNormalBodyの応答選択ではs2/s3と
' 同じ応答を返すが、上表の「最初の1呼出」判定はs2/s3の呼出だけを数える
' (表の「適用step」欄が厳密にs2/s3とだけ書いているため。s2r/s3rでの発火は
' 対象外)。
'
' モジュール分割: 1モジュール30,000字契約(12章§2)のため、定型項目が多い
' S3以降の6応答(+count_violation派生)はmodMockLlm2へ切り出した。判定ロジック
' はこのモジュールに閉じ、modMockLlm2は単純な文字列返却関数の集まりに徹する。
'
' 移植元: 新規(PoCに対応物なし。12章§2)。
' ==============================================================================

' 「最初の1呼出のみ」型障害の発火済みフラグ(乱数を使わないための状態保持)。
Private mBrokenFired As Boolean
Private mEnumS2Fired As Boolean
Private mCountS3Fired As Boolean

' ==============================================================================
' MockResponse - 唯一の公開口(14章§4・T-14契約)
' ==============================================================================
Public Function MockResponse(ByVal stepName As String, ByVal variantName As String, _
                             ByVal fault As String) As String
    Dim stepKey As String
    Dim f As String
    Dim body As String

    stepKey = LCase$(Trim$(stepName))
    f = Trim$(fault)

    If LenB(f) = 0 Then
        MockResponse = NormalBody(stepKey, variantName)
        Exit Function
    End If

    ' 「呼出のたび毎回」型のうち、正常応答を組み立てる前に決まる2値。
    If f = "empty" Then
        MockResponse = ""
        Exit Function
    ElseIf f = "limit" Then
        MockResponse = LimitFaultText()
        Exit Function
    End If

    body = NormalBody(stepKey, variantName)

    Select Case f
        Case "fake_err"
            ' 帯域外成否規約の検査用。ok=Trueのまま素通しされる想定
            ' (14章§4・§6。判定材料にしないのはmodGatewayRPN側の責務)。
            MockResponse = "#ERR:E0201:偽装エラーです" & vbLf & body
        Case "broken_json"
            If Not mBrokenFired Then
                mBrokenFired = True
                MockResponse = BreakJsonTail(body)
            Else
                MockResponse = body
            End If
        Case "enum_violation"
            If stepKey = "s2" And Not mEnumS2Fired Then
                mEnumS2Fired = True
                MockResponse = InjectEnumViolationS2(body)
            Else
                MockResponse = body
            End If
        Case "count_violation"
            If stepKey = "s3" And Not mCountS3Fired Then
                mCountS3Fired = True
                MockResponse = modMockLlm2.BuildS3CountViolationJson()
            Else
                MockResponse = body
            End If
        Case "ghost_id"
            If stepKey = "s3" Then
                MockResponse = InjectGhostIdS3(body)
            Else
                MockResponse = body
            End If
        Case Else
            ' 未知の値は正常応答へフォールバック(config入力ミスの暴走防止)。
            MockResponse = body
    End Select
End Function

' ------------------------------------------------------------------------------
' NormalBody - fault抜きの正常応答を組み立てる(15章§8.1)。stepKeyは小文字化・
'   トリム済みの前提(呼び元のMockResponseで正規化済み)。
' ------------------------------------------------------------------------------
Private Function NormalBody(ByVal stepKey As String, ByVal variantName As String) As String
    Dim v As String
    v = LCase$(Trim$(variantName))

    Select Case stepKey
        Case "s1"
            If v = "renewal" Then
                NormalBody = BuildS1RnwJson()
            Else
                NormalBody = BuildS1NewJson()
            End If
        Case "s2", "s2r"
            If v = "renewal" Then
                NormalBody = BuildS2RnwJson()
            Else
                NormalBody = BuildS2NewJson()
            End If
        Case "s3", "s3r"
            NormalBody = modMockLlm2.BuildS3Json()
        Case "s4"
            NormalBody = modMockLlm2.BuildS4Json()
        Case "pf"
            NormalBody = modMockLlm2.BuildPfJson()
        Case "s2c"
            If v = "hit" Then
                NormalBody = modMockLlm2.BuildS2CHitJson()
            Else
                NormalBody = modMockLlm2.BuildS2CCleanJson()
            End If
        Case "s3c"
            If v = "hit" Then
                NormalBody = modMockLlm2.BuildS3CHitJson()
            Else
                NormalBody = modMockLlm2.BuildS3CCleanJson()
            End If
        Case "sp"
            NormalBody = SpTextMock()
        Case Else
            ' 未定義のstepName(wt/fg等のPhase 1.5含む)は空文字を返す。
            ' 呼び元のClassifyResponseがE0202(空応答)として扱う。
            NormalBody = ""
    End Select
End Function

' ------------------------------------------------------------------------------
' 障害注入ヘルパー(15章§8.2)
' ------------------------------------------------------------------------------
Private Function LimitFaultText() As String
    LimitFaultText = "本日のAI利用回数が上限に達しました。時間をおいて再度お試しください。"
End Function

' 末尾の閉じ括弧を1文字落として不完全JSONにする(broken_json)。
' Build*Jsonの最終チャンクは末尾にvbLfを付けない契約のため、正常応答の末尾は
' 必ずJSON最外殻の"}"そのものである。
Private Function BreakJsonTail(ByVal body As String) As String
    If Right$(body, 1) = "}" Then
        BreakJsonTail = Left$(body, Len(body) - 1)
    Else
        BreakJsonTail = body
    End If
End Function

' risk_no=1のcategoryをenum外の値へ差し替える(enum_violation。V-S2-03発火用)。
' 対象部分文字列はBuildS2*Jsonの生成時にチャンク境界の保護対象としており、
' 実行時の戻り値の中でも分断されず1本のまま現れる。
Private Function InjectEnumViolationS2(ByVal body As String) As String
    InjectEnumViolationS2 = Replace(body, """category"":""manufacturing_quality""", """category"":""quality""", 1, 1)
End Function

' story_no=1のmenu_idsに実在しない"M-9999"を混ぜる(ghost_id。V-S3-03発火用)。
' 対象部分文字列はBuildS3Jsonの生成時にチャンク境界の保護対象としている。
Private Function InjectGhostIdS3(ByVal body As String) As String
    InjectGhostIdS3 = Replace(body, """menu_ids"":[""M-0012""]", """menu_ids"":[""M-0012"",""M-9999""]", 1, 1)
End Function

' 壁打ち(sp)のmock応答。スキーマを持たない自由対話のため、決定的な短い
' 助言文を返す(15章§6.5)。
Private Function SpTextMock() As String
    Dim s As String
    s = s & "まずは浸水による操業停止と原料調達の不安定化、この2つが重なっている点が気になります。" & vbLf
    s = s & "工場の止水対策の状況と、原料の仕入先分散の有無を確認できると、提案の骨太さが変わってきます。"
    SpTextMock = s
End Function

' ==============================================================================
' 正常応答本体(15章§8.1)。s = s & "..." & vbLf 方式(Constは使わない。VBAの
' 1論理行1,023字・行継続25本の制約を避けるため)。
' ==============================================================================

Public Function BuildS1NewJson() As String
    Dim s As String

    s = s & "{""company_name"":""株式会社浜松スイーツファクトリー"",""business_summary"":""静岡県浜松市に本社を置く洋菓子・和菓子の製造販売企業。自社工場2拠点で焼き菓子を中心に生産し、直営店・卸売に加えEC直販を伸ばしている。"",""main_products"":[""季節限定焼き菓子ギフトセット"",""洋菓子詰め合わせ(EC限定)"",""和菓子詰め合わせ""],""processes"":[""浜松本社工場で主力の焼き菓子を一貫生産している"",""積志第二工場は繁忙期のギフト商品増産に対応している"",""EC直販分は本社工場から直送する体制である""],""locations"":[{""name"":""浜松本社工場"",""type"":""工場"",""address"":""不明"",""hazard_note"":""浸水想定区域内(想定浸水深0.5メートルから3.0メートル)との記載あり"",""notes"":""主力ラインを持つ最大拠点""},{""name"":""積志第二工場"",""type"":""工場"",""address"":""不明"",""hazard_note"":""不明"",""notes""" & vbLf
    s = s & ":""繁忙期の増産用ライン""},{""name"":""本社直営店"",""type"":""店舗"",""address"":""不明"",""hazard_note"":""不明"",""notes"":""不明""}],""supply_chain"":{""key_materials"":[""小麦粉"",""バター"",""卵"",""国産果実(みかん・いちご)""],""notes"":""主要原材料は複数の国内商社経由で調達しており、乳製品・果実の一部は季節や産地不作の影響を受けやすいと記載がある""},""customers"":{""segments"":[""個人ギフト需要"",""法人向け贈答需要"",""EC直販の個人顧客""],""channels"":[""直営店"",""卸売(菓子問屋)"",""自社ECサイト""]},""workforce_notes"":""製造ラインはパート従業員の比率が高く、繁忙期は季節雇用で補っていると採用ページに記載がある"",""management_notes"":""EC直販比率の拡大を経営方針として掲げ、直近で自社ECサイトを刷新したと社長挨拶で述べている"",""strategy_outlook"":{""mvv""" & vbLf
    s = s & ":""地域に根差した菓子づくりで顧客の特別な日に寄り添うことを掲げている"",""aspirations"":[""EC直販比率のさらなる拡大"",""季節限定商品の開発強化"",""衛生管理体制の高度化""],""market_context"":""国内の菓子市場は縮小傾向だがギフト需要とEC市場は堅調に推移していると業界記事にある""},""current_coverage"":[],""field_insights"":[{""note"":""社長は先代からの工場を大事にしており設備更新には慎重だと聞いている"",""tag"":""constraint""},{""note"":""EC直販の物流は外部委託先1社に依存しておりトラブル時の代替が無いらしい"",""tag"":""risk_clue""},{""note"":""競合の同業他社が値上げに踏み切ったため価格面では当社が優位に見えるとのこと"",""tag"":""competitor""}],""missing_info"":[{""item"":""浜松本社工場の建物構造(耐火・耐震等級)"",""why_needed"":""施設・自然災害リスクの評価に必要なため""}" & vbLf
    s = s & ",{""item"":""EC物流委託先との契約内容(損害時の責任分担)"",""why_needed"":""サプライチェーンリスクの評価に必要なため""},{""item"":""直近の食品衛生関連の指摘・自主回収の有無"",""why_needed"":""製造・品質リスクの評価に必要なため""}],""input_quality"":{""coverage"":[{""aspect"":""profile"",""status"":""ok""},{""aspect"":""business"",""status"":""ok""},{""aspect"":""sites"",""status"":""partial""},{""aspect"":""history"",""status"":""ok""},{""aspect"":""news"",""status"":""partial""},{""aspect"":""hr"",""status"":""partial""},{""aspect"":""finance_risk"",""status"":""missing""},{""aspect"":""sales_memo"",""status"":""ok""},{""aspect"":""sns""" & vbLf
    s = s & ",""status"":""missing""},{""aspect"":""competitors"",""status"":""missing""},{""aspect"":""market"",""status"":""partial""},{""aspect"":""finance"",""status"":""missing""},{""aspect"":""insurance_ctx"",""status"":""ok""},{""aspect"":""hazard"",""status"":""partial""}],""overall"":""mid"",""advice"":""有価証券報告書相当の財務・リスク情報とSNS評判、競合動向、市況情報を追加すると仮説の精度が上がる""},""research_requests"":[{""purpose"":""finance_risk観点(財務・事業リスクの記載)を埋めるための調査"",""prompt_text"":""静岡県浜松市の菓子メーカーである株式会社浜松スイーツファクトリー(本社所在地:静岡県浜松市)について、EDINETまたは同社の公式IRページに掲載されている有価証券報告書または決算公告の「事業等のリスク」に相当する記載内容を調査してください。該当する事実が見当たらない場合は「見当たらない」、取得できない項目は「取得できず」と明記してください。各項目には出典URLを付けてください。まとめサイト・就活情報サイト・個人ブログは情報源に使わないでください。有価証券報告書や決算公告が存在しない場合はその旨を報告してください。""" & vbLf
    s = s & "},{""purpose"":""sns観点(SNS・口コミの評判傾向)を埋めるための調査"",""prompt_text"":""静岡県浜松市の菓子メーカーである株式会社浜松スイーツファクトリー(本社所在地:静岡県浜松市)について、SNSや口コミサイトでの評判傾向(品質・接客・労働環境・炎上の有無)を調査してください。該当する事実が見当たらない場合は「見当たらない」、取得できない項目は「取得できず」と明記してください。各項目には出典URLを付けてください。まとめサイト・就活情報サイト・個人ブログは情報源に使わないでください。""}]}"

    BuildS1NewJson = s
End Function

Public Function BuildS1RnwJson() As String
    Dim s As String

    s = s & "{""company_name"":""株式会社浜松スイーツファクトリー"",""business_summary"":""静岡県浜松市に本社を置く洋菓子・和菓子の製造販売企業。自社工場2拠点で焼き菓子を中心に生産し、直営店・卸売に加えEC直販を伸ばしている。"",""main_products"":[""季節限定焼き菓子ギフトセット"",""洋菓子詰め合わせ(EC限定)"",""和菓子詰め合わせ""],""processes"":[""浜松本社工場で主力の焼き菓子を一貫生産している"",""積志第二工場は繁忙期のギフト商品増産に対応している"",""EC直販分は本社工場から直送する体制である""],""locations"":[{""name"":""浜松本社工場"",""type"":""工場"",""address"":""不明"",""hazard_note"":""浸水想定区域内(想定浸水深0.5メートルから3.0メートル)との記載あり"",""notes"":""主力ラインを持つ最大拠点""},{""name"":""積志第二工場"",""type"":""工場"",""address"":""不明"",""hazard_note"":""不明"",""notes""" & vbLf
    s = s & ":""繁忙期の増産用ライン""},{""name"":""本社直営店"",""type"":""店舗"",""address"":""不明"",""hazard_note"":""不明"",""notes"":""不明""}],""supply_chain"":{""key_materials"":[""小麦粉"",""バター"",""卵"",""国産果実(みかん・いちご)""],""notes"":""主要原材料は複数の国内商社経由で調達しており、乳製品・果実の一部は季節や産地不作の影響を受けやすいと記載がある""},""customers"":{""segments"":[""個人ギフト需要"",""法人向け贈答需要"",""EC直販の個人顧客""],""channels"":[""直営店"",""卸売(菓子問屋)"",""自社ECサイト""]},""workforce_notes"":""製造ラインはパート従業員の比率が高く、繁忙期は季節雇用で補っていると採用ページに記載がある"",""management_notes"":""EC直販比率の拡大を経営方針として掲げ、直近で自社ECサイトを刷新したと社長挨拶で述べている"",""strategy_outlook"":{""mvv""" & vbLf
    s = s & ":""地域に根差した菓子づくりで顧客の特別な日に寄り添うことを掲げている"",""aspirations"":[""EC直販比率のさらなる拡大"",""季節限定商品の開発強化"",""衛生管理体制の高度化""],""market_context"":""国内の菓子市場は縮小傾向だがギフト需要とEC市場は堅調に推移していると業界記事にある""},""current_coverage"":[{""line_name"":""火災保険(工場物件)"",""coverage_summary"":""浜松本社工場の建物および設備を対象とする火災保険"",""limit_note"":""建物3億円・設備1億円"",""special_note"":""地震保険は付帯なし""},{""line_name"":""生産物賠償責任保険(PL保険)"",""coverage_summary"":""製造した菓子製品に起因する対人対物賠償を担保"",""limit_note"":""1事故あたり1億円"",""special_note"":""リコール費用特約なし""},{""line_name"":""労働災害総合保険"",""coverage_summary"":""従業員の業務災害を法定外補償で上乗せ""" & vbLf
    s = s & ",""limit_note"":""不明"",""special_note"":""パート従業員の加入状況は不明""}],""field_insights"":[{""note"":""社長は先代からの工場を大事にしており設備更新には慎重だと聞いている"",""tag"":""constraint""},{""note"":""EC直販の物流は外部委託先1社に依存しておりトラブル時の代替が無いらしい"",""tag"":""risk_clue""}],""missing_info"":[{""item"":""止水板などの水災対策の導入状況"",""why_needed"":""施設・自然災害リスクの評価に必要なため""},{""item"":""EC物流委託先との契約内容(損害時の責任分担)"",""why_needed"":""サプライチェーンリスクの評価に必要なため""}],""input_quality"":{""coverage"":[{""aspect"":""profile"",""status"":""ok""},{""aspect"":""business"",""status"":""ok""},{""aspect"":""sites"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""history"",""status"":""ok""},{""aspect"":""news"",""status"":""ok""},{""aspect"":""hr"",""status"":""ok""},{""aspect"":""finance_risk"",""status"":""ok""},{""aspect"":""sales_memo"",""status"":""ok""},{""aspect"":""sns"",""status"":""partial""},{""aspect"":""competitors"",""status"":""ok""},{""aspect"":""market"",""status"":""partial""},{""aspect"":""finance"",""status"":""ok""},{""aspect"":""insurance_ctx"",""status"":""ok""},{""aspect"":""hazard"",""status"":""ok""}],""overall"":""high"",""advice"":""追加不要""},""research_requests"":[]}"

    BuildS1RnwJson = s
End Function

Public Function BuildS2NewJson() As String
    Dim s As String

    s = s & "{""risks"":[{""risk_no"":1,""category"":""manufacturing_quality"",""risk_name"":""アレルゲン表示誤りによる自主回収"",""scenario"":""繁忙期の増産で表示チェック体制が薄まり、アレルゲン表示を誤った商品が出荷されて自主回収に至る"",""status"":""proposed"",""frequency"":""mid"",""impact"":""large"",""frequency_score"":3,""impact_score"":4,""evidence"":{""quote"":""季節限定商品の開発強化"",""source"":""hp""},""insurability"":{""transferability"":""cover"",""line_note"":""PL保険のリコール費用特約の有無を確認"",""control_note"":""表示チェック体制の二重化""},""loss_scale_note"":"""",""check_points"":[""表示チェック工程のダブルチェック運用有無"",""アレルゲン専用ライン分離の有無""],""preventions""" & vbLf
    s = s & ":[{""measure"":""出荷前のアレルゲン表示ダブルチェック体制を構築する"",""related_menu_id"":""M-0012""}]},{""risk_no"":2,""category"":""supply_chain"",""risk_name"":""乳製品・果実の調達難による生産停滞"",""scenario"":""産地不作や輸入価格の高騰で主要原材料の調達が滞り、生産計画に遅れが生じる"",""status"":""proposed"",""frequency"":""mid"",""impact"":""mid"",""frequency_score"":3,""impact_score"":3,""evidence"":{""quote"":""季節や産地不作の影響を受けやすい"",""source"":""hp""},""insurability"":{""transferability"":""partial"",""line_note"":""仕入先費用保険の適用可否を確認"",""control_note"":""仕入先の複数化・在庫積み増し""},""loss_scale_note"":"""",""check_points"":[" & vbLf
    s = s & """主要原材料の仕入先数"",""在庫の積み増し余地""],""preventions"":[{""measure"":""主要原材料の仕入先複数化を検討する"",""related_menu_id"":""""}]},{""risk_no"":3,""category"":""facility_bcp"",""risk_name"":""浸水による浜松本社工場の操業停止"",""scenario"":""豪雨により本社工場周辺が浸水想定深に達し、主力ラインが長期間停止する"",""status"":""proposed"",""frequency"":""low"",""impact"":""large"",""frequency_score"":2,""impact_score"":4,""evidence"":{""quote"":""浸水想定区域内(想定浸水深0.5メートルから3.0メートル)"",""source"":""hp""},""insurability"":{""transferability"":""cover"",""line_note"":""火災保険の水災担保・利益保険の要否を確認"",""control_note"":""止水板の設置・BCP発動基準の整備""" & vbLf
    s = s & "},""loss_scale_note"":"""",""check_points"":[""止水板等の水災対策の有無"",""BCP発動基準の整備状況""],""preventions"":[{""measure"":""止水板設置とBCP発動基準を整備する"",""related_menu_id"":""M-0012""},{""measure"":""主力ラインの一部を高所へ移設することを検討する"",""related_menu_id"":""""}]},{""risk_no"":4,""category"":""hr_labor"",""risk_name"":""パート比率の高さによる繁忙期の人手不足"",""scenario"":""繁忙期にパート従業員が確保できず、増産計画に対応できなくなる"",""status"":""proposed"",""frequency"":""high"",""impact"":""mid"",""frequency_score"":4,""impact_score"":3,""evidence"":{""quote"":""パート従業員の比率が高く"",""source"":""hp""},""insurability"":{""transferability""" & vbLf
    s = s & ":""hard"",""line_note"":""労災上乗せの対象範囲を確認"",""control_note"":""繁忙期の応援体制整備・多能工化""},""loss_scale_note"":"""",""check_points"":[""繁忙期の応援体制の有無"",""多能工化の進捗""],""preventions"":[{""measure"":""多能工化研修を実施する"",""related_menu_id"":""""}]},{""risk_no"":5,""category"":""digital_info"",""risk_name"":""ECサイト刷新に伴う個人情報漏えい"",""scenario"":""刷新した自社ECサイトの脆弱性を突かれ、顧客の個人情報が外部に流出する"",""status"":""proposed"",""frequency"":""mid"",""impact"":""large"",""frequency_score"":3,""impact_score"":4,""evidence"":{""quote"":""自社ECサイトを刷新した"",""source"":""hp""},""insurability"":{""transferability""" & vbLf
    s = s & ":""cover"",""line_note"":""サイバー保険の個人情報漏えい対応費用担保を確認"",""control_note"":""脆弱性診断の定期実施""},""loss_scale_note"":"""",""check_points"":[""ECサイトの脆弱性診断の実施状況"",""個人情報保護体制の整備状況""],""preventions"":[{""measure"":""ECサイトの脆弱性診断を定期実施する"",""related_menu_id"":""""}]},{""risk_no"":6,""category"":""strategy_market"",""risk_name"":""国内菓子市場の縮小による需要減"",""scenario"":""国内の菓子市場が緩やかに縮小し、既存商品の販売数量が伸び悩む"",""status"":""proposed"",""frequency"":""mid"",""impact"":""mid"",""frequency_score"":3,""impact_score"":3,""evidence"":{""quote"":""国内の菓子市場は縮小傾向"",""source"":""knowledge""" & vbLf
    s = s & "},""insurability"":{""transferability"":""hard"",""line_note"":""需要減そのものは保険で移転できない"",""control_note"":""EC直販拡大・新商品開発での需要創出""},""loss_scale_note"":"""",""check_points"":[""EC直販比率の推移"",""新商品の開発ペース""],""preventions"":[{""measure"":""EC直販チャネルの拡大を継続する"",""related_menu_id"":""""}]},{""risk_no"":7,""category"":""legal_regulatory"",""risk_name"":""食品表示法改正への対応遅れ"",""scenario"":""食品表示制度の改正に表示ルールの改訂が追いつかず、指摘や回収につながる"",""status"":""proposed"",""frequency"":""low"",""impact"":""mid"",""frequency_score"":2,""impact_score"":3,""evidence"":{""quote"":""食品表示""," & vbLf
    s = s & """source"":""knowledge""},""insurability"":{""transferability"":""partial"",""line_note"":""法令違反由来の損害は保険対象外の場合が多く要確認"",""control_note"":""表示ルールの定期チェック体制""},""loss_scale_note"":"""",""check_points"":[""表示ルール改訂への追随体制""],""preventions"":[{""measure"":""表示ルールの定期チェック体制を整備する"",""related_menu_id"":""""}]},{""risk_no"":8,""category"":""sales_customer"",""risk_name"":""EC物流委託先の代替不能による配送遅延"",""scenario"":""EC直販の配送を担う外部委託先1社にトラブルが起き、代替手段が無いまま配送が滞る"",""status"":""proposed"",""frequency"":""mid"",""impact"":""mid"",""frequency_score"":3,""impact_score""" & vbLf
    s = s & ":3,""evidence"":{""quote"":""外部委託先1社に依存しており"",""source"":""memo""},""insurability"":{""transferability"":""partial"",""line_note"":""利益保険(BI)の契約者外要因担保の可否を確認"",""control_note"":""物流委託先の複線化検討""},""loss_scale_note"":"""",""check_points"":[""物流委託先の複線化の可否""],""preventions"":[{""measure"":""物流委託先の複線化を検討する"",""related_menu_id"":""""}]}],""gaps"":[],""emerging_risks"":[{""risk_name"":""気候変動による原料調達難と工場の高温化"",""category"":""facility_bcp"",""horizon"":""mid_long"",""scenario"":""果実や乳製品の産地不作が頻発すると原料調達が不安定化し、浜松2工場の夏季高温化が空調負荷や品質管理コストを押し上げる可能性がある"",""evidence_quote""" & vbLf
    s = s & ":""季節や産地不作の影響を受けやすい"",""evidence_source"":""hp"",""proposal_hint"":""原料調達の多元化提案と工場の空調・BCP診断メニューの提案に接続できる""}],""open_questions"":[""主要原材料の仕入先は何社に分散しているか"",""EC物流の代替委託先を確保する計画はあるか""]}"

    BuildS2NewJson = s
End Function

Public Function BuildS2RnwJson() As String
    Dim s As String

    s = s & "{""risks"":[{""risk_no"":1,""category"":""manufacturing_quality"",""risk_name"":""アレルゲン表示誤りによる自主回収"",""scenario"":""繁忙期の増産で表示チェック体制が薄まり、アレルゲン表示を誤った商品が出荷されて自主回収に至る"",""status"":""proposed"",""frequency"":""mid"",""impact"":""large"",""frequency_score"":3,""impact_score"":4,""evidence"":{""quote"":""季節限定商品の開発強化"",""source"":""hp""},""insurability"":{""transferability"":""cover"",""line_note"":""PL保険のリコール費用特約の有無を確認"",""control_note"":""表示チェック体制の二重化""},""loss_scale_note"":"""",""check_points"":[""表示チェック工程のダブルチェック運用有無"",""アレルゲン専用ライン分離の有無""],""preventions""" & vbLf
    s = s & ":[{""measure"":""出荷前のアレルゲン表示ダブルチェック体制を構築する"",""related_menu_id"":""M-0012""}]},{""risk_no"":2,""category"":""supply_chain"",""risk_name"":""乳製品・果実の調達難による生産停滞"",""scenario"":""産地不作や輸入価格の高騰で主要原材料の調達が滞り、生産計画に遅れが生じる"",""status"":""proposed"",""frequency"":""mid"",""impact"":""mid"",""frequency_score"":3,""impact_score"":3,""evidence"":{""quote"":""季節や産地不作の影響を受けやすい"",""source"":""hp""},""insurability"":{""transferability"":""partial"",""line_note"":""仕入先費用保険の適用可否を確認"",""control_note"":""仕入先の複数化・在庫積み増し""},""loss_scale_note"":"""",""check_points"":[" & vbLf
    s = s & """主要原材料の仕入先数"",""在庫の積み増し余地""],""preventions"":[{""measure"":""主要原材料の仕入先複数化を検討する"",""related_menu_id"":""""}]},{""risk_no"":3,""category"":""facility_bcp"",""risk_name"":""浸水による浜松本社工場の操業停止"",""scenario"":""豪雨により本社工場周辺が浸水想定深に達し、主力ラインが長期間停止する"",""status"":""proposed"",""frequency"":""low"",""impact"":""large"",""frequency_score"":2,""impact_score"":4,""evidence"":{""quote"":""浸水想定区域内(想定浸水深0.5メートルから3.0メートル)"",""source"":""hp""},""insurability"":{""transferability"":""cover"",""line_note"":""火災保険の水災担保・利益保険の要否を確認"",""control_note"":""止水板の設置・BCP発動基準の整備""" & vbLf
    s = s & "},""loss_scale_note"":"""",""check_points"":[""止水板等の水災対策の有無"",""BCP発動基準の整備状況""],""preventions"":[{""measure"":""止水板設置とBCP発動基準を整備する"",""related_menu_id"":""M-0012""},{""measure"":""主力ラインの一部を高所へ移設することを検討する"",""related_menu_id"":""""}]},{""risk_no"":4,""category"":""hr_labor"",""risk_name"":""パート比率の高さによる繁忙期の人手不足"",""scenario"":""繁忙期にパート従業員が確保できず、増産計画に対応できなくなる"",""status"":""proposed"",""frequency"":""high"",""impact"":""mid"",""frequency_score"":4,""impact_score"":3,""evidence"":{""quote"":""パート従業員の比率が高く"",""source"":""hp""},""insurability"":{""transferability""" & vbLf
    s = s & ":""hard"",""line_note"":""労災上乗せの対象範囲を確認"",""control_note"":""繁忙期の応援体制整備・多能工化""},""loss_scale_note"":"""",""check_points"":[""繁忙期の応援体制の有無"",""多能工化の進捗""],""preventions"":[{""measure"":""多能工化研修を実施する"",""related_menu_id"":""""}]},{""risk_no"":5,""category"":""digital_info"",""risk_name"":""ECサイト刷新に伴う個人情報漏えい"",""scenario"":""刷新した自社ECサイトの脆弱性を突かれ、顧客の個人情報が外部に流出する"",""status"":""proposed"",""frequency"":""mid"",""impact"":""large"",""frequency_score"":3,""impact_score"":4,""evidence"":{""quote"":""自社ECサイトを刷新した"",""source"":""hp""},""insurability"":{""transferability""" & vbLf
    s = s & ":""cover"",""line_note"":""サイバー保険の個人情報漏えい対応費用担保を確認"",""control_note"":""脆弱性診断の定期実施""},""loss_scale_note"":"""",""check_points"":[""ECサイトの脆弱性診断の実施状況"",""個人情報保護体制の整備状況""],""preventions"":[{""measure"":""ECサイトの脆弱性診断を定期実施する"",""related_menu_id"":""""}]},{""risk_no"":6,""category"":""strategy_market"",""risk_name"":""国内菓子市場の縮小による需要減"",""scenario"":""国内の菓子市場が緩やかに縮小し、既存商品の販売数量が伸び悩む"",""status"":""proposed"",""frequency"":""mid"",""impact"":""mid"",""frequency_score"":3,""impact_score"":3,""evidence"":{""quote"":""国内の菓子市場は縮小傾向"",""source"":""knowledge""" & vbLf
    s = s & "},""insurability"":{""transferability"":""hard"",""line_note"":""需要減そのものは保険で移転できない"",""control_note"":""EC直販拡大・新商品開発での需要創出""},""loss_scale_note"":"""",""check_points"":[""EC直販比率の推移"",""新商品の開発ペース""],""preventions"":[{""measure"":""EC直販チャネルの拡大を継続する"",""related_menu_id"":""""}]},{""risk_no"":7,""category"":""legal_regulatory"",""risk_name"":""食品表示法改正への対応遅れ"",""scenario"":""食品表示制度の改正に表示ルールの改訂が追いつかず、指摘や回収につながる"",""status"":""proposed"",""frequency"":""low"",""impact"":""mid"",""frequency_score"":2,""impact_score"":3,""evidence"":{""quote"":""食品表示""," & vbLf
    s = s & """source"":""knowledge""},""insurability"":{""transferability"":""partial"",""line_note"":""法令違反由来の損害は保険対象外の場合が多く要確認"",""control_note"":""表示ルールの定期チェック体制""},""loss_scale_note"":"""",""check_points"":[""表示ルール改訂への追随体制""],""preventions"":[{""measure"":""表示ルールの定期チェック体制を整備する"",""related_menu_id"":""""}]},{""risk_no"":8,""category"":""sales_customer"",""risk_name"":""EC物流委託先の代替不能による配送遅延"",""scenario"":""EC直販の配送を担う外部委託先1社にトラブルが起き、代替手段が無いまま配送が滞る"",""status"":""proposed"",""frequency"":""mid"",""impact"":""mid"",""frequency_score"":3,""impact_score""" & vbLf
    s = s & ":3,""evidence"":{""quote"":""外部委託先1社に依存しており"",""source"":""memo""},""insurability"":{""transferability"":""partial"",""line_note"":""利益保険(BI)の契約者外要因担保の可否を確認"",""control_note"":""物流委託先の複線化検討""},""loss_scale_note"":"""",""check_points"":[""物流委託先の複線化の可否""],""preventions"":[{""measure"":""物流委託先の複線化を検討する"",""related_menu_id"":""""}]}],""gaps"":[{""gap_no"":1,""gap_type"":""uninsured"",""target"":""サイバー(個人情報漏えい)"",""description"":""EC刷新に伴う個人情報漏えいに備えるサイバー保険が現契約に見当たらない"",""risk_evidence"":""自社ECサイトを刷新した"",""coverage_evidence"":""該当契約なし""},{""gap_no""" & vbLf
    s = s & ":2,""gap_type"":""underinsured"",""target"":""生産物賠償責任保険(PL保険)"",""description"":""リコール費用特約が付いておらず自主回収時の費用が担保されない疑いがある"",""risk_evidence"":""アレルゲン表示誤りによる自主回収"",""coverage_evidence"":""1事故あたり1億円・リコール費用特約なし""},{""gap_no"":3,""gap_type"":""overlap"",""target"":""労働災害総合保険"",""description"":""パート従業員向けの法定外補償の範囲が別契約と重複している可能性がある"",""risk_evidence"":""パート従業員の比率が高く"",""coverage_evidence"":""パート従業員の加入状況は不明""}],""emerging_risks"":[],""open_questions"":[""主要原材料の仕入先は何社に分散しているか"",""EC物流の代替委託先を確保する計画はあるか""]}"

    BuildS2RnwJson = s
End Function
