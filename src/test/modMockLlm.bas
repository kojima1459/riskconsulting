Attribute VB_Name = "modMockLlm"
Option Explicit

' broken_json_once 専用の状態1bit(15章§8.2の状態レス原則の唯一の例外)。
' 初回呼出だけ破損応答を返し、以降は正常応答。ResetFaultOnce でリセット。
Private mOnceFired As Boolean

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
' 公開契約(14章§6。3本):
'   MockResponse(stepName, variantName, fault) : modGatewayRPN のmock分岐が呼ぶ
'     ゲートウェイ入口。応答本文を持たず下の2本へ振り分けるだけ。stepName は
'     s1/s2/s2r/s3/s3r/s4/pf/s2c/s3c/sp、variantName は ResolveMockVariant の
'     戻り値(new/renewal/hit/clean/common)、fault は config mock_fault の値。
'   ResponseById(mockId)          : 15章§8.1の11 IDで正常応答。表外のIDは ""。
'   FaultResponse(faultKind, step): 15章§8.2の11値の障害注入応答。空なら ""。
'   戻り値は生応答そのもの。JSON防衛線(抽出・検証)は呼び出し側の責務であり、
'   ここでは通さない(14章§5)。
'
' 正常応答(15章§8.1。7 step種・11応答): MK-S1-NEW / MK-S1-RNW / MK-S2-NEW /
'   MK-S2-RNW / MK-S3 / MK-S4 / MK-PF / MK-S2C-HIT / MK-S2C-CLEAN /
'   MK-S3C-HIT / MK-S3C-CLEAN。各応答は自分のバリアント文脈(NEWはcase_type=new、
'   RNWはrenewal。共通応答は両文脈)でmodValidateに合格する(受入条件1)。
'   sp(壁打ち)はスキーマを持たないため自由文を返す(表外。IDを持たない)。
'
' 障害注入(15章§8.2の表が正。config mock_fault。既定は空=正常応答のみ)。11値と
' 適用stepは FaultBody の Select Case が実体で、broken_json=末尾の閉じ括弧欠落 /
' enum_violation=s2のcategoryをenum外"quality"へ / count_violation=s3のstories
' 2件 / ghost_id=s3のmenu_idsに"M-9999"混入 / empty=空文字 / limit=固定の上限
' 文字列 / fake_err=先頭行"#ERR:E0201:"+正常JSON本体 / ribbon_429・
' ribbon_disconnect・ribbon_content_filter=実リボンの定型失敗文(裁定書24 A-1。
' 実体は modMockLlm2)。11値以外は正常応答へフォールバックする(config入力ミス
' でE2E全体を暴走させないため)。
'
' 【状態レス規約(15章§8.2)】fault指定中は毎回同じ応答を返す。「最初の1回だけ
' 壊す」型の内部カウンタは持たない(乱数・現在時刻を使わないのと同じ理由=再現性。
' カウンタはテスト実行順に結果が依存し、単体テストからは初期化できず失敗の再現が
' できなくなる)。修復リトライ経路(2回目は正常)の検証は、config mock_fault を
' 当該値から空へ切り替えた2ラン構成で行う。
'
' モジュール分割(modMockLlm1..n。12章§2): 1モジュール30,000字契約のため、定型
' 項目が多いS3以降の6応答(+count_violation派生)はmodMockLlm2へ切り出した。
' 判定ロジックは本モジュールに閉じ、modMockLlm2は文字列返却関数の集まりに徹する。
'
' 移植元: 新規(PoCに対応物なし。12章§2)。
' ==============================================================================

' ==============================================================================
' MockResponse - ゲートウェイ入口(14章§6・§4・T-14契約)
' ------------------------------------------------------------------------------
'   自分では応答本文を持たず、fault の有無で ResponseById / FaultResponse へ
'   振り分けるだけ。fault 指定時も variantName を尊重した本文へ障害を注入する
'   (renewal案件にnew文脈の本文を返すと、fake_err のように「本文は正常」で
'   あるべき障害でmodValidateが落ちてしまうため)。
' ==============================================================================
Public Function MockResponse(ByVal stepName As String, ByVal variantName As String, _
                             ByVal fault As String) As String
    Dim stepKey As String
    Dim f As String

    stepKey = LCase$(Trim$(stepName))
    f = Trim$(fault)

    If LenB(f) = 0 Then
        MockResponse = BodyFor(stepKey, variantName)
    Else
        MockResponse = FaultBody(f, stepKey, variantName)
    End If
End Function

' ==============================================================================
' ResponseById - 15章§8.1の表の mock ID で正常応答を返す(14章§6)
' ------------------------------------------------------------------------------
'   キーの正は15章§8.1の11 ID と15章§5.6 の MK-S5(計12 ID)。表に無いIDは "" を返す(呼び元の
'   ClassifyResponse がE0202として扱う)。決定的=乱数・現在時刻・呼び出し回数に
'   依存しない。壁打ち(sp)は表にIDを持たないため本関数の対象外。
' ==============================================================================
Public Function ResponseById(ByVal mockId As String) As String
    Select Case UCase$(Trim$(mockId))
        Case "MK-S1-NEW"
            ResponseById = BuildS1NewJson()
        Case "MK-S1-RNW"
            ResponseById = BuildS1RnwJson()
        Case "MK-S2-NEW"
            ResponseById = modMockLlm3.BuildS2NewJson()
        Case "MK-S2-RNW"
            ResponseById = modMockLlm3.BuildS2RnwJson()
        Case "MK-S3"
            ResponseById = modMockLlm2.BuildS3Json()
        Case "MK-S4"
            ResponseById = modMockLlm2.BuildS4Json()
        Case "MK-S5"
            ResponseById = modMockLlm4.BuildS5Json()
        Case "MK-PF"
            ResponseById = modMockLlm2.BuildPfJson()
        Case "MK-S2C-HIT"
            ResponseById = modMockLlm2.BuildS2CHitJson()
        Case "MK-S2C-CLEAN"
            ResponseById = modMockLlm2.BuildS2CCleanJson()
        Case "MK-S3C-HIT"
            ResponseById = modMockLlm2.BuildS3CHitJson()
        Case "MK-S3C-CLEAN"
            ResponseById = modMockLlm2.BuildS3CCleanJson()
        Case Else
            ResponseById = ""
    End Select
End Function

' ==============================================================================
' FaultResponse - 15章§8.2の障害注入応答(14章§6)
' ------------------------------------------------------------------------------
'   【状態レス】同じ引数なら常に同じ応答を返す。faultKind が空なら ""(15章§8.2
'   の「mock_fault が空のときは正常応答のみ」はゲートウェイ入口 MockResponse の
'   責務であって、本関数は障害の形だけを返す口である)。
'   バリアントは各stepの既定(s1/s2=new、s2c/s3c=clean)を使う。
' ==============================================================================
Public Function FaultResponse(ByVal faultKind As String, ByVal stepName As String) As String
    Dim f As String

    f = Trim$(faultKind)
    If LenB(f) = 0 Then
        FaultResponse = ""
        Exit Function
    End If
    FaultResponse = FaultBody(f, LCase$(Trim$(stepName)), "")
End Function

' FaultBody - 障害注入の実体(15章§8.2)。MockResponse と FaultResponse が共有し
'   注入の規則を1箇所に閉じる。stepKey は小文字化・トリム済みの前提。
Private Function FaultBody(ByVal f As String, ByVal stepKey As String, _
                           ByVal variantName As String) As String
    Dim body As String

    ' 正常応答を組み立てる前に決まる2値。
    If f = "empty" Then
        FaultBody = ""
        Exit Function
    ElseIf f = "limit" Then
        FaultBody = LimitFaultText()
        Exit Function
    ElseIf f = "ribbon_429" Then
        FaultBody = modMockLlm2.RibbonErr429Text()
        Exit Function
    ElseIf f = "ribbon_disconnect" Then
        FaultBody = modMockLlm2.RibbonDisconnectText()
        Exit Function
    ElseIf f = "ribbon_content_filter" Then
        FaultBody = modMockLlm2.RibbonContentFilterText()
        Exit Function
    End If

    body = BodyFor(stepKey, variantName)

    Select Case f
        Case "fake_err"
            ' 帯域外成否規約の検査用。ok=Trueのまま素通しされる想定
            ' (14章§4・§6。判定材料にしないのはmodGatewayRPN側の責務)。
            FaultBody = "#ERR:E0201:偽装エラーです" & vbLf & body
        Case "broken_json"
            FaultBody = BreakJsonTail(body)
        Case "broken_json_once"
            ' 初回のみ破損(修復リトライの成功系=validate_result=repairedの検証用)。
            If mOnceFired Then
                FaultBody = body
            Else
                mOnceFired = True
                FaultBody = BreakJsonTail(body)
            End If
        Case "enum_violation"
            If stepKey = "s2" Then
                FaultBody = InjectEnumViolationS2(body)
            Else
                FaultBody = body
            End If
        Case "count_violation"
            If stepKey = "s3" Then
                FaultBody = modMockLlm2.BuildS3CountViolationJson()
            Else
                FaultBody = body
            End If
        Case "ghost_id"
            If stepKey = "s3" Then
                FaultBody = InjectGhostIdS3(body)
            Else
                FaultBody = body
            End If
        Case Else
            ' 未知の値は正常応答へフォールバック(config入力ミスの暴走防止)。
            FaultBody = body
    End Select
End Function

' BodyFor - stepKey とバリアントから正常応答本文を得る(15章§8.1)。表にIDを
'   持たない壁打ち(sp)だけは直接返し、それ以外は mock ID へ写して ResponseById
'   へ渡す(応答本文の分岐を1箇所=ResponseById に閉じる)。
Private Function BodyFor(ByVal stepKey As String, ByVal variantName As String) As String
    If stepKey = "sp" Then
        BodyFor = SpTextMock()
    Else
        BodyFor = ResponseById(MockIdFor(stepKey, variantName))
    End If
End Function

' MockIdFor - (stepKey, variantName) -> 15章§8.1の mock ID。表に無い組合せは ""
'   (呼び元の ResponseById が "" を返し、ClassifyResponse がE0202とする)。
'   s2r / s3r(15章§4.7 改訂パス)は s2 / s3 と同じ応答を使う。
Private Function MockIdFor(ByVal stepKey As String, ByVal variantName As String) As String
    Dim v As String
    v = LCase$(Trim$(variantName))

    Select Case stepKey
        Case "s1"
            If v = "renewal" Then MockIdFor = "MK-S1-RNW" Else MockIdFor = "MK-S1-NEW"
        Case "s2", "s2r"
            If v = "renewal" Then MockIdFor = "MK-S2-RNW" Else MockIdFor = "MK-S2-NEW"
        Case "s3", "s3r"
            MockIdFor = "MK-S3"
        Case "s4"
            MockIdFor = "MK-S4"
        Case "s5"
            MockIdFor = "MK-S5"
        Case "pf"
            MockIdFor = "MK-PF"
        Case "s2c"
            If v = "hit" Then MockIdFor = "MK-S2C-HIT" Else MockIdFor = "MK-S2C-CLEAN"
        Case "s3c"
            If v = "hit" Then MockIdFor = "MK-S3C-HIT" Else MockIdFor = "MK-S3C-CLEAN"
        Case Else
            ' 未定義のstepName(wt/fg等のPhase 1.5含む)は空文字を返す。
            MockIdFor = ""
    End Select
End Function

' ------------------------------------------------------------------------------
' 障害注入ヘルパー(15章§8.2)
' ------------------------------------------------------------------------------
' 利用上限を示す応答の実体(15章§8.2で1文字列に固定)。14章§2の
' LooksLikeLimitError はこの文字列だけを見る(語彙を2箇所に書かない)。
Private Function LimitFaultText() As String
    LimitFaultText = "#LIMIT: 本日のAIリボン利用上限に達しました(LimitCheck)"
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

    s = s & "{""company_name"":""株式会社浜松スイーツファクトリー"",""business_summary"":""静岡県浜松市に本社を置く洋菓子・和菓子の製造販売企業。自社工場2拠点で焼き菓子を中心に生産し、直営店・卸売に加えEC直販を伸ばしている。"",""main_products"":[""季節限定焼き菓子ギフトセット"",""洋菓子詰め合わせ(EC限定)"",""和菓子詰め合わせ""],""processes"":[""浜松本社工場で主力の焼き菓子を一貫生産している"",""積志第二工場は繁忙期のギフト商品増産に対応している"",""EC直販分は本社工場から直送する体制である""],""locations"":[{""name"":""浜松本社工場"",""type"":""工場"",""address"":""不明"",""hazard_note"":""浸水想定区域内(想定浸水深0.5メートルから3.0メートル)との記載あり"",""notes"":""主力ラインを持つ最大拠点""" & vbLf
    s = s & "},{""name"":""積志第二工場"",""type"":""工場"",""address"":""不明"",""hazard_note"":""不明"",""notes""" & vbLf
    s = s & ":""繁忙期の増産用ライン""" & vbLf
    s = s & "},{""name"":""本社直営店"",""type"":""店舗"",""address"":""不明"",""hazard_note"":""不明"",""notes"":""不明""}],""supply_chain"":{""key_materials"":[""小麦粉"",""バター"",""卵"",""国産果実(みかん・いちご)""],""notes"":""主要原材料は複数の国内商社経由で調達しており、乳製品・果実の一部は季節や産地不作の影響を受けやすいと記載がある""" & vbLf
    s = s & "},""customers"":{""segments"":[""個人ギフト需要"",""法人向け贈答需要"",""EC直販の個人顧客""],""channels"":[""直営店"",""卸売(菓子問屋)"",""自社ECサイト""]},""workforce_notes"":""製造ラインはパート従業員の比率が高く、繁忙期は季節雇用で補っていると採用ページに記載がある"",""management_notes"":""EC直販比率の拡大を経営方針として掲げ、直近で自社ECサイトを刷新したと社長挨拶で述べている"",""strategy_outlook"":{""mvv""" & vbLf
    s = s & ":""地域に根差した菓子づくりで顧客の特別な日に寄り添うことを掲げている"",""aspirations"":[""EC直販比率のさらなる拡大"",""季節限定商品の開発強化"",""衛生管理体制の高度化""],""market_context"":""国内の菓子市場は縮小傾向だがギフト需要とEC市場は堅調に推移していると業界記事にある""" & vbLf
    s = s & "},""current_coverage"":[{""line_name"":""労働災害総合保険(見立て)"",""coverage_summary"":""(見立て)元請の包括契約に労災上乗せが乗っている模様"",""limit_note"":""不明"",""special_note"":""不明"",""certainty"":""assumed""}]," & vbLf
    s = s & """financials"":{""fiscal_year"":""不明"",""net_assets"":""不明"",""sales"":""不明"",""operating_profit"":""不明"",""source"":""unknown"",""note"":""不明""" & vbLf
    s = s & "},""field_insights"":[{""note"":""社長は先代からの工場を大事にしており設備更新には慎重だと聞いている"",""tag"":""constraint""" & vbLf
    s = s & "},{""note"":""EC直販の物流は外部委託先1社に依存しておりトラブル時の代替が無いらしい"",""tag"":""risk_clue""" & vbLf
    s = s & "},{""note"":""競合の同業他社が値上げに踏み切ったため価格面では当社が優位に見えるとのこと"",""tag"":""competitor""}],""missing_info"":[{""item"":""浜松本社工場の建物構造(耐火・耐震等級)"",""why_needed"":""施設・自然災害リスクの評価に必要なため"",""kind"":""not_found""}" & vbLf
    s = s & ",{""item"":""EC物流委託先との契約内容(損害時の責任分担)"",""why_needed"":""サプライチェーンリスクの評価に必要なため"",""kind"":""hearing_only""" & vbLf
    s = s & "},{""item"":""直近の食品衛生関連の指摘・自主回収の有無"",""why_needed"":""製造・品質リスクの評価に必要なため"",""kind"":""undisclosed""}],""input_quality"":{""coverage"":[{""aspect"":""profile"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""business"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""sites"",""status"":""partial""" & vbLf
    s = s & "},{""aspect"":""history"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""news"",""status"":""partial""" & vbLf
    s = s & "},{""aspect"":""hr"",""status"":""partial""" & vbLf
    s = s & "},{""aspect"":""finance_risk"",""status"":""missing""" & vbLf
    s = s & "},{""aspect"":""sales_memo"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""sns""" & vbLf
    s = s & ",""status"":""missing""" & vbLf
    s = s & "},{""aspect"":""competitors"",""status"":""missing""" & vbLf
    s = s & "},{""aspect"":""market"",""status"":""partial""" & vbLf
    s = s & "},{""aspect"":""finance"",""status"":""missing""" & vbLf
    s = s & "},{""aspect"":""insurance_ctx"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""hazard"",""status"":""partial""}],""overall"":""mid"",""advice"":""有価証券報告書相当の財務・リスク情報とSNS評判、競合動向、市況情報を追加すると仮説の精度が上がる""" & vbLf
    s = s & "},""research_requests"":[{""purpose"":""finance_risk観点(財務・事業リスクの記載)を埋めるための調査"",""prompt_text"":""静岡県浜松市の菓子メーカーである株式会社浜松スイーツファクトリー(本社所在地:静岡県浜松市)について、EDINETまたは同社の公式IRページに掲載されている有価証券報告書または決算公告の「事業等のリスク」に相当する記載内容を調査してください。該当する事実が見当たらない場合は「見当たらない」、取得できない項目は「取得できず」と明記してください。各項目には出典URLを付けてください。まとめサイト・就活情報サイト・個人ブログは情報源に使わないでください。有価証券報告書や決算公告が存在しない場合はその旨を報告してください。""" & vbLf
    s = s & "},{""purpose"":""sns観点(SNS・口コミの評判傾向)を埋めるための調査"",""prompt_text"":""静岡県浜松市の菓子メーカーである株式会社浜松スイーツファクトリー(本社所在地:静岡県浜松市)について、SNSや口コミサイトでの評判傾向(品質・接客・労働環境・炎上の有無)を調査してください。該当する事実が見当たらない場合は「見当たらない」、取得できない項目は「取得できず」と明記してください。各項目には出典URLを付けてください。まとめサイト・就活情報サイト・個人ブログは情報源に使わないでください。""}]," & vbLf
    s = s & """sources"":[{""label"":""会社概要(公式HP)"",""url"":""https://example.co.jp/company/"",""aspect"":""profile""" & vbLf
    s = s & "},{""label"":""工場紹介(公式HP)"",""url"":""https://example.co.jp/factory/"",""aspect"":""sites""}]}"

    BuildS1NewJson = s
End Function

Public Function BuildS1RnwJson() As String
    Dim s As String

    s = s & "{""company_name"":""株式会社浜松スイーツファクトリー"",""business_summary"":""静岡県浜松市に本社を置く洋菓子・和菓子の製造販売企業。自社工場2拠点で焼き菓子を中心に生産し、直営店・卸売に加えEC直販を伸ばしている。"",""main_products"":[""季節限定焼き菓子ギフトセット"",""洋菓子詰め合わせ(EC限定)"",""和菓子詰め合わせ""],""processes"":[""浜松本社工場で主力の焼き菓子を一貫生産している"",""積志第二工場は繁忙期のギフト商品増産に対応している"",""EC直販分は本社工場から直送する体制である""],""locations"":[{""name"":""浜松本社工場"",""type"":""工場"",""address"":""不明"",""hazard_note"":""浸水想定区域内(想定浸水深0.5メートルから3.0メートル)との記載あり"",""notes"":""主力ラインを持つ最大拠点""" & vbLf
    s = s & "},{""name"":""積志第二工場"",""type"":""工場"",""address"":""不明"",""hazard_note"":""不明"",""notes""" & vbLf
    s = s & ":""繁忙期の増産用ライン""" & vbLf
    s = s & "},{""name"":""本社直営店"",""type"":""店舗"",""address"":""不明"",""hazard_note"":""不明"",""notes"":""不明""}],""supply_chain"":{""key_materials"":[""小麦粉"",""バター"",""卵"",""国産果実(みかん・いちご)""],""notes"":""主要原材料は複数の国内商社経由で調達しており、乳製品・果実の一部は季節や産地不作の影響を受けやすいと記載がある""" & vbLf
    s = s & "},""customers"":{""segments"":[""個人ギフト需要"",""法人向け贈答需要"",""EC直販の個人顧客""],""channels"":[""直営店"",""卸売(菓子問屋)"",""自社ECサイト""]},""workforce_notes"":""製造ラインはパート従業員の比率が高く、繁忙期は季節雇用で補っていると採用ページに記載がある"",""management_notes"":""EC直販比率の拡大を経営方針として掲げ、直近で自社ECサイトを刷新したと社長挨拶で述べている"",""strategy_outlook"":{""mvv""" & vbLf
    s = s & ":""地域に根差した菓子づくりで顧客の特別な日に寄り添うことを掲げている"",""aspirations"":[""EC直販比率のさらなる拡大"",""季節限定商品の開発強化"",""衛生管理体制の高度化""],""market_context"":""国内の菓子市場は縮小傾向だがギフト需要とEC市場は堅調に推移していると業界記事にある""" & vbLf
    s = s & "},""current_coverage"":[{""line_name"":""火災保険(工場物件)"",""coverage_summary"":""浜松本社工場の建物および設備を対象とする火災保険"",""limit_note"":""建物3億円・設備1億円"",""special_note"":""地震保険は付帯なし"",""certainty"":""confirmed""" & vbLf
    s = s & "},{""line_name"":""生産物賠償責任保険(PL保険)"",""coverage_summary"":""製造した菓子製品に起因する対人対物賠償を担保"",""limit_note"":""1事故あたり1億円"",""special_note"":""リコール費用特約なし"",""certainty"":""confirmed""" & vbLf
    s = s & "},{""line_name"":""労働災害総合保険"",""coverage_summary"":""従業員の業務災害を法定外補償で上乗せ""" & vbLf
    s = s & ",""limit_note"":""不明"",""special_note"":""パート従業員の加入状況は不明"",""certainty"":""confirmed""}],""financials"":{""fiscal_year"":""2025年3月期"",""net_assets"":""12億円"",""sales"":""85億円"",""operating_profit"":""3億2000万円"",""source"":""kessan_kokoku"",""note"":""決算公告の貸借対照表要旨から転記""" & vbLf
    s = s & "},""field_insights"":[{""note"":""社長は先代からの工場を大事にしており設備更新には慎重だと聞いている"",""tag"":""constraint""" & vbLf
    s = s & "},{""note"":""EC直販の物流は外部委託先1社に依存しておりトラブル時の代替が無いらしい"",""tag"":""risk_clue""}],""missing_info"":[{""item"":""止水板などの水災対策の導入状況"",""why_needed"":""施設・自然災害リスクの評価に必要なため"",""kind"":""hearing_only""" & vbLf
    s = s & "},{""item"":""売上高が資料間で食い違っている(85億円と82億円)"",""why_needed"":""規模前提が変わると補償額の妥当性が変わるため"",""kind"":""conflict""" & vbLf
    s = s & "},{""item"":""EC物流委託先との契約内容(損害時の責任分担)"",""why_needed"":""サプライチェーンリスクの評価に必要なため"",""kind"":""undisclosed""}],""input_quality"":{""coverage"":[{""aspect"":""profile"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""business"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""sites"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""history"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""news"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""hr"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""finance_risk"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""sales_memo"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""sns"",""status"":""partial""" & vbLf
    s = s & "},{""aspect"":""competitors"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""market"",""status"":""partial""" & vbLf
    s = s & "},{""aspect"":""finance"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""insurance_ctx"",""status"":""ok""" & vbLf
    s = s & "},{""aspect"":""hazard"",""status"":""ok""}],""overall"":""high"",""advice"":""追加不要""" & vbLf
    s = s & "},""research_requests"":[]," & vbLf
    s = s & """sources"":[{""label"":""会社概要(公式HP)"",""url"":""https://example.co.jp/company/"",""aspect"":""profile""" & vbLf
    s = s & "},{""label"":""決算公告"",""url"":""https://example.co.jp/ir/koukoku2025.html"",""aspect"":""finance""" & vbLf
    s = s & "},{""label"":""重ねるハザードマップ(浜松本社工場の住所)"",""url"":""https://disaportal.gsi.go.jp/"",""aspect"":""hazard""}]}"

    BuildS1RnwJson = s
End Function

' broken_json_once の状態リセット(14章§6)。テスト・E2Eシナリオの冒頭で呼ぶ。
Public Sub ResetFaultOnce()
    mOnceFired = False
End Sub
