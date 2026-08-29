Attribute VB_Name = "modGatewayRPN"
Option Explicit

' ==============================================
' modGatewayRPN - LLM呼び出しの唯一の窓口(14章§1・§2・§4・§6)
' ----------------------------------------------
' 役割: CallStep / CallChat の2本だけが外向きの呼び出し口。呼び出し元は経路を
'   意識しない。3経路分岐・アドイン検出・利用上限チェック・run_log メタ作成。
'
' R3(12章§2): Application.Run を書いてよい唯一のモジュール。R4: core層だが
'   その責務のため Application.Run と Application.AddIns だけ使ってよい9本の
'   1つ(12章§4)。Worksheets/Range/MsgBox には触れない。
'
' 成否の帯域外規約(14章§6・最重要): 成否は ByRef ok As Boolean だけで運ぶ。
'   戻り値文字列の内容(先頭の "#ERR:")では判定しない。LLM出力は内容を誘導
'   できるため、平文プレフィクスで成否を決めるとエラーUIを騙った任意文面表示が
'   成立する(姉妹PJ監査で実証)。したがって【LLM応答が "#ERR:" で始まっていても
'   ok=True のまま素通しする】。"#ERR:Exxxx:説明" は ok=False のときだけ作る。
'   例外は投げない(失敗= ok=False + 説明文字列 + modLog記録)。
'
' 移植元: PoC src/core/modGateway.bas のアドイン検出/上限文字列判定の流儀。
'   CallLLM を CallStep/CallChat へ分割し、引数値をconfig駆動に徹底(NFR-M3)、
'   帯域外成否規約と run_log を追加、RAG検索系(埋め込み)は非移植。
'
' 純ロジック(経路決定・toolN組立・待ち時間解決・応答判定・mockバリアント決定・
'   履歴切詰め)は Excel非依存の Public 関数へ切り出す(17章§1 の(c)で叩ける)。
' ==============================================

' 経路の名前(13章§2.3 llm_transport の enum)。
Public Const GW_RIBBON As String = "ribbon"
Public Const GW_DIRECT As String = "direct"
Public Const GW_MOCK As String = "mock"
' 自由対話(PL-04)の履歴区切り。台帳の確定方式(新しい順・";;;"連結)。
Public Const GW_HIST_SEP As String = ";;;"

' 上限系の定型拒否文は短い。長文の正当な応答を誤爆させないための閾値。
Private Const GW_LIMIT_PROBE_MAX As Long = 120
' 待ち時間の安全域(秒)。既定値は 13章§2.3 の llm_wait_sec と同値で、config が
' 読めないときの最後の砦(通常はconfigの値が使われる。NFR-M3)。
Private Const GW_WAIT_MIN As Long = 30
Private Const GW_WAIT_MAX As Long = 7200
Private Const GW_WAIT_DEFAULT As Long = 1200

' リボン検出のセッションキャッシュ(裁定D2)。
Private mRibbonChecked As Boolean
Private mRibbonAvailable As Boolean
' mock の批判step用バリアント交替カウンタ(乱数不使用・15章§8.1)。s2c と s3c で
' 別々に数える。1本の通し番号にすると片方が常に同じバリアントになり、改訂パスと
' 改訂スキップの両方を通すという T-28 の受入条件を満たせない。
Private mMockRoundS2c As Long
Private mMockRoundS3c As Long
' run_log の未確定行(1行=1呼び出し)。詳細は StageRun のコメント。
Private mPendingRun As TRunLogRec
Private mHasPendingRun As Boolean
' 呼び出し文脈。CallStep のシグネチャ(14章§6)には案件IDが無いため、run_log に
' 必要な「gatewayが知り得ない列」を呼び出し前に預かる。
Private mCtxCaseId As String
Private mCtxRoundNo As String
Private mCtxCaseType As String
Private mCtxKbIds As String
Private mCtxOperator As String

' ==============================================
' CallStep - Step呼び出しの唯一の口(14章§6)
' ----------------------------------------------
'   stepName  : 19章§4の12値(s1/s2/s3/s4/s2c/s3c/s2r/s3r/pf/sp/wt/fg)
'   playId    : PL-01..08。run_log の play 列へそのまま入る
'   schemaJson: direct経路の json_schema strict 用。ribbon/mockでは未使用
'   ok        : 成否の【唯一の】判定材料。戻り値の内容で判定してはならない
'   戻り値    : ok=True なら生応答そのまま(先頭が "#ERR:" でもそのまま)。
'               ok=False なら "#ERR:Exxxx:説明"(人間向け・判定材料ではない)
'   JSON防衛線(14章§5)はここでは通さない(抽出・正規化・検証は呼び出し側)。
' ==============================================
Public Function CallStep(ByVal stepName As String, ByVal playId As String, _
                         ByVal systemPrompt As String, ByVal userPrompt As String, _
                         ByVal schemaJson As String, ByRef ok As Boolean, _
                         Optional ByRef latencyMs As Long = 0) As String
    Dim t0 As Double
    Dim route As String
    Dim rawBody As String
    Dim errCode As String
    Dim errMsg As String
    Dim modelUsed As String
    Dim detailText As String

    ok = False
    latencyMs = 0
    t0 = Timer

    On Error GoTo Failed

    ' 前回の未回収行を先に落とす(14章§1「全呼び出しを1行記録」の取りこぼし防止)
    FlushPendingRun

    route = CurrentTransport()
    Select Case route
        Case GW_MOCK
            rawBody = MockStep(stepName, modelUsed, detailText, errCode, errMsg)
        Case GW_DIRECT
            rawBody = DirectStep(stepName, systemPrompt, userPrompt, schemaJson, _
                                 modelUsed, errCode, errMsg)
        Case Else
            rawBody = RibbonStep(stepName, systemPrompt, userPrompt, _
                                 modelUsed, errCode, errMsg)
    End Select
    GoTo Finish

Failed:
    errCode = "E0202"
    errMsg = Err.Description
    modLog.LogError errCode, "modGatewayRPN.CallStep", _
                    "step=" & stepName & " transport=" & route, Err.Number
    Resume Finish

Finish:
    latencyMs = CLng(modUtil.ElapsedMsSince(t0))
    If LenB(errCode) = 0 Then
        ' 生応答が "#ERR:" で始まっていても【昇格させない】(帯域外規約)。
        ok = True
        CallStep = rawBody
    Else
        ok = False
        CallStep = "#ERR:" & errCode & ":" & errMsg
    End If

    StageRun stepName, playId, route, modelUsed, latencyMs, _
             Len(systemPrompt) + Len(userPrompt), Len(rawBody), _
             BuildRunDetail(errCode, detailText)
End Function

' ==============================================
' CallChat - 自由対話(PL-04)専用。リボンの会話継続引数 prevU/prevA を使う
'            唯一の関数(14章§6)。JSONスキーマは使わない。
' ----------------------------------------------
'   histU/histA: 「新しい順」に GW_HIST_SEP 連結した履歴。config
'                sparring_max_turns 往復を超える分はここで切り捨てる(16章E-44)。
'   ok         : CallStepと同格の帯域外規約(自由対話は最も偽装しやすい経路)。
'   errCode    : ok=False のときだけ E02xx を帯域外で返す(E-44の「往復数を
'                減らして再開」分岐は E0204 で行う)。ok=True のときは ""。
'   PII走査(16章E-05/E-31)は【呼び出し側の責務】。modPii は app層であり
'   core層から参照できない(R1)ため、送信直前の走査は app層で行うこと。
'   run_log: 自由対話には検証段が無いので本関数だけは自分で1行書く。
' ==============================================
Public Function CallChat(ByVal caseId As String, ByVal systemPrompt As String, _
                         ByVal userMsg As String, ByVal histU As String, ByVal histA As String, _
                         ByRef ok As Boolean, Optional ByRef errCode As String = "", _
                         Optional ByRef latencyMs As Long = 0) As String
    Dim t0 As Double
    Dim route As String
    Dim rawBody As String
    Dim errMsg As String
    Dim modelUsed As String
    Dim maxTurns As Long
    Dim prevU As String
    Dim prevA As String
    Dim rec As TRunLogRec

    ok = False
    errCode = ""
    latencyMs = 0
    t0 = Timer

    On Error GoTo Failed

    FlushPendingRun

    maxTurns = modConfig.GetLong("sparring_max_turns", 12)
    prevU = TrimHistoryPairs(histU, maxTurns)
    prevA = TrimHistoryPairs(histA, maxTurns)

    route = CurrentTransport()
    Select Case route
        Case GW_MOCK
            modelUsed = GW_MOCK
            rawBody = MockChat(errCode, errMsg)
        Case GW_DIRECT
            rawBody = DirectStep("sp", systemPrompt, userMsg, "", _
                                 modelUsed, errCode, errMsg)
        Case Else
            rawBody = RibbonChat(systemPrompt, userMsg, prevU, prevA, _
                                 modelUsed, errCode, errMsg)
    End Select
    GoTo Finish

Failed:
    errCode = "E0202"
    errMsg = Err.Description
    modLog.LogError errCode, "modGatewayRPN.CallChat", _
                    "transport=" & route, Err.Number
    Resume Finish

Finish:
    latencyMs = CLng(modUtil.ElapsedMsSince(t0))
    If LenB(errCode) = 0 Then
        ok = True
        CallChat = rawBody
    Else
        ok = False
        CallChat = "#ERR:" & errCode & ":" & errMsg
    End If

    rec.case_id = caseId
    rec.round_no = ""
    rec.stepName = "sp"
    rec.play = "PL-04"
    rec.transport = route
    rec.model = modelUsed
    rec.latency_ms = latencyMs
    rec.input_chars = Len(systemPrompt) + Len(userMsg) + Len(prevU) + Len(prevA)
    rec.output_chars = Len(rawBody)
    rec.injected_kb_ids = mCtxKbIds
    If ok Then rec.validate_result = "ok" Else rec.validate_result = "failed"
    rec.detail = BuildRunDetail(errCode, "turns=" & CStr(maxTurns))
    rec.operator = mCtxOperator
    modLog.LogRun rec
End Function

' ==============================================
' RibbonAvailable - リボンアドインの検出(14章§2・裁定D2)
' ----------------------------------------------
'   公式作法: Application.AddIns をループし、config ribbon_addin_name の
'   部分一致 + Installed で判定する(API呼び出し不要・即時)。結果はセッション
'   キャッシュ。AddIns へアクセスできない環境(LibreOffice等)は「検出失敗」で
'   False を返すが E0201 のエラー扱いにはしない(誤ブロック防止。E0201 を
'   記録するのは実際に呼び出す側)。
' ==============================================
Public Function RibbonAvailable() As Boolean
    Dim addinName As String
    Dim ai As Object

    If mRibbonChecked Then
        RibbonAvailable = mRibbonAvailable
        Exit Function
    End If
    mRibbonChecked = True
    mRibbonAvailable = False

    ' 既定値は 13章§2.3 の ribbon_addin_name と同値(社内AIリボンのアドイン名で
    ' あって本製品の名前ではない)。空を明示設定したときだけ検出を行わない。
    addinName = modConfig.GetStr("ribbon_addin_name", "リボンちゃん")
    If LenB(addinName) = 0 Then
        RibbonAvailable = False
        Exit Function
    End If

    On Error GoTo DetectFail
    For Each ai In Application.AddIns
        If InStr(ai.Name, addinName) > 0 Then
            If ai.Installed Then
                mRibbonAvailable = True
                Exit For
            End If
        End If
    Next ai
    RibbonAvailable = mRibbonAvailable
    Exit Function

DetectFail:
    modLog.LogUsage "ribbon_detect_fail", "", "AddIns走査に失敗: " & Err.Description
    mRibbonAvailable = False
    RibbonAvailable = False
End Function

' ==============================================
' RunLimitCheck - リボン公式 LimitCheck() の唯一の呼び出し口(裁定D3)
'   戻り値: True=続行不可(利用期限切れ等) / False=続行可
' ----------------------------------------------
'   起動時(12章§2.1 modBoot手順(7))と実行時の案内に使う。True でも起動は
'   止めない。ribbon経路以外 / config limit_check=FALSE / リボン未検出 /
'   Application.Run 失敗(古いリボン)は、すべて False(続行可)へ倒す穏当運用
'   (誤ブロック防止。エラー扱いにせず usage_log への情報記録に留める)。
' ==============================================
Public Function RunLimitCheck() As Boolean
    Dim res As Variant

    RunLimitCheck = False

    If CurrentTransport() <> GW_RIBBON Then Exit Function
    If Not modConfig.GetBool("limit_check", True) Then Exit Function
    If Not RibbonAvailable() Then Exit Function

    On Error GoTo CheckFail
    res = Application.Run("LimitCheck")
    RunLimitCheck = CBool(res)
    Exit Function

CheckFail:
    modLog.LogUsage "limit_check_skip", "", "LimitCheck呼び出しに失敗: " & Err.Description
    RunLimitCheck = False
End Function

' ==============================================
' run_log の受け渡し(14章§1「1行=1 LLM呼び出し」を1行に保つための仕掛け)
' ----------------------------------------------
'   validate_result は【検証が終わるまで確定しない】が、transport / model /
'   latency_ms は gateway しか知らない。1呼び出し=1行に保つには受け渡しが要る:
'     (1) 呼び出し側が SetRunContext で case_id / round_no / 注入ID等を預ける
'     (2) CallStep は「gatewayが知る列」を埋めた行を保留する(StageRun)
'     (3) 検証を終えた呼び出し側が TakeLastRun で受け取り、validate_result を
'         埋めて modLog.LogRun へ1行だけ書く
'     (4) 誰も回収しないまま次の呼び出しが来たら FlushPendingRun が
'         validate_result 空のまま書き出す(行の取りこぼしを作らない)
'   CallChat だけは検証段が無いので自分で書く。
' ==============================================
Public Sub SetRunContext(ByVal caseId As String, ByVal roundNo As String, _
                         ByVal caseType As String, ByVal injectedKbIds As String, _
                         ByVal operatorName As String)
    mCtxCaseId = caseId
    mCtxRoundNo = roundNo
    mCtxCaseType = caseType
    mCtxKbIds = injectedKbIds
    mCtxOperator = operatorName
End Sub

Public Sub ClearRunContext()
    mCtxCaseId = ""
    mCtxRoundNo = ""
    mCtxCaseType = ""
    mCtxKbIds = ""
    mCtxOperator = ""
End Sub

' 保留中の行を受け取る(受け取ると保留は解除)。戻り値 False=保留なし。
Public Function TakeLastRun(ByRef rec As TRunLogRec) As Boolean
    If Not mHasPendingRun Then
        TakeLastRun = False
        Exit Function
    End If
    rec = mPendingRun
    mHasPendingRun = False
    TakeLastRun = True
End Function

' 保留中の行があれば validate_result 空のまま書き出す。
Public Sub FlushPendingRun()
    If Not mHasPendingRun Then Exit Sub
    mHasPendingRun = False
    modLog.LogRun mPendingRun
End Sub

' ==============================================
' 純ロジック(Excel非依存。LibreOffice実行テストで直接叩ける)
' ==============================================

' 経路決定(14章§1)。mock_llm=TRUE が最優先。llm_transport の想定外の値は
' 既定の ribbon へ倒す(direct へ勝手に流れないのが安全側)。
Public Function ResolveTransport(ByVal mockLlm As Boolean, ByVal transportCfg As String) As String
    Dim v As String

    If mockLlm Then
        ResolveTransport = GW_MOCK
        Exit Function
    End If

    v = LCase$(Trim$(transportCfg))
    Select Case v
        Case GW_RIBBON, GW_DIRECT, GW_MOCK
            ResolveTransport = v
        Case Else
            ResolveTransport = GW_RIBBON
    End Select
End Function

' toolN の組立(14章§2・裁定D1)。接頭辞は config app_tool_prefix 由来。core層に
' 製品名を焼かないため接頭辞が空でも動く(stepNameだけを送る)。
Public Function BuildToolName(ByVal toolPrefix As String, ByVal stepName As String) As String
    BuildToolName = Trim$(toolPrefix) & Trim$(stepName)
End Function

' 待ち時間の解決(秒)。config が 0 や負値・異常値でも呼び出しを壊さない。
Public Function ResolveWaitSec(ByVal cfgWaitSec As Long) As Long
    If cfgWaitSec <= 0 Then
        ResolveWaitSec = GW_WAIT_DEFAULT
        Exit Function
    End If
    ResolveWaitSec = modUtil.ClampLong(cfgWaitSec, GW_WAIT_MIN, GW_WAIT_MAX)
End Function

' MaxTokens の解決。Step別上書き(s1_max_tokens 等)を優先し、無ければ全体値。
' 0 は「リボン側の既定に従う」を意味するのでそのまま通す。
Public Function ResolveMaxTokens(ByVal stepMaxTokens As Long, ByVal globalMaxTokens As Long) As Long
    If stepMaxTokens > 0 Then
        ResolveMaxTokens = stepMaxTokens
    ElseIf globalMaxTokens > 0 Then
        ResolveMaxTokens = globalMaxTokens
    Else
        ResolveMaxTokens = 0
    End If
End Function

' 応答が「利用上限の定型拒否文」らしいか(16章E-15)。定型拒否文は短いので短い
' 応答に限って判定する(約款文言は「上限」「回数」を普通に含み、長文の正当な
' 応答が誤爆していた実機事故がある)。JSONらしい応答も除外する。
Public Function LooksLikeLimitError(ByVal response As String) As Boolean
    Dim s As String

    If Len(response) > GW_LIMIT_PROBE_MAX Then Exit Function
    If LooksLikeJsonBody(response) Then Exit Function

    s = LCase$(response)
    LooksLikeLimitError = (InStr(s, "上限") > 0) Or (InStr(s, "limit") > 0) Or _
                          (InStr(s, "回数") > 0) Or (InStr(s, "rate") > 0) Or _
                          (InStr(s, "quota") > 0)
End Function

' 応答判定。戻り値 "" = 正常 / "E0202" = 空応答 / "E0204" = 利用上限。
' 【重要】ここでは "#ERR:" プレフィクスを一切見ない。LLMが "#ERR:E0201:..." で
' 始まる本文を返しても正常応答として扱う(14章§4 mock_fault=fake_err・§6)。
Public Function ClassifyResponse(ByVal response As String) As String
    If LenB(Trim$(response)) = 0 Then
        ClassifyResponse = "E0202"
        Exit Function
    End If
    If LooksLikeLimitError(response) Then
        ClassifyResponse = "E0204"
        Exit Function
    End If
    ClassifyResponse = ""
End Function

' mock のバリアント決定(15章§8.1)。乱数・現在時刻を使わない決定的な規則。
'   s1 / s2 / s2r : 案件区分(new / renewal)で分ける
'   s2c / s3c     : 呼び出し回数の偶奇で hit(指摘あり)と clean(0件)を交替させ、
'                   改訂パスと改訂スキップの両経路を1回のE2Eで通す
'   その他        : common(new/renewal どちらの文脈でも合格する共通応答)
Public Function ResolveMockVariant(ByVal stepName As String, ByVal caseType As String, _
                                   ByVal roundIndex As Long) As String
    Select Case LCase$(Trim$(stepName))
        Case "s1", "s2", "s2r"
            If LCase$(Trim$(caseType)) = "renewal" Then
                ResolveMockVariant = "renewal"
            Else
                ResolveMockVariant = "new"
            End If
        Case "s2c", "s3c"
            If (roundIndex Mod 2) = 1 Then
                ResolveMockVariant = "hit"
            Else
                ResolveMockVariant = "clean"
            End If
        Case Else
            ResolveMockVariant = "common"
    End Select
End Function

' 会話履歴の切詰め(16章E-44)。履歴は「新しい順」に GW_HIST_SEP 連結されて
' 来るので先頭から maxTurns 個だけ残す。
Public Function TrimHistoryPairs(ByVal hist As String, ByVal maxTurns As Long) As String
    Dim parts As Variant
    Dim i As Long
    Dim n As Long
    Dim acc As String

    If LenB(hist) = 0 Then Exit Function
    If maxTurns <= 0 Then Exit Function

    parts = Split(hist, GW_HIST_SEP)
    n = UBound(parts) - LBound(parts) + 1
    If n <= maxTurns Then
        TrimHistoryPairs = hist
        Exit Function
    End If

    For i = 0 To maxTurns - 1
        If i > 0 Then acc = acc & GW_HIST_SEP
        acc = acc & parts(LBound(parts) + i)
    Next i
    TrimHistoryPairs = acc
End Function

' エラーコードに対応する人間向け説明(16章§1)。ok=False のときだけ使う文面で
' あって、成否の判定材料ではない。
Public Function ErrMessageFor(ByVal errCode As String) As String
    Select Case errCode
        Case "E0201"
            ErrMessageFor = "AIリボンが見つかりません"
        Case "E0202"
            ErrMessageFor = "応答が空でした。時間をおいて再実行してください"
        Case "E0203"
            ErrMessageFor = "APIが混雑しています。時間をおいて再実行してください"
        Case "E0204"
            ErrMessageFor = "本日のAI利用枠の上限です"
        Case "E0205"
            ErrMessageFor = "direct経路は開発者専用です"
        Case "E0206"
            ErrMessageFor = "応答が拒否または途中終了しました"
        Case Else
            ErrMessageFor = "呼び出しに失敗しました"
    End Select
End Function

' run_log の detail(400字上限・本文非記録=NFR-S3)。メタだけを連結する。
Public Function BuildRunDetail(ByVal errCode As String, ByVal extra As String) As String
    Dim acc As String

    If LenB(errCode) > 0 Then acc = "err=" & errCode
    If LenB(extra) > 0 Then
        If LenB(acc) > 0 Then acc = acc & ";"
        acc = acc & extra
    End If
    BuildRunDetail = modUtil.SafeLeft(acc, 400)
End Function

' ==============================================
' 内部ヘルパー
' ==============================================

' 現在の経路(config駆動)。
Private Function CurrentTransport() As String
    CurrentTransport = ResolveTransport(modConfig.GetBool("mock_llm", False), _
                                        modConfig.GetStr("llm_transport", GW_RIBBON))
End Function

' JSONらしい本文か(上限誤爆の除外用)。先頭が波括弧・角括弧・コードフェンス。
Private Function LooksLikeJsonBody(ByVal response As String) As Boolean
    Dim t As String
    Dim c As String

    t = Trim$(response)
    If LenB(t) = 0 Then Exit Function
    c = Left$(t, 1)
    LooksLikeJsonBody = (c = "{") Or (c = "[") Or (Left$(t, 3) = "```")
End Function

' 推論調整の解決。Step別上書き(s2_effort 等)を優先し、無ければ全体値。config
' reasoning_tuning=FALSE のときは空文字を送る(エスケープハッチ)。
Private Function ResolveTuning(ByVal stepName As String, ByVal kind As String) As String
    Dim stepKey As String

    If Not modConfig.GetBool("reasoning_tuning", True) Then Exit Function

    stepKey = stepName & "_" & kind
    If modConfig.HasKey(stepKey) Then
        ResolveTuning = modConfig.GetStr(stepKey, "")
        Exit Function
    End If

    If kind = "effort" Then
        ResolveTuning = modConfig.GetStr("reasoning_effort", "medium")
    Else
        ResolveTuning = modConfig.GetStr("reasoning_verbosity", "low")
    End If
End Function

' ----------------------------------------------
' RibbonStep - 主経路(本番)。14章§2の12引数呼び出し。
' ----------------------------------------------
'   引数値はすべて modConfig 参照(10章 NFR-M3。リテラル直書き禁止)。
'   位置引数の意味(確定台帳 RIBBON_API_CONFIRMED.md §1 #1・裁定D1):
'     1 Text  2 roleSystem  3 Temperature(Double)  4 MaxTokens(Long。0=既定)
'     5 Wait(秒)  6 optModel  7 prevU  8 prevA(Stepは各回独立のため "")
'     9 toolN(管理側ログ識別。config app_tool_prefix + stepName)
'    10 reasoning_effort  11 verbosity
'   リボン未検出は E0201 で停止し direct へ自動フォールバックしない(16章E-14)。
'
'   16章E-50: この Application.Run は最大 llm_wait_sec 秒 VBAをブロックする。
'   進捗の確定表示(modUIProgress.SetStage)は【呼び出し側がこの前に】済ませて
'   おく規約(core層から ui層は参照できない=R1)。ここでは書き切った表示を画面へ
'   反映させるため呼出直前に DoEvents を1回挟む。ゴースト化抑止は ui層の担当。
' ----------------------------------------------
Private Function RibbonStep(ByVal stepName As String, ByVal systemPrompt As String, _
                            ByVal userPrompt As String, ByRef modelUsed As String, _
                            ByRef errCode As String, ByRef errMsg As String) As String
    Dim temperature As Double
    Dim maxTok As Long
    Dim waitSec As Long
    Dim toolN As String
    Dim effort As String
    Dim verbosity As String
    Dim res As Variant
    Dim s As String

    modelUsed = modConfig.GetStr("recommended_model", "")

    If Not RibbonAvailable() Then
        errCode = "E0201"
        errMsg = ErrMessageFor(errCode)
        modLog.LogError errCode, "modGatewayRPN.CallStep", "step=" & stepName
        Exit Function
    End If

    temperature = modConfig.GetDouble("temperature", 0.3)
    maxTok = ResolveMaxTokens(modConfig.GetLong(stepName & "_max_tokens", 0), _
                              modConfig.GetLong("llm_max_tokens", 0))
    waitSec = ResolveWaitSec(modConfig.GetLong("llm_wait_sec", 1200))
    toolN = BuildToolName(modConfig.GetStr("app_tool_prefix", ""), stepName)
    effort = ResolveTuning(stepName, "effort")
    verbosity = ResolveTuning(stepName, "verbosity")

    On Error GoTo RunFailed
    DoEvents
    res = Application.Run("ChatGPT", userPrompt, systemPrompt, temperature, maxTok, _
                          waitSec, modelUsed, "", "", toolN, effort, verbosity)
    s = CStr(res)
    GoTo Classify

RunFailed:
    errCode = "E0202"
    errMsg = Err.Description
    modLog.LogError errCode, "modGatewayRPN.CallStep", _
                    "step=" & stepName & " tool=" & toolN, Err.Number
    Resume ExitPoint

Classify:
    errCode = ClassifyResponse(s)
    If LenB(errCode) > 0 Then
        errMsg = ErrMessageFor(errCode)
        modLog.LogError errCode, "modGatewayRPN.CallStep", _
                        "step=" & stepName & " len=" & CStr(Len(s))
        Exit Function
    End If
    RibbonStep = s

ExitPoint:
End Function

' RibbonChat - 自由対話(PL-04)のリボン呼び出し。prevU/prevA を使う点だけが
'   RibbonStep と異なる。toolN の stepName は "sp" 固定。
Private Function RibbonChat(ByVal systemPrompt As String, ByVal userMsg As String, _
                            ByVal prevU As String, ByVal prevA As String, _
                            ByRef modelUsed As String, ByRef errCode As String, _
                            ByRef errMsg As String) As String
    Dim temperature As Double
    Dim maxTok As Long
    Dim waitSec As Long
    Dim toolN As String
    Dim effort As String
    Dim verbosity As String
    Dim res As Variant
    Dim s As String

    modelUsed = modConfig.GetStr("recommended_model", "")

    If Not RibbonAvailable() Then
        errCode = "E0201"
        errMsg = ErrMessageFor(errCode)
        modLog.LogError errCode, "modGatewayRPN.CallChat", "step=sp"
        Exit Function
    End If

    temperature = modConfig.GetDouble("temperature", 0.3)
    maxTok = ResolveMaxTokens(modConfig.GetLong("sp_max_tokens", 0), _
                              modConfig.GetLong("llm_max_tokens", 0))
    waitSec = ResolveWaitSec(modConfig.GetLong("llm_wait_sec", 1200))
    toolN = BuildToolName(modConfig.GetStr("app_tool_prefix", ""), "sp")
    effort = ResolveTuning("sp", "effort")
    verbosity = ResolveTuning("sp", "verbosity")

    On Error GoTo RunFailed
    DoEvents
    res = Application.Run("ChatGPT", userMsg, systemPrompt, temperature, maxTok, _
                          waitSec, modelUsed, prevU, prevA, toolN, effort, verbosity)
    s = CStr(res)
    GoTo Classify

RunFailed:
    errCode = "E0202"
    errMsg = Err.Description
    modLog.LogError errCode, "modGatewayRPN.CallChat", "step=sp", Err.Number
    Resume ExitPoint

Classify:
    errCode = ClassifyResponse(s)
    If LenB(errCode) > 0 Then
        errMsg = ErrMessageFor(errCode)
        modLog.LogError errCode, "modGatewayRPN.CallChat", "step=sp len=" & CStr(Len(s))
        Exit Function
    End If
    RibbonChat = s

ExitPoint:
End Function

' ----------------------------------------------
' DirectStep - 開発・検証用(14章§3)。HTTP・strict・リトライ・キー読込の実体は
'   modGatewayDirect(T-13)が持ち、本モジュールは経路を選ぶだけ。
' ----------------------------------------------
'   T-13 への呼び出し契約(本モジュールが前提とする公開シグネチャ):
'     CallDirect(stepName, systemPrompt, userPrompt, schemaJson, _
'                ByRef modelUsed, ByRef errCode, ByRef errMsg) As String
'   errCode は "" = 成功 / E0203(429・5xx リトライ尽き)/ E0205(キー無し)/
'   E0206(refusal・finish_reason が stop 以外)/ E0202(その他)。err_log 記録と
'   HTTPステータスの保持は modGatewayDirect 側の責務。
' ----------------------------------------------
Private Function DirectStep(ByVal stepName As String, ByVal systemPrompt As String, _
                            ByVal userPrompt As String, ByVal schemaJson As String, _
                            ByRef modelUsed As String, ByRef errCode As String, _
                            ByRef errMsg As String) As String
    modelUsed = modConfig.GetStr("direct_model", "")
    DirectStep = modGatewayDirect.CallDirect(stepName, systemPrompt, userPrompt, _
                                             schemaJson, modelUsed, errCode, errMsg)
    If LenB(errCode) > 0 Then
        If LenB(errMsg) = 0 Then errMsg = ErrMessageFor(errCode)
        DirectStep = ""
    End If
End Function

' ----------------------------------------------
' MockStep - 本体内mockトランスポート(14章§4(a))。呼ぶ相手は modMockLlm だけ。
' ----------------------------------------------
'   T-14 への呼び出し契約(本モジュールが前提とする公開シグネチャ):
'     MockResponse(stepName, variantName, fault) As String
'   variantName は ResolveMockVariant の戻り値(new / renewal / hit / clean /
'   common)。fault は config mock_fault の値(空なら正常応答)。
'   障害注入の応答も【正常な戻り値】として扱い、空応答=E0202・上限文字列=
'   E0204 の判定は ribbon経路と同じ ClassifyResponse に通す。fake_err
'   (本文が "#ERR:E0201:..." で始まる正常JSON)は判定に使わないので ok=True の
'   まま素通しする(14章§4・§6)。
' ----------------------------------------------
Private Function MockStep(ByVal stepName As String, ByRef modelUsed As String, _
                          ByRef detailText As String, ByRef errCode As String, _
                          ByRef errMsg As String) As String
    Dim fault As String
    Dim variantName As String
    Dim s As String

    modelUsed = GW_MOCK
    fault = Trim$(modConfig.GetStr("mock_fault", ""))

    variantName = ResolveMockVariant(stepName, mCtxCaseType, NextMockRound(stepName))
    detailText = "variant=" & variantName
    If LenB(fault) > 0 Then detailText = detailText & ";fault=" & fault

    s = modMockLlm.MockResponse(stepName, variantName, fault)

    errCode = ClassifyResponse(s)
    If LenB(errCode) > 0 Then
        errMsg = ErrMessageFor(errCode)
        modLog.LogError errCode, "modGatewayRPN.CallStep", _
                        "step=" & stepName & " " & detailText
        Exit Function
    End If
    MockStep = s
End Function

' 批判stepの呼び出し回数を1つ進めて返す(それ以外のstepは 0 で無関係)。
Private Function NextMockRound(ByVal stepName As String) As Long
    Select Case LCase$(Trim$(stepName))
        Case "s2c"
            mMockRoundS2c = mMockRoundS2c + 1
            NextMockRound = mMockRoundS2c
        Case "s3c"
            mMockRoundS3c = mMockRoundS3c + 1
            NextMockRound = mMockRoundS3c
        Case Else
            NextMockRound = 0
    End Select
End Function

' 自由対話のmock。step は "sp"、バリアントは common 固定。
Private Function MockChat(ByRef errCode As String, ByRef errMsg As String) As String
    Dim fault As String
    Dim s As String

    fault = Trim$(modConfig.GetStr("mock_fault", ""))
    s = modMockLlm.MockResponse("sp", "common", fault)

    errCode = ClassifyResponse(s)
    If LenB(errCode) > 0 Then
        errMsg = ErrMessageFor(errCode)
        modLog.LogError errCode, "modGatewayRPN.CallChat", "step=sp"
        Exit Function
    End If
    MockChat = s
End Function

' StageRun - run_log の行を組み立てて保留する(書き出しはしない)。
'   validate_result は呼び出し側が検証後に埋める列なのでここでは空のまま。
Private Sub StageRun(ByVal stepName As String, ByVal playId As String, _
                     ByVal route As String, ByVal modelUsed As String, _
                     ByVal latencyMs As Long, ByVal inChars As Long, _
                     ByVal outChars As Long, ByVal detailText As String)
    mPendingRun.run_at = ""
    mPendingRun.case_id = mCtxCaseId
    mPendingRun.round_no = mCtxRoundNo
    mPendingRun.stepName = stepName
    mPendingRun.play = playId
    mPendingRun.transport = route
    mPendingRun.model = modelUsed
    mPendingRun.latency_ms = latencyMs
    mPendingRun.input_chars = inChars
    mPendingRun.output_chars = outChars
    mPendingRun.injected_kb_ids = mCtxKbIds
    mPendingRun.validate_result = ""
    mPendingRun.detail = detailText
    mPendingRun.operator = mCtxOperator
    mHasPendingRun = True
End Sub
