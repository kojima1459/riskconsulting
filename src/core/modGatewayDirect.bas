Attribute VB_Name = "modGatewayDirect"
Option Explicit

' ==============================================
' modGatewayDirect - direct経路(開発・検証用)の実体(14章§3・17章T-13)
' ----------------------------------------------
' 役割: modGatewayRPN.DirectStep から呼ばれる唯一の相手。chat/completions への
'   HTTP POST・response_format=json_schema(strict)・指数バックオフ・キー読込を
'   ここへ閉じ込める。経路選択やmodel/errCode/errMsgの解釈はmodGatewayRPNの
'   責務(本モジュールは知らない)。
'
' 公開契約(modGatewayRPN.DirectStep のコメントに書かれた前提。14章§6):
'   CallDirect(stepName, systemPrompt, userPrompt, schemaJson, _
'              ByRef modelUsed, ByRef errCode, ByRef errMsg) As String
'   errCode: "" = 成功 / E0203(429・5xxのリトライ尽き) / E0205(キー無し) /
'            E0206(refusal・finish_reasonがstop以外) / E0202(その他=408・
'            タイムアウト・その他4xx・通信例外)。
'   戻り値: 成功時は choices[0].message.content(生テキスト)。失敗時は ""。
'   err_log記録とHTTPステータスの保持は本モジュールの責務(modGatewayRPN
'   コメントの明記どおり)。
'
' NFR-S2(16章): キーは %APPDATA%\RPN\api_key.txt の1行目のみ。ブック・config・
'   ログ・画面・リポジトリ・配布物のどこにも書かない。難読化して埋め込む方式
'   (PoCのOBF1)は禁止。本モジュールはファイルから読んだ値をそのまま
'   Authorizationヘッダへ渡すだけで、値を変数以外(セル・ログ・戻り値)へ
'   一切書かない。
'
' R4: modGatewayDirectはExcelトークン許可モジュール9本(12章§4)に含まれない。
'   Worksheets/Range/Application./ThisWorkbook/MsgBox/ActiveSheetのいずれにも
'   触れない(CreateObjectでのCOM生成はExcelトークンではないため可)。
'
' 純関数の分離方針(このタスクの指示どおり3種を分ける。いずれもExcel非依存の
'   Public関数でLibreOffice実行テストから直接叩ける):
'   (a) バックオフ計算  : ComputeBackoffMs / IsRetryableServerStatus /
'                          ShouldRetryDirect / RetryWaitMs / ClassifyDirectErrorCode
'   (b) リクエストボディ組立: BuildRequestBody / BuildChatCompletionsUrl /
'                          ResolveMaxTokensDirect / IsOSeriesModel / FormatTemperature
'   (c) キーファイルのパース: ParseApiKeyFirstLine / ExpandAppDataToken
'   HTTP実体(CreateObject("MSXML2.ServerXMLHTTP.6.0"))とファイルI/O(Open/Close)・
'   Sleepは上記の外側にあるPrivateの薄い手続きへ閉じ込め、純関数からは呼ばない。
'
' schemaJson引数について: modSchemas(15章)は未実装のため、本モジュールはJSON
'   スキーマ本文を丸ごと文字列で受け取るだけで、スキーマの妥当性検査はしない
'   (14章§3が「15章の全スキーマはstrict要件を満たす」と保証する前提に乗る)。
'   schemaJsonが空(壁打ちPL-04のCallChat経由等)のときは response_format を
'   送らない。
' ==============================================

' 429/500/502/503のバックオフ間隔(14章§3。ミリ秒)。
Private Const RETRY_BACKOFF_1_MS As Long = 2000
Private Const RETRY_BACKOFF_2_MS As Long = 4000
Private Const RETRY_BACKOFF_3_MS As Long = 8000
' 429/5xxは最大3回まで再試行(初回+3回=最大4試行)。
Private Const RETRY_MAX_ATTEMPTS_5XX As Long = 4
' 408/タイムアウトは1回だけ再試行(初回+1回=最大2試行)。
Private Const RETRY_MAX_ATTEMPTS_TIMEOUT As Long = 2
' 408/タイムアウト再試行時の待ち時間(仕様は「1回」とだけ定め間隔は明記しない
' ため、429系の1段目と同じ2秒を用いる)。
Private Const RETRY_TIMEOUT_WAIT_MS As Long = 2000

' MSXML2.ServerXMLHTTP.6.0 の setTimeouts 引数(resolve/connect固定・send/
' receiveはconfig direct_http_timeout_ms。13章§2.3既定120000)。
Private Const HTTP_RESOLVE_TIMEOUT_MS As Long = 5000
Private Const HTTP_CONNECT_TIMEOUT_MS As Long = 10000

' Sleep(kernel32)。32bit/64bit両対応(#If VBA7)。LongLongは使わない(タスク制約)。
#If VBA7 Then
    Private Declare PtrSafe Sub WinApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub WinApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

' ==============================================
' CallDirect - direct経路の唯一の入口(14章§3・§6)。
' ==============================================
Public Function CallDirect(ByVal stepName As String, ByVal systemPrompt As String, _
                           ByVal userPrompt As String, ByVal schemaJson As String, _
                           ByRef modelUsed As String, ByRef errCode As String, _
                           ByRef errMsg As String) As String
    errCode = vbNullString
    errMsg = vbNullString
    CallDirect = vbNullString

    Dim apiKey As String
    apiKey = LoadApiKey()
    If LenB(apiKey) = 0 Then
        errCode = "E0205"
        modLog.LogError errCode, "modGatewayDirect.CallDirect", _
                        "step=" & stepName & " api_key_missing"
        Exit Function
    End If

    Dim model As String
    model = Trim$(modelUsed)
    If LenB(model) = 0 Then model = Trim$(modConfig.GetStr("direct_model", vbNullString))
    modelUsed = model

    Dim apiUrl As String
    apiUrl = BuildChatCompletionsUrl(modConfig.GetStr("direct_api_base", vbNullString))

    Dim temperature As Double
    temperature = modConfig.GetDouble("temperature", 0.3)

    Dim maxTokens As Long
    maxTokens = ResolveMaxTokensDirect(modConfig.GetLong(stepName & "_max_tokens", 0), _
                                       modConfig.GetLong("llm_max_tokens", 0))

    Dim timeoutMs As Long
    timeoutMs = modConfig.GetLong("direct_http_timeout_ms", 120000)
    If timeoutMs < 1000 Then timeoutMs = 1000

    Dim body As String
    body = BuildRequestBody(stepName, model, systemPrompt, userPrompt, schemaJson, _
                            temperature, Not IsOSeriesModel(model), maxTokens)

    Dim attemptsSoFar As Long
    Dim httpStatus As Long
    Dim respText As String
    Dim wasTimeout As Boolean
    Dim comErrNum As Long
    Dim waitMs As Long

    Do
        attemptsSoFar = attemptsSoFar + 1
        SendOnce apiUrl, apiKey, body, timeoutMs, httpStatus, respText, wasTimeout, comErrNum

        If httpStatus = 200 And Not wasTimeout Then Exit Do

        If Not ShouldRetryDirect(httpStatus, wasTimeout, attemptsSoFar) Then
            errCode = ClassifyDirectErrorCode(httpStatus, wasTimeout)
            modLog.LogError errCode, "modGatewayDirect.CallDirect", _
                            "step=" & stepName & " attempts=" & CStr(attemptsSoFar), _
                            comErrNum, httpStatus
            Exit Function
        End If

        waitMs = RetryWaitMs(httpStatus, wasTimeout, attemptsSoFar)
        modLog.LogUsage "direct_retry", vbNullString, _
                       "step=" & stepName & " attempt=" & CStr(attemptsSoFar) & _
                       " status=" & CStr(httpStatus) & " wait_ms=" & CStr(waitMs)
        SleepMs waitMs
    Loop

    ' 200 OK。refusal / finish_reason の判定(14章§3。ok=Falseへ倒す唯一の理由)。
    Dim refusal As String
    refusal = modJsonLite.GetStr(respText, "refusal")
    Dim finishReason As String
    finishReason = modJsonLite.GetStr(respText, "finish_reason")
    If LenB(refusal) > 0 Or (LenB(finishReason) > 0 And finishReason <> "stop") Then
        errCode = "E0206"
        modLog.LogError errCode, "modGatewayDirect.CallDirect", _
                        "step=" & stepName & " finish_reason=" & finishReason
        Exit Function
    End If

    CallDirect = modJsonLite.GetStr(respText, "content")
End Function

' ==============================================
' (a) バックオフ計算(純関数。Excel非依存)
' ==============================================

' 429/5xx用バックオフ間隔(ms)。retryIndex=1回目の再試行前/2/3。範囲外は0。
Public Function ComputeBackoffMs(ByVal retryIndex As Long) As Long
    Select Case retryIndex
        Case 1
            ComputeBackoffMs = RETRY_BACKOFF_1_MS
        Case 2
            ComputeBackoffMs = RETRY_BACKOFF_2_MS
        Case 3
            ComputeBackoffMs = RETRY_BACKOFF_3_MS
        Case Else
            ComputeBackoffMs = 0
    End Select
End Function

' 429/500/502/503か(14章§3の対象コード)。
Public Function IsRetryableServerStatus(ByVal httpStatus As Long) As Boolean
    IsRetryableServerStatus = (httpStatus = 429) Or (httpStatus = 500) Or _
                              (httpStatus = 502) Or (httpStatus = 503)
End Function

' 次の試行を行ってよいか。attemptsSoFar=直前までに完了した試行回数(1始まり)。
'   408/タイムアウト: 1回だけ再試行(合計2試行)。
'   429/5xx         : 最大3回まで再試行(合計4試行)。
'   その他4xx等      : 再試行なし。
Public Function ShouldRetryDirect(ByVal httpStatus As Long, ByVal wasTimeout As Boolean, _
                                  ByVal attemptsSoFar As Long) As Boolean
    If wasTimeout Or httpStatus = 408 Then
        ShouldRetryDirect = (attemptsSoFar < RETRY_MAX_ATTEMPTS_TIMEOUT)
    ElseIf IsRetryableServerStatus(httpStatus) Then
        ShouldRetryDirect = (attemptsSoFar < RETRY_MAX_ATTEMPTS_5XX)
    Else
        ShouldRetryDirect = False
    End If
End Function

' 次の試行までの待ち時間(ms)。ShouldRetryDirectがTrueのときだけ意味を持つ。
Public Function RetryWaitMs(ByVal httpStatus As Long, ByVal wasTimeout As Boolean, _
                            ByVal attemptsSoFar As Long) As Long
    If wasTimeout Or httpStatus = 408 Then
        RetryWaitMs = RETRY_TIMEOUT_WAIT_MS
    Else
        RetryWaitMs = ComputeBackoffMs(attemptsSoFar)
    End If
End Function

' リトライを使い切った(または対象外だった)ときのerrCode。
'   429/5xx由来のみE0203、それ以外(408/タイムアウト/その他4xx/通信例外)は
'   E0202(modGatewayRPNコメントの表記どおり「その他」)。
Public Function ClassifyDirectErrorCode(ByVal httpStatus As Long, ByVal wasTimeout As Boolean) As String
    If (Not wasTimeout) And IsRetryableServerStatus(httpStatus) Then
        ClassifyDirectErrorCode = "E0203"
    Else
        ClassifyDirectErrorCode = "E0202"
    End If
End Function

' ==============================================
' (b) リクエストボディ組立(純関数。Excel非依存)
' ==============================================

' chat/completions のURL(direct_api_base末尾の"/"重複を吸収)。
Public Function BuildChatCompletionsUrl(ByVal apiBase As String) As String
    Dim b As String
    b = Trim$(apiBase)
    If Len(b) > 0 Then
        If Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    End If
    BuildChatCompletionsUrl = b & "/chat/completions"
End Function

' MaxTokensの解決。Step別上書き(s1_max_tokens等)を優先し、無ければ全体値。
' 0は「max_tokensキー自体を送らない」を意味するのでそのまま通す(14章§3)。
Public Function ResolveMaxTokensDirect(ByVal stepMaxTokens As Long, ByVal globalMaxTokens As Long) As Long
    If stepMaxTokens > 0 Then
        ResolveMaxTokensDirect = stepMaxTokens
    ElseIf globalMaxTokens > 0 Then
        ResolveMaxTokensDirect = globalMaxTokens
    Else
        ResolveMaxTokensDirect = 0
    End If
End Function

' o系モデル名(先頭が"o")か。o系にはtemperatureを送らない(14章§3の逐語)。
Public Function IsOSeriesModel(ByVal modelName As String) As Boolean
    Dim t As String
    t = Trim$(modelName)
    If LenB(t) = 0 Then
        IsOSeriesModel = False
    Else
        IsOSeriesModel = (LCase$(Left$(t, 1)) = "o")
    End If
End Function

' Doubleを"."小数点で文字列化する(カンマ小数点ロケールでもJSONを壊さない)。
Public Function FormatTemperature(ByVal v As Double) As String
    FormatTemperature = Replace(CStr(v), ",", ".")
End Function

' リクエストボディ本体(14章§3のJSON形。32bitメモリ制約=16章E-26のためバッファ
' 連結方式=modUtil.BufAdd/BufTextを使い、`s = s & ...`の逐次連結はしない)。
'   sendTemperature: IsOSeriesModelの否定を呼び出し側から渡す(このFunctionは
'     modelを見て自ら判定せず、渡された値のみを使う=1判断1箇所)。
'   schemaJsonが空文字(Trim後)のときは response_format を出力しない。
Public Function BuildRequestBody(ByVal stepName As String, ByVal model As String, _
                                 ByVal systemPrompt As String, ByVal userPrompt As String, _
                                 ByVal schemaJson As String, ByVal temperature As Double, _
                                 ByVal sendTemperature As Boolean, ByVal maxTokens As Long) As String
    Dim buf() As String
    Dim n As Long
    modUtil.BufInit buf, n

    modUtil.BufAdd buf, n, "{""model"":"""
    modUtil.BufAdd buf, n, modJsonLite.EscapeJsonStr(model)
    modUtil.BufAdd buf, n, """"

    If sendTemperature Then
        modUtil.BufAdd buf, n, ",""temperature"":" & FormatTemperature(temperature)
    End If

    If maxTokens > 0 Then
        modUtil.BufAdd buf, n, ",""max_tokens"":" & CStr(maxTokens)
    End If

    modUtil.BufAdd buf, n, ",""messages"":[{""role"":""system"",""content"":"""
    modUtil.BufAdd buf, n, modJsonLite.EscapeJsonStr(systemPrompt)
    modUtil.BufAdd buf, n, """},{""role"":""user"",""content"":"""
    modUtil.BufAdd buf, n, modJsonLite.EscapeJsonStr(userPrompt)
    modUtil.BufAdd buf, n, """}]"

    Dim schemaTrimmed As String
    schemaTrimmed = Trim$(schemaJson)
    If LenB(schemaTrimmed) > 0 Then
        modUtil.BufAdd buf, n, ",""response_format"":{""type"":""json_schema"",""json_schema"":{""name"":"""
        modUtil.BufAdd buf, n, modJsonLite.EscapeJsonStr(stepName)
        modUtil.BufAdd buf, n, """,""strict"":true,""schema"":"
        modUtil.BufAdd buf, n, schemaTrimmed
        modUtil.BufAdd buf, n, "}}"
    End If

    modUtil.BufAdd buf, n, "}"
    BuildRequestBody = modUtil.BufText(buf, n)
End Function

' ==============================================
' (c) キーファイルのパース(純関数。Excel非依存)
' ==============================================

' pathText中の"%APPDATA%"トークンをappDataValueへ置換する。トークンが無ければ
' そのまま返す(direct_key_pathへ絶対パスを直接設定する運用も通す)。
Public Function ExpandAppDataToken(ByVal pathText As String, ByVal appDataValue As String) As String
    If InStr(1, pathText, "%APPDATA%", vbTextCompare) = 0 Then
        ExpandAppDataToken = pathText
    Else
        ExpandAppDataToken = Replace(pathText, "%APPDATA%", appDataValue, 1, -1, vbTextCompare)
    End If
End Function

' キーファイルの生テキストから1行目のキーだけを取り出す(NFR-S2: 1行目のみが
' 正)。UTF-8 BOM・CRLF/LFいずれも吸収する。前後空白は落とす。
Public Function ParseApiKeyFirstLine(ByVal rawContent As String) As String
    Dim s As String
    s = rawContent
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF& Then s = Mid$(s, 2)
    End If

    Dim lfPos As Long
    lfPos = InStr(1, s, vbLf, vbBinaryCompare)
    Dim firstLine As String
    If lfPos > 0 Then
        firstLine = Left$(s, lfPos - 1)
    Else
        firstLine = s
    End If
    firstLine = Replace(firstLine, vbCr, vbNullString)
    ParseApiKeyFirstLine = Trim$(firstLine)
End Function

' ==============================================
' 内部ヘルパー(Excel依存・COM/ファイルI/O。純関数からは呼ばない)
' ==============================================

' キーファイルの1行目を返す。無い/読めない/空はすべて""(呼び出し側でE0205)。
' 値そのものは変数以外(セル・ログ)へ一切出さない(NFR-S2)。
Private Function LoadApiKey() As String
    LoadApiKey = vbNullString

    Dim path As String
    path = KeyFilePath()
    If LenB(path) = 0 Then Exit Function

    Dim fnum As Long
    fnum = FreeFile

    On Error GoTo Failed
    Open path For Input As #fnum

    Dim lineText As String
    If Not EOF(fnum) Then
        Line Input #fnum, lineText
    Else
        lineText = vbNullString
    End If
    Close #fnum
    On Error GoTo 0

    LoadApiKey = ParseApiKeyFirstLine(lineText)
    Exit Function

Failed:
    ' ハンドラ稼働中はOn Error Resume Nextが効かない(vba_lintの実機教訓と同じ)。
    ' Resumeで一旦ハンドラを抜けてから後始末する。
    Resume Cleanup

Cleanup:
    ' Open失敗時はfnumが未オープンのまま(Closeは無害。ファイル無し=よくある
    ' 正常系のため、ここは例外を握り潰してよい唯一の箇所)。
    On Error Resume Next
    Close #fnum
    On Error GoTo 0
    LoadApiKey = vbNullString
End Function

' キーファイルのパス。config direct_key_path(既定 %APPDATA%\RPN\api_key.txt。
' 13章§2.3)を優先し、未設定ならこの既定値を使う。%APPDATA%は実行時に
' Environ$で展開する(config側は環境変数を含む文字列のまま持ってよい)。
Private Function KeyFilePath() As String
    Dim configured As String
    configured = Trim$(modConfig.GetStr("direct_key_path", vbNullString))
    If LenB(configured) = 0 Then configured = "%APPDATA%\RPN\api_key.txt"
    KeyFilePath = ExpandAppDataToken(configured, Environ$("APPDATA"))
End Function

' HTTP実体。MSXML2.ServerXMLHTTP.6.0(setTimeoutsを持つのでNW瞬断で無限待ちに
' ならない)。成功/失敗いずれもhttpStatus/respText/wasTimeout/comErrNumへ
' 帯域外で返す(例外は外へ投げない=呼び出し側のリトライループを単純に保つ)。
Private Sub SendOnce(ByVal apiUrl As String, ByVal apiKey As String, ByVal body As String, _
                     ByVal timeoutMs As Long, ByRef httpStatus As Long, ByRef respText As String, _
                     ByRef wasTimeout As Boolean, ByRef comErrNum As Long)
    httpStatus = 0
    respText = vbNullString
    wasTimeout = False
    comErrNum = 0

    On Error GoTo Failed
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts HTTP_RESOLVE_TIMEOUT_MS, HTTP_CONNECT_TIMEOUT_MS, timeoutMs, timeoutMs
    http.Open "POST", apiUrl, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.Send body
    httpStatus = CLng(http.Status)
    respText = CStr(http.responseText)

Cleanup:
    On Error Resume Next
    Set http = Nothing
    On Error GoTo 0
    Exit Sub

Failed:
    ' 通信例外(NW瞬断・DNS失敗・接続拒否等)はすべてタイムアウト扱いへ倒す
    ' (14章§3の「408/タイムアウト=1回」の枠。HTTPステータスを持たない失敗は
    ' この枠以外に受け皿が無い)。Err.Clearの前に番号を退避する(社内プロキシの
    ' 拒否等を握り潰さないため。modGatewayDirect PoCの実機教訓と同じ理由)。
    comErrNum = Err.Number
    wasTimeout = True
    Resume Cleanup
End Sub

' バックオフ待ち。kernel32 Sleepを使う(Excel Application.Waitは秒未満を
' 扱えずR4のExcelトークン禁止にも抵触するため使わない)。
Private Sub SleepMs(ByVal ms As Long)
    If ms > 0 Then WinApiSleep ms
End Sub
