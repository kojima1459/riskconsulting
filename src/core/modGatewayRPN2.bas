Attribute VB_Name = "modGatewayRPN2"
Option Explicit

' ============================================================================
' modGatewayRPN2 - modGatewayRPN の純関数の受け皿(core層・T-63・裁定書34)
' ----------------------------------------------------------------------------
' なぜ2本目か(30,000字契約・12章§2):
'   W12-A(裁定書34 §1.2)で modGatewayRPN.CallChat に「どのStepとして呼ぶか」
'   (壁打ち sp / 案件チャット ch)の引数を通す必要が出たが、modGatewayRPN は
'   29,952字で残り48字しかなかった。**コメントを削って字数を稼ぐことはしない**
'   規約(CLAUDE.md)に従い、Excel にもシートにも config にも触らない
'   「引数だけで答えが決まる4本」をこちらへ移して空きを作った。
'
' 移した4本(挙動は1文字も変えていない。純層テストの参照先だけが変わる):
'   BuildToolName    リボンへ渡す toolN の組立(14章§2・裁定D1)
'   ResolveWaitSec   待ち時間(秒)の解決とクランプ
'   ResolveMaxTokens Step別上書き -> 全体値 -> 0(リボン既定に従う)
'   TrimHistoryPairs 会話履歴の切詰め(16章E-44)
'
' ここに置かないもの: config を読む関数(ResolveTuning / CurrentTransport)と、
'   応答文字列の判定(ClassifyResponse / DecideOk 等)。前者は純関数ではなく、
'   後者は modGatewayRPN 本体の責務(経路の成否判定)から切り離すと読みにくい。
'
' R1/R4(12章§4): core層。Excelトークンを1つも持たない(層(a)から直接叩ける)。
' ============================================================================

' 自由対話(PL-04)の履歴区切り。**値源は modGatewayRPN.GW_HIST_SEP の1箇所**で
' あり、ここでは参照するだけ(同じ値を二重に書かない)。
' 待ち時間の安全域(秒)。既定値は 13章§2.3 の llm_wait_sec と同値で、config が
' 読めないときの最後の砦(通常はconfigの値が使われる。NFR-M3)。
Private Const GW2_WAIT_MIN As Long = 30
Private Const GW2_WAIT_MAX As Long = 7200
Private Const GW2_WAIT_DEFAULT As Long = 1200

' toolN の組立(14章§2・裁定D1)。接頭辞は config app_tool_prefix 由来。core層に
' 製品名を焼かないため接頭辞が空でも動く(stepNameだけを送る)。
Public Function BuildToolName(ByVal toolPrefix As String, ByVal stepName As String) As String
    BuildToolName = Trim$(toolPrefix) & Trim$(stepName)
End Function

' 待ち時間の解決(秒)。config が 0 や負値・異常値でも呼び出しを壊さない。
Public Function ResolveWaitSec(ByVal cfgWaitSec As Long) As Long
    If cfgWaitSec <= 0 Then
        ResolveWaitSec = GW2_WAIT_DEFAULT
        Exit Function
    End If
    ResolveWaitSec = modUtil.ClampLong(cfgWaitSec, GW2_WAIT_MIN, GW2_WAIT_MAX)
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

' 会話履歴の切詰め(16章E-44)。履歴は「新しい順」に modGatewayRPN.GW_HIST_SEP
' 連結されて来るので先頭から maxTurns 個だけ残す。
Public Function TrimHistoryPairs(ByVal hist As String, ByVal maxTurns As Long) As String
    Dim parts As Variant
    Dim i As Long
    Dim n As Long
    Dim acc As String

    If LenB(hist) = 0 Then Exit Function
    If maxTurns <= 0 Then Exit Function

    parts = Split(hist, modGatewayRPN.GW_HIST_SEP)
    n = UBound(parts) - LBound(parts) + 1
    If n <= maxTurns Then
        TrimHistoryPairs = hist
        Exit Function
    End If

    For i = 0 To maxTurns - 1
        If i > 0 Then acc = acc & modGatewayRPN.GW_HIST_SEP
        acc = acc & parts(LBound(parts) + i)
    Next i
    TrimHistoryPairs = acc
End Function

' リボン定型失敗文(E0203/E0204/E0207)の先頭80字。err_log detail に残し原因を
' 追えるようにする(裁定書36)。改行/タブは半角空白へ均してから
' modUtilText.SanitizeInput(私用領域除去・■■■置換)を通し、最後にLeft$で
' 80字へ切る(NFR-S3: 本文ではないが必ず切る)。空応答なら空文字を返す。
Public Function RibbonHead(ByVal response As String) As String
    Dim t As String
    t = Trim$(response)
    If LenB(t) = 0 Then Exit Function
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    t = modUtilText.SanitizeInput(t)
    If Len(t) > 80 Then t = Left$(t, 80)
    RibbonHead = t
End Function
