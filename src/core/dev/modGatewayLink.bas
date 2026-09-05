Attribute VB_Name = "modGatewayLink"
Option Explicit

' ============================================================================
' modGatewayLink(dev版) - direct経路への接続点。裁定書30 裁定1(b)。
' ----------------------------------------------------------------------------
' 開発ビルド(--dev)だけがこのソースを使う(build/modules.json の dev_src)。
' 配布ビルド(--prod)は src/core/modGatewayLink.bas(E0209 を返すだけの版)を
' 使い、modGatewayDirect ごと配布物から落とす。
'
' 本ファイルの責務は転送1本だけである。HTTP・strict・リトライ・キー読込の実体は
' modGatewayDirect(T-13)が持ち、errCode / errMsg の解釈は modGatewayRPN が持つ。
' ここに判断を足さないこと(足すと prod 版と dev 版でふるまいが割れる)。
' ============================================================================

Public Function CallDirect(ByVal stepName As String, ByVal systemPrompt As String, _
                           ByVal userPrompt As String, ByVal schemaJson As String, _
                           ByRef modelUsed As String, ByRef errCode As String, _
                           ByRef errMsg As String) As String
    CallDirect = modGatewayDirect.CallDirect(stepName, systemPrompt, userPrompt, _
                                             schemaJson, modelUsed, errCode, errMsg)
End Function
