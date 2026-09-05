Attribute VB_Name = "modGatewayLink"
Option Explicit

' ============================================================================
' modGatewayLink - direct経路への接続点(配布版)。裁定書30 裁定1(b)。
' ----------------------------------------------------------------------------
' なぜ在るのか:
'   direct経路の実体 modGatewayDirect は、社内AVが重く見るHTTPのCOM生成を持つ。
'   配布物(prod)からはモジュールごと落とす(modules.json の
'   ship:false)。ところが VBA は**プロジェクト全体**をコンパイルするため、
'   modGatewayRPN が modGatewayDirect を名前で参照したままだと配布物が実機で
'   コンパイルエラーになる。そこで「呼び先を1本だけ持つ薄い接続モジュール」を
'   間に挟み、**ソースをビルドモードで差し替える**。
'     prod = 本ファイル(src/core/modGatewayLink.bas)。E0209 を返すだけ。
'     dev  = src/core/dev/modGatewayLink.bas。modGatewayDirect へ転送する。
'   どちらを使うかは build/modules.json の dev_src が唯一の値源であり、
'   Application.Run や文字列ディスパッチによる実行時の切替は**しない**
'   (実行時に名前で呼ぶ形そのものが AV の見る形であり、禁止事項でもある)。
'
' 契約(14章§6 の CallDirect と同一。呼び出し側 modGatewayRPN.DirectStep は
'   prod/dev のどちらでも同じ書き方で呼べる):
'   CallDirect(stepName, systemPrompt, userPrompt, schemaJson, _
'              ByRef modelUsed, ByRef errCode, ByRef errMsg) As String
'
' prod版のふるまい(裁定書30 裁定1(c)・司令塔の手直し):
'   config の llm_transport=direct が設定されていたら E0209 を返し、
'   **err_log にも1行残す**(16章 E-62)。**リボン経路へ黙って倒さない**
'   (設定ミスを隠すと、利用者は自分が何経路で動いているのか分からなくなる)。
'   detail は "direct_not_shipped" の一語だけで、設定値も本文も書かない
'   (NFR-S3)。
' ============================================================================

' 配布版で direct 経路が呼ばれたときのエラーコード(16章§1・14章§3)。
Public Const GL_ERR_DIRECT_OFF As String = "E0209"

' 逐語(16章§1 E0209)。原則③の2文。
Private Const GL_MSG_DIRECT_OFF As String = _
    "この配布では direct 経路は使えません。llm_transport=ribbon にしてください"

Public Function CallDirect(ByVal stepName As String, ByVal systemPrompt As String, _
                           ByVal userPrompt As String, ByVal schemaJson As String, _
                           ByRef modelUsed As String, ByRef errCode As String, _
                           ByRef errMsg As String) As String
    ' 引数は契約を合わせるために受けるだけで、配布版では1つも使わない。
    errCode = GL_ERR_DIRECT_OFF
    errMsg = GL_MSG_DIRECT_OFF
    modelUsed = vbNullString
    CallDirect = vbNullString

    ' 設定ミスの記録(16章 E-62)。detail は固定語のみ=NFR-S3。
    modLog.LogError GL_ERR_DIRECT_OFF, "modGatewayLink.CallDirect", _
                    "direct_not_shipped"
End Function
