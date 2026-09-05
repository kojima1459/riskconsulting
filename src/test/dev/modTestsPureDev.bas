Attribute VB_Name = "modTestsPureDev"
Option Explicit

' ============================================================================
' modTestsPureDev - dev専用の純ロジックテスト(17章§4-1 層(a))。裁定書30 裁定1(d)。
' ----------------------------------------------------------------------------
' 役割:
'   配布物から外れたモジュール(direct経路の modGatewayDirect。modules.json の
'   ship:false)だけを検査する。**配布ブックには載らない**(本ファイル自身も
'   ship:false)ので、prod の純層本数には1本も入らない。
'   入口は Public Sub RunAll()。dev ビルドでは modTestsPureHook(dev版)の
'   RunAll から呼ばれ、Linux 側では tools/run_lo_tests.py --pure-set dev-only
'   (ゲート lo-pure-dev)が直接呼ぶ。
'
' 本ファイルのテスト本数: 17本(G11)。wintest/tests_expected.txt の
'   dev_only 行がこの本数の正である。
'
' 出自: modTestsPure2.bas の G11 群(W1・裁定書5 C-12)をそのまま移設した。
'   アサーション・期待値・テスト名は一字も変えていない(移設で回帰網の意味が
'   変わらないようにするため)。
'
' 設計判断(R4準拠): Worksheets/Range/Application/ThisWorkbook/MsgBox/
'   ActiveSheet には一切触れない。改行は vbLf 基準。
'
' 本ファイルが前提とする公開契約(すべて14章§6にある):
'   modGatewayDirect.BackoffMs / RetryBudgetFor / ParseKeyLine /
'                    IsOSeriesModel / BuildRequestBody
'   modJsonLite.GetStr / GetBoolJ
' ============================================================================

Public Sub RunAll()
    On Error GoTo F11
    T_GatewayDirect
GDone:
    On Error GoTo 0
    Exit Sub

F11:
    GroupFail "G11 GatewayDirect"
    Resume GDone
End Sub

Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

' 文字列一致の1本。期待値と実際値の両方をレポートへ残す。
Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), "期待=[" & want & "] 実際=[" & act & "]"
End Sub

' 数値一致の1本。
Private Sub ChkN(ByVal nm As String, ByVal act As Long, ByVal want As Long)
    modTestRunner.Check nm, (act = want), "期待=" & want & " 実際=" & act
End Sub

' ----------------------------------------------------------------------------
' G11: modGatewayDirect の純ロジック(14章§3)。
'   HTTP送信そのものは層(a)の対象外。リトライ回数・待ち時間・
'   ボディ組立・キーファイル1行目の取り出しだけを固定する。
'   ボディの検査は書式(空白の有無)に依存しないよう modJsonLite 経由で読む。
'   BackoffMs / RetryBudgetFor / ParseKeyLine は14章§6で確定した名前(裁定書5 A-3)。
' ----------------------------------------------------------------------------
Private Sub T_GatewayDirect()
    Dim body As String
    Dim sch As String

    sch = "{""type"":""object""}"

    ' 指数バックオフ 2s/4s/8s(1回目/2回目/3回目)
    ChkN "BackoffMs_1回目は2000ms_14章§3", modGatewayDirect.BackoffMs(1), 2000
    ChkN "BackoffMs_2回目は4000ms_14章§3", modGatewayDirect.BackoffMs(2), 4000
    ChkN "BackoffMs_3回目は8000ms_14章§3", modGatewayDirect.BackoffMs(3), 8000

    ' 429/500/502/503=最大3回、408=1回、その他4xx=リトライなし
    ChkN "RetryBudget_429は3回_14章§3", modGatewayDirect.RetryBudgetFor(429), 3
    ChkN "RetryBudget_408は1回_14章§3", modGatewayDirect.RetryBudgetFor(408), 1
    ChkN "RetryBudget_その他4xxはリトライなし_14章§3", _
        modGatewayDirect.RetryBudgetFor(404), 0
    ChkN "RetryBudget_500は3回_14章§3", modGatewayDirect.RetryBudgetFor(500), 3

    ' o系モデル判定(14章§3「o系モデル名(先頭"o")では temperature を送らない」)。
    ' BuildRequestBody は渡された sendTemperature しか見ない契約(1判断1箇所)な
    ' ので、モデル名からフラグを起こす側をここで単体固定する(裁定書5 C-12)。
    ' 期待値 o3=True / gpt-4.1=False は同節の根拠文言そのもの。
    modTestRunner.Check "IsOSeriesModel_o3はTrue_14章§3", _
        (modGatewayDirect.IsOSeriesModel("o3") = True), _
        "先頭が o のモデル名を o系と判定していない"
    modTestRunner.Check "IsOSeriesModel_gpt41はFalse_14章§3", _
        (modGatewayDirect.IsOSeriesModel("gpt-4.1") = False), _
        "o系でないモデル名を o系と誤判定している"

    body = modGatewayDirect.BuildRequestBody("s1", "gpt-4.1", "S", "U", sch, 0.3, _
        (Not modGatewayDirect.IsOSeriesModel("gpt-4.1")), 0)
    ChkS "BuildRequestBody_modelをconfig値で載せる_14章§3", _
        modJsonLite.GetStr(body, "model"), "gpt-4.1"
    ' llm_max_tokens=0 のときはキー自体を送らない(14章§3の但し書き)。
    modTestRunner.Check "BuildRequestBody_max_tokens0はキーごと送らない_14章§3", _
        (InStr(body, """max_tokens""") = 0), "max_tokensキーが送出されている"
    modTestRunner.Check "BuildRequestBody_json_schemaはstrictがtrue_14章§3", _
        (modJsonLite.GetBoolJ(body, "strict", False) = True), "strict:trueでない"
    ChkS "BuildRequestBody_json_schemaのnameはstep名_14章§3", _
        modJsonLite.GetStr(body, "name"), "s1"

    ' o系モデル名(先頭"o")では temperature を送らない。
    body = modGatewayDirect.BuildRequestBody("s1", "o3", "S", "U", sch, 0.3, _
        (Not modGatewayDirect.IsOSeriesModel("o3")), 0)
    modTestRunner.Check "BuildRequestBody_o系モデルはtemperatureを送らない_14章§3", _
        (InStr(body, """temperature""") = 0), "o系なのにtemperatureが載っている"

    ' 外部由来テキストがボディを壊さないこと(引用符はJSONエスケープされる)。
    body = modGatewayDirect.BuildRequestBody("s1", "gpt-4.1", "S", _
        "彼は""はい""と答えた", sch, 0.3, _
        (Not modGatewayDirect.IsOSeriesModel("gpt-4.1")), 0)
    modTestRunner.Check "BuildRequestBody_user本文の引用符をエスケープ_14章§3", _
        (InStr(body, "\""") > 0), "引用符がエスケープされていない"

    ' キーはファイル1行目のみを使う(ブック・config・ログ・配布物に一切残さない
    ' =16章 NFR-S2。ローテ不能の借用キーのため一度露出したら恒久被害)。
    ChkS "ParseKeyLine_LF区切りの1行目を返す_14章§3", _
        modGatewayDirect.ParseKeyLine("KEY-1行目" & vbLf & "KEY-2行目"), "KEY-1行目"
    ChkS "ParseKeyLine_CRLF区切りでもCRを残さない_14章§3", _
        modGatewayDirect.ParseKeyLine("KEY-1行目" & vbCrLf & "KEY-2行目"), "KEY-1行目"
End Sub
