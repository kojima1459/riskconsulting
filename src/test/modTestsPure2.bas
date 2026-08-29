Attribute VB_Name = "modTestsPure2"
Option Explicit

' ============================================================================
' modTestsPure2 - 純ロジックモジュールのユニットテスト(17章§4-1 層(a))G6〜G11
' ----------------------------------------------------------------------------
' 役割:
'   modTestsPure.bas が30,000字契約の警告域(29,851字)に達したため、
'   mock/log/gateway系のグループ(G6〜G11)をここへ切り出した。アサーション・
'   期待値・テスト名は modTestsPure.bas から一字も変えていない。グループ単位の
'   失敗隔離(On Error構造)もそのまま移設した。
'
'   入口は Public Sub RunAll()。modTestRunner.RunAllPureTests から直接は
'   呼ばれず、modTestsPure.RunAll の末尾から呼ばれる(呼び出し側=ランナーに
'   分割を意識させない。外部公開の入口は modTestsPure.RunAll ひとつのまま)。
'
' 設計判断(R4準拠): Worksheets/Range/Application/ThisWorkbook/MsgBox/
'   ActiveSheet には一切触れない。改行は vbLf 基準。
'
' 本ファイルのテスト本数: 70本(G6 21 / G7 9 / G8 5 / G9 4 / G10 10 / G11 16 /
'   Fnv1a64Hex系はG3で modTestsPure.bas 側)。modTestsPure.bas(G0〜G5) 45本と
'   合わせて総記載 115本 = 実行 115本 / SKIP 0本(wintest/tests_expected.txt)。
'
' グループ / 本数 / 根拠章:
'   G6  modMockLlm 正常系11応答                21本  15章§8.1
'   G7  modMockLlm 障害注入                     9本  15章§8.2 / 14章§4
'   G8  modConfig 型変換の既定値フォールバック  5本  13章§2.3 / 16章 E-52 / 14章§2
'   G9  modLog 純ロジック                       4本  13章§2.4 / 16章 NFR-S3
'   G10 modGatewayRPN 経路分岐・上限・DecideOk 10本  14章§1 / §2 / §6
'   G11 modGatewayDirect 純ロジック            16本  14章§3 / 17章 T-13
'
' 本ファイルが前提とする公開契約(すべて14章§6にある。★は裁定書5で§6へ追記):
'   modConfig.GetStr / GetLong / GetDouble / GetBool
'   modLog.TruncDetail ★ / ShouldRotate ★
'   modGatewayRPN.ResolveTransport / LooksLikeLimitError / DecideOk ★
'   modGatewayDirect.BackoffMs ★ / RetryBudgetFor ★ / ParseKeyLine ★ /
'                    IsOSeriesModel ★ / BuildRequestBody ★
'   modMockLlm.MockResponse ★ / ResponseById ★ / FaultResponse ★
' ============================================================================

Public Sub RunAll()
    On Error GoTo F6
    T_MockNormal
G7:
    On Error GoTo F7
    T_MockFault
G8:
    On Error GoTo F8
    T_Config
G9:
    On Error GoTo F9
    T_Log
G10:
    On Error GoTo F10
    T_GatewayRpn
G11:
    On Error GoTo F11
    T_GatewayDirect
GDone:
    On Error GoTo 0
    Exit Sub

F6:
    GroupFail "G6 MockNormal"
    Resume G7
F7:
    GroupFail "G7 MockFault"
    Resume G8
F8:
    GroupFail "G8 Config"
    Resume G9
F9:
    GroupFail "G9 Log"
    Resume G10
F10:
    GroupFail "G10 GatewayRpn"
    Resume G11
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

' 配列要素数。未実装・未注入・Nothing はすべて -1 に落として1件の失敗にする。
Private Function ArrCount(ByVal js As String, ByVal ky As String) As Long
    Dim c As Collection
    On Error GoTo Bad
    Set c = modJsonLite.GetArrayItems(js, ky)
    If c Is Nothing Then
        ArrCount = -1
    Else
        ArrCount = c.Count
    End If
    Exit Function
Bad:
    ArrCount = -1
End Function

' 15章§8.1のmock応答を14章§5の防衛線(1)に通した本体JSON。呼び先は14章§6の
' modMockLlm.ResponseById(mockId)。mockId は15章§8.1の表の11 IDが正。
Private Function MockJson(ByVal mockId As String) As String
    MockJson = modJsonLite.ExtractJsonBlock(modMockLlm.ResponseById(mockId))
End Function

' ----------------------------------------------------------------------------
' G6: modMockLlm 正常系(15章§8.1の7 step種・計11応答)。
'   固定したいのは (a)全11応答が防衛線(1)を通ること (b)表が件数を明記した
'   主要キーが自バリアント文脈で揃っていること。
'   呼び先は14章§6で確定した ResponseById(mockId)。mockID の正は15章§8.1の表。
' ----------------------------------------------------------------------------
Private Sub T_MockNormal()
    Dim js As String

    modTestRunner.Check "MKS1NEW_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S1-NEW") <> ""), "抽出できない"
    modTestRunner.Check "MKS1RNW_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S1-RNW") <> ""), "抽出できない"
    modTestRunner.Check "MKS2NEW_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S2-NEW") <> ""), "抽出できない"
    modTestRunner.Check "MKS2RNW_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S2-RNW") <> ""), "抽出できない"
    modTestRunner.Check "MKS3_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S3") <> ""), "抽出できない"
    modTestRunner.Check "MKS4_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S4") <> ""), "抽出できない"
    modTestRunner.Check "MKPF_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-PF") <> ""), "抽出できない"
    modTestRunner.Check "MKS2CHIT_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S2C-HIT") <> ""), "抽出できない"
    modTestRunner.Check "MKS2CCLEAN_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S2C-CLEAN") <> ""), "抽出できない"
    modTestRunner.Check "MKS3CHIT_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S3C-HIT") <> ""), "抽出できない"
    modTestRunner.Check "MKS3CCLEAN_ExtractJsonBlockで抽出可能_15章§8.1", _
        (MockJson("MK-S3C-CLEAN") <> ""), "抽出できない"

    ' new文脈: current_coverage は必須キーだが空配列(15章§0原則6)。
    js = MockJson("MK-S1-NEW")
    ChkN "MKS1NEW_current_coverageが0件_15章§8.1", ArrCount(js, "current_coverage"), 0

    ChkN "MKS1RNW_current_coverageが3件_15章§8.1", _
        ArrCount(MockJson("MK-S1-RNW"), "current_coverage"), 3

    js = MockJson("MK-S2-NEW")
    ChkN "MKS2NEW_emerging_risksが1件_15章§8.1", ArrCount(js, "emerging_risks"), 1

    js = MockJson("MK-S2-RNW")
    ChkN "MKS2RNW_gapsが3件_15章§8.1", ArrCount(js, "gaps"), 3
    ' 空配列が合格であることをmockで担保する(15章§11の emerging_risks 件数ケースが
    ' 0件では発火しないこと)。ケースIDは17章§4-2の照合対象なので本文には書かない。
    ChkN "MKS2RNW_emerging_risksが0件_15章§8.1", ArrCount(js, "emerging_risks"), 0

    js = MockJson("MK-S3")
    ChkN "MKS3_storiesが3件_15章§8.1", ArrCount(js, "stories"), 3
    ' 正常系mockが使ってよいIDは M-0012 / L-03 / S-0004 / K-0003 / P9 / MC-0107 のみ。
    ' 障害注入 ghost_id 専用の M-9999 が正常系へ混ざっていないこと。
    modTestRunner.Check "MKS3_幽霊IDを含まない_15章§8.1", _
        (InStr(js, "M-9999") = 0), "M-9999が正常系mockに混入している"

    ChkN "MKS4_slidesが5枚_15章§8.1", ArrCount(MockJson("MK-S4"), "slides"), 5
    ChkS "MKPF_survivalがmid_15章§8.1", _
        modJsonLite.GetStr(MockJson("MK-PF"), "survival"), "mid"
    ' 改訂スキップ経路(15章§4.5の issues 0件=合格判定)の素材は issues 0件であること。
    ChkN "MKS2CCLEAN_issuesが0件_15章§8.1", _
        ArrCount(MockJson("MK-S2C-CLEAN"), "issues"), 0
End Sub

' ----------------------------------------------------------------------------
' G7: modMockLlm 障害注入(15章§8.2 / 14章§4の表)。
'   固定するのは「mockが返す応答の形」だけ。E0302/E0301/E0204への昇格・修復
'   リトライ回数はパイプライン側(T-24)の担当で層(a)では扱わない。
'   呼び先は14章§6で確定した FaultResponse(faultKind, stepName)。状態レス契約
'   (15章§8.2)なので、G6が先に回ったか否かに結果が依存しない。
' ----------------------------------------------------------------------------
Private Sub T_MockFault()
    Dim body As String

    ChkS "mockfault_broken_jsonは抽出不可で空文字_15章§8.2", _
        modJsonLite.ExtractJsonBlock(modMockLlm.FaultResponse("broken_json", "s1")), ""

    body = modMockLlm.FaultResponse("enum_violation", "s2")
    modTestRunner.Check "mockfault_enum_violationは未定義値qualityを含む_15章§8.2", _
        (InStr(body, """category"":""quality""") > 0), "category=quality(enum外の完全形)が入っていない"
    modTestRunner.Check "mockfault_enum_violationはJSONとして抽出可能_15章§8.2", _
        (modJsonLite.ExtractJsonBlock(body) <> ""), _
        "enum違反はJSON構造としては正常であるべき"

    modTestRunner.Check "mockfault_ghost_idは不実在のM9999を含む_15章§8.2", _
        (InStr(modMockLlm.FaultResponse("ghost_id", "s3"), "M-9999") > 0), _
        "不実在ID M-9999 が入っていない"

    ChkN "mockfault_count_violationはstoriesが2件_15章§8.2", _
        ArrCount(modJsonLite.ExtractJsonBlock( _
            modMockLlm.FaultResponse("count_violation", "s3")), "stories"), 2

    ChkS "mockfault_emptyは空文字_15章§8.2", _
        modMockLlm.FaultResponse("empty", "s1"), ""

    modTestRunner.Check "mockfault_limitは上限系文字列と判定される_15章§8.2", _
        (modGatewayRPN.LooksLikeLimitError(modMockLlm.FaultResponse("limit", "s1")) = True), _
        "E0204へ落ちる素材になっていない"

    ' fake_err: 帯域外成否規約(14章§6)の素材。先頭は #ERR: だが中身は正常JSON。
    body = modMockLlm.FaultResponse("fake_err", "s1")
    ChkS "mockfault_fake_errは先頭がERRプレフィクス_15章§8.2", _
        Left(body, Len("#ERR:E0201:")), "#ERR:E0201:"
    ' ok=True を保つ根拠: 本文からJSONが取れる=正常応答として処理を続けられる。
    ' 平文プレフィクスで成否を判定すると偽装エラーUIが成立する(14章§6)。
    modTestRunner.Check "CallStep_fake_errでもJSON本文を抽出できる_14章§6帯域外規約", _
        (modJsonLite.ExtractJsonBlock(body) <> ""), _
        "偽装エラー行のせいで本文JSONが失われている"

    modMockLlm.ResetFaultOnce
    modTestRunner.Check "mockfault_broken_json_onceは初回のみ破損_15章§8.2", _
        (modJsonLite.ExtractJsonBlock(modMockLlm.FaultResponse("broken_json_once", "s1")) = ""), _
        "初回呼出が破損応答になっていない"
    modTestRunner.Check "mockfault_broken_json_onceの2回目は正常_15章§8.2", _
        (modJsonLite.ExtractJsonBlock(modMockLlm.FaultResponse("broken_json_once", "s1")) <> ""), _
        "2回目呼出(修復想定)が正常応答になっていない"
End Sub

' ----------------------------------------------------------------------------
' G8: modConfig の型変換(13章§2.3の既定値表・16章 E-52)。
'   未登録キーは必ず呼び出し側の既定値へ落ちること(config欠損で製品を止めない)。
' ----------------------------------------------------------------------------
Private Sub T_Config()
    ChkS "modConfig_GetStr未登録キーは既定値_E52", _
        modConfig.GetStr("rpn_no_such_key", "既定"), "既定"
    ChkN "modConfig_GetLong未登録キーは既定値_E52", _
        modConfig.GetLong("rpn_no_such_key", 1200), 1200
    ' 14章§2の入れ子既定値(Step別上書き→共通→0)がconfig未登録でも0に収束すること。
    ChkN "modConfig_GetLong既定値0でもフォールバックする_14章§2", _
        modConfig.GetLong("rpn_no_such_key1", modConfig.GetLong("rpn_no_such_key2", 0)), 0
    modTestRunner.Check "modConfig_GetDouble未登録キーは既定値_E52", _
        (modConfig.GetDouble("rpn_no_such_key", 0.3) = 0.3), _
        "実際=" & modConfig.GetDouble("rpn_no_such_key", 0.3)
    modTestRunner.Check "modConfig_GetBool未登録キーは既定値_E52", _
        (modConfig.GetBool("rpn_no_such_key", True) = True) And _
        (modConfig.GetBool("rpn_no_such_key", False) = False), _
        "TrueとFalseの双方で既定値へ落ちること"
End Sub

' ----------------------------------------------------------------------------
' G9: modLog の純ロジック(13章§2.4 / 16章 NFR-S3)。
'   detail は最大400字のメタのみ。log_max_rows(既定2000)は err/usage/run 共通の
'   ローテ閾値。閾値ちょうど(rowCount = log_max_rows)の扱いは14章§6で「真」と
'   確定したが、本ファイルの判定式は原本のまま閾値の前後だけを見る。
' ----------------------------------------------------------------------------
Private Sub T_Log()
    ChkN "modLog_detail400字超は400字へ切詰め_13章§2.4", _
        Len(modLog.TruncDetail(String(500, "x"))), 400
    ChkN "modLog_detail400字以下は素通し_13章§2.4", _
        Len(modLog.TruncDetail(String(399, "x"))), 399
    ChkS "modLog_detail空文字は空文字_NFRS3", modLog.TruncDetail(""), ""
    modTestRunner.Check "modLog_ローテ判定は閾値の前後で反転する_13章log_max_rows", _
        (modLog.ShouldRotate(1999, 2000) = False) And _
        (modLog.ShouldRotate(2001, 2000) = True), _
        "log_max_rows=2000 の前後で判定が反転しない"
    modTestRunner.Check "modLog_ローテ判定は閾値ちょうどで真_14章§6", _
        (modLog.ShouldRotate(2000, 2000) = True), "rowCount=maxRowsで回転しない"
End Sub

' ----------------------------------------------------------------------------
' G10: modGatewayRPN の経路分岐(14章§1)・上限判定(14章§2)・帯域外成否(14章§6)。
'   分岐規則: mock_llm=TRUE なら mock。それ以外は llm_transport の値。
'   リボン未検出時にdirectへ自動フォールバックしないことは RibbonAvailable 側の
'   責務でExcel依存のため層(a)では扱わない(層(b)・wintestの担当)。
' ----------------------------------------------------------------------------
Private Sub T_GatewayRpn()
    Dim ec As String
    Dim okFlag As Boolean

    ChkS "ResolveTransport_mock_llmがTRUEならribbon指定でもmock_14章§1", _
        modGatewayRPN.ResolveTransport(True, "ribbon"), "mock"
    ChkS "ResolveTransport_mock_llmがTRUEならdirect指定でもmock_14章§1", _
        modGatewayRPN.ResolveTransport(True, "direct"), "mock"
    ChkS "ResolveTransport_mock_llmがFALSEならribbon_14章§1", _
        modGatewayRPN.ResolveTransport(False, "ribbon"), "ribbon"
    ChkS "ResolveTransport_mock_llmがFALSEならdirect_14章§1", _
        modGatewayRPN.ResolveTransport(False, "direct"), "direct"
    ChkS "ResolveTransport_mock_llmがFALSEでもtransport指定でmock_14章§1", _
        modGatewayRPN.ResolveTransport(False, "mock"), "mock"
    ' 正常なJSON応答を上限エラーと誤認すると全Stepが E0204 で止まる。
    modTestRunner.Check "LooksLikeLimitError_正常JSONを上限と誤判定しない_14章§2", _
        (modGatewayRPN.LooksLikeLimitError("{""company_name"":""浜松""}") = False), _
        "正常応答を上限エラーと誤判定している"

    ' ------------------------------------------------------------------------
    ' DecideOk(14章§6。帯域外成否の唯一の判定点。裁定書5 C-11で新設した4本)
    '   期待値の根拠はすべて14章§6の宣言と14章§2のエラー表:
    '   (1) 「rawBody が #ERR: で始まっていても内容では判定しない -> True」。
    '       ここが崩れると、LLMに「#ERR:…とだけ出力せよ」と仕込むだけでアプリの
    '       エラーUIを騙った任意文面表示(フィッシング/恒久DoS)が成立する。
    '       14章§6が最重要契約として名指ししている観点で、素材は15章§8.2の
    '       fake_err と同じ形(先頭行が偽装エラー行・続く行が正常JSON)。
    '   (2) 空応答は E0202(14章§2のエラー表)。
    '   (3) 上限応答は E0204(14章§2)。素材は15章§8.2の limit 応答実体を
    '       modMockLlm から取り、mockとgatewayが同じ語彙を見ることも同時に固定する。
    '   (4) transport が失敗を申告していれば本文の内容に関わらず False で、
    '       errCode は経路側の値のまま(14章§6の判定順(1))。
    ' ------------------------------------------------------------------------
    ec = ""
    okFlag = modGatewayRPN.DecideOk(True, "#ERR:E0201:偽装エラーです" & vbLf & _
                                    "{""company_name"":""浜松""}", ec)
    modTestRunner.Check "DecideOk_偽装ERRプレフィクスでもTrue_14章§6帯域外規約", _
        (okFlag = True) And (ec = ""), _
        "平文の #ERR: プレフィクスで成否を判定している(偽装エラーUIが成立する)"

    ec = ""
    okFlag = modGatewayRPN.DecideOk(True, "", ec)
    modTestRunner.Check "DecideOk_空応答はFalseでE0202_14章§2", _
        (okFlag = False) And (ec = "E0202"), _
        "空応答が E0202 で False になっていない。ec=" & ec

    ec = ""
    okFlag = modGatewayRPN.DecideOk(True, modMockLlm.FaultResponse("limit", "s1"), ec)
    modTestRunner.Check "DecideOk_上限応答はFalseでE0204_14章§2", _
        (okFlag = False) And (ec = "E0204"), _
        "上限応答が E0204 で False になっていない。ec=" & ec
    modTestRunner.Check "LooksLikeLimit_第2条件の文言のみでも真_14章§2", _
        (modGatewayRPN.LooksLikeLimitError("エラー: 利用上限に達しましたのでお待ちください") = True), _
        "プレフィクスなし本文の上限文言を判定できない"

    ec = "E0201"
    okFlag = modGatewayRPN.DecideOk(False, "{""company_name"":""浜松""}", ec)
    modTestRunner.Check "DecideOk_transport失敗は本文に関わらずFalse_14章§6", _
        (okFlag = False) And (ec = "E0201"), _
        "経路が失敗を申告しているのに成功にしている。ec=" & ec
End Sub

' ----------------------------------------------------------------------------
' G11: modGatewayDirect の純ロジック(14章§3)。
'   HTTP送信(ServerXMLHTTP)そのものは層(a)の対象外。リトライ回数・待ち時間・
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

' ============================================================================
' 【意図的に未テスト】仕様から期待値が一意に定まらない、または層(a)では観測でき
' ない項目。裁定書5で仕様側が確定した点は「確定済」と明記するが、**本ファイルに
' アサーションは足さない**(追加が許されたのは C-11 / C-12 の6本だけ。増やすときは
' 司令塔の指示と tests_expected の同時更新が要る):
'   ・確定済で未テストのまま置くもの: SanitizeForCell の適用順と ' 前置を切詰めに
'     含めるか(16章NFR-S7(1))/ ExtractJsonBlock が引用符内の波括弧を数えないこと
'     (14章§5)/ GetStr がエスケープを解いて返すこと・GetArrayItems のキー不在は
'     空Collection・ShouldRotate の閾値ちょうどは真・BackoffMs の attemptNo は
'     1始まり・ParseKeyLine は前後Trim・fault空の FaultResponse は ""(14章§6)。
'   ・層(a)では観測できないもの: ExtractJsonBlock の extra_json / 前処理P0の
'     fw_normalized / E-49の除去件数の run_log detail 記録(ログ書込はExcel依存。
'     検問は17章 T-24 のDoD)。
'   ・CallStep / CallChat の ByRef ok そのもの(config設定とrun_log記録を伴い
'     Excel依存のため層(b)・T-24の担当)。ただし ok の決定ロジックは DecideOk へ
'     切り出されたので、G10の4本が層(a)で規約を押さえている。
' ============================================================================
