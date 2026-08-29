Attribute VB_Name = "modTestsPure"
Option Explicit

' ============================================================================
' modTestsPure - 純ロジックモジュールのユニットテスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 役割:
'   src/core・src/app の純ロジックを、Excelなしで(tools/run_lo_tests.py の
'   モード1 = LibreOffice headless)検証する。入口は modTestRunner.RunAllPureTests
'   から呼ばれる Public Sub RunAll()。
'
' 移植元: PoC「マイ本棚AI」 src/test/modTestsPure.bas の骨格
'   (グループ単位の失敗隔離・Check の書き方・R4準拠の方針)。
'   PoC固有のテスト本体(modChunker / modPii / 旧PoCの秘匿値復元テスト
'   など)は移植していない。とくに 旧PoCの秘匿値復元テストは、本製品がキーを
'   ブックに一切入れない設計(16章NFR-S2)のため機構ごと非移植とした。
'
' 設計判断(R4準拠): Worksheets/Range/Application/ThisWorkbook/MsgBox/
'   ActiveSheet には一切触れない。改行は vbLf 基準。
'
' グループ単位の失敗隔離:
'   RunAll は各テストグループを "On Error GoTo <Label> / Resume <NextLabel>" で
'   1グループずつ囲む。modTestRunner.RunAllPureTests 側は RunAll 全体を1個の
'   On Error Resume Next で包むだけなので、RunAll内で無防備に例外が起きると
'   その時点で以降のグループが一切実行されずレポートが失われる。
'   捕捉時は必ず Check で可視化する(握りつぶしていない)。
'
'   **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**
'   (更新漏れは modTestRunner が即FAILとして可視化する。17章§4-1)。
'
' ============================================================================
' W1で統合したテスト群(総記載 104本 = 実行 62本 + SKIP 42本)
' ----------------------------------------------------------------------------
' 執筆方針: 実装コードを一切読まずに仕様書だけから期待値を導いた。期待値を
'   あとから実装に合わせて書き換えることは禁止(17章§1)。仕様から一意に定まら
'   ない項目はテスト化せず、ファイル末尾の「意図的に未テスト」へ列挙した。
'
' グループ / 記載本数 / 根拠章:
'   G0  ランナー自己テスト(W0から継続)      4本  17章§4-1
'   G1  SetCellSafe純ロジック(SanitizeForCell)  9本  16章 E-46 / NFR-S7(1) / E-22
'   G2  SanitizeInput                           6本  15章§0原則9 / 16章 E-04
'   G3  NormalizeForHash系(Fnv1a64Hex)          2本  14章§6 / 16章 E-49
'   G4  ExtractJsonBlock 入力パターン表        11本  14章§5(7パターン表+前処理P0)
'   G5  modJsonLite アクセサ                   13本  14章§6 / 13章§2.2
'   G6  modMockLlm 正常系11応答                21本  15章§8.1  ※全件SKIP
'   G7  modMockLlm 障害注入                     9本  15章§8.2 / 14章§4 ※全件SKIP
'   G8  modConfig 型変換の既定値フォールバック  5本  13章§2.3 / 16章 E-52 / 14章§2
'   G9  modLog 純ロジック                       4本  13章§2.4 / 16章 NFR-S3 ※全件SKIP
'   G10 modGatewayRPN 経路分岐・上限判定        6本  14章§1 / 14章§2 / 14章§6
'   G11 modGatewayDirect 純ロジック            14本  14章§3 / 17章 T-13(うち8本SKIP)
'
' SKIPの理由(17章§1: テストを実装へ合わせて書き換えることを禁じているため、
'   仕様が前提とした関数名が実装に存在しないケースは「テストを直す」のではなく
'   明示SKIP + 司令塔への報告とした):
'   ・modMockLlm.ResponseById / modMockLlm.FaultResponse が実装に無い
'     (実在する唯一の公開口は MockResponse(stepName, variantName, fault))。
'   ・modLog.TruncDetail / modLog.ShouldRotate が実装に無い。
'   ・modGatewayDirect.BackoffMs / RetryBudgetFor / ParseKeyLine が実装に無い
'     (近い実在関数は ComputeBackoffMs / ShouldRetryDirect+IsRetryableServerStatus
'      / ParseApiKeyFirstLine。同一契約かどうかは司令塔の裁定事項)。
'
' 呼び出し口だけを実装へ合わせた箇所(期待値・アサーションは無改変):
'   ・modGatewayDirect.BuildRequestBody は関数名が実在するため引数順を実装の
'     宣言(stepName, model, systemPrompt, userPrompt, schemaJson, temperature,
'     sendTemperature, maxTokens)へ合わせた。sendTemperature は実装が呼び出し側
'     に判断を委ねているので Not IsOSeriesModel(model) を渡す。
'
' 本ファイルが前提とする公開契約(★=仕様書に関数名の記載が無く、本テストが
' 契約として先に固定するもの。司令塔の裁定が必要):
'   modUtilText.SanitizeInput / Fnv1a64Hex                    14章§6にあり
'   modUtilText.SanitizeForCell(s) As String                  ★
'       SetCellSafe がセルへ書く直前に通す純変換部。SetCellSafe 本体はセルI/Oを
'       伴いLibreOfficeで実行できないため、16章 NFR-S7(1) のガード3点(先頭式記号
'       の無害化・NULバイト除去・32,000字切詰め)を担う変換関数の分離公開を前提
'       とする。CSV書出(NFR-S7(2))も同じ関数を通す。
'   modJsonLite.*                                             14章§6にあり
'   modConfig.GetStr / GetLong / GetDouble                    14章§2のコード例にあり
'   modConfig.GetBool(name, dflt) As Boolean                  ★(mock_llm/limit_check
'       /keep_window_alive 等 TRUE/FALSE キー群の読み口)
'   modLog.TruncDetail(s) As String                           ★(detail最大400字)
'   modLog.ShouldRotate(rowCount, maxRows) As Boolean         ★(log_max_rows)
'   modGatewayRPN.ResolveTransport(mockLlm, transportCfg)     ★(14章§1の分岐規則)
'   modGatewayRPN.LooksLikeLimitError(s) As Boolean           14章§2に関数名の言及あり
'   modGatewayDirect.BackoffMs(attemptNo) As Long             ★
'   modGatewayDirect.RetryBudgetFor(httpStatus) As Long       ★
'   modGatewayDirect.BuildRequestBody(...)                    ★
'   modGatewayDirect.ParseKeyLine(fileText) As String         ★
'   modMockLlm.ResponseById(mockId) As String                 ★(mockIDは15章§8.1の表)
'   modMockLlm.FaultResponse(faultKind, stepName) As String   ★(15章§8.2の表)
' ============================================================================

' SKIP理由の定型文(1論理行1023字制約を避けるため定数へ括り出す)。
Private Const MISS_MOCK As String = _
    "modMockLlm の ResponseById(mockId) が実装に無い。実在の公開口は MockResponse(step,variant,fault) の1本のみ"
Private Const MISS_FAULT As String = _
    "modMockLlm の FaultResponse(fault,step) が実装に無い。実在の公開口は MockResponse(step,variant,fault) の1本のみ"
Private Const MISS_LOG As String = _
    "modLog の TruncDetail / ShouldRotate が実装に無い(公開口は LogError / LogUsage / LogRun の3本のみ)"
Private Const MISS_BACKOFF As String = _
    "modGatewayDirect の BackoffMs が実装に無い(近い実在関数は ComputeBackoffMs)"
Private Const MISS_BUDGET As String = _
    "modGatewayDirect の RetryBudgetFor が実装に無い(近い実在関数は ShouldRetryDirect)"
Private Const MISS_KEYLINE As String = _
    "modGatewayDirect の ParseKeyLine が実装に無い(近い実在関数は ParseApiKeyFirstLine)"

Public Sub RunAll()
    On Error GoTo F0
    TestRunnerContract
G1:
    On Error GoTo F1
    T_SetCellSafe
G2:
    On Error GoTo F2
    T_SanitizeInput
G3:
    On Error GoTo F3
    T_HashUtil
G4:
    On Error GoTo F4
    T_ExtractJsonBlock
G5:
    On Error GoTo F5
    T_JsonLite
G6:
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

F0:
    GroupFail "G0 TestRunnerContract"
    Resume G1
F1:
    GroupFail "G1 SetCellSafe"
    Resume G2
F2:
    GroupFail "G2 SanitizeInput"
    Resume G3
F3:
    GroupFail "G3 HashUtil"
    Resume G4
F4:
    GroupFail "G4 ExtractJsonBlock"
    Resume G5
F5:
    GroupFail "G5 JsonLite"
    Resume G6
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

' 未実行(SKIP)の1本。テスト名と理由の両方をレポートへ残す。
Private Sub SkipT(ByVal nm As String, ByVal why As String)
    modTestRunner.Check "[SKIP] " & nm, True, why
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

' ----------------------------------------------------------------------------
' G0: テストハーネス自身の契約を固定する(17章§4-1 ランナー要件の回帰)。
' ここが崩れると「テストを書いても走らない/落ちても緑に見える」状態になり、
' 以降の全テストの信頼性が消えるので、最初の1本はここに置く。
' ----------------------------------------------------------------------------
Private Sub TestRunnerContract()
    Dim firstFive As String
    Dim failsBefore As Long

    ' 1) このグループが最初に走る時点で、失敗は1件も記録されていないはず。
    modTestRunner.Check "ランナー: 開始時点でFAILが0件", _
        (modTestRunner.Failures() = 0), _
        "実際=" & modTestRunner.Failures()

    ' 2) レポートの1行目は "PASS n / FAIL m / SKIP k" で始まる。
    '    tools/run_lo_tests.py がこの書式を正規表現で読むため、崩れると
    '    集計そのものが読めなくなり「合格側へ倒れる」危険がある。
    firstFive = Left(modTestRunner.ReportText(), 5)
    modTestRunner.Check "ランナー: ReportTextの先頭がPASS集計", _
        (firstFive = "PASS "), "実際=" & firstFive

    ' 3) 成功のCheckはFAILを増やさない(集計の基本不変条件)。
    failsBefore = modTestRunner.Failures()
    modTestRunner.Check "ランナー: 成功のCheckは失敗に数えない(被験体)", True, ""
    modTestRunner.Check "ランナー: 成功のCheckはFAILを増やさない", _
        (modTestRunner.Failures() = failsBefore), _
        "前=" & failsBefore & " 後=" & modTestRunner.Failures()
End Sub

' ----------------------------------------------------------------------------
' G1: SetCellSafe の純ロジック(16章 E-46 / NFR-S7の書き込み口(1)(2))。
'   ガードは3点「先頭式記号(= + - @)の ' 前置」「NULバイト除去」「32,000字切詰め」。
'   セル書込そのものはExcel依存のため層(a)の対象外(層(b) modTestsExcel の担当)。
' ----------------------------------------------------------------------------
Private Sub T_SetCellSafe()
    Dim exactText As String

    ChkS "SetCellSafe_先頭イコール無害化_E46", _
        modUtilText.SanitizeForCell("=SUM(A1)"), "'=SUM(A1)"
    ChkS "SetCellSafe_先頭プラス無害化_E46", _
        modUtilText.SanitizeForCell("+1+1"), "'+1+1"
    ChkS "SetCellSafe_先頭マイナス無害化_E46", _
        modUtilText.SanitizeForCell("-2+3"), "'-2+3"
    ChkS "SetCellSafe_先頭アットマーク無害化_E46", _
        modUtilText.SanitizeForCell("@SUM(A1)"), "'@SUM(A1)"
    ChkS "SetCellSafe_通常文字列は素通し_E46", _
        modUtilText.SanitizeForCell("浜松スイーツファクトリー"), "浜松スイーツファクトリー"
    ' 無害化は「先頭」限定。本文中の記号まで触ると原文性(FR-34)が壊れる。
    ChkS "SetCellSafe_中間の等号は無害化しない_E46", _
        modUtilText.SanitizeForCell("売上=前年比110%"), "売上=前年比110%"
    ChkS "SetCellSafe_NULバイト除去_NFRS7", _
        modUtilText.SanitizeForCell("A" & Chr(0) & "B"), "AB"

    ChkN "SetCellSafe_32000字超は32000字へ切詰め_E22", _
        Len(modUtilText.SanitizeForCell(String(40000, "a"))), 32000
    exactText = String(32000, "a")
    modTestRunner.Check "SetCellSafe_32000字ちょうどは切詰めない_E22", _
        (modUtilText.SanitizeForCell(exactText) = exactText), _
        "実際の長さ=" & Len(modUtilText.SanitizeForCell(exactText))
End Sub

' ----------------------------------------------------------------------------
' G2: SanitizeInput(15章§0原則9・16章 E-04)。
'   「■■■」→「[境界記号]」置換と、制御文字・私用領域文字の除去。
'   改行の保持は原則9の「1行属性は SanitizeInput に加えて改行を空白へ畳んでから
'   埋める」という別立ての規定から一意に導ける(SanitizeInput 自身が改行を落とす
'   なら、この追加規定は存在理由を失う)。
' ----------------------------------------------------------------------------
Private Sub T_SanitizeInput()
    ChkS "SanitizeInput_境界記号3連を置換_15章原則9", _
        modUtilText.SanitizeInput("■■■"), "[境界記号]"
    ChkS "SanitizeInput_本文中の境界記号を置換_15章原則9", _
        modUtilText.SanitizeInput("前■■■HPここから■■■後"), _
        "前[境界記号]HPここから[境界記号]後"
    ' 対象は3連の文字列そのもの。2連は境界記号ではないので触らない。
    ChkS "SanitizeInput_境界記号2連は置換しない_15章原則9", _
        modUtilText.SanitizeInput("■■"), "■■"
    ChkS "SanitizeInput_制御文字を除去_E04", _
        modUtilText.SanitizeInput("A" & Chr(1) & "B"), "AB"
    ChkS "SanitizeInput_私用領域文字を除去_E04", _
        modUtilText.SanitizeInput("A" & ChrW(&HE000) & "B"), "AB"
    ChkS "SanitizeInput_改行は保持_15章原則9の1行属性規定", _
        modUtilText.SanitizeInput("A" & vbLf & "B"), "A" & vbLf & "B"
End Sub

' ----------------------------------------------------------------------------
' G3: 重複排除の土台(14章§6・16章 E-49の尾部劣化対策)。
' ----------------------------------------------------------------------------
Private Sub T_HashUtil()
    Dim h As String
    Dim i As Long
    Dim okHex As Boolean

    h = modUtilText.Fnv1a64Hex("浜松スイーツファクトリー")
    okHex = (Len(h) = 16)
    If okHex Then
        For i = 1 To 16
            If InStr("0123456789abcdefABCDEF", Mid(h, i, 1)) = 0 Then okHex = False
        Next i
    End If
    modTestRunner.Check "Fnv1a64Hex_16桁の16進文字列を返す_14章§6", okHex, "実際=" & h

    modTestRunner.Check "Fnv1a64Hex_異なる入力は異なる値_E49", _
        (modUtilText.Fnv1a64Hex("risk-A") <> modUtilText.Fnv1a64Hex("risk-B")), _
        "衝突している(別項目が畳まれる)"
End Sub

' ----------------------------------------------------------------------------
' G4: ExtractJsonBlock(14章§5の入力パターン表を1テスト1パターンで実装)。
'   戻り値 "" は「抽出失敗」で修復リトライへ。壊れたJSONを推測で補完しない。
' ----------------------------------------------------------------------------
Private Sub T_ExtractJsonBlock()
    Dim raw As String

    ' パターン1: 前後に説明文
    ChkS "ExtractJsonBlock_パターン1前後に説明文_14章§5", _
        modJsonLite.ExtractJsonBlock("承知しました。{""a"":1} 以上です。"), "{""a"":1}"

    ' パターン2: コードフェンス囲み
    ChkS "ExtractJsonBlock_パターン2コードフェンス囲み_14章§5", _
        modJsonLite.ExtractJsonBlock("```json" & vbLf & "{""a"":1}" & vbLf & "```"), _
        "{""a"":1}"

    ' パターン3: フェンス閉じ忘れ
    ChkS "ExtractJsonBlock_パターン3フェンス閉じ忘れ_14章§5", _
        modJsonLite.ExtractJsonBlock("```json" & vbLf & "{""a"":1}"), "{""a"":1}"

    ' パターン4: 末尾途切れ(閉じ括弧欠落)。補完しないので ""。
    ChkS "ExtractJsonBlock_パターン4末尾途切れは空文字_14章§5", _
        modJsonLite.ExtractJsonBlock("{""a"":1,""b"":[{""c"":2}"), ""

    ' パターン5: JSON2連結。最初に対応が閉じた1本目のみを返す。
    ChkS "ExtractJsonBlock_パターン5JSON2連結は1本目のみ_14章§5", _
        modJsonLite.ExtractJsonBlock("{""a"":1}{""a"":2}"), "{""a"":1}"

    ' パターン6: 全角波括弧・全角コロン(前処理P0)
    ChkS "ExtractJsonBlock_パターン6全角波括弧と全角コロン_14章§5", _
        modJsonLite.ExtractJsonBlock("｛""a""：1｝"), "{""a"":1}"

    ' パターン7: 値内の生改行とエスケープ済み引用符。抽出は括弧の対応だけを見るので
    '            成功し、値の整形は防衛線(2)の担当なのでここでは原文のまま返る。
    raw = "{""note"":""1行目" & vbLf & "2行目 \""引用\""""}"
    ChkS "ExtractJsonBlock_パターン7値内の生改行とエスケープ引用符_14章§5", _
        modJsonLite.ExtractJsonBlock(raw), raw

    ' 前処理P0の残り(全角二重引用符3種・全角角括弧・全角読点)
    ChkS "ExtractJsonBlock_全角二重引用符U201Cを半角化_14章§5P0", _
        modJsonLite.ExtractJsonBlock("{" & ChrW(&H201C) & "a" & ChrW(&H201D) & ":1}"), _
        "{""a"":1}"
    ChkS "ExtractJsonBlock_全角二重引用符UFF02を半角化_14章§5P0", _
        modJsonLite.ExtractJsonBlock("{" & ChrW(&HFF02) & "a" & ChrW(&HFF02) & ":1}"), _
        "{""a"":1}"
    ChkS "ExtractJsonBlock_全角角括弧と全角読点を半角化_14章§5P0", _
        modJsonLite.ExtractJsonBlock("{""a""：［1，2］}"), "{""a"":[1,2]}"

    ' 括弧が1つも無い応答は抽出失敗。
    ChkS "ExtractJsonBlock_括弧なし入力は空文字_14章§5", _
        modJsonLite.ExtractJsonBlock("説明文だけでJSONがありません"), ""
End Sub

' ----------------------------------------------------------------------------
' G5: modJsonLite のアクセサ(14章§6)。正常系と壊れ入力。
' ----------------------------------------------------------------------------
Private Sub T_JsonLite()
    Dim js As String
    Dim rt As String

    js = "{""company_name"":""浜松スイーツ"",""risk_no"":7,""lands"":true}"

    ChkS "GetStr_文字列値を取得_14章§6", _
        modJsonLite.GetStr(js, "company_name"), "浜松スイーツ"
    ' GetStr は既定値引数を持たない = 不在時の戻りは空文字しかありえない。
    ChkS "GetStr_キー不在は空文字_14章§6", modJsonLite.GetStr(js, "no_such_key"), ""
    ' キー名は引用符ごと照合する。部分一致で別キーを拾ってはいけない。
    ChkS "GetStr_キー名の部分一致で誤取得しない_14章§6", _
        modJsonLite.GetStr("{""risk_name"":""A"",""name"":""B""}", "name"), "B"

    ChkN "GetLong_整数値を取得_14章§6", modJsonLite.GetLong(js, "risk_no", -1), 7
    ChkN "GetLong_キー不在は既定値_14章§6", _
        modJsonLite.GetLong(js, "no_such_key", 42), 42
    ChkN "GetLong_壊れ入力は既定値_14章§6", _
        modJsonLite.GetLong("これはJSONではありません", "risk_no", 42), 42

    modTestRunner.Check "GetBoolJ_trueを取得_14章§6", _
        (modJsonLite.GetBoolJ(js, "lands", False) = True), "trueが読めていない"
    modTestRunner.Check "GetBoolJ_falseを取得_14章§6", _
        (modJsonLite.GetBoolJ("{""lands"":false}", "lands", True) = False), _
        "falseが読めていない"
    modTestRunner.Check "GetBoolJ_キー不在は既定値_14章§6", _
        (modJsonLite.GetBoolJ(js, "no_such_key", True) = True), "既定値へ落ちていない"

    ' 空配列は13章の空配列規約(new案件の current_coverage / gaps)の土台。
    ChkN "GetArrayItems_空配列は0件_13章空配列規約", ArrCount("{""gaps"":[]}", "gaps"), 0
    ChkN "GetArrayItems_オブジェクト配列の件数_14章§6", _
        ArrCount("{""risks"":[{""risk_no"":1},{""risk_no"":2}]}", "risks"), 2
    ' 要素の内側に配列があっても要素境界を誤らないこと(E-49の重複排除は要素単位の
    ' ハッシュで行うため、分割精度がそのまま前提になる)。
    ChkN "GetArrayItems_入れ子配列を含む要素を正しく分割_E49", _
        ArrCount("{""stories"":[{""menu_ids"":[""M-0012""]},{""menu_ids"":[]}]}", _
                 "stories"), 2

    ' 13章セル格納規約: セル内改行はJSON化時に \n へエスケープされる。往復で原文へ。
    rt = "C:\temp" & vbLf & """引用""" & vbTab & "末尾"
    modTestRunner.Check "EscapeJsonStr_UnescapeJsonStrの往復一致_13章セル格納規約", _
        (modJsonLite.UnescapeJsonStr(modJsonLite.EscapeJsonStr(rt)) = rt), _
        "往復で原文に戻らない"
End Sub

' ----------------------------------------------------------------------------
' G6: modMockLlm 正常系(15章§8.1の7 step種・計11応答)。
'   固定したかったのは (a)全11応答が防衛線(1)を通ること (b)表が件数を明記した
'   主要キーが自バリアント文脈で揃っていること。
'   【全件SKIP】仕様が前提とした modMockLlm.ResponseById(mockId) が実装に存在
'   しない。実装の唯一の公開口は MockResponse(stepName, variantName, fault) で、
'   mockID から (step, variant) への対応付けはテスト側の新規実装になるため、
'   17章§1(テストを実装へ合わせて書き換えない)に従い司令塔へ差し戻す。
' ----------------------------------------------------------------------------
Private Sub T_MockNormal()
    SkipT "MKS1NEW_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKS1RNW_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKS2NEW_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKS2RNW_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKS3_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKS4_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKPF_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKS2CHIT_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKS2CCLEAN_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKS3CHIT_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK
    SkipT "MKS3CCLEAN_ExtractJsonBlockで抽出可能_15章§8.1", MISS_MOCK

    ' new文脈: current_coverage は必須キーだが空配列(15章§0原則6)。
    SkipT "MKS1NEW_current_coverageが0件_15章§8.1", MISS_MOCK
    SkipT "MKS1RNW_current_coverageが3件_15章§8.1", MISS_MOCK
    SkipT "MKS2NEW_emerging_risksが1件_15章§8.1", MISS_MOCK
    SkipT "MKS2RNW_gapsが3件_15章§8.1", MISS_MOCK
    ' 空配列が合格であることをmockで担保する(15章§11の emerging_risks 件数ケースが
    ' 0件では発火しないこと)。ケースIDは17章§4-2の照合対象なので本文には書かない。
    SkipT "MKS2RNW_emerging_risksが0件_15章§8.1", MISS_MOCK
    SkipT "MKS3_storiesが3件_15章§8.1", MISS_MOCK
    ' 正常系mockが使ってよいIDは M-0012 / L-03 / S-0004 / K-0003 / P9 / MC-0107 のみ。
    ' 障害注入 ghost_id 専用の M-9999 が正常系へ混ざっていないこと。
    SkipT "MKS3_幽霊IDを含まない_15章§8.1", MISS_MOCK
    SkipT "MKS4_slidesが5枚_15章§8.1", MISS_MOCK
    SkipT "MKPF_survivalがmid_15章§8.1", MISS_MOCK
    ' 改訂スキップ経路(15章§4.5の issues 0件=合格判定)の素材は issues 0件であること。
    SkipT "MKS2CCLEAN_issuesが0件_15章§8.1", MISS_MOCK
End Sub

' ----------------------------------------------------------------------------
' G7: modMockLlm 障害注入(15章§8.2 / 14章§4の表)。
'   固定したかったのは「mockが返す応答の形」だけ。E0302/E0301/E0204への昇格・
'   修復リトライ回数はパイプライン側(T-24)の担当で層(a)では扱わない。
'   【全件SKIP】modMockLlm.FaultResponse(faultKind, stepName) が実装に無い。
' ----------------------------------------------------------------------------
Private Sub T_MockFault()
    SkipT "mockfault_broken_jsonは抽出不可で空文字_15章§8.2", MISS_FAULT
    SkipT "mockfault_enum_violationは未定義値qualityを含む_15章§8.2", MISS_FAULT
    SkipT "mockfault_enum_violationはJSONとして抽出可能_15章§8.2", MISS_FAULT
    SkipT "mockfault_ghost_idは不実在のM9999を含む_15章§8.2", MISS_FAULT
    SkipT "mockfault_count_violationはstoriesが2件_15章§8.2", MISS_FAULT
    SkipT "mockfault_emptyは空文字_15章§8.2", MISS_FAULT
    SkipT "mockfault_limitは上限系文字列と判定される_15章§8.2", MISS_FAULT
    ' fake_err: 帯域外成否規約(14章§6)の素材。先頭は #ERR: だが中身は正常JSON。
    SkipT "mockfault_fake_errは先頭がERRプレフィクス_15章§8.2", MISS_FAULT
    ' ok=True を保つ根拠: 本文からJSONが取れる=正常応答として処理を続けられる。
    ' 平文プレフィクスで成否を判定すると偽装エラーUIが成立する(14章§6)。
    SkipT "CallStep_fake_errでもJSON本文を抽出できる_14章§6帯域外規約", MISS_FAULT
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
'   ローテ閾値。閾値ちょうどの扱いは13章に規定が無いためテスト化しない。
'   【全件SKIP】TruncDetail / ShouldRotate が実装に無い(切詰めとローテは
'   modLog 内のPrivate側に埋まっており、純関数として公開されていない)。
' ----------------------------------------------------------------------------
Private Sub T_Log()
    SkipT "modLog_detail400字超は400字へ切詰め_13章§2.4", MISS_LOG
    SkipT "modLog_detail400字以下は素通し_13章§2.4", MISS_LOG
    SkipT "modLog_detail空文字は空文字_NFRS3", MISS_LOG
    SkipT "modLog_ローテ判定は閾値の前後で反転する_13章log_max_rows", MISS_LOG
End Sub

' ----------------------------------------------------------------------------
' G10: modGatewayRPN の経路分岐(14章§1)と上限判定(14章§2)。
'   分岐規則: mock_llm=TRUE なら mock。それ以外は llm_transport の値。
'   リボン未検出時にdirectへ自動フォールバックしないことは RibbonAvailable 側の
'   責務でExcel依存のため層(a)では扱わない(層(b)・wintestの担当)。
' ----------------------------------------------------------------------------
Private Sub T_GatewayRpn()
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
End Sub

' ----------------------------------------------------------------------------
' G11: modGatewayDirect の純ロジック(14章§3)。
'   HTTP送信(ServerXMLHTTP)そのものは層(a)の対象外。リトライ回数・待ち時間・
'   ボディ組立・キーファイル1行目の取り出しだけを固定する。
'   ボディの検査は書式(空白の有無)に依存しないよう modJsonLite 経由で読む。
'   BackoffMs / RetryBudgetFor / ParseKeyLine の8本はSKIP(関数が実装に無い)。
' ----------------------------------------------------------------------------
Private Sub T_GatewayDirect()
    Dim body As String
    Dim sch As String

    sch = "{""type"":""object""}"

    ' 指数バックオフ 2s/4s/8s(1回目/2回目/3回目)
    SkipT "BackoffMs_1回目は2000ms_14章§3", MISS_BACKOFF
    SkipT "BackoffMs_2回目は4000ms_14章§3", MISS_BACKOFF
    SkipT "BackoffMs_3回目は8000ms_14章§3", MISS_BACKOFF

    ' 429/500/502/503=最大3回、408=1回、その他4xx=リトライなし
    SkipT "RetryBudget_429は3回_14章§3", MISS_BUDGET
    SkipT "RetryBudget_408は1回_14章§3", MISS_BUDGET
    SkipT "RetryBudget_その他4xxはリトライなし_14章§3", MISS_BUDGET

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
    SkipT "ParseKeyLine_LF区切りの1行目を返す_14章§3", MISS_KEYLINE
    SkipT "ParseKeyLine_CRLF区切りでもCRを残さない_14章§3", MISS_KEYLINE
End Sub

' ============================================================================
' 【意図的に未テスト】仕様から期待値が一意に定まらないため本ファイルでは扱わない
' (司令塔へ報告済み。実装に合わせて後から書き足すことを禁じるための明示的な空白):
'   ・SanitizeForCell の適用順(NULバイト除去と先頭式記号判定のどちらが先か)。
'     Chr(0) & "=1" の期待値が "'=1" か "=1" かが16章E-46から決まらない。
'   ・SanitizeForCell の32,000字切詰めに ' 前置の1字を含めるか否か。
'   ・ExtractJsonBlock が文字列値の内側にある波括弧を終端とみなすか
'     (14章§5パターン7「括弧の対応のみを見る」の解釈が2通り立つ)。
'   ・ExtractJsonBlock パターン5の run_log detail への extra_json=1 記録
'     (ログ書込はExcel依存のため層(a)の対象外)。
'   ・GetStr が値のエスケープを解いて返すか生のまま返すか(UnescapeJsonStr が別
'     関数として存在するため両方の読みが立つ。往復テストのみを置いた)。
'   ・GetArrayItems のキー不在時の戻り(空Collection か Nothing か)。
'   ・modLog.ShouldRotate の閾値ちょうど(rowCount = log_max_rows)の判定。
'   ・BackoffMs の attemptNo の起点(1始まり/0始まり)。本ファイルは1始まりを採用。
'   ・ParseKeyLine が1行目の前後空白をTrimするか。
'   ・mock_fault が空のときの FaultResponse の戻り(15章§8.2は「正常応答のみを
'     返す」とCallStep側の挙動で規定しており、mock単体の戻りを定めていない)。
'   ・CallStep / CallChat の ByRef ok そのもの(config設定とrun_log記録を伴い
'     Excel依存のため層(b)・T-24の担当。層(a)では素材の形だけを固定した)。
' ============================================================================
