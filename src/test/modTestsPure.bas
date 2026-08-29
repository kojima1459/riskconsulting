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
' 移植元: PoC「マイ本棚AI」 src/test/modTestsPure.bas の骨格(グループ単位の
'   失敗隔離・Check の書き方・R4準拠の方針)。PoC固有のテスト本体(modChunker /
'   modPii / 旧PoCの秘匿値復元テストなど)は移植していない。とくに秘匿値復元
'   テストは、本製品がキーをブックに入れない設計(16章NFR-S2)のため機構ごと非移植。
'
' 設計判断(R4準拠): Worksheets/Range/Application/ThisWorkbook/MsgBox/
'   ActiveSheet には一切触れない。改行は vbLf 基準。
'
' グループ単位の失敗隔離:
'   RunAll は各テストグループを "On Error GoTo <Label> / Resume <NextLabel>" で
'   1グループずつ囲む。modTestRunner.RunAllPureTests 側は RunAll 全体を1個の
'   On Error Resume Next で包むだけなので、RunAll内で無防備に例外が起きるとその
'   時点で以降のグループが実行されずレポートが失われる。捕捉時は必ず Check で
'   可視化する(握りつぶしていない)。
'
'   **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**
'   (更新漏れは modTestRunner が即FAILとして可視化する。17章§4-1)。
'
' ============================================================================
' W1で統合したテスト群(総記載 110本 = 実行 110本 / SKIP 0本)
' ----------------------------------------------------------------------------
' 執筆方針: 実装コードを一切読まずに仕様書だけから期待値を導いた。期待値を
'   あとから実装に合わせて書き換えることは禁止(17章§1)。仕様から一意に定まら
'   ない項目はテスト化せず、ファイル末尾の「意図的に未テスト」へ列挙した。
'
' グループ / 本数 / 根拠章:
'   G0  ランナー自己テスト(W0から継続)      4本  17章§4-1
'   G1  SetCellSafe純ロジック(SanitizeForCell)  9本  16章 E-46 / NFR-S7(1) / E-22
'   G2  SanitizeInput                           6本  15章§0原則9 / 16章 E-04
'   G3  NormalizeForHash系(Fnv1a64Hex)          2本  14章§6 / 16章 E-49
'   G4  ExtractJsonBlock 入力パターン表        11本  14章§5(7パターン表+前処理P0)
'   G5  modJsonLite アクセサ                   13本  14章§6 / 13章§2.2
'   G6  modMockLlm 正常系11応答                21本  15章§8.1
'   G7  modMockLlm 障害注入                     9本  15章§8.2 / 14章§4
'   G8  modConfig 型変換の既定値フォールバック  5本  13章§2.3 / 16章 E-52 / 14章§2
'   G9  modLog 純ロジック                       4本  13章§2.4 / 16章 NFR-S3
'   G10 modGatewayRPN 経路分岐・上限・DecideOk 10本  14章§1 / §2 / §6
'   G11 modGatewayDirect 純ロジック            16本  14章§3 / 17章 T-13
'
' 裁定書5(W1整合)による変更点。アサーションの期待値・判定式は無改変:
'   ・呼び先の関数名を14章§6の確定名へ合わせ、SKIP 42本をすべて実行へ戻した
'     (modMockLlm の ResponseById / FaultResponse = A-1、modLog の TruncDetail /
'      ShouldRotate = A-2、modGatewayDirect の BackoffMs / RetryBudgetFor /
'      ParseKeyLine = A-3)。
'   ・司令塔指示による新規は6本のみ: G10に DecideOk 4本(C-11。14章§6の帯域外
'     成否規約の回帰網)、G11に IsOSeriesModel 2本(C-12。14章§3の根拠文言)。
'   ・BuildRequestBody の引数順は14章§6の宣言どおり。sendTemperature は実装が
'     呼び出し側に判断を委ねる契約(1判断1箇所)なので Not IsOSeriesModel(model)
'     を渡す。その判定自体は C-12 の2本が単体で押さえる。
'
' 本ファイルが前提とする公開契約(すべて14章§6にある。★は裁定書5で§6へ追記):
'   modUtilText.SanitizeInput / Fnv1a64Hex / SanitizeForCell ★
'   modJsonLite.*(引数名も§6の json / key / dflt に一致)
'   modConfig.GetStr / GetLong / GetDouble / GetBool
'   modLog.TruncDetail ★ / ShouldRotate ★
'   modGatewayRPN.ResolveTransport / LooksLikeLimitError / DecideOk ★
'   modGatewayDirect.BackoffMs ★ / RetryBudgetFor ★ / ParseKeyLine ★ /
'                    IsOSeriesModel ★ / BuildRequestBody ★
'   modMockLlm.MockResponse ★ / ResponseById ★ / FaultResponse ★
' ============================================================================

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
