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
' モジュール分割(30,000字契約対応。アサーション・期待値・テスト名は無改変):
'   本ファイルは G0〜G5 のみを保持する。G6〜G11(mock/log/gateway系)は
'   modTestsPure2.bas へ移設した。グループ単位の失敗隔離(On Error構造)ごと
'   移し、本ファイルの RunAll 末尾から modTestsPure2.RunAll を呼ぶ。
'   modTestRunner.RunAllPureTests からの入口は従来どおり本ファイルの
'   RunAll ひとつのまま(呼び出し側に分割を意識させない)。
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
' W1で統合したテスト群(総記載 115本 = 実行 115本 / SKIP 0本)。
' 内訳: 本ファイル(G0〜G5) 45本 + modTestsPure2(G6〜G11) 70本 = 115本。
' ----------------------------------------------------------------------------
' 執筆方針: 実装コードを一切読まずに仕様書だけから期待値を導いた。期待値を
'   あとから実装に合わせて書き換えることは禁止(17章§1)。仕様から一意に定まら
'   ない項目はテスト化せず、modTestsPure2.bas 末尾の「意図的に未テスト」へ列挙した。
'
' グループ / 本数 / 根拠章:
'   G0  ランナー自己テスト(W0から継続)      4本  17章§4-1
'   G1  SetCellSafe純ロジック(SanitizeForCell)  9本  16章 E-46 / NFR-S7(1) / E-22
'   G2  SanitizeInput                           6本  15章§0原則9 / 16章 E-04
'   G3  NormalizeForHash系(Fnv1a64Hex)          2本  14章§6 / 16章 E-49
'   G4  ExtractJsonBlock 入力パターン表        11本  14章§5(7パターン表+前処理P0)
'   G5  modJsonLite アクセサ                   13本  14章§6 / 13章§2.2
'   (G6〜G11は modTestsPure2.bas を参照)
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
GDone:
    On Error GoTo 0
    modTestsPure2.RunAll
    modTestsPure3.RunAll
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
