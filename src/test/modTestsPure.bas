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
'   PoC固有のテスト本体(modChunker / modPii / 難読化キー復元 TestDeobfuscateSecret
'   など)は移植していない。とくに **TestDeobfuscateSecret 系はOBF1機構ごと廃止**
'   したため意図的に持ち込まない(16章NFR-S2: 本製品はキーをブックに入れない)。
'
' W0時点の中身:
'   src/core・src/app のモジュールはまだ1本も実装されていないため、ここには
'   「テストハーネス自身の契約」を固定する最小の自己テストだけを置く。
'   W1以降、移植したモジュールのテストをここ(および modTestsPure2..n)へ足す。
'   **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**
'   (更新漏れは modTestRunner が即FAILとして可視化する。17章§4-1)。
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
' ============================================================================

Public Sub RunAll()
    On Error GoTo RunnerFail
    TestRunnerContract
NextDone:
    On Error GoTo 0
    Exit Sub

RunnerFail:
    modTestRunner.Check "TestRunnerContract(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
End Sub

' ----------------------------------------------------------------------------
' テストハーネス自身の契約を固定する(17章§4-1 ランナー要件の回帰)。
' ここが崩れると「テストを書いても走らない/落ちても緑に見える」状態になり、
' 以降の全テストの信頼性が消えるので、最初の1本はここに置く。
' ----------------------------------------------------------------------------
Private Sub TestRunnerContract()
    ' 1) このグループが最初に走る時点で、失敗は1件も記録されていないはず。
    modTestRunner.Check "ランナー: 開始時点でFAILが0件", _
        (modTestRunner.Failures() = 0), _
        "実際=" & modTestRunner.Failures()

    ' 2) レポートの1行目は "PASS n / FAIL m / SKIP k" で始まる。
    '    tools/run_lo_tests.py がこの書式を正規表現で読むため、崩れると
    '    集計そのものが読めなくなり「合格側へ倒れる」危険がある。
    Dim firstFive As String
    firstFive = Left(modTestRunner.ReportText(), 5)
    modTestRunner.Check "ランナー: ReportTextの先頭がPASS集計", _
        (firstFive = "PASS "), "実際=" & firstFive

    ' 3) 成功のCheckはFAILを増やさない(集計の基本不変条件)。
    Dim failsBefore As Long
    failsBefore = modTestRunner.Failures()
    modTestRunner.Check "ランナー: 成功のCheckは失敗に数えない(被験体)", True, ""
    modTestRunner.Check "ランナー: 成功のCheckはFAILを増やさない", _
        (modTestRunner.Failures() = failsBefore), _
        "前=" & failsBefore & " 後=" & modTestRunner.Failures()
End Sub
