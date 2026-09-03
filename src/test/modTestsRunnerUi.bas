Attribute VB_Name = "modTestsRunnerUi"
Option Explicit

' ============================================================================
' modTestsRunnerUi - ブック内テスト実行(17章 T-48・裁定書14 裁定5)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   社内PCではターミナル・PowerShell・Pythonが使えないため、
'   wintest/run_excel_tests.ps1 と同じ検問を**ブックの中だけ**で回せる口が要る。
'   ps1(開発側の正)は存続させ、本モジュールは**利用者向けの同等検問**として
'   24章の正になる。
'
' 合否の4条件(ps1と同一。どれか1つでも欠ければNG):
'   (1) FAIL 0件  (2) SKIP 0件
'   (3) 純層(modTestsPure*)の実行本数 = 期待本数(config!tests_expected)
'   (4) 層(b)(modTestsExcel*)が1本以上走っている
'
' 実行順(modTestsExcel.RunAllExcelTests は ResetTests を呼ばない設計であり、
'   純層の**後**に呼んで結果を積み増す前提で書かれている。この順序を守る):
'   SetExpectedCount -> RunAllPureTests -> 純層の実行本数を控える
'   -> RunAllExcelTests -> 4条件を判定 -> 使い方タブへ書込 -> MsgBoxで1行
'
' fail-closed(裁定書14 裁定5):
'   期待本数(config!tests_expected)が空・非数値なら**テストを実行せず**NGで終える。
'   「読めなかったから全緑」という抜け道を作らない。
'
' 本モジュールは test層だが、結果の表示と確認ダイアログのためにExcelへ触れる
'   (modTestRunner / modTestsPure* の純ロジック規律は一切変えていない)。
' ============================================================================

' 期待本数の値源(裁定書27 W9-A)。配布方式Bでは隠しシート vba_src そのものが
' 無くなったため、旧 vba_src!E2 から config シートの `tests_expected` 行へ移した。
' config は A=name / B=value の縦持ちなので、A列を name で走査して B を読む
' (行番号をコードへ焼かない。13章§2.3)。
Private Const TR_SHEET_CFG As String = "config"
Private Const TR_EXPECTED_KEY As String = "tests_expected"
Private Const TR_CFG_MAX_ROWS As Long = 400
Private Const TR_RESULT_NAME As String = "gd_test_result"
Private Const TR_RESULT_ROWS As Long = 20
Private Const TR_TITLE As String = "リスク提案ナビ 自己テスト"

Private Const TR_CONFIRM As String = _
    "テストを実行します。数分かかり、実行中は操作できません。" & _
    "記録シート（err_log）に検査の行が残ります。よろしいですか。"
Private Const TR_NO_EXPECTED As String = _
    "期待本数が読めませんでした。開発担当へご連絡ください。"

' ============================================================================
' RunAllTestsFromBook - 使い方タブの[テストを実行]の OnAction。
'   引数なしなので Alt+F8(マクロ一覧)からも実行できる。
' ============================================================================
Public Sub RunAllTestsFromBook()
    On Error GoTo Failed

    If MsgBox(TR_CONFIRM, vbYesNo + vbQuestion, TR_TITLE) <> vbYes Then Exit Sub

    Dim expected As Long
    expected = ExpectedCount()
    If expected <= 0 Then
        WriteResult TR_NO_EXPECTED, vbNullString
        MsgBox TR_NO_EXPECTED, vbExclamation, TR_TITLE
        Exit Sub
    End If

    modTestRunner.SetExpectedCount expected
    modTestRunner.RunAllPureTests

    Dim pureExecuted As Long
    pureExecuted = modTestRunner.ExecutedCount()

    modTestsExcel.RunAllExcelTests

    Dim excelExecuted As Long
    excelExecuted = modTestRunner.ExecutedCount() - pureExecuted

    Dim summary As String
    summary = SummaryText(expected, pureExecuted, excelExecuted)

    WriteResult summary, modTestRunner.ReportText()
    MsgBox summary, vbInformation, TR_TITLE
    Exit Sub

Failed:
    NoticeAbort Err.Number
End Sub

' 実行中に想定外のエラーが出たときの後始末(ハンドラ稼働中に On Error を重ねない
' ため、別Subへ切り出す)。
Private Sub NoticeAbort(ByVal errNumber As Long)
    On Error Resume Next
    Dim errText As String
    errText = "テストの実行中に問題が起きました（" & CStr(errNumber) & "）。" & _
              "開発担当へご連絡ください。"
    WriteResult errText, vbNullString
    MsgBox errText, vbExclamation, TR_TITLE
End Sub

' ============================================================================
' 合否サマリ1行(ps1と同じ4条件)。
' ============================================================================
Private Function SummaryText(ByVal expected As Long, ByVal pureExecuted As Long, _
                             ByVal excelExecuted As Long) As String
    Dim failN As Long
    Dim skipN As Long
    failN = modTestRunner.FailCount()
    skipN = modTestRunner.SkipCount()

    Dim detail As String
    detail = "FAIL " & CStr(failN) & " / SKIP " & CStr(skipN) & _
             " / 純層 " & CStr(pureExecuted) & " = 期待 " & CStr(expected) & _
             " / 層(b) " & CStr(excelExecuted) & " 本"

    If failN = 0 And skipN = 0 And pureExecuted = expected And excelExecuted >= 1 Then
        SummaryText = "全PASS（" & detail & "）"
    Else
        SummaryText = "NG（" & detail & "）"
    End If
End Function

' ============================================================================
' 期待本数(config!tests_expected)。空・非数値・0以下は 0 を返す(=fail-closed)。
'   config シートが無い/キーが無い/読めない、はすべて 0(=テストを実行しない)。
'   「読めなかったから全緑」の抜け道を作らない(裁定書14 裁定5)。
' ============================================================================
Private Function ExpectedCount() As Long
    On Error GoTo Failed

    Dim ws As Object
    Set ws = modUISheet.SheetOf(TR_SHEET_CFG)
    If ws Is Nothing Then Exit Function

    Dim raw As String
    Dim r As Long
    For r = 2 To TR_CFG_MAX_ROWS
        If Trim$(CStr(ws.Cells(r, 1).Value)) = TR_EXPECTED_KEY Then
            raw = Trim$(CStr(ws.Cells(r, 2).Value))
            Exit For
        End If
    Next r
    If LenB(raw) = 0 Then Exit Function
    If Not IsNumeric(raw) Then Exit Function

    Dim n As Long
    n = CLng(Val(raw))
    If n <= 0 Then Exit Function
    ExpectedCount = n
    Exit Function

Failed:
    ExpectedCount = 0
End Function

' ============================================================================
' 結果の書込(使い方タブの gd_test_result・縦20行)。
'   1行目 = 合否サマリ。2行目以降 = ReportText の先頭から。
'   FAILがあるときは FAIL行(先頭が "NG: ")を先に並べる(20行で切れても
'   落ちた検査が読める)。あふれた分は最終行を「(以下省略・残りN行)」にする。
' ============================================================================
Private Sub WriteResult(ByVal summary As String, ByVal reportText As String)
    On Error Resume Next

    Dim cell As Object
    Set cell = modUISheet.NamedCell(TR_RESULT_NAME)
    If cell Is Nothing Then Exit Sub

    Dim ws As Object
    Set ws = cell.Worksheet
    If ws Is Nothing Then Exit Sub

    Dim baseRow As Long
    Dim baseCol As Long
    baseRow = cell.row
    baseCol = cell.Column

    Dim lines() As String
    lines = ResultLines(summary, reportText)

    Dim i As Long
    For i = 0 To TR_RESULT_ROWS - 1
        If i <= UBound(lines) Then
            modUtilText.SetCellSafe ws.Cells(baseRow + i, baseCol), lines(i), _
                                    "modTestsRunnerUi/" & TR_RESULT_NAME
        Else
            ws.Cells(baseRow + i, baseCol).ClearContents
        End If
    Next i
End Sub

' 書き込む行を組み立てる(1行目=サマリ / FAIL行を優先 / 20行超は打切り表示)。
Private Function ResultLines(ByVal summary As String, ByVal reportText As String) As String()
    Dim src() As String
    Dim outLines() As String
    ReDim outLines(0 To TR_RESULT_ROWS - 1)

    outLines(0) = summary

    Dim n As Long
    n = 1
    If LenB(reportText) > 0 Then
        src = Split(reportText, vbLf)

        ' (1) FAIL行を先に。
        Dim i As Long
        For i = LBound(src) To UBound(src)
            If n >= TR_RESULT_ROWS Then Exit For
            If InStr(1, src(i), "NG: ", vbBinaryCompare) = 1 Then
                outLines(n) = src(i)
                n = n + 1
            End If
        Next i

        ' (2) 残りを先頭から(FAIL行は重複させない)。
        For i = LBound(src) To UBound(src)
            If n >= TR_RESULT_ROWS Then Exit For
            If InStr(1, src(i), "NG: ", vbBinaryCompare) <> 1 Then
                outLines(n) = src(i)
                n = n + 1
            End If
        Next i

        ' (3) 収まらなかった行数を最終行に出す。最終行は打切り表示で**置換**される
        '     ため、そこにあった本文1行も欠落数に数える(数えないと1行少なく申告
        '     することになる)。
        '     例: 本文25行なら 1行目=サマリ+本文19行を置いた時点で n=20、
        '         rest=25-19=6。最終行が打切り表示に化けて本文は18行しか残らない
        '         ので、申告は 25-18=7 行が正しい。
        Dim rest As Long
        rest = (UBound(src) - LBound(src) + 1) - (n - 1)
        If rest > 0 Then
            rest = rest + 1
            outLines(TR_RESULT_ROWS - 1) = "（以下省略・残り" & CStr(rest) & "行）"
        End If
    End If

    ResultLines = outLines
End Function
