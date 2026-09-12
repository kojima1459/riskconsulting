Attribute VB_Name = "modTestRunner"
Option Explicit

' ============================================================================
' modTestRunner - 純ロジックテストの実行エンジン(17章§4-1 ランナー要件)
' ----------------------------------------------------------------------------
' 役割:
'   modTestsPure1..n が呼ぶ「Check」を受け止め、合否を集計してレポート文字列を
'   作る、テストの土台。ExcelでもLibreOffice(ヘッドレス)でも同じコードで動く
'   ことが最重要なモジュールなので、他のどのモジュールより厳格にR4(純ロジック)
'   を守る。
'
' 移植元: PoC「マイ本棚AI」 src/test/modTestRunner.bas。
'   PoCの SKIP別集計(SKIPをPASSにも総本数にも数えない)はそのまま維持し、
'   17章§4-1が追加要求する次の2点を新設した。
'     (2) 期待テスト総本数 tests_expected と実行本数の一致検査
'     (3) SKIPが1件でもあればレポートの冒頭と末尾に理由込みで全件列挙
'
' 設計判断:
'   ・R4準拠: Worksheets/Range/Application/ThisWorkbook/MsgBox に一切触れない。
'     結果はすべてPrivateなモジュール変数に蓄積し、ReportText が文字列として
'     返す。呼び出し側(Excelなら診断シート、LOならファイル書き出し)が好きな
'     方法で表示すればよい。
'   ・改行は vbLf で統一(LOのStarBasicはvbCrLfも解釈できるが、テキスト比較や
'     LO側での表示ズレを避けるため vbLf に統一する)。
'   ・tests_expected の実値は wintest/tests_expected.txt(1行目に10進整数のみ)
'     が正。VBA側からファイル位置を解決するとブックのパスに依存してしまうため、
'     読み取りは実行ハーネス(tools/run_lo_tests.py / wintest/run_excel_tests.ps1)
'     の責務とし、ここへは SetExpectedCount で渡してもらう。未設定のまま
'     RunAllPureTests を回した場合は「0件実行の全緑」を成立させないため
'     ランナー自身がFAILを1件立てる。
' ============================================================================

' ----------------------------------------------------------------------------
' SKIP集計
'   実行環境の制限で本体を1行も実行できないテスト群が、PoCでは
'   `Check "...スキップ", True` として【PASSに計上】されていた。レポートには
'   「実行されなかった」という情報が一切残らず、PASS件数が水増しされていた。
'   呼び出し側は  Check "[SKIP] modXxx: 理由", True, "詳細"  と書く。
'   接頭辞つき かつ cond=True のものは PASS にも総数にも数えず、SKIP として
'   別に数えて ReportText の冒頭と末尾の一覧に出す。
' ----------------------------------------------------------------------------
Private Const SKIP_PREFIX As String = "[SKIP]"

Private mTotalCount As Long     ' Check呼び出し総数(=実行本数。SKIPは含めない)
Private mFailCount As Long      ' 失敗数
Private mFailLines() As String  ' 失敗の詳細(1件1行)
Private mFailCap As Long        ' mFailLinesの現在容量
Private mSkipCount As Long      ' 未実行(SKIP)数
Private mSkipLines() As String  ' 未実行の一覧(1件1行)
Private mSkipCap As Long        ' mSkipLinesの現在容量
Private mStarted As Boolean     ' ResetTests済みかどうか(未Reset時の誤集計防止)
Private mExpected As Long       ' 期待テスト総本数(tests_expected)
Private mExpectedSet As Boolean ' SetExpectedCountが呼ばれたか

' ----------------------------------------------------------------------------
' ResetTests: 集計をゼロから始める
'   mExpected / mExpectedSet はここでは消さない。ハーネスは
'   SetExpectedCount -> RunAllPureTests の順で呼ぶため、消すと期待値が失われる。
' ----------------------------------------------------------------------------
Public Sub ResetTests()
    mTotalCount = 0
    mFailCount = 0
    mFailCap = 16
    ReDim mFailLines(1 To mFailCap)
    mSkipCount = 0
    mSkipCap = 16
    ReDim mSkipLines(1 To mSkipCap)
    mStarted = True
End Sub

' ----------------------------------------------------------------------------
' SetExpectedCount: 期待テスト総本数(wintest/tests_expected.txt の値)を渡す
' ----------------------------------------------------------------------------
Public Sub SetExpectedCount(ByVal n As Long)
    mExpected = n
    mExpectedSet = True
End Sub

' ----------------------------------------------------------------------------
' Check: 1件のテスト結果を記録する
'   testName : テスト名(検証ルール表のケースIDを含めること。17章§4-2)
'   cond     : True=成功 / False=失敗
'   detail   : 失敗時に添える補足(期待値/実際値など)。省略可
' ----------------------------------------------------------------------------
Public Sub Check(ByVal testName As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
    If Not mStarted Then ResetTests

    ' 未実行の申告([SKIP]接頭辞)は PASS にも総数にも数えない。
    ' cond=False で来た場合だけは本物の失敗として下へ流す(接頭辞を付けた
    ' まま落ちるテストを黙って隠さないため)。
    If cond Then
        If Left(testName, Len(SKIP_PREFIX)) = SKIP_PREFIX Then
            mSkipCount = mSkipCount + 1
            If mSkipCount > mSkipCap Then
                mSkipCap = mSkipCap * 2
                ReDim Preserve mSkipLines(1 To mSkipCap)
            End If
            mSkipLines(mSkipCount) = "SKIP: " & Trim(Mid(testName, Len(SKIP_PREFIX) + 1))
            If Len(detail) > 0 Then
                mSkipLines(mSkipCount) = mSkipLines(mSkipCount) & " -- " & detail
            End If
            Exit Sub
        End If
    End If

    mTotalCount = mTotalCount + 1
    If Not cond Then AddFailure "NG: " & testName, detail
End Sub

' ----------------------------------------------------------------------------
' AddFailure: 失敗を1件積む(実行本数 mTotalCount は増やさない)
'   ランナー自身の失敗(tests_expected 不一致など)を、テスト本数の水増しを
'   起こさずに記録するための内部口。
' ----------------------------------------------------------------------------
Private Sub AddFailure(ByVal headLine As String, ByVal detail As String)
    mFailCount = mFailCount + 1
    If mFailCount > mFailCap Then
        mFailCap = mFailCap * 2
        ReDim Preserve mFailLines(1 To mFailCap)
    End If

    Dim s As String
    s = headLine
    If Len(detail) > 0 Then s = s & " -- " & detail
    mFailLines(mFailCount) = s
End Sub

' ----------------------------------------------------------------------------
' Failures: 現在までの失敗件数
' ----------------------------------------------------------------------------
Public Function Failures() As Long
    Failures = mFailCount
End Function

' ----------------------------------------------------------------------------
' 計数の公開アクセサ(裁定書14 裁定5)
'   ブック内テスト実行(modTestsRunnerUi)が ps1 と同じ4条件(FAIL 0 / SKIP 0 /
'   純層の実行本数=期待 / 層(b) 1本以上)をVBA側で判定するための読み出し口。
'   **読み出すだけ**であり集計の仕方は変えない(R4=純ロジックのまま。
'   LibreOffice実行テストへの影響も無い)。
' ----------------------------------------------------------------------------
Public Function PassCount() As Long
    PassCount = mTotalCount - mFailCount
End Function

Public Function FailCount() As Long
    FailCount = mFailCount
End Function

Public Function SkipCount() As Long
    SkipCount = mSkipCount
End Function

Public Function ExecutedCount() As Long
    ExecutedCount = mTotalCount
End Function

' ----------------------------------------------------------------------------
' ReportText: 集計行に続けて未実行一覧・失敗一覧・未実行一覧(再掲)を返す
'   例:
'     PASS 41 / FAIL 1 / SKIP 2
'     EXECUTED 42 / EXPECTED 42
'     --- 未実行(SKIP) 2件 [冒頭] ---
'     SKIP: modXxx: LO環境の既知の制限により未実行
'     ...
'     NG: V-S2-03_件数下限 -- 期待=4件 実際=3件
'     --- 未実行(SKIP) 2件 [末尾] ---
'     ...
'   1行目の書式は tools/run_lo_tests.py が正規表現で読むので変更しないこと。
' ----------------------------------------------------------------------------
Public Function ReportText() As String
    If Not mStarted Then ResetTests

    Dim passN As Long
    passN = mTotalCount - mFailCount

    Dim expectedText As String
    If mExpectedSet Then
        expectedText = CStr(mExpected)
    Else
        expectedText = "(未設定)"
    End If

    Dim n As Long
    n = 2 + mFailCount
    If mSkipCount > 0 Then n = n + (mSkipCount + 1) * 2

    Dim parts() As String
    ReDim parts(0 To n - 1)

    Dim k As Long
    parts(0) = "PASS " & passN & " / FAIL " & mFailCount & " / SKIP " & mSkipCount
    parts(1) = "EXECUTED " & mTotalCount & " / EXPECTED " & expectedText
    k = 2

    Dim i As Long
    If mSkipCount > 0 Then
        parts(k) = "--- 未実行(SKIP) " & mSkipCount & "件 [冒頭] ---"
        k = k + 1
        For i = 1 To mSkipCount
            parts(k) = mSkipLines(i)
            k = k + 1
        Next i
    End If

    For i = 1 To mFailCount
        parts(k) = mFailLines(i)
        k = k + 1
    Next i

    If mSkipCount > 0 Then
        parts(k) = "--- 未実行(SKIP) " & mSkipCount & "件 [末尾] ---"
        k = k + 1
        For i = 1 To mSkipCount
            parts(k) = mSkipLines(i)
            k = k + 1
        Next i
    End If

    ReportText = Join(parts, vbLf)
End Function

' ----------------------------------------------------------------------------
' RunAllPureTests: 各 modTestsPure* の Run系を列挙呼び出しする。
'   呼び出し先が未実装/未注入/実行時エラーでも本Subは落ちず、その事実を
'   1件のテスト失敗として可視化する(「無かったことにする」のではなく
'   「失敗として見せる」)。
'   最後に実行本数と tests_expected を照合し、ズレていればランナー自身が
'   FAILを立てる(17章§4-1 ランナー要件(2))。
' ----------------------------------------------------------------------------
Public Sub RunAllPureTests()
    ResetTests

    On Error Resume Next
    Err.Clear
    modTestsPure.RunAll
    If Err.Number <> 0 Then
        Check "modTestsPure.RunAll", False, "呼び出しでエラー: " & Err.Description & _
              " (Err=" & Err.Number & ") ※未実装/未注入の可能性"
        Err.Clear
    End If
    On Error GoTo 0

    ' dev専用の純層テスト(裁定書30 裁定1(d))。配布ビルドでは modTestsPureHook が
    ' 何も実行しない版に差し替わるので、ここは prod/dev のどちらでも同じ1行で
    ' 済む(実行時の名前ディスパッチは使わない)。
    On Error Resume Next
    Err.Clear
    modTestsPureHook.RunAll
    If Err.Number <> 0 Then
        Check "modTestsPureHook.RunAll", False, _
              "呼び出しでエラー: " & Err.Description & " (Err=" & Err.Number & ")"
        Err.Clear
    End If
    On Error GoTo 0

    ' HTML画面(裁定書34 W12-A)の純層。modTestsPure* の連鎖とは別に、navi の
    ' 純関数(JSONの検証と組立・action許可・data_key)を叩く24本を持つ。
    On Error Resume Next
    Err.Clear
    modTestsPureNavi.RunAll
    If Err.Number <> 0 Then
        Check "modTestsPureNavi.RunAll", False, _
              "呼び出しでエラー: " & Err.Description & " (Err=" & Err.Number & ")"
        Err.Clear
    End If
    On Error GoTo 0

    ' W14(裁定書37 班2)の純層。modTestsPure* の数珠つなぎとは別に呼ぶ
    ' (班1の modTestsPure23 と同じ行を奪い合わないため。裁定書37 §2)。
    On Error Resume Next
    Err.Clear
    modTestsPure24.RunAll
    If Err.Number <> 0 Then
        Check "modTestsPure24.RunAll", False, _
              "呼び出しでエラー: " & Err.Description & " (Err=" & Err.Number & ")"
        Err.Clear
    End If
    On Error GoTo 0

    ' W15(裁定書38 班C)の純層。CheckS5・提案書テンプレ・確認必須・対訳表。
    ' modTestsPure* の数珠つなぎとは別に呼ぶ(他班と同じ行を奪い合わないため)。
    On Error Resume Next
    Err.Clear
    modTestsPure27.RunAll
    If Err.Number <> 0 Then
        Check "modTestsPure27.RunAll", False, _
              "呼び出しでエラー: " & Err.Description & " (Err=" & Err.Number & ")"
        Err.Clear
    End If
    On Error GoTo 0

    ' ---- 実行本数の照合(0件実行の「全緑」を成立させない) ----
    If Not mExpectedSet Then
        AddFailure "NG: tests_expected が未設定です", _
                   "wintest/tests_expected.txt を読み、RunAllPureTests の前に " & _
                   "modTestRunner.SetExpectedCount を呼んでください(17章§4-1)"
    ElseIf mTotalCount <> mExpected Then
        AddFailure "NG: 実行本数が tests_expected と不一致", _
                   "期待=" & mExpected & " 実際=" & mTotalCount & _
                   " (テストを増減したら wintest/tests_expected.txt を更新すること)"
    End If
End Sub
