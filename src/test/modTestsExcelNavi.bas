Attribute VB_Name = "modTestsExcelNavi"
Option Explicit

' Spec 8.3/10: 実セルを使う層(b)。ネットワークもフォームの起動も使わない。
' 裁定書34 §1.5(W12-A)で7本目(NAVI-B7)を足した。modNaviState.BuildAppState は
'   シート・ナレッジ・ThisWorkbook を触るので純関数ではなく、層(a)では書けない。
Private Const FIXTURE As String = "C-97990101-990"

Public Function RunExcelTestsNavi() As Long
    Dim n As Long
    Dim ready As Boolean
    Dim passed As Boolean
    ready = FixtureIsFree()
    For n = 1 To 8
        passed = False
        If ready Then passed = CheckFixture(n)
        modTestRunner.Check "NAVI-B" & CStr(n) & " " & TestTitle(n), passed, _
            "専用テストID=" & FIXTURE & " (既存データがある場合は変更せずFAIL)"
    Next n
    If ready Then CleanupFixture
    RunExcelTestsNavi = 8
End Function

Private Function TestTitle(ByVal n As Long) As String
    Select Case n
        Case 1: TestTitle = "chat_u timestamp and escaped text round trip"
        Case 2: TestTitle = "chat_a 32001 characters split round trip"
        Case 3: TestTitle = "nav_basics JSON round trip"
        Case 4: TestTitle = "edited Step JSON has display precedence"
        Case 5: TestTitle = "leading equals is stored as text"
        Case 6: TestTitle = "ui preference persists and restores"
        Case 7: TestTitle = "app_display_name reaches the initialize response"
        Case 8: TestTitle = "ActPasteMaterial honors pii_paste_policy(warn continues/block stops. 裁定書44 B-1)"
    End Select
End Function

Private Function FixtureIsFree() As Boolean
    Dim ws As Object
    Dim block As Variant
    Dim r As Long
    Dim c As Long
    On Error GoTo Failed
    Set ws = modCaseStore2.SheetOf("case_data")
    If ws Is Nothing Then Exit Function
    block = modCaseStore2.ReadBlock(ws, modCaseStore2.LastRowOf(ws))
    c = modUtil.FindHeaderCol(block, "case_id")
    If c = 0 Then Exit Function
    For r = 2 To modCaseStore2.LastRowOf(ws)
        If CStr(block(r, c)) = FIXTURE Then Exit Function
    Next r
    FixtureIsFree = True
    Exit Function
Failed:
    FixtureIsFree = False
End Function

Private Function CheckFixture(ByVal n As Long) As Boolean
    Dim srcValue As String
    Dim dataKey As String
    On Error GoTo Failed
    Select Case n
        Case 1
            dataKey = "chat_u"
            srcValue = "1" & vbTab & "2026-09-08 10:00:00" & vbTab & "質問\n二行目"
        Case 2
            dataKey = "chat_a"
            srcValue = String$(32000, "A") & "終"
        Case 3
            dataKey = "nav_basics"
            srcValue = "{""address"":""東京都千代田区"",""copied_at"":{}}"
        Case 4
            If Not modCaseStore.SaveData(FIXTURE, "s1_json", "{""source"":""base""}") Then Exit Function
            If Not modCaseStore.SaveData(FIXTURE, "s1_edited", "{""source"":""edited""}") Then Exit Function
            CheckFixture = modCaseStore.ResolveStepJson(FIXTURE, 1) = "{""source"":""edited""}"
            Exit Function
        Case 5
            dataKey = "chat_a"
            srcValue = "=1+1"
        Case 6
            CheckFixture = CheckSettingsRoundTrip()
            Exit Function
        Case 7
            CheckFixture = CheckDisplayNameInState()
            Exit Function
        Case 8
            CheckFixture = CheckPiiPastePolicyWiring()
            Exit Function
    End Select
    If Not modCaseStore.SaveData(FIXTURE, dataKey, srcValue) Then Exit Function
    CheckFixture = (modCaseStore.LoadData(FIXTURE, dataKey) = srcValue)
    If n = 5 Then CheckFixture = CheckFixture And NoFormulaInFixture()
    Exit Function
Failed:
    CheckFixture = False
End Function

Private Function NoFormulaInFixture() As Boolean
    Dim ws As Object
    Dim block As Variant
    Dim cId As Long
    Dim cText As Long
    Dim r As Long
    On Error GoTo Failed
    Set ws = modCaseStore2.SheetOf("case_data")
    block = modCaseStore2.ReadBlock(ws, modCaseStore2.LastRowOf(ws))
    cId = modUtil.FindHeaderCol(block, "case_id")
    cText = modUtil.FindHeaderCol(block, "content")
    For r = 2 To modCaseStore2.LastRowOf(ws)
        If CStr(block(r, cId)) = FIXTURE Then
            If ws.Cells(r, cText).HasFormula Then Exit Function
        End If
    Next r
    NoFormulaInFixture = True
    Exit Function
Failed:
    NoFormulaInFixture = False
End Function

Private Function CheckSettingsRoundTrip() As Boolean
    Dim before As String
    Dim saved As Boolean
    Dim restored As Boolean
    On Error GoTo Failed
    before = modConfig.GetStr("ui_font_scale", "medium")
    saved = modConfig.SetValue("ui_font_scale", "large")
    saved = saved And (modConfig.GetStr("ui_font_scale", vbNullString) = "large")
    restored = modConfig.SetValue("ui_font_scale", before)
    CheckSettingsRoundTrip = saved And restored
    Exit Function
Failed:
    RestoreSetting before
End Function

' 裁定書34 §0.4/§1.5: 画面へ出す製品名の値源は config app_display_name の1箇所で、
'   HTML は initialize 応答の state.app.display_name を読む。ここでは
'   (1) modBootNavi.AppDisplayName が config を読むこと
'   (2) その値が BuildAppState の app ブロックに載ること
'   の2つを、**実在しない印**へ一時的に差し替えて確かめる(識別子は
'   identity.display_name にも同じキー名があるため、値で見分ける)。
'   終わったら必ず元へ戻す。
Private Function CheckDisplayNameInState() As Boolean
    Dim before As String
    Dim marker As String
    Dim stateJson As String
    Dim seen As Boolean
    marker = "T47検査用表示名"
    On Error GoTo Failed
    before = modConfig.GetStr("app_display_name", "リスク提案ナビ")
    If Not modConfig.SetValue("app_display_name", marker) Then Exit Function
    If modBootNavi.AppDisplayName() <> marker Then GoTo Restore0
    stateJson = modNaviState.BuildAppState(vbNullString)
    seen = (InStr(1, stateJson, Chr$(34) & "display_name" & Chr$(34) & ":" & _
                  Chr$(34) & marker & Chr$(34), vbBinaryCompare) > 0)
    CheckDisplayNameInState = seen
Restore0:
    RestoreDisplayName before
    Exit Function
Failed:
    RestoreDisplayName before
    CheckDisplayNameInState = False
End Function

' 裁定書44 B-1: modNaviActions.ActPasteMaterial の配線(pii_paste_policyの
'   読取〜ブロック可否)を実Excelで確かめる(config読取・EnsureCase・ログはR4の
'   純関数では叩けない)。FIXTUREは「案件一覧」に登録していないIDなので、
'   EnsureCaseは常にFalse=Failure("", "E0101")で止まる。これはPIIチェックの
'   **あと**なので、E0103で止まる(=block)かE0101まで進む(=warn継続)かで
'   configの効きを見分けられる(案件は作らないのでCleanupFixture不要)。
Private Function CheckPiiPastePolicyWiring() As Boolean
    Dim piiText As String, dataJson As String, before As String
    Dim cidWarn As String, cidBlock As String, respWarn As String, respBlock As String
    Dim warnOk As Boolean, blockOk As Boolean
    On Error GoTo Failed
    piiText = "田中様と佐藤様に09012345678までご連絡ください。"
    dataJson = "{""slot"":""hp"",""text"":" & modNaviJson.Q(piiText) & "}"
    before = modConfig.GetStr("pii_paste_policy", "warn")

    If Not modConfig.SetValue("pii_paste_policy", "warn") Then GoTo Restore1
    cidWarn = FIXTURE
    respWarn = modNaviActions.ActPasteMaterial(cidWarn, dataJson)
    warnOk = (InStr(1, respWarn, "E0103", vbBinaryCompare) = 0)

    If Not modConfig.SetValue("pii_paste_policy", "block") Then GoTo Restore1
    cidBlock = FIXTURE
    respBlock = modNaviActions.ActPasteMaterial(cidBlock, dataJson)
    blockOk = (InStr(1, respBlock, "E0103", vbBinaryCompare) > 0) And _
              (InStr(1, respBlock, "人名らしき語", vbBinaryCompare) > 0)

    CheckPiiPastePolicyWiring = warnOk And blockOk
Restore1:
    modConfig.SetValue "pii_paste_policy", before
    Exit Function
Failed:
    modConfig.SetValue "pii_paste_policy", before
    CheckPiiPastePolicyWiring = False
End Function

Private Sub RestoreDisplayName(ByVal srcValue As String)
    On Error Resume Next
    If LenB(srcValue) > 0 Then modConfig.SetValue "app_display_name", srcValue
End Sub

Private Sub RestoreSetting(ByVal srcValue As String)
    On Error Resume Next
    modConfig.SetValue "ui_font_scale", srcValue
End Sub

Private Sub CleanupFixture()
    Dim ws As Object
    Dim block As Variant
    Dim c As Long
    Dim r As Long
    On Error Resume Next
    Set ws = modCaseStore2.SheetOf("case_data")
    If ws Is Nothing Then Exit Sub
    block = modCaseStore2.ReadBlock(ws, modCaseStore2.LastRowOf(ws))
    c = modUtil.FindHeaderCol(block, "case_id")
    If c = 0 Then Exit Sub
    For r = modCaseStore2.LastRowOf(ws) To 2 Step -1
        If CStr(block(r, c)) = FIXTURE Then ws.Rows(r).Delete
    Next r
End Sub
