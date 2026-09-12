Attribute VB_Name = "modNaviState"
Option Explicit
' Specification 7.4: the active case is read through existing stores, never sheet selection.
Private gDraft As String
Private gSparringCase As String
Private gSparringNote As String
Public Sub SetDraft(ByVal json As String)
    If Not modNaviJson.IsValidJson(json) Then Err.Raise 5, "modNaviState", "invalid_draft"
    gDraft = json
End Sub
Public Function DraftJson() As String
    If Len(gDraft) = 0 Then gDraft = "{}"
    DraftJson = gDraft
End Function
Public Function BuildAppState(ByVal caseId As String) As String
    Dim buf() As String, count As Long, dataDir As String, kb As Boolean
    Dim menus As Long, schemes As Long, examples As Long, industry As String, note As String
    Dim route As String
    industry = "00"
    If Len(caseId) > 0 Then industry = modCaseRead.CaseColumnOf(caseId, "industry_code")
    menus = modPipeline.KbRowCount(modKnowledge.MenusFor(industry, 0))
    schemes = modPipeline.KbRowCount(modKnowledge.SchemesFor(industry, 0))
    examples = modPipeline.KbRowCount(modKnowledge.CasesFor(industry, 0))
    kb = (menus + schemes + examples > 0)
    If kb Then
        note = modBoot.KbAutoNote() & "読込OK（メニュー" & CStr(menus) & "/型" & CStr(schemes) & "/事例" & CStr(examples) & "）"
    Else
        note = "ナレッジブック.xlsx を本体と同じフォルダーに置いて[ナレッジを読み直す]を押してください。"
    End If
    dataDir = modUtil.ResolveDataDir(modConfig.GetStr("data_dir", ""), ThisWorkbook.Path)
    route = modGatewayRPN.ResolveTransport(modConfig.GetBool("mock_llm", False), modConfig.GetStr("llm_transport", "ribbon"))
    modUtil.BufInit buf, count
    ' 裁定書34 §0.4/§1.2: 画面へ出す製品名の値源は config app_display_name の1箇所。
    '   HTML 側は state.app.display_name を読んで題名と見出しに使う。
    modUtil.BufAdd buf, count, "{""app"":{""version"":""0.3.0"",""display_name"":" & modNaviJson.Q(modBootNavi.AppDisplayName()) & ",""transport"":" & modNaviJson.Q(route) & ",""model"":" & modNaviJson.Q(modConfig.GetStr("recommended_model", "gpt-5.5")) & ","
    modUtil.BufAdd buf, count, """ribbon_available"":" & modNaviJson.Flag(modGatewayRPN.RibbonAvailable()) & ",""data_dir"":" & modNaviJson.Q(dataDir) & ","
    modUtil.BufAdd buf, count, """data_dir_last_resort"":" & modNaviJson.Flag(modUtil.DataDirIsLastResort(dataDir, ThisWorkbook.Path)) & ","
    modUtil.BufAdd buf, count, """kb_status"":" & modNaviJson.Q(note) & ",""kb_ready"":" & modNaviJson.Flag(kb) & ",""ui_locked"":" & modNaviJson.Flag(modUIProgress.IsUiLocked()) & ","
    modUtil.BufAdd buf, count, """max_wait_sec"":" & CStr(modConfig.GetLong("llm_wait_sec", 1200)) & ",""advanced"":" & modNaviJson.Flag(modConfig.GetBool("ui_advanced", False)) & "},"
    modUtil.BufAdd buf, count, """identity"":{""display_name"":" & modNaviJson.Q(modConfig.GetStr("operator", modCaseStore2.OwnerName())) & ",""windows_user"":" & modNaviJson.Q(Environ$("USERNAME")) & "},"
    modUtil.BufAdd buf, count, """enums"":" & modNaviState2.BuildEnums() & ",""industries"":" & modNaviState2.BuildIndustries() & ","
    modUtil.BufAdd buf, count, """prompt_templates"":" & BuildPrompts("", "{}") & ",""cases"":" & modNaviStore.ListCases() & ","
    modUtil.BufAdd buf, count, """current"":" & BuildCaseState(caseId, kb) & ",""settings"":" & SettingsJson() & "}"
    BuildAppState = modUtil.BufText(buf, count)
End Function
Private Function SettingsJson() As String
    SettingsJson = "{""ui_mode"":" & modNaviJson.Q(modConfig.GetStr("ui_mode", "sheet")) & _
        ",""ui_font_scale"":" & modNaviJson.Q(modConfig.GetStr("ui_font_scale", "medium")) & _
        ",""chat_include_materials"":" & modNaviJson.Flag(modConfig.GetBool("chat_include_materials", False)) & "}"
End Function
Public Function BuildCaseState(ByVal caseId As String, Optional ByVal kbReady As Boolean) As String
    Dim ctx As String, basics As String, company As String, status As String, materials As String
    Dim buf() As String, count As Long, n As Long, stepNo As Long, rule As Long, anyArea As Boolean
    Dim raw As String, failure As String, report As String, reportHtml As String
    If Len(caseId) = 0 Then
        ctx = DraftJson(): basics = DraftJson()
    Else
        ctx = modNaviState2.CaseContext(caseId)
        basics = modCaseStore.LoadData(caseId, "nav_basics")
        If Not modNaviJson.IsValidJson(basics) Then basics = "{}"
    End If
    company = modJsonLite.GetStr(ctx, "company")
    status = modJsonLite.GetStr(ctx, "status")
    materials = modNaviState2.MaterialsJson(caseId, company, anyArea)
    rule = modUINav.StepRuleOf(kbReady, modUIProgress.IsUiLocked(), Len(company) > 0, anyArea, status)
    stepNo = modUINav.StepFor(kbReady, modUIProgress.IsUiLocked(), Len(company) > 0, anyArea, status)
    modUtil.BufInit buf, count
    modUtil.BufAdd buf, count, "{""case_id"":" & modNaviJson.Q(caseId) & ",""ctx"":" & ctx & ",""basics"":" & basics & ","
    modUtil.BufAdd buf, count, """prompts"":" & BuildPrompts(company, basics, modJsonLite.GetStr(ctx, "industry_name")) & ",""materials"":" & materials & ","
    modUtil.BufAdd buf, count, """step"":{""no"":" & CStr(stepNo) & ",""next_action"":" & modNaviJson.Q(modNaviHost.DisplayMessage(modUINav.StepActionOf(rule))) & _
        ",""text"":" & modNaviJson.Q(modNaviHost.DisplayMessage(modUINav.StepText(stepNo))) & "},"
    modUtil.BufAdd buf, count, """stages"":" & BuildStageList(caseId) & ","
    failure = "{"
    For n = 1 To 4
        raw = ""
        If Len(caseId) > 0 Then raw = modCaseStore.ResolveStepJson(caseId, n)
        modUtil.BufAdd buf, count, modNaviJson.Q("s" & CStr(n)) & ":" & modNaviJson.Q(modUICase.RestoreNames(raw, company)) & ","
        If n > 1 Then failure = failure & ","
        raw = ""
        If Len(caseId) > 0 Then raw = modCaseStore.LoadData(caseId, "s" & CStr(n) & "_json_failed")
        failure = failure & modNaviJson.Q("s" & CStr(n)) & ":" & modNaviJson.Q(modUICase.RestoreNames(raw, company))
    Next n
    modUtil.BufAdd buf, count, """failed"":" & failure & "},""s2_prev"":null,"
    If Len(caseId) = 0 Then
        modUtil.BufAdd buf, count, """chat"":[],""sparring"":[],""runs"":[],""logs"":{},""feedback"":[],""judgements"":[],"
    Else
        modUtil.BufAdd buf, count, """chat"":" & modNaviChat.History(caseId) & ",""sparring"":" & modNaviState2.SparringJson(caseId, company) & ","
        modUtil.BufAdd buf, count, """runs"":" & modNaviStore.RowsJson("run_log", caseId) & ",""logs"":" & modNaviStore.LogRowsOf(caseId) & ","
        modUtil.BufAdd buf, count, """feedback"":" & modNaviStore.RowsJson("フィードバック", caseId) & ",""judgements"":" & modNaviStore.ListJudgements(caseId) & ","
        report = modCaseRead.CaseColumnOf(caseId, "report_path")
        ' The path originates in the business store, never from an HTML path parameter.
        If Len(report) > 0 Then
            If modUtil.FileExistsAt(report) Then reportHtml = modUtil.ReadUtf8File(report)
        End If
    End If
    modUtil.BufAdd buf, count, """sparring_note"":" & modNaviJson.Q(SparringNote(caseId)) & ","
    modUtil.BufAdd buf, count, """outputs"":{""report_path"":" & modNaviJson.Q(report) & ",""report_html"":" & modNaviJson.Q(reportHtml) & _
        ",""hearing_built_at"":" & modNaviJson.Q(modNaviState2.HearingBuiltAt(caseId)) & "}}"
    BuildCaseState = modUtil.BufText(buf, count)
End Function
Public Function BuildPrompts(ByVal company As String, ByVal basics As String, Optional ByVal industryName As String) As String
    Dim n As Long, template As String, srcText As String, result As String, titles As Variant
    Dim industry As String, limitChars As Long, chars As Long
    titles = Array("1本目 会社の基本", "2本目 リスクの兆候", "3本目 調達・仕入れの構造", "業界と競合", "世の中の動きとの関係", "前回の更新からの変化", "拠点の災害リスク", "決算のハイライト")
    industry = industryName
    If Len(industry) = 0 Then industry = modJsonLite.GetStr(basics, "industry_name")
    ' 裁定書37 B-09: 展開後の字数(chars)と上限超過(over)を各プロンプトへ足す。
    ' 上限は config dr_input_max_chars(既定2,000)。判定は純関数 IsOverDrLimit に閉じる。
    limitChars = modConfig.GetLong("dr_input_max_chars", 2000)
    result = "["
    For n = 1 To 8
        template = modUISheet.ReadNamed("gd_prompt_" & Format$(n, "00"))
        srcText = modUIResearch.FillTemplate(template, company, modJsonLite.GetStr(basics, "address"), _
            industry, modJsonLite.GetStr(basics, "sec_code"), modJsonLite.GetStr(basics, "sites"))
        chars = Len(srcText)
        If n > 1 Then result = result & ","
        result = result & "{""no"":" & CStr(n) & ",""title"":" & modNaviJson.Q(CStr(titles(n - 1))) & _
            ",""template"":" & modNaviJson.Q(template) & ",""text"":" & modNaviJson.Q(srcText) & _
            ",""chars"":" & CStr(chars) & ",""over"":" & modNaviJson.Flag(IsOverDrLimit(chars, limitChars)) & _
            ",""copied_at"":" & modNaviJson.Q(modJsonLite.GetStr(basics, "copied_" & CStr(n))) & "}"
    Next n
    BuildPrompts = result & "]"
End Function

' 裁定書37 B-09: 展開後の字数が dr_input_max_chars を超えたか(純関数・副作用なし)。
' modTestsPureNavi.bas の NAVI-P33〜P35(1999/2000/2001の3点)が正。
Public Function IsOverDrLimit(ByVal chars As Long, ByVal limitChars As Long) As Boolean
    IsOverDrLimit = (chars > limitChars)
End Function
Public Function BuildStageList(ByVal caseId As String) As String
    Dim n As Long, result As String, state As String, stageName As String, runs As Collection
    Dim row As Variant, v As String, latency As Long, failedStep As String
    If Len(caseId) > 0 Then
        Set runs = modJsonLite.GetArrayItems("{""items"":" & modNaviStore.RowsJson("run_log", caseId) & "}", "items")
        failedStep = modCaseRead.CaseColumnOf(caseId, "failed_step")
    End If
    result = "["
    For n = 1 To 4
        state = "pending": v = "": latency = 0
        If Len(caseId) > 0 Then
            If Len(modCaseStore.ResolveStepJson(caseId, n)) > 0 Then state = "done"
            If failedStep = CStr(n) Or failedStep = "s" & CStr(n) Then state = "failed"
        End If
        If Not runs Is Nothing Then
            For Each row In runs
                If modJsonLite.GetStr(CStr(row), "step") = "s" & CStr(n) Then
                    v = modJsonLite.GetStr(CStr(row), "validate_result")
                    latency = modJsonLite.GetLong(CStr(row), "latency_ms", 0)
                End If
            Next row
        End If
        Select Case n
        Case 1: stageName = "企業プロファイル分析"
        Case 2: stageName = "リスク分析"
        Case 3: stageName = "提案分析"
        Case 4: stageName = "骨子作成"
        End Select
        If n > 1 Then result = result & ","
        result = result & "{""no"":" & CStr(n) & ",""name"":" & modNaviJson.Q(stageName) & ",""state"":" & modNaviJson.Q(state) & _
            ",""validate"":" & modNaviJson.Q(v) & ",""latency_ms"":" & CStr(latency) & "}"
    Next n
    BuildStageList = result & "]"
End Function

Public Sub SetSparringNote(ByVal caseId As String, ByVal note As String)
    gSparringCase = caseId
    gSparringNote = note
End Sub
Public Function SparringNote(ByVal caseId As String) As String
    If caseId = gSparringCase Then SparringNote = gSparringNote
End Function
