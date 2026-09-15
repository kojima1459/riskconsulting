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
    modUtil.BufAdd buf, count, """max_wait_sec"":" & CStr(modConfig.GetLong("llm_wait_sec", 1200)) & ",""advanced"":" & modNaviJson.Flag(modConfig.GetBool("ui_advanced", False)) & ","
    ' 裁定書39 R2-14: 貼付1本あたりの上限は config で変えられる。画面の
    ' 「N字を超えています」を焼き付けないよう、実値を state へ載せる。
    modUtil.BufAdd buf, count, """dr_input_max_chars"":" & CStr(modConfig.GetLong("dr_input_max_chars", 2000)) & ","
    modUtil.BufAdd buf, count, """feature_inbox"":" & modNaviJson.Flag(modConfig.GetBool("feature_inbox", True)) & ",""feature_judgelog"":" & modNaviJson.Flag(modConfig.GetBool("feature_judgelog", True)) & "},"
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
    Dim raw As String, failure As String, report As String, reportHtml As String, proposal As String
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
        ' 提案書の保存先は**案件が採番されているときだけ**読む(下書きのJSONは
        ' 画面から来るので、そこに紛れた proposal_path を区画④へ出さない)。
        proposal = ProposalPathIn(basics)
        ' The path originates in the business store, never from an HTML path parameter.
        If Len(report) > 0 Then
            If modUtil.FileExistsAt(report) Then reportHtml = modUtil.ReadUtf8File(report)
        End If
    End If
    modUtil.BufAdd buf, count, """sparring_note"":" & modNaviJson.Q(SparringNote(caseId)) & ","
    modUtil.BufAdd buf, count, """outputs"":" & OutputsJson(report, reportHtml, proposal, _
        modNaviState2.HearingBuiltAt(caseId)) & "}"
    BuildCaseState = modUtil.BufText(buf, count)
End Function

' OutputsJson - 区画④「出力」に出す成果物3件(レポート・提案書・ヒアリング)を
'   1つの JSON へ組む純関数(11章§3.8.2c・裁定書39 R2-06)。提案書のパスを
'   落とすと、出したファイルに利用者がたどり着けない(会社PCはエクスプローラ
'   操作が限られる。docs/24・29)。ui/views.js の outputList がこの3件を読む。
Public Function OutputsJson(ByVal reportPath As String, ByVal reportHtml As String, _
                            ByVal proposalPath As String, ByVal hearingAt As String) As String
    OutputsJson = "{""report_path"":" & modNaviJson.Q(reportPath) & _
        ",""report_html"":" & modNaviJson.Q(reportHtml) & _
        ",""proposal_path"":" & modNaviJson.Q(proposalPath) & _
        ",""hearing_built_at"":" & modNaviJson.Q(hearingAt) & "}"
End Function

' ProposalPlanOf - [提案書（お客さま向け）を出す]を押したときの段取りを決める
'   純関数(裁定書39 R2-01)。ボタンは1本のままにするため、S5(お客さま向け
'   提案書の文章)が無いときは**その場で作ってから**出力する。
'     "export"           : s5_json か s5_edited がある → そのまま出力する
'     "run_s5"           : S5 は無いが S1+S2+S3 がそろっている → 先に S5 を作る
'     "upstream_missing" : S1〜S3 のどれかが無い → [まとめて分析]が先
'   S4(骨子)は提案書の必須入力ではない(15章§5.6 の user は S1/S2/S3 と実数)。
Public Function ProposalPlanOf(ByVal s5Json As String, ByVal s5Edited As String, _
                               ByVal s1Json As String, ByVal s2Json As String, _
                               ByVal s3Json As String) As String
    If LenB(Trim$(s5Json)) > 0 Or LenB(Trim$(s5Edited)) > 0 Then
        ProposalPlanOf = "export"
        Exit Function
    End If
    If LenB(Trim$(s1Json)) = 0 Or LenB(Trim$(s2Json)) = 0 Or LenB(Trim$(s3Json)) = 0 Then
        ProposalPlanOf = "upstream_missing"
        Exit Function
    End If
    ProposalPlanOf = "run_s5"
End Function

' SetProposalPath / ProposalPathOf - 提案書の保存先(裁定書39 R2-06・裁定書40 R-m1)。
'   **案件データ(nav_basics)へ書いて残す**。画面のセッション変数に置くと
'   ブックを開き直した時点で消え、区画④の出力一覧が「未出力」に戻る
'   (会社PCはエクスプローラの操作が限られるので、画面に出さないと利用者は
'   自分が出したファイルへたどり着けない。docs/24・29)。値源は
'   ActExportProposal の成功時1箇所のまま。読み口は BuildCaseState(区画④の
'   出力一覧)と ActOpenReport(given=proposal)の2つ。
'   13章§2.2 の nav_basics の記載も同期すること(13章の改訂は別班の担当)。
Public Sub SetProposalPath(ByVal caseId As String, ByVal pathText As String)
    Dim basicsJson As String
    If LenB(caseId) = 0 Then Exit Sub
    basicsJson = modCaseStore.LoadData(caseId, "nav_basics")
    If Not modNaviJson.IsValidJson(basicsJson) Then basicsJson = "{}"
    modCaseStore.SaveData caseId, "nav_basics", _
        modNaviStore.MergeBasics(basicsJson, "{}", pathText)
End Sub
Public Function ProposalPathOf(ByVal caseId As String) As String
    If LenB(caseId) = 0 Then Exit Function
    ProposalPathOf = ProposalPathIn(modCaseStore.LoadData(caseId, "nav_basics"))
End Function

' ProposalPathIn - nav_basics(JSON)から提案書の保存先だけを取り出す純関数
'   (裁定書40 R-m1)。保存されていなければ空文字=区画④は「未出力」と出す。
Public Function ProposalPathIn(ByVal basicsJson As String) As String
    ProposalPathIn = modJsonLite.GetStr(basicsJson, "proposal_path")
End Function

' ProposalStepsOf - ProposalPlanOf が決めた段取りを、ActExportProposal が実際に
'   踏む手順の並び(";"区切り・前から順に実行)へ開く純関数(裁定書40 R-M2)。
'   **「S5 を作ってから出力する」という順序の正はここ1箇所**であり、
'   ActExportProposal はこの並びをそのまま回すだけなので、順序や本数を
'   書き換えると純テスト(NAVI-P54〜P56・P68〜P70)が落ちる。
'     "run_s5" -> "run_s5;export"(その場で S5 を作ってから出力する)
'     "export" -> "export"(S5 は作り直さない。AI呼出を1回むだにしない)
'     それ以外(upstream_missing など) -> ""(1手も実行しない)
Public Function ProposalStepsOf(ByVal plan As String) As String
    Select Case plan
    Case "run_s5"
        ProposalStepsOf = "run_s5;export"
    Case "export"
        ProposalStepsOf = "export"
    End Select
End Function
' BuildPrompts - 区画①の調査指示文8本を state へ組む。
'   裁定書38 Z-49 で各プロンプトへ "warning" を1本足した。**16章 E-69 の
'   (b)「現契約サマリ・営業メモと20字以上一致する断片」の下見はここには無い**:
'   裁定書39 R1-06 のとおり、BuildPrompts は画面更新のたびに呼ばれるのに対し、
'   (b) は最大30万字 × 8本の走査なので、コピーの直前だけで足りる。
'   実体は modNaviActions2.CopyWarningOf(copy_prompt の action)が持つ。
'   ここに残すのは (a) modPii 走査だけ(対象は展開後のプロンプト本文=数百字)。
Public Function BuildPrompts(ByVal company As String, ByVal basics As String, _
                              Optional ByVal industryName As String) As String
    Dim n As Long, template As String, srcText As String, result As String, titles As Variant
    Dim industry As String, limitChars As Long, chars As Long, warnText As String
    titles = Array("1本目 会社の基本", "2本目 リスクの兆候", "3本目 調達・仕入れの構造", "業界と競合", "世の中の動きとの関係", "前回の更新からの変化", "拠点の災害リスク", "決算のハイライト")
    industry = industryName
    If Len(industry) = 0 Then industry = modJsonLite.GetStr(basics, "industry_name")
    ' 裁定書37 B-09: 展開後の字数(chars)と上限超過(over)を各プロンプトへ足す。
    ' 上限は config dr_input_max_chars(既定2,000)。判定は純関数 IsOverDrLimit に閉じる。
    limitChars = modConfig.GetLong("dr_input_max_chars", 2000)
    result = "["
    For n = 1 To 8
        template = modUISheet.ReadNamed(PromptNameOf(n))
        srcText = FillPrompt(template, company, basics, industry)
        chars = Len(srcText)
        warnText = vbNullString
        If modPii.HasPii(srcText) Then
            warnText = "個人情報らしき記述が含まれています。"
        End If
        If n > 1 Then result = result & ","
        result = result & "{""no"":" & CStr(n) & ",""title"":" & modNaviJson.Q(CStr(titles(n - 1))) & _
            ",""template"":" & modNaviJson.Q(template) & ",""text"":" & modNaviJson.Q(srcText) & _
            ",""chars"":" & CStr(chars) & ",""over"":" & modNaviJson.Flag(IsOverDrLimit(chars, limitChars)) & _
            ",""warning"":" & modNaviJson.Q(warnText) & _
            ",""holes"":" & HolesJsonOf(modUIResearch.HolesOf(srcText)) & _
            ",""copied_at"":" & modNaviJson.Q(modJsonLite.GetStr(basics, "copied_" & CStr(n))) & "}"
    Next n
    BuildPrompts = result & "]"
End Function

' HolesJsonOf - 裁定書47 I-1(b): HolesOf の`;`区切りを文字列配列のJSONへ組む。
'   コピー前に見せる(未入力の穴の名前一覧)ため BuildPrompts が使う。
Private Function HolesJsonOf(ByVal holesText As String) As String
    If LenB(holesText) = 0 Then
        HolesJsonOf = "[]"
        Exit Function
    End If
    Dim names() As String, i As Long, result As String
    names = Split(holesText, ";")
    result = "["
    For i = LBound(names) To UBound(names)
        If i > LBound(names) Then result = result & ","
        result = result & modNaviJson.Q(names(i))
    Next i
    HolesJsonOf = result & "]"
End Function

' PromptNameOf / FillPrompt - 指示文1本の「名前定義」と「差し込み」。
'   BuildPrompts(8本まとめて)と PromptTextOf(1本だけ)の両方がここを通るので、
'   差し込む項目の並びは1箇所にしかない。
Private Function PromptNameOf(ByVal n As Long) As String
    PromptNameOf = "gd_prompt_" & Format$(n, "00")
End Function
Private Function FillPrompt(ByVal template As String, ByVal company As String, _
                            ByVal basics As String, ByVal industry As String) As String
    ' 裁定書47 I-1(a): 公式サイトURL・会社の規模・直近決算期も nav_basics から読む
    ' (会社情報フォームの新3項目。空ならFillTemplateが従来どおり穴を返す)。
    FillPrompt = modUIResearch.FillTemplate(template, company, modJsonLite.GetStr(basics, "address"), _
        industry, modJsonLite.GetStr(basics, "sec_code"), modJsonLite.GetStr(basics, "sites"), _
        modJsonLite.GetStr(basics, "official_url"), modJsonLite.GetStr(basics, "company_size"), _
        modJsonLite.GetStr(basics, "fiscal_term"))
End Function

' PromptTextOf - 指示文 n 本目の展開後の本文(裁定書39 R1-06)。
'   [コピー]の action(modNaviActions2.ActCopyPrompt)が、コピーされる1本だけを
'   下見するために使う。画面更新のために8本ぶん作り直すのは BuildPrompts。
Public Function PromptTextOf(ByVal n As Long, ByVal company As String, _
                             ByVal basics As String, ByVal industryName As String) As String
    Dim industry As String
    industry = industryName
    If Len(industry) = 0 Then industry = modJsonLite.GetStr(basics, "industry_name")
    PromptTextOf = FillPrompt(modUISheet.ReadNamed(PromptNameOf(n)), company, basics, industry)
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
