Attribute VB_Name = "modTestsPureNavi"
Option Explicit

' Spec 7.2/7.5/8.3/10. 24 assertions; baseline 780 -> tests_expected 804.
' 裁定書36で RibbonHead/InWindow を追加(+8 assertions -> tests_expected 812)。
' 裁定書37 B-09で IsOverDrLimit(NAVI-P33〜P35)を追加(+3 assertions)。
Public Sub RunAll()
    On Error GoTo Failed
    CheckN "NAVI-P01 empty object", modNaviJson.IsValidJson("{}")
    CheckN "NAVI-P02 nested value", modNaviJson.IsValidJson("{""a"":[1,true,null,{""b"":""日本語""}]}")
    CheckN "NAVI-P03 truncated object rejected", Not modNaviJson.IsValidJson("{""a"":1")
    CheckN "NAVI-P04 trailing comma rejected", Not modNaviJson.IsValidJson("{""a"":1,}")
    CheckN "NAVI-P05 multiple root rejected", Not modNaviJson.IsValidJson("{} {}")
    CheckN "NAVI-P06 raw newline rejected", Not modNaviJson.IsValidJson("{""a"":""a" & vbLf & "b""}")
    CheckN "NAVI-P07 duplicate key rejected", Not modNaviJson.IsValidJson("{""a"":1,""a"":2}")
    CheckN "NAVI-P08 invalid escape rejected", Not modNaviJson.IsValidJson("{""a"":""\q""}")
    CheckN "NAVI-P09 empty JSON string", modNaviJson.Q(vbNullString) = Chr$(34) & Chr$(34)
    CheckN "NAVI-P10 quotes round trip", RoundTrip("A""B")
    CheckN "NAVI-P11 newlines round trip", RoundTrip("A" & vbCrLf & "B")
    CheckN "NAVI-P12 slash round trip", RoundTrip("C:\資料\会社")
    CheckN "NAVI-P13 closing script round trip", RoundTrip("</script><script>alert(1)</script>")
    CheckN "NAVI-P14 nested field", modNaviJson.RawField("{""data"":{""s1"":""{}""},""action"":""initialize""}", "action") = """initialize"""
    CheckN "NAVI-P15 nested object", modNaviJson.ObjectField("{""data"":{""a"":1}}", "data") = "{""a"":1}"
    CheckN "NAVI-P16 initialize allowed", modNaviHost.IsAllowed("initialize")
    CheckN "NAVI-P17 pipeline allowed", modNaviHost.IsAllowed("run_pipeline")
    CheckN "NAVI-P18 chat allowed", modNaviHost.IsAllowed("chat")
    CheckN "NAVI-P19 arbitrary macro rejected", Not modNaviHost.IsAllowed("Application.Run")
    CheckN "NAVI-P20 action case exact", Not modNaviHost.IsAllowed("INITIALIZE")
    CheckN "NAVI-P21 business vocabulary", modNaviHost.DisplayMessage("まとめて作る") = "まとめて分析"
    CheckN "NAVI-P22 data keys", KeyCountAndPresence()
    CheckN "NAVI-P23 setting whitelist", modConfig.IsSettingsKeyAllowed("ui_mode") And modConfig.IsSettingsKeyAllowed("ui_font_scale") And modConfig.IsSettingsKeyAllowed("chat_include_materials") And Not modConfig.IsSettingsKeyAllowed("ch_effort")
    CheckN "NAVI-P24 keep newest history", modGatewayRPN2.TrimHistoryPairs("new;;;middle;;;old", 2) = "new;;;middle"
    CheckN "NAVI-P25 RibbonHead trims normal text", modGatewayRPN2.RibbonHead(" (error:500) timeout ") = "(error:500) timeout"
    CheckN "NAVI-P26 RibbonHead collapses CR/LF/tab to space", _
        modGatewayRPN2.RibbonHead("a" & vbCr & "b" & vbLf & "c" & vbTab & "d") = "a b c d"
    CheckN "NAVI-P27 RibbonHead truncates over 80 chars", _
        modGatewayRPN2.RibbonHead(String$(90, "x")) = String$(80, "x")
    CheckN "NAVI-P28 RibbonHead empty for blank input rejected", _
        modGatewayRPN2.RibbonHead(vbNullString) = vbNullString And modGatewayRPN2.RibbonHead("   ") = vbNullString
    CheckN "NAVI-P29 InWindow inside window", _
        modNaviStore.InWindow("2026-01-01 10:00:00", "2026-01-01 09:59:00", "2026-01-01 10:01:00")
    CheckN "NAVI-P30 InWindow boundary exact both ends", _
        modNaviStore.InWindow("2026-01-01 09:59:00", "2026-01-01 09:59:00", "2026-01-01 10:01:00") And _
        modNaviStore.InWindow("2026-01-01 10:01:00", "2026-01-01 09:59:00", "2026-01-01 10:01:00")
    CheckN "NAVI-P31 InWindow one second outside rejected", _
        Not modNaviStore.InWindow("2026-01-01 09:58:59", "2026-01-01 09:59:00", "2026-01-01 10:01:00") And _
        Not modNaviStore.InWindow("2026-01-01 10:01:01", "2026-01-01 09:59:00", "2026-01-01 10:01:00")
    CheckN "NAVI-P32 InWindow invalid date rejected", _
        Not modNaviStore.InWindow("not-a-date", "2026-01-01 09:59:00", "2026-01-01 10:01:00") And _
        Not modNaviStore.InWindow("2026-01-01 10:00:00", vbNullString, "2026-01-01 10:01:00")
    CheckN "NAVI-P33 IsOverDrLimit under limit", Not modNaviState.IsOverDrLimit(1999, 2000)
    CheckN "NAVI-P34 IsOverDrLimit at limit is not over", Not modNaviState.IsOverDrLimit(2000, 2000)
    CheckN "NAVI-P35 IsOverDrLimit over limit", modNaviState.IsOverDrLimit(2001, 2000)
    ' 裁定書39 R2-01: 提案書は S5 の作成込みで長時間処理になるので、進捗表示と
    ' UIロックの対象(IsLongAction)に入れる。入っていないと押しても何も起きて
    ' いないように見え、二度押しで二重実行になる。
    CheckN "NAVI-P36 export_proposal is a long action", _
        modNaviActions.IsLongAction("export_proposal")
    CheckN "NAVI-P37 export_report stays short", _
        Not modNaviActions.IsLongAction("export_report")
    ' 裁定書39 R2-08: 案件が開けない致命エラーのときこそ報告が要るので、
    ' report_mail は案件必須から外す。ほかの出力は案件必須のまま。
    CheckN "NAVI-P38 report_mail works without a case", _
        Not modNaviActions.NeedsCase("report_mail")
    CheckN "NAVI-P39 other outputs still need a case", _
        modNaviActions.NeedsCase("export_proposal") And modNaviActions.NeedsCase("export_report")
    ' 裁定書39 R2-01: 1ボタンで完結させる段取りの判定(純関数)。
    CheckN "NAVI-P40 ProposalPlanOf exports when s5_json exists", _
        modNaviState.ProposalPlanOf("{""a"":1}", "", "{}", "{}", "{}") = "export"
    CheckN "NAVI-P41 ProposalPlanOf exports when only s5_edited exists", _
        modNaviState.ProposalPlanOf("", "{""a"":1}", "{}", "{}", "{}") = "export"
    CheckN "NAVI-P42 ProposalPlanOf runs S5 when upstream is complete", _
        modNaviState.ProposalPlanOf("", "", "{""s"":1}", "{""s"":2}", "{""s"":3}") = "run_s5"
    CheckN "NAVI-P43 ProposalPlanOf refuses when S2 is missing", _
        modNaviState.ProposalPlanOf("", "", "{""s"":1}", "", "{""s"":3}") = "upstream_missing"
    CheckN "NAVI-P44 ProposalPlanOf treats blank-only json as missing", _
        modNaviState.ProposalPlanOf("   ", "", "{""s"":1}", "{""s"":2}", "{""s"":3}") = "run_s5"
    ' 裁定書39 R2-06: 出力一覧は提案書の保存先も持つ。
    CheckN "NAVI-P45 OutputsJson carries proposal_path", _
        InStr(1, modNaviState.OutputsJson("R", "", "P", "H"), """proposal_path"":""P""", 0) > 0
    CheckN "NAVI-P46 OutputsJson keeps report and hearing", _
        InStr(1, modNaviState.OutputsJson("R", "", "P", "H"), """report_path"":""R""", 0) > 0 And _
        InStr(1, modNaviState.OutputsJson("R", "", "P", "H"), """hearing_built_at"":""H""", 0) > 0
    ' 裁定書39 R1-06 / 16章 E-69: [コピー]直前の下見(純関数)。
    CheckN "NAVI-P47 CopyWarningOf flags a 20-char overlap", _
        modNaviActions2.CopyWarningOf(Left$(NaviFragSample(), 20), NaviFragSample()) = _
        "現契約・営業メモと20字以上一致する記述が含まれています。"
    CheckN "NAVI-P48 CopyWarningOf ignores a 19-char overlap", _
        modNaviActions2.CopyWarningOf(Left$(NaviFragSample(), 19), NaviFragSample()) = vbNullString
    CheckN "NAVI-P49 CopyWarningOf is quiet without a source", _
        modNaviActions2.CopyWarningOf(Left$(NaviFragSample(), 20), vbNullString) = vbNullString
    Exit Sub
Failed:
    modTestRunner.Check "NAVI pure unexpected error", False, CStr(Err.Number) & " " & Err.Description
End Sub

Private Sub CheckN(ByVal title As String, ByVal passed As Boolean)
    modTestRunner.Check title, passed, vbNullString
End Sub

' 下見テスト用の下敷き(modTestsPure26 の SharesLongFragment と同じ素材)。
Private Function NaviFragSample() As String
    NaviFragSample = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
End Function

Private Function RoundTrip(ByVal srcValue As String) As Boolean
    Dim json As String
    json = "{""x"":" & modNaviJson.Q(srcValue) & "}"
    RoundTrip = modNaviJson.IsValidJson(json) And (modNaviJson.StringField(json, "x") = srcValue)
End Function

Private Function KeyCountAndPresence() As Boolean
    Dim keys As String
    keys = ";" & modCaseStore3.DataKeys() & ";"
    ' 13章§2.2 の data_key は 36値(v2.8・裁定書38 班A s1_json_prev / 班C s5_json・s5_edited・s5_json_failed)。
    KeyCountAndPresence = (UBound(Split(modCaseStore3.DataKeys(), ";")) = 35) And _
        InStr(1, keys, ";chat_u;", 0) > 0 And InStr(1, keys, ";chat_a;", 0) > 0 And _
        InStr(1, keys, ";nav_basics;", 0) > 0 And InStr(1, keys, ";s1_json_prev;", 0) > 0
End Function
