Attribute VB_Name = "modTestsPureNavi"
Option Explicit

' Spec 7.2/7.5/8.3/10. 24 assertions; baseline 780 -> tests_expected 804.
' 裁定書36で RibbonHead/InWindow を追加(+8 assertions -> tests_expected 812)。
' 裁定書37 B-09で IsOverDrLimit(NAVI-P33〜P35)を追加(+3 assertions)。
' 裁定書39 R2-01/R2-06/R2-08/R1-06 で NAVI-P36〜P49(+14 assertions)。
' 裁定書40 R-M2/R-m1/R-m3 で NAVI-P50〜P61(+12 assertions)。
' 裁定書40 R-M2 の通し固定で NAVI-P68〜P70(+3 assertions)。
' 裁定書40 の横展開(同型)で NAVI-P62〜P67(+6 assertions)。
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
    ' 裁定書40 R-M2 / R-m3: [提案書（お客さま向け）を出す]の4分岐を、出る側と
    ' 出ない側の**両方向**で固定する。ActExportProposal はこの3本の純関数
    ' (ProposalBlockOf / ProposalStepsOf / ProposalS5FailedJson)だけで判断と
    ' 順序を決めるので、分岐を書き換えるとここが落ちる。
    CheckN "NAVI-P50 an unreviewed case is refused with E0603", _
        modNaviActions2.ProposalBlockOf(modExportProposal.NeedsReviewMessage(vbNullString), "export") = _
        "{""ok"":false,""message"":" & modNaviJson.Q(modExportProposal.NeedsReviewMessage(vbNullString)) & _
        ",""error_code"":""E0603""}"
    CheckN "NAVI-P51 a reviewed case is not refused", _
        modNaviActions2.ProposalBlockOf(modExportProposal.NeedsReviewMessage("山田太郎"), "export") = vbNullString And _
        modNaviActions2.ProposalBlockOf(vbNullString, "run_s5") = vbNullString
    CheckN "NAVI-P52 upstream_missing asks for the analysis first", _
        modNaviActions2.ProposalBlockOf(vbNullString, "upstream_missing") = _
        "{""ok"":false,""message"":""先に[まとめて分析]を実行してください。"",""error_code"":""E0101""}"
    CheckN "NAVI-P53 the review check wins over upstream_missing", _
        modNaviActions2.ProposalBlockOf("内容を確認してから出力してください。", "upstream_missing") = _
        "{""ok"":false,""message"":""内容を確認してから出力してください。"",""error_code"":""E0603""}"
    CheckN "NAVI-P54 run_s5 builds S5 before the export", _
        modNaviState.ProposalStepsOf("run_s5") = "run_s5;export"
    CheckN "NAVI-P55 an existing S5 is not rebuilt", _
        modNaviState.ProposalStepsOf("export") = "export"
    CheckN "NAVI-P56 a blocked plan runs no step at all", _
        modNaviState.ProposalStepsOf("upstream_missing") = vbNullString And _
        modNaviState.ProposalStepsOf(vbNullString) = vbNullString
    CheckN "NAVI-P57 a failed S5 returns the 16 E-71 guidance", _
        modNaviActions2.ProposalS5FailedJson() = "{""ok"":false,""message"":""提案書を作れませんでした。" & _
        "今回は提案骨子（4. 提案の骨子）をご利用ください。"",""error_code"":""E0302""}"
    ' 裁定書40 R-m1: 提案書の保存先は案件データ(nav_basics)に残る=ブックを
    ' 開き直しても区画④の出力一覧に出る。画面から来た JSON では上書きできない。
    CheckN "NAVI-P58 proposal_path survives in nav_basics", _
        modNaviState.ProposalPathIn("{""proposal_path"":" & modNaviJson.Q(NaviPathSample()) & "}") = NaviPathSample()
    CheckN "NAVI-P59 a case without an export has no proposal_path", _
        modNaviState.ProposalPathIn("{""address"":""東京都""}") = vbNullString And _
        modNaviState.ProposalPathIn(vbNullString) = vbNullString
    CheckN "NAVI-P60 saving the company form keeps the stored proposal_path", _
        modJsonLite.GetStr(modNaviStore.MergeBasics("{""proposal_path"":" & modNaviJson.Q(NaviPathSample()) & "}", _
            "{""address"":""大阪市"",""proposal_path"":""C:\\evil\\x.html""}"), "proposal_path") = NaviPathSample()
    CheckN "NAVI-P61 only a new export overwrites the stored proposal_path", _
        modJsonLite.GetStr(modNaviStore.MergeBasics("{""proposal_path"":" & modNaviJson.Q(NaviPathSample()) & "}", _
            "{}", NaviPathSample() & "2"), "proposal_path") = NaviPathSample() & "2"
    ' 裁定書40 R-M2: 案件データ -> 段取り -> 手順の並び を**つないだまま**固定する
    ' (途中の1本だけを見ていると、つなぎ目を書き換えても誰も気づけない)。
    CheckN "NAVI-P68 a case without S5 builds it first, then exports", _
        modNaviState.ProposalStepsOf(modNaviState.ProposalPlanOf(vbNullString, vbNullString, _
            "{""s"":1}", "{""s"":2}", "{""s"":3}")) = "run_s5;export"
    CheckN "NAVI-P69 a case that already has S5 only exports", _
        modNaviState.ProposalStepsOf(modNaviState.ProposalPlanOf("{""a"":1}", vbNullString, _
            "{""s"":1}", "{""s"":2}", "{""s"":3}")) = "export"
    CheckN "NAVI-P70 an unanalysed case is stopped with the guidance", _
        modNaviState.ProposalStepsOf(modNaviState.ProposalPlanOf(vbNullString, vbNullString, _
            "{""s"":1}", vbNullString, "{""s"":3}")) = vbNullString And _
        modNaviActions2.ProposalBlockOf(vbNullString, modNaviState.ProposalPlanOf(vbNullString, _
            vbNullString, "{""s"":1}", vbNullString, "{""s"":3}")) = _
        "{""ok"":false,""message"":""先に[まとめて分析]を実行してください。"",""error_code"":""E0101""}"
    ' 裁定書40 の横展開(同型): R2-06 の防波堤(画面から来たパスは開かない)と
    ' R1-06 の走査上限も、純関数へ出して両方向で固定する。
    CheckN "NAVI-P62 the proposal row opens the stored proposal", _
        modNaviActions2.OpenTargetOf("R.html", NaviPathSample(), NaviPathSample()) = NaviPathSample()
    CheckN "NAVI-P63 a path the case never produced is refused", _
        modNaviActions2.OpenTargetOf("R.html", NaviPathSample(), "C:\\windows\\x.html") = vbNullString And _
        modNaviActions2.OpenTargetOf(vbNullString, vbNullString, NaviPathSample()) = vbNullString
    CheckN "NAVI-P64 the report row opens the report", _
        modNaviActions2.OpenTargetOf("R.html", vbNullString, "R.html") = "R.html" And _
        modNaviActions2.OpenTargetOf("R.html", NaviPathSample(), vbNullString) = "R.html"
    CheckN "NAVI-P65 no report means nothing to open", _
        modNaviActions2.OpenTargetOf(vbNullString, vbNullString, vbNullString) = vbNullString
    CheckN "NAVI-P66 the copy pre-scan source is capped at 30000 chars", _
        Len(modNaviActions2.CapSourceText(String$(30001, "x"))) = 30000
    CheckN "NAVI-P67 a short source is not cut", _
        modNaviActions2.CapSourceText("abc") = "abc" And _
        Len(modNaviActions2.CapSourceText(String$(30000, "x"))) = 30000
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

' 提案書の保存先テスト用の実パス相当(区切り文字が JSON のエスケープを通ることも
' あわせて見る。裁定書40 R-m1)。
Private Function NaviPathSample() As String
    NaviPathSample = "D:\データ\提案書_20260912.html"
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
