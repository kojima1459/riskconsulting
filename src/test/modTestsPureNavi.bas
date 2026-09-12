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
    Exit Sub
Failed:
    modTestRunner.Check "NAVI pure unexpected error", False, CStr(Err.Number) & " " & Err.Description
End Sub

Private Sub CheckN(ByVal title As String, ByVal passed As Boolean)
    modTestRunner.Check title, passed
End Sub

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
