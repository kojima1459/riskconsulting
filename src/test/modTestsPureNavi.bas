Attribute VB_Name = "modTestsPureNavi"
Option Explicit

' Spec 7.2/7.5/8.3/10. 24 assertions; baseline 780 -> tests_expected 804.
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
    KeyCountAndPresence = (UBound(Split(modCaseStore3.DataKeys(), ";")) = 31) And _
        InStr(1, keys, ";chat_u;", 0) > 0 And InStr(1, keys, ";chat_a;", 0) > 0 And _
        InStr(1, keys, ";nav_basics;", 0) > 0
End Function
