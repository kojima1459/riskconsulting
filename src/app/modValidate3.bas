Attribute VB_Name = "modValidate3"
Option Explicit

' ============================================================================
' modValidate3 - S1の「落とさない警告」検査(15章§2 V-S1-14 / V-S1-15)
' ----------------------------------------------------------------------------
' 12章§2: modValidate(残328字)・modValidate2(残352字)が30,000字契約で満杯の
'   ため新設した(裁定書38 班A・B班報告 §6 C-3)。
'
' 責務は2件だけである。
'   V-S1-14(B-04): sources[].url が**貼付原文**に実在するか(InStr)。
'   V-S1-15(B-12): 接頭辞・出所の不整合(「(見立て)」の欠落・financials の
'                  出所と値の食い違い)。
'
' **戻り値を modValidate.CheckS1 に混ぜない**: CheckS1 の戻り値は
'   modPipeline.Defend の errText となり、非空なら修復リトライと
'   validate_result=failed を引き起こす。この2件は15章§11で「警告」であり、
'   出力を落としてはならないので、呼び出しは modPipeline3.DefendNotes(S1成功時)
'   からの**注記経路**だけとする(run_log の detail と HTMLレポートの
'   meta.s1_warn へ印を残す)。15章§2 の注記が正。
'
' 照合に使う「貼付原文」は modPipeline3.BuildHaystack(caseId, "") である
'   (第2引数に s1Json を渡さない。S1の出力を混ぜると、捏造したURLが自分自身と
'   一致して検査が無意味になる)。
'
' R4(12章§2): Excelトークン・config・シートに触れない純文字列モジュール。
' CP932準拠(15章§0 原則7): 本文・注釈ともに CP932 内の文字だけで書く。
' ============================================================================

' 注記の区切り(modPipeline.AddNote / modGround と同じ ";" 規約)。
Private Const V3_SEP As String = ";"

' 検査するケースIDの並び(WarnNoteOf の出力順。追番のみ・番号は再利用しない)。
Private Const V3_CASE_IDS As String = "V-S1-14|V-S1-15"

' URLの末尾に付きやすい句読点・括弧(文末の「。」や引用の「）」まで含んだ
'   URLを返してくる回があるため、照合の前に落とす)。半角空白・タブも含む。
Private Const V3_URL_TAIL As String = "。、．，.,;:)）」』】>＞ " & vbTab

' 「(見立て)」の接頭辞(15章§2 ルール9)。全角括弧で書く回もあるので両方見る。
Private Const V3_HEARSAY1 As String = "(見立て)"
Private Const V3_HEARSAY2 As String = "（見立て）"

' financials の「不明」(15章§2 ルール10)。
Private Const V3_UNKNOWN As String = "不明"

' ============================================================================
' CheckS1Notes - V-S1-14 / V-S1-15 の警告行(改行区切り)を返す。空=指摘なし。
'   各行は15章§0 原則10 のとおり "[ケースID] " で始まる。
'   haystack が空のときは V-S1-14 を**検査しない**(貼付が空のときに全件を
'   未照合で埋めない。modPipeline3.GroundHook と同じ fail-open)。
' ============================================================================
Public Function CheckS1Notes(ByVal json As String, ByVal haystack As String) As String
    Dim r As String

    r = UrlNotes(json, haystack)
    r = Join2(r, PrefixNotes(json))
    CheckS1Notes = r
End Function

' --- V-S1-14: sources[].url の貼付原文実在 ---------------------------------
Private Function UrlNotes(ByVal json As String, ByVal haystack As String) As String
    Dim it As Variant, idx As Long, urlText As String, r As String

    If LenB(Trim$(haystack)) = 0 Then Exit Function

    idx = 0
    For Each it In modJsonLite.GetArrayItems(json, "sources")
        urlText = TrimUrl(modJsonLite.GetStr(CStr(it), "url"))
        ' 空のURLは「原文に実在するURL」ではない(15章§2 ルール11は
        ' 「URLの無い資料は sources に入れない」と定める)ので警告に載せる。
        If LenB(urlText) = 0 Or InStr(1, haystack, urlText, vbBinaryCompare) = 0 Then
            r = Join2(r, "[V-S1-14] sources[" & idx & "].url が貼付原文に見当たりません: " & urlText)
        End If
        idx = idx + 1
    Next it
    UrlNotes = r
End Function

' --- V-S1-15: 接頭辞・出所の不整合 -----------------------------------------
Private Function PrefixNotes(ByVal json As String) As String
    Dim it As Variant, idx As Long, oneText As String, r As String
    Dim srcText As String

    idx = 0
    For Each it In modJsonLite.GetArrayItems(json, "current_coverage")
        oneText = CStr(it)
        If LCase$(Trim$(modJsonLite.GetStr(oneText, "certainty"))) = "assumed" Then
            If InStr(1, oneText, V3_HEARSAY1, vbBinaryCompare) = 0 _
               And InStr(1, oneText, V3_HEARSAY2, vbBinaryCompare) = 0 Then
                r = Join2(r, "[V-S1-15] 接頭辞・出所の不整合があります: current_coverage[" & _
                             idx & "] は certainty=assumed ですが「(見立て)」がありません")
            End If
        End If
        idx = idx + 1
    Next it

    srcText = LCase$(Trim$(modJsonLite.GetStr(json, "source")))
    If LenB(srcText) > 0 And srcText <> "unknown" Then
        If AllUnknown(json) Then
            r = Join2(r, "[V-S1-15] 接頭辞・出所の不整合があります: financials.source=" & _
                         srcText & " ですが4項目すべてが「不明」です")
        End If
    End If
    PrefixNotes = r
End Function

' financials の4項目がすべて「不明」か(前後の空白は落として完全一致で見る)。
Private Function AllUnknown(ByVal json As String) As Boolean
    AllUnknown = (Trim$(modJsonLite.GetStr(json, "fiscal_year")) = V3_UNKNOWN) _
             And (Trim$(modJsonLite.GetStr(json, "net_assets")) = V3_UNKNOWN) _
             And (Trim$(modJsonLite.GetStr(json, "sales")) = V3_UNKNOWN) _
             And (Trim$(modJsonLite.GetStr(json, "operating_profit")) = V3_UNKNOWN)
End Function

' ============================================================================
' WarnNoteOf - 警告行を run_log の detail 用の1語へ畳む。
'   "V-S1-14:2;V-S1-15:1" の形(0件のケースは出さない。全部0件なら空文字)。
'   並びは V3_CASE_IDS の順(出現順ではない=実行ごとに揺れない)。
' ============================================================================
Public Function WarnNoteOf(ByVal notesText As String) As String
    Dim ids() As String, i As Long, n As Long, r As String

    If LenB(notesText) = 0 Then Exit Function
    ids = Split(V3_CASE_IDS, "|")
    For i = LBound(ids) To UBound(ids)
        n = CountOf(notesText, "[" & ids(i) & "] ")
        If n > 0 Then
            If LenB(r) > 0 Then r = r & V3_SEP
            r = r & ids(i) & ":" & CStr(n)
        End If
    Next i
    WarnNoteOf = r
End Function

' TrimUrl - URLの前後の空白と、末尾の句読点・閉じ括弧を落とす。
'   (「...です(出典: https://example.com/a)。」のような書き方で括弧や句点まで
'    URLに含めてくる回があり、そのままでは原文と一致しない)。
Public Function TrimUrl(ByVal rawUrl As String) As String
    Dim t As String

    t = Trim$(rawUrl)
    Do While LenB(t) > 0
        If InStr(1, V3_URL_TAIL, Right$(t, 1), vbBinaryCompare) = 0 Then Exit Do
        t = Left$(t, Len(t) - 1)
    Loop
    TrimUrl = t
End Function

' 改行区切りの積み上げ(空は足さない)。
Private Function Join2(ByVal acc As String, ByVal oneText As String) As String
    Join2 = acc
    If LenB(oneText) = 0 Then Exit Function
    If LenB(Join2) > 0 Then Join2 = Join2 & vbLf
    Join2 = Join2 & oneText
End Function

' 部分文字列の出現件数。
Private Function CountOf(ByVal hay As String, ByVal needle As String) As Long
    Dim p As Long, n As Long

    If LenB(needle) = 0 Then Exit Function
    p = InStr(1, hay, needle, vbBinaryCompare)
    Do While p > 0
        n = n + 1
        p = InStr(p + Len(needle), hay, needle, vbBinaryCompare)
    Loop
    CountOf = n
End Function
