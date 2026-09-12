Attribute VB_Name = "modValidate3"
Option Explicit

' ============================================================================
' modValidate3 - S1の「落とさない警告」検査(15章§2 V-S1-14 / V-S1-15)
' ----------------------------------------------------------------------------
' 12章§2: modValidate(残328字)・modValidate2(残352字)が30,000字契約で満杯の
'   ため新設した(裁定書38 班A・B班報告 §6 C-3)。
'
' 責務は次の4件である(2〜4件目は裁定書39 で追加)。
'   V-S1-14(B-04): sources[].url が**貼付原文**に実在するか(InStr)。
'   V-S1-15(B-12): 接頭辞・出所の不整合(「(見立て)」の欠落・financials の
'                  出所と値の食い違い)。
'   V-S1-16 / V-S1-17(R1-09 / X-1): sources の欠落と missing_info[].kind の
'                  enum 外。**この2件も他の警告と同じ注記チャネル**に出す
'                  (裁定書40 P-M1。当初は modValidate.CheckS1 の戻り値へ
'                  連結していたが、それは下の「混ぜない」規約に反していた)。
'   PostNormalize(R1-09 / 裁定書41 §2): S1 の正規化直後に sources の空配列と
'                  missing_info[].kind の既定値 not_found を補う fail-open。
'
' **戻り値を modValidate.CheckS1 に混ぜない**: CheckS1 の戻り値は
'   modPipeline.Defend の errText となり、非空なら修復リトライ1回と
'   PL_RES_FAILED -> FailStep(案件 status=error)を引き起こす。本モジュールの
'   4件は15章§11で「警告」であり、出力を落としてはならない。したがって出口は
'   CheckS1Notes 1本(呼ぶのは modPipeline3.S1Notes(S1成功時)と
'   modExportHtml)とし、run_log の detail と HTMLレポートの meta.s1_warn へ
'   印を残すだけにする。15章§2 の注記が正。
'
' **新設キーの欠落を不合格にしない(fail-open)**: sources(v2.7)も
'   missing_info[].kind(v2.7)も、実運用のリボン経路ではスキーマを強制できない
'   ためモデルが落としうる。落としたことを不合格・警告のどちらにしても案件が
'   前へ進まなくなる/警告が常時鳴って本物が埋もれるので、**欠落は既定値で
'   補って黙って続け、値が壊れているときだけ警告する**。補填はどちらも
'   PostNormalize の1本で行う(sources は空配列、missing_info[].kind の欠落と
'   空は "not_found")。**「みなす」だけでは足りない**(裁定書41 §2): JSON へ
'   実際に書かないと 18章 SEC-04 の「種別」列が modHtmlTemplate1.LB の引き当てに
'   失敗して**空欄**になり、利用者には種別が消えたようにしか見えない。
'
' 裁定書39 R1-10 / G-2: 値は**対象オブジェクトを切り出してから**読む。
'   json 全体へ GetStr(json,"source") を掛けると「最初に現れた同名キー」(15章§14)
'   を拾うためスキーマのキー順に依存し、生のJSON片へ InStr を掛けるとスキーマ外の
'   フィールドに現れた「(見立て)」で警告が消える。どちらもリボン経路(スキーマ
'   強制なし)で現実に起きる。
'
' 照合に使う「貼付原文」は modPipeline3.BuildHaystack(caseId, "") である
'   (第2引数に s1Json を渡さない。S1の出力を混ぜると、捏造したURLが自分自身と
'   一致して検査が無意味になる)。
'
' R4(12章§2): Excelトークン・config・シートに触れない純文字列モジュール。
' CP932準拠(15章§0 原則7): 本文・注釈ともに CP932 内の文字だけで書く。
' ============================================================================

' 注記の区切りは "," (run_log detail の項目区切り ";" と衝突させない。司令塔検収 W15)。
Private Const V3_SEP As String = ","

' 検査するケースIDの並び(WarnNoteOf の出力順。追番のみ・番号は再利用しない)。
'   裁定書40 P-M1 で V-S1-16 / V-S1-17 も注記チャネルへ移したので並びへ足す。
Private Const V3_CASE_IDS As String = "V-S1-14|V-S1-15|V-S1-16|V-S1-17"

' URLの末尾に付きやすい句読点・括弧(文末の「。」や引用の「）」まで含んだ
'   URLを返してくる回があるため、照合の前に落とす)。半角空白・タブも含む。
Private Const V3_URL_TAIL As String = "。、．，.,;:)）」』】>＞ " & vbTab

' 「(見立て)」の接頭辞(15章§2 ルール9)。全角括弧で書く回もあるので両方見る。
Private Const V3_HEARSAY1 As String = "(見立て)"
Private Const V3_HEARSAY2 As String = "（見立て）"

' financials の「不明」(15章§2 ルール10)。
Private Const V3_UNKNOWN As String = "不明"

' 裁定書39 G-2: 「(見立て)」を探す**対象の値**(Schema-S1 の current_coverage[])。
'   生のJSON片へ InStr を掛けると、スキーマ外のフィールド(リボン経路はスキーマを
'   強制しないので混ざりうる)に現れた「(見立て)」を拾って警告を握りつぶす。
Private Const V3_CC_VALUE_KEYS As String = "line_name|coverage_summary|limit_note|special_note"

' 裁定書39 R1-10: financials の4項目(AllUnknown の判定対象)。
Private Const V3_FIN_KEYS As String = "fiscal_year|net_assets|sales|operating_profit"

' 裁定書39 X-1: missing_info[].kind の enum(19章§3・Schema-S1。V-S1-17)。
Private Const V3_MI_KIND As String = "|conflict|undisclosed|not_found|hearing_only|"

' 空白とみなす文字(JSONのトークン間)。
Private Const V3_WS As String = " " & vbTab & vbCr & vbLf

' ============================================================================
' CheckS1Notes - V-S1-14 / V-S1-15 / V-S1-16 / V-S1-17 の警告行(改行区切り)を
'   返す。空=指摘なし。各行は15章§0 原則10 のとおり "[ケースID] " で始まる。
'   haystack が空のときは V-S1-14 を**検査しない**(貼付が空のときに全件を
'   未照合で埋めない。modPipeline3.GroundHook と同じ fail-open)。
'   V-S1-16 / V-S1-17 は haystack を使わない(貼付原文と無関係の構造検査)ので
'   haystack が空でも判定する。
' ============================================================================
Public Function CheckS1Notes(ByVal json As String, ByVal haystack As String) As String
    Dim r As String

    r = UrlNotes(json, haystack)
    r = Join2(r, PrefixNotes(json))
    r = Join2(r, SoftNotesS1(json))
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
    Dim srcText As String, finText As String

    idx = 0
    For Each it In modJsonLite.GetArrayItems(json, "current_coverage")
        oneText = CStr(it)
        If LCase$(Trim$(modJsonLite.GetStr(oneText, "certainty"))) = "assumed" Then
            If Not HasHearsay(oneText) Then
                r = Join2(r, "[V-S1-15] 接頭辞・出所の不整合があります: current_coverage[" & _
                             idx & "] は certainty=assumed ですが「(見立て)」がありません")
            End If
        End If
        idx = idx + 1
    Next it

    ' 裁定書39 R1-10: GetStr は「最初に現れた同名キー」を返す(15章§14)ので、
    '   json 全体へ "source" を引くと sources[] の要素やトップレベルの同名キーを
    '   拾う。リボン経路はキー順を強制できない以上、**financials を切り出してから**
    '   読む(W14 で潰した「スキーマ順に依存した読み」を再導入しない)。
    finText = ObjOf(json, "financials")
    If LenB(finText) = 0 Then
        PrefixNotes = r
        Exit Function
    End If
    srcText = LCase$(Trim$(modJsonLite.GetStr(finText, "source")))
    If LenB(srcText) > 0 And srcText <> "unknown" Then
        If AllUnknown(finText) Then
            r = Join2(r, "[V-S1-15] 接頭辞・出所の不整合があります: financials.source=" & _
                         srcText & " ですが4項目すべてが「不明」です")
        End If
    End If
    PrefixNotes = r
End Function

' current_coverage[] の1要素について、**値**のどれかに「(見立て)」があるか
'   (裁定書39 G-2)。キー名やスキーマ外のフィールドは見ない。
Private Function HasHearsay(ByVal itemJson As String) As Boolean
    Dim keyList() As String, i As Long, v As String

    keyList = Split(V3_CC_VALUE_KEYS, "|")
    For i = LBound(keyList) To UBound(keyList)
        v = modJsonLite.GetStr(itemJson, keyList(i))
        If LenB(v) > 0 Then
            If InStr(1, v, V3_HEARSAY1, vbBinaryCompare) > 0 Then
                HasHearsay = True
                Exit Function
            End If
            If InStr(1, v, V3_HEARSAY2, vbBinaryCompare) > 0 Then
                HasHearsay = True
                Exit Function
            End If
        End If
    Next i
End Function

' financials の4項目がすべて「不明」か(前後の空白は落として完全一致で見る)。
'   引数は **financials オブジェクトそのもの**(ObjOf で切り出したもの)。
Private Function AllUnknown(ByVal finJson As String) As Boolean
    Dim keyList() As String, i As Long

    keyList = Split(V3_FIN_KEYS, "|")
    For i = LBound(keyList) To UBound(keyList)
        If Trim$(modJsonLite.GetStr(finJson, keyList(i))) <> V3_UNKNOWN Then Exit Function
    Next i
    AllUnknown = True
End Function

' ============================================================================
' ObjOf - トップレベル(深さ1)のキー keyName が持つオブジェクト値 "{...}" を
'   そのまま切り出す(不在・オブジェクト以外・壊れたJSONは "")。
'   modJsonLite にオブジェクト取り出しの口が無いための最小ヘルパ。
' ============================================================================
Private Function ObjOf(ByVal srcJson As String, ByVal keyName As String) As String
    Dim valPos As Long, endPos As Long

    valPos = TopValuePos(srcJson, keyName)
    If valPos = 0 Then Exit Function
    If Mid$(srcJson, valPos, 1) <> "{" Then Exit Function
    endPos = ObjEndPos(srcJson, valPos)
    If endPos = 0 Then Exit Function
    ObjOf = Mid$(srcJson, valPos, endPos - valPos + 1)
End Function

' TopValuePos - トップレベル(深さ1)のキー keyName の**値の開始位置**(0=不在)。
'   modJsonLite.GetStr が「最初に現れた同名キー」を返す(深さを見ない)のに対し、
'   ここは深さ1だけを見る(裁定書39 R1-09 / R1-10)。
Private Function TopValuePos(ByVal srcJson As String, ByVal keyName As String) As Long
    Dim n As Long, i As Long, depth As Long
    Dim ch As String, endPos As Long, rawKey As String

    n = Len(srcJson)
    i = 1
    Do While i <= n
        ch = Mid$(srcJson, i, 1)
        If ch = """" Then
            endPos = StrEndPos(srcJson, i)
            If endPos = 0 Then Exit Function
            rawKey = Mid$(srcJson, i + 1, endPos - i - 1)
            i = SkipWs(srcJson, endPos + 1)
            If i <= n Then
                If Mid$(srcJson, i, 1) = ":" And depth = 1 And rawKey = keyName Then
                    i = SkipWs(srcJson, i + 1)
                    If i > n Then Exit Function
                    TopValuePos = i
                    Exit Function
                End If
            End If
        ElseIf ch = "{" Or ch = "[" Then
            depth = depth + 1
            i = i + 1
        ElseIf ch = "}" Or ch = "]" Then
            depth = depth - 1
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End Function

' 空白を読み飛ばした次の位置。
Private Function SkipWs(ByVal s As String, ByVal pos As Long) As Long
    Dim n As Long, i As Long
    n = Len(s)
    i = pos
    Do While i <= n
        If InStr(1, V3_WS, Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Do
        i = i + 1
    Loop
    SkipWs = i
End Function

' quotePos の開き引用符に対応する閉じ引用符の位置(0=見つからない)。
'   "\" のエスケープは2文字まとめて読み飛ばす。
Private Function StrEndPos(ByVal s As String, ByVal quotePos As Long) As Long
    Dim n As Long, i As Long, ch As String
    n = Len(s)
    i = quotePos + 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "\" Then
            i = i + 2
        ElseIf ch = """" Then
            StrEndPos = i
            Exit Function
        Else
            i = i + 1
        End If
    Loop
End Function

' openPos の "{" に対応する "}" の位置(0=見つからない)。文字列の中の波括弧は数えない。
Private Function ObjEndPos(ByVal s As String, ByVal openPos As Long) As Long
    Dim n As Long, i As Long, depth As Long, ch As String, e As Long
    n = Len(s)
    i = openPos
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = """" Then
            e = StrEndPos(s, i)
            If e = 0 Then Exit Function
            i = e + 1
        ElseIf ch = "{" Then
            depth = depth + 1
            i = i + 1
        ElseIf ch = "}" Then
            depth = depth - 1
            If depth = 0 Then
                ObjEndPos = i
                Exit Function
            End If
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End Function

' ============================================================================
' WarnNoteOf - 警告行を run_log の detail 用の1語へ畳む。
'   "V-S1-14:2,V-S1-15:1" の形(0件のケースは出さない。全部0件なら空文字)。
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

' ============================================================================
' SoftNotesS1 - V-S1-16 / V-S1-17(裁定書39 R1-09 / X-1)。**警告**なので
'   CheckS1Notes(注記チャネル)からだけ呼ぶ(裁定書40 P-M1)。空="指摘なし"。
'   modValidate.CheckS1 の戻り値には**載せない**(載せると修復リトライ ->
'   FailStep で案件 status=error になる)。
' ============================================================================
Public Function SoftNotesS1(ByVal json As String) As String
    Dim r As String, it As Variant, idx As Long, kindText As String

    ' V-S1-16: sources の欠落。Schema-S1 のルート required には残すが(direct 経路
    '   では強制できる)、CheckS1 の**不合格**からは外した。リボン経路でモデルが
    '   新設キーを落とすと、不合格 -> 修復リトライ1回 -> 失敗 で案件が二度と
    '   S1 を通せなくなるため(mock は必ず返すのでゲートでは露見しない)。
    If Not HasTopKey(json, "sources") Then
        r = Join2(r, "[V-S1-16] sources がありません(空として続けます)")
    End If

    ' V-S1-17: missing_info[].kind の enum。リボン経路はスキーマを強制しないので
    '   "conflicted" のような値が素通りし、SEC-03/04 の分離表示(kind==='conflict')
    '   から静かに外れていた。
    '   **欠落・空は不正としない**(裁定書40 P-M1): kind は v2.7 の新設キーで、
    '   モデルが落とすのは常態である。未記入は PostNormalize が "not_found"
    '   (見つからない)で**実際に補填**する(sources を空配列で補うのと同じ
    '   fail-open。裁定書41 §2)。補填値は kind!=="conflict" なので SEC-03/04 の
    '   分離表示でも通常の不足情報として扱われ、辻褄が合う。
    '   **値が enum 外のときだけ**警告する。なお本関数は補填の**前**の JSON を
    '   渡されることもある(呼び口は保存済みの sN_json)ので、欠落・空の判定は
    '   ここにも残す(補填が効いていれば空には当たらない)。
    idx = 0
    For Each it In modJsonLite.GetArrayItems(json, "missing_info")
        kindText = Trim$(modJsonLite.GetStr(CStr(it), "kind"))
        If LenB(kindText) > 0 Then
            If InStr(1, V3_MI_KIND, "|" & kindText & "|", vbBinaryCompare) = 0 Then
                r = Join2(r, "[V-S1-17] missing_info[" & idx & "].kind が不正です: " & kindText)
            End If
        End If
        idx = idx + 1
    Next it
    SoftNotesS1 = r
End Function

' ============================================================================
' PostNormalize - modValidate.NormalizeLlmJson が正規化の直後に掛ける後処理。
'   S1 の sources 補填(裁定書39 R1-09)と missing_info[].kind の補填
'   (裁定書41 §2)の2本。他の step は素通し。
' ============================================================================
Public Function PostNormalize(ByVal stepName As String, ByVal json As String) As String
    PostNormalize = json
    If LCase$(Trim$(stepName)) <> "s1" Then Exit Function
    PostNormalize = FillMissingKind(FillEmptySources(json))
End Function

' トップレベルに "sources" が無ければ空配列を足す。オブジェクトとして読めない
'   文字列(抽出失敗・空・配列)は触らない。
Private Function FillEmptySources(ByVal json As String) As String
    FillEmptySources = json
    If HasTopKey(json, "sources") Then Exit Function
    FillEmptySources = AddPair(json, """sources"":[]")
End Function

' ============================================================================
' FillMissingKind - missing_info[] の各要素に kind が無い/空のとき "not_found"
'   を実際に書き込む(裁定書41 §2。裁定書40 P-M1「空なら not_found」の残り半分)。
' ----------------------------------------------------------------------------
'   SoftNotesS1 の V-S1-17 は「欠落・空を警告しない」だけで値を直していなかった。
'   18章 SEC-04 の「種別」列は modHtmlTemplate2 が LB(LMK, mi[k].kind) で引き、
'   modHtmlTemplate1.LB は未知キーをそのまま返すので、空のままだと**空欄**で
'   出る(補填すれば LMK.not_found の「未取得」が出る)。sources と同じ場所で
'   同じように補う。
'
'   触らない(素通しする)場合(fail-open。壊れた入力で JSON を壊さない):
'     ・トップレベルに missing_info が無い/値が配列でない/閉じていない
'     ・要素にオブジェクト以外(文字列・数値)が混ざっている
'     ・補う要素が1件も無い(そのときは文字列を作り直さない)
'   要素は**配列を切り出してから** modJsonLite.GetArrayItems に渡す。同関数は
'   深さを見ず「最初に現れた同名キー」を拾うため(裁定書39 R1-10 と同じ理由)。
' ============================================================================
Private Function FillMissingKind(ByVal json As String) As String
    Dim valPos As Long, endPos As Long, arrText As String
    Dim it As Variant, t As String, newT As String, outText As String
    Dim n As Long, changed As Boolean

    FillMissingKind = json
    valPos = TopValuePos(json, "missing_info")
    If valPos = 0 Then Exit Function
    If Mid$(json, valPos, 1) <> "[" Then Exit Function
    endPos = ArrEndPos(json, valPos)
    If endPos = 0 Then Exit Function
    arrText = Mid$(json, valPos, endPos - valPos + 1)

    For Each it In modJsonLite.GetArrayItems("{""missing_info"":" & arrText & "}", _
                                             "missing_info")
        t = Trim$(CStr(it))
        If Left$(t, 1) <> "{" Or Right$(t, 1) <> "}" Then Exit Function
        If LenB(Trim$(modJsonLite.GetStr(t, "kind"))) = 0 Then
            newT = SetKindNotFound(t)
            If newT <> t Then
                t = newT
                changed = True
            End If
        End If
        If n > 0 Then outText = outText & ","
        outText = outText & t
        n = n + 1
    Next it
    If Not changed Then Exit Function

    FillMissingKind = Left$(json, valPos - 1) & "[" & outText & "]" & _
                      Mid$(json, endPos + 1)
End Function

' ============================================================================
' SetKindNotFound - 要素オブジェクト t の kind を "not_found" にする。
'   キーが無ければ足し、**空文字・空白だけの文字列なら値を置き換える**
'   (足すだけだと同名キーが2つ並び、15章§14「最初に現れた同名キー」の規則で
'   空のほうが勝って補填が効かない)。kind が文字列でない(null・数値・配列)
'   ときは触らない(fail-open。読めない形を推測で書き換えない)。
' ============================================================================
Private Function SetKindNotFound(ByVal t As String) As String
    Dim valPos As Long, e As Long

    SetKindNotFound = t
    valPos = TopValuePos(t, "kind")
    If valPos = 0 Then
        SetKindNotFound = AddPair(t, """kind"":""not_found""")
        Exit Function
    End If
    If Mid$(t, valPos, 1) <> """" Then Exit Function
    e = StrEndPos(t, valPos)
    If e = 0 Then Exit Function
    SetKindNotFound = Left$(t, valPos - 1) & """not_found""" & Mid$(t, e + 1)
End Function

' ============================================================================
' AddPair - オブジェクト文字列 t の末尾へ "キー":値 を1組足す。t がオブジェクト
'   として読めなければ**触らない**(fail-open)。sources の空配列補填と
'   missing_info[].kind の補填が同じ1本を通る(同じ規則を2箇所に持たない)。
' ============================================================================
Private Function AddPair(ByVal t As String, ByVal pairText As String) As String
    Dim s As String, p As Long

    AddPair = t
    s = RTrim$(t)
    If LenB(s) = 0 Then Exit Function
    If Left$(s, 1) <> "{" Then Exit Function
    If Right$(s, 1) <> "}" Then Exit Function

    ' 閉じ "}" の直前が "{" なら空オブジェクトなので "," を置かない。
    p = Len(s) - 1
    Do While p >= 1
        If InStr(1, V3_WS, Mid$(s, p, 1), vbBinaryCompare) = 0 Then Exit Do
        p = p - 1
    Loop
    If p < 1 Then Exit Function
    If Mid$(s, p, 1) = "{" Then
        AddPair = Left$(s, Len(s) - 1) & pairText & "}"
    Else
        AddPair = Left$(s, Len(s) - 1) & "," & pairText & "}"
    End If
End Function

' ArrEndPos - openPos の "[" に対応する "]" の位置(0=閉じていない)。ObjEndPos の
'   配列版(文字列リテラルの中の括弧は数えない)。
Private Function ArrEndPos(ByVal s As String, ByVal openPos As Long) As Long
    Dim n As Long, i As Long, depth As Long, ch As String, e As Long
    n = Len(s)
    i = openPos
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = """" Then
            e = StrEndPos(s, i)
            If e = 0 Then Exit Function
            i = e + 1
        ElseIf ch = "[" Then
            depth = depth + 1
            i = i + 1
        ElseIf ch = "]" Then
            depth = depth - 1
            If depth = 0 Then
                ArrEndPos = i
                Exit Function
            End If
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End Function

' トップレベル(深さ1)に keyName があるか。
Private Function HasTopKey(ByVal srcJson As String, ByVal keyName As String) As Boolean
    HasTopKey = (TopValuePos(srcJson, keyName) > 0)
End Function

' HeadOverlap - 顧客語の先頭1〜maxLen 文字が、社内語の直前の本文の
'   末尾と重複するか(語頭側の二重を防ぐ。UndoNeeded の語尾側と対の関係)。
Public Function HeadOverlap(ByVal dstText As String, ByVal hay As String, _
                            ByVal startPos As Long, ByVal maxLen As Long) As Boolean
    Dim k As Long

    For k = 1 To maxLen
        If k <= Len(dstText) And startPos - k >= 1 Then
            If Mid$(hay, startPos - k, k) = Left$(dstText, k) Then
                HeadOverlap = True
                Exit Function
            End If
        End If
    Next k
End Function
