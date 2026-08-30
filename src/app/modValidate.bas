Attribute VB_Name = "modValidate"
Option Explicit

' ============================================================================
' modValidate - LLM応答の検証(15章§11の検証ルール表。T-22)
' ----------------------------------------------------------------------------
' 正: 15章§11(全64ケースの一覧) / 15章の各Check節の検証ルール表(条件・判定・
'   エラー文テンプレ) / 15章§0 原則10(判定3値) / 19章§3(enum) / 14章§6
'   (公開シグネチャ) / 16章 E-49(NormalizeLlmJson) / 17章 T-22(DoD)。
'
' 戻り値の規約(15章§0 原則10): **不合格・警告**のエラー文を vbLf 区切りで連結
'   した文字列(発火なしは "")。各行は必ず "[ケースID] " で始まる(17章§4-2の
'   照合スクリプトが機械照合する)。1ケースが複数件の欠陥に当たるときは**同じ
'   ケースIDの行を件数分**出す。判定「合格」の V-S2C-05 / V-S3C-05 は
'   エラー文を持たないため戻り値に現れない。
'
' **純関数**であること(12章§2・17章§4-1 層(a)): シート・config・ログ・
'   modKnowledge に一切触れない。ID実在検査の一覧(menusText 等)と、JSONの外側の
'   文脈(case_type / dossier_tier / 前ラウンドのS2 / 現場メモの有無)は
'   **引数で受ける**(14章§6 が引数渡し設計を正式宣言。裁定書7 A-1)。実在ID一覧
'   テキスト(15章の1行書式)は呼出側=modPipeline が modKnowledge から取得して渡す。
'
' ID実在検査は**fail-closed**(裁定書7 A-2): 非空のIDがあるのに一覧が空(未提供)
'   なら「[ケースID] ID実在検査が実行できません(ID一覧未提供)」で**不合格**とする。
'   渡し忘れを黙って合格に見せない(16章 E-07 のKPI「すり抜け0件」)。
'
' 30,000字契約(12章§2)により CheckS4 / CheckPF / CheckS2C / CheckS3C の本体と
'   NormalizeLlmJson の正規化エンジンは **modValidate2** にある。14章§6が
'   宣言した公開関数8本の入口は仕様どおり本モジュールに置き、本体へ委譲する。
'
' R4準拠(Excelトークン不使用)・CP932準拠(15章§0 原則7)・例外を投げない(14章§6)。
' 配列添字の {i} は**0始まり**(JSONパスの読み方に合わせる)。{n}/{no}/{value}
'   はJSONの実値。{p} は百分率の整数部(切り捨て)。
' ============================================================================

' --- enum(19章§3。前後を "|" で挟んだ照合用の並び) ---
Private Const VE_LOC_TYPE As String = "|工場|本社|店舗|倉庫|その他|"
Private Const VE_ASPECT As String = "|profile|business|sites|history|news|hr|finance_risk|sales_memo|sns|competitors|market|finance|insurance_ctx|hazard|"
Private Const VE_CATEGORY As String = "|strategy_market|supply_chain|manufacturing_quality|sales_customer|facility_bcp|hr_labor|digital_info|legal_regulatory|finance_counterparty|brand_social|"
Private Const VE_RISK_STATUS As String = "|proposed|confirmed|rejected|new|"
Private Const VE_FREQ As String = "|high|mid|low|"
Private Const VE_IMPACT As String = "|large|mid|small|"
Private Const VE_SOURCE As String = "|hp|yuho|memo|contract|prev_renewal|knowledge|inference|"
Private Const VE_TRANSFER As String = "|cover|partial|hard|"
Private Const VE_GAP_TYPE As String = "|uninsured|underinsured|overlap|"
Private Const VE_HORIZON As String = "|already|near|mid_long|"
Private Const VE_KIND As String = "|upsell|cross_sell|scheme|"

' --- 15章 Schema-S1 の required 15キー(V-S1-01) ---
Private Const VS1_REQUIRED As String = "company_name|business_summary|main_products|processes|locations|supply_chain|customers|workforce_notes|management_notes|strategy_outlook|current_coverage|field_insights|missing_info|input_quality|research_requests"

' --- 件数・字数のしきい値(15章の各ルール表) ---
Private Const VS1_ASPECT_N As Long = 14
Private Const VS1_PROMPT_MAX As Long = 1800
Private Const VS1_REQ_MAX As Long = 6
Private Const VS2_RISK_MIN As Long = 5
Private Const VS2_RISK_MAX As Long = 20
Private Const VS2_PREV_MIN As Long = 1
Private Const VS2_PREV_MAX As Long = 3
Private Const VS2_EMERGING_MAX As Long = 3
Private Const VS3_STORY_N As Long = 3

' --- 前ラウンド無し(初回実行)を表す値(15章§3 prevS2Json の「なし」) ---
Private Const V_NONE_TEXT As String = "なし"

' --- ID実在検査が実行できないときのエラー文(裁定書7 A-2 の fail-closed) ---
Private Const V_NOLIST As String = "ID実在検査が実行できません(ID一覧未提供)"

' NormalizeLlmJson - 尾部劣化(同名項目の重複出力)の正規化(14章§5(2.5)・16章E-49)
'   本体は modValidate2.NormalizeCore(30,000字契約による分割)。
Public Function NormalizeLlmJson(ByVal stepName As String, ByVal json As String, _
                                 ByRef removedCount As Long) As String
    NormalizeLlmJson = modValidate2.NormalizeCore(stepName, json, removedCount)
End Function

' CheckS1 - 企業プロファイル構造化の検証(15章§2 CheckS1 検証ルール表。11件)
'   fieldNoteProvided: 現場メモの有無(V-S1-11 は提供時のみ判定。15章§2 補足)。
'     s1Json から判別できないので modPipeline が渡す。**既定は True**(V-S1-11 は
'     「警告」なので不明なら黙らず出す。15章§0 原則10)。
Public Function CheckS1(ByVal json As String, ByVal caseType As String, _
                        Optional ByVal fieldNoteProvided As Boolean = True) As String
    Dim r As String
    Dim keyList() As String
    Dim i As Long
    Dim idx As Long
    Dim itemsCol As Collection
    Dim it As Variant
    Dim sVal As String
    Dim nCount As Long
    Dim seenAspect As String
    Dim aspectList() As String
    Dim ctype As String

    ctype = LCase$(Trim$(caseType))

    ' --- V-S1-01: required 15キーのいずれかが欠落 ---
    keyList = Split(VS1_REQUIRED, "|")
    For i = LBound(keyList) To UBound(keyList)
        If Not TopKeyExists(json, keyList(i)) Then
            Ap r, "[V-S1-01] 必須キー " & keyList(i) & " がありません"
        End If
    Next i

    ' --- V-S1-02: locations[].type が enum 外 ---
    Set itemsCol = modJsonLite.GetArrayItems(json, "locations")
    idx = 0
    For Each it In itemsCol
        sVal = modJsonLite.GetStr(CStr(it), "type")
        If Not InEnum(sVal, VE_LOC_TYPE) Then
            Ap r, "[V-S1-02] locations[" & idx & "].type が不正です: " & sVal
        End If
        idx = idx + 1
    Next it

    ' --- V-S1-03 / V-S1-04: current_coverage と case_type ---
    nCount = modJsonLite.GetArrayItems(json, "current_coverage").Count
    If ctype = "renewal" And nCount = 0 Then
        Ap r, "[V-S1-03] 更新案件ですが current_coverage が0件です"
    End If
    If ctype = "new" And nCount > 0 Then
        Ap r, "[V-S1-04] 新規案件ですが current_coverage が" & nCount & "件あります"
    End If

    ' --- V-S1-05: missing_info が0件 ---
    If modJsonLite.GetArrayItems(json, "missing_info").Count = 0 Then
        Ap r, "[V-S1-05] missing_info が0件です"
    End If

    ' --- V-S1-06 / V-S1-07: input_quality.coverage の件数と14 aspect ---
    Set itemsCol = modJsonLite.GetArrayItems(json, "coverage")
    nCount = itemsCol.Count
    If nCount <> VS1_ASPECT_N Then
        Ap r, "[V-S1-06] input_quality.coverage が" & nCount & "件です(14件必要)"
    End If
    If nCount > 0 Then
        seenAspect = "|"
        For Each it In itemsCol
            sVal = modJsonLite.GetStr(CStr(it), "aspect")
            If (Not InEnum(sVal, VE_ASPECT)) Or InStr(seenAspect, "|" & sVal & "|") > 0 Then
                Ap r, "[V-S1-07] input_quality.coverage の aspect に欠落または重複があります: " & sVal
            Else
                seenAspect = seenAspect & sVal & "|"
            End If
        Next it
        aspectList = Split(Mid$(VE_ASPECT, 2, Len(VE_ASPECT) - 2), "|")
        For i = LBound(aspectList) To UBound(aspectList)
            If InStr(seenAspect, "|" & aspectList(i) & "|") = 0 Then
                Ap r, "[V-S1-07] input_quality.coverage の aspect に欠落または重複があります: " & aspectList(i)
            End If
        Next i
    End If

    ' --- V-S1-08 / V-S1-09 / V-S1-10: research_requests ---
    sVal = modJsonLite.GetStr(json, "overall")
    Set itemsCol = modJsonLite.GetArrayItems(json, "research_requests")
    nCount = itemsCol.Count
    If sVal <> "high" And nCount = 0 Then
        Ap r, "[V-S1-08] overall=" & sVal & " ですが research_requests が0件です"
    End If
    idx = 0
    For Each it In itemsCol
        i = Len(modJsonLite.GetStr(CStr(it), "prompt_text"))
        If i > VS1_PROMPT_MAX Then
            Ap r, "[V-S1-09] research_requests[" & idx & "].prompt_text が" & i & "字です(1800字以内)"
        End If
        idx = idx + 1
    Next it
    If nCount > VS1_REQ_MAX Then
        Ap r, "[V-S1-10] research_requests が" & nCount & "件です(6件以内)"
    End If

    ' --- V-S1-11: 現場メモ提供ありで field_insights が0件 ---
    If fieldNoteProvided Then
        If modJsonLite.GetArrayItems(json, "field_insights").Count = 0 Then
            Ap r, "[V-S1-11] 現場メモがありますが field_insights が0件です"
        End If
    End If

    CheckS1 = r
End Function

' CheckS2 - リスク仮説＋付保ギャップの検証(15章§3 CheckS2 検証ルール表。17件)
'   menusText : S2へ注入した menusSummary(15章§3の1行書式。V-S2-06 の実在判定)
'   prevS2Json: 前ラウンドのS2 JSON。"" または "なし" を初回実行とみなす
'               (V-S2-09 は初回のみ / V-S2-10 は第2ラウンド以降のみ)
'   ※ Optional String に `= ""` を書かない。VBAの省略時と同義だが LibreOffice
'     Basic は空文字既定を束縛できず省略呼び出しが実行時エラー13になる(層(c))。
Public Function CheckS2(ByVal json As String, ByVal caseType As String, _
                        Optional ByVal menusText As String, _
                        Optional ByVal prevS2Json As String) As String
    Dim r As String
    Dim risksCol As Collection
    Dim gapsCol As Collection
    Dim emCol As Collection
    Dim prevCol As Collection
    Dim it As Variant
    Dim prevItem As Variant
    Dim rj As String
    Dim noText As String
    Dim sVal As String
    Dim nRisk As Long
    Dim nCount As Long
    Dim idx As Long
    Dim seenNos As String
    Dim infCount As Long
    Dim hardCount As Long
    Dim firstRound As Boolean
    Dim ctype As String
    Dim prevText As String

    ctype = LCase$(Trim$(caseType))
    prevText = Trim$(prevS2Json)
    firstRound = (LenB(prevText) = 0 Or prevText = V_NONE_TEXT)

    Set risksCol = modJsonLite.GetArrayItems(json, "risks")
    Set gapsCol = modJsonLite.GetArrayItems(json, "gaps")
    Set emCol = modJsonLite.GetArrayItems(json, "emerging_risks")
    nRisk = risksCol.Count

    ' --- V-S2-01: risks の件数 ---
    If nRisk < VS2_RISK_MIN Or nRisk > VS2_RISK_MAX Then
        Ap r, "[V-S2-01] risks が" & nRisk & "件です(5～20件)"
    End If

    seenNos = "|"
    For Each it In risksCol
        rj = CStr(it)
        noText = Trim$(modJsonLite.GetStr(rj, "risk_no"))

        ' --- V-S2-02: risk_no の重複 ---
        If InStr(seenNos, "|" & noText & "|") > 0 Then
            Ap r, "[V-S2-02] risk_no " & noText & " が重複しています"
        Else
            seenNos = seenNos & noText & "|"
        End If

        ' --- V-S2-03: risks[] の enum 6キー ---
        ChkRiskEnum r, rj, noText, "category", "category", VE_CATEGORY
        ChkRiskEnum r, rj, noText, "status", "status", VE_RISK_STATUS
        ChkRiskEnum r, rj, noText, "frequency", "frequency", VE_FREQ
        ChkRiskEnum r, rj, noText, "impact", "impact", VE_IMPACT
        ChkRiskEnum r, rj, noText, "source", "evidence.source", VE_SOURCE
        ChkRiskEnum r, rj, noText, "transferability", "insurability.transferability", VE_TRANSFER

        ' --- V-S2-04: evidence.quote が空 ---
        If LenB(modJsonLite.GetStr(rj, "quote")) = 0 Then
            Ap r, "[V-S2-04] risk_no " & noText & " の evidence.quote が空です"
        End If

        ' --- V-S2-05 / V-S2-06: preventions ---
        nCount = 0
        For Each prevItem In modJsonLite.GetArrayItems(rj, "preventions")
            nCount = nCount + 1
            sVal = Trim$(modJsonLite.GetStr(CStr(prevItem), "related_menu_id"))
            If LenB(sVal) > 0 Then
                If Not ListGiven(menusText) Then
                    Ap r, "[V-S2-06] " & V_NOLIST
                ElseIf Not IdExistsIn(menusText, sVal) Then
                    Ap r, "[V-S2-06] risk_no " & noText & " の related_menu_id " & sVal & " は実在しません"
                End If
            End If
        Next prevItem
        If nCount < VS2_PREV_MIN Or nCount > VS2_PREV_MAX Then
            Ap r, "[V-S2-05] risk_no " & noText & " の preventions が" & nCount & "件です(1～3件)"
        End If

        ' --- V-S2-07 / V-S2-08: スコアとバンドの整合 ---
        sVal = modJsonLite.GetStr(rj, "frequency")
        idx = modJsonLite.GetLong(rj, "frequency_score", 0)
        If Not ScoreFits(idx, sVal, True) Then
            Ap r, "[V-S2-07] risk_no " & noText & " の frequency_score " & idx & _
                  " が frequency=" & sVal & " と不整合です"
        End If
        sVal = modJsonLite.GetStr(rj, "impact")
        idx = modJsonLite.GetLong(rj, "impact_score", 0)
        If Not ScoreFits(idx, sVal, False) Then
            Ap r, "[V-S2-08] risk_no " & noText & " の impact_score " & idx & _
                  " が impact=" & sVal & " と不整合です"
        End If

        ' --- V-S2-09: 初回ラウンドの status ---
        sVal = modJsonLite.GetStr(rj, "status")
        If firstRound And sVal <> "proposed" Then
            Ap r, "[V-S2-09] 初回実行ですが risk_no " & noText & " の status が " & sVal & " です"
        End If

        If modJsonLite.GetStr(rj, "source") = "inference" Then infCount = infCount + 1
        If modJsonLite.GetStr(rj, "transferability") = "hard" Then hardCount = hardCount + 1
    Next it

    ' --- V-S2-10: 第2ラウンド以降で risks 件数が減少 ---
    If Not firstRound Then
        Set prevCol = modJsonLite.GetArrayItems(prevS2Json, "risks")
        If nRisk < prevCol.Count Then
            Ap r, "[V-S2-10] risks が前ラウンド" & prevCol.Count & "件から" & nRisk & _
                  "件に減りました(rejected の削除禁止)"
        End If
    End If

    ' --- V-S2-11 / V-S2-12: gaps と case_type ---
    nCount = gapsCol.Count
    If ctype = "renewal" And nCount = 0 Then
        Ap r, "[V-S2-11] 更新案件ですが gaps が0件です"
    End If
    If ctype = "new" And nCount > 0 Then
        Ap r, "[V-S2-12] 新規案件ですが gaps が" & nCount & "件あります"
    End If

    ' --- V-S2-13: gap_no の重複 / gap_type の enum ---
    seenNos = "|"
    For Each it In gapsCol
        rj = CStr(it)
        noText = Trim$(modJsonLite.GetStr(rj, "gap_no"))
        If InStr(seenNos, "|" & noText & "|") > 0 Then
            Ap r, "[V-S2-13] gaps の gap_no が不正です: " & noText
        Else
            seenNos = seenNos & noText & "|"
        End If
        sVal = modJsonLite.GetStr(rj, "gap_type")
        If Not InEnum(sVal, VE_GAP_TYPE) Then
            Ap r, "[V-S2-13] gaps の gap_type が不正です: " & sVal
        End If
    Next it

    ' --- V-S2-14: inference 比率(百分率は整数部で表示する) ---
    If nRisk > 0 Then
        If infCount * 2 > nRisk Then
            Ap r, "[V-S2-14] inference 比率が" & Int(infCount * 100 / nRisk) & "%です(50%以下が目安)"
        End If
    End If

    ' --- V-S2-15: transferability=hard が0件 ---
    If hardCount = 0 Then
        Ap r, "[V-S2-15] transferability=hard のリスクが0件です"
    End If

    ' --- V-S2-16: emerging_risks の上限(0件は正常) ---
    nCount = emCol.Count
    If nCount > VS2_EMERGING_MAX Then
        Ap r, "[V-S2-16] emerging_risks が" & nCount & "件です(0～3件)"
    End If

    ' --- V-S2-17: emerging_risks[] の enum 3キー ---
    idx = 0
    For Each it In emCol
        rj = CStr(it)
        ChkEmergEnum r, rj, idx, "category", VE_CATEGORY
        ChkEmergEnum r, rj, idx, "horizon", VE_HORIZON
        ChkEmergEnum r, rj, idx, "evidence_source", VE_SOURCE
        idx = idx + 1
    Next it

    CheckS2 = r
End Function

' CheckS3 - 提案マッチングの検証(15章§4 CheckS3 検証ルール表。13件。最重要)
'   s2Json     : 審査対象のS2 JSON(risk_no / gap_no の実在判定)
'   menusText / linesText / schemesText / casesText: S3へ注入した一覧
'     (15章§4の1行書式。行頭 "[ID] ")。V-S3-03～V-S3-06 の実在判定に使う。
'     空文字は「一覧未提供」= fail-closed で不合格(裁定書7 A-2。V_NOLIST)。
'   caseType  : V-S3-13 用。省略時は IsRenewalCtx が s2Json から導く。
'   ※ Optional String の既定値は CheckS2 の注記と同じ理由で書かない。
Public Function CheckS3(ByVal json As String, ByVal s2Json As String, _
                        Optional ByVal menusText As String, _
                        Optional ByVal linesText As String, _
                        Optional ByVal schemesText As String, _
                        Optional ByVal casesText As String, _
                        Optional ByVal caseType As String) As String
    Dim r As String
    Dim storiesCol As Collection
    Dim topicsCol As Collection
    Dim it As Variant
    Dim elemVal As Variant
    Dim sj As String
    Dim noText As String
    Dim sVal As String
    Dim headline As String
    Dim riskSet As String
    Dim gapSet As String
    Dim seenNos As String
    Dim idx As Long
    Dim nStory As Long
    Dim usedCount As Long
    Dim kindHit As Long
    Dim storyNo As Long

    riskSet = CollectNos(s2Json, "risks", "risk_no")
    gapSet = CollectNos(s2Json, "gaps", "gap_no")

    Set storiesCol = modJsonLite.GetArrayItems(json, "stories")
    Set topicsCol = modJsonLite.GetArrayItems(json, "do_not_propose")
    nStory = storiesCol.Count

    ' --- V-S3-01: stories は3件固定 ---
    If nStory <> VS3_STORY_N Then
        Ap r, "[V-S3-01] stories が" & nStory & "件です(3件固定)"
    End If

    idx = 0
    seenNos = "|"
    For Each it In storiesCol
        sj = CStr(it)
        idx = idx + 1
        storyNo = modJsonLite.GetLong(sj, "story_no", 0)
        noText = Trim$(modJsonLite.GetStr(sj, "story_no"))

        ' --- V-S3-11: story_no が 1..3 の連番でない/重複 ---
        If storyNo <> idx Or storyNo < 1 Or storyNo > VS3_STORY_N _
           Or InStr(seenNos, "|" & noText & "|") > 0 Then
            Ap r, "[V-S3-11] story_no が1..3の連番ではありません: " & noText
        End If
        seenNos = seenNos & noText & "|"

        ' --- V-S3-02: proposal_kind の enum ---
        sVal = modJsonLite.GetStr(sj, "proposal_kind")
        If Not InEnum(sVal, VE_KIND) Then
            Ap r, "[V-S3-02] story_no " & noText & " の proposal_kind が不正です: " & sVal
        End If
        If sVal = "upsell" Or sVal = "cross_sell" Then kindHit = kindHit + 1

        ' --- V-S3-03: menu_ids[] の実在 ---
        For Each elemVal In modJsonLite.GetArrayItems(sj, "menu_ids")
            usedCount = usedCount + 1
            If Not ListGiven(menusText) Then
                Ap r, "[V-S3-03] " & V_NOLIST
            ElseIf Not IdExistsIn(menusText, Trim$(CStr(elemVal))) Then
                Ap r, "[V-S3-03] story_no " & noText & " の menu_id " & CStr(elemVal) & " は実在しません"
            End If
        Next elemVal

        ' --- V-S3-04: line_ids[] の実在 ---
        For Each elemVal In modJsonLite.GetArrayItems(sj, "line_ids")
            If Not ListGiven(linesText) Then
                Ap r, "[V-S3-04] " & V_NOLIST
            ElseIf Not IdExistsIn(linesText, Trim$(CStr(elemVal))) Then
                Ap r, "[V-S3-04] story_no " & noText & " の line_id " & CStr(elemVal) & " は実在しません"
            End If
        Next elemVal

        ' --- V-S3-05: scheme_id の実在("" は許す) ---
        sVal = Trim$(modJsonLite.GetStr(sj, "scheme_id"))
        If LenB(sVal) > 0 Then
            usedCount = usedCount + 1
            If Not ListGiven(schemesText) Then
                Ap r, "[V-S3-05] " & V_NOLIST
            ElseIf Not IdExistsIn(schemesText, sVal) Then
                Ap r, "[V-S3-05] story_no " & noText & " の scheme_id " & sVal & " は実在しません"
            End If
        End If

        ' --- V-S3-06: similar_case_id の実在("" は許す) ---
        sVal = Trim$(modJsonLite.GetStr(sj, "similar_case_id"))
        If LenB(sVal) > 0 Then
            If Not ListGiven(casesText) Then
                Ap r, "[V-S3-06] " & V_NOLIST
            ElseIf Not IdExistsIn(casesText, sVal) Then
                Ap r, "[V-S3-06] story_no " & noText & " の similar_case_id " & sVal & " は実在しません"
            End If
        End If

        ' --- V-S3-07 / V-S3-08: target_risk_nos / target_gap_nos の実在 ---
        For Each elemVal In modJsonLite.GetArrayItems(sj, "target_risk_nos")
            If InStr(riskSet, "|" & Trim$(CStr(elemVal)) & "|") = 0 Then
                Ap r, "[V-S3-07] story_no " & noText & " の target_risk_no " & CStr(elemVal) & _
                      " が S2 に存在しません"
            End If
        Next elemVal
        For Each elemVal In modJsonLite.GetArrayItems(sj, "target_gap_nos")
            If InStr(gapSet, "|" & Trim$(CStr(elemVal)) & "|") = 0 Then
                Ap r, "[V-S3-08] story_no " & noText & " の target_gap_no " & CStr(elemVal) & _
                      " が S2 に存在しません"
            End If
        Next elemVal

        ' --- V-S3-10: do_not_propose の topic との重複 ---
        headline = modJsonLite.GetStr(sj, "headline")
        For Each elemVal In topicsCol
            sVal = Trim$(modJsonLite.GetStr(CStr(elemVal), "topic"))
            If LenB(sVal) > 0 Then
                If InStr(headline, sVal) > 0 Or IdInArray(sj, "menu_ids", sVal) _
                   Or IdInArray(sj, "line_ids", sVal) Then
                    Ap r, "[V-S3-10] story_no " & noText & " が do_not_propose の topic " & _
                          sVal & " と重複しています"
                End If
            End If
        Next elemVal
    Next it

    ' --- V-S3-09: 当社メニュー・型が3本合計で0 ---
    If usedCount = 0 Then
        Ap r, "[V-S3-09] 全ストーリーで当社メニュー・型がひとつも使われていません"
    End If

    ' --- V-S3-12: unmatched_risks[].risk_no の実在 ---
    For Each it In modJsonLite.GetArrayItems(json, "unmatched_risks")
        noText = Trim$(modJsonLite.GetStr(CStr(it), "risk_no"))
        If InStr(riskSet, "|" & noText & "|") = 0 Then
            Ap r, "[V-S3-12] unmatched_risks の risk_no " & noText & " が S2 に存在しません"
        End If
    Next it

    ' --- V-S3-13: 更新案件で upsell も cross_sell も0本 ---
    If IsRenewalCtx(caseType, s2Json) And kindHit = 0 Then
        Ap r, "[V-S3-13] 更新案件ですが upsell/cross_sell が0本です"
    End If

    CheckS3 = r
End Function

' CheckS4 / CheckPF / CheckS2C / CheckS3C - 本体は modValidate2(30,000字契約)。
' dossierTier 省略時は「ティア不明」= V-S4-01/02 の両方を当てる(CheckS4Core 注記)。
Public Function CheckS4(ByVal json As String, _
                        Optional ByVal dossierTier As String, _
                        Optional ByVal maxSlidesT2 As Long = 10) As String
    CheckS4 = modValidate2.CheckS4Core(json, dossierTier, maxSlidesT2)
End Function

Public Function CheckS2C(ByVal json As String, Optional ByVal s2Json As String) As String
    CheckS2C = modValidate2.CheckS2CCore(json, s2Json)
End Function

Public Function CheckS3C(ByVal json As String) As String
    CheckS3C = modValidate2.CheckS3CCore(json)
End Function

Public Function CheckPF(ByVal json As String, Optional ByVal refIdsText As String) As String
    CheckPF = modValidate2.CheckPFCore(json, refIdsText)
End Function

' --- 内部ヘルパ(本モジュール専用の純関数) ---

' エラー行を vbLf 区切りで積む(15章§0 原則10)。
Private Sub Ap(ByRef outText As String, ByVal lineText As String)
    If LenB(outText) = 0 Then
        outText = lineText
    Else
        outText = outText & vbLf & lineText
    End If
End Sub

' enum の所属判定。空文字は常に不所属(必須キー欠落も enum 違反として現れる)。
Private Function InEnum(ByVal sVal As String, ByVal enumList As String) As Boolean
    If LenB(sVal) = 0 Then Exit Function
    InEnum = (InStr(enumList, "|" & sVal & "|") > 0)
End Function

' 更新案件の文脈か(V-S3-13 の唯一の判断点)。caseType があれば 13章§2.1 の2値を
' 使い、省略時は s2Json の gaps 非空で判定する(V-S2-12「新規案件で gaps が1件
' 以上」は不合格なので、gaps 非空のS2は更新案件のものと確定できる)。
Private Function IsRenewalCtx(ByVal caseType As String, ByVal s2Json As String) As Boolean
    If LenB(Trim$(caseType)) > 0 Then
        IsRenewalCtx = (LCase$(Trim$(caseType)) = "renewal")
    Else
        IsRenewalCtx = (modJsonLite.GetArrayItems(s2Json, "gaps").Count > 0)
    End If
End Function

' ID一覧が渡されたか(fail-closed の判断点。裁定書7 A-2)。空=未提供であって
' 「一覧に無い」ではないため、呼び出し側は V_NOLIST の不合格を出す。15章§6.1 の
' "(登録なし)" は空ではないので通常の実在検査が走る。
Private Function ListGiven(ByVal listText As String) As Boolean
    ListGiven = (LenB(Trim$(listText)) > 0)
End Function

' 注入テキスト(1行1件・行頭 "[ID] ")に当該IDの行があるか。ID実在検査
' (V-S2-06 / V-S3-03～06)の照合部。一覧が空のときは False(未提供の扱いは
' ListGiven で先に分けるので、ここへ来た時点で「一覧に無い」を意味する)。
Private Function IdExistsIn(ByVal listText As String, ByVal idText As String) As Boolean
    If LenB(idText) = 0 Then Exit Function
    If LenB(listText) = 0 Then Exit Function
    If Left$(listText, Len(idText) + 2) = "[" & idText & "]" Then
        IdExistsIn = True
        Exit Function
    End If
    IdExistsIn = (InStr(listText, vbLf & "[" & idText & "]") > 0)
End Function

' risks / gaps などの採番キーを "|1|2|" 形式に集める。
Private Function CollectNos(ByVal srcJson As String, ByVal arrKey As String, _
                            ByVal noKey As String) As String
    Dim outText As String
    Dim it As Variant
    outText = "|"
    For Each it In modJsonLite.GetArrayItems(srcJson, arrKey)
        outText = outText & Trim$(modJsonLite.GetStr(CStr(it), noKey)) & "|"
    Next it
    CollectNos = outText
End Function

' 文字列配列の要素に idText と完全一致するものがあるか(V-S3-10)。
Private Function IdInArray(ByVal objJson As String, ByVal arrKey As String, _
                           ByVal idText As String) As Boolean
    Dim it As Variant
    For Each it In modJsonLite.GetArrayItems(objJson, arrKey)
        If Trim$(CStr(it)) = idText Then
            IdInArray = True
            Exit Function
        End If
    Next it
End Function

' V-S2-03 の1キー分(risks[] の enum)。
Private Sub ChkRiskEnum(ByRef outText As String, ByVal objJson As String, _
                        ByVal noText As String, ByVal jsonKey As String, _
                        ByVal dispKey As String, ByVal enumList As String)
    Dim sVal As String
    sVal = modJsonLite.GetStr(objJson, jsonKey)
    If Not InEnum(sVal, enumList) Then
        Ap outText, "[V-S2-03] risk_no " & noText & " の " & dispKey & " が不正です: " & sVal
    End If
End Sub

' V-S2-17 の1キー分(emerging_risks[] の enum)。
Private Sub ChkEmergEnum(ByRef outText As String, ByVal objJson As String, _
                         ByVal idx As Long, ByVal jsonKey As String, _
                         ByVal enumList As String)
    Dim sVal As String
    sVal = modJsonLite.GetStr(objJson, jsonKey)
    If Not InEnum(sVal, enumList) Then
        Ap outText, "[V-S2-17] emerging_risks[" & idx & "] の " & jsonKey & _
                    " が不正です: " & sVal
    End If
End Sub

' V-S2-07 / V-S2-08 のバンド判定。isFreq=True は frequency(low/mid/high)、
' False は impact(small/mid/large)。低位バンド={1,2} / mid=3 / 高位バンド={4,5}。
Private Function ScoreFits(ByVal scoreVal As Long, ByVal bandText As String, _
                           ByVal isFreq As Boolean) As Boolean
    Dim lowName As String
    Dim highName As String
    If isFreq Then
        lowName = "low"
        highName = "high"
    Else
        lowName = "small"
        highName = "large"
    End If
    If bandText = lowName Then
        ScoreFits = (scoreVal = 1 Or scoreVal = 2)
    ElseIf bandText = "mid" Then
        ScoreFits = (scoreVal = 3)
    ElseIf bandText = highName Then
        ScoreFits = (scoreVal = 4 Or scoreVal = 5)
    Else
        ' バンド自体が enum 外のときは V-S2-03 が報告するため、ここでは黙る。
        ScoreFits = True
    End If
End Function

' JSONの最外オブジェクト直下(深さ1)に当該キーがあるか(V-S1-01)。
' 入れ子の同名キーを「あった」と数えないため、modJsonLite ではなくここで走査する。
Private Function TopKeyExists(ByVal srcJson As String, ByVal keyName As String) As Boolean
    Dim n As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim endPos As Long
    Dim rawKey As String

    n = Len(srcJson)
    i = 1
    Do While i <= n
        ch = Mid$(srcJson, i, 1)
        If ch = """" Then
            endPos = StrEndPos(srcJson, i)
            If endPos = 0 Then Exit Function
            rawKey = Mid$(srcJson, i + 1, endPos - i - 1)
            i = endPos + 1
            Do While i <= n
                If InStr(" " & vbTab & vbCr & vbLf, Mid$(srcJson, i, 1)) = 0 Then Exit Do
                i = i + 1
            Loop
            If i <= n Then
                If Mid$(srcJson, i, 1) = ":" And depth = 1 And rawKey = keyName Then
                    TopKeyExists = True
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

' 開き引用符 startPos に対応する閉じ引用符の位置(無ければ0)。
Private Function StrEndPos(ByVal srcJson As String, ByVal startPos As Long) As Long
    Dim n As Long
    Dim i As Long
    Dim ch As String
    n = Len(srcJson)
    i = startPos + 1
    Do While i <= n
        ch = Mid$(srcJson, i, 1)
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
