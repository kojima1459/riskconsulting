Attribute VB_Name = "modUICase2"
Option Explicit

' ============================================================================
' modUICase2 - S1～S4シートの描画と逆シリアライズ(ui層・T-31)
' ----------------------------------------------------------------------------
' 30,000字契約(12章§2)による modUICase の分割先。**呼んでよいのは modUICase /
' modUIHome** であり、変換表(19章§3)は modUICase から引く(表を2箇所に持たない)。
'
' 本モジュールが持つ規約:
'   ・13章§2.12-§2.15 の列定義(Cols*)と 13章§2.2 のセル格納規約の変換関数は
'     modUICaseFmt が唯一持ち、本モジュールはそれを**書き出しと逆シリアライズの
'     両方で同じように読む**(片方だけ直すことができない)。
'   ・13章§2.2 逆シリアライズ規約5「空行が表の終端」・規約2「順序列の昇順」・
'     規約3「順序列はJSONへ書き戻さない」。
'   ・保存前に必ず modValidate.CheckSN を通し、不合格なら sN_edited を保存しない
'     (規約4)。下流は従前の参照優先(sNr_json / sN_json)を使い続ける。
'
' ブロックの行数(部屋)は**次のブロックのアンカー行から動的に決める**。13章§2.9が
' 「行番号を仮定しない」と定めているため、確保行数を定数で持たない(持つと台帳の
' reserve_rows を変えた瞬間に隣のブロックの見出しを消す)。
' ============================================================================

Private Const U2_SRC As String = "modUICase2"
Private Const U2_WARN As String = "hm_warning"    ' 13章§2.10 HOMEの警告欄
Private Const U2_MSG_MISMATCH As String = "画面の案件と保存先が一致しません。再描画してください"

' 画面制御用のモジュール変数(裁定書9 B6。永続でない**描画状態のフラグ**であり、
' 14章§6の「状態保持の例外」への登録は要らない)。
'   gDrawStep    = いま DrawStep が描いている Step 番号(DrawArrBlock への文脈)
'   gTruncStep() = 直近の描画で部屋あふれ(切捨て)が起きた Step
'   gTruncNote   = 切捨ての内訳(利用者向け文言の材料)
Private gDrawStep As Long
Private gTruncStep(1 To 4) As Boolean
Private gTruncNote As String
Private Const U2_MAX_ROOM As Long = 200        ' 最終ブロックの部屋の上限
Private Const U2_HDR_WIDTH As Long = 24        ' 見出し行の探索幅(列番号ではない)
Private Const U2_FIRST_COL As Long = 1         ' ブロックの左端列(build/sheets_main.json)

Public Function SheetNameOf(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 1
        SheetNameOf = "S1_企業プロファイル"
    Case 2
        SheetNameOf = "S2_リスク仮説"
    Case 3
        SheetNameOf = "S3_提案"
    Case 4
        SheetNameOf = "S4_骨子"
    End Select
End Function

' ============================================================================
' SerializeStep - シート -> JSON(13章§2.2 逆シリアライズ規約)
'   組めなければ ""(捏造しない)。
' ============================================================================
Public Function SerializeStep(ByVal stepNo As Long) As String
    On Error GoTo Failed
    Select Case stepNo
    Case 1
        SerializeStep = SerializeS1()
    Case 2
        SerializeStep = SerializeS2()
    Case 3
        SerializeStep = SerializeS3()
    Case 4
        SerializeStep = SerializeS4()
    End Select
    Exit Function
Failed:
    modLog.LogError "E0302", U2_SRC & ".SerializeStep", _
                    "serialize_failed:s" & CStr(stepNo), Err.Number
    SerializeStep = vbNullString
End Function

' S1: 15章 SchemaS1 のプロパティ順に組む(13章§2.2 規約1)。
Private Function SerializeS1() As String
    Dim basicSpec As String
    basicSpec = modUICaseFmt.ColsS1Basic()

    Dim vals As Variant
    vals = ReadSingleRow("s1_basic", basicSpec)
    If IsEmpty(vals) Then Exit Function

    Dim s As String
    s = ""
    modUICaseFmt.AddFrag s, "company_name", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "company_name"))
    modUICaseFmt.AddFrag s, "business_summary", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "business_summary"))
    modUICaseFmt.AddFrag s, "main_products", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "main_products"))
    modUICaseFmt.AddFrag s, "processes", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "processes"))
    modUICaseFmt.AddFrag s, "locations", _
            BlockArrJson("s1_locations", modUICaseFmt.ColsS1Locations())

    Dim sub1 As String
    sub1 = ""
    modUICaseFmt.AddFrag sub1, "key_materials", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "supply_chain_key_materials"))
    modUICaseFmt.AddFrag sub1, "notes", modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "supply_chain_notes"))
    modUICaseFmt.AddFrag s, "supply_chain", "{" & sub1 & "}"

    sub1 = ""
    modUICaseFmt.AddFrag sub1, "segments", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "customers_segments"))
    modUICaseFmt.AddFrag sub1, "channels", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "customers_channels"))
    modUICaseFmt.AddFrag s, "customers", "{" & sub1 & "}"

    modUICaseFmt.AddFrag s, "workforce_notes", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "workforce_notes"))
    modUICaseFmt.AddFrag s, "management_notes", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "management_notes"))

    sub1 = ""
    modUICaseFmt.AddFrag sub1, "mvv", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "mvv"))
    modUICaseFmt.AddFrag sub1, "aspirations", _
            modUICaseFmt.ArrJson(ValueOf(basicSpec, vals, "aspirations"))
    modUICaseFmt.AddFrag sub1, "market_context", _
            modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "market_context"))
    modUICaseFmt.AddFrag s, "strategy_outlook", "{" & sub1 & "}"

    modUICaseFmt.AddFrag s, "current_coverage", _
            BlockArrJson("s1_current_coverage", modUICaseFmt.ColsS1Coverage())
    modUICaseFmt.AddFrag s, "field_insights", _
            BlockArrJson("s1_field_insights", modUICaseFmt.ColsS1Insights())
    modUICaseFmt.AddFrag s, "missing_info", _
            BlockArrJson("s1_missing_info", modUICaseFmt.ColsS1Missing())

    sub1 = ""
    modUICaseFmt.AddFrag sub1, "coverage", _
            BlockArrJson("s1_input_quality", modUICaseFmt.ColsS1Quality())
    modUICaseFmt.AddFrag sub1, "overall", _
            modUICaseFmt.EnumJson("input_quality_overall", ValueOf(basicSpec, vals, "input_quality_overall"))
    modUICaseFmt.AddFrag sub1, "advice", modUICaseFmt.StrJson(ValueOf(basicSpec, vals, "input_quality_advice"))
    modUICaseFmt.AddFrag s, "input_quality", "{" & sub1 & "}"

    modUICaseFmt.AddFrag s, "research_requests", _
            BlockArrJson("s1_research_requests", modUICaseFmt.ColsS1Research())
    SerializeS1 = "{" & s & "}"
End Function

Private Function SerializeS2() As String
    Dim s As String
    s = ""
    modUICaseFmt.AddFrag s, "risks", BlockArrJson("s2_risks", modUICaseFmt.ColsS2Risks())
    modUICaseFmt.AddFrag s, "gaps", BlockArrJson("s2_gaps", modUICaseFmt.ColsS2Gaps())
    modUICaseFmt.AddFrag s, "emerging_risks", BlockArrJson("s2_emerging", modUICaseFmt.ColsS2Emerging())
    modUICaseFmt.AddFrag s, "open_questions", ScalarArrJson("s2_open_questions", "question")
    SerializeS2 = "{" & s & "}"
End Function

Private Function SerializeS3() As String
    Dim s As String
    s = ""
    modUICaseFmt.AddFrag s, "stories", BlockArrJson("s3_stories", modUICaseFmt.ColsS3Stories())
    modUICaseFmt.AddFrag s, "unmatched_risks", BlockArrJson("s3_unmatched_risks", modUICaseFmt.ColsS3Unmatched())
    modUICaseFmt.AddFrag s, "do_not_propose", BlockArrJson("s3_do_not_propose", modUICaseFmt.ColsS3DoNot())
    SerializeS3 = "{" & s & "}"
End Function

Private Function SerializeS4() As String
    Dim vals As Variant
    vals = ReadSingleRow("s4_meta", modUICaseFmt.ColsS4Meta())
    If IsEmpty(vals) Then Exit Function

    Dim s As String
    s = ""
    modUICaseFmt.AddFrag s, "file_title", modUICaseFmt.StrJson(ValueOf(modUICaseFmt.ColsS4Meta(), vals, "file_title"))
    modUICaseFmt.AddFrag s, "slides", BlockArrJson("s4_slides", modUICaseFmt.ColsS4Slides())
    modUICaseFmt.AddFrag s, "hearing_questions", BlockArrJson("s4_hearing_questions", modUICaseFmt.ColsS4Hearing())
    SerializeS4 = "{" & s & "}"
End Function

' ============================================================================
' SaveEditedStep - 13章§2.2 逆シリアライズ規約4
'   シート -> JSON -> CheckSN -> 合格なら sN_edited を保存。
'   不合格は保存せず errText に不合格理由(呼び出し側が画面へ出す)。
' ============================================================================
Public Function SaveEditedStep(ByVal caseId As String, ByVal stepNo As Long, _
                               ByRef errText As String) As Boolean
    On Error GoTo Failed
    errText = vbNullString

    If stepNo < 1 Or stepNo > 4 Then Exit Function
    If Not modCaseStore.IsValidCaseId(caseId) Then
        errText = "案件IDが不正です。"
        Exit Function
    End If

    ' 裁定書9 B1: 画面が別案件の内容のまま sN_edited を確定させない。読取専用の
    ' 案件ID表示セル(13章§2.12 sN_case_id)と引数が一致しなければ1セルも保存しない。
    ' 呼び出し側で二重に文言を出さないよう errText は空のままにし、画面へ直接出す。
    If Trim$(modUISheet.ReadNamed(CaseIdCellOf(stepNo))) <> caseId Then
        modUISheet.WriteNamed U2_WARN, U2_MSG_MISMATCH
        modLog.LogError "E0302", U2_SRC & ".SaveEditedStep", _
                        "case_id_mismatch:s" & CStr(stepNo)
        Exit Function
    End If

    ' 裁定書9 B6: 部屋あふれで表示しきれていない行があるあいだは保存しない
    ' (切り詰められた画面が参照優先の最上位に居座るのを防ぐ)。
    If gTruncStep(stepNo) Then
        modUISheet.WriteNamed U2_WARN, "表示しきれていない行があるため、Step" & _
            CStr(stepNo) & " の編集は保存しませんでした（" & gTruncNote & "）。"
        Exit Function
    End If

    Dim jsonText As String
    jsonText = SerializeStep(stepNo)
    If LenB(jsonText) = 0 Then
        errText = "画面から内容を読み取れませんでした（ブロックの見出しをご確認ください）。"
        Exit Function
    End If

    errText = ValidateEdited(caseId, stepNo, jsonText)
    If LenB(errText) > 0 Then
        modLog.LogError "E0302", U2_SRC & ".SaveEditedStep", _
                        "edited_rejected:s" & CStr(stepNo)
        Exit Function
    End If

    If Not modCaseStore.SaveData(caseId, "s" & CStr(stepNo) & "_edited", jsonText) Then
        errText = "編集内容を保存できませんでした。"
        Exit Function
    End If

    modLog.LogUsage "sheet_edited_saved", caseId, "step=s" & CStr(stepNo)

    ' 裁定書9 B14(16章NFR-S7 設計原則(2)): 一次のfail-closed(注入テキスト照合)を
    ' 残したまま、ナレッジブック本体との**二次照合**を後段で行う。不一致は保存を
    ' ブロックせず hm_warning へ出す(注入の切詰めやKB更新でも差分は生じうるため)。
    Dim unknownIds As String
    unknownIds = UnknownKbIds(stepNo, jsonText)
    If LenB(unknownIds) > 0 Then
        modUISheet.WriteNamed U2_WARN, "ナレッジブックに見当たらないIDがあります（" & _
            modUtil.SafeLeft(unknownIds, 200) & "）。内容をご確認ください。"
    End If

    SaveEditedStep = True
    Exit Function

Failed:
    errText = "編集内容の保存中にエラーが発生しました。"
    modLog.LogError "E0302", U2_SRC & ".SaveEditedStep", "save_failed", Err.Number
    SaveEditedStep = False
End Function

' 編集JSONの検証。modPipeline の防衛線(3)と同じ Check* を同じ材料で呼ぶ
' (ID実在検査は fail-closed なので一覧テキストを必ず渡す。16章E-07)。
Private Function ValidateEdited(ByVal caseId As String, ByVal stepNo As Long, _
                                ByVal jsonText As String) As String
    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String
    Dim s4Variant As String
    Dim tierText As String

    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        ValidateEdited = "案件一覧からこの案件を読めませんでした。"
        Exit Function
    End If

    Select Case stepNo
    Case 1
        ValidateEdited = modValidate.CheckS1(jsonText, ctx.case_type, _
            (LenB(Trim$(modCaseStore.LoadData(caseId, "input_field_notes"))) > 0))
    Case 2
        ValidateEdited = modValidate.CheckS2(jsonText, ctx.case_type, _
            modKnowledge.MenusSummaryFor(ctx.industry_code, 0), _
            modCaseStore.LoadData(caseId, "s2_prev_json"))
    Case 3
        ValidateEdited = modValidate.CheckS3(jsonText, _
            modCaseStore.ResolveStepJson(caseId, 2), _
            modKnowledge.MenusFor(ctx.industry_code, 0), _
            modKnowledge.LinesText(0), _
            modKnowledge.SchemesFor(ctx.industry_code, 0), _
            modKnowledge.CasesFor(ctx.industry_code, 0), ctx.case_type)
    Case 4
        ValidateEdited = modValidate.CheckS4(jsonText, tierText, _
            modConfig.GetLong("ppt_max_slides_t2", 10))
    End Select
End Function

' 二次照合(裁定書9 B14)。JSONに現れるナレッジIDのうち modKnowledge 本体に実在
'   しないものを ";" 区切りで返す(全て実在なら "")。走査表は
'   `配列キー|要素内のキー(空=要素そのものがID)|種別` の vbLf 区切り。
Private Function UnknownKbIds(ByVal stepNo As Long, ByVal jsonText As String) As String
    On Error GoTo Skip1

    Dim spec As String
    If stepNo = 2 Then
        spec = "risks>preventions|related_menu_id|m"
    ElseIf stepNo = 3 Then
        spec = "stories|menu_ids|m" & vbLf & "stories|line_ids|l" & vbLf & _
               "stories|scheme_id|s" & vbLf & "stories|similar_case_id|c"
    Else
        Exit Function
    End If

    Dim buf() As String
    Dim cnt As Long
    modUtil.BufInit buf, cnt

    Dim rows1() As String
    rows1 = Split(spec, vbLf)
    Dim i As Long
    For i = LBound(rows1) To UBound(rows1)
        ScanIds jsonText, rows1(i), buf, cnt
    Next i

    UnknownKbIds = Replace(modUtil.BufText(buf, cnt), vbLf, ";")
    Exit Function
Skip1:
    UnknownKbIds = vbNullString
End Function

' 走査表1行ぶん。外側配列(">"で入れ子を1段)の各要素からIDを取り出して照合する。
Private Sub ScanIds(ByVal jsonText As String, ByVal rowText As String, _
                    ByRef buf() As String, ByRef cnt As Long)
    Dim flds() As String
    flds = Split(rowText, "|")
    If UBound(flds) - LBound(flds) <> 2 Then Exit Sub

    Dim outerKey As String
    Dim innerKey As String
    Dim keyName As String
    Dim kindText As String
    outerKey = flds(LBound(flds))
    keyName = flds(LBound(flds) + 1)
    kindText = flds(LBound(flds) + 2)

    Dim p As Long
    p = InStr(1, outerKey, ">", vbBinaryCompare)
    If p > 0 Then
        innerKey = Mid$(outerKey, p + 1)
        outerKey = Left$(outerKey, p - 1)
    End If

    Dim outer1 As Collection
    Set outer1 = modJsonLite.GetArrayItems(jsonText, outerKey)

    Dim i As Long
    For i = 1 To outer1.count
        If LenB(innerKey) > 0 Then
            Dim inner1 As Collection
            Dim j As Long
            Set inner1 = modJsonLite.GetArrayItems(CStr(outer1(i)), innerKey)
            For j = 1 To inner1.count
                AddIfMissing modJsonLite.GetStr(CStr(inner1(j)), keyName), kindText, buf, cnt
            Next j
        Else
            Dim ids As Collection
            Dim k As Long
            Set ids = modJsonLite.GetArrayItems(CStr(outer1(i)), keyName)
            If ids.count > 0 Then
                For k = 1 To ids.count
                    AddIfMissing CStr(ids(k)), kindText, buf, cnt
                Next k
            Else
                AddIfMissing modJsonLite.GetStr(CStr(outer1(i)), keyName), kindText, buf, cnt
            End If
        End If
    Next i
End Sub

' 種別ごとの実在検査(14章§6 modKnowledge の5本のうち画面に現れる4本)。
Private Sub AddIfMissing(ByVal idText As String, ByVal kindText As String, _
                         ByRef buf() As String, ByRef cnt As Long)
    Dim id1 As String
    id1 = Trim$(idText)
    If LenB(id1) = 0 Then Exit Sub

    Dim okId As Boolean
    Select Case kindText
    Case "m"
        okId = modKnowledge.MenuIdExists(id1)
    Case "l"
        okId = modKnowledge.LineIdExists(id1)
    Case "s"
        okId = modKnowledge.SchemeIdExists(id1)
    Case Else
        okId = modKnowledge.CaseLibIdExists(id1)
    End Select
    If Not okId Then modUtil.BufAdd buf, cnt, id1
End Sub

' 13章§2.12: S1～S4の読取専用の案件ID表示セル(名前付きレンジ)。
Private Function CaseIdCellOf(ByVal stepNo As Long) As String
    CaseIdCellOf = "s" & CStr(stepNo) & "_case_id"
End Function

' ============================================================================
' DrawStep - JSON -> シート(参照優先の解決は modCaseStore.ResolveStepJson)
' ============================================================================
Public Function DrawStep(ByVal caseId As String, ByVal stepNo As Long) As Boolean
    On Error GoTo Failed

    If stepNo < 1 Or stepNo > 4 Then Exit Function
    If Not modCaseStore.IsValidCaseId(caseId) Then Exit Function

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String
    Dim s4Variant As String
    Dim tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        Exit Function
    End If

    ' 裁定書9 B1/B6: この描画で「どの案件を」「切り詰めずに」描けたかを記録する。
    gDrawStep = stepNo
    gTruncStep(stepNo) = False
    gTruncNote = vbNullString
    modUISheet.WriteNamed CaseIdCellOf(stepNo), caseId

    Dim jsonText As String
    jsonText = modCaseStore.ResolveStepJson(caseId, stepNo)
    If LenB(Trim$(jsonText)) = 0 Then
        modLog.LogUsage "sheet_draw_skipped", caseId, "step=s" & CStr(stepNo) & ";empty"
        Exit Function
    End If

    ' 16章E-31: 画面表示の直前に実名へ戻す({{POLICY_NO}} は復元表が無いので残す)。
    jsonText = modUICase.RestoreNames(jsonText, ctx.company)

    Select Case stepNo
    Case 1
        DrawS1 jsonText
    Case 2
        DrawS2 jsonText
    Case 3
        DrawS3 jsonText
    Case 4
        DrawS4 jsonText
    End Select

    modLog.LogUsage "sheet_drawn", caseId, "step=s" & CStr(stepNo)
    DrawStep = True
    Exit Function

Failed:
    modLog.LogError "E0603", U2_SRC & ".DrawStep", "draw_failed:s" & CStr(stepNo), Err.Number
    DrawStep = False
End Function

Private Sub DrawS1(ByVal jsonText As String)
    DrawSingleRow "s1_basic", modUICaseFmt.ColsS1Basic(), jsonText
    DrawArrBlock "s1_locations", modUICaseFmt.ColsS1Locations(), jsonText, "locations"
    DrawArrBlock "s1_current_coverage", modUICaseFmt.ColsS1Coverage(), jsonText, "current_coverage"
    DrawArrBlock "s1_field_insights", modUICaseFmt.ColsS1Insights(), jsonText, "field_insights"
    DrawArrBlock "s1_missing_info", modUICaseFmt.ColsS1Missing(), jsonText, "missing_info"
    DrawArrBlock "s1_input_quality", modUICaseFmt.ColsS1Quality(), _
                 modUICaseFmt.SubJson(jsonText, "input_quality"), "coverage"
    DrawArrBlock "s1_research_requests", modUICaseFmt.ColsS1Research(), jsonText, "research_requests"
    DrawQualityBanner jsonText
    DrawResearchButtons
End Sub

Private Sub DrawS2(ByVal jsonText As String)
    DrawArrBlock "s2_gaps", modUICaseFmt.ColsS2Gaps(), jsonText, "gaps"
    DrawArrBlock "s2_risks", modUICaseFmt.ColsS2Risks(), jsonText, "risks"
    DrawScalarArr "s2_open_questions", "question", jsonText, "open_questions"
    DrawArrBlock "s2_emerging", modUICaseFmt.ColsS2Emerging(), jsonText, "emerging_risks"
End Sub

Private Sub DrawS3(ByVal jsonText As String)
    DrawArrBlock "s3_stories", modUICaseFmt.ColsS3Stories(), jsonText, "stories"
    DrawArrBlock "s3_unmatched_risks", modUICaseFmt.ColsS3Unmatched(), jsonText, "unmatched_risks"
    DrawArrBlock "s3_do_not_propose", modUICaseFmt.ColsS3DoNot(), jsonText, "do_not_propose"
End Sub

Private Sub DrawS4(ByVal jsonText As String)
    DrawSingleRow "s4_meta", modUICaseFmt.ColsS4Meta(), jsonText
    DrawArrBlock "s4_slides", modUICaseFmt.ColsS4Slides(), jsonText, "slides"
    DrawArrBlock "s4_hearing_questions", modUICaseFmt.ColsS4Hearing(), jsonText, "hearing_questions"
End Sub

' 11章: 実行後はS1の入力充足度診断がS1シート上部に表示される。13章§2.12は
'   この行に列も名前付きレンジも与えていないので、データ面を汚さないよう
'   ラベル図形で描く(表示専用)。
Private Sub DrawQualityBanner(ByVal jsonText As String)
    On Error Resume Next
    Dim ws As Object
    Set ws = modUISheet.SheetOf(SheetNameOf(1))
    If ws Is Nothing Then Exit Sub

    Dim iq As String
    iq = modUICaseFmt.SubJson(jsonText, "input_quality")

    Dim overallJa As String
    overallJa = modUICase.EnumJa("input_quality_overall", modJsonLite.GetStr(iq, "overall"))

    Dim bannerText As String
    bannerText = "充足度: " & overallJa & " ｜ 助言: " & modJsonLite.GetStr(iq, "advice")

    modUISheet.EnsureLabel ws, "lbl_s1_quality", bannerText, 1, 3, 460#, 16#
End Sub

' ============================================================================
' ブロック単位の読み書き
' ============================================================================

' 単一行ブロック(s1_basic / s4_meta)の1行を読む。読めなければ Empty。
Private Function ReadSingleRow(ByVal anchorName As String, ByVal colSpec As String) As Variant
    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Function

    Dim cols As Variant
    cols = ColIndexes(ws, headerRow, colSpec)
    If IsEmpty(cols) Then Exit Function

    ReadSingleRow = RowValues(ws, headerRow + 1, cols)
End Function

' 配列ブロックを読んでJSON配列本文にする。0行なら "[]"(13章§2.2 空配列規約)。
Private Function BlockArrJson(ByVal anchorName As String, ByVal colSpec As String) As String
    BlockArrJson = "[]"

    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Function

    Dim cols As Variant
    cols = ColIndexes(ws, headerRow, colSpec)
    If IsEmpty(cols) Then Exit Function

    Dim lastRow As Long
    lastRow = modUISheet.BlockLastRow(ws, headerRow, U2_FIRST_COL, _
                                      ColCount(colSpec), RoomOf(anchorName))

    ' 13章§2.2 逆シリアライズ規約2「行順=配列順。順序列の昇順に並べ替えてから
    ' 配列化する」。並べ替えの鍵は当該表が持つ順序列(risk_no / gap_no / story_no /
    ' slide_no / emg_no / seq)で、無ければ行順のまま。
    Dim keyCol As Long
    keyCol = modUICaseFmt.OrderColOf(colSpec)

    Dim items() As String
    Dim keys() As Long
    Dim n As Long
    ReDim items(0 To lastRow - headerRow)
    ReDim keys(0 To lastRow - headerRow)

    Dim r As Long
    Dim vals As Variant
    For r = headerRow + 1 To lastRow
        vals = RowValues(ws, r, cols)
        items(n) = modUICaseFmt.RowObjJson(colSpec, vals)
        If keyCol >= 0 Then
            keys(n) = CLng(Val(CStr(vals(keyCol + LBound(vals)))))
        Else
            keys(n) = n + 1
        End If
        n = n + 1
    Next r
    If n = 0 Then Exit Function

    modUICaseFmt.SortItems items, keys, n

    Dim acc As String
    Dim i As Long
    For i = 0 To n - 1
        If LenB(acc) > 0 Then acc = acc & ","
        acc = acc & items(i)
    Next i

    BlockArrJson = "[" & acc & "]"
End Function

' 文字列だけの配列ブロック(s2_open_questions)。
Private Function ScalarArrJson(ByVal anchorName As String, ByVal colName As String) As String
    ScalarArrJson = "[]"

    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Function

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, U2_FIRST_COL, U2_HDR_WIDTH)
    Dim colNo As Long
    colNo = modUISheet.ColOf(hdr, U2_FIRST_COL, colName)
    If colNo <= 0 Then Exit Function

    Dim lastRow As Long
    lastRow = modUISheet.BlockLastRow(ws, headerRow, U2_FIRST_COL, 2, RoomOf(anchorName))

    Dim seqCol As Long
    seqCol = modUISheet.ColOf(hdr, U2_FIRST_COL, "seq")

    Dim items() As String
    Dim keys() As Long
    Dim n As Long
    ReDim items(0 To lastRow - headerRow)
    ReDim keys(0 To lastRow - headerRow)

    Dim r As Long
    Dim v As String
    For r = headerRow + 1 To lastRow
        v = modUISheet.CellText(ws, r, colNo)
        If LenB(Trim$(v)) > 0 Then
            items(n) = modUICaseFmt.StrJson(v)
            keys(n) = n + 1
            If seqCol > 0 Then keys(n) = CLng(Val(modUISheet.CellText(ws, r, seqCol)))
            n = n + 1
        End If
    Next r
    If n = 0 Then Exit Function

    modUICaseFmt.SortItems items, keys, n

    Dim acc As String
    Dim i As Long
    For i = 0 To n - 1
        If LenB(acc) > 0 Then acc = acc & ","
        acc = acc & items(i)
    Next i

    ScalarArrJson = "[" & acc & "]"
End Function

' 単一行ブロックへ書く。
Private Sub DrawSingleRow(ByVal anchorName As String, ByVal colSpec As String, _
                          ByVal jsonText As String)
    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Sub

    Dim cols As Variant
    cols = ColIndexes(ws, headerRow, colSpec)
    If IsEmpty(cols) Then Exit Sub

    WriteRow ws, headerRow + 1, cols, colSpec, jsonText, 1, anchorName
End Sub

' 配列ブロックへ書く(部屋を超える分は書かない。書けなかった件数は記録する)。
Private Sub DrawArrBlock(ByVal anchorName As String, ByVal colSpec As String, _
                         ByVal jsonText As String, ByVal arrayKey As String)
    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Sub

    Dim cols As Variant
    cols = ColIndexes(ws, headerRow, colSpec)
    If IsEmpty(cols) Then Exit Sub

    Dim room As Long
    room = RoomOf(anchorName)
    modUISheet.ClearBlock ws, headerRow, U2_FIRST_COL, ColCount(colSpec), room

    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(jsonText, arrayKey)

    Dim i As Long
    For i = 1 To items.count
        If i > room Then Exit For
        WriteRow ws, headerRow + i, cols, colSpec, CStr(items(i)), i, anchorName
    Next i

    If items.count > room Then
        modLog.LogError "E0604", U2_SRC & ".DrawArrBlock", _
                        anchorName & ":rows=" & CStr(items.count) & ";room=" & CStr(room)
        NoteTruncation anchorName, items.count, room
    End If
End Sub

' 裁定書9 B6: 部屋あふれを画面へ出し、当該Stepの保存をブロックする印を立てる。
Private Sub NoteTruncation(ByVal anchorName As String, ByVal total1 As Long, ByVal room As Long)
    On Error Resume Next

    If gDrawStep >= 1 And gDrawStep <= 4 Then gTruncStep(gDrawStep) = True

    Dim one As String
    one = anchorName & ": " & CStr(total1) & "件のうち" & CStr(room) & "件しか表示できていません"
    If LenB(gTruncNote) > 0 Then gTruncNote = gTruncNote & " ／ "
    gTruncNote = gTruncNote & one

    modUISheet.WriteNamed U2_WARN, gTruncNote & _
        "。編集を保存すると残りが失われるため、この画面の保存は行いません。"
End Sub

' 文字列配列ブロックへ書く。
Private Sub DrawScalarArr(ByVal anchorName As String, ByVal colName As String, _
                          ByVal jsonText As String, ByVal arrayKey As String)
    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Sub

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, U2_FIRST_COL, U2_HDR_WIDTH)

    Dim seqCol As Long
    Dim colNo As Long
    seqCol = modUISheet.ColOf(hdr, U2_FIRST_COL, "seq")
    colNo = modUISheet.ColOf(hdr, U2_FIRST_COL, colName)
    If colNo <= 0 Then Exit Sub

    Dim room As Long
    room = RoomOf(anchorName)
    modUISheet.ClearBlock ws, headerRow, U2_FIRST_COL, 2, room

    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(jsonText, arrayKey)

    Dim i As Long
    For i = 1 To items.count
        If i > room Then Exit For
        If seqCol > 0 Then
            modUISheet.PutText ws, headerRow + i, seqCol, CStr(i), anchorName & "/seq"
        End If
        modUISheet.PutText ws, headerRow + i, colNo, CStr(items(i)), _
                           anchorName & "/" & colName
    Next i
End Sub

' 1行ぶんを書く(順序列は行番号、他はスキーマ写像から取り出す)。
Private Sub WriteRow(ByVal ws As Object, ByVal rowNo As Long, ByVal cols As Variant, _
                     ByVal colSpec As String, ByVal itemJson As String, _
                     ByVal ordinal As Long, ByVal whereNote As String)
    Dim specs() As String
    specs = Split(colSpec, ";")

    Dim i As Long
    Dim physName As String
    Dim kindText As String
    Dim extraText As String
    Dim pathText As String
    Dim cellValue As String

    For i = LBound(specs) To UBound(specs)
        modUICaseFmt.SplitCol specs(i), physName, kindText, extraText, pathText
        If kindText = "o" Then
            cellValue = CStr(ordinal)
        Else
            cellValue = modUICaseFmt.CellFromJson(kindText, extraText, pathText, itemJson)
        End If
        modUISheet.PutText ws, rowNo, CLng(cols(i - LBound(specs))), cellValue, _
                           whereNote & "/" & physName
    Next i
End Sub

' 列定義に対応する絶対列番号の配列。1つでも見つからなければ Empty。
Private Function ColIndexes(ByVal ws As Object, ByVal headerRow As Long, _
                            ByVal colSpec As String) As Variant
    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, U2_FIRST_COL, U2_HDR_WIDTH)
    If IsEmpty(hdr) Then Exit Function

    Dim specs() As String
    specs = Split(colSpec, ";")

    Dim cols() As Long
    ReDim cols(0 To UBound(specs) - LBound(specs))

    Dim i As Long
    Dim physName As String
    Dim kindText As String
    Dim extraText As String
    Dim pathText As String
    For i = LBound(specs) To UBound(specs)
        modUICaseFmt.SplitCol specs(i), physName, kindText, extraText, pathText
        cols(i - LBound(specs)) = modUISheet.ColOf(hdr, U2_FIRST_COL, physName)
        If cols(i - LBound(specs)) <= 0 Then Exit Function
    Next i

    ColIndexes = cols
End Function

' 1行の値を列定義順に読む。
Private Function RowValues(ByVal ws As Object, ByVal rowNo As Long, _
                           ByVal cols As Variant) As Variant
    Dim vals() As String
    ReDim vals(LBound(cols) To UBound(cols))

    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        vals(i) = modUISheet.CellText(ws, rowNo, CLng(cols(i)))
    Next i
    RowValues = vals
End Function

' 列定義の物理名で値を引く(位置ではなく名前で引く=列順の変更に強い)。
Private Function ValueOf(ByVal colSpec As String, ByVal vals As Variant, _
                         ByVal physWanted As String) As String
    Dim specs() As String
    specs = Split(colSpec, ";")

    Dim i As Long
    Dim physName As String
    Dim kindText As String
    Dim extraText As String
    Dim pathText As String
    For i = LBound(specs) To UBound(specs)
        modUICaseFmt.SplitCol specs(i), physName, kindText, extraText, pathText
        If physName = physWanted Then
            ValueOf = CStr(vals(i - LBound(specs) + LBound(vals)))
            Exit Function
        End If
    Next i
End Function

Private Function ColCount(ByVal colSpec As String) As Long
    Dim specs() As String
    specs = Split(colSpec, ";")
    ColCount = UBound(specs) - LBound(specs) + 1
End Function

' ブロックの「部屋」(見出しの下に使ってよい行数)。次のブロックの見出し行の
'   1行手前(空行=表の終端)までを上限にする。最後のブロックは U2_MAX_ROOM。
Private Function RoomOf(ByVal anchorName As String) As Long
    RoomOf = U2_MAX_ROOM

    Dim stepNo As Long
    Dim names() As String
    Dim i As Long
    Dim here As Long
    Dim nxt As Long

    For stepNo = 1 To 4
        names = Split(modUICaseFmt.AnchorsOf(stepNo), ";")
        For i = LBound(names) To UBound(names)
            If names(i) = anchorName Then
                If i < UBound(names) Then
                    here = modUISheet.BlockRow(anchorName)
                    nxt = modUISheet.BlockRow(names(i + 1))
                    If here > 0 And nxt > here Then
                        RoomOf = nxt - here - 2      ' 間の空行を1行残す
                        If RoomOf < 1 Then RoomOf = 1
                    End If
                End If
                Exit Function
            End If
        Next i
    Next stepNo
End Function

' ============================================================================
' S1～S4シートの図形ボタン(11章§5・11章§2の各ワイヤー)
' ----------------------------------------------------------------------------
' 実行系のハンドラは modUIHome が持つ(HOMEの[S1][S2][S3][S4]と同じ動作を
' 各シートのボタンからも起こすだけなので、実装を2箇所に置かない)。
' ============================================================================
Public Sub EnsureStepButtons()
    On Error Resume Next

    Dim ws As Object

    Set ws = modUISheet.SheetOf(SheetNameOf(1))
    modUISheet.EnsureButton ws, "btn_s1_rerun", "S1から再実行", 1, 1, 84#, "modUIHome.HomeRunS1"
    modUISheet.EnsureButton ws, "btn_s1_next", "S2へ進む", 1, 2, 84#, "modUIHome.HomeRunS2"

    Set ws = modUISheet.SheetOf(SheetNameOf(2))
    modUISheet.EnsureButton ws, "btn_s2_rerun", "S2から再実行", 1, 1, 84#, "modUIHome.HomeRunS2"
    modUISheet.EnsureButton ws, "btn_s2_next", "S3へ進む", 1, 2, 84#, "modUIHome.HomeRunS3"

    Set ws = modUISheet.SheetOf(SheetNameOf(3))
    modUISheet.EnsureButton ws, "btn_s3_rerun", "S3から再実行", 1, 1, 84#, "modUIHome.HomeRunS3"
    modUISheet.EnsureButton ws, "btn_s3_next", "S4へ進む", 1, 2, 84#, "modUIHome.HomeRunS4"

    Set ws = modUISheet.SheetOf(SheetNameOf(4))
    modUISheet.EnsureButton ws, "btn_s4_rerun", "S4から再実行", 1, 1, 84#, "modUIHome.HomeRunS4"
    modUISheet.EnsureButton ws, "btn_s4_hearing", "ヒアリングシート生成", 1, 2, 128#, _
                            "modUIHome.HomeBuildHearing"
End Sub

' S1の追加収集ブロックの各行に[コピー]ボタンを置く(11章。案件入力側と同じ動作)。
Private Sub DrawResearchButtons()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.BlockSheet("s1_research_requests")
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow("s1_research_requests")
    If headerRow <= 0 Then Exit Sub

    modUISheet.DropShapesByPrefix ws, "btncopy_"

    Dim room As Long
    room = RoomOf("s1_research_requests")

    Dim lastRow As Long
    lastRow = modUISheet.BlockLastRow(ws, headerRow, U2_FIRST_COL, 3, room)

    Dim r As Long
    For r = headerRow + 1 To lastRow
        modUISheet.EnsureButton ws, "btncopy_" & CStr(r), "コピー", r, 4, 52#, _
                                "modUICase3.CopyResearchRow"
    Next r
End Sub
