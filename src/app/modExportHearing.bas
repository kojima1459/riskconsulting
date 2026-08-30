Attribute VB_Name = "modExportHearing"
Option Explicit

' ==========================================================
' modExportHearing - ヒアリングシートの生成(13章§2.16・14章§6)
' ------------------------------------------------
' 責務: S4 の `hearing_questions`(参照優先は13章§2.2に従い s4_edited > s4_json)
'   から、本体ブックの `ヒアリングシート`(A4縦印刷用)を整形生成する。
'   **LLMを呼ばない**(整形のみ)。HTMLレポート(18章)とは別物で、あちらは
'   S1+S2+S3だけから描き S4 を読まない(14章§6・18章§3の注記)。
'
' 書き込むもの:
'   印刷ヘッダ(名前付きレンジ) hs_case_id / hs_company / hs_industry_name /
'     hs_printed_at。**hs_visit_date は手入力欄なので触らない**。
'   ブロック `hearing_questions`(アンカー=見出し行の名前付きレンジ)の
'     q_no / question / purpose。**answer_memo は生成直後は必ず空**にする
'     (訪問時の記入欄。印刷して手書きする運用も想定)。
'
' 本シートから case_data への自動書き戻しは行わない(回答の取捨は人が行う。
'   13章§2.16)。次ラウンドへの経路は案件入力の「ヒアリング回答」欄への貼付。
'
' R4(12章§2・§4): 本モジュールはExcelトークン許可モジュールの1つ。許可の幅は
'   「本体ブック内のヒアリングシート1枚の読み書き」に限る。案件データは
'   modCaseStore / modCaseRead 経由でしか取らない。
' NFR-S7①(16章): 外部由来テキスト(LLM出力)のセル書込は必ず
'   modUtilText.SetCellSafe を通す。定数・ヘッダの書込には ' SAFE:const を付す。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ==========================================================

Private Const EH_SRC As String = "modExportHearing"
Private Const EH_SHEET As String = "ヒアリングシート"
Private Const EH_BLOCK As String = "hearing_questions"
Private Const EH_MAX_Q As Long = 10
Private Const EH_HDR_WIDTH As Long = 6

' ==========================================================
' BuildHearingSheet - 14章§6の契約。True=生成できた / False=生成しなかった。
'   S4未実行・シート不在・ブロックアンカー不在はいずれも False(捏造しない)。
' ==========================================================
Public Function BuildHearingSheet(ByVal caseId As String) As Boolean
    On Error GoTo Failed

    If Not modCaseStore.IsValidCaseId(caseId) Then
        modLog.LogError "E0101", EH_SRC & ".BuildHearingSheet", "invalid_case_id"
        Exit Function
    End If

    Dim s4Text As String
    s4Text = modCaseStore.ResolveStepJson(caseId, 4)
    If LenB(Trim$(s4Text)) = 0 Then
        modLog.LogUsage "hearing_sheet_skipped", caseId, "s4_empty"
        Exit Function
    End If

    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(s4Text, EH_BLOCK)
    If items.count = 0 Then
        modLog.LogUsage "hearing_sheet_skipped", caseId, "no_questions"
        Exit Function
    End If

    Dim ws As Object
    Set ws = SheetOf(EH_SHEET)
    If ws Is Nothing Then
        modLog.LogError "E0603", EH_SRC & ".BuildHearingSheet", "sheet_missing"
        Exit Function
    End If

    Dim headerRow As Long
    headerRow = BlockHeaderRow(EH_BLOCK)
    If headerRow <= 0 Then
        modLog.LogError "E0603", EH_SRC & ".BuildHearingSheet", "anchor_missing"
        Exit Function
    End If

    Dim hdr As Variant
    hdr = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, EH_HDR_WIDTH)).Value

    Dim cNo As Long, cQ As Long, cP As Long, cA As Long
    cNo = modUtil.FindHeaderCol(hdr, "q_no")
    cQ = modUtil.FindHeaderCol(hdr, "question")
    cP = modUtil.FindHeaderCol(hdr, "purpose")
    cA = modUtil.FindHeaderCol(hdr, "answer_memo")
    If cNo <= 0 Or cQ <= 0 Or cP <= 0 Or cA <= 0 Then
        modLog.LogError "E0603", EH_SRC & ".BuildHearingSheet", "header_missing"
        Exit Function
    End If

    ' 前回生成分を消してから積み直す(古い設問が残って混ざらないようにする)。
    ClearRows ws, headerRow, cNo, cQ, cP, cA

    WriteHeaderFields caseId

    Dim wrote As Long
    Dim i As Long
    For i = 1 To items.count
        If wrote >= EH_MAX_Q Then Exit For
        Dim itemJson As String
        itemJson = CStr(items(i))
        Dim qText As String, pText As String
        qText = modJsonLite.GetStr(itemJson, "question")
        pText = modJsonLite.GetStr(itemJson, "purpose")
        If LenB(Trim$(qText)) > 0 Then
            wrote = wrote + 1
            Dim r As Long
            r = headerRow + wrote
            ws.Cells(r, cNo).Value = wrote ' SAFE:const
            modUtilText.SetCellSafe ws.Cells(r, cQ), qText, EH_SHEET & "/question"
            modUtilText.SetCellSafe ws.Cells(r, cP), pText, EH_SHEET & "/purpose"
        End If
    Next i

    If wrote = 0 Then
        modLog.LogUsage "hearing_sheet_skipped", caseId, "no_valid_questions"
        Exit Function
    End If

    modLog.LogUsage "hearing_sheet", caseId, "questions=" & CStr(wrote)
    BuildHearingSheet = True
    Exit Function

Failed:
    modLog.LogError "E0603", EH_SRC & ".BuildHearingSheet", "build_failed", Err.Number
    BuildHearingSheet = False
End Function

' ==========================================================
' AnswerMemoCount - 14章§6の契約(裁定書9 N8・B12、裁定書10 M5改)。
'   ヒアリングシートの answer_memo 列の**非空行数**を返す。シート不在・
'   アンカー不在・見出し不在の場合は 0。
'   裁定書10 M5: hs_case_id と caseId の不一致で 0 を返す旧仕様は廃止
'   (別案件の手書き回答が入った状態で確認なしに BuildHearingSheet が全行を
'   空へ戻す fail-open を塞ぐ)。本関数は案件を問わず常に数える。引数 caseId
'   は14章§6の契約シグネチャ維持のため残すが、判定には使わない。
'   呼出側(modUIHome)は modUISheet.ReadNamed("hs_case_id") を読み、
'   caseId と不一致かつ本関数>0のとき「別案件(<hs_case_id>)の手書き回答が
'   残っています」の確認文言にする(裁定書10 M5)。
'   本関数は数えるだけで、シートを1セルも書き換えない。
' ==========================================================
Public Function AnswerMemoCount(ByVal caseId As String) As Long
    On Error GoTo Zero0

    Dim ws As Object
    Set ws = SheetOf(EH_SHEET)
    If ws Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = BlockHeaderRow(EH_BLOCK)
    If headerRow <= 0 Then Exit Function

    Dim hdr As Variant
    hdr = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, EH_HDR_WIDTH)).Value

    Dim cA As Long
    cA = modUtil.FindHeaderCol(hdr, "answer_memo")
    If cA <= 0 Then Exit Function

    Dim n As Long
    Dim r As Long
    For r = headerRow + 1 To headerRow + EH_MAX_Q
        If LenB(Trim$(CStr(ws.Cells(r, cA).Value))) > 0 Then n = n + 1
    Next r
    AnswerMemoCount = n
    Exit Function

Zero0:
    AnswerMemoCount = 0
End Function

' 印刷ヘッダ(名前付きレンジ)。案件一覧からの複写3項目と生成日時。
'   hs_visit_date は手入力欄なので**書かない**(13章§2.16)。
Private Sub WriteHeaderFields(ByVal caseId As String)
    On Error Resume Next

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String, s4Variant As String, tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        Exit Sub
    End If

    PutNamed "hs_case_id", caseId
    PutNamed "hs_company", ctx.company
    PutNamed "hs_industry_name", ctx.industry_name
    PutNamed "hs_printed_at", modUtil.NowStamp()
End Sub

' 名前付きレンジ1点への書込。会社名は案件入力の自由記述なので SetCellSafe を通す。
Private Sub PutNamed(ByVal rangeName As String, ByVal valueText As String)
    On Error Resume Next
    Dim cell As Object
    Set cell = NamedCell(rangeName)
    If cell Is Nothing Then Exit Sub
    modUtilText.SetCellSafe cell, valueText, EH_SHEET & "/" & rangeName
End Sub

' ブロックの予約行(10行)ぶんを空へ戻す。answer_memo も消す(生成直後は必ず空)。
Private Sub ClearRows(ByVal ws As Object, ByVal headerRow As Long, ByVal cNo As Long, _
                      ByVal cQ As Long, ByVal cP As Long, ByVal cA As Long)
    On Error Resume Next
    Dim r As Long
    For r = headerRow + 1 To headerRow + EH_MAX_Q
        ws.Cells(r, cNo).Value = vbNullString ' SAFE:const
        ws.Cells(r, cQ).Value = vbNullString ' SAFE:const
        ws.Cells(r, cP).Value = vbNullString ' SAFE:const
        ws.Cells(r, cA).Value = vbNullString ' SAFE:const
    Next r
End Sub

Private Function SheetOf(ByVal sheetName As String) As Object
    On Error GoTo NoSheet
    Set SheetOf = ThisWorkbook.Worksheets(sheetName)
    Exit Function
NoSheet:
    Set SheetOf = Nothing
End Function

' ブロックアンカー(名前付きレンジ)が指す見出し行の行番号。不在は0。
'   13章§2.9: 複数ブロックのシートは行番号を仮定せずアンカーで解決する。
Private Function BlockHeaderRow(ByVal anchorName As String) As Long
    On Error GoTo NoAnchor
    Dim cell As Object
    Set cell = NamedCell(anchorName)
    If cell Is Nothing Then Exit Function
    BlockHeaderRow = cell.row
    Exit Function
NoAnchor:
    BlockHeaderRow = 0
End Function

Private Function NamedCell(ByVal rangeName As String) As Object
    On Error GoTo NoName
    Set NamedCell = ThisWorkbook.Names(rangeName).RefersToRange.Cells(1, 1)
    Exit Function
NoName:
    Set NamedCell = Nothing
End Function
