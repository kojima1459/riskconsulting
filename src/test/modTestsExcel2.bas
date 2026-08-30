Attribute VB_Name = "modTestsExcel2"
Option Explicit

' ============================================================================
' modTestsExcel2 - テスト3層の層(b)の分割先(30,000字契約・12章§2)
' ----------------------------------------------------------------------------
' 入口は modTestsExcel.RunAllExcelTests から呼ばれる RunExcelTests2 の1本
' (14章§6。wintest からの唯一の入口は従来どおり RunAllExcelTests)。戻り値は
' **本モジュールが打った Check の本数**で、呼出側の自己照合(TE_EXPECTED)へ足す。
'
' 本モジュールが受け持つ検査(裁定書11):
'   Q9  = 貼付欄の描画->保存の往復で原文が1字も変わらない(13章§2.11・16章E-22)。
'         連結側 modUICase3.JoinField は Private かつシートI/Oのため層(a)から
'         原理的に到達できず、区切り文字の再混入は純層のどの変異でも捕まらない。
'   Q1  = 続き欄に収まらない描画(overflow)では ci_case_id を空のままにし、
'         SavePasteFields の一致検査が保存を止める(13章§2.11)。**ガードを
'         無効化する変異はこの1本でのみ検出できる**(揮発フラグではなく
'         ブックに残るセルが根拠であることの回帰)。
'
' フィクスチャ規律は modTestsExcel と同じ: 自前で作り、成功・失敗のどちらの
' 経路でも後始末する。Notice(MsgBox)を出す経路は無人実行を止めるため呼ばない。
' LibreOffice では走らない(層(b)は実Excel専用)。
' ============================================================================

' フィクスチャ案件(書式は IsValidCaseId 合格・実在しない未来日付系)。
Private Const T2_CASE As String = "C-97990102-903"
Private Const T2_COMPANY As String = "T47検査用商事"
Private Const T2_SHEET As String = "案件入力"

' 貼付欄の1セル上限(modUICase3.U3_CHUNK と同値。ここを超えると続き欄へ回る)。
Private Const T2_CHUNK As Long = 32000

' 13章§2.12 B1 と同一の不一致文言(裁定書11 Q1 で案件入力側も同じ作法にした)。
Private Const T2_MSG_MISMATCH As String = "画面の案件と保存先が一致しません。再描画してください"

Private m2Run As Long

' ============================================================================
' RunExcelTests2 - 層(b)の分割先の入口。戻り値=打った Check の本数。
' ============================================================================
Public Function RunExcelTests2() As Long
    m2Run = 0
    TestQ9PasteRoundTrip
    TestV5DraftRowNotCounted
    RunExcelTests2 = m2Run
End Function

Private Sub ECheck(ByVal testName As String, ByVal cond As Boolean, _
                   Optional ByVal detail As String)
    m2Run = m2Run + 1
    modTestRunner.Check testName, cond, detail
End Sub

' ============================================================================
' Q9 + Q1: 貼付欄の往復一致と overflow の永続ガード
'   (1) 33,000字(=本欄32,000+続き欄1,000)を case_data へ置いて描画->保存し、
'       読み戻した原文が1字も変わらないこと(境界に vbLf が増えない)。
'   (2) 続き欄を持たない欄(input_memo)へ33,000字を置いて描画すると overflow に
'       なり、ci_case_id が**空のまま**になること。
'   (3) その状態の[保存]は1欄も書かず、hm_warning へ不一致文言・err_log へ
'       E0302 を残し、case_data の原文が消えないこと。
' ============================================================================
Private Sub TestQ9PasteRoundTrip()
    Dim wsCases As Object
    Dim rowNo As Long
    Dim warnOrig As String
    Dim hpText As String
    Dim memoText As String
    Dim tail As Long
    Dim actOrig As String
    On Error GoTo Crashed

    ' 裁定書13 W5: CaseSave は成功・ブロックのどちらでも ShowSheet "HOME" を通る
    ' ため、本テストの途中で活性シートが移る。実行前のシート名を覚えておき、
    ' Cleanup で必ず戻す(以降の層(b)テストの前提を壊さない)。画面遷移は ui層の
    ' 責務なので、素の Activate ではなく modUISheet.ShowSheet を通す。
    actOrig = ActiveSheetName()

    Set wsCases = SheetByName("案件一覧")
    Dim cCase As Long
    Dim hdr As Variant
    If Not wsCases Is Nothing Then
        hdr = Hdr1(wsCases, 32)
        cCase = modUtil.FindHeaderCol(hdr, "case_id")
    End If
    ECheck "T47B-Q9-01_案件一覧と貼付欄の前提が揃っている", _
           (Not wsCases Is Nothing) And cCase > 0 And _
           Not (modUISheet.NamedCell("ci_paste_hp_1") Is Nothing)
    If wsCases Is Nothing Or cCase <= 0 Then Exit Sub

    warnOrig = modUISheet.ReadNamed("hm_warning")
    rowNo = LastRowA(wsCases) + 1
    PutCell wsCases, rowNo, cCase, T2_CASE
    PutNamedCol wsCases, hdr, rowNo, "company", T2_COMPANY
    PutNamedCol wsCases, hdr, rowNo, "industry_code", "T47"
    PutNamedCol wsCases, hdr, rowNo, "industry_name", "検査用"
    PutNamedCol wsCases, hdr, rowNo, "case_type", "new"
    PutNamedCol wsCases, hdr, rowNo, "status", "draft"

    ' (1) 往復一致。本欄を1字だけ超える長さにして境界をまたがせる。
    hpText = "先頭" & String$(T2_CHUNK - 4, "あ") & "境界" & String$(996, "い") & "末尾"
    modCaseStore.SaveData T2_CASE, "input_hp", hpText
    modCaseStore.SaveData T2_CASE, "input_memo", vbNullString

    modUICase3.DrawCaseInput T2_CASE
    ECheck "T47B-Q9-02_収まる描画では ci_case_id に案件IDが入る", _
           Trim$(modUISheet.ReadNamed("ci_case_id")) = T2_CASE, _
           "実際=" & modUISheet.ReadNamed("ci_case_id")

    Dim ok As Boolean
    ok = modUICase3.SaveCaseInput(T2_CASE)
    ECheck "T47B-Q9-03_描き切った画面は保存できる", ok

    Dim back As String
    back = modCaseStore.LoadData(T2_CASE, "input_hp")
    ECheck "T47B-Q9-04_描画->保存の往復で字数が変わらない(境界にvbLfが増えない)", _
           Len(back) = Len(hpText), _
           "元=" & CStr(Len(hpText)) & " 往復後=" & CStr(Len(back))
    ECheck "T47B-Q9-05_描画->保存の往復で原文が1字も変わらない", _
           StrComp(back, hpText, vbBinaryCompare) = 0

    ' (2) overflow(続き欄を持たない input_memo に本欄超えを置く)。
    memoText = "冒頭" & String$(T2_CHUNK - 2, "う") & "あふれ"
    modCaseStore.SaveData T2_CASE, "input_memo", memoText

    modUICase3.DrawCaseInput T2_CASE
    ECheck "T47B-Q1-06_overflow描画では ci_case_id を空のままにする", _
           LenB(Trim$(modUISheet.ReadNamed("ci_case_id"))) = 0, _
           "実際=" & modUISheet.ReadNamed("ci_case_id")

    ' (3) その画面からの保存は1欄も書かない(ガード無効化の変異はここで落ちる)。
    tail = ErrLogTail()
    ok = modUICase3.SaveCaseInput(T2_CASE)
    ECheck "T47B-Q1-07_表示case_idが空の画面からは保存しない(False)", Not ok
    ECheck "T47B-Q1-08_hm_warningへ不一致文言を書く", _
           modUISheet.ReadNamed("hm_warning") = T2_MSG_MISMATCH, _
           "実際=" & modUISheet.ReadNamed("hm_warning")
    ECheck "T47B-Q1-09_err_logへE0302 case_id_mismatch:ci", _
           ErrLogged(tail, "E0302", "case_id_mismatch:ci")
    ECheck "T47B-Q1-10_ブロックされた保存で原文が消えない", _
           StrComp(modCaseStore.LoadData(T2_CASE, "input_hp"), hpText, _
                   vbBinaryCompare) = 0

    ' (4) 裁定書12 V1(13章§2.11): CaseSave の3値判定。ci_case_id が空のまま
    '     [保存して戻る]を押しても**新規採番へ倒さない**。3値判定を「空->採番」
    '     へ戻す変異はここで落ちる(案件一覧に行が1本増えるため)。
    '     裁定書13 W5: CaseSave は冒頭で TryEnterUiLock を取る。ロックが他所で
    '     握られたままだと本文へ入らず無言で戻るため、hm_warning が空のまま
    '     V1-12 だけが赤になる(誤検知)。**先にロックの空きを確かめ**、取れなければ
    '     2本とも SKIP 扱いにする(本数は不変=TE_EXPECTED の自己照合を崩さない)。
    Dim lockFree As Boolean
    lockFree = modUIProgress.TryEnterUiLock("T47B-V1の前提確認")
    If lockFree Then modUIProgress.ExitUiLock

    Dim rowsBefore As Long
    rowsBefore = LastRowA(wsCases)
    modUISheet.WriteNamed "hm_warning", vbNullString
    If lockFree Then
        modUICase3.CaseSave
        ECheck "T47B-V1-11_表示case_idが空の[保存して戻る]は採番しない", _
               LastRowA(wsCases) = rowsBefore, _
               "案件一覧の最終行 前=" & CStr(rowsBefore) & " 後=" & CStr(LastRowA(wsCases))
        ECheck "T47B-V1-12_ブロック時はhm_warningへ警告文を出す(V3)", _
               modUISheet.ReadNamed("hm_warning") = T2_MSG_MISMATCH, _
               "実際=" & modUISheet.ReadNamed("hm_warning")
    Else
        ECheck "T47B-V1-11_表示case_idが空の[保存して戻る]は採番しない", True, _
               "SKIP: UiLockが握られており CaseSave を無人実行できない"
        ECheck "T47B-V1-12_ブロック時はhm_warningへ警告文を出す(V3)", True, _
               "SKIP: 同上"
    End If

Cleanup:
    On Error Resume Next
    modCaseStore.SaveData T2_CASE, "input_hp", vbNullString
    modCaseStore.SaveData T2_CASE, "input_memo", vbNullString
    ClearNamed "ci_paste_hp_1"
    ClearNamed "ci_paste_hp_2"
    ClearNamed "ci_paste_hp_3"
    ClearNamed "ci_paste_memo_1"
    ' 裁定書12 V8: 属性欄も戻す。フィクスチャ企業名が画面に残ったままだと、
    ' 実機で[保存して戻る]を押したときの挙動が検査用の値に引きずられる。
    ClearNamed "ci_case_type"
    ClearNamed "ci_company"
    ClearNamed "ci_industry_code"
    ClearNamed "ci_industry_name"
    ClearNamed "ci_dossier_tier"
    ClearNamed "ci_channel"
    ClearNamed "ci_kanji"
    ClearNamed "ci_bid"
    ClearNamed "ci_reins"
    ClearNamed "ci_other_insurers"
    ClearNamed "ci_s4_variant"
    modUISheet.WriteNamed "ci_case_id", vbNullString
    modUISheet.WriteNamed "hm_warning", warnOrig
    DropFixtureRow wsCases, cCase
    ' 裁定書13 W5: 活性シートを実行前へ戻す(CaseSave が HOME へ移していることがある)。
    If LenB(actOrig) > 0 Then modUISheet.ShowSheet actOrig
    Exit Sub
Crashed:
    ECheck "T47B-Q9-99_想定外エラー", False, _
           "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Cleanup
End Sub

' ============================================================================
' V5(裁定書12・13章§2.6): 下書き行に status を書いても CountByStatus が
'   数えない。HOMEの hm_inbox_* は CountByStatus だけが値源であり、下書き行を
'   数えると「一括診断は0件なのにHOMEは1件」という食い違いが出る。下書き行の
'   スキップを無効化する変異はこの1本でのみ検出できる(層(a)から到達できない)。
' ============================================================================
Private Sub TestV5DraftRowNotCounted()
    Dim ws As Object
    Dim addedRow As Long
    Dim before As Long
    Dim after1 As Long
    Dim okAll As Boolean
    Dim detail As String
    On Error GoTo Crashed

    Set ws = SheetByName("受信箱")
    If ws Is Nothing Then
        detail = "受信箱シートが無い"
        GoTo Report
    End If

    Dim hdr As Variant
    Dim cId As Long
    Dim cStatus As Long
    hdr = Hdr1(ws, 32)
    cId = modUtil.FindHeaderCol(hdr, "inbox_id")
    cStatus = modUtil.FindHeaderCol(hdr, "status")
    If cId <= 0 Or cStatus <= 0 Then
        detail = "inbox_id / status 列が引けない"
        GoTo Report
    End If

    before = modUIInbox.CountByStatus("undiagnosed")

    ' 末尾へ「マーカー付き かつ status=undiagnosed」の行を1本だけ足す。
    addedRow = LastRowA(ws) + 1
    PutCell ws, addedRow, cId, modInboxStore.IB_DRAFT_MARK
    PutCell ws, addedRow, cStatus, "undiagnosed"

    after1 = modUIInbox.CountByStatus("undiagnosed")
    okAll = (after1 = before)
    detail = "追加前=" & CStr(before) & " 追加後=" & CStr(after1)

Report:
    ECheck "T47B-V5-13_下書き行にstatusを書いてもCountByStatusが数えない", _
           okAll, detail
    On Error Resume Next
    If addedRow >= 2 And Not ws Is Nothing Then ws.Rows(addedRow).Delete
    Exit Sub
Crashed:
    okAll = False
    detail = "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Report
End Sub

' ============================================================================
' 共通の下請け(すべて Private。公開口は RunExcelTests2 のみ=14章§6)
' ============================================================================

' 案件一覧からフィクスチャ案件の行を全削除する(下から回す)。
Private Sub DropFixtureRow(ByVal ws As Object, ByVal cCase As Long)
    On Error Resume Next
    If ws Is Nothing Or cCase <= 0 Then Exit Sub
    Dim r As Long
    For r = LastRowA(ws) To 2 Step -1
        If Trim$(CellStr(ws, r, cCase)) = T2_CASE Then ws.Rows(r).Delete
    Next r
End Sub

Private Sub ClearNamed(ByVal rangeName As String)
    On Error Resume Next
    If modUISheet.NamedCell(rangeName) Is Nothing Then Exit Sub
    modUISheet.WriteNamed rangeName, vbNullString
End Sub

Private Sub PutCell(ByVal ws As Object, ByVal r As Long, ByVal c As Long, _
                    ByVal valueText As String)
    If c <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(r, c), valueText, "T47B/案件一覧"
End Sub

Private Sub PutNamedCol(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                        ByVal colName As String, ByVal valueText As String)
    PutCell ws, r, modUtil.FindHeaderCol(hdr, colName), valueText
End Sub

' 実行前の活性シート名(取れなければ "")。裁定書13 W5 の復帰用。
Private Function ActiveSheetName() As String
    On Error GoTo NoActive
    ActiveSheetName = CStr(ActiveSheet.Name)
    Exit Function
NoActive:
    ActiveSheetName = vbNullString
End Function

' シートを名前で引く(非表示・veryHidden も対象。無ければ Nothing)。
Private Function SheetByName(ByVal sheetTitle As String) As Object
    On Error GoTo NoSheet
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.count
        If ThisWorkbook.Worksheets(i).Name = sheetTitle Then
            Set SheetByName = ThisWorkbook.Worksheets(i)
            Exit Function
        End If
    Next i
    Exit Function
NoSheet:
    Set SheetByName = Nothing
End Function

' A列基準の最終行(データ無しは1)。
Private Function LastRowA(ByVal ws As Object) As Long
    On Error GoTo One1
    LastRowA = ws.Cells(ws.Rows.count, 1).End(-4162).row
    If LastRowA < 1 Then LastRowA = 1
    Exit Function
One1:
    LastRowA = 1
End Function

' 見出し行(1行目)を 1..scanCols で読む。
Private Function Hdr1(ByVal ws As Object, ByVal scanCols As Long) As Variant
    On Error GoTo Empty0
    Hdr1 = ws.Range(ws.Cells(1, 1), ws.Cells(1, scanCols)).Value
    Exit Function
Empty0:
    Hdr1 = Empty
End Function

' 1セルを文字列で読む(読めなければ "")。
Private Function CellStr(ByVal ws As Object, ByVal r As Long, ByVal c As Long) As String
    On Error GoTo Blank0
    CellStr = CStr(ws.Cells(r, c).Value)
    Exit Function
Blank0:
    CellStr = vbNullString
End Function

' err_log の現在の最終行(この行より後を「テストが積んだ分」として検査する)。
Private Function ErrLogTail() As Long
    Dim ws As Object
    Set ws = SheetByName("err_log")
    If ws Is Nothing Then Exit Function
    ErrLogTail = LastRowA(ws)
End Function

' err_log の fromRow より後に「err_code一致 かつ detail部分一致」の行があるか。
Private Function ErrLogged(ByVal fromRow As Long, ByVal codeText As String, _
                           ByVal detailPart As String) As Boolean
    On Error GoTo No0
    Dim ws As Object
    Set ws = SheetByName("err_log")
    If ws Is Nothing Then Exit Function
    Dim hdr As Variant
    hdr = Hdr1(ws, 8)
    Dim cCode As Long
    Dim cDetail As Long
    cCode = modUtil.FindHeaderCol(hdr, "err_code")
    cDetail = modUtil.FindHeaderCol(hdr, "detail")
    If cCode <= 0 Or cDetail <= 0 Then Exit Function
    Dim r As Long
    Dim lastRow As Long
    lastRow = LastRowA(ws)
    For r = fromRow + 1 To lastRow
        If CStr(ws.Cells(r, cCode).Value) = codeText Then
            If InStr(1, CStr(ws.Cells(r, cDetail).Value), detailPart, vbBinaryCompare) > 0 Then
                ErrLogged = True
                Exit Function
            End If
        End If
    Next r
    Exit Function
No0:
    ErrLogged = False
End Function
