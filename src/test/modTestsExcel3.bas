Attribute VB_Name = "modTestsExcel3"
Option Explicit

' ============================================================================
' modTestsExcel3 - テスト3層の層(b)の3本目(30,000字契約・12章§2)
' ----------------------------------------------------------------------------
' 入口は modTestsExcel.RunAllExcelTests から呼ばれる RunExcelTests3 の1本
' (14章§6。wintest からの唯一の入口は従来どおり RunAllExcelTests)。戻り値は
' **本モジュールが打った Check の本数**で、呼出側の自己照合(TE_EXPECTED)へ足す。
'
' なぜ3本目か: modTestsExcel2 に本群を足すと 33,190 字となり 30,000字契約を
'   超えたため(vba_lint が ERROR で止める)。分け方は「W10=企業ファイルの
'   往復」というウェーブ単位で、フィクスチャも自前で持つ。
'
' 本モジュールが受け持つ検査(裁定書28 W10・13章§2.8 v2.7):
'   W10 = 企業ファイル(1社1.xlsx)の往復一致と、その照合が効いていることの
'         変異注入(列を1つ落とすと落ちる)。
'
' フィクスチャ規律は modTestsExcel と同じ: 自前で作り、成功・失敗のどちらの
' 経路でも後始末する。Notice(MsgBox)を出す経路は呼ばない(無人実行を止めない)。
' LibreOffice では走らない(層(b)は実Excel専用)。
' ============================================================================

' フィクスチャ案件(書式は IsValidCaseId 合格・実在しない未来日付系)。
Private Const T3_CASE As String = "C-97990102-904"
Private Const T3_COMPANY As String = "T47検査用W10商事"
Private Const T3_SAFE_TEXT As String = "当社は1952年の創業以来、静岡県浜松市を拠点に事業を営んでいます。"

' 企業ファイル往復で落とす列。13章§2.1 の案件一覧の列であり、企業ファイル側
'   dossier_case の見出しにも同名で並ぶ(第2ラウンド情報)。
Private Const T3_DROP_COL As String = "focus_line_ids"

Private m3Run As Long

' ============================================================================
' RunExcelTests3 - 層(b)の3本目の入口。戻り値=打った Check の本数。
' ============================================================================
Public Function RunExcelTests3() As Long
    m3Run = 0
    TestW10CompanyFileRoundTrip
    ' 裁定書34 §1.1(W12-A): HTML画面の層(b)6本。modTestsExcel は 29,902字で
    '   満杯のため、結線はここから行う(modTestsExcel には TE_EXPECTED の
    '   数字1つだけを直す)。戻り値は自分が打った本数へ足して返す。
    m3Run = m3Run + modTestsExcelNavi.RunExcelTestsNavi()
    RunExcelTests3 = m3Run
End Function

Private Sub ECheck(ByVal testName As String, ByVal cond As Boolean, _
                   Optional ByVal detail As String)
    m3Run = m3Run + 1
    modTestRunner.Check testName, cond, detail
End Sub

' ============================================================================
' IsMacExcel - Mac版Excelで動いているか(裁定書29 裁定4)。
' ----------------------------------------------------------------------------
'   Application.OperatingSystem は Mac版が "Macintosh ..." を返す。判定を
'   **この1関数に閉じる**ことで、ホストの見分け方が散らばるのを防ぐ。
'   読めない環境(LibreOffice等)では False =「Macではない」側へ倒す
'   (Windowsの検問に新しい逃げ道を作らないため)。
'
'   なぜ core(modUtilPath)ではなく**テスト層**に置くのか: 用途は層(b)の
'   SKIP判定だけで、製品コードは1箇所も呼ばない。core へ置くと 12章§4 の
'   R4(Excelトークンの許可モジュール)を1本広げることになるため、
'   司令塔の裁定(2026-09-05)で test層へ移した。test層は R4 適用外である。
'
'   本モジュールの他の Public は RunExcelTests3 だけだが、層(b)の3本
'   (modTestsExcel / modTestsExcel2 / modTestsExcel3)から共用するため
'   Public にする(層(b)どうしの参照はテスト層内なので R1 に触れない)。
' ============================================================================
Public Function IsMacExcel() As Boolean
    On Error GoTo NotMac
    IsMacExcel = (InStr(1, Application.OperatingSystem, "Macintosh", vbTextCompare) = 1)
    Exit Function
NotMac:
    IsMacExcel = False
End Function

' ============================================================================
' W10: 企業ファイルの往復一致(裁定書28・13章§2.8 v2.7・17章§4-1 層(b))
' ----------------------------------------------------------------------------
'   (1) 案件を企業ファイルへ書き出し、**本体側と企業ファイル側の全項目**
'       (案件一覧の全列 + case_data の全 data_key)が一致すること。
'   (2) **変異注入**: 書き出した企業ファイルの dossier_case から列を1つ
'       (focus_line_ids)落とすと、(1)の照合が**落ちること**。落ちなければ
'       「照合が列を見ていない」ことになり、往復テストが空回りしている。
'
'   層(a)に置けない理由: 一致の中身は .xlsx を実際に開いて書いて読み直した
'   結果でしか作れない(LibreOffice では Workbooks.Add / SaveAs の経路を
'   確かめられない)。
' ============================================================================
Private Sub TestW10CompanyFileRoundTrip()
    Dim wsCases As Object
    Dim hdr As Variant
    Dim cCase As Long
    Dim rowNo As Long
    Dim dirText As String
    Dim pathText As String
    Dim okSame As Boolean, okDrop As Boolean
    Dim detSame As String, detDrop As String
    On Error GoTo Crashed

    detSame = "前提不成立"
    detDrop = "前提不成立"

    Set wsCases = SheetByName("案件一覧")
    If wsCases Is Nothing Then GoTo Report
    hdr = Hdr1(wsCases, 32)
    cCase = modUtil.FindHeaderCol(hdr, "case_id")
    If cCase <= 0 Then GoTo Report

    rowNo = LastRowA(wsCases) + 1
    PutCell wsCases, rowNo, cCase, T3_CASE
    PutNamedCol wsCases, hdr, rowNo, "company", T3_COMPANY
    PutNamedCol wsCases, hdr, rowNo, "industry_code", "T47"
    PutNamedCol wsCases, hdr, rowNo, "industry_name", "検査用"
    PutNamedCol wsCases, hdr, rowNo, "case_type", "new"
    PutNamedCol wsCases, hdr, rowNo, "status", "s3_done"
    PutNamedCol wsCases, hdr, rowNo, "round_no", "2"
    ' 第2ラウンド情報(裁定書28 が企業ファイルへ含めることを求めた2列)。
    PutNamedCol wsCases, hdr, rowNo, "adopted_story_nos", "1;3"
    PutNamedCol wsCases, hdr, rowNo, T3_DROP_COL, "L-03;L-07"

    ' case_data 側は「S1〜S4 の編集後JSON」を含む枠を埋める(全 data_key を
    ' 埋める必要はない。空の枠は本体側も企業ファイル側も空で一致する)。
    modCaseStore.SaveData T3_CASE, "s1_edited", "{""mvv"":""検査用""}"
    modCaseStore.SaveData T3_CASE, "input_memo", T3_SAFE_TEXT

    dirText = modUtilPath.TempDir()
    If LenB(dirText) = 0 Then GoTo Report

    pathText = modCompanyFile.ExportCompanyFile(T3_CASE, dirText, vbNullString)
    If LenB(pathText) = 0 Then
        detSame = "ExportCompanyFile が空を返した(書き出しに失敗)"
        GoTo Report
    End If

    okSame = (modCompanyFile3.FileFingerprint(pathText, T3_CASE) = _
              modCompanyFile3.CaseFingerprint(T3_CASE))
    detSame = "本体側と企業ファイル側の全項目(案件一覧の全列+case_data の全 data_key)"

    ' 変異注入: 企業ファイルの dossier_case から列を1つ落とす。
    okDrop = DropDossierColumn(pathText, T3_DROP_COL)
    If okDrop Then
        okDrop = (modCompanyFile3.FileFingerprint(pathText, T3_CASE) <> _
                  modCompanyFile3.CaseFingerprint(T3_CASE))
        detDrop = "列 " & T3_DROP_COL & " を落としたら往復照合が落ちること"
    Else
        detDrop = "変異注入そのものに失敗(企業ファイルを書き換えられなかった)"
    End If

Report:
    ECheck "T47B-W10-01_企業ファイルの往復で全項目が一致する", okSame, detSame
    ECheck "T47B-W10-02_企業ファイルの列を1つ落とすと往復照合が落ちる", okDrop, detDrop

    On Error Resume Next
    If LenB(pathText) > 0 Then Kill pathText
    modCaseStore.SaveData T3_CASE, "s1_edited", vbNullString
    modCaseStore.SaveData T3_CASE, "input_memo", vbNullString
    DropFixtureRow wsCases, cCase
    Exit Sub
Crashed:
    okSame = False
    okDrop = False
    detSame = "Err=" & CStr(Err.Number) & " " & Err.Description
    detDrop = detSame
    Resume Report
End Sub

' 企業ファイルの dossier_case から見出しを1つ消す(変異注入の実体)。True=消せた。
Private Function DropDossierColumn(ByVal pathText As String, ByVal colName As String) As Boolean
    Dim wb As Object
    On Error GoTo Failed0

    Set wb = modCompanyFile2.DossierOpen(pathText, False)
    If wb Is Nothing Then Exit Function

    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, "dossier_case")
    If ws Is Nothing Then GoTo Failed0

    Dim blk As Variant
    blk = modCompanyFile2.SheetBlock(ws, modCompanyFile2.SheetLastRow(ws))
    Dim c As Long
    c = modUtil.FindHeaderCol(blk, colName)
    If c <= 0 Then GoTo Failed0

    ws.Columns(c).Delete
    DropDossierColumn = modCompanyFile2.DossierSaveAndClose(wb, pathText)
    Exit Function

Failed0:
    CloseQuietly wb
    DropDossierColumn = False
End Function

' ハンドラの外で閉じる(ハンドラ稼働中は再捕捉できない)。
Private Sub CloseQuietly(ByVal wb As Object)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Sub

' ----------------------------------------------------------------------------
' フィクスチャの小道具(modTestsExcel2 と同じ作法。本モジュールに閉じた写し)
' ----------------------------------------------------------------------------

Private Function SheetByName(ByVal sheetTitle As String) As Object
    On Error GoTo NoSheet
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
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
    LastRowA = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
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

Private Function CellStr(ByVal ws As Object, ByVal r As Long, ByVal c As Long) As String
    On Error GoTo Blank0
    CellStr = CStr(ws.Cells(r, c).Value)
    Exit Function
Blank0:
    CellStr = vbNullString
End Function

Private Sub PutCell(ByVal ws As Object, ByVal r As Long, ByVal c As Long, _
                    ByVal valueText As String)
    If c <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(r, c), valueText, "T47B/案件一覧"
End Sub

Private Sub PutNamedCol(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                        ByVal colName As String, ByVal valueText As String)
    PutCell ws, r, modUtil.FindHeaderCol(hdr, colName), valueText
End Sub

' フィクスチャ行の後始末(下から回すのは行ずれ回避)。
Private Sub DropFixtureRow(ByVal ws As Object, ByVal cCase As Long)
    On Error Resume Next
    If ws Is Nothing Or cCase <= 0 Then Exit Sub
    Dim r As Long
    For r = LastRowA(ws) To 2 Step -1
        If Trim$(CellStr(ws, r, cCase)) = T3_CASE Then ws.Rows(r).Delete
    Next r
End Sub
