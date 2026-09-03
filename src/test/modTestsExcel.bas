Attribute VB_Name = "modTestsExcel"
Option Explicit

' ============================================================================
' modTestsExcel - テスト3層の層(b): Excel固有テスト(17章 T-47・裁定書10 §3)
' ----------------------------------------------------------------------------
' 実Excel(Windows実機)でのみ走る。wintest/run_excel_tests.ps1 が
' modTestRunner.RunAllPureTests の後に RunAllExcelTests(14章§6 test層契約。
' 公開はこの1本のみ)を呼ぶ。集計は modTestRunner.Check へ相乗りし、
' PASS/FAIL/SKIP の作法を層(a)と1本化する。
'
' 本数の扱い(17章§4-1との整合):
'   wintest/tests_expected.txt は**純層(modTestsPure)の本数のまま**とする
'   (裁定書10 §3)。純層の本数照合は RunAllPureTests の末尾で完結しており、
'   本モジュールは自前の TE_EXPECTED と実際に打った本数を末尾で自己照合する
'   (0件実行の「全緑」や途中クラッシュの取りこぼしを成立させない)。
'
' フィクスチャ規律:
'   各テストは自前でフィクスチャを作り、成功・失敗のどちらの経路でも後始末
'   する(実データを壊さない)。err_log / usage_log への追記だけは「実行の記録」
'   として残す(ログは追記専用であり削除しない)。
'
' 期待値の根拠(実装に合わせて甘くしない):
'   B1  = 裁定書9 B1(SaveEditedStep の sN_case_id 不一致ガード。13章§2.12)
'   B3  = 裁定書9 B3(起票失敗で下書き行が残る。13章§4)
'   B11 = 裁定書9 B11(ReadKbSheets 0件でスナップショット不触。16章E-08)
'   B12 = 裁定書9 B12 + 裁定書10 M5(AnswerMemoCount は案件を問わず数える)
'   B13 = 裁定書9 B13 + 裁定書10 M6(2相書込の途中失敗痕跡は E0604 で止める)
'   C1  = 裁定書10 C1 + 補遺P7(投函下書き行の常設と、投函の失敗経路での残存)
'   Q9/Q1 = 裁定書11(貼付欄の往復一致と overflow の永続ガード)。実体は
'           modTestsExcel2(30,000字契約による分割先)にあり、本数だけ合流する
'
' LibreOffice(run_lo_tests.py)では走らない(層(b)は実Excel専用。tools/
' run_lo_tests.py のモジュール一覧にも含めない)。
' ============================================================================

' フィクスチャ用の案件ID(書式は IsValidCaseId 合格・実在しない未来日付系)。
Private Const TE_CASE_A As String = "C-97990101-901"
Private Const TE_CASE_B As String = "C-97990101-902"

' 裁定書9 B1 の不一致文言(13章§2.12。実装から写さず裁定書の文言を正とする)。
Private Const TE_MSG_B1 As String = "画面の案件と保存先が一致しません。再描画してください"

' 判断台帳の下書きマーカーとスナップショットの生存マーカー。
Private Const TE_DRAFT_MARK As String = "T47下書きマーカー"
Private Const TE_SNAP_MARK As String = "T47_SNAP_MARK"

' 仮seq帯の先頭値 = modCaseStore.CS_SEQ_BAND(100000) + 1。帯の値は実装裁量
' (裁定書9 B13)だが、M6検査はこの帯の実値に対する回帰なので定数で固定する。
Private Const TE_BAND_SEQ As Long = 100001

' 本モジュールが打つ Check の総本数(自己照合用。テストを増減したら更新)。
' (裁定書12: V5で1本・V1で2本を modTestsExcel2 へ追加し 44 -> 47)
' (裁定書22 M4: W6.1のナビ貼付3本を modTestsExcel2 へ追加し 47 -> 50)
' (裁定書22 仕上げ: v3.1の予約行方式を前提にしていた modTestsExcel2 の
'  旧 Q9/Q1/V1 群 12本を撤去し、新経路(保管+プレビュー)の Q9N 5本・V1N 2本へ
'  書き換えたため 50 -> 45。内訳 = 本モジュール34本 + modTestsExcel2 11本。
'  Q1(overflowの永続ガード)は v3.2 で**事象そのものが消えた**ため書き換え先を
'  持たない(跡地の理由は modTestsExcel2 の「跡地」節が持つ))
' (裁定書25 S3 / T-56: 7欄目 input_finance の往復1本を modTestsExcel2 へ
'  追加し 45 -> 46。内訳 = 本モジュール34本 + modTestsExcel2 12本)
' (裁定書26 A/D / W8.1: コーチ帯の図形の中の文字とフッター図形の存在の2本を
'  modTestsExcel2 へ追加し 46 -> 48。内訳 = 本モジュール34本 +
'  modTestsExcel2 14本。図形の描画は LibreOffice で確かめられないため、
'  この2本は層(b)にしか置けない)
Private Const TE_EXPECTED As Long = 48

Private mRun As Long    ' ECheck が数える実行本数

' ============================================================================
' RunAllExcelTests - 層(b)の唯一の入口(14章§6・17章 T-47)
'   ResetTests は呼ばない(純層の結果に積み増す。レポートは1本にまとまる)。
' ============================================================================
Public Sub RunAllExcelTests()
    mRun = 0
    TestB1SaveGuard
    TestB3JudgeDraft
    TestB11SnapshotGuard
    TestB12AnswerMemo
    TestB13TwoPhase
    TestC1DraftRow
    ' 裁定書11 Q9/Q1: 30,000字契約による分割先(modTestsExcel2)の本数を足す
    ' (wintest からの入口は本モジュールの1本のままにする=14章§6)。
    mRun = mRun + modTestsExcel2.RunExcelTests2()
    ' 自己照合はランナーへ直接打つ(mRun には数えない)。
    modTestRunner.Check "T47-00_層(b)本数の自己照合(" & CStr(TE_EXPECTED) & "本)", _
        mRun = TE_EXPECTED, "実際=" & CStr(mRun) & _
        " (不足は途中クラッシュ・前提不成立でテストが最後まで走っていない兆候)"
End Sub

' Check の相乗り口。層(b)の実行本数を自前で数える。
Private Sub ECheck(ByVal testName As String, ByVal cond As Boolean, _
                   Optional ByVal detail As String)
    mRun = mRun + 1
    modTestRunner.Check testName, cond, detail
End Sub

' ============================================================================
' B1: SaveEditedStep の case_id 不一致ガード(裁定書9 B1)
'   s1_case_id 表示セルへ別案件IDを置き、引数と不一致のとき
'   (1)保存しない (2)errText は空のまま (3)hm_warning へ規定文言
'   (4)err_log へ E0302 case_id_mismatch:s1、を検査する。
' ============================================================================
Private Sub TestB1SaveGuard()
    Dim warnOrig As String, idOrig As String
    Dim restoreNeeded As Boolean
    On Error GoTo Crashed

    Dim cell As Object
    Set cell = modUISheet.NamedCell("s1_case_id")
    ECheck "T47-B1-01_s1_case_id表示セルが存在する(13章§2.12)", Not (cell Is Nothing)
    If cell Is Nothing Then Exit Sub

    warnOrig = modUISheet.ReadNamed("hm_warning")
    idOrig = modUISheet.ReadNamed("s1_case_id")
    restoreNeeded = True
    modUISheet.WriteNamed "s1_case_id", TE_CASE_B

    Dim tail As Long
    tail = ErrLogTail()
    Dim errText As String
    Dim ok As Boolean
    ok = modUICase2.SaveEditedStep(TE_CASE_A, 1, errText)

    ECheck "T47-B1-02_case_id不一致では保存しない(False)", Not ok
    ECheck "T47-B1-03_不一致時のerrTextは空(文言は画面へ直書き)", LenB(errText) = 0, _
           "errText=" & errText
    ECheck "T47-B1-04_hm_warningへ不一致文言を書く", _
           modUISheet.ReadNamed("hm_warning") = TE_MSG_B1, _
           "実際=" & modUISheet.ReadNamed("hm_warning")
    ECheck "T47-B1-05_err_logへE0302 case_id_mismatch:s1", _
           ErrLogged(tail, "E0302", "case_id_mismatch:s1")

Cleanup:
    On Error Resume Next
    If restoreNeeded Then
        modUISheet.WriteNamed "s1_case_id", idOrig
        modUISheet.WriteNamed "hm_warning", warnOrig
        ' ガードが壊れていた場合に備え、誤って作られた s1_edited を掃除する。
        modCaseStore.SaveData TE_CASE_A, "s1_edited", vbNullString
    End If
    Exit Sub
Crashed:
    ECheck "T47-B1-99_想定外エラー", False, _
           "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Cleanup
End Sub

' ============================================================================
' B3: 起票失敗で下書き行が残る(裁定書9 B3)
'   判断台帳の末尾へ下書き(situationのみ)を置き、必須欠落の NewJudgement が
'   (1)空IDで失敗 (2)下書きセルを触らない (3)judge_id を書かない
'   (4)行を増やさない (5)E0101 を記録する、を検査する。続けて必須充足の
'   起票が judge_id 付きの行を1行だけ追加することも確かめる。
'   ※ 台帳は削除禁止(13章§4)だが、ここで消すのはテストが作った
'      フィクスチャ行のみ(実データには触れない)。
' ============================================================================
Private Sub TestB3JudgeDraft()
    Dim ws As Object
    Dim d As Long, last0 As Long, addedRow As Long
    Dim cId As Long, cSit As Long
    On Error GoTo Crashed

    Set ws = SheetByName("判断台帳")
    Dim hdr As Variant
    If Not ws Is Nothing Then
        hdr = Hdr1(ws, 16)
        cId = modUtil.FindHeaderCol(hdr, "judge_id")
        cSit = modUtil.FindHeaderCol(hdr, "situation")
    End If
    ECheck "T47-B3-01_判断台帳の見出し(judge_id/situation)が引ける", _
           (Not ws Is Nothing) And cId > 0 And cSit > 0
    If ws Is Nothing Or cId <= 0 Or cSit <= 0 Then Exit Sub

    last0 = LastRowA(ws)            ' judge_id 列 = A列(13章§2.7)
    d = last0 + 1
    modUtilText.SetCellSafe ws.Cells(d, cSit), TE_DRAFT_MARK, "T47/判断台帳draft"

    Dim tail As Long
    tail = ErrLogTail()
    Dim rec As TJudgement
    rec.line_id = "T47-LINE"
    rec.case_ref = TE_CASE_A
    rec.situation = TE_DRAFT_MARK
    rec.decision = vbNullString     ' 必須欠落 -> 起票失敗(裁定書9 B3 の失敗系)
    rec.key_reason = "T47決め手"
    rec.recorded_by = "T47"

    Dim retId As String
    retId = modJudgeStore.NewJudgement(rec)
    ECheck "T47-B3-02_必須欠落の起票は空IDで失敗", LenB(retId) = 0, "ret=" & retId
    ECheck "T47-B3-03_失敗しても下書き行の入力が残る", _
           CellStr(ws, d, cSit) = TE_DRAFT_MARK, "実際=" & CellStr(ws, d, cSit)
    ECheck "T47-B3-04_失敗時にjudge_idを書き込まない", _
           LenB(Trim$(CellStr(ws, d, cId))) = 0
    ECheck "T47-B3-05_失敗時に台帳の行が増えない", LastRowA(ws) = last0, _
           "before=" & CStr(last0) & " after=" & CStr(LastRowA(ws))
    ECheck "T47-B3-06_err_logへE0101 empty_required", _
           ErrLogged(tail, "E0101", "empty_required")

    ' 成功系: 下書きマーカーを外してから起票する(NewJudgement は台帳末尾=
    ' この位置へ確定行を書くため、マーカーを残すと確定行と混ざる)。
    ws.Cells(d, cSit).ClearContents

    rec.decision = "keep"
    retId = modJudgeStore.NewJudgement(rec)
    ECheck "T47-B3-07_必須充足の起票はjudge_idを返す", _
           modJudgeStore.IsValidJudgeId(retId), "ret=" & retId
    If modJudgeStore.IsValidJudgeId(retId) Then
        addedRow = LastRowA(ws)
        ECheck "T47-B3-08_起票行が台帳末尾へ1行だけ追加される", _
               (addedRow = last0 + 1) And (Trim$(CellStr(ws, addedRow, cId)) = retId), _
               "lastRow=" & CStr(addedRow) & " id=" & CellStr(ws, addedRow, cId)
    Else
        ECheck "T47-B3-08_起票行が台帳末尾へ1行だけ追加される", False, "起票が失敗したため検査不能"
        addedRow = 0
    End If

Cleanup:
    On Error Resume Next
    If Not ws Is Nothing Then
        If addedRow >= 2 Then ws.Rows(addedRow).Delete
        If d >= 2 And cSit > 0 Then
            If CellStr(ws, d, cSit) = TE_DRAFT_MARK Then ws.Cells(d, cSit).ClearContents
        End If
    End If
    Exit Sub
Crashed:
    ECheck "T47-B3-99_想定外エラー", False, _
           "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Cleanup
End Sub

' ============================================================================
' B11: ReadKbSheets 0件でスナップショット不触(裁定書9 B11・16章E-08)
'   kb_snapshot へ生存マーカーを置き、データ行0件のナレッジブック
'   (見出しのみの「業種マスタ」1枚)へ kb_path を差し替えて LoadKnowledge を
'   呼ぶ。ガードが生きていれば SaveSnapshot(Cells.Clear を伴う)は走らず
'   マーカーが生き残り、E0401 kb_zero_rows が記録される。
' ============================================================================
Private Sub TestB11SnapshotGuard()
    Dim kbOrig As String, tmpPath As String
    Dim wsSnap As Object, wbT As Object
    Dim createdSnap As Boolean, swapped As Boolean, marked As Boolean
    Dim alertsOrig As Boolean
    On Error GoTo Crashed
    alertsOrig = Application.DisplayAlerts

    kbOrig = modConfig.GetStr("kb_path", vbNullString)

    Set wsSnap = SheetByName("kb_snapshot")
    If wsSnap Is Nothing Then
        Set wsSnap = modUISheet.EnsureHiddenSheet("kb_snapshot")
        createdSnap = True
    End If
    If wsSnap Is Nothing Then
        ECheck "T47-B11-01_kb_pathを0件フィクスチャへ差替できる", False, _
               "kb_snapshot シートを用意できない"
        Exit Sub
    End If
    ' マーカーは4列目: SaveSnapshot(1..3列)にも RestoreSnapshot(3列読み)にも
    ' 掛からない位置で、Cells.Clear が走ったときだけ消える。
    modUtilText.SetCellSafe wsSnap.Cells(1, 4), TE_SNAP_MARK, "T47/kb_snapshot"
    marked = True

    tmpPath = TempFilePath("t47_kb_zero.xlsx")
    On Error Resume Next
    Kill tmpPath
    On Error GoTo Crashed
    Application.DisplayAlerts = False
    Set wbT = Application.Workbooks.Add
    wbT.Worksheets(1).Name = "業種マスタ"
    modUtilText.SetCellSafe wbT.Worksheets(1).Cells(1, 1), "industry_code", "T47/kbfixture"
    wbT.SaveAs tmpPath, 51
    wbT.Close False
    Set wbT = Nothing
    Application.DisplayAlerts = alertsOrig

    modConfig.SetValue "kb_path", tmpPath
    swapped = True
    ECheck "T47-B11-01_kb_pathを0件フィクスチャへ差替できる", _
           modConfig.GetStr("kb_path", vbNullString) = tmpPath

    Dim tail As Long
    tail = ErrLogTail()
    modKnowledge.LoadKnowledge      ' 戻り値は退避の有無に依存するため検査しない

    ECheck "T47-B11-02_0件読込でスナップショットを触らない(マーカー生存)", _
           CellStr(wsSnap, 1, 4) = TE_SNAP_MARK, "実際=" & CellStr(wsSnap, 1, 4)
    ECheck "T47-B11-03_err_logへE0401 kb_zero_rows", _
           ErrLogged(tail, "E0401", "kb_zero_rows")

Cleanup:
    On Error Resume Next
    If Not wbT Is Nothing Then wbT.Close False
    If marked Then wsSnap.Cells(1, 4).ClearContents
    If createdSnap And Not wsSnap Is Nothing Then
        Application.DisplayAlerts = False
        wsSnap.Delete
    End If
    Application.DisplayAlerts = alertsOrig
    If swapped Then modConfig.SetValue "kb_path", kbOrig
    If LenB(tmpPath) > 0 Then Kill tmpPath
    modKnowledge.LoadKnowledge      ' 元のkb_pathでキャッシュを回復(ベストエフォート)
    Exit Sub
Crashed:
    ECheck "T47-B11-99_想定外エラー", False, _
           "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Cleanup
End Sub

' ============================================================================
' B12: AnswerMemoCount の値検査(裁定書9 B12・裁定書10 M5)
'   answer_memo 非空行(空白のみは数えない)を、hs_case_id と引数 caseId が
'   不一致でも 0 へ倒さず常に数えること(M5: fail-open の閉鎖)を検査する。
' ============================================================================
Private Sub TestB12AnswerMemo()
    Dim ws As Object
    Dim headerRow As Long, cA As Long
    Dim r1 As Long, r2 As Long, r3 As Long
    Dim o1 As String, o2 As String, o3 As String
    Dim hsOrig As String
    Dim touched As Boolean
    On Error GoTo Crashed

    headerRow = modUISheet.BlockRow("hearing_questions")
    Set ws = modUISheet.BlockSheet("hearing_questions")
    If headerRow > 0 And Not ws Is Nothing Then
        Dim hdr As Variant
        hdr = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 6)).Value
        cA = modUtil.FindHeaderCol(hdr, "answer_memo")
    End If
    ECheck "T47-B12-01_hearing_questionsのanswer_memo列が引ける", _
           headerRow > 0 And cA > 0
    If headerRow <= 0 Or cA <= 0 Or ws Is Nothing Then Exit Sub

    Dim baseN As Long
    baseN = modExportHearing.AnswerMemoCount(TE_CASE_A)

    r1 = headerRow + 1
    r2 = headerRow + 2
    r3 = headerRow + 3
    o1 = CellStr(ws, r1, cA)
    o2 = CellStr(ws, r2, cA)
    o3 = CellStr(ws, r3, cA)
    hsOrig = modUISheet.ReadNamed("hs_case_id")
    touched = True

    Dim origN As Long
    If LenB(Trim$(o1)) > 0 Then origN = origN + 1
    If LenB(Trim$(o2)) > 0 Then origN = origN + 1
    If LenB(Trim$(o3)) > 0 Then origN = origN + 1

    modUtilText.SetCellSafe ws.Cells(r1, cA), "T47回答1", "T47/answer_memo"
    modUtilText.SetCellSafe ws.Cells(r2, cA), "T47回答2", "T47/answer_memo"
    modUtilText.SetCellSafe ws.Cells(r3, cA), "   ", "T47/answer_memo"

    Dim want As Long
    want = baseN - origN + 2
    ECheck "T47-B12-02_非空行を数える(空白のみは数えない)", _
           modExportHearing.AnswerMemoCount(TE_CASE_A) = want, _
           "期待=" & CStr(want) & " 実際=" & CStr(modExportHearing.AnswerMemoCount(TE_CASE_A))
    modUISheet.WriteNamed "hs_case_id", TE_CASE_B
    ECheck "T47-B12-03_M5: hs_case_id不一致でも0へ倒さず数える", _
           modExportHearing.AnswerMemoCount(TE_CASE_A) = want, _
           "期待=" & CStr(want) & " 実際=" & CStr(modExportHearing.AnswerMemoCount(TE_CASE_A))

    modUtilText.SetCellSafe ws.Cells(r1, cA), o1, "T47/answer_memo"
    modUtilText.SetCellSafe ws.Cells(r2, cA), o2, "T47/answer_memo"
    modUtilText.SetCellSafe ws.Cells(r3, cA), o3, "T47/answer_memo"
    modUISheet.WriteNamed "hs_case_id", hsOrig
    touched = False
    ECheck "T47-B12-04_後始末後は元の件数へ戻る", _
           modExportHearing.AnswerMemoCount(TE_CASE_A) = baseN

Cleanup:
    On Error Resume Next
    If touched Then
        modUtilText.SetCellSafe ws.Cells(r1, cA), o1, "T47/answer_memo"
        modUtilText.SetCellSafe ws.Cells(r2, cA), o2, "T47/answer_memo"
        modUtilText.SetCellSafe ws.Cells(r3, cA), o3, "T47/answer_memo"
        modUISheet.WriteNamed "hs_case_id", hsOrig
    End If
    Exit Sub
Crashed:
    ECheck "T47-B12-99_想定外エラー", False, _
           "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Cleanup
End Sub

' ============================================================================
' B13: 2相書込と E0604(裁定書9 B13・裁定書10 M6)
'   (1)32,007字の分割保存・結合が可逆 (2)空保存で全断片削除
'   (3)欠番(seq2欠落)は詰めず E0604 seq_gap で ""(fail-closed)
'   (4)本seq帯が空で仮seq帯だけ残る途中失敗痕跡は E0604 seq_band_orphan で ""。
'   途中失敗の状態は case_data へフィクスチャ行を直接置いて再現する。
' ============================================================================
Private Sub TestB13TwoPhase()
    Dim ws As Object
    Dim cCase As Long, cKey As Long, cSeq As Long, cBody As Long
    Dim touched As Boolean
    On Error GoTo Crashed

    Set ws = SheetByName("case_data")
    Dim hdr As Variant
    If Not ws Is Nothing Then
        hdr = Hdr1(ws, 8)
        cCase = modUtil.FindHeaderCol(hdr, "case_id")
        cKey = modUtil.FindHeaderCol(hdr, "data_key")
        cSeq = modUtil.FindHeaderCol(hdr, "seq")
        cBody = modUtil.FindHeaderCol(hdr, "content")
    End If
    ECheck "T47-B13-01_case_dataの見出し4列が引ける", _
           (Not ws Is Nothing) And cCase > 0 And cKey > 0 And cSeq > 0 And cBody > 0
    If ws Is Nothing Or cCase <= 0 Or cKey <= 0 Or cSeq <= 0 Or cBody <= 0 Then Exit Sub

    touched = True

    ' --- 正常系: 複数断片の往復(1セル32,000字 + 尾片) ---
    Dim content As String
    content = String$(32000, "a") & "T47TAIL"
    ECheck "T47-B13-02_2相書込SaveDataが成功する", _
           modCaseStore.SaveData(TE_CASE_A, "s1_edited", content)
    Dim got As String
    got = modCaseStore.LoadData(TE_CASE_A, "s1_edited")
    ECheck "T47-B13-03_断片分割の往復が可逆(32,007字)", got = content, _
           "期待len=" & CStr(Len(content)) & " 実際len=" & CStr(Len(got))
    ECheck "T47-B13-04_空保存で全断片を削除できる", _
           modCaseStore.SaveData(TE_CASE_A, "s1_edited", vbNullString)
    ECheck "T47-B13-05_削除後のLoadDataは空", _
           LenB(modCaseStore.LoadData(TE_CASE_A, "s1_edited")) = 0

    ' --- 欠番(B13): seq=1,3 だけを置く -> 詰めずに fail-closed ---
    Dim r As Long, tail As Long
    r = LastRowA(ws)
    PutDataRow ws, r + 1, cCase, cKey, cSeq, cBody, 1, "T47A"
    PutDataRow ws, r + 2, cCase, cKey, cSeq, cBody, 3, "T47C"
    tail = ErrLogTail()
    ECheck "T47-B13-06_欠番(seq2欠落)は詰めず空を返す", _
           LenB(modCaseStore.LoadData(TE_CASE_A, "s1_edited")) = 0
    ECheck "T47-B13-07_欠番はE0604 seq_gapを記録する", _
           ErrLogged(tail, "E0604", "seq_gap:s1_edited")
    CleanCaseData ws, cCase

    ' --- 途中失敗痕跡(M6): 本seq帯が空で仮seq帯だけ残る -> 静かに""へ倒さない ---
    r = LastRowA(ws)
    PutDataRow ws, r + 1, cCase, cKey, cSeq, cBody, TE_BAND_SEQ, "T47X"
    tail = ErrLogTail()
    ECheck "T47-B13-08_M6: 本帯空+仮帯残存は空を返す", _
           LenB(modCaseStore.LoadData(TE_CASE_A, "s1_edited")) = 0
    ECheck "T47-B13-09_M6: E0604 seq_band_orphanを記録する", _
           ErrLogged(tail, "E0604", "seq_band_orphan:s1_edited")
    CleanCaseData ws, cCase
    touched = False

Cleanup:
    On Error Resume Next
    If touched And Not ws Is Nothing Then
        CleanCaseData ws, cCase
        modCaseStore.SaveData TE_CASE_A, "s1_edited", vbNullString
    End If
    Exit Sub
Crashed:
    ECheck "T47-B13-99_想定外エラー", False, _
           "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Cleanup
End Sub

' ============================================================================
' C1: 投函下書き行(裁定書10 C1・補遺P7・13章§2.6)
'   (1)起動シーケンス(EnsureInboxButtons)の後、受信箱の先頭データ行の id 列が
'      下書きマーカーであること(下書き行の常設)。
'   (2)投函の失敗経路(body が空)にあたる状態で起動シーケンスを再実行しても、
'      行が増えず・マーカーが残ること(下書き行は消さない・二重に作らない)。
'   (3)下書き行は受信箱のデータではないので未診断一覧に混ざらないこと。
'   ※ InboxPost 自体は失敗時に MsgBox(モーダル)を出し無人実行を止めるため
'      直接は呼ばない。検査は下書き行の不変条件と走査の継ぎ目で行う。
' ============================================================================
Private Sub TestC1DraftRow()
    Dim ws As Object
    Dim cId As Long, cBody As Long, cStat As Long
    Dim origBody As String, origStat As String
    Dim lastBefore As Long
    Dim touched As Boolean
    On Error GoTo Crashed

    Set ws = SheetByName("受信箱")
    Dim hdr As Variant
    If Not ws Is Nothing Then
        hdr = Hdr1(ws, 24)
        cId = modUtil.FindHeaderCol(hdr, "inbox_id")
        cBody = modUtil.FindHeaderCol(hdr, "body")
    End If
    ECheck "T47-C1-01_受信箱の見出し(inbox_id/body)が引ける", _
           (Not ws Is Nothing) And cId > 0 And cBody > 0
    If ws Is Nothing Or cId <= 0 Or cBody <= 0 Then Exit Sub

    modUIInbox.EnsureInboxButtons
    ECheck "T47-C1-02_起動後の先頭データ行id列が投函下書きマーカー", _
           Trim$(CellStr(ws, 2, cId)) = modInboxStore.IB_DRAFT_MARK, _
           "実際=" & CellStr(ws, 2, cId)

    origBody = CellStr(ws, 2, cBody)
    touched = True
    ws.Cells(2, cBody).ClearContents        ' 投函の失敗経路(body 空)の入力状態
    lastBefore = LastRowA(ws)

    modUIInbox.EnsureInboxButtons           ' 冪等であること(下書き行を作り直さない)
    ECheck "T47-C1-03_body空でも受信箱の行が増えない", LastRowA(ws) = lastBefore, _
           "before=" & CStr(lastBefore) & " after=" & CStr(LastRowA(ws))
    ECheck "T47-C1-04_body空でも下書き行のマーカーが残る", _
           Trim$(CellStr(ws, 2, cId)) = modInboxStore.IB_DRAFT_MARK, _
           "実際=" & CellStr(ws, 2, cId)

    ' 走査スキップの検査は status を undiagnosed にした状態で行う(status が空
    ' のままだと「未診断ではない」という別の理由で外れ、IsDraftId の回帰に
    ' ならない)。cStat が引けないときは status を触らずに検査する。
    modUtilText.SetCellSafe ws.Cells(2, cBody), "T47下書き本文", "T47/受信箱body"
    cStat = modUtil.FindHeaderCol(hdr, "status")
    If cStat > 0 Then
        origStat = CellStr(ws, 2, cStat)
        modUtilText.SetCellSafe ws.Cells(2, cStat), "undiagnosed", "T47/受信箱status"
    End If
    ECheck "T47-C1-05_下書き行は未診断一覧に混ざらない(走査スキップ)", _
           InStr(1, modInboxStore.UndiagnosedIds(), modInboxStore.IB_DRAFT_MARK, _
                 vbBinaryCompare) = 0, _
           "実際=" & modInboxStore.UndiagnosedIds()

Cleanup:
    On Error Resume Next
    If touched Then
        If LenB(origBody) = 0 Then
            ws.Cells(2, cBody).ClearContents
        Else
            modUtilText.SetCellSafe ws.Cells(2, cBody), origBody, "T47/受信箱body"
        End If
        If cStat > 0 Then
            If LenB(origStat) = 0 Then
                ws.Cells(2, cStat).ClearContents
            Else
                modUtilText.SetCellSafe ws.Cells(2, cStat), origStat, "T47/受信箱status"
            End If
        End If
    End If
    Exit Sub
Crashed:
    ECheck "T47-C1-99_想定外エラー", False, _
           "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Cleanup
End Sub

' case_data へフィクスチャ断片行を1行置く(途中失敗状態の再現用)。
Private Sub PutDataRow(ByVal ws As Object, ByVal r As Long, ByVal cCase As Long, _
                       ByVal cKey As Long, ByVal cSeq As Long, ByVal cBody As Long, _
                       ByVal seqNo As Long, ByVal bodyText As String)
    modUtilText.SetCellSafe ws.Cells(r, cCase), TE_CASE_A, "T47/case_data"
    modUtilText.SetCellSafe ws.Cells(r, cKey), "s1_edited", "T47/case_data"
    ws.Cells(r, cSeq).Value = seqNo    ' SAFE:const 内部生成のLong
    modUtilText.SetCellSafe ws.Cells(r, cBody), bodyText, "T47/case_data"
End Sub

' case_data からフィクスチャ案件(TE_CASE_A)の行を全削除する(下から回す)。
Private Sub CleanCaseData(ByVal ws As Object, ByVal cCase As Long)
    On Error Resume Next
    Dim r As Long
    For r = LastRowA(ws) To 2 Step -1
        If Trim$(CellStr(ws, r, cCase)) = TE_CASE_A Then ws.Rows(r).Delete
    Next r
End Sub

' ============================================================================
' 共通の下請け(すべて Private。公開口は RunAllExcelTests のみ=14章§6)
' ============================================================================

' シートを名前で引く(非表示・veryHidden も対象。無ければ Nothing)。
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
    Dim cCode As Long, cDetail As Long
    cCode = modUtil.FindHeaderCol(hdr, "err_code")
    cDetail = modUtil.FindHeaderCol(hdr, "detail")
    If cCode <= 0 Or cDetail <= 0 Then Exit Function
    Dim r As Long, lastRow As Long
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

' 一時ファイルの置き場(%TEMP%。取れなければブックと同じフォルダ)。
Private Function TempFilePath(ByVal fileName As String) As String
    Dim d As String
    d = Environ$("TEMP")
    If LenB(d) = 0 Then d = ThisWorkbook.Path
    If Right$(d, 1) = "\" Then d = Left$(d, Len(d) - 1)
    TempFilePath = d & "\" & fileName
End Function
