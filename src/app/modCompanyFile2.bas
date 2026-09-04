Attribute VB_Name = "modCompanyFile2"
Option Explicit

' ============================================================================
' modCompanyFile2 - 企業ドシエファイルのブック・シート下位I/O(app層・T-29)
' ----------------------------------------------------------------------------
' なぜ2本に分かれているのか:
'   T-29の実装を1モジュールに収めると約36,000字になり、開発憲法の
'   【1モジュール30,000字契約】(12章§2・CONTRIBUTING.md §2.5)を超えた。
'   そこで「何を書くか」(modCompanyFile: 13章§2.8の5シートの中身・PII走査・
'   2段検証の判定)と「どこへどう書くか」(本モジュール: .xlsxを開く/作る/
'   保存する、行を消す、seq順に結合する)の境目で割った。既存の
'   modValidate / modValidate2 と同じ分け方である。
'   ※ 12章§2のモジュール一覧に `modCompanyFile2` は無いため vba_lint が
'     WARN を出す。分割の是非は司令塔の裁定を仰ぐ(concerns 参照)。
'
' R4(12章§4): 企業ドシエファイル(1社1.xlsx)のシートI/Oが責務そのもの。触るのは
'   【引数で渡された企業ファイル側のシート】だけで、本体ブックには触らない。
'   セル書込は全て modUtilText.SetCellSafe(16章 NFR-S7(1))。
'
' 13章§2.8の物理レイアウト(シート名と1行目ヘッダ)は本モジュールが持つ。
'   modCompanyFile 側はシート名だけを知っていればよい。
' ============================================================================

Private Const C2_SRC As String = "modCompanyFile2"

' 13章§2.8の5シートと、その1行目ヘッダ(物理名 snake_case。列番号は書かない)。
' 裁定書28(W10)で dossier_case / dossier_data / dossier_judge の3枚を足した
' (企業ファイルが「蓄積の正」になったため、案件一覧の全列・case_data の全
'  data_key・判断台帳を1ファイルに収める。13章§2.8)。既存5枚は不変。
Private Const C2_SHEETS As String = "dossier_meta;dossier_profile;dossier_rounds;dossier_facts;dossier_notes;dossier_case;dossier_data;dossier_judge"
Private Const C2_HDR_META As String = "company_id;company;industry_code;schema_version;created_at;updated_at;pii_scan_result;pii_confirmed_at;owner"
Private Const C2_HDR_PROFILE As String = "case_id;round_no;seq;s1_json;mvv;aspirations;market_context;updated_at"
Private Const C2_HDR_ROUNDS As String = "round_no;case_id;executed_at;dossier_tier;seq;s1_json;s2_json;s3_json;report_file;round_summary"
Private Const C2_HDR_FACTS As String = "round_no;case_id;visited_at;event;used_proposals;customer_quote;terms_summary;loss_note;recorded_by;recorded_at;row_hash"
Private Const C2_HDR_NOTES As String = "round_no;case_id;note_kind;tag;seq;content;saved_at"
' 裁定書28(W10)の3枚。dossier_case は 13章§2.1 案件一覧の全列(物理順)＋
'   schema_version、dossier_data は 13章§2.2 case_data と同じ縦持ち、
'   dossier_judge は 13章§2.7 判断台帳の全列＋case_id/round_no。
Private Const C2_HDR_CASE As String = "case_id;case_type;dossier_tier;parent_case_id;company;industry_code;industry_name;channel;kanji;bid;reins;other_insurers;status;created_at;updated_at;owner;adopted_story_nos;focus_line_ids;ppt_path;report_path;note;s4_variant;round_no;last_ok_step;failed_step;schema_version"
Private Const C2_HDR_DATA As String = "case_id;data_key;seq;content;saved_at"
Private Const C2_HDR_JUDGE As String = "case_id;round_no;judge_id;judged_at;line_id;recorded_by;case_ref;situation;decision;factor_note;key_reason;result;post_loss"

' xlOpenXMLWorkbook(マクロ無し.xlsx) / xlUp。組込定数名を書くとLOの構文検査で
' 未定義名になるため数値で持つ。
Private Const C2_XLSX_FORMAT As Long = 51
Private Const C2_DIR_UP As Long = -4162
' 裁定書28: dossier_case が26列になったため 20 -> 30 へ広げた(見出しの右端を
'   読み落とすと FindHeaderCol が列を見つけられず、その列が黙って空になる)。
Private Const C2_SCAN_COLS As Long = 30
Private Const C2_SEP As String = ";"

' ============================================================================
' DossierOpenOrCreate - 既存なら開き、無ければ5シート構成の新規.xlsxを作る。
'   戻り値は開いた状態のブック(失敗は Nothing)。保存は DossierSaveAndClose。
' ============================================================================
Public Function DossierOpenOrCreate(ByVal pathText As String) As Object
    On Error GoTo Failed

    If FileExists(pathText) Then
        Set DossierOpenOrCreate = DossierOpen(pathText, False)
        EnsureAllSheets DossierOpenOrCreate
        Exit Function
    End If

    Dim wb As Object
    Set wb = Application.Workbooks.Add
    EnsureAllSheets wb
    RemoveForeignSheets wb
    Set DossierOpenOrCreate = wb
    Exit Function

Failed:
    Set DossierOpenOrCreate = Nothing
End Function

' 既存ファイルだけを開く(不在・失敗は Nothing)。
'   裁定書27 W9-B7(d): 開く前に DisplayAlerts を退避して False にする。
'   他者ロック・読取専用推奨・リンク更新の**モーダル**が出ると、無人で進む
'   一括実行がそこで固まる(16章 E-51「モーダルを出さない」)。戻しは成功でも
'   失敗でも必ず通す。
Public Function DossierOpen(ByVal pathText As String, ByVal readOnlyMode As Boolean) As Object
    Dim prevAlerts As Boolean
    On Error GoTo Failed
    If Not FileExists(pathText) Then Exit Function
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set DossierOpen = Application.Workbooks.Open(pathText, 0, readOnlyMode)
    Application.DisplayAlerts = prevAlerts
    Exit Function
Failed:
    RestoreAlerts prevAlerts
    Set DossierOpen = Nothing
End Function

' ハンドラ稼働中に On Error Resume Next は書けないので、戻しは別Subへ切り出す。
Private Sub RestoreAlerts(ByVal prevAlerts As Boolean)
    On Error Resume Next
    Application.DisplayAlerts = prevAlerts
End Sub

' 上書き確認を出さずにマクロ無し.xlsxとして保存して閉じる。
'   裁定書9 B8・N3(14章§6): SaveAs の失敗を握り潰さず Boolean で返す
'   (True=SaveAs と Close が成功)。共有フォルダの読取専用・他者ロック・
'   パス長超過で現実に起きる。失敗時もブックは保存せずに閉じて残さない
'   (開いたままのダーティなブックを VerifyRoundTrip が読むと、ディスクでは
'   なくメモリ上の未保存内容と突合して合格してしまうため)。
Public Function DossierSaveAndClose(ByVal wb As Object, ByVal pathText As String) As Boolean
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.DisplayAlerts = False
    wb.SaveAs pathText, C2_XLSX_FORMAT
    Application.DisplayAlerts = prevAlerts
    wb.Close False
    DossierSaveAndClose = True
    Exit Function
Failed:
    modLog.LogError "E0603", C2_SRC & ".DossierSaveAndClose", "save_failed", Err.Number
    Resume CleanUp0
CleanUp0:
    ' Resume でハンドラを抜けてから後始末する(ハンドラ稼働中は再捕捉できない)。
    On Error Resume Next
    Application.DisplayAlerts = prevAlerts
    wb.Close False
    DossierSaveAndClose = False
End Function

' 保存せずに閉じる(読み取り目的で開いたブックの後始末)。
Public Sub DossierClose(ByVal wb As Object)
    On Error GoTo Ignore0
    wb.Close False
    Exit Sub
Ignore0:
End Sub

' シートを名前で引く(無ければ Nothing)。
Public Function DossierSheet(ByVal wb As Object, ByVal sheetTitle As String) As Object
    On Error GoTo NoSheet
    If wb Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).Name = sheetTitle Then
            Set DossierSheet = wb.Worksheets(i)
            Exit Function
        End If
    Next i
    Exit Function
NoSheet:
    Set DossierSheet = Nothing
End Function

' A列基準の最終行。データが無ければ1(見出し行)。
Public Function SheetLastRow(ByVal ws As Object) As Long
    On Error GoTo One1
    SheetLastRow = ws.Cells(ws.Rows.Count, 1).End(C2_DIR_UP).Row
    If SheetLastRow < 1 Then SheetLastRow = 1
    Exit Function
One1:
    SheetLastRow = 1
End Function

' 見出し行を含む矩形を一度だけ読む。1セルだけのRangeは2次元配列にならないので
' 必ず2行以上を読む。
Public Function SheetBlock(ByVal ws As Object, ByVal lastRow As Long) As Variant
    On Error GoTo Empty0
    Dim hi As Long
    hi = lastRow
    If hi < 2 Then hi = 2
    SheetBlock = ws.Range(ws.Cells(1, 1), ws.Cells(hi, C2_SCAN_COLS)).Value
    Exit Function
Empty0:
    SheetBlock = Empty
End Function

' 見出し行の下を全部消す(dossier_profile の積み直し用)。
Public Sub SheetClearRows(ByVal ws As Object)
    On Error GoTo Ignore0
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = SheetLastRow(ws)
    If lastRow < 2 Then Exit Sub
    ws.Range(ws.Rows(2), ws.Rows(lastRow)).Delete
    Exit Sub
Ignore0:
End Sub

' 同じ (case_id, round_no) の行を消す。再保存が二重行にならないための置換で
' あり、別ラウンドの行には触らない(=「追記して育てる」を壊さない)。
Public Sub SheetDropRound(ByVal ws As Object, ByVal caseId As String, ByVal roundNo As Long)
    On Error GoTo Ignore0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = SheetLastRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim blk As Variant
    blk = SheetBlock(ws, lastRow)
    Dim cCase As Long, cRound As Long
    cCase = modUtil.FindHeaderCol(blk, "case_id")
    cRound = modUtil.FindHeaderCol(blk, "round_no")
    If cCase <= 0 Or cRound <= 0 Then Exit Sub

    Dim r As Long
    For r = lastRow To 2 Step -1
        If Trim$(BlockText(blk, r, cCase)) = caseId Then
            If ToLongSafe(blk(r, cRound)) = roundNo Then ws.Rows(r).Delete
        End If
    Next r
    Exit Sub
Ignore0:
End Sub

' 列名で位置を引いて1セル書く。列が無ければ何もしない。書込は必ず SetCellSafe。
Public Sub SheetPutText(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                        ByVal colName As String, ByVal textValue As String)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(r, c), textValue, ws.Name & "/" & colName
End Sub

' 数値列(round_no / seq)。内部生成のLongで数値型を保つ必要があるため
' SetCellSafe(文字列を返す)は通さない(modCaseStore.PutNum と同じ理由)。
Public Sub SheetPutNum(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                       ByVal colName As String, ByVal numValue As Long)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    ws.Cells(r, c).Value = numValue   ' SAFE:const 内部生成のLong(外部由来テキストではない)
End Sub

' ============================================================================
' SheetJoinBySeq - round_no(と note_kind)で絞った行を seq の【値】で並べ直して
'   結合する。行の並び順に頼らないので、利用者が手で行を動かしたファイルでも
'   欠落なく復元できる(32,000字分割の往復一致の要。16章 E-22)。
'   noteKind が "" なら note_kind 列を見ない(dossier_rounds 用)。
' ============================================================================
Public Function SheetJoinBySeq(ByVal ws As Object, ByVal roundNo As Long, _
                               ByVal noteKind As String, ByVal colName As String) As String
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = SheetLastRow(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = SheetBlock(ws, lastRow)
    Dim cRound As Long, cSeq As Long, cVal As Long, cKind As Long
    cRound = modUtil.FindHeaderCol(blk, "round_no")
    cSeq = modUtil.FindHeaderCol(blk, "seq")
    cVal = modUtil.FindHeaderCol(blk, colName)
    cKind = modUtil.FindHeaderCol(blk, "note_kind")
    If cRound <= 0 Or cSeq <= 0 Or cVal <= 0 Then Exit Function

    ' 1周目で seq の最大値(=断片数)を、2周目で seq を添字にして配置する。
    Dim maxSeq As Long
    Dim r As Long, sq As Long
    For r = 2 To lastRow
        If RowMatches(blk, r, cRound, cKind, roundNo, noteKind) Then
            sq = ToLongSafe(blk(r, cSeq))
            If sq > maxSeq Then maxSeq = sq
        End If
    Next r
    If maxSeq < 1 Then Exit Function

    Dim hits() As String
    ReDim hits(1 To maxSeq)
    For r = 2 To lastRow
        If RowMatches(blk, r, cRound, cKind, roundNo, noteKind) Then
            sq = ToLongSafe(blk(r, cSeq))
            If sq >= 1 And sq <= maxSeq Then hits(sq) = BlockText(blk, r, cVal)
        End If
    Next r

    SheetJoinBySeq = modUtil.JoinCellChunks(hits)
End Function

' ============================================================================
' SheetDropCase - 同じ case_id の行を全部消す(裁定書28)。dossier_case /
'   dossier_data / dossier_facts / dossier_judge の積み直し用。SheetDropRound と
'   違い**ラウンドを見ない**(これらの枚は「その案件の最新の全体」を持つため、
'   再保存のたびに丸ごと置き換えるのが正しい)。他の案件の行には触らない。
' ============================================================================
Public Sub SheetDropCase(ByVal ws As Object, ByVal caseId As String)
    On Error GoTo Ignore0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = SheetLastRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim blk As Variant
    blk = SheetBlock(ws, lastRow)
    Dim cCase As Long
    cCase = modUtil.FindHeaderCol(blk, "case_id")
    If cCase <= 0 Then Exit Sub

    Dim r As Long
    For r = lastRow To 2 Step -1
        If Trim$(BlockText(blk, r, cCase)) = caseId Then ws.Rows(r).Delete
    Next r
    Exit Sub
Ignore0:
End Sub

' ============================================================================
' SheetJoinByKey - dossier_data の (case_id, data_key) を seq の【値】で並べ直して
'   結合する(SheetJoinBySeq の data_key 版。13章§2.2 の 32,000字分割の復元)。
'   行の並び順に頼らないので、利用者が手で行を動かしたファイルでも欠落しない。
' ============================================================================
Public Function SheetJoinByKey(ByVal ws As Object, ByVal caseId As String, _
                               ByVal dataKey As String) As String
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = SheetLastRow(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = SheetBlock(ws, lastRow)
    Dim cCase As Long, cKey As Long, cSeq As Long, cVal As Long
    cCase = modUtil.FindHeaderCol(blk, "case_id")
    cKey = modUtil.FindHeaderCol(blk, "data_key")
    cSeq = modUtil.FindHeaderCol(blk, "seq")
    cVal = modUtil.FindHeaderCol(blk, "content")
    If cCase <= 0 Or cKey <= 0 Or cSeq <= 0 Or cVal <= 0 Then Exit Function

    Dim maxSeq As Long
    Dim r As Long, sq As Long
    For r = 2 To lastRow
        If KeyMatches(blk, r, cCase, cKey, caseId, dataKey) Then
            sq = ToLongSafe(blk(r, cSeq))
            If sq > maxSeq Then maxSeq = sq
        End If
    Next r
    If maxSeq < 1 Then Exit Function

    Dim hits() As String
    ReDim hits(1 To maxSeq)
    For r = 2 To lastRow
        If KeyMatches(blk, r, cCase, cKey, caseId, dataKey) Then
            sq = ToLongSafe(blk(r, cSeq))
            If sq >= 1 And sq <= maxSeq Then hits(sq) = BlockText(blk, r, cVal)
        End If
    Next r

    SheetJoinByKey = modUtil.JoinCellChunks(hits)
End Function

Private Function KeyMatches(ByVal blk As Variant, ByVal r As Long, _
                            ByVal cCase As Long, ByVal cKey As Long, _
                            ByVal caseId As String, ByVal dataKey As String) As Boolean
    If Trim$(BlockText(blk, r, cCase)) <> caseId Then Exit Function
    If Trim$(BlockText(blk, r, cKey)) <> dataKey Then Exit Function
    KeyMatches = True
End Function

' そのシートが持つ最大の round_no(0=行無し)。
Public Function SheetMaxRound(ByVal ws As Object) As Long
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = SheetLastRow(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = SheetBlock(ws, lastRow)
    Dim cRound As Long
    cRound = modUtil.FindHeaderCol(blk, "round_no")
    If cRound <= 0 Then Exit Function

    Dim r As Long, v As Long
    For r = 2 To lastRow
        v = ToLongSafe(blk(r, cRound))
        If v > SheetMaxRound Then SheetMaxRound = v
    Next r
End Function

' 読み込み済みブロックの1セルを文字列で返す(列0・範囲外は "")。
Public Function BlockText(ByVal blk As Variant, ByVal r As Long, ByVal c As Long) As String
    On Error GoTo Empty0
    If c <= 0 Then Exit Function
    BlockText = CStr(blk(r, c))
    Exit Function
Empty0:
    BlockText = vbNullString
End Function

' ----------------------------------------------------------------------------
' 内部
' ----------------------------------------------------------------------------

Private Sub EnsureAllSheets(ByVal wb As Object)
    If wb Is Nothing Then Exit Sub
    EnsureSheet wb, "dossier_meta", C2_HDR_META
    EnsureSheet wb, "dossier_profile", C2_HDR_PROFILE
    EnsureSheet wb, "dossier_rounds", C2_HDR_ROUNDS
    EnsureSheet wb, "dossier_facts", C2_HDR_FACTS
    EnsureSheet wb, "dossier_notes", C2_HDR_NOTES
    EnsureSheet wb, "dossier_case", C2_HDR_CASE
    EnsureSheet wb, "dossier_data", C2_HDR_DATA
    EnsureSheet wb, "dossier_judge", C2_HDR_JUDGE
End Sub

' シートが無ければ末尾に足して1行目へ物理名ヘッダを書く。既存なら触らない
' (利用者が育てたファイルの列をこちらの都合で並べ替えない)。
Private Sub EnsureSheet(ByVal wb As Object, ByVal sheetTitle As String, ByVal hdrText As String)
    On Error GoTo Failed

    Dim ws As Object
    Set ws = DossierSheet(wb, sheetTitle)
    If Not ws Is Nothing Then Exit Sub

    Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetTitle

    Dim cols() As String
    cols = modUtil.SplitKeepNonEmpty(hdrText, C2_SEP)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        modUtilText.SetCellSafe ws.Cells(1, i - LBound(cols) + 1), cols(i), sheetTitle & "/header"
    Next i
    Exit Sub

Failed:
    modLog.LogError "E0603", C2_SRC & ".EnsureSheet", "add_failed:" & sheetTitle, Err.Number
    Resume Ignore0
Ignore0:
End Sub

' Workbooks.Add が付けた既定シート(Sheet1 等)を消す。5枚を先に作ってから消すので
' 「可視シートが0枚」にはならない。
Private Sub RemoveForeignSheets(ByVal wb As Object)
    On Error GoTo Failed

    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Dim i As Long
    For i = wb.Worksheets.Count To 1 Step -1
        If InStr(1, C2_SEP & C2_SHEETS & C2_SEP, _
                 C2_SEP & wb.Worksheets(i).Name & C2_SEP, vbBinaryCompare) = 0 Then
            wb.Worksheets(i).Delete
        End If
    Next i

    Application.DisplayAlerts = prevAlerts
    Exit Sub

Failed:
    modLog.LogError "E0603", C2_SRC & ".RemoveForeignSheets", "delete_failed", Err.Number
    Resume Ignore0
Ignore0:
End Sub

Private Function FileExists(ByVal pathText As String) As Boolean
    On Error GoTo NotFound
    If LenB(pathText) = 0 Then Exit Function
    FileExists = (LenB(Dir$(pathText)) > 0)
    Exit Function
NotFound:
    FileExists = False
End Function

Private Function RowMatches(ByVal blk As Variant, ByVal r As Long, _
                            ByVal cRound As Long, ByVal cKind As Long, _
                            ByVal roundNo As Long, ByVal noteKind As String) As Boolean
    If ToLongSafe(blk(r, cRound)) <> roundNo Then Exit Function
    If LenB(noteKind) > 0 Then
        If cKind <= 0 Then Exit Function
        If Trim$(BlockText(blk, r, cKind)) <> noteKind Then Exit Function
    End If
    RowMatches = True
End Function

Private Function ToLongSafe(ByVal v As Variant) As Long
    On Error GoTo Zero0
    Dim t As String
    t = Trim$(CStr(v))
    If LenB(t) = 0 Then Exit Function
    If Not IsNumeric(t) Then Exit Function
    ToLongSafe = CLng(t)
    Exit Function
Zero0:
    ToLongSafe = 0
End Function
