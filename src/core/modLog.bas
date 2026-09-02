Attribute VB_Name = "modLog"
Option Explicit

' ============================================================================
' modLog - err_log / usage_log / run_log への記録(3系統)とローテーション
' ----------------------------------------------------------------------------
' 役割(13章§2.4):
'   err_log  : logged_at / err_code / source / detail / err_number / http_status
'   usage_log: logged_at / event / case_id / detail
'   run_log  : 1行=1 LLM呼び出し。run_at / case_id / round_no / step / play /
'              transport / model / latency_ms / input_chars / output_chars /
'              injected_kb_ids / validate_result / detail / operator
'
' 設計の要点:
'   ・detail は最大400字・【本文非記録】(NFR-S3)。入力やLLM生成の本文を
'     ログへ落とさない。切詰めはここで一括して行うので、呼び出し側は素の
'     メタ文字列を渡してよい。
'   ・列アクセスは列名ベース(13章§6。列番号のハードコード禁止)。見出し行を
'     1回読み、modUtil.FindHeaderCol で位置を引く。見つからない列は黙って
'     飛ばす(ログの1列欠落でアプリを止めない)。
'   ・ログ書き込み自体の失敗はアプリの処理を止めない(「ログで死なない」)。
'     ただし Debug.Print には必ず残す。これが唯一の握り潰し例外であり、
'     ログ以外の失敗を握り潰す口実に使わない。
'   ・セルへの書込は全て modUtilText.SetCellSafe を通す(16章 NFR-S7①)。
'     数値列だけは内部生成のLongであり外部由来テキストではないので、
'     行末に ' SAFE:const を明示したうえで PutNum が直接書く(1箇所に閉じる)。
'
' 移植元: PoC「マイ本棚AI」 src/core/modLog.bas。
'   LogError / LogUsage / EnsureLogSheet / TrimLog(古い側から削る)の骨格を
'   維持し、RPN向けに (a) run_log の3系統目を追加 (b) 列名ベースのアクセスへ
'   変更 (c) 利用者向けメッセージ表(FriendlyMessage / ShowError)を非移植と
'   した。(c)の理由: 19章§4のメッセージ本文はナレッジブック・受信箱といった
'   製品固有の語彙を含み、12章§4が core層への持ち込みを禁じているため。
'   文言表は app/ui 層に置く(本モジュールは記録だけを担う)。
'
' R4: err_log / usage_log / run_log シートへの記録が責務そのもののため、
'   Excelトークンの使用を明示的に許可されたモジュール(12章§4の9本の1つ)。
' ============================================================================

Private Const LOG_SHEET_ERR As String = "err_log"
Private Const LOG_SHEET_USAGE As String = "usage_log"
Private Const LOG_SHEET_RUN As String = "run_log"

' detail の上限(13章§2.4。本文は書かない=NFR-S3)。source も長すぎる値を
' 貼らせないため上限を設ける。
Private Const LOG_DETAIL_MAX As Long = 400
Private Const LOG_SOURCE_MAX As Long = 200

' 見出し行を読む幅。run_log の14列に余裕を見た固定幅で読む(列位置は
' 名前で引くため、この数値は「探索範囲」であって列番号ではない)。
Private Const LOG_HEADER_SCAN_COLS As Long = 24

' xlUp / xlSheetHidden の数値。Excel組み込み定数名を書かずに済ませ、
' LibreOffice側の構文チェックで未定義名にならないようにする。
Private Const LOG_DIR_UP As Long = -4162
Private Const LOG_SHEET_HIDDEN As Long = 0

' ローテーション既定(config log_max_rows。13章§2.3)。
Private Const LOG_MAX_ROWS_DEFAULT As Long = 2000

' ============================================================================
' 純ロジック(Excel非依存。LibreOffice実行テストで直接叩ける。14章§6)
' ----------------------------------------------------------------------------
' 切詰めとローテ判定は「本文を残さない」「古い行から消す」という規約の実体で
' あり、シートI/Oと混ぜると層(a)から一切検査できなくなる。純関数として公開し、
' 記録側(LogError / LogUsage / LogRun / TrimLog)は必ずこれを通す。
' ============================================================================

' detail 列の切詰め(最大400字。13章§2.4・16章 NFR-S3)。本文は残さない規約の
' 実体で、記録側はこれを通してからシートへ書く。
Public Function TruncDetail(ByVal s As String) As String
    TruncDetail = modUtil.SafeLeft(s, LOG_DETAIL_MAX)
End Function

' ローテすべきか(config log_max_rows)。rowCount >= maxRows で True
' (閾値ちょうどで回す)。maxRows <= 0 はローテ無効で常に False。
Public Function ShouldRotate(ByVal rowCount As Long, ByVal maxRows As Long) As Boolean
    If maxRows <= 0 Then Exit Function
    ShouldRotate = (rowCount >= maxRows)
End Function

' ============================================================================
' LogError - err_log へ1行記録する(16章の全エラーコード共通の口)。
' ----------------------------------------------------------------------------
'   errCode   : E01xx-E07xx(正は16章§1の code 列)
'   source    : 発生箇所(モジュール名.関数名 等)
'   detail    : メタ情報のみ。入力本文・LLM生成本文は絶対に渡さない(NFR-S3)
'   errNumber : Err.Number(VBA実行時エラー番号。無ければ0)
'   httpStatus: HTTP応答コード(direct経路のみ。無ければ0)
' ============================================================================
Public Sub LogError(ByVal errCode As String, ByVal source As String, ByVal detail As String, _
                    Optional ByVal errNumber As Long = 0, Optional ByVal httpStatus As Long = 0)
    On Error GoTo Failed

    Dim ws As Object
    Set ws = EnsureLogSheet(LOG_SHEET_ERR)
    If ws Is Nothing Then GoTo Failed

    Dim hdr As Variant
    hdr = HeaderRowOf(ws)
    Dim r As Long
    r = NextRow(ws)

    PutText ws, hdr, r, "logged_at", modUtil.NowStamp()
    PutText ws, hdr, r, "err_code", errCode
    PutText ws, hdr, r, "source", modUtil.SafeLeft(source, LOG_SOURCE_MAX)
    PutText ws, hdr, r, "detail", TruncDetail(detail)
    PutNum ws, hdr, r, "err_number", errNumber
    PutNum ws, hdr, r, "http_status", httpStatus

    TrimLog ws
    Exit Sub

Failed:
    ' ログ書込の失敗でアプリを止めない(唯一の握り潰し例外)。VBEでは追える。
    Debug.Print "[modLog.LogError:書込失敗] " & errCode & " " & source & " : " & detail
End Sub

' ============================================================================
' LogUsage - usage_log へ1行記録する(利用状況・整備優先度シグナル)。
'   列は logged_at / event / case_id / detail の4つ(13章§2.4)。
' ============================================================================
Public Sub LogUsage(ByVal eventName As String, ByVal caseId As String, ByVal detail As String)
    On Error GoTo Failed

    Dim ws As Object
    Set ws = EnsureLogSheet(LOG_SHEET_USAGE)
    If ws Is Nothing Then GoTo Failed

    Dim hdr As Variant
    hdr = HeaderRowOf(ws)
    Dim r As Long
    r = NextRow(ws)

    PutText ws, hdr, r, "logged_at", modUtil.NowStamp()
    PutText ws, hdr, r, "event", eventName
    PutText ws, hdr, r, "case_id", caseId
    PutText ws, hdr, r, "detail", TruncDetail(detail)

    TrimLog ws
    Exit Sub

Failed:
    Debug.Print "[modLog.LogUsage:書込失敗] " & eventName & " " & caseId & " : " & detail
End Sub

' ============================================================================
' LogRun - run_log へ1行記録する(1行=1 LLM呼び出し。NFR-S5 監査可能性)。
' ----------------------------------------------------------------------------
'   14列あるため引数ではなく modTypes.TRunLogRec で受ける(順序の取り違えが
'   コンパイルを通ってしまうのを防ぐ)。rec.run_at が空なら記録時刻で埋める。
'   rec.stepName は run_log の step 列へ書く(Step はVBAの予約語のため
'   フィールド名だけを変えてある)。
'   round_no は受信箱起点(pf/wt/fg)では空を書く(13章§2.4)。
' ============================================================================
Public Sub LogRun(ByRef rec As TRunLogRec)
    On Error GoTo Failed

    Dim ws As Object
    Set ws = EnsureLogSheet(LOG_SHEET_RUN)
    If ws Is Nothing Then GoTo Failed

    Dim hdr As Variant
    hdr = HeaderRowOf(ws)
    Dim r As Long
    r = NextRow(ws)

    Dim stamp As String
    stamp = rec.run_at
    If LenB(stamp) = 0 Then stamp = modUtil.NowStamp()

    PutText ws, hdr, r, "run_at", stamp
    PutText ws, hdr, r, "case_id", rec.case_id
    PutText ws, hdr, r, "round_no", rec.round_no
    PutText ws, hdr, r, "step", rec.stepName
    PutText ws, hdr, r, "play", rec.play
    PutText ws, hdr, r, "transport", rec.transport
    PutText ws, hdr, r, "model", rec.model
    PutNum ws, hdr, r, "latency_ms", rec.latency_ms
    PutNum ws, hdr, r, "input_chars", rec.input_chars
    PutNum ws, hdr, r, "output_chars", rec.output_chars
    PutText ws, hdr, r, "injected_kb_ids", rec.injected_kb_ids
    PutText ws, hdr, r, "validate_result", rec.validate_result
    PutText ws, hdr, r, "detail", TruncDetail(rec.detail)
    PutText ws, hdr, r, "operator", rec.operator

    TrimLog ws
    Exit Sub

Failed:
    Debug.Print "[modLog.LogRun:書込失敗] " & rec.stepName & " " & rec.case_id
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------
' 13章§2.4の列並び。シートが無いときに作る雛形としてのみ使う(既存シートの
' 列位置はここではなく見出し行の実物から引く)。
Private Function HeaderDefOf(ByVal sheetName As String) As Variant
    Select Case sheetName
        Case LOG_SHEET_ERR
            HeaderDefOf = Array("logged_at", "err_code", "source", "detail", _
                                "err_number", "http_status")
        Case LOG_SHEET_USAGE
            HeaderDefOf = Array("logged_at", "event", "case_id", "detail")
        Case Else
            HeaderDefOf = Array("run_at", "case_id", "round_no", "step", "play", _
                                "transport", "model", "latency_ms", "input_chars", _
                                "output_chars", "injected_kb_ids", "validate_result", _
                                "detail", "operator")
    End Select
End Function

Private Function EnsureLogSheet(ByVal sheetName As String) As Object
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error GoTo Failed
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
        Dim header As Variant
        header = HeaderDefOf(sheetName)
        Dim i As Long
        For i = LBound(header) To UBound(header)
            ws.Cells(1, i + 1).Value = header(i)   ' SAFE:const 13章§2.4の固定ヘッダ
        Next i
        HideSheetQuietly ws
        On Error GoTo 0
    End If
    Set EnsureLogSheet = ws
    Exit Function

Failed:
    Set EnsureLogSheet = Nothing
End Function

Private Sub HideSheetQuietly(ByVal ws As Object)
    On Error Resume Next
    ws.Visible = LOG_SHEET_HIDDEN
    On Error GoTo 0
End Sub

Private Function HeaderRowOf(ByVal ws As Object) As Variant
    On Error GoTo Empty0
    HeaderRowOf = ws.Range(ws.Cells(1, 1), ws.Cells(1, LOG_HEADER_SCAN_COLS)).Value
    Exit Function
Empty0:
    HeaderRowOf = Empty
End Function

Private Function NextRow(ByVal ws As Object) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.count, 1).End(LOG_DIR_UP).row
    If r < 1 Then r = 1
    NextRow = r + 1
End Function

' 列名で位置を引いて1セル書く。列が無ければ何もしない(ログの1列欠落で
' アプリを止めない)。書込は必ず modUtilText.SetCellSafe を通す。
Private Sub PutText(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                    ByVal colName As String, ByVal textValue As String)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(r, c), textValue, ws.Name & "/" & colName
End Sub

' 数値列。値は内部生成のLong(経過ミリ秒・字数・エラー番号)であって外部由来
' テキストではないため、テキスト化せず数値のまま書く。集計・並べ替えのために
' 数値型を保つ必要があり、SetCellSafe(文字列を返す)は通せない。
Private Sub PutNum(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                   ByVal colName As String, ByVal numValue As Long)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    ws.Cells(r, c).Value = numValue   ' SAFE:const 内部生成のLong(外部由来テキストではない)
End Sub

' ----------------------------------------------------------------------------
' TrimLog - 古い行を落として上限行数に収める(13章§2.3 log_max_rows)。
' ----------------------------------------------------------------------------
'   新しい行は末尾へ積むので、落とすのは【上側=古い方】。ここを逆にすると
'   障害直後に最も見たい行から消えるという最悪の挙動になる。
'   毎回削ると重いので、上限を1割超えてから上限ちょうどまで一気に削る。
'   log_max_rows <= 0 でローテ無効。
' ----------------------------------------------------------------------------
Private Sub TrimLog(ByVal ws As Object)
    On Error Resume Next
    Dim maxRows As Long
    maxRows = modConfig.GetLong("log_max_rows", LOG_MAX_ROWS_DEFAULT)
    If maxRows <= 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(LOG_DIR_UP).row
    Dim dataRows As Long
    dataRows = lastRow - 1
    ' 判定は ShouldRotate に一本化する(閾値の意味を2箇所に書かない)。毎回削ると
    ' 重いので、実際に削り始めるのは上限を1割超えてからにする=閾値側に余裕を
    ' 足して渡す。削る先は上限ちょうど。
    If Not ShouldRotate(dataRows, maxRows + CLng(Fix(maxRows / 10)) + 1) Then Exit Sub

    ' 残すのは末尾 maxRows 行。ヘッダ(1行目)の直下から余った分だけ消す。
    Dim dropCount As Long
    dropCount = dataRows - maxRows
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + dropCount, 1)).EntireRow.Delete
    On Error GoTo 0
End Sub
