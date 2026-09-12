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

' 裁定書28: data_dir の下のログ置き場(csv の複製先)。
Private Const LOG_CSV_FOLDER As String = "ログ"

' EditRatio(17章 Z-51)の文字計数表。文字コードを添字にした計数表と、その
'   要素が「今回の呼び出しで書かれたか」を示す世代印。呼び出しごとに
'   65,536要素を 0 で埋め直す費用を避けるためだけの作業領域で、呼び出しを
'   またいで意味のある状態は持たない(結果は毎回入力だけで決まる)。
Private gErCount() As Long
Private gErStampAt() As Long
Private gErStamp As Long
Private gErReady As Boolean

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
' CSV の1行組立て(純関数。裁定書28・司令塔裁定「標準CSVで書く」)
' ----------------------------------------------------------------------------
' なぜ要るのか: 本体xlsm は毎朝 D: から消える(裁定書28 確定事実)ので、
'   シートに書いたログも一緒に消える。同じ1行を data_dir\ログ\*.csv へも
'   足しておけば、手順書の「err_log の最後の行を送ってください」が、消えた
'   あとでも成立する。
' 形の規約: **情シス・開発担当が Excel やメモ帳でそのまま開ける標準CSV**に
'   する(司令塔裁定)。したがって
'   ・区切りは**カンマ**。組立ては**フィールドの配列**から行い、値の中の
'     カンマは引用符の中にそのまま残す(列は増えない)
'   ・各フィールドは**常に**二重引用符で囲み、値の中の `"` は `""` にする
'     (囲むか囲まないかを値で分けない=読む側の実装差で崩れない)
'   ・改行(CR/LF)とタブは半角空白へ潰す(1行=1レコードを崩さない。壊れた
'     1行より、1行に収まった読みにくい1行のほうが後で使える)
'   ・先頭の式記号(= + - @)は E-46 と同じ理由で `'` を前置する(CSVを
'     Excelで開いた管理者側で式として実行されないように)
' ============================================================================

' 1フィールドぶんの無害化(常に二重引用符で囲んで返す)。
Public Function CsvField(ByVal s As String) As String
    Dim t As String
    t = Replace(Replace(Replace(s, vbCrLf, " "), vbCr, " "), vbLf, " ")
    t = Replace(t, Chr$(9), " ")

    If LenB(t) > 0 Then
        Select Case Left$(t, 1)
        Case "=", "+", "-", "@"
            t = "'" & t
        End Select
    End If

    CsvField = """" & Replace(t, """", """""") & """"
End Function

' 1行ぶん。**フィールドの配列**を受け取り、1つずつ CsvField に通して
'   カンマでつなぐ。値を先に1本の文字列へ連結してから割り直すことはしない
'   (割り直すと、値の中のカンマや区切り文字がそこで列を分けてしまう。
'    標準CSVでは囲まれたフィールドの中のカンマは列を分けない)。
'   配列は VBA の規則により ByRef で受ける(tools/vba_lint.py の
'   check_udt_byval_param)。
Public Function CsvLineOf(ByRef fields() As String) As String
    Dim outText As String
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        If i > LBound(fields) Then outText = outText & ","
        outText = outText & CsvField(fields(i))
    Next i
    CsvLineOf = outText
End Function

' ============================================================================
' EditRatio - AI原案と人の修正後の「文字ベースの差分率」(0〜100の整数)
' ----------------------------------------------------------------------------
' 17章 Z-51(効果測定の機械計測)の計測核。`sN_json`(AI原案)と `sN_edited`
'   (人が直したもの)を渡すと「どれだけ直したか」を百分率で返す。
'
' 定義(**ここが唯一の値源**): 両方を**文字の多重集合**(どの文字が何個ある
'   か)とみなし、共通する個数 common を数えて
'       差分率 = 100 x (LenA + LenB - 2 x common) / (LenA + LenB)
'   を四捨五入した整数を返す。同一なら 0、共通文字が1つも無ければ 100。
'   両方空なら 0、片方だけ空なら 100。
'
' なぜ編集距離(レーベンシュタイン)にしないか: 編集距離は O(|a|x|b|) で、
'   1セル上限 32,000字(16章 E-22)どうしでは 10億回の走査になり、案件保存の
'   たびに数分固まる(裁定書40 P-M3 で同じ轍を踏んだ)。多重集合の差は
'   O(|a|+|b|) で、**文字の並べ替えだけの修正は 0% と出る**代わりに加筆・
'   削除・書き換えの量を実用十分な精度で表す。この割切りは「どれだけ直したか」
'   の傾向を月次で集計する Z-51 の用途に合わせたもので、厳密な差分表示には
'   使わない(画面に差分を出す機能はこの関数を使わないこと)。
'
' 数え方の実装: 文字コード(0〜65535)を添字にした計数表をモジュール変数として
'   1度だけ確保し、呼び出しごとに増える**世代印**で各要素の有効/無効を決める。
'   毎回 65,536要素を 0 で埋め直す費用を避けるための作法であり、計算結果は
'   埋め直す版と同じである(世代印が上限へ近づいたら表ごと作り直す)。
'   AscW は符号付き16bitを返すので 65536 を足して正へ直す
'   (modKnowledgeRank.CharCodes / modPii.CodePointOf と同じ作法)。
' ============================================================================
Public Function EditRatio(ByVal draftText As String, ByVal editedText As String) As Long
    Dim la As Long, lb As Long
    Dim i As Long, v As Long, common As Long, diff As Long

    la = Len(draftText)
    lb = Len(editedText)

    ' 両方空 = 直していない(0%)。片方だけ空 = 全とっかえ(100%)。
    If la = 0 Then
        If lb = 0 Then Exit Function
        EditRatio = 100
        Exit Function
    End If
    If lb = 0 Then
        EditRatio = 100
        Exit Function
    End If
    ' 完全一致の近道(大小文字・かな漢字はそのまま比較する=vbBinaryCompare)。
    If StrComp(draftText, editedText, vbBinaryCompare) = 0 Then Exit Function

    If Not gErReady Then
        ReDim gErStampAt(0 To 65535)
        ReDim gErCount(0 To 65535)
        gErReady = True
        gErStamp = 0
    End If
    If gErStamp >= 2000000000 Then
        ' 世代印の桁が尽きる前に表ごと 0 へ戻す(ReDim は全要素を 0 にする)。
        ReDim gErStampAt(0 To 65535)
        gErStamp = 0
    End If
    gErStamp = gErStamp + 1

    For i = 1 To la
        v = AscW(Mid$(draftText, i, 1))
        If v < 0 Then v = v + 65536
        If gErStampAt(v) = gErStamp Then
            gErCount(v) = gErCount(v) + 1
        Else
            gErStampAt(v) = gErStamp
            gErCount(v) = 1
        End If
    Next i

    For i = 1 To lb
        v = AscW(Mid$(editedText, i, 1))
        If v < 0 Then v = v + 65536
        If gErStampAt(v) = gErStamp Then
            If gErCount(v) > 0 Then
                common = common + 1
                gErCount(v) = gErCount(v) - 1
            End If
        End If
    Next i

    diff = la + lb - 2 * common
    ' 四捨五入(VBA の Round は銀行丸めなので使わない)。
    EditRatio = Int((100# * diff) / (la + lb) + 0.5)
End Function

' EditRatioNote - usage_log の detail に載せる1項目 `edit_ratio_sN=R` を組む
'   (17章 Z-51・13章§2.4)。書式を1箇所に閉じ、呼び出し側が独自に組み立てる
'   のを防ぐ。stepNo が 1〜5 の外、ratio が 0〜100 の外なら空文字を返す。
'   複数段を並べるときの区切りは呼び出し側が ";" でつなぐ。
Public Function EditRatioNote(ByVal stepNo As Long, ByVal ratio As Long) As String
    If stepNo < 1 Or stepNo > 5 Then Exit Function
    If ratio < 0 Or ratio > 100 Then Exit Function
    EditRatioNote = "edit_ratio_s" & CStr(stepNo) & "=" & CStr(ratio)
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

    Dim stamp As String
    stamp = modUtil.NowStamp()

    PutText ws, hdr, r, "logged_at", stamp
    PutText ws, hdr, r, "err_code", errCode
    PutText ws, hdr, r, "source", modUtil.SafeLeft(source, LOG_SOURCE_MAX)
    PutText ws, hdr, r, "detail", TruncDetail(detail)
    PutNum ws, hdr, r, "err_number", errNumber
    PutNum ws, hdr, r, "http_status", httpStatus

    TrimLog ws

    ' csv 複製(裁定書28)。13章§2.4 の err_log の列順そのままに並べる。
    Dim csvFields(0 To 5) As String
    csvFields(0) = stamp
    csvFields(1) = errCode
    csvFields(2) = modUtil.SafeLeft(source, LOG_SOURCE_MAX)
    csvFields(3) = TruncDetail(detail)
    csvFields(4) = CStr(errNumber)
    csvFields(5) = CStr(httpStatus)
    AppendCsv LOG_SHEET_ERR, CsvLineOf(csvFields)
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

    Dim stamp As String
    stamp = modUtil.NowStamp()

    PutText ws, hdr, r, "logged_at", stamp
    PutText ws, hdr, r, "event", eventName
    PutText ws, hdr, r, "case_id", caseId
    PutText ws, hdr, r, "detail", TruncDetail(detail)

    TrimLog ws

    ' csv 複製(裁定書28 と同じ理由。13章§2.4 の usage_log の列順そのまま)。
    ' 17章 Z-51 の効果測定は各自の data_dir からこの csv を集めて数えるので、
    ' シートにしか残らないと集められない(本体xlsm は毎朝 D: から消える)。
    Dim csvFields(0 To 3) As String
    csvFields(0) = stamp
    csvFields(1) = eventName
    csvFields(2) = caseId
    csvFields(3) = TruncDetail(detail)
    AppendCsv LOG_SHEET_USAGE, CsvLineOf(csvFields)
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

    ' csv 複製(裁定書28)。13章§2.4 の run_log の14列をその順で並べる。
    Dim csvFields(0 To 13) As String
    csvFields(0) = stamp
    csvFields(1) = rec.case_id
    csvFields(2) = rec.round_no
    csvFields(3) = rec.stepName
    csvFields(4) = rec.play
    csvFields(5) = rec.transport
    csvFields(6) = rec.model
    csvFields(7) = CStr(rec.latency_ms)
    csvFields(8) = CStr(rec.input_chars)
    csvFields(9) = CStr(rec.output_chars)
    csvFields(10) = rec.injected_kb_ids
    csvFields(11) = rec.validate_result
    csvFields(12) = TruncDetail(rec.detail)
    csvFields(13) = rec.operator
    AppendCsv LOG_SHEET_RUN, CsvLineOf(csvFields)
    Exit Sub

Failed:
    Debug.Print "[modLog.LogRun:書込失敗] " & rec.stepName & " " & rec.case_id
End Sub

' ----------------------------------------------------------------------------
' AppendCsv - data_dir\ログ\<kind>.csv へ1行足す(裁定書28。13章§2.4)
'   本体のシートへの記録が済んだ**あと**に呼ぶ。ここでの失敗は握り潰し、
'   シート側の1行だけを残す(ログの複製でアプリを止めない)。
'   UTF-8 は modUtilText.Utf8Bytes、フォルダは modUtil.EnsureFolder が持つ。
'   `Print #` は CP932 で書くため使わない(Binary で末尾へ足す)。
' ----------------------------------------------------------------------------
Private Sub AppendCsv(ByVal kind As String, ByVal lineText As String)
    Dim fileNo As Long
    On Error GoTo Failed
    If LenB(lineText) = 0 Then Exit Sub

    Dim dataDir As String
    dataDir = modUtil.ResolveDataDir(modConfig.GetStr("data_dir", vbNullString))
    If LenB(dataDir) = 0 Then Exit Sub

    Dim logDir As String
    logDir = dataDir & modUtil.PathSep() & LOG_CSV_FOLDER
    If Not modUtil.EnsureFolder(logDir) Then Exit Sub

    Dim pathText As String
    pathText = logDir & modUtil.PathSep() & kind & ".csv"

    Dim body As String
    body = lineText & vbCrLf
    Dim n As Long
    n = modUtilText.Utf8Len(body, False)
    If n <= 0 Then Exit Sub

    Dim buf() As Byte
    buf = modUtilText.Utf8Bytes(body, False)

    fileNo = FreeFile
    Open pathText For Binary Access Write As #fileNo
    Put #fileNo, LOF(fileNo) + 1, buf
    Close #fileNo
    Exit Sub

Failed:
    CloseCsvQuiet fileNo
    Debug.Print "[modLog:csv複製の失敗] " & kind
End Sub

' ハンドラ内で On Error Resume Next は効かないため、後始末は別Subへ。
Private Sub CloseCsvQuiet(ByVal fileNo As Long)
    On Error Resume Next
    If fileNo > 0 Then Close #fileNo
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
