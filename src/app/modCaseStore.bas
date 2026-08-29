Attribute VB_Name = "modCaseStore"
Option Explicit

' ============================================================================
' modCaseStore - 案件一覧 / case_data の唯一の入出力口(app層・T-20)
' 責務(12章§2・13章§2.1/§2.2・14章§6): NewCase(13章§1・E-23) / SaveData・
'   LoadData(分割保存と透過結合。E-22) / ResolveStepJson(参照優先。13章§2.2) /
'   SetStatus / InvalidateDownstream(E-10) / RepairStates(E-12) / FreezeRound。
' 配置: app層(12章§2)。シート名という製品固有語彙を持つため core層 へ置けない。
' R4: 案件シートI/Oが責務そのものでExcelトークン許可10本の1つ(12章§4)。見るのは
'   案件一覧 と case_data の2枚だけ。再描画と確認ダイアログは ui層 modUICase。
' 列アクセスは列名ベース(13章冒頭)。セル書込は全て modUtilText.SetCellSafe
'   (16章 NFR-S7①)。数値列だけは内部生成のLongなので ' SAFE:const を明示。
' 純ロジックの分離: 採番・参照優先・状態遷移・無効化対象はシートI/Oを含まない
'   関数へ切り出し、うち BuildCaseId / IsValidCaseId / CanTransition /
'   ResolveDataKey の4本は14章§6が公開を宣言している(裁定書6 項目7)。規約を
'   シートI/Oに閉じ込めると層(a)から検査できない(W2aの実害)。
' ============================================================================

' --- シート名(13章§2.1・§2.2) ---
Private Const CS_SHEET_CASES As String = "案件一覧"
Private Const CS_SHEET_DATA As String = "case_data"

' 1セルに入れる最大字数(16章 E-22。物理上限32,767字の手前で切る)。
Private Const CS_CHUNK_CHARS As Long = 32000

' xlUp の数値(組込定数名を書かず LO の構文チェックで未定義名にしない)。
Private Const CS_DIR_UP As Long = -4162

' 見出し行を読む幅。案件一覧23列に余裕を見た「探索範囲」で列番号ではない。
Private Const CS_SCAN_COLS As Long = 32

' 同日連番の上限(16章 E-23「999まで対応」)。
Private Const CS_SERIAL_MAX As Long = 999

' data_key / 状態のリスト区切り。modUtil.AppendIdList と同じ ";" を使う。
Private Const CS_SEP As String = ";"

' 新規案件の既定値と error(13章§2.1)。
Private Const CS_STATUS_NEW As String = "draft"
Private Const CS_TIER_DEFAULT As String = "t1_quick"
Private Const CS_VARIANT_DEFAULT As String = "proposal"
Private Const CS_STATUS_ERROR As String = "error"

' case_type の enum(13章§2.1・19章§3)。
Private Const CS_CASE_TYPES As String = "new;renewal"

' status の enum 全8値(同)。
Private Const CS_STATUSES As String = "draft;s1_done;s2_done;s3_done;s4_done;exported;feedback_done;error"

' status の正順(11章§4)。error はこの並びの外側(任意の状態から入り、
' failed_step の再実行で並びのどこへでも戻る)。
Private Const CS_STATUS_ORDER As String = "draft;s1_done;s2_done;s3_done;s4_done;exported;feedback_done"

' status のうち s4_done より上位(RepairStates で降格させない2値)。
Private Const CS_STATUSES_ABOVE_S4 As String = "exported;feedback_done"

' data_key の enum 全28値(13章§2.2・19章§3と完全一致)。表に無いキーは拒否。
Private Const CS_DATA_KEYS As String = "input_hp;input_yuho;input_memo;input_contract;input_prev_renewal;input_dossier;input_field_notes;input_coverage_note;input_hearing_answers;s1_json;s2_json;s3_json;s4_json;s2c_json;s3c_json;s2r_json;s3r_json;s2_prev_json;s1_edited;s2_edited;s3_edited;s4_edited;s1_json_failed;s2_json_failed;s3_json_failed;s4_json_failed;sparring_u;sparring_a"

' --- 純ロジック(Excel非依存)。シート操作側は必ずここを通す ---

' BuildCaseId - 案件IDの組立(13章§1: C-YYYYMMDD-NNN。14章§6)。dayText は
' yyyymmdd の8桁ちょうど(数字のみ)、seq は 1..999。範囲外は ""(呼び出し側が
' 枯渇・不正として扱う。例外は投げない)。引数名が dayText なのは datePart が
' VBA組込関数 DatePart と衝突するため(14章§6に同じ注記がある)。
Public Function BuildCaseId(ByVal dayText As String, ByVal seq As Long) As String
    If Len(dayText) <> 8 Then Exit Function
    If Not IsAllDigits(dayText) Then Exit Function
    If seq < 1 Or seq > CS_SERIAL_MAX Then Exit Function
    BuildCaseId = "C-" & dayText & "-" & Right$("00" & CStr(seq), 3)
End Function

' IsValidCaseId - "C-"+数字8桁+"-"+数字3桁(001..999)ちょうどか(14章§6)。
' 前後の空白は許さない。
Public Function IsValidCaseId(ByVal id As String) As Boolean
    If Len(id) <> 14 Then Exit Function
    If Left$(id, 2) <> "C-" Then Exit Function
    If Mid$(id, 11, 1) <> "-" Then Exit Function
    If Not IsAllDigits(Mid$(id, 3, 8)) Then Exit Function
    Dim tailText As String
    tailText = Right$(id, 3)
    If Not IsAllDigits(tailText) Then Exit Function
    IsValidCaseId = (CLng(tailText) >= 1)
End Function

' 半角数字だけで構成されているか(空は False)。
Private Function IsAllDigits(ByVal s As String) As Boolean
    If LenB(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        If InStr(1, "0123456789", Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsAllDigits = True
End Function

' CanTransition - 案件statusの遷移可否(11章§4 + 16章 E-12。14章§6)。許すのは
' (1)正順の1段進み (2)任意の状態から error へ (3)error からの復帰(failed_step
' の再実行なので戻り先は正順のどこでもよい。10章NFR-R3)。自己遷移・段飛ばし・
' 巻き戻しは False。
Public Function CanTransition(ByVal fromStatus As String, ByVal toStatus As String) As Boolean
    Dim f As String, tgt As String
    f = Trim$(fromStatus)
    tgt = Trim$(toStatus)
    If Not IsListedValue(CS_STATUSES, f) Then Exit Function
    If Not IsListedValue(CS_STATUSES, tgt) Then Exit Function
    If StrComp(f, tgt, vbBinaryCompare) = 0 Then Exit Function
    If StrComp(tgt, CS_STATUS_ERROR, vbBinaryCompare) = 0 Then
        CanTransition = True
        Exit Function
    End If
    If StrComp(f, CS_STATUS_ERROR, vbBinaryCompare) = 0 Then
        CanTransition = True
        Exit Function
    End If
    CanTransition = (OrderOf(tgt) = OrderOf(f) + 1)
End Function

' CS_STATUS_ORDER 上の位置(0始まり)。並びに無い(=error)は -1。
Private Function OrderOf(ByVal statusText As String) As Long
    OrderOf = -1
    Dim parts As Variant
    Dim i As Long
    parts = Split(CS_STATUS_ORDER, CS_SEP)
    For i = 0 To UBound(parts)
        If StrComp(CStr(parts(i)), statusText, vbBinaryCompare) = 0 Then OrderOf = i
    Next i
End Function

' ResolveDataKey - 下流Stepが読むべき data_key を1つ返す純核(13章§2.2・14章§6)。
' N=2,3 は sN_edited > sNr_json > sN_json / N=1,4 は sN_edited > sN_json
' (S1/S4に改訂パスは無いので hasRevised は見ない)。無ければ ""。範囲外も ""。
Public Function ResolveDataKey(ByVal stepNo As Long, ByVal hasEdited As Boolean, _
                               ByVal hasRevised As Boolean, ByVal hasJson As Boolean) As String
    If stepNo < 1 Or stepNo > 4 Then Exit Function
    Dim n As String
    n = CStr(stepNo)
    If hasEdited Then
        ResolveDataKey = "s" & n & "_edited"
        Exit Function
    End If
    If stepNo = 2 Or stepNo = 3 Then
        If hasRevised Then
            ResolveDataKey = "s" & n & "r_json"
            Exit Function
        End If
    End If
    If hasJson Then ResolveDataKey = "s" & n & "_json"
End Function

' 案件IDから同日ぶんの連番を取り出す(日付不一致・形違いは0)。IDの形の判定は
' IsValidCaseId が唯一の値源(2箇所で書かない)。
Private Function SerialOfCaseId(ByVal caseId As String, ByVal dayText As String) As Long
    If Not IsValidCaseId(caseId) Then Exit Function
    If Mid$(caseId, 3, 8) <> dayText Then Exit Function
    SerialOfCaseId = CLng(Right$(caseId, 3))
End Function

' ";" 区切りの enum リストに値が含まれるか(前後空白は無視・大小文字は区別)。
Private Function IsListedValue(ByVal listText As String, ByVal valueText As String) As Boolean
    Dim v As String
    v = Trim$(valueText)
    If LenB(v) = 0 Then Exit Function
    IsListedValue = (InStr(1, CS_SEP & listText & CS_SEP, CS_SEP & v & CS_SEP, vbBinaryCompare) > 0)
End Function

' その1Stepぶんの成果物 data_key(無効化の単位)。input_* と s2_prev_json は含めない。
Private Function StepArtifactKeys(ByVal stepNo As Long) As String
    Dim n As String
    Select Case stepNo
        Case 1, 4
            n = CStr(stepNo)
            StepArtifactKeys = "s" & n & "_json;s" & n & "_edited;s" & n & "_json_failed"
        Case 2, 3
            n = CStr(stepNo)
            StepArtifactKeys = "s" & n & "_json;s" & n & "_edited;s" & n & "r_json;s" & _
                               n & "c_json;s" & n & "_json_failed"
    End Select
End Function

' fromStepNo より下流の成果物 data_key を列挙する(16章 E-10)。=4 なら空。
Private Function DownstreamKeys(ByVal fromStepNo As Long) As String
    Dim acc As String
    Dim i As Long
    For i = fromStepNo + 1 To 4
        Dim grp As String
        grp = StepArtifactKeys(i)
        If LenB(grp) > 0 Then
            If LenB(acc) = 0 Then
                acc = grp
            Else
                acc = acc & CS_SEP & grp
            End If
        End If
    Next i
    DownstreamKeys = acc
End Function

' last_ok_step から status を再導出(13章§2.1)。0 -> draft、1..4 -> sN_done
' (CS_STATUS_ORDER の並びがそのまま索引になる)。
Private Function StatusForStep(ByVal stepNo As Long) As String
    StatusForStep = CS_STATUS_NEW
    If stepNo < 1 Or stepNo > 4 Then Exit Function
    Dim parts As Variant
    parts = Split(CS_STATUS_ORDER, CS_SEP)
    StatusForStep = CStr(parts(stepNo))
End Function

' 16章 E-12 の状態再導出。failed_step が非空なら error を維持。s4_done より上位
' (exported / feedback_done)は effStep が4なら降格させない(出力の事実を消さない)。
Private Function RepairedStatus(ByVal currentStatus As String, ByVal failedStep As String, _
                                ByVal effStep As Long) As String
    If LenB(Trim$(failedStep)) > 0 Then
        RepairedStatus = CS_STATUS_ERROR
        Exit Function
    End If
    If effStep >= 4 Then
        If IsListedValue(CS_STATUSES_ABOVE_S4, currentStatus) Then
            RepairedStatus = Trim$(currentStatus)
            Exit Function
        End If
    End If
    RepairedStatus = StatusForStep(effStep)
End Function

' NewCase - 案件を1件起こして case_id を返す(13章§1・16章 E-23)。採番は
'   C-YYYYMMDD-NNN。当日の使用済み最大連番の次から始め、衝突したら E0605 を
'   記録して次の番号へ(999で枯渇)。失敗時は ""(例外は投げない)。channel /
'   kanji / bid / reins / industry_name は §6 のシグネチャが受け取らないため
'   空のまま起こす(13章§2.1「必須(○)の適用時点」= 実行前検証で埋まっていれば
'   よい。裁定書6 項目9)。
Public Function NewCase(ByVal company As String, ByVal industryCode As String, _
                        ByVal caseType As String) As String
    On Error GoTo Failed

    If LenB(Trim$(company)) = 0 Or LenB(Trim$(industryCode)) = 0 Then
        modLog.LogError "E0101", "modCaseStore.NewCase", "empty_required:company/industry_code"
        Exit Function
    End If
    If Not IsListedValue(CS_CASE_TYPES, caseType) Then
        modLog.LogError "E0101", "modCaseStore.NewCase", "invalid_case_type"
        Exit Function
    End If

    Dim ws As Object
    Set ws = SheetOf(CS_SHEET_CASES)
    If ws Is Nothing Then
        modLog.LogError "E0603", "modCaseStore.NewCase", "sheet_missing:cases"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    Dim blk As Variant
    blk = ReadBlock(ws, lastRow)
    Dim cId As Long
    cId = modUtil.FindHeaderCol(blk, "case_id")
    If cId <= 0 Then
        modLog.LogError "E0603", "modCaseStore.NewCase", "header_missing:case_id"
        Exit Function
    End If

    Dim dayText As String
    dayText = modUtilText.IsoDateCompact(Date)

    ' 当日ぶんの使用済み最大連番。
    Dim usedMax As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim sn As Long
        sn = SerialOfCaseId(Trim$(CStr(blk(r, cId))), dayText)
        If sn > usedMax Then usedMax = sn
    Next r

    Dim serialNo As Long
    Dim newId As String
    serialNo = usedMax
    Do
        serialNo = serialNo + 1
        If serialNo > CS_SERIAL_MAX Then
            modLog.LogError "E0605", "modCaseStore.NewCase", "serial_exhausted:" & dayText
            Exit Function
        End If
        newId = BuildCaseId(dayText, serialNo)
        If LenB(newId) = 0 Then
            modLog.LogError "E0605", "modCaseStore.NewCase", "id_build_failed:" & dayText
            Exit Function
        End If
        If RowOfCase(blk, lastRow, cId, newId) <= 0 Then Exit Do
        ' 既存と衝突(E-23)。記録して次の番号へ。
        modLog.LogError "E0605", "modCaseStore.NewCase", "id_collision:" & CStr(serialNo)
    Loop

    Dim wr As Long
    wr = lastRow + 1
    Dim stampText As String
    stampText = modUtil.NowStamp()

    PutText ws, blk, wr, "case_id", newId
    PutText ws, blk, wr, "case_type", Trim$(caseType)
    PutText ws, blk, wr, "dossier_tier", CS_TIER_DEFAULT
    PutText ws, blk, wr, "company", company
    PutText ws, blk, wr, "industry_code", industryCode
    PutText ws, blk, wr, "status", CS_STATUS_NEW
    PutText ws, blk, wr, "created_at", stampText
    PutText ws, blk, wr, "updated_at", stampText
    PutText ws, blk, wr, "owner", OwnerName()
    PutText ws, blk, wr, "s4_variant", CS_VARIANT_DEFAULT
    PutNum ws, blk, wr, "round_no", 1
    PutNum ws, blk, wr, "last_ok_step", 0

    NewCase = newId
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.NewCase", "unexpected", Err.Number
    NewCase = vbNullString
End Function

' SaveData - case_data へ1本のテキストを分割保存する(13章§2.2・16章 E-22)。
'   同じ (case_id, data_key) の既存行を全て消してから32,000字ごとの断片を
'   seq=1.. で積み直す(追記ではなく置換)。content が空なら削除だけで True。
'   断片は SetCellSafe を通り先頭が = + - @ なら "'" が前置されるが、.Value で
'   読み戻すと元へ戻る(往復一致は保たれる。NFR-S7①)。
Public Function SaveData(ByVal caseId As String, ByVal dataKey As String, _
                         ByVal content As String) As Boolean
    On Error GoTo Failed

    If LenB(Trim$(caseId)) = 0 Then
        modLog.LogError "E0101", "modCaseStore.SaveData", "empty_case_id"
        Exit Function
    End If
    If Not IsListedValue(CS_DATA_KEYS, dataKey) Then
        modLog.LogError "E0101", "modCaseStore.SaveData", "invalid_data_key:" & dataKey
        Exit Function
    End If

    Dim ws As Object
    Set ws = SheetOf(CS_SHEET_DATA)
    If ws Is Nothing Then
        modLog.LogError "E0603", "modCaseStore.SaveData", "sheet_missing:case_data"
        Exit Function
    End If

    DropRowsOf ws, Trim$(caseId), Trim$(dataKey)

    If LenB(content) = 0 Then
        SaveData = True
        Exit Function
    End If

    Dim parts() As String
    parts = modUtil.SplitForCells(content, CS_CHUNK_CHARS)

    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    Dim hdr As Variant
    hdr = ReadBlock(ws, 2)
    Dim stampText As String
    stampText = modUtil.NowStamp()

    Dim i As Long
    Dim seqNo As Long
    For i = LBound(parts) To UBound(parts)
        seqNo = i - LBound(parts) + 1
        Dim wr As Long
        wr = lastRow + seqNo
        PutText ws, hdr, wr, "case_id", Trim$(caseId)
        PutText ws, hdr, wr, "data_key", Trim$(dataKey)
        PutNum ws, hdr, wr, "seq", seqNo
        PutText ws, hdr, wr, "content", parts(i)
        PutText ws, hdr, wr, "saved_at", stampText
    Next i

    SaveData = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.SaveData", "write_failed:" & dataKey, Err.Number
    SaveData = False
End Function

' LoadData - 断片を seq 順に結合して返す(透過処理)。行の並び順ではなく seq の値で
'   並べ直すので手で行を動かしたブックでも復元できる。該当が無ければ ""。
Public Function LoadData(ByVal caseId As String, ByVal dataKey As String) As String
    On Error GoTo Failed

    If LenB(Trim$(caseId)) = 0 Then Exit Function

    Dim ws As Object
    Set ws = SheetOf(CS_SHEET_DATA)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = ReadBlock(ws, lastRow)
    Dim cCase As Long, cKey As Long, cSeq As Long, cBody As Long
    cCase = modUtil.FindHeaderCol(blk, "case_id")
    cKey = modUtil.FindHeaderCol(blk, "data_key")
    cSeq = modUtil.FindHeaderCol(blk, "seq")
    cBody = modUtil.FindHeaderCol(blk, "content")
    If cCase <= 0 Or cKey <= 0 Or cSeq <= 0 Or cBody <= 0 Then
        modLog.LogError "E0603", "modCaseStore.LoadData", "header_missing:case_data"
        Exit Function
    End If

    Dim wantCase As String, wantKey As String
    wantCase = Trim$(caseId)
    wantKey = Trim$(dataKey)

    ' 1周目: seq の最大値=断片数を求める。
    Dim maxSeq As Long
    Dim r As Long
    For r = 2 To lastRow
        If MatchesRow(blk, r, cCase, cKey, wantCase, wantKey) Then
            Dim sq1 As Long
            sq1 = ToLongSafe(blk(r, cSeq))
            If sq1 > maxSeq Then maxSeq = sq1
        End If
    Next r
    If maxSeq < 1 Then Exit Function

    ' 2周目: seq を添字として配置する(欠番は空断片で残る)。
    Dim parts() As String
    ReDim parts(1 To maxSeq)
    For r = 2 To lastRow
        If MatchesRow(blk, r, cCase, cKey, wantCase, wantKey) Then
            Dim sq2 As Long
            sq2 = ToLongSafe(blk(r, cSeq))
            If sq2 >= 1 And sq2 <= maxSeq Then parts(sq2) = CStr(blk(r, cBody))
        End If
    Next r

    LoadData = modUtil.JoinCellChunks(parts)
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.LoadData", "read_failed:" & dataKey, Err.Number
    LoadData = vbNullString
End Function

' ResolveStepJson - 下流Stepが参照すべきJSONを一元解決する(13章§2.2)。
'   優先順の判断は純核 ResolveDataKey が唯一の値源。ここは「読んで渡す」だけで
'   分岐を持たない(答えを2箇所に書かない)。
Public Function ResolveStepJson(ByVal caseId As String, ByVal stepNo As Long) As String
    If stepNo < 1 Or stepNo > 4 Then Exit Function
    Dim n As String
    n = CStr(stepNo)

    Dim editedText As String, revisedText As String, jsonText As String
    editedText = LoadData(caseId, "s" & n & "_edited")
    If stepNo = 2 Or stepNo = 3 Then revisedText = LoadData(caseId, "s" & n & "r_json")
    jsonText = LoadData(caseId, "s" & n & "_json")

    Select Case ResolveDataKey(stepNo, LenB(Trim$(editedText)) > 0, _
                               LenB(Trim$(revisedText)) > 0, LenB(Trim$(jsonText)) > 0)
        Case "s" & n & "_edited"
            ResolveStepJson = editedText
        Case "s" & n & "r_json"
            ResolveStepJson = revisedText
        Case "s" & n & "_json"
            ResolveStepJson = jsonText
    End Select
End Function

' SetStatus - 案件一覧の status を書き換える(13章§2.1の enum 8値のみ)。
Public Function SetStatus(ByVal caseId As String, ByVal statusText As String) As Boolean
    On Error GoTo Failed

    If Not IsListedValue(CS_STATUSES, statusText) Then
        modLog.LogError "E0101", "modCaseStore.SetStatus", "invalid_status:" & statusText
        Exit Function
    End If

    Dim ws As Object
    Set ws = SheetOf(CS_SHEET_CASES)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    Dim blk As Variant
    blk = ReadBlock(ws, lastRow)
    Dim cId As Long
    cId = modUtil.FindHeaderCol(blk, "case_id")
    If cId <= 0 Then Exit Function

    Dim rowNo As Long
    rowNo = RowOfCase(blk, lastRow, cId, Trim$(caseId))
    If rowNo <= 0 Then Exit Function

    PutText ws, blk, rowNo, "status", Trim$(statusText)
    PutText ws, blk, rowNo, "updated_at", modUtil.NowStamp()
    SetStatus = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.SetStatus", "write_failed", Err.Number
    SetStatus = False
End Function

' InvalidateDownstream - 上流を再実行するとき下流の成果物を無効化する(E-10)。
'   fromStepNo より下流の成果物 data_key を消し、last_ok_step を fromStepNo へ
'   戻して status を再導出する。消したキーは E0601 の detail へ列挙(黙って
'   消さない)。input_* と s2_prev_json は成果物ではないので消さない。再描画と
'   確認ダイアログは ui層 modUICase。
Public Sub InvalidateDownstream(ByVal caseId As String, ByVal fromStepNo As Long)
    On Error GoTo Failed

    Dim keepStep As Long
    keepStep = fromStepNo
    If keepStep < 0 Then keepStep = 0
    If keepStep > 4 Then keepStep = 4

    Dim keyList As String
    keyList = DownstreamKeys(keepStep)
    If LenB(keyList) = 0 Then Exit Sub

    Dim ws As Object
    Set ws = SheetOf(CS_SHEET_DATA)
    If ws Is Nothing Then Exit Sub

    Dim keyArr() As String
    keyArr = Split(keyList, CS_SEP)

    Dim removed As String
    Dim i As Long
    For i = LBound(keyArr) To UBound(keyArr)
        If DropRowsOf(ws, Trim$(caseId), keyArr(i)) > 0 Then
            removed = modUtil.AppendIdList(removed, keyArr(i))
        End If
    Next i

    ApplyRepairedState Trim$(caseId), keepStep

    ' 無効化した data_key を必ず記録する(E-10。キー名を残す)。
    modLog.LogError "E0601", "modCaseStore.InvalidateDownstream", _
                    "from_step=" & CStr(keepStep) & " invalidated=" & removed
    Exit Sub

Failed:
    modLog.LogError "E0603", "modCaseStore.InvalidateDownstream", "failed", Err.Number
End Sub

' RepairStates - 起動時の状態整合修復(16章 E-12・12章§2.1 手順③)。
'   (1) sN_json(合格済のみ)の最大Stepと last_ok_step を突合し小さいほうを採る
'   (2) failed_step が非空なら status=error を維持し failed_step 自体は
'       書き換えない(失敗の記憶を消さない) (3) 導出に使うのは sN_json だけ
'   (4) ui_lock は modUIProgress のモジュール変数でプロセス終了時に自然消滅
'       (E-11/E-51)。app層から ui層は呼べない(R1)ので何もしない (5) 修復件数を
'       返す(0件でも起動は続行)。
Public Function RepairStates() As Long
    On Error GoTo Failed

    Dim ws As Object
    Set ws = SheetOf(CS_SHEET_CASES)
    If ws Is Nothing Then
        modLog.LogError "E0603", "modCaseStore.RepairStates", "sheet_missing:cases"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = ReadBlock(ws, lastRow)
    Dim cId As Long, cStatus As Long, cLastOk As Long, cFailed As Long
    cId = modUtil.FindHeaderCol(blk, "case_id")
    cStatus = modUtil.FindHeaderCol(blk, "status")
    cLastOk = modUtil.FindHeaderCol(blk, "last_ok_step")
    cFailed = modUtil.FindHeaderCol(blk, "failed_step")
    If cId <= 0 Or cStatus <= 0 Or cLastOk <= 0 Then
        modLog.LogError "E0603", "modCaseStore.RepairStates", "header_missing:cases"
        Exit Function
    End If

    Dim fixedCount As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim rowCase As String
        rowCase = Trim$(CStr(blk(r, cId)))
        If LenB(rowCase) > 0 Then
            Dim effStep As Long
            Dim storedOk As Long
            storedOk = ToLongSafe(blk(r, cLastOk))
            ' E-12(1) の「突合」= sN_json 側と last_ok_step 側の小さいほう。
            effStep = MaxOkStepOf(rowCase)
            If storedOk < effStep Then effStep = storedOk
            If effStep < 0 Then effStep = 0

            Dim failedText As String
            failedText = vbNullString
            If cFailed > 0 Then failedText = Trim$(CStr(blk(r, cFailed)))

            Dim curStatus As String, wantStatus As String
            curStatus = Trim$(CStr(blk(r, cStatus)))
            wantStatus = RepairedStatus(curStatus, failedText, effStep)

            Dim didFix As Boolean
            didFix = False
            If storedOk <> effStep Then
                PutNum ws, blk, r, "last_ok_step", effStep
                didFix = True
            End If
            If curStatus <> wantStatus Then
                PutText ws, blk, r, "status", wantStatus
                didFix = True
            End If
            If didFix Then
                PutText ws, blk, r, "updated_at", modUtil.NowStamp()
                fixedCount = fixedCount + 1
            End If
        End If
    Next r

    If fixedCount > 0 Then
        modLog.LogError "E0603", "modCaseStore.RepairStates", "repaired=" & CStr(fixedCount)
    End If
    RepairStates = fixedCount
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.RepairStates", "repair_failed", Err.Number
    RepairStates = 0
End Function

' FreezeRound - ラウンドを確定する(FR-35マルチラウンド・13章§2.2)。
'   (1) 参照優先で解決したS2の1本を s2_prev_json へ退避(1本のみ・世代管理なし)
'   (2) round_no を+1 (3) 新しい round_no を返す。案件が無い・退避に失敗したら
'   0 を返し round_no を進めない(退避できていないのに進む状態を作らない)。
Public Function FreezeRound(ByVal caseId As String) As Long
    On Error GoTo Failed

    Dim ws As Object
    Set ws = SheetOf(CS_SHEET_CASES)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    Dim blk As Variant
    blk = ReadBlock(ws, lastRow)
    Dim cId As Long, cRound As Long
    cId = modUtil.FindHeaderCol(blk, "case_id")
    cRound = modUtil.FindHeaderCol(blk, "round_no")
    If cId <= 0 Or cRound <= 0 Then
        modLog.LogError "E0603", "modCaseStore.FreezeRound", "header_missing:round_no"
        Exit Function
    End If

    Dim rowNo As Long
    rowNo = RowOfCase(blk, lastRow, cId, Trim$(caseId))
    If rowNo <= 0 Then Exit Function

    Dim s2Text As String
    s2Text = ResolveStepJson(caseId, 2)
    If LenB(Trim$(s2Text)) > 0 Then
        If Not SaveData(caseId, "s2_prev_json", s2Text) Then
            modLog.LogError "E0603", "modCaseStore.FreezeRound", "prev_save_failed"
            Exit Function
        End If
    End If

    Dim curRound As Long
    curRound = ToLongSafe(blk(rowNo, cRound))
    If curRound < 1 Then curRound = 1

    PutNum ws, blk, rowNo, "round_no", curRound + 1
    PutText ws, blk, rowNo, "updated_at", modUtil.NowStamp()
    FreezeRound = curRound + 1
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.FreezeRound", "freeze_failed", Err.Number
    FreezeRound = 0
End Function

' --- 内部ヘルパー(シートI/O) ---

' 名前でシートを取る。無ければ Nothing(実行時に生やすと列定義の欠けた表になる)。
Private Function SheetOf(ByVal sheetName As String) As Object
    On Error GoTo NoSheet
    Set SheetOf = ThisWorkbook.Worksheets(sheetName)
    Exit Function
NoSheet:
    Set SheetOf = Nothing
End Function

' Application.UserName(13章§2.1 owner)。取得できない環境では ""。
Private Function OwnerName() As String
    On Error GoTo NoName
    OwnerName = CStr(Application.UserName)
    Exit Function
NoName:
    OwnerName = vbNullString
End Function

' A列基準の最終行。データが無ければ1(見出し行)。
Private Function LastRowOf(ByVal ws As Object) As Long
    Dim r As Long
    On Error GoTo One1
    r = ws.Cells(ws.Rows.count, 1).End(CS_DIR_UP).row
    If r < 1 Then r = 1
    LastRowOf = r
    Exit Function
One1:
    LastRowOf = 1
End Function

' 見出し行を含む矩形を一度だけ読む。1セルだけの Range は2次元配列にならないので
' 必ず2行以上を読む。
Private Function ReadBlock(ByVal ws As Object, ByVal lastRow As Long) As Variant
    On Error GoTo Empty0
    Dim hi As Long
    hi = lastRow
    If hi < 2 Then hi = 2
    ReadBlock = ws.Range(ws.Cells(1, 1), ws.Cells(hi, CS_SCAN_COLS)).Value
    Exit Function
Empty0:
    ReadBlock = Empty
End Function

' 読み込み済みブロックから case_id 一致行を探す。見つからなければ0。
Private Function RowOfCase(ByVal blk As Variant, ByVal lastRow As Long, _
                           ByVal caseCol As Long, ByVal caseId As String) As Long
    On Error GoTo NotFound
    If caseCol <= 0 Then Exit Function
    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(blk(r, caseCol))) = caseId Then
            RowOfCase = r
            Exit Function
        End If
    Next r
    Exit Function
NotFound:
    RowOfCase = 0
End Function

' case_data の1行が (case_id, data_key) に一致するか。
Private Function MatchesRow(ByVal blk As Variant, ByVal r As Long, _
                            ByVal caseCol As Long, ByVal keyCol As Long, _
                            ByVal caseId As String, ByVal dataKey As String) As Boolean
    On Error GoTo NoMatch
    If Trim$(CStr(blk(r, caseCol))) <> caseId Then Exit Function
    If Trim$(CStr(blk(r, keyCol))) <> dataKey Then Exit Function
    MatchesRow = True
    Exit Function
NoMatch:
    MatchesRow = False
End Function

' 同じ (case_id, data_key) の行を全削除し件数を返す(下から回すのは行ずれ回避)。
Private Function DropRowsOf(ByVal ws As Object, ByVal caseId As String, _
                            ByVal dataKey As String) As Long
    On Error GoTo Done0
    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = ReadBlock(ws, lastRow)
    Dim cCase As Long, cKey As Long
    cCase = modUtil.FindHeaderCol(blk, "case_id")
    cKey = modUtil.FindHeaderCol(blk, "data_key")
    If cCase <= 0 Or cKey <= 0 Then Exit Function

    Dim n As Long
    Dim r As Long
    For r = lastRow To 2 Step -1
        If MatchesRow(blk, r, cCase, cKey, caseId, Trim$(dataKey)) Then
            ws.Rows(r).Delete
            n = n + 1
        End If
    Next r
    DropRowsOf = n
    Exit Function
Done0:
    DropRowsOf = 0
End Function

' sN_json(検証合格済のみ)が示す最大Step。1本も無ければ0(E-12(1)(3))。
Private Function MaxOkStepOf(ByVal caseId As String) As Long
    Dim n As Long
    For n = 4 To 1 Step -1
        If LenB(Trim$(LoadData(caseId, "s" & CStr(n) & "_json"))) > 0 Then
            MaxOkStepOf = n
            Exit Function
        End If
    Next n
End Function

' last_ok_step と status を「そのStepまで進んだ状態」へ揃える(判定は RepairedStatus)。
Private Function ApplyRepairedState(ByVal caseId As String, ByVal keepStep As Long) As Boolean
    On Error GoTo Failed
    Dim ws As Object
    Set ws = SheetOf(CS_SHEET_CASES)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    Dim blk As Variant
    blk = ReadBlock(ws, lastRow)
    Dim cId As Long, cStatus As Long, cFailed As Long
    cId = modUtil.FindHeaderCol(blk, "case_id")
    cStatus = modUtil.FindHeaderCol(blk, "status")
    cFailed = modUtil.FindHeaderCol(blk, "failed_step")
    If cId <= 0 Or cStatus <= 0 Then Exit Function

    Dim rowNo As Long
    rowNo = RowOfCase(blk, lastRow, cId, caseId)
    If rowNo <= 0 Then Exit Function

    Dim failedText As String
    failedText = vbNullString
    If cFailed > 0 Then failedText = Trim$(CStr(blk(rowNo, cFailed)))

    PutNum ws, blk, rowNo, "last_ok_step", keepStep
    PutText ws, blk, rowNo, "status", _
            RepairedStatus(Trim$(CStr(blk(rowNo, cStatus))), failedText, keepStep)
    PutText ws, blk, rowNo, "updated_at", modUtil.NowStamp()
    ApplyRepairedState = True
    Exit Function

Failed:
    ApplyRepairedState = False
End Function

' 列名で位置を引いて1セル書く。列が無ければ何もしない。書込は必ず SetCellSafe。
Private Sub PutText(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                    ByVal colName As String, ByVal textValue As String)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(r, c), textValue, ws.Name & "/" & colName
End Sub

' 数値列(seq / round_no / last_ok_step)。内部生成のLongで数値型を保つ必要があり
' SetCellSafe(文字列を返す)は通さない(modLog.PutNum と同じ理由)。
Private Sub PutNum(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                   ByVal colName As String, ByVal numValue As Long)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    ws.Cells(r, c).Value = numValue   ' SAFE:const 内部生成のLong(外部由来テキストではない)
End Sub

' セル値を Long へ。空・非数値・取得失敗は0。
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
