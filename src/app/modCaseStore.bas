Attribute VB_Name = "modCaseStore"
Option Explicit

' ============================================================================
' modCaseStore - 案件一覧 / case_data の唯一の入出力口(app層・T-20)
' 責務(12章§2・13章§2.1/§2.2・14章§6): NewCase(13章§1・E-23) / SaveData・
'   LoadData(分割保存と透過結合。E-22) / ResolveStepJson(参照優先。13章§2.2) /
'   SetStatus / SetStepOutcome(E-06の書込口) / InvalidateDownstream(E-10) /
'   RepairStates(E-12) / FreezeRound。
' 配置: app層(12章§2)。R4: 案件一覧/case_data 2枚のI/Oが責務そのもので12章
'   §4の許可枠の1つ。再描画と確認ダイアログは ui層 modUICase。
' 下位I/O 10本は modCaseStore2 へ分離(裁定書8 A-2。modCompanyFile2 と同型)。
'   素のExcelトークンは残らないが責務は不変のため12章§4の許可枠から外さない。
' 列アクセスは列名ベース(13章冒頭)。セル書込は全て modUtilText.SetCellSafe
'   (16章 NFR-S7①)。数値列だけは内部生成のLongなので ' SAFE:const を明示。
' 純ロジック(採番・参照優先・状態遷移・無効化対象)はシートI/O無しの関数へ分離。
'   BuildCaseId / IsValidCaseId / CanTransition / ResolveDataKey の4本は14章
'   §6が公開を宣言(裁定書6 項目7。シートI/Oに閉じると層(a)から検査できない)。
' ============================================================================

' --- シート名(13章§2.1・§2.2) ---
Private Const CS_SHEET_CASES As String = "案件一覧"
Private Const CS_SHEET_DATA As String = "case_data"

' 1セルに入れる最大字数(16章 E-22。物理上限32,767字の手前で切る)。
Private Const CS_CHUNK_CHARS As Long = 32000

' 2相書込(裁定書9 B13)の仮seq帯の起点。超過seq=書き途中の仮行(LoadData対象外)。
Private Const CS_SEQ_BAND As Long = 100000

' dossier_tier の enum(13章§2.1・19章§3)。
Private Const CS_TIERS As String = "t1_quick;t2_full;t3_sparring"

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

' failed_step の enum(13章§2.1)。空は「失敗なし」なのでこの表には含めない。
Private Const CS_FAILED_STEPS As String = "s1;s2;s3;s4;s2c;s3c"

' status のうち s4_done より上位(RepairStates で降格させない2値)。
Private Const CS_STATUSES_ABOVE_S4 As String = "exported;feedback_done"

' data_key の enum 全29値の一覧は modCaseStore3.DataKeys() が持つ(13章§2.2・
' 19章§3と完全一致)。表に無いキーは拒否。本モジュールが30,000字契約の上限に
' 達したため v2.6 で分割先へ移した(値も並びも1つも変えていない)。

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
'   空のまま起こす(13章§2.1=実行前検証で埋まればよい。裁定書6 項目9)。
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
    Set ws = modCaseStore2.SheetOf(CS_SHEET_CASES)
    If ws Is Nothing Then
        modLog.LogError "E0603", "modCaseStore.NewCase", "sheet_missing:cases"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = modCaseStore2.LastRowOf(ws)
    Dim blk As Variant
    blk = modCaseStore2.ReadBlock(ws, lastRow)
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
        If modCaseStore2.RowOfCase(blk, lastRow, cId, newId) <= 0 Then Exit Do
        ' 既存と衝突(E-23)。記録して次の番号へ。
        modLog.LogError "E0605", "modCaseStore.NewCase", "id_collision:" & CStr(serialNo)
    Loop

    Dim wr As Long
    wr = lastRow + 1
    Dim stampText As String
    stampText = modUtil.NowStamp()

    modCaseStore2.PutText ws, blk, wr, "case_id", newId
    modCaseStore2.PutText ws, blk, wr, "case_type", Trim$(caseType)
    modCaseStore2.PutText ws, blk, wr, "dossier_tier", CS_TIER_DEFAULT
    modCaseStore2.PutText ws, blk, wr, "company", company
    modCaseStore2.PutText ws, blk, wr, "industry_code", industryCode
    modCaseStore2.PutText ws, blk, wr, "status", CS_STATUS_NEW
    modCaseStore2.PutText ws, blk, wr, "created_at", stampText
    modCaseStore2.PutText ws, blk, wr, "updated_at", stampText
    modCaseStore2.PutText ws, blk, wr, "owner", modCaseStore2.OwnerName()
    modCaseStore2.PutText ws, blk, wr, "s4_variant", CS_VARIANT_DEFAULT
    modCaseStore2.PutNum ws, blk, wr, "round_no", 1
    modCaseStore2.PutNum ws, blk, wr, "last_ok_step", 0

    NewCase = newId
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.NewCase", "unexpected", Err.Number
    NewCase = vbNullString
End Function

' SaveData - case_data へ1本のテキストを分割保存する(13章§2.2・16章 E-22)。
'   32,000字ごとの断片 seq=1.. で置換。content が空なら削除だけで True。
'   **2相書込**(裁定書9 B13): 新断片を仮seq帯へ書き切ってから旧行を消す(実体
'   は modCaseStore2.ReplaceSeqRows。途中失敗で旧行は消えず False)。断片は
'   SetCellSafe を通る("'" 前置は .Value 読み戻しで元へ戻る)。
Public Function SaveData(ByVal caseId As String, ByVal dataKey As String, _
                         ByVal content As String) As Boolean
    On Error GoTo Failed

    If LenB(Trim$(caseId)) = 0 Then
        modLog.LogError "E0101", "modCaseStore.SaveData", "empty_case_id"
        Exit Function
    End If
    If Not IsListedValue(modCaseStore3.DataKeys(), dataKey) Then
        modLog.LogError "E0101", "modCaseStore.SaveData", "invalid_data_key:" & dataKey
        Exit Function
    End If

    Dim ws As Object
    Set ws = modCaseStore2.SheetOf(CS_SHEET_DATA)
    If ws Is Nothing Then
        modLog.LogError "E0603", "modCaseStore.SaveData", "sheet_missing:case_data"
        Exit Function
    End If

    Dim idText As String, keyText As String
    idText = Trim$(caseId)
    keyText = Trim$(dataKey)

    If LenB(content) = 0 Then
        ' 空=削除の意図(仮行も含め全行を消す)。
        modCaseStore2.DropRowsOf ws, idText, keyText
        SaveData = True
        Exit Function
    End If

    Dim parts() As String
    parts = modUtil.SplitForCells(content, CS_CHUNK_CHARS)

    If Not modCaseStore2.ReplaceSeqRows(ws, idText, keyText, parts, CS_SEQ_BAND) Then GoTo Failed

    SaveData = True
    ' 17章 Z-51。成否確定後=記録失敗で保存を失敗にしない(詳細は3側)。
    modCaseStore3.LogEditRatioOnSave idText, keyText, content
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.SaveData", "write_failed:" & dataKey, Err.Number
    SaveData = False
End Function

' LoadData - 断片を seq 順に結合して返す(透過処理)。該当が無ければ ""。
'   仮seq帯(CS_SEQ_BAND超=2相書込の書き途中)は読まない。完全性検査(裁定書9
'   B13): seq 1..maxSeq に欠番があれば詰めずに E0604 を記録して ""(fail-closed)。
'   裁定書10 M6: 本seq帯が空でも仮seq帯に行が残っていれば E0604 を記録して ""。
Public Function LoadData(ByVal caseId As String, ByVal dataKey As String) As String
    On Error GoTo Failed

    If LenB(Trim$(caseId)) = 0 Then Exit Function

    Dim ws As Object
    Set ws = modCaseStore2.SheetOf(CS_SHEET_DATA)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = modCaseStore2.LastRowOf(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = modCaseStore2.ReadBlock(ws, lastRow)
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

    ' 1周目: seq の最大値=断片数を求める(仮seq帯は対象外)。
    Dim maxSeq As Long
    Dim bandLeft As Boolean
    Dim r As Long
    For r = 2 To lastRow
        If modCaseStore2.MatchesRow(blk, r, cCase, cKey, wantCase, wantKey) Then
            Dim sq1 As Long
            sq1 = modCaseStore2.ToLongSafe(blk(r, cSeq))
            If sq1 > maxSeq And sq1 <= CS_SEQ_BAND Then maxSeq = sq1
            If sq1 > CS_SEQ_BAND Then bandLeft = True
        End If
    Next r
    If maxSeq < 1 Then
        ' M6: 本帯空+仮帯残存=相2後の失敗痕跡。
        If bandLeft Then
            modLog.LogError "E0604", "modCaseStore.LoadData", _
                            "seq_band_orphan:" & wantKey
        End If
        Exit Function
    End If

    ' 2周目: seq を添字として配置し、行の実在を seen へ記録する。
    Dim parts() As String
    Dim seen() As Boolean
    ReDim parts(1 To maxSeq)
    ReDim seen(1 To maxSeq)
    For r = 2 To lastRow
        If modCaseStore2.MatchesRow(blk, r, cCase, cKey, wantCase, wantKey) Then
            Dim sq2 As Long
            sq2 = modCaseStore2.ToLongSafe(blk(r, cSeq))
            If sq2 >= 1 And sq2 <= maxSeq Then
                parts(sq2) = CStr(blk(r, cBody))
                seen(sq2) = True
            End If
        End If
    Next r

    ' 完全性検査(B13): 欠番=途中失敗・手作業破損。黙って詰めない。
    Dim chk As Long
    For chk = 1 To maxSeq
        If Not seen(chk) Then
            modLog.LogError "E0604", "modCaseStore.LoadData", _
                            "seq_gap:" & wantKey & " missing=" & CStr(chk)
            Exit Function
        End If
    Next chk

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
    Dim blk As Variant
    Dim rowNo As Long
    If Not modCaseStore2.LocateRow(CS_SHEET_CASES, caseId, ws, blk, rowNo) Then Exit Function

    modCaseStore2.PutText ws, blk, rowNo, "status", Trim$(statusText)
    modCaseStore2.PutText ws, blk, rowNo, "updated_at", modUtil.NowStamp()
    SetStatus = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.SetStatus", "write_failed", Err.Number
    SetStatus = False
End Function

' PromoteTier - 案件一覧 dossier_tier の**唯一の書込口**(裁定書9 N2・A-1)。
'   tierText は enum(CS_TIERS)のみ。不一致は1列も書かず False(E0101)。行不在も
'   False。status は動かさない。方向(昇格/降格)は判定しない。app層からの呼出は
'   modSparring.ResumeSparring の1点のみ(14章§6)。
Public Function PromoteTier(ByVal caseId As String, ByVal tierText As String) As Boolean
    On Error GoTo Failed

    Dim t As String
    t = Trim$(tierText)
    If Not IsListedValue(CS_TIERS, t) Then
        modLog.LogError "E0101", "modCaseStore.PromoteTier", "invalid_tier:" & t
        Exit Function
    End If

    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not modCaseStore2.LocateRow(CS_SHEET_CASES, caseId, ws, blk, rowNo) Then Exit Function

    modCaseStore2.PutText ws, blk, rowNo, "dossier_tier", t
    modCaseStore2.PutText ws, blk, rowNo, "updated_at", modUtil.NowStamp()
    PromoteTier = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.PromoteTier", "write_failed", Err.Number
    PromoteTier = False
End Function

' SetStepOutcome - 16章 E-06 の last_ok_step / failed_step の【書込口】(裁定書8
'   A-2 で14章§6へ宣言)。modPipeline の成功・失敗経路がここを通る。status は
'   動かさない(遷移の唯一の口は SetStatus)。
'   lastOkStep: 0..4 を書く。**負値は「更新しない」**(E-06「失敗時は更新
'     しない」を引数で表す)。
'   failedStep: "" は失敗の記憶を消す(13章§2.1)。非空は enum s1/s2/s3/s4/
'     s2c/s3c のみ。表に無い値は E0101 で拒否して1列も書かない。
Public Function SetStepOutcome(ByVal caseId As String, ByVal lastOkStep As Long, _
                               ByVal failedStep As String) As Boolean
    On Error GoTo Failed

    Dim failText As String
    failText = Trim$(failedStep)
    If LenB(failText) > 0 Then
        If Not IsListedValue(CS_FAILED_STEPS, failText) Then
            modLog.LogError "E0101", "modCaseStore.SetStepOutcome", _
                            "invalid_failed_step:" & failText
            Exit Function
        End If
    End If
    If lastOkStep > 4 Then
        modLog.LogError "E0101", "modCaseStore.SetStepOutcome", _
                        "invalid_last_ok_step:" & CStr(lastOkStep)
        Exit Function
    End If

    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not modCaseStore2.LocateRow(CS_SHEET_CASES, caseId, ws, blk, rowNo) Then Exit Function

    If lastOkStep >= 0 Then
        modCaseStore2.PutNum ws, blk, rowNo, "last_ok_step", lastOkStep
    End If
    modCaseStore2.PutText ws, blk, rowNo, "failed_step", failText
    modCaseStore2.PutText ws, blk, rowNo, "updated_at", modUtil.NowStamp()
    SetStepOutcome = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.SetStepOutcome", "write_failed", Err.Number
    SetStepOutcome = False
End Function

' SetReportPath - 案件一覧 report_path の【書込口】(18章§1.1⑦・14章§6の
'   GenerateHtmlReport が「確定パスを案件一覧 report_path に記録する」と定める)。
'   modExportHtml はR4によりシートに触れないため、記録はここを通す。status は
'   動かさない(出力の成否は状態遷移に影響しない。16章 E-48)。
'   記録に失敗しても生成済みファイルは残る=呼び出し側は警告に留める。
Public Function SetReportPath(ByVal caseId As String, ByVal pathText As String) As Boolean
    On Error GoTo Failed

    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not modCaseStore2.LocateRow(CS_SHEET_CASES, caseId, ws, blk, rowNo) Then Exit Function

    modCaseStore2.PutText ws, blk, rowNo, "report_path", pathText
    modCaseStore2.PutText ws, blk, rowNo, "updated_at", modUtil.NowStamp()
    SetReportPath = True
    modCaseStore3.LogEditRatiosOnReport caseId
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.SetReportPath", "write_failed", Err.Number
    SetReportPath = False
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
    Set ws = modCaseStore2.SheetOf(CS_SHEET_DATA)
    If ws Is Nothing Then Exit Sub

    Dim keyArr() As String
    keyArr = Split(keyList, CS_SEP)

    Dim removed As String
    Dim i As Long
    For i = LBound(keyArr) To UBound(keyArr)
        If modCaseStore2.DropRowsOf(ws, Trim$(caseId), keyArr(i)) > 0 Then
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
'   (1) sN_json(合格済のみ)の最大Stepと last_ok_step の小さいほうを採る
'   (2) failed_step 非空なら status=error 維持・failed_step は書き換えない
'   (3) 導出は sN_json だけ (4) ui_lock は modUIProgress のモジュール変数で
'       自然消滅(E-11/E-51。R1により何もしない) (5) 修復件数を返す。
Public Function RepairStates() As Long
    On Error GoTo Failed

    Dim ws As Object
    Set ws = modCaseStore2.SheetOf(CS_SHEET_CASES)
    If ws Is Nothing Then
        modLog.LogError "E0603", "modCaseStore.RepairStates", "sheet_missing:cases"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = modCaseStore2.LastRowOf(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = modCaseStore2.ReadBlock(ws, lastRow)
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
            storedOk = modCaseStore2.ToLongSafe(blk(r, cLastOk))
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
                modCaseStore2.PutNum ws, blk, r, "last_ok_step", effStep
                didFix = True
            End If
            If curStatus <> wantStatus Then
                modCaseStore2.PutText ws, blk, r, "status", wantStatus
                didFix = True
            End If
            If didFix Then
                modCaseStore2.PutText ws, blk, r, "updated_at", modUtil.NowStamp()
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
    Set ws = modCaseStore2.SheetOf(CS_SHEET_CASES)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = modCaseStore2.LastRowOf(ws)
    Dim blk As Variant
    blk = modCaseStore2.ReadBlock(ws, lastRow)
    Dim cId As Long, cRound As Long
    cId = modUtil.FindHeaderCol(blk, "case_id")
    cRound = modUtil.FindHeaderCol(blk, "round_no")
    If cId <= 0 Or cRound <= 0 Then
        modLog.LogError "E0603", "modCaseStore.FreezeRound", "header_missing:round_no"
        Exit Function
    End If

    Dim rowNo As Long
    rowNo = modCaseStore2.RowOfCase(blk, lastRow, cId, Trim$(caseId))
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
    curRound = modCaseStore2.ToLongSafe(blk(rowNo, cRound))
    If curRound < 1 Then curRound = 1

    modCaseStore2.PutNum ws, blk, rowNo, "round_no", curRound + 1
    modCaseStore2.PutText ws, blk, rowNo, "updated_at", modUtil.NowStamp()
    ' 13章§2.1(v2.6・裁定書25 S4): 採用提案番号と深掘り種目を写す。
    modCaseStore3.ApplyRoundFocus caseId, ResolveStepJson(caseId, 3)
    FreezeRound = curRound + 1
    Exit Function

Failed:
    modLog.LogError "E0603", "modCaseStore.FreezeRound", "freeze_failed", Err.Number
    FreezeRound = 0
End Function

' --- 内部ヘルパー(状態導出。素のシートI/Oは modCaseStore2 側) ---

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
'   裁定書9 B20: 冒頭で現在値との min を取り、下げる方向にしか動かさない。
Private Function ApplyRepairedState(ByVal caseId As String, ByVal keepStep As Long) As Boolean
    On Error GoTo Failed
    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not modCaseStore2.LocateRow(CS_SHEET_CASES, caseId, ws, blk, rowNo) Then Exit Function

    Dim cStatus As Long, cFailed As Long, cLastOk As Long
    cStatus = modUtil.FindHeaderCol(blk, "status")
    cFailed = modUtil.FindHeaderCol(blk, "failed_step")
    cLastOk = modUtil.FindHeaderCol(blk, "last_ok_step")
    If cStatus <= 0 Then Exit Function

    ' B20: min(keepStep, 現在の last_ok_step)。
    If cLastOk > 0 Then
        Dim curOk As Long
        curOk = modCaseStore2.ToLongSafe(blk(rowNo, cLastOk))
        If keepStep > curOk Then keepStep = curOk
    End If
    If keepStep < 0 Then keepStep = 0

    Dim failedText As String
    failedText = vbNullString
    If cFailed > 0 Then failedText = Trim$(CStr(blk(rowNo, cFailed)))

    modCaseStore2.PutNum ws, blk, rowNo, "last_ok_step", keepStep
    modCaseStore2.PutText ws, blk, rowNo, "status", _
            RepairedStatus(Trim$(CStr(blk(rowNo, cStatus))), failedText, keepStep)
    modCaseStore2.PutText ws, blk, rowNo, "updated_at", modUtil.NowStamp()
    ApplyRepairedState = True
    Exit Function

Failed:
    ApplyRepairedState = False
End Function

