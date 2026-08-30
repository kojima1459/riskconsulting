Attribute VB_Name = "modJudgeStore"
Option Explicit

' ============================================================================
' modJudgeStore - 判断台帳シートの唯一の入出力口(app層・T-26・裁定書8 B-8)
' ----------------------------------------------------------------------------
' 正: 13章§1(採番 J-YYYYMM-NNN)/13章§2.7(判断台帳の列)/19章§3(decision enum)/
'   14章§6(NewJudgement の宣言)/16章E-05(4)(situation・key_reasonの保存前
'   PII走査)/10章FR-22(UW判断の事実記録)。
' 責務: UW判断の起票(judge_id採番)・再表示のための読取・事後結果(result/
'   post_loss)の記録。**判断台帳1枚しか触らない**。案件一覧との紐付けは
'   case_ref(自由記述の参照)のみで、案件の実在検査はしない(13章§2.7)。
' 削除は無い: 13章§2.14「受信箱・判断台帳・run_logは削除しない(累積が資産)」
'   のとおり、本モジュールは Delete に相当する公開関数を持たない。
' R4(12章§2但し書き): Excelトークン許可モジュールの1つ(許可の幅は「判断台帳
'   1枚」)。セル書込は全て modUtilText.SetCellSafe(16章 NFR-S7①)を通す。
' 保存前PII走査(16章E-05(4)): situation / key_reason の2列を対象に、書込の
'   直前で modPii.HasPii を通す。**検知したら1列も書かない**(E-05(1)-(4)は
'   ブロック仕様。customer_quoteのような伏字差し替えの例外はここには無い)。
'   E0103 を err_log へ記録し、detail は modPii.ScanReport の返り値(検知種別
'   と文字位置のみ。本文は含めない=NFR-S3)をそのまま使う。
' line_id の種目マスタ実在(13章§2.7の注記)はUI層の入力規則(ドロップダウン)
'   に委ねる(modCaseStore.NewCase が industry_code の実在をチェックしないの
'   と同じ扱い。案件一覧・受信箱と同じく「非空であること」までが本モジュール
'   の責務)。ここを機械的に検査するには種目マスタ(ナレッジブック)への参照が
'   要るが、裁定書8 B-8はその結線を求めていない。
' situation の「200字以内」(13章§2.7)は入力補助の目安であり、本モジュールは
'   打切らない(業務文書の途中欠落を防ぐ。案件入力のHP<200字警告と同じく
'   UI側の文字数カウンタの役目)。セル上限(32,767字)の保護は SetCellSafe が
'   別途担う。
' 列アクセスは列名ベース(13章冒頭。列番号のハードコード禁止)。例外は投げず、
'   読めない・書けないは False / "" で返す。
' ============================================================================

' --- シート名(13章§2.7) ---
Private Const JG_SHEET As String = "判断台帳"

' xlUp の数値(組込定数名を書かず LO の構文チェックで未定義名にしない)。
Private Const JG_DIR_UP As Long = -4162

' 見出し行を読む幅(判断台帳11列に余裕を見た探索範囲。列番号ではない)。
Private Const JG_SCAN_COLS As Long = 16

' 同月連番の上限(受信箱・案件と同じ3桁採番)。
Private Const JG_SERIAL_MAX As Long = 999

' enum の区切り(modUtil.AppendIdList と同じ ";")。
Private Const JG_SEP As String = ";"

' 19章§3 / 13章§2.7 の enum。
Private Const JG_DECISIONS As String = "raise;close;restrict;keep;improve"
Private Const JG_RESULTS As String = "won;lost;pending"

' --- 純ロジック(Excel非依存。シート操作側は必ずここを通す。14章§6) ---

' BuildJudgeId - 判断台帳IDの組立(13章§1: J-YYYYMM-NNN)。monthText は yyyymm
'   の6桁ちょうど(数字のみ)、seq は 1..999。範囲外は ""(呼び出し側が枯渇・
'   不正として扱う。例外は投げない)。modInboxStore.BuildInboxId と同型。
Public Function BuildJudgeId(ByVal monthText As String, ByVal seq As Long) As String
    If Len(monthText) <> 6 Then Exit Function
    If Not IsAllDigits(monthText) Then Exit Function
    If seq < 1 Or seq > JG_SERIAL_MAX Then Exit Function
    BuildJudgeId = "J-" & monthText & "-" & Right$("00" & CStr(seq), 3)
End Function

' IsValidJudgeId - "J-"+数字6桁+"-"+数字3桁(001..999)ちょうどか。前後空白は
'   許さない。19章§4の注記どおり「判断基準ID(J-NN)」とは形が異なるため
'   桁数で区別する。
Public Function IsValidJudgeId(ByVal id As String) As Boolean
    If Len(id) <> 12 Then Exit Function
    If Left$(id, 2) <> "J-" Then Exit Function
    If Mid$(id, 9, 1) <> "-" Then Exit Function
    If Not IsAllDigits(Mid$(id, 3, 6)) Then Exit Function
    Dim tailText As String
    tailText = Right$(id, 3)
    If Not IsAllDigits(tailText) Then Exit Function
    IsValidJudgeId = (CLng(tailText) >= 1)
End Function

' IsValidDecision - 19章§3の decision enum(raise/close/restrict/keep/improve)
'   に一致するか。空は False(decisionは必須列)。
Public Function IsValidDecision(ByVal decisionText As String) As Boolean
    IsValidDecision = IsListedValue(JG_DECISIONS, decisionText)
End Function

' IsValidJudgeResult - 13章§2.7の result enum(won/lost/pending)に一致するか。
'   空は False。result列そのものは任意入力なので、空を許すかどうかの判断は
'   呼び出し側(NewJudgement/SetJudgementResult)が「空なら検査をスキップする」
'   形で行う。
Public Function IsValidJudgeResult(ByVal resultText As String) As Boolean
    IsValidJudgeResult = IsListedValue(JG_RESULTS, resultText)
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

' ";" 区切りの enum リストに値が含まれるか(前後空白は無視・大小文字は区別)。
Private Function IsListedValue(ByVal listText As String, ByVal valueText As String) As Boolean
    Dim v As String
    v = Trim$(valueText)
    If LenB(v) = 0 Then Exit Function
    IsListedValue = (InStr(1, JG_SEP & listText & JG_SEP, JG_SEP & v & JG_SEP, vbBinaryCompare) > 0)
End Function

' --- 判断台帳シートI/O ---

' NewJudgement - UW判断を1件起票して judge_id を返す(13章§1・§2.7・14章§6)。
'   採番は J-YYYYMM-NNN。当月の使用済み最大連番の次から始め、衝突したら
'   E0605 を記録して次の番号へ(999で枯渇)。失敗時は ""(例外は投げない)。
'   必須列(judge_id/judged_at以外): line_id・situation・decision・
'   key_reason・recorded_by。空が1つでもあれば1列も書かない。
'   decision は19章§3のenumのみ受け付ける(不一致は保存ブロック)。result は
'   任意だが、非空なら13章§2.7のenum(won/lost/pending)のみ受け付ける。
'   保存直前に situation・key_reason を modPii へ通す(16章E-05(4))。検知
'   したら1列も書かず E0103 を記録する。
Public Function NewJudgement(ByVal rec As TJudgement) As String
    On Error GoTo Failed

    If LenB(Trim$(rec.line_id)) = 0 Or LenB(Trim$(rec.situation)) = 0 _
            Or LenB(Trim$(rec.decision)) = 0 Or LenB(Trim$(rec.key_reason)) = 0 _
            Or LenB(Trim$(rec.recorded_by)) = 0 Then
        modLog.LogError "E0101", "modJudgeStore.NewJudgement", _
                        "empty_required:line_id/situation/decision/key_reason/recorded_by"
        Exit Function
    End If

    If Not IsValidDecision(rec.decision) Then
        modLog.LogError "E0101", "modJudgeStore.NewJudgement", "invalid_decision:" & Trim$(rec.decision)
        Exit Function
    End If

    Dim resultText As String
    resultText = Trim$(rec.result)
    If LenB(resultText) > 0 Then
        If Not IsValidJudgeResult(resultText) Then
            modLog.LogError "E0101", "modJudgeStore.NewJudgement", "invalid_result:" & resultText
            Exit Function
        End If
    End If

    ' 16章E-05(4): 保存前PII走査。situation/key_reasonのどちらかで検知したら
    ' 1列も書かない(ブロック。伏字差し替えの余地は無い)。
    If modPii.HasPii(rec.situation) Then
        modLog.LogError "E0103", "modJudgeStore.NewJudgement", _
                        modPii.ScanReport(rec.situation, JG_SHEET & "/situation")
        Exit Function
    End If
    If modPii.HasPii(rec.key_reason) Then
        modLog.LogError "E0103", "modJudgeStore.NewJudgement", _
                        modPii.ScanReport(rec.key_reason, JG_SHEET & "/key_reason")
        Exit Function
    End If

    Dim ws As Object
    Set ws = JgSheet()
    If ws Is Nothing Then
        modLog.LogError "E0603", "modJudgeStore.NewJudgement", "sheet_missing:judge"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = JgLastRow(ws)
    Dim blk As Variant
    blk = JgBlock(ws, lastRow)
    Dim cId As Long
    cId = modUtil.FindHeaderCol(blk, "judge_id")
    If cId <= 0 Then
        modLog.LogError "E0603", "modJudgeStore.NewJudgement", "header_missing:judge_id"
        Exit Function
    End If

    Dim monthText As String
    monthText = Left$(modUtilText.IsoDateCompact(Date), 6)

    ' 当月ぶんの使用済み最大連番。
    Dim usedMax As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim sn As Long
        sn = SerialOfJudgeId(Trim$(CStr(blk(r, cId))), monthText)
        If sn > usedMax Then usedMax = sn
    Next r

    Dim serialNo As Long
    Dim newId As String
    serialNo = usedMax
    Do
        serialNo = serialNo + 1
        If serialNo > JG_SERIAL_MAX Then
            modLog.LogError "E0605", "modJudgeStore.NewJudgement", "serial_exhausted:" & monthText
            Exit Function
        End If
        newId = BuildJudgeId(monthText, serialNo)
        If LenB(newId) = 0 Then
            modLog.LogError "E0605", "modJudgeStore.NewJudgement", "id_build_failed:" & monthText
            Exit Function
        End If
        If JgRowOf(blk, lastRow, cId, newId) <= 0 Then Exit Do
        modLog.LogError "E0605", "modJudgeStore.NewJudgement", "id_collision:" & CStr(serialNo)
    Loop

    Dim wr As Long
    wr = lastRow + 1
    JgPut ws, blk, wr, "judge_id", newId
    JgPut ws, blk, wr, "judged_at", modUtil.NowStamp()
    JgPut ws, blk, wr, "line_id", Trim$(rec.line_id)
    JgPut ws, blk, wr, "case_ref", rec.case_ref
    JgPut ws, blk, wr, "situation", rec.situation
    JgPut ws, blk, wr, "decision", Trim$(rec.decision)
    JgPut ws, blk, wr, "factor_note", rec.factor_note
    JgPut ws, blk, wr, "key_reason", rec.key_reason
    JgPut ws, blk, wr, "result", resultText
    JgPut ws, blk, wr, "post_loss", rec.post_loss
    JgPut ws, blk, wr, "recorded_by", Trim$(rec.recorded_by)

    NewJudgement = newId
    Exit Function

Failed:
    modLog.LogError "E0603", "modJudgeStore.NewJudgement", "unexpected", Err.Number
    NewJudgement = vbNullString
End Function

' ReadJudgement - 1件の判断を読む(入力→保存→再表示一致のための唯一の口。
'   17章T-26 DoD)。戻り値 False=シート・見出し・当該行が無い(呼び出し側は
'   fail-closedで中止する)。False のとき rec は全列空へ戻す。
Public Function ReadJudgement(ByVal judgeId As String, ByRef rec As TJudgement) As Boolean
    On Error GoTo Failed

    Dim blank As TJudgement
    rec = blank
    If Not IsValidJudgeId(judgeId) Then Exit Function

    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not JgLocate(judgeId, ws, blk, rowNo) Then Exit Function

    rec.line_id = JgValue(blk, rowNo, "line_id")
    rec.case_ref = JgValue(blk, rowNo, "case_ref")
    rec.situation = JgValue(blk, rowNo, "situation")
    rec.decision = JgValue(blk, rowNo, "decision")
    rec.factor_note = JgValue(blk, rowNo, "factor_note")
    rec.key_reason = JgValue(blk, rowNo, "key_reason")
    rec.result = JgValue(blk, rowNo, "result")
    rec.post_loss = JgValue(blk, rowNo, "post_loss")
    rec.recorded_by = JgValue(blk, rowNo, "recorded_by")
    ReadJudgement = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modJudgeStore.ReadJudgement", "read_failed", Err.Number
    ReadJudgement = False
End Function

' SetJudgementResult - 事後結果(result/post_loss)を記録する(13章§2.7・
'   FR-22)。situation/decision/key_reason 等の記録内容はここでは動かさない
'   (起票時に確定させた事実記録を後から書き換えない。動くのは結果の2列のみ)。
'   resultText が非空なら13章§2.7のenum(won/lost/pending)のみ受け付ける
'   (不一致は保存ブロック)。空文字列を渡すと result 列を空へ戻す(pendingの
'   取消)。result/post_loss は自由記述ではないPII対象外の2列(16章E-05(4)の
'   走査対象は situation/key_reason のみ)なのでmodPiiは通さない。
Public Function SetJudgementResult(ByVal judgeId As String, ByVal resultText As String, _
                                   ByVal postLoss As String) As Boolean
    On Error GoTo Failed

    Dim tgt As String
    tgt = Trim$(resultText)
    If LenB(tgt) > 0 Then
        If Not IsValidJudgeResult(tgt) Then
            modLog.LogError "E0101", "modJudgeStore.SetJudgementResult", "invalid_result:" & tgt
            Exit Function
        End If
    End If

    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not JgLocate(judgeId, ws, blk, rowNo) Then Exit Function

    JgPut ws, blk, rowNo, "result", tgt
    JgPut ws, blk, rowNo, "post_loss", postLoss
    SetJudgementResult = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modJudgeStore.SetJudgementResult", "write_failed", Err.Number
    SetJudgementResult = False
End Function

' --- 内部ヘルパー ---

' 判断台帳IDから当月ぶんの連番を取り出す(月不一致・形違いは0)。IDの形の判定は
' IsValidJudgeId が唯一の値源(2箇所で書かない)。
Private Function SerialOfJudgeId(ByVal judgeId As String, ByVal monthText As String) As Long
    If Not IsValidJudgeId(judgeId) Then Exit Function
    If Mid$(judgeId, 3, 6) <> monthText Then Exit Function
    SerialOfJudgeId = CLng(Right$(judgeId, 3))
End Function

' 名前でシートを取る。無ければ Nothing(実行時に生やすと列定義の欠けた表になる)。
Private Function JgSheet() As Object
    On Error GoTo NoSheet
    Set JgSheet = ThisWorkbook.Worksheets(JG_SHEET)
    Exit Function
NoSheet:
    Set JgSheet = Nothing
End Function

' A列基準の最終行。データが無ければ1(見出し行)。
Private Function JgLastRow(ByVal ws As Object) As Long
    Dim n As Long
    On Error GoTo One1
    n = ws.Cells(ws.Rows.count, 1).End(JG_DIR_UP).row
    If n < 1 Then n = 1
    JgLastRow = n
    Exit Function
One1:
    JgLastRow = 1
End Function

' 見出し行を含む矩形を一度だけ読む(1セルだけの Range は2次元配列にならない)。
Private Function JgBlock(ByVal ws As Object, ByVal lastRow As Long) As Variant
    On Error GoTo Empty1
    Dim hi As Long
    hi = lastRow
    If hi < 2 Then hi = 2
    JgBlock = ws.Range(ws.Cells(1, 1), ws.Cells(hi, JG_SCAN_COLS)).Value
    Exit Function
Empty1:
    JgBlock = Empty
End Function

' 読み込み済みブロックから judge_id 一致行を探す(見つからなければ0)。
Private Function JgRowOf(ByVal blk As Variant, ByVal lastRow As Long, _
                         ByVal idCol As Long, ByVal judgeId As String) As Long
    On Error GoTo NotFound
    Dim n As Long
    For n = 2 To lastRow
        If Trim$(CStr(blk(n, idCol))) = judgeId Then
            JgRowOf = n
            Exit Function
        End If
    Next n
    Exit Function
NotFound:
    JgRowOf = 0
End Function

' 「判断台帳の1行を触る」ための場所決め(シート -> 最終行 -> 矩形 -> 該当行)。
'   どこかで欠けたら False(呼び出し側は黙って中止する。ここでログは書かない)。
Private Function JgLocate(ByVal judgeId As String, ByRef ws As Object, _
                          ByRef blk As Variant, ByRef rowNo As Long) As Boolean
    rowNo = 0
    Set ws = JgSheet()
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = JgLastRow(ws)
    blk = JgBlock(ws, lastRow)

    Dim cId As Long
    cId = modUtil.FindHeaderCol(blk, "judge_id")
    If cId <= 0 Then Exit Function

    rowNo = JgRowOf(blk, lastRow, cId, Trim$(judgeId))
    JgLocate = (rowNo > 0)
End Function

' 列名で1セルを引く(13章冒頭。列が無ければ "")。
Private Function JgValue(ByVal blk As Variant, ByVal r As Long, _
                         ByVal headerName As String) As String
    On Error GoTo Empty0
    Dim c As Long
    c = modUtil.FindHeaderCol(blk, headerName)
    If c <= 0 Then Exit Function
    JgValue = Trim$(CStr(blk(r, c)))
    Exit Function
Empty0:
    JgValue = vbNullString
End Function

' 列名で位置を引いて1セル書く。列が無ければ何もしない。書込は必ず SetCellSafe
' (16章 NFR-S7①)。
Private Sub JgPut(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                  ByVal colName As String, ByVal textValue As String)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(r, c), textValue, JG_SHEET & "/" & colName
End Sub
