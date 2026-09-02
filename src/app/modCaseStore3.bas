Attribute VB_Name = "modCaseStore3"
Option Explicit

' ============================================================================
' modCaseStore3 - data_key の一覧と、ラウンド確定時の種目絞り込み(app層・T-56)
' ----------------------------------------------------------------------------
' なぜ3本目なのか(30,000字契約。12章§2・17章§3の分割バックログ2番):
'   modCaseStore が 29,997字(残り3字)で満杯であり、v2.6(裁定書25 S3/S4)の
'   3点 -- data_key への `input_finance` 追加・案件一覧の `adopted_story_nos` /
'   `focus_line_ids` の写し込み -- を本体へ1文字も足せない。そこで
'     modCaseStore  = 「何を書くか」(採番・参照優先・状態遷移・分割保存の規約)
'     modCaseStore2 = 「どこへどう書くか」(素のシートI/O)
'     modCaseStore3 = 「data_key の一覧」と「ラウンド確定の付帯処理」
'   の3本に割った。modCaseStore.SaveData の data_key 検査は DataKeys() を、
'   modCaseStore.FreezeRound は ApplyRoundFocus() を呼ぶ(依存は
'   modCaseStore -> modCaseStore3 の一方向)。14章§6の公開契約面には載せない
'   (modCaseStore2 / modCompanyFile2 と同じ扱い)。
'
' 13章§2.1(v2.6・裁定書25 S4)の写し込みの筋:
'   商談の記録(フィードバック)の `used_proposals`(採用した提案の番号・「;」区切り)
'   -> 案件一覧 `adopted_story_nos` -> 採用storyの `line_ids` -> `focus_line_ids`。
'   **利用者に種目を選ばせない**。`focus_line_ids` は種目マスタに実在するIDだけを
'   書く(実在しないIDが 15章§1.2c BLOCK_ROUND2_FOCUS の {{focus_line_ids}} へ
'   流れると、S2/S3が存在しない種目を深掘りする)。
'
' R4(12章§4): 案件一覧・フィードバックの読み書きが責務。許可の幅は modCaseStore
'   と同じ本体ブック内であり広がっていない。セル書込は modCaseStore2.PutText
'   (= modUtilText.SetCellSafe)に集約する。
' ============================================================================

Private Const CS3_SHEET_CASES As String = "案件一覧"
Private Const CS3_SHEET_FB As String = "フィードバック"
Private Const CS3_SEP As String = ";"
Private Const CS3_SRC As String = "modCaseStore3"

' 13章§2.2 の data_key(**全29値**。19章§3と完全一致させる)。v2.6 で
'   `input_finance`(決算・財務。裁定書25 S3)を input_coverage_note の次へ足した。
'   並びは13章§2.2 の列挙順そのままで、modCaseStore.SaveData の許可リストになる。
Private Const CS3_DATA_KEYS As String = "input_hp;input_yuho;input_memo;input_contract;input_prev_renewal;input_dossier;input_field_notes;input_coverage_note;input_finance;input_hearing_answers;s1_json;s2_json;s3_json;s4_json;s2c_json;s3c_json;s2r_json;s3r_json;s2_prev_json;s1_edited;s2_edited;s3_edited;s4_edited;s1_json_failed;s2_json_failed;s3_json_failed;s4_json_failed;sparring_u;sparring_a"

' DataKeys - 13章§2.2 の data_key 一覧(「;」区切り)。値源はここ1箇所。
Public Function DataKeys() As String
    DataKeys = CS3_DATA_KEYS
End Function

' ============================================================================
' 純関数(層(a)テスト対象。Excelもナレッジも触らない)
' ============================================================================

' NormalizeStoryNos - `used_proposals` を `adopted_story_nos` の形へ整える。
'   ・区切りは「;」。前後の空白は落とす
'   ・10進整数として読めない要素は捨てる(利用者の走り書きを案件一覧へ入れない)
'   ・重複は先に出たものだけ残す。入力の並びは変えない
'   ・1つも残らなければ空文字("採用の記録なし")
Public Function NormalizeStoryNos(ByVal usedProposals As String) As String
    Dim parts() As String
    parts = Split(Replace(usedProposals, vbTab, CS3_SEP), CS3_SEP)

    Dim acc As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim one As String
        one = Trim$(parts(i))
        If IsPositiveInt(one) Then
            If Not IsListed(acc, one) Then
                If LenB(acc) > 0 Then acc = acc & CS3_SEP
                acc = acc & one
            End If
        End If
    Next i
    NormalizeStoryNos = acc
End Function

' CollectLineIds - 採用storyの `line_ids` を集める(13章§2.1 focus_line_ids の導出)。
'   s3Json = S3の解決済みJSON(stories[] を含む)。storyNos = `1;3` の形。
'   ・stories[] を配列順に走査し、story_no が storyNos に載っている行だけを採る
'   ・その行の line_ids(「; 」区切り)を分解して「;」で連結する
'   ・重複は先に出たものだけ残す(story 1 と 3 が同じ種目を指しても1本)
'   ・storyNos が空なら空文字(絞り込みなし)
'   **実在チェックはここではしない**(ナレッジを触らない純関数に保つため)。
'   実在で絞るのは ExistingLineIds。
Public Function CollectLineIds(ByVal s3Json As String, ByVal storyNos As String) As String
    If LenB(Trim$(storyNos)) = 0 Then Exit Function
    If LenB(Trim$(s3Json)) = 0 Then Exit Function

    Dim acc As String
    Dim storyVal As Variant
    For Each storyVal In modJsonLite.GetArrayItems(s3Json, "stories")
        Dim storyJson As String
        storyJson = CStr(storyVal)
        If LenB(storyJson) > 0 Then
            Dim noText As String
            noText = CStr(modJsonLite.GetLong(storyJson, "story_no", 0))
            If IsListed(storyNos, noText) Then
                Dim idVal As Variant
                For Each idVal In modJsonLite.GetArrayItems(storyJson, "line_ids")
                    AppendId acc, Trim$(CStr(idVal))
                Next idVal
            End If
        End If
    Next storyVal
    CollectLineIds = acc
End Function

' ExistingLineIds - 種目マスタに実在するIDだけを残す(13章§2.1)。
'   実在しないIDは**書かない**(黙って落とす。空になっても案件一覧へは空を書く)。
Public Function ExistingLineIds(ByVal lineIds As String) As String
    Dim parts() As String
    parts = Split(lineIds, CS3_SEP)

    Dim acc As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim one As String
        one = Trim$(parts(i))
        If LenB(one) > 0 Then
            If modKnowledge.LineIdExists(one) Then
                If Not IsListed(acc, one) Then
                    If LenB(acc) > 0 Then acc = acc & CS3_SEP
                    acc = acc & one
                End If
            End If
        End If
    Next i
    ExistingLineIds = acc
End Function

' ============================================================================
' ApplyRoundFocus - ラウンド確定の付帯処理(13章§2.1・§2.5。modCaseStore.FreezeRound
'   から呼ばれる)。商談の記録の used_proposals を案件一覧へ写し、採用storyの
'   line_ids から focus_line_ids を導出して書く。
'   採用の記録が無ければ2列とも空にする(前ラウンドの絞り込みを残さない)。
' ============================================================================
Public Sub ApplyRoundFocus(ByVal caseId As String, ByVal s3Json As String)
    On Error GoTo Failed

    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not modCaseStore2.LocateRow(CS3_SHEET_CASES, caseId, ws, blk, rowNo) Then Exit Sub

    Dim adopted As String
    adopted = NormalizeStoryNos(UsedProposalsOf(caseId))

    Dim focus As String
    focus = ExistingLineIds(CollectLineIds(s3Json, adopted))

    modCaseStore2.PutText ws, blk, rowNo, "adopted_story_nos", adopted
    modCaseStore2.PutText ws, blk, rowNo, "focus_line_ids", focus
    Exit Sub

Failed:
    modLog.LogError "E0603", CS3_SRC & ".ApplyRoundFocus", "focus_failed", Err.Number
End Sub

' UsedProposalsOf - フィードバック(13章§2.5)から当該案件の used_proposals を読む。
'   同じ案件に複数行あるときは**最後に書かれた非空の行**を採る(記録は追記式で、
'   後の商談ほど新しい採用の事実であるため)。
Private Function UsedProposalsOf(ByVal caseId As String) As String
    Dim ws As Object
    Set ws = modCaseStore2.SheetOf(CS3_SHEET_FB)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = modCaseStore2.LastRowOf(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = modCaseStore2.ReadBlock(ws, lastRow)

    Dim cId As Long, cUsed As Long
    cId = modUtil.FindHeaderCol(blk, "case_id")
    cUsed = modUtil.FindHeaderCol(blk, "used_proposals")
    If cId <= 0 Or cUsed <= 0 Then Exit Function

    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(blk(r, cId))), Trim$(caseId), vbTextCompare) = 0 Then
            Dim v As String
            v = Trim$(CStr(blk(r, cUsed)))
            If LenB(v) > 0 Then UsedProposalsOf = v
        End If
    Next r
End Function

' ============================================================================
' 内部ヘルパー(純関数)
' ============================================================================

' 「;」区切りの一覧に値が(要素として)載っているか。
Private Function IsListed(ByVal listText As String, ByVal valueText As String) As Boolean
    If LenB(valueText) = 0 Then Exit Function
    IsListed = (InStr(1, CS3_SEP & listText & CS3_SEP, _
                      CS3_SEP & valueText & CS3_SEP, vbBinaryCompare) > 0)
End Function

' 10進の正整数か(前後の空白は呼出側で落としてある前提)。
Private Function IsPositiveInt(ByVal s As String) As Boolean
    If LenB(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        Dim c As String
        c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    IsPositiveInt = (Val(s) > 0)
End Function

' 1本のIDを acc(「;」区切り)へ足す。空と重複は捨てる。
Private Sub AppendId(ByRef acc As String, ByVal idText As String)
    If LenB(idText) = 0 Then Exit Sub
    If IsListed(acc, idText) Then Exit Sub
    If LenB(acc) > 0 Then acc = acc & CS3_SEP
    acc = acc & idText
End Sub
