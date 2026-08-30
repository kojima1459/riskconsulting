Attribute VB_Name = "modUICase4"
Option Explicit

' ============================================================================
' modUICase4 - フィードバック・判断台帳(ui層・T-31)
' ----------------------------------------------------------------------------
' 30,000字契約(12章§2)による modUICase の分割先。12章§2 が modUICase の責務と
' 定める「FB・判断台帳の描画・読取」をそのまま引き受ける。
'
' どちらのシートも 13章§2.5 / §2.7 の**単票テーブル**であり、enum列は機械値の
' まま持つ(19章§3の日本語ラベルへは変換しない=tools/enum_check.py の EXCLUDED に
' 理由つきで宣言済み)。したがって本モジュールは変換表を引かない。
'
' 起票の唯一の口:
'   判断台帳は modJudgeStore.NewJudgement(採番・enum検証・16章E-05(4)のPII走査を
'   内蔵)が唯一の口。画面は「下書き行」を読み、**その行を消してから**起票する
'   (同じ内容の行が2本できない)。
'   フィードバックは 13章§2.5 のとおり事実のみの表で、store系モジュールを持た
'   ないため本モジュールが行を仕上げる(case_id / recorded_by / recorded_at)。
'   16章 E-05(5): customer_quote だけはブロックせず伏字案を提示して選ばせる。
' ============================================================================

Private Const U4_SRC As String = "modUICase4"
Private Const U4_FB As String = "フィードバック"
Private Const U4_JUDGE As String = "判断台帳"
Private Const U4_SCAN_COLS As Long = 32
Private Const U4_CASEIN As String = "案件入力"

' ============================================================================
' 図形ボタン(11章 HOMEワイヤーの[商談の記録][判断台帳]から開いた先の操作)
' ============================================================================
Public Sub EnsureRecordButtons()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(U4_FB)
    modUISheet.EnsureButton ws, "btn_fb_save", "商談の記録を保存", 1, 11, 128#, _
                            "modUICase4.FeedbackSave"

    Set ws = modUISheet.SheetOf(U4_JUDGE)
    modUISheet.EnsureButton ws, "btn_jl_save", "判断を起票", 1, 13, 100#, _
                            "modUICase4.JudgeSave"
    modUISheet.EnsureButton ws, "btn_jl_result", "結果を記録", 1, 14, 100#, _
                            "modUICase4.JudgeSaveResult"
End Sub

' ============================================================================
' フィードバック(13章§2.5・16章 E-05(5))
' ============================================================================
Public Sub FeedbackSave()
    If Not modUIProgress.TryEnterUiLock("商談の記録") Then Exit Sub
    On Error GoTo Done

    Dim ws As Object
    Set ws = modUISheet.SheetOf(U4_FB)
    If ws Is Nothing Then GoTo Done

    ' 裁定書10補遺 P1: ParkFocus は活性シートの Cells(1,1) を選択するため、先に
    ' 通すと利用者の行選択が必ず行1へ潰れて以後の処理が成立しない。選択行の
    ' 取得を先に済ませ、フォーカス退避はその後(および完了時)に行う。
    Dim rowNo As Long
    rowNo = modUISheet.SelectedRow(ws)

    modUIProgress.ParkFocus

    If rowNo < 2 Then
        Notice "記録する行を選んでから押してください（1行目は見出しです）。"
        GoTo Done
    End If

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, 1, 1, U4_SCAN_COLS)
    If IsEmpty(hdr) Then GoTo Done

    Dim quoteCol As Long
    quoteCol = modUtil.FindHeaderCol(hdr, "customer_quote")

    ' 16章 E-05(5): customer_quote は**ブロックせず**伏字案を提示して選ばせる。
    If quoteCol > 0 Then
        Dim quoteText As String
        quoteText = modUISheet.CellText(ws, rowNo, quoteCol)
        If LenB(quoteText) > 0 Then
            If modPii.HasPii(quoteText) Then
                modLog.LogError "E0103", U4_SRC & ".FeedbackSave", _
                                modPii.ScanReport(quoteText, "customer_quote")
                Dim answer As Long
                answer = MsgBox("顧客の発言に個人情報らしき記述があります。" & vbLf & _
                                "伏字に置き換えて保存しますか？（いいえ＝原文のまま保存）", _
                                vbYesNo + vbExclamation, "個人情報の確認")
                If answer = vbYes Then
                    modUISheet.PutText ws, rowNo, quoteCol, modPii.MaskText(quoteText), _
                                       U4_FB & "/customer_quote"
                End If
            End If
        End If
    End If

    PutIfEmpty ws, hdr, rowNo, "case_id", modUISheet.ReadNamed("hm_case_id")
    PutIfEmpty ws, hdr, rowNo, "recorded_by", OwnerName()
    PutAlways ws, hdr, rowNo, "recorded_at", modUtil.NowStamp()

    ' 13章§2.1・14章§6(裁定書9 B16(b)): フィードバック保存の**成功分岐**が
    ' case_status の feedback_done を立てる唯一の点。案件IDが取れないときは
    ' 状態を動かさない(記録そのものは成功として扱う。16章 E-48)。
    Dim caseId As String
    caseId = modUISheet.ReadNamed("hm_case_id")
    If modCaseStore.IsValidCaseId(caseId) Then
        modCaseStore.SetStatus caseId, "feedback_done"
    End If

    modLog.LogUsage "feedback_saved", caseId, "row=" & CStr(rowNo)
    Notice "商談の記録を保存しました。"

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Private Sub PutIfEmpty(ByVal ws As Object, ByVal hdr As Variant, ByVal rowNo As Long, _
                       ByVal colName As String, ByVal valueText As String)
    Dim colNo As Long
    colNo = modUtil.FindHeaderCol(hdr, colName)
    If colNo <= 0 Then Exit Sub
    If LenB(Trim$(modUISheet.CellText(ws, rowNo, colNo))) > 0 Then Exit Sub
    modUISheet.PutText ws, rowNo, colNo, valueText, ws.Name & "/" & colName
End Sub

Private Sub PutAlways(ByVal ws As Object, ByVal hdr As Variant, ByVal rowNo As Long, _
                      ByVal colName As String, ByVal valueText As String)
    Dim colNo As Long
    colNo = modUtil.FindHeaderCol(hdr, colName)
    If colNo <= 0 Then Exit Sub
    modUISheet.PutText ws, rowNo, colNo, valueText, ws.Name & "/" & colName
End Sub

Private Function OwnerName() As String
    On Error GoTo NoName
    OwnerName = CStr(Application.UserName)
    Exit Function
NoName:
    OwnerName = vbNullString
End Function

' ============================================================================
' 判断台帳(13章§2.7・14章§6 modJudgeStore)
' ----------------------------------------------------------------------------
' 起票は modJudgeStore.NewJudgement が唯一の口(採番・enum検証・PII走査を持つ)。
' 画面は最終行の次の空行を「下書き行」として使い、[判断を起票]でその行を読み、
' **下書き行を消してから**正式に起票する(同じ内容の行が2本できない)。
' ============================================================================
Public Sub JudgeSave()
    If Not modUIProgress.TryEnterUiLock("判断を起票") Then Exit Sub
    On Error GoTo Done

    Dim ws As Object
    Set ws = modUISheet.SheetOf(U4_JUDGE)
    If ws Is Nothing Then GoTo Done

    ' 裁定書10補遺 P1: 選択行の取得は ParkFocus より必ず先に行う。
    Dim rowNo As Long
    rowNo = modUISheet.SelectedRow(ws)

    modUIProgress.ParkFocus

    If rowNo < 2 Then
        Notice "起票する下書き行を選んでから押してください（1行目は見出しです）。"
        GoTo Done
    End If

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, 1, 1, U4_SCAN_COLS)
    If IsEmpty(hdr) Then GoTo Done

    If LenB(Trim$(ColText(ws, hdr, rowNo, "judge_id"))) > 0 Then
        Notice "この行は起票済みです（judge_id が入っています）。"
        GoTo Done
    End If

    Dim rec As TJudgement
    rec.line_id = ColText(ws, hdr, rowNo, "line_id")
    rec.case_ref = ColText(ws, hdr, rowNo, "case_ref")
    rec.situation = ColText(ws, hdr, rowNo, "situation")
    rec.decision = ColText(ws, hdr, rowNo, "decision")
    rec.factor_note = ColText(ws, hdr, rowNo, "factor_note")
    rec.key_reason = ColText(ws, hdr, rowNo, "key_reason")
    rec.result = ColText(ws, hdr, rowNo, "result")
    rec.post_loss = ColText(ws, hdr, rowNo, "post_loss")
    rec.recorded_by = OwnerName()

    ' 裁定書9 B3: **起票が成功したときだけ**下書き行を消す。NewJudgement は必須列の
    ' 欠落・enum外・PII検知など7通りで空を返し、VBAの行削除は Undo できないため、
    ' 先に消すと利用者の入力が復元不能のまま失われる(13章§4「判断台帳は削除しない」)。
    Dim judgeId As String
    judgeId = modJudgeStore.NewJudgement(rec)
    If LenB(judgeId) = 0 Then
        Notice "起票できませんでした。種目・状況・判断・決め手が埋まっているか、" & _
               "個人情報が含まれていないかをご確認ください（下書き行はそのまま残しています）。"
        GoTo Done
    End If

    ' 起票済みの行が2本にならないよう、成功を確かめてから下書き行を消す。
    ' 裁定書10補遺 P2: NewJudgement は judge_id 列基準の最終行の**次**へ確定行を書く
    ' ため、下書き行が台帳末尾直下にあると確定行と同じ位置になる。削除前に当該行の
    ' judge_id が空であることを確かめ、空でなければ(=その行が確定行へ昇格した)
    ' 削除しない。
    If LenB(Trim$(ColText(ws, hdr, rowNo, "judge_id"))) = 0 Then
        ws.Rows(rowNo).Delete
    Else
        modLog.LogError "E0603", U4_SRC & ".JudgeSave", "draft_row_promoted:" & judgeId
    End If

    Notice "判断台帳へ起票しました: " & judgeId

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' FR-22: 事後結果(result / post_loss)だけを更新する。
Public Sub JudgeSaveResult()
    If Not modUIProgress.TryEnterUiLock("結果を記録") Then Exit Sub
    On Error GoTo Done

    Dim ws As Object
    Set ws = modUISheet.SheetOf(U4_JUDGE)
    If ws Is Nothing Then GoTo Done

    Dim rowNo As Long
    rowNo = modUISheet.SelectedRow(ws)
    If rowNo < 2 Then
        Notice "結果を記録する行を選んでから押してください。"
        GoTo Done
    End If

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, 1, 1, U4_SCAN_COLS)
    If IsEmpty(hdr) Then GoTo Done

    Dim judgeId As String
    judgeId = ColText(ws, hdr, rowNo, "judge_id")
    If Not modJudgeStore.IsValidJudgeId(judgeId) Then
        Notice "judge_id のある行を選んでください。"
        GoTo Done
    End If

    If modJudgeStore.SetJudgementResult(judgeId, ColText(ws, hdr, rowNo, "result"), _
                                        ColText(ws, hdr, rowNo, "post_loss")) Then
        Notice "結果を記録しました: " & judgeId
    Else
        Notice "結果を記録できませんでした（result は won / lost / pending のみです）。"
    End If

Done:
    modUIProgress.ExitUiLock
End Sub

Private Function ColText(ByVal ws As Object, ByVal hdr As Variant, ByVal rowNo As Long, _
                         ByVal colName As String) As String
    Dim colNo As Long
    colNo = modUtil.FindHeaderCol(hdr, colName)
    If colNo <= 0 Then Exit Function
    ColText = modUISheet.CellText(ws, rowNo, colNo)
End Function

' ============================================================================
' 追加収集の[コピー](17章 T-31 DoD)
' ----------------------------------------------------------------------------
' 裁定書11 Q1: 30,000字契約(12章§2)により modUICase3 から本モジュールへ移した
' (案件入力の overflow 永続ガードを入れる余白が modUICase3 に無かった)。
' 呼出は図形ボタンの OnAction "modUICase4.CopyResearchRow" のみ(14章§6)。
' ----------------------------------------------------------------------------
' クリックされた図形の位置から行を決める(OnActionは引数を運べないため、
' Application.Caller が返す図形名で当該行を特定する)。
' ============================================================================
Public Sub CopyResearchRow()
    If Not modUIProgress.TryEnterUiLock("調査プロンプトのコピー") Then Exit Sub
    On Error GoTo Done

    Dim ws As Object
    Set ws = ActiveSheet
    If ws Is Nothing Then GoTo Done

    Dim shapeKey As String
    shapeKey = CStr(Application.Caller)
    If LenB(shapeKey) = 0 Then GoTo Done

    Dim rowNo As Long
    rowNo = ws.Shapes(shapeKey).TopLeftCell.row
    If rowNo <= 0 Then GoTo Done

    Dim colNo As Long
    colNo = PromptColOn(ws)
    If colNo <= 0 Then GoTo Done

    Dim payload As String
    payload = modUISheet.CellText(ws, rowNo, colNo)
    If LenB(payload) = 0 Then GoTo Done

    If Not modUISheet.CopyToClipboard(payload) Then
        ws.Cells(rowNo, colNo).Select
        Notice "クリップボードへ入れられませんでした。選択したセルを Ctrl+C でコピーしてください。"
    End If

Done:
    modUIProgress.ExitUiLock
End Sub

' 調査プロンプト本文の列。案件入力は ci_research_anchor の1つ右、
' S1は s1_research_requests ブロックの prompt_text 列。
Private Function PromptColOn(ByVal ws As Object) As Long
    On Error GoTo NoCol

    If ws.Name = U4_CASEIN Then
        Dim anchor As Object
        Set anchor = modUISheet.NamedCell("ci_research_anchor")
        If anchor Is Nothing Then Exit Function
        PromptColOn = anchor.Column + 1
        Exit Function
    End If

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow("s1_research_requests")
    If headerRow <= 0 Then Exit Function

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, 1, U4_SCAN_COLS)
    PromptColOn = modUISheet.ColOf(hdr, 1, "prompt_text")
    Exit Function
NoCol:
    PromptColOn = 0
End Function

' 利用者への案内。HOMEの警告欄へ書き、ダイアログでも知らせる。
Private Sub Notice(ByVal messageText As String)
    On Error Resume Next
    modUISheet.WriteNamed "hm_warning", messageText
    MsgBox messageText, vbInformation, "リスク提案ナビ"
End Sub
