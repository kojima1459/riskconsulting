Attribute VB_Name = "modUISparring"
Option Explicit

' ============================================================================
' modUISparring - 商談の予行演習シートの描画・発話入出力・「受信箱へ」送信(ui層・T-34)
' (旧称: 壁打ち。利用者向け表示は Z-48 で「商談の予行演習」へ改名。内部識別子は不変)
' ----------------------------------------------------------------------------
' 11章 商談の予行演習ワイヤーと 13章§2.17 が正。実行制御(履歴管理・CallChat呼出・受信箱
' 登録)は app層 modSparring が持ち、本モジュールは画面だけを受け持つ。
'
'   [予行演習を開始/再開] -> modSparring.ResumeSparring(履歴件数と注入文言を返す)
'   [送信]              -> modSparring.SendSparring(PII走査は app層の責務)
'   [受信箱へ]          -> modSparring.SendToInbox(戻り値の inbox_id を行へ書く)
'
' 保存形式(14章§6): 1発話=1行= `seq <TAB> spoke_at <TAB> 本文`、行区切りは vbLf、
'   本文は EscapeJsonStr で畳まれている。seq は user と ai が共有する通し連番
'   なので、両方を読んで seq 昇順に1本の表へ組み直す(13章§2.17「物理行は seq 昇順
'   の追記型」)。11章の「新しい順」表示は最終行へスクロールして実現する。
'
' 絵文字(11章§5・§8.6): `.bas` に絵文字リテラルを書かない。**W9.2 で
'   サロゲートペア組立(ChrW(&HD83D&) & ChrW(&HDCA1&))も撤去した**。禁止の趣旨は
'   「ソースに絵文字を持ち込まない」であって「別の書き方で持ち込んでよい」では
'   ない。ChrW にサロゲートの片割れを渡す形は環境によって扱いが割れ(Mac の実Excel
'   では起動直後に実行時エラー5が出た)、ボタン名は絵文字が無くても意味が通る。
' ============================================================================

Private Const US2_SRC As String = "modUISparring"
Private Const US2_SHEET As String = "商談の予行演習"
Private Const US2_BLOCK As String = "sparring_log"
Private Const US2_SCAN_COLS As Long = 12
Private Const US2_ROOM As Long = 400          ' 1ブロックだけのシートなので広く取る
Private Const US2_MAX_ITEMS As Long = 800

' ============================================================================
' 図形ボタン(11章 商談の予行演習ワイヤー)
' ============================================================================
Public Sub EnsureSparringButtons()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(US2_SHEET)
    If ws Is Nothing Then Exit Sub

    modUISheet.EnsureButton ws, "btn_sp_resume", "予行演習を開始/再開", _
                            1, 8, 160#, "modUISparring.SparringResume"
    modUISheet.EnsureButton ws, "btn_sp_send", "送信", 2, 8, 72#, _
                            "modUISparring.SparringSend"
    modUISheet.EnsureButton ws, "btn_sp_inbox", "受信箱へ", 3, 8, 120#, _
                            "modUISparring.SparringToInbox"
    modUISheet.EnsureButton ws, "btn_sp_s3", "S3を開く", 4, 8, 92#, _
                            "modUISparring.SparringOpenS3"
End Sub

' ============================================================================
' OpenSparring の撤去(W15 Round3・裁定書41 §1)
' ----------------------------------------------------------------------------
' 唯一の呼出元は v3.2 で廃止した HOMEシートのハンドラ modUIHome2.HomeOpenSparring
' であり、それを落としたので本Subも一緒に落とした。現行の入口は
'   使い方タブ⑦上級の[表示する](modUIGuide.ShowAdvanced1 -> 商談の予行演習シートを可視化)
'   -> 商談の予行演習シートの[予行演習を開始/再開](SparringResume)
' で、対象案件は CaseIdOnSheet() が sp_case_id -> hm_case_id の順で解決するため
' 動作は変わらない。
' ============================================================================

' ============================================================================
' [予行演習を開始/再開](11章)。dossier_tier の t3_sparring 昇格は 14章§6 の
'   ResumeSparring が「書込口が無い」ため行わない契約(同節の未解決事項)。
' ============================================================================
Public Sub SparringResume()
    If Not modUIProgress.TryEnterUiLock("商談の予行演習の開始/再開") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    StartOrResume CaseIdOnSheet()

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Private Sub StartOrResume(ByVal caseId As String)
    On Error Resume Next

    If Not modCaseStore.IsValidCaseId(caseId) Then
        Notice "対象案件が選ばれていません（商談の予行演習シートの対象案件IDをご確認ください）。"
        Exit Sub
    End If

    Dim contextNote As String
    Dim turns As Long
    turns = modSparring.ResumeSparring(caseId, contextNote)
    If turns < 0 Then
        Notice "案件一覧からこの案件を読めませんでした。商談の予行演習を開始できません。"
        Exit Sub
    End If

    modUISheet.WriteNamed "sp_context_note", contextNote
    DrawLog caseId
End Sub

' ============================================================================
' [送信](11章)。PII走査・履歴保存・上限系の案内は modSparring が持つ。
' ============================================================================
Public Sub SparringSend()
    If Not modUIProgress.TryEnterUiLock("商談の予行演習の送信") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus

    Dim caseId As String
    caseId = CaseIdOnSheet()
    If Not modCaseStore.IsValidCaseId(caseId) Then
        Notice "対象案件が選ばれていません。"
        GoTo Done
    End If

    Dim utterance As String
    utterance = modUISheet.ReadNamed("sp_input")
    If LenB(Trim$(utterance)) = 0 Then
        Notice "発話入力欄が空です。"
        GoTo Done
    End If

    modUIProgress.SetStage "商談の予行演習（1往復）を実行中", modConfig.GetLong("llm_wait_sec", 1200)

    Dim replyText As String
    Dim errCode As String
    If modSparring.SendSparring(caseId, utterance, replyText, errCode) Then
        modUISheet.WriteNamed "sp_input", vbNullString
        DrawLog caseId
    Else
        Notice SendErrorText(errCode)
    End If

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' 帯域外で返るエラーコードを利用者の言葉にする(16章 E-05/E-15/E-44)。
Private Function SendErrorText(ByVal errCode As String) As String
    Select Case errCode
    Case "E0103"
        SendErrorText = "個人情報らしき記述を検知したため送信しませんでした。発話を直してください。"
    Case "E0204"
        SendErrorText = "本日のAI利用枠の上限に達した可能性があります。" & _
                        "往復数を減らして再開してください。"
    Case "E0101"
        SendErrorText = "商談の予行演習を始める前提が足りません（案件の選択と発話をご確認ください）。"
    Case "E0604"
        SendErrorText = "履歴を保存できませんでした。err_log をご確認ください。"
    Case Else
        SendErrorText = "商談の予行演習の応答を受け取れませんでした（" & errCode & "）。"
    End Select
End Function

' ============================================================================
' [受信箱へ](11章。各発話に付くボタン)。二重送信の抑止は inbox_id 列(13章§2.17)。
' ============================================================================
Public Sub SparringToInbox()
    If Not modUIProgress.TryEnterUiLock("受信箱へ送信") Then Exit Sub
    On Error GoTo Done

    Dim ws As Object
    Set ws = modUISheet.BlockSheet(US2_BLOCK)
    If ws Is Nothing Then GoTo Done

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(US2_BLOCK)
    If headerRow <= 0 Then GoTo Done

    ' 裁定書10補遺 P1: 選択行の取得は ParkFocus より必ず先に行う。
    Dim rowNo As Long
    rowNo = modUISheet.SelectedRow(ws)

    modUIProgress.ParkFocus

    If rowNo <= headerRow Then
        Notice "送りたい発話の行を選んでから押してください。"
        GoTo Done
    End If

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, 1, US2_SCAN_COLS)
    If IsEmpty(hdr) Then GoTo Done

    Dim idCol As Long
    idCol = modUISheet.ColOf(hdr, 1, "inbox_id")

    If idCol > 0 Then
        If LenB(Trim$(modUISheet.CellText(ws, rowNo, idCol))) > 0 Then
            Notice "この発話は送信済みです（inbox_id が入っています）。"
            GoTo Done
        End If
    End If

    Dim caseId As String
    caseId = CaseIdOnSheet()

    Dim seqNo As Long
    Dim roleKind As String
    seqNo = CLng(Val(CellOf(ws, hdr, rowNo, "seq")))
    roleKind = modUICase.EnumEn("sparring_role", CellOf(ws, hdr, rowNo, "role"))
    If seqNo <= 0 Or LenB(roleKind) = 0 Then
        Notice "この行からは発話を特定できませんでした。"
        GoTo Done
    End If

    Dim inboxId As String
    inboxId = modSparring.SendToInbox(caseId, roleKind, seqNo)
    If LenB(inboxId) = 0 Then
        Notice "受信箱へ登録できませんでした（個人情報の検知・前提不足の可能性があります）。"
        GoTo Done
    End If

    If idCol > 0 Then
        modUISheet.PutText ws, rowNo, idCol, inboxId, US2_SHEET & "/inbox_id"
    End If
    modUIInbox.RefreshInbox
    modUIHome.RefreshHome
    Notice "受信箱へ送りました: " & inboxId

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' [S3を開く](11章「結論はS2/S3シートに自分で反映」)。
Public Sub SparringOpenS3()
    If Not modUIProgress.TryEnterUiLock("S3を開く") Then Exit Sub
    On Error GoTo Done
    modUISheet.ShowSheet modUICase2.SheetNameOf(3)
Done:
    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' 履歴の描画(13章§2.17 sparring_log)
' ============================================================================
Private Sub DrawLog(ByVal caseId As String)
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.BlockSheet(US2_BLOCK)
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(US2_BLOCK)
    If headerRow <= 0 Then Exit Sub

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, 1, US2_SCAN_COLS)
    If IsEmpty(hdr) Then Exit Sub

    Dim cSeq As Long
    Dim cRole As Long
    Dim cUtter As Long
    Dim cAt As Long
    cSeq = modUISheet.ColOf(hdr, 1, "seq")
    cRole = modUISheet.ColOf(hdr, 1, "role")
    cUtter = modUISheet.ColOf(hdr, 1, "utterance")
    cAt = modUISheet.ColOf(hdr, 1, "spoke_at")
    If cSeq <= 0 Or cRole <= 0 Or cUtter <= 0 Or cAt <= 0 Then Exit Sub

    ' 案件を切り替えたら全面再描画する(前案件の行を残さない。13章§2.17)。
    modUISheet.ClearBlock ws, headerRow, 1, 5, US2_ROOM

    Dim rows1() As String
    Dim cnt As Long
    ReDim rows1(0 To US2_MAX_ITEMS)
    cnt = 0
    CollectRows modSparring.HistoryOf(caseId, "user"), "user", rows1, cnt
    CollectRows modSparring.HistoryOf(caseId, "ai"), "ai", rows1, cnt
    If cnt = 0 Then Exit Sub

    SortBySeq rows1, cnt

    Dim i As Long
    Dim seqText As String
    Dim atText As String
    Dim roleKind As String
    Dim bodyText As String
    For i = 0 To cnt - 1
        If i >= US2_ROOM Then Exit For
        SplitRow rows1(i), seqText, atText, roleKind, bodyText
        modUISheet.PutText ws, headerRow + i + 1, cSeq, seqText, US2_SHEET & "/seq"
        modUISheet.PutText ws, headerRow + i + 1, cRole, _
                           modUICase.EnumJa("sparring_role", roleKind), US2_SHEET & "/role"
        modUISheet.PutText ws, headerRow + i + 1, cUtter, bodyText, US2_SHEET & "/utterance"
        modUISheet.PutText ws, headerRow + i + 1, cAt, atText, US2_SHEET & "/spoke_at"
    Next i

    ' 11章の「新しい順」表示は最終行へスクロールして実現する。
    ScrollTo ws, headerRow + cnt
End Sub

' 保存形式(`seq <TAB> spoke_at <TAB> 本文`)を `seq <TAB> spoke_at <TAB> role <TAB>
'   本文` の内部形式へ移し替える。本文のエスケープはここで解く。
Private Sub CollectRows(ByVal storedText As String, ByVal roleKind As String, _
                        ByRef rows1() As String, ByRef cnt As Long)
    If LenB(storedText) = 0 Then Exit Sub

    Dim lines() As String
    lines = Split(storedText, vbLf)

    Dim i As Long
    Dim p1 As Long
    Dim p2 As Long
    For i = LBound(lines) To UBound(lines)
        If cnt > US2_MAX_ITEMS Then Exit Sub
        If LenB(Trim$(lines(i))) > 0 Then
            p1 = InStr(1, lines(i), vbTab, vbBinaryCompare)
            If p1 > 1 Then p2 = InStr(p1 + 1, lines(i), vbTab, vbBinaryCompare)
            If p1 > 1 And p2 > p1 Then
                rows1(cnt) = Left$(lines(i), p1 - 1) & vbTab & _
                             Mid$(lines(i), p1 + 1, p2 - p1 - 1) & vbTab & _
                             roleKind & vbTab & _
                             modJsonLite.UnescapeJsonStr(Mid$(lines(i), p2 + 1))
                cnt = cnt + 1
            End If
        End If
    Next i
End Sub

Private Sub SplitRow(ByVal rowText As String, ByRef seqText As String, _
                     ByRef atText As String, ByRef roleKind As String, _
                     ByRef bodyText As String)
    seqText = vbNullString
    atText = vbNullString
    roleKind = vbNullString
    bodyText = vbNullString

    Dim flds() As String
    flds = Split(rowText, vbTab)
    If UBound(flds) - LBound(flds) < 3 Then Exit Sub
    seqText = flds(LBound(flds))
    atText = flds(LBound(flds) + 1)
    roleKind = flds(LBound(flds) + 2)
    bodyText = flds(LBound(flds) + 3)
End Sub

' seq(先頭フィールド)の昇順へ並べ替える(件数が小さいので挿入ソート)。
Private Sub SortBySeq(ByRef rows1() As String, ByVal cnt As Long)
    Dim i As Long
    Dim j As Long
    Dim keyText As String
    Dim keyNo As Long

    For i = 1 To cnt - 1
        keyText = rows1(i)
        keyNo = SeqOf(keyText)
        j = i - 1
        Do While j >= 0
            If SeqOf(rows1(j)) <= keyNo Then Exit Do
            rows1(j + 1) = rows1(j)
            j = j - 1
        Loop
        rows1(j + 1) = keyText
    Next i
End Sub

Private Function SeqOf(ByVal rowText As String) As Long
    Dim p1 As Long
    p1 = InStr(1, rowText, vbTab, vbBinaryCompare)
    If p1 <= 1 Then Exit Function
    SeqOf = CLng(Val(Left$(rowText, p1 - 1)))
End Function

Private Sub ScrollTo(ByVal ws As Object, ByVal rowNo As Long)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ActiveWindow.ScrollRow = rowNo
End Sub

' ============================================================================
' 小物
' ============================================================================
' 商談の予行演習シートの対象案件ID。空ならHOMEの選択案件へ落とす。
Private Function CaseIdOnSheet() As String
    Dim v As String
    v = modUISheet.ReadNamed("sp_case_id")
    If LenB(v) = 0 Then v = modUISheet.ReadNamed("hm_case_id")
    CaseIdOnSheet = v
End Function

Private Function CellOf(ByVal ws As Object, ByVal hdr As Variant, ByVal rowNo As Long, _
                        ByVal colName As String) As String
    Dim colNo As Long
    colNo = modUISheet.ColOf(hdr, 1, colName)
    If colNo <= 0 Then Exit Function
    CellOf = modUISheet.CellText(ws, rowNo, colNo)
End Function

Private Sub Notice(ByVal messageText As String)
    On Error Resume Next
    modUISheet.WriteNamed "hm_warning", messageText
    MsgBox messageText, vbInformation, "商談の予行演習"
End Sub
