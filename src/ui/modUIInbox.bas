Attribute VB_Name = "modUIInbox"
Option Explicit

' ============================================================================
' modUIInbox - 受信箱の描画・投函・診断結果表示・判定入力(ui層・T-32)
' ----------------------------------------------------------------------------
' 11章 受信箱ワイヤーと 13章§2.6 が正。受信箱シートは**データ兼画面**であり
' 一覧そのものが表なので、本モジュールは「一覧を作る」のではなく
'   (1) 投函(＋投函)          -> modInboxStore.NewInboxItem
'   (2) 未診断の一括診断      -> modPlayOps.RunPreflightAll
'   (3) 選択行の診断全文表示  -> pf_json を読み下してラベル図形へ
'   (4) 選択行の判定入力      -> modInboxStore.SetInboxJudgement(E-41 fail-closed)
'   (5) 関心度の表示(FR-17)   -> modInboxStore.InterestText
' を担う。
'
' 11章のワイヤーが持つ「関心度」「詳細ペイン」は 13章§2.6 に列も名前付きレンジも
' 無いため、**ラベル図形**で描く(表のデータ面を勝手に増やさない)。
'
' 受信箱の語彙(status / source_kind / drop_type / revive_tag)は 13章§2.6 が
' **機械値のまま**持つ列なので日本語変換は行わない(19章§3の変換表の対象外。
' tools/enum_check.py の EXCLUDED に理由つきで宣言してある)。
' 例外は判定の**入力列** judge_to で、ここだけは利用者が選ぶ列なので日本語ラベル
' (採択/条件付き保留/却下)で見せ、modUICase.EnumEn("inbox_judge_to", ...) で
' 機械値へ戻してから store へ渡す(裁定書10 M7・19章§3 inbox.judge_to)。
' ============================================================================

Private Const UI2_SRC As String = "modUIInbox"
Private Const UI2_SHEET As String = "受信箱"
Private Const UI2_SCAN_COLS As Long = 24
Private Const UI2_LBL_INTEREST As String = "lbl_inbox_interest"
Private Const UI2_LBL_DIAG As String = "lbl_inbox_diagnosis"

' 受信箱 body の上限(13章§2.6・16章 E-22: 32,000字で打切り＋警告)。
Private Const UI2_BODY_MAX As Long = 32000

' ============================================================================
' 図形ボタン(11章 受信箱ワイヤー: [＋投函] [未診断を一括診断])
' ============================================================================
Public Sub EnsureInboxButtons()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UI2_SHEET)
    If ws Is Nothing Then Exit Sub

    modUISheet.EnsureButton ws, "btn_ib_post", "＋投函", 1, 18, 72#, _
                            "modUIInbox.InboxPost"
    modUISheet.EnsureButton ws, "btn_ib_diag", "未診断を一括診断", 1, 19, 120#, _
                            "modUIInbox.InboxDiagnoseAll"
    modUISheet.EnsureButton ws, "btn_ib_show", "診断を表示", 1, 20, 92#, _
                            "modUIInbox.InboxShowDiagnosis"
    modUISheet.EnsureButton ws, "btn_ib_judge", "判定を保存", 1, 21, 92#, _
                            "modUIInbox.InboxSaveJudgement"

    modUISheet.EnsureLabel ws, UI2_LBL_INTEREST, "関心度: (集計なし)", 2, 18, 260#, 16#
    modUISheet.EnsureLabel ws, UI2_LBL_DIAG, "診断: (行を選んで［診断を表示］)", _
                           4, 18, 320#, 90#

    ' 裁定書10 C1: 投函下書き行を常設する(起動のたびに冪等に確かめる)。
    EnsureDraftRow ws
End Sub

' ============================================================================
' 投函下書き行(13章§2.6・裁定書10 C1)
' ----------------------------------------------------------------------------
' 受信箱テーブルの**先頭データ行(ヘッダ直下)**を[＋投函]の入力欄として常設し、
' inbox_id 列に固定マーカーを置く。判断台帳の下書き行と同じ作法であり、
' InputBox の255字上限を回避しつつ、名前付きレンジ(旧 ib_body_draft)を作れない
' フラットシートでも成立する(台帳の現機構は columns だけのシートに
' header_fields の名前付きレンジを作れない。builder は変更しない)。
' 戻り値=下書き行の行番号(0=見出しが読めない等で用意できなかった)。
' ============================================================================
Private Function EnsureDraftRow(ByVal ws As Object) As Long
    On Error GoTo Zero0
    If ws Is Nothing Then Exit Function

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, 1, 1, UI2_SCAN_COLS)
    If IsEmpty(hdr) Then Exit Function

    Dim colNo As Long
    colNo = modUtil.FindHeaderCol(hdr, "inbox_id")
    If colNo <= 0 Then Exit Function

    ' 既にマーカー行があればそれが下書き行(行番号には依存しない)。
    Dim lastRow As Long
    lastRow = modUISheet.LastRowOf(ws)
    Dim r As Long
    For r = 2 To lastRow
        If IsDraftRow(ws, colNo, r) Then
            EnsureDraftRow = r
            Exit Function
        End If
    Next r

    ' 裁定書11 Q3(b)・13章§2.6の回復規定: 利用者がマーカーを消すと、
    ' inbox_id が空で body に書きかけの本文が残った**孤児行**になる。新しい
    ' 下書き行を挿すと孤児行は誰からも触られないまま残置するので、先頭から
    ' 探して見つかればその行をマーカーで下書き行として復帰させる。
    ' 裁定書12 V6: 回復条件へ「status列も空」を足す。status に値のある行は
    ' 投函済みのデータ行(id列だけを消してしまった行)であり、下書き行として
    ' マーキングすると次の投函成功で入力列が消され status だけの幽霊行になる。
    Dim bodyCol As Long
    Dim stCol As Long
    bodyCol = modUtil.FindHeaderCol(hdr, "body")
    stCol = modUtil.FindHeaderCol(hdr, "status")
    If bodyCol > 0 Then
        For r = 2 To lastRow
            If LenB(Trim$(modUISheet.CellText(ws, r, colNo))) = 0 And _
               IsBlankCol(ws, stCol, r) Then
                If LenB(Trim$(modUISheet.CellText(ws, r, bodyCol))) > 0 Then
                    modUISheet.PutText ws, r, colNo, modInboxStore.IB_DRAFT_MARK, _
                                       UI2_SHEET & "/inbox_id"
                    modLog.LogError "E0603", UI2_SRC & ".EnsureDraftRow", _
                                    "draft_row_remarked:row=" & CStr(r)
                    EnsureDraftRow = r
                    Exit Function
                End If
            End If
        Next r
    End If

    ' 無ければヘッダ直下へ用意する。既存データが1行目の下に居るときだけ行を
    ' 差し込む(空のシートに空行を増やさない)。
    If LenB(Trim$(modUISheet.CellText(ws, 2, colNo))) > 0 Then
        ws.Rows(2).Insert
        ' 裁定書11 Q3(a): 挿入行は直上=見出し行から書式を継ぐため、利用者が
        ' 選ぶ列(source_kind / judge_to)の入力規則が付かない。受信箱の
        ' フラット表(f種別)の束ねを張り直して下書き行にもDVを行き渡らせる。
        modUICase.RebindFlatValidation UI2_SHEET
    End If
    modUISheet.PutText ws, 2, colNo, modInboxStore.IB_DRAFT_MARK, _
                       UI2_SHEET & "/inbox_id"
    EnsureDraftRow = 2
    Exit Function
Zero0:
    EnsureDraftRow = 0
End Function

' 当該列が空か(列が見つからないときは True=条件に加えない)。
Private Function IsBlankCol(ByVal ws As Object, ByVal colNo As Long, _
                            ByVal rowNo As Long) As Boolean
    If colNo <= 0 Then
        IsBlankCol = True
        Exit Function
    End If
    IsBlankCol = (LenB(Trim$(modUISheet.CellText(ws, rowNo, colNo))) = 0)
End Function

' 当該行が投函下書き行か(13章§2.6。判定基準は id列マーカーの1点)。
Private Function IsDraftRow(ByVal ws As Object, ByVal idCol As Long, _
                            ByVal rowNo As Long) As Boolean
    If idCol <= 0 Then Exit Function
    IsDraftRow = (Trim$(modUISheet.CellText(ws, rowNo, idCol)) = modInboxStore.IB_DRAFT_MARK)
End Function

' ============================================================================
' RefreshInbox - 関心度(FR-17)の表示とHOMEの件数の更新
' ============================================================================
Public Sub RefreshInbox()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UI2_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim summary As String
    summary = modInboxStore.InterestText()
    If LenB(summary) = 0 Then summary = "(集計なし)"
    modUISheet.SetLabelText ws, UI2_LBL_INTEREST, "関心度: " & summary
End Sub

' ============================================================================
' CountByStatus - 13章§2.6 status 列の件数(HOMEの hm_inbox_* が使う)
' ============================================================================
Public Function CountByStatus(ByVal statusText As String) As Long
    On Error GoTo Zero0

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UI2_SHEET)
    If ws Is Nothing Then Exit Function

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, 1, 1, UI2_SCAN_COLS)
    If IsEmpty(hdr) Then Exit Function

    Dim colNo As Long
    colNo = modUtil.FindHeaderCol(hdr, "status")
    If colNo <= 0 Then Exit Function

    ' 裁定書11 Q2(13章§2.6): 件数・集計を行う経路も下書き行を必ずスキップする。
    ' status 列には入力規則もシート保護も無いので、下書き行に値が入りうる。
    ' 数えてしまうと「一括診断は0件なのにHOMEは1件」という食い違いになる。
    Dim idCol As Long
    idCol = modUtil.FindHeaderCol(hdr, "inbox_id")

    Dim lastRow As Long
    lastRow = modUISheet.LastRowOf(ws)

    Dim n As Long
    Dim r As Long
    For r = 2 To lastRow
        If Not IsDraftRow(ws, idCol, r) Then
            If Trim$(modUISheet.CellText(ws, r, colNo)) = statusText Then n = n + 1
        End If
    Next r
    CountByStatus = n
    Exit Function
Zero0:
    CountByStatus = 0
End Function

' ============================================================================
' 投函(11章 [＋投函])。16章 E-05(2): 本文のPII検知は登録をブロックする。
' ----------------------------------------------------------------------------
' 裁定書10 C1(13章§2.6): 本文は InputBox(既定フォントで255字前後が上限)でも
'   名前付きレンジ(旧 N6 ib_body_draft。**廃止**)でもなく、受信箱テーブルの
'   **投函下書き行**(ヘッダ直下・id列マーカー)の body 列から読む。
'   32,000字超は打ち切って警告(16章 E-22)。起票に成功したら下書き行の入力列を
'   空へ戻す(行そのものは消さない)。テーマは下書き行の theme 列に書けるが、
'   短文なので空なら従来どおり InputBox で尋ねる。
' ============================================================================
Public Sub InboxPost()
    If Not modUIProgress.TryEnterUiLock("投函") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UI2_SHEET)
    If ws Is Nothing Then GoTo Done

    Dim draftRow As Long
    draftRow = EnsureDraftRow(ws)
    If draftRow <= 0 Then
        Notice "受信箱シートの見出しを読み取れませんでした。ブックの構成をご確認ください。"
        GoTo Done
    End If

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, 1, 1, UI2_SCAN_COLS)
    If IsEmpty(hdr) Then GoTo Done

    Dim bodyText As String
    bodyText = Trim$(ColText(ws, hdr, draftRow, "body"))
    If LenB(bodyText) = 0 Then
        Notice "受信箱の投函下書き行（id列が「" & modInboxStore.IB_DRAFT_MARK & _
               "」の行）の body 欄に本文を貼ってから押してください。"
        GoTo Done
    End If

    Dim truncated As Boolean
    If Len(bodyText) > UI2_BODY_MAX Then
        bodyText = modUtil.SafeLeft(bodyText, UI2_BODY_MAX)
        truncated = True
    End If

    Dim themeText As String
    themeText = Trim$(ColText(ws, hdr, draftRow, "theme"))
    If LenB(themeText) = 0 Then
        themeText = Trim$(CStr(InputBox("テーマ（短く。関心度の集計はこの文言でまとめます）", _
                                        "受信箱への投函")))
    End If
    If LenB(themeText) = 0 Then GoTo Done

    ' 裁定書11 Q4(19章§3 inbox.source_kind): source_kind は下書き行の**入力列**
    ' なので judge_to と同作法にする。日本語ラベル(部内投稿/現場の声/ウォッチ)で
    ' 選ばせ、modUICase.EnumEn で機械値へ戻す。表に無いラベルは "" が返るので
    ' 推測で寄せずここで止める(fail-closed。store の E0101 で原因不明の案内に
    ' なるのを避ける)。空欄は従来どおり投函の既定値。
    Dim sourceKindJa As String
    Dim sourceKind As String
    sourceKindJa = Trim$(ColText(ws, hdr, draftRow, "source_kind"))
    If LenB(sourceKindJa) = 0 Then
        sourceKind = "member_post"
    Else
        sourceKind = modUICase.EnumEn("inbox_source_kind", sourceKindJa)
        If LenB(sourceKind) = 0 Then
            Notice "発信元の種別は一覧から選んでください（部内投稿／現場の声／ウォッチ）: " & _
                   sourceKindJa
            GoTo Done
        End If
    End If

    If modPii.HasPii(themeText & vbLf & bodyText) Then
        modLog.LogError "E0103", UI2_SRC & ".InboxPost", _
                        modPii.ScanReport(themeText & vbLf & bodyText, "inbox_body")
        Notice "個人情報らしき記述を検知したため投函しませんでした。該当箇所を直してください。"
        GoTo Done
    End If

    Dim inboxId As String
    inboxId = modInboxStore.NewInboxItem(sourceKind, themeText, bodyText)
    If LenB(inboxId) = 0 Then
        Notice "投函できませんでした。err_log をご確認ください。"
        GoTo Done
    End If

    ' 裁定書10 C1: 起票に**成功したときだけ**下書き行の入力列を空へ戻す
    ' (失敗したら利用者の入力を残す。判断台帳の下書き行と同じ作法=B3)。
    ' マーカーと行そのものは常設なので消さない。
    ClearDraftCol ws, hdr, draftRow, "body"
    ClearDraftCol ws, hdr, draftRow, "theme"
    ClearDraftCol ws, hdr, draftRow, "source_kind"

    RefreshInbox
    modUIHome.RefreshHome
    If truncated Then
        Notice "投函しました: " & inboxId & _
               "（本文が32,000字を超えたため、超過分は打ち切りました）"
    Else
        Notice "投函しました: " & inboxId
    End If

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' ============================================================================
' 未診断の一括診断(11章 [未診断を一括診断]・16章 E-40)
' ============================================================================
Public Sub InboxDiagnoseAll()
    If Not modUIProgress.TryEnterUiLock("受信箱の一括診断") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    modUIProgress.SetStage "未診断の投函をプリフライト診断中", _
                           modConfig.GetLong("llm_wait_sec", 1200)

    Dim n As Long
    n = modPlayOps.RunPreflightAll()

    RefreshInbox
    modUIHome.RefreshHome
    Notice "プリフライト診断: " & CStr(n) & "件を診断しました。"

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' ============================================================================
' 選択行の診断結果表示(11章「行選択→詳細ペイン: 診断全文・組み替え案」)
' ----------------------------------------------------------------------------
' 15章§6 の PF スキーマから、生存見込み・予測類型・所見・組み替え案を読み下す。
' ============================================================================
Public Sub InboxShowDiagnosis()
    If Not modUIProgress.TryEnterUiLock("診断の表示") Then Exit Sub
    On Error GoTo Done

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UI2_SHEET)
    If ws Is Nothing Then GoTo Done

    Dim rowNo As Long
    rowNo = modUISheet.SelectedRow(ws)
    If rowNo < 2 Then
        Notice "診断を見たい行を選んでから押してください（1行目は見出しです）。"
        GoTo Done
    End If

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, 1, 1, UI2_SCAN_COLS)
    If IsEmpty(hdr) Then GoTo Done

    Dim pfJson As String
    pfJson = ColText(ws, hdr, rowNo, "pf_json")
    If LenB(Trim$(pfJson)) = 0 Then
        modUISheet.SetLabelText ws, UI2_LBL_DIAG, "診断: この投函はまだ診断されていません。"
        GoTo Done
    End If

    modUISheet.SetLabelText ws, UI2_LBL_DIAG, DiagnosisText(ws, hdr, rowNo, pfJson)

Done:
    modUIProgress.ExitUiLock
End Sub

' 診断の読み下し(1つの表示文字列)。本文が長いので先頭のみを見せる。
Private Function DiagnosisText(ByVal ws As Object, ByVal hdr As Variant, _
                               ByVal rowNo As Long, ByVal pfJson As String) As String
    Dim acc As String
    acc = "ID: " & ColText(ws, hdr, rowNo, "inbox_id") & vbLf
    acc = acc & "生存見込み: " & modPlayOps.PfSurvivalOf(pfJson) & vbLf
    acc = acc & "予測類型: " & modPlayOps.PfPredTypesOf(pfJson) & vbLf

    Dim reasonText As String
    reasonText = modJsonLite.GetStr(pfJson, "survival_reason")
    If LenB(reasonText) > 0 Then acc = acc & "理由: " & modUtil.SafeLeft(reasonText, 400) & vbLf

    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(pfJson, "rework_ideas")
    Dim i As Long
    For i = 1 To items.count
        acc = acc & "組替案" & CStr(i) & ": " & _
              modUtil.SafeLeft(CStr(items(i)), 200) & vbLf
    Next i

    DiagnosisText = modUtil.SafeLeft(acc, 1800)
End Function

' ============================================================================
' 判定入力(11章「判定入力(T0-T10)」・16章 E-41 統制語彙の必須化)
' ----------------------------------------------------------------------------
' 必須検査は modInboxStore.JudgementError が唯一の判定であり、ここでは
' **判定結果を作らず**、シート上で選ばれた語彙をそのまま渡す(fail-closed。
' 足りなければ SetInboxJudgement が1列も書かずに戻る)。
' 裁定書9 B2(13章§2.6・N5): 判定先は **入力列 judge_to** から読む。status 列は
' 判定の**結果**であり利用者に触らせない(ここでも読まない・書かない。status の
' 書換と judge_to のクリアは SetInboxJudgement 成功時に store 側が行う)。
' 旧方式(status 列を読む)は遷移検査の from と to が同じセル由来になり、
' 自己遷移として必ず弾かれて判定が1件も保存できなかった。
' ============================================================================
Public Sub InboxSaveJudgement()
    If Not modUIProgress.TryEnterUiLock("判定の保存") Then Exit Sub
    On Error GoTo Done

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UI2_SHEET)
    If ws Is Nothing Then GoTo Done

    ' 裁定書10補遺 P1: 選択行の取得は ParkFocus より必ず先に行う。
    Dim rowNo As Long
    rowNo = modUISheet.SelectedRow(ws)

    modUIProgress.ParkFocus

    If rowNo < 2 Then
        Notice "判定する行を選んでから押してください（1行目は見出しです）。"
        GoTo Done
    End If

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, 1, 1, UI2_SCAN_COLS)
    If IsEmpty(hdr) Then GoTo Done

    Dim inboxId As String
    inboxId = ColText(ws, hdr, rowNo, "inbox_id")
    If Not modInboxStore.IsValidInboxId(inboxId) Then
        Notice "inbox_id のある行を選んでください。"
        GoTo Done
    End If

    Dim judgeToJa As String
    Dim judgeTo As String
    Dim dropType As String
    Dim reviveTag As String
    Dim dueText As String
    judgeToJa = Trim$(ColText(ws, hdr, rowNo, "judge_to"))
    dropType = ColText(ws, hdr, rowNo, "drop_type")
    reviveTag = ColText(ws, hdr, rowNo, "revive_tag")
    dueText = ColText(ws, hdr, rowNo, "revive_due")

    If LenB(judgeToJa) = 0 Then
        Notice "judge_to 列で判定先（採択／条件付き保留／却下）を選んでから押してください。"
        GoTo Done
    End If

    ' 裁定書10 M7: judge_to 列は日本語ラベルで選ばせる(19章§3 inbox.judge_to)。
    ' 機械値へ戻すのは modUICase.EnumEn だけであり、表に無いラベルは "" が返る
    ' (推測で近い値へ寄せない。13章§2.2 と同じ fail-closed)。
    judgeTo = modUICase.EnumEn("inbox_judge_to", judgeToJa)
    If LenB(judgeTo) = 0 Then
        Notice "judge_to 列は一覧から選んでください（採択／条件付き保留／却下）: " & judgeToJa
        GoTo Done
    End If

    ' 16章 E-41 の必須検査(唯一の判定点)を先に通し、理由を画面へ出す。
    Dim errText As String
    errText = modInboxStore.JudgementError(judgeTo, dropType, reviveTag, _
                                           IsDateText(dueText))
    If LenB(errText) > 0 Then
        Notice "保存できません: " & errText
        GoTo Done
    End If

    Dim due As Date
    If IsDateText(dueText) Then due = CDate(dueText)

    If modInboxStore.SetInboxJudgement(inboxId, judgeTo, dropType, reviveTag, due) Then
        RefreshInbox
        modUIHome.RefreshHome
        Notice "判定を保存しました: " & inboxId
    Else
        Notice "判定を保存できませんでした（統制語彙・遷移をご確認ください）。"
    End If

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Private Function IsDateText(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Then Exit Function
    IsDateText = IsDate(s)
End Function

' 下書き行の入力列を1つ空へ戻す(列が無ければ何もしない)。
Private Sub ClearDraftCol(ByVal ws As Object, ByVal hdr As Variant, _
                          ByVal rowNo As Long, ByVal colName As String)
    Dim colNo As Long
    colNo = modUtil.FindHeaderCol(hdr, colName)
    If colNo <= 0 Then Exit Sub
    modUISheet.PutText ws, rowNo, colNo, vbNullString, UI2_SHEET & "/" & colName
End Sub

Private Function ColText(ByVal ws As Object, ByVal hdr As Variant, ByVal rowNo As Long, _
                         ByVal colName As String) As String
    Dim colNo As Long
    colNo = modUtil.FindHeaderCol(hdr, colName)
    If colNo <= 0 Then Exit Function
    ColText = modUISheet.CellText(ws, rowNo, colNo)
End Function

Private Sub Notice(ByVal messageText As String)
    On Error Resume Next
    modUISheet.WriteNamed "hm_warning", messageText
    MsgBox messageText, vbInformation, "受信箱"
End Sub
