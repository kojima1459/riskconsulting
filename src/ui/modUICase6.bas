Attribute VB_Name = "modUICase6"
Option Explicit

' ============================================================================
' modUICase6 - ナビ区画②「貼る」の保管＋プレビュー(ui層・T-49)
' ----------------------------------------------------------------------------
' 11章v3.2 §3.3 / §7.2(a)(e) と 13章§2.11 が正。
'
' 方式(v3.2で予約行方式を撤回した理由は11章§3.3.0):
'   [ここに貼る]はクリップボードの**文字だけ**を取り出し、既知フッターを落として
'   から **case_data へ直接保存**する。画面へ書くのは状態行とプレビュー5行だけで、
'   セルへ本文を展開しない(何字貼っても溢れない・画面が重くならない)。
'
' 保存経路は既存契約のまま(11章§7.2(a)): modCaseStore.SaveData が
'   modUtil.SplitForCells(content, 32,000) で case_data の複数行へ入れ、読み出しは
'   modCaseStore.LoadData(区切りなし連結)。**新しい可逆性の契約を1つも増やさない**。
'
' 純関数は core の modNavText / modUIGeom が持つ(層(a)でテストする)。画面へ描く
'   のは modUINavDraw。本モジュールは画面・クリップボード・store をつなぐだけ。
' ============================================================================

Private Const U6_SRC As String = "modUICase6"
Private Const U6_SHEET As String = "ナビ"
Private Const U6_PREVIEW_LINES As Long = 5
' 裁定書22 i1: プレビューの見出し「┈┈┈ 先頭だけお見せします ┈┈┈」の値源は
'   build/sheets_main.json の ci_prev_* の label だけである(ビルドがセルへ焼く)。
'   VBA側の定数は使われないまま二重の値源になっていたので撤去した。
Private Const U6_LOCK_PASTE As String = "貼り付け"
Private Const U6_LOCK_SHOW As String = "中身の表示"
Private Const U6_LOCK_CLEAR As String = "貼ったものの取り消し"
Private Const U6_LOCK_SAVE As String = "貼ったものの保存"

' 図形名の接頭辞(13章§2.11(b))。置き直す前にまとめて落とす。
Private Const U6_BTN_PREFIX As String = "ci_btn_"
Private Const U6_BTN_COL_FIRST As Long = 4
Private Const U6_BTN_COL_LAST As Long = 60
Private Const U6_BTN_GAP As Double = 8#

' 11章§3.3.7 の逐語(2文)。
Private Const U6_MSG_NO_CLIP As String = _
    "貼り付けられませんでした。もう一度コピーしてから、[ここに貼る]を押してください。"
Private Const U6_MSG_NO_SAVE As String = _
    "貼った文章を保存できませんでした。もう一度[ここに貼る]を押してください。" & _
    "直らないときは、使い方タブの[記録を見る]を押して、いちばん下の行を開発担当へ送ってください。"
Private Const U6_MSG_MISMATCH As String = _
    "画面の案件と保存先が合いません。いちばん上の帯で、案件を選び直してください。"
Private Const U6_MSG_PII As String = _
    "個人のお名前らしい記述が見つかったため保存しませんでした。" & _
    "該当の行を消してから、もう一度保存してください。"
Private Const U6_MSG_NOTEPAD_NG As String = _
    "中身を開けませんでした。貼った文章はちゃんと保存されていますので、そのまま先へ進んでください。"
Private Const U6_MSG_FOOTER As String = "（末尾のシステムの表示は取り除きました）"
Private Const U6_MSG_NOT_YET As String = "まだ貼っていません"
' 11章§3.3.7(裁定書22 m3): 画面を描き切れていないときの2文。ナビは1画面なので
'   「HOMEへ戻って」ではなく帯の[ナレッジを読み直す]か開き直しへ誘導する。
Private Const U6_MSG_NO_DRAW As String = _
    "画面を正しく開けませんでした。いちばん上の帯の[ナレッジを読み直す]を押すか、" & _
    "ブックを開き直してください。"
' 11章§3.3.7(裁定書22 M2): 現場メモが60行に入りきらないときの2文。
Private Const U6_MSG_MEMO_ROWS As String = _
    "現場メモが枠に入りきりませんでした。いちばん下の余った行を切り取って、" & _
    "別の見出しの下へ貼ってください。"
Private Const U6_MSG_SAVE_NG As String = "保存できませんでした"

' 一時ファイル(11章§3.3.4(2))。7日より古いものは起動時に消す。
Private Const U6_TMP_PREFIX As String = "rpn_view_"
Private Const U6_TMP_KEEP_DAYS As Long = 7

' 6欄の定義(13章§2.11(a))。data_key 9本は不変。現場メモだけはプレビューを持た
' ない(実体の入力枠そのものが画面)。
'   欄キー|data_key|画面ラベル|プレビュー|直貼り枠|見張り行|状態行
Public Function AreaTable() As String
    Dim s As String
    s = s & "dossier|input_dossier|調べた結果|ci_prev_dossier|ci_raw_dossier|" & _
            "ci_sent_dossier|ci_count_dossier" & vbLf
    s = s & "hp|input_hp|会社のホームページ|ci_prev_hp|ci_raw_hp|" & _
            "ci_sent_hp|ci_count_hp" & vbLf
    s = s & "yuho|input_yuho|有価証券報告書のリスクの章|ci_prev_yuho|ci_raw_yuho|" & _
            "ci_sent_yuho|ci_count_yuho" & vbLf
    s = s & "field_notes|input_memo|現場メモ||ci_area_field_notes|" & _
            "ci_sent_field_notes|ci_count_field_notes" & vbLf
    s = s & "contract|input_contract|いまの契約|ci_prev_contract|ci_raw_contract|" & _
            "ci_sent_contract|ci_count_contract" & vbLf
    s = s & "hearing_answers|input_hearing_answers|ヒアリング回答|ci_prev_hearing_answers|" & _
            "ci_raw_hearing_answers|ci_sent_hearing_answers|ci_count_hearing_answers"
    AreaTable = s
End Function

' 1行を7つの欄へ割る。
Private Sub SplitArea(ByVal lineText As String, ByRef areaKey As String, _
                      ByRef dataKey As String, ByRef labelText As String, _
                      ByRef prevRange As String, ByRef rawRange As String, _
                      ByRef sentRange As String, ByRef countRange As String)
    Dim f() As String
    f = Split(lineText, "|")
    areaKey = vbNullString
    If UBound(f) - LBound(f) < 6 Then Exit Sub
    areaKey = f(0)
    dataKey = f(1)
    labelText = f(2)
    prevRange = f(3)
    rawRange = f(4)
    sentRange = f(5)
    countRange = f(6)
End Sub

' 欄キーから1行を引く。見つからなければ ""。
Private Function AreaLineOf(ByVal areaKey As String) As String
    Dim lines() As String
    lines = Split(AreaTable(), vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If InStr(1, lines(i), areaKey & "|", vbBinaryCompare) = 1 Then
            AreaLineOf = lines(i)
            Exit Function
        End If
    Next i
End Function

' 欄キーの一覧(vbLf区切り)。描画側が並び順どおりに回すための唯一の口。
Public Function AreaKeys() As String
    Dim lines() As String
    lines = Split(AreaTable(), vbLf)
    Dim i As Long
    Dim acc As String
    For i = LBound(lines) To UBound(lines)
        Dim f() As String
        f = Split(lines(i), "|")
        If LenB(acc) > 0 Then acc = acc & vbLf
        acc = acc & f(0)
    Next i
    AreaKeys = acc
End Function

' 欄の定義の1項目(0=キー 1=data_key 2=画面ラベル 3=プレビュー 4=直貼り枠
'   5=見張り行 6=状態行)。表の読み方を描画側へ写さないための口。
Public Function AreaField(ByVal areaKey As String, ByVal idx As Long) As String
    Dim f() As String
    f = Split(AreaLineOf(areaKey), "|")
    If idx < 0 Then Exit Function
    If idx > UBound(f) Then Exit Function
    AreaField = f(idx)
End Function

' その欄に保管された本文(現場メモは枠の中身)。描画側が字数とプレビューに使う。
Public Function AreaBody(ByVal caseId As String, ByVal areaKey As String) As String
    If StrComp(areaKey, "field_notes", vbBinaryCompare) = 0 Then
        AreaBody = ReadFieldNotesArea()
        Exit Function
    End If
    If LenB(caseId) = 0 Then Exit Function
    AreaBody = LoadArea(caseId, AreaField(areaKey, 1))
End Function

' ハンドラ名(dossier -> PasteIntoDossier)。図形の OnAction を組み立てるのに使う。
Public Function HandlerName(ByVal stem As String, ByVal areaKey As String) As String
    Dim parts() As String
    parts = Split(areaKey, "_")
    Dim i As Long
    Dim acc As String
    For i = LBound(parts) To UBound(parts)
        If LenB(parts(i)) > 0 Then
            acc = acc & UCase$(Left$(parts(i), 1)) & Mid$(parts(i), 2)
        End If
    Next i
    HandlerName = stem & acc
End Function

' 保管と読み出し(11章§7.2(a))。契約は既存のまま。

' StoreArea - 本文を case_data へ保存する。True=保存できた。
'   16章 E-04 の浄化と E-31 の匿名化を通してから store へ渡す。
Public Function StoreArea(ByVal caseId As String, ByVal dataKey As String, _
                          ByVal content As String) As Boolean
    On Error GoTo Failed

    Dim body As String
    body = modUtilText.SanitizeInput(content)

    Dim hits As Long
    body = modUICase.AnonymizeText(body, modUISheet.ReadNamed("ci_company"), hits)
    If hits > 0 Then
        modLog.LogError "E0104", U6_SRC & ".StoreArea", "replaced=" & CStr(hits)
    End If

    StoreArea = modCaseStore.SaveData(caseId, dataKey, body)
    Exit Function
Failed:
    modLog.LogError "E0603", U6_SRC & ".StoreArea", "store_failed:" & dataKey, Err.Number
    StoreArea = False
End Function

' LoadArea - case_data から全文を戻す(区切りなし連結。M3の可逆性そのもの)。
Public Function LoadArea(ByVal caseId As String, ByVal dataKey As String) As String
    LoadArea = modCaseStore.LoadData(caseId, dataKey)
End Function

' ReadDirectPaste - Ctrl+V 直貼り枠を上から走査し、最終の非空行までを vbLf で
'   連結して返す(11章§7.2(a))。overflow は見張り行が空でないことで立つ。
Public Function ReadDirectPaste(ByVal rawRange As String, ByVal sentRange As String, _
                                ByRef overflow As Boolean) As String
    overflow = False
    On Error GoTo Failed

    overflow = (LenB(modUISheet.ReadNamed(sentRange)) > 0)

    Dim head As Object
    Set head = modUISheet.NamedCell(rawRange)
    If head Is Nothing Then Exit Function

    Dim rows As Long
    rows = modUISheet.NamedRows(rawRange)
    If rows <= 0 Then Exit Function

    Dim ws As Object
    Set ws = head.Worksheet

    Dim buf() As String
    Dim cnt As Long
    modUtil.BufInit buf, cnt

    Dim i As Long
    Dim last As Long
    For i = 1 To rows
        If LenB(Trim$(CStr(ws.Cells(head.row + i - 1, head.Column).Value))) > 0 Then last = i
    Next i
    For i = 1 To last
        modUtil.BufAdd buf, cnt, CStr(ws.Cells(head.row + i - 1, head.Column).Value)
    Next i

    ReadDirectPaste = modUtil.BufText(buf, cnt)
    Exit Function
Failed:
    ReadDirectPaste = vbNullString
End Function

' SentinelCheck - 見張り行が空でない(=枠に入りきらなかった)か。True=はみ出した。
Public Function SentinelCheck(ByVal areaKey As String) As Boolean
    Dim a As String, d As String, l As String, p As String
    Dim rw As String, sn As String, ct As String
    SplitArea AreaLineOf(areaKey), a, d, l, p, rw, sn, ct
    If LenB(a) = 0 Then Exit Function
    SentinelCheck = (LenB(modUISheet.ReadNamed(sn)) > 0)
End Function

' MergedOrShapeCheck - 直貼り枠に結合セル・図形・画像が入っていないか。
'   True=入っている。実体は modUICase7.MergedOrShapeAt(裁定書22 m2 で図形の
'   交差検知を足したため、30,000字契約により分割先へ置いた)。
Public Function MergedOrShapeCheck(ByVal areaKey As String) As Boolean
    MergedOrShapeCheck = modUICase7.MergedOrShapeAt(areaKey)
End Function

' PasteIntoArea - [ここに貼る]の本体(11章§3.3.4(1)。順序が正)。
Public Sub PasteIntoArea(ByVal areaKey As String)
    On Error GoTo Failed

    ' (0) 裁定書22 m3(B15): 描き切れていない画面から保存しない。
    If Not modUINav.DrawOk() Then
        modUIToast.ShowToast U6_MSG_NO_DRAW, "error"
        Exit Sub
    End If

    Dim a As String, dataKey As String, labelText As String, prevRange As String
    Dim rawRange As String, sentRange As String, countRange As String
    SplitArea AreaLineOf(areaKey), a, dataKey, labelText, prevRange, rawRange, _
              sentRange, countRange
    If LenB(a) = 0 Then Exit Sub

    ' (1) クリップボードからテキストだけを取り出す(書式は読まない)。
    Dim okFlag As Boolean
    Dim raw As String
    raw = modUISheet.PasteFromClipboard(okFlag)

    ' (2) 取り出せなければ枠に一切触らずに終わる(既存の中身を消さない)。
    If Not okFlag Then
        modUIToast.ShowToast U6_MSG_NO_CLIP, "error"
        Exit Sub
    End If

    ' (3) 既知フッターを落とす。(4) 改行を vbLf へ揃える。
    Dim body As String
    body = modNavText.NormalizeEol(modNavText.StripDrFooter(raw))
    Dim footerCut As Boolean
    footerCut = (Len(body) < Len(modNavText.NormalizeEol(raw)))

    ' 個人情報の検査も[ここに貼る]の中で行う(11章§3.3.7。保存まで待たない)。
    '   裁定書22 m8: **採番より前**に行う。個人情報で弾く貼り付けのために案件を
    '   採番すると、中身の無い幽霊案件が案件一覧へ積まれる。
    If modPii.HasPii(body) Then
        modLog.LogError "E0103", U6_SRC & ".PasteIntoArea", _
                        modPii.ScanReport(body, U6_SHEET & ":" & labelText)
        modUIToast.ShowToast U6_MSG_PII, "error"
        Exit Sub
    End If

    ' 画面の案件と保存先を突き合わせる。合わなければ**1字も保存しない**。
    '   画面が新規モード(固定マーカー)なら、ここで採番してから保存する
    '   (13章§2.11(e) の書き手(3)。先に[貼ったものを保存する]を押させないと
    '    1本目が貼れない、という順序の罠を作らないため)。
    Dim caseId As String
    caseId = EnsureCaseId()
    If LenB(caseId) = 0 Then
        modUIToast.ShowToast CaseIdFailText(), "error"
        Exit Sub
    End If

    ' 現場メモだけは実体の入力枠なので、case_data へ直接は書かず枠へ差し込む。
    '   60行に入りきらなければ**1行も書かず**、その欄だけをブロックする(M2)。
    If StrComp(areaKey, "field_notes", vbBinaryCompare) = 0 Then
        If Not AppendToFieldNotes(body) Then
            modUIHome.ShowWarning U6_MSG_MEMO_ROWS
            Exit Sub
        End If
        modUINavDraw.RefreshArea caseId, areaKey
        modUIToast.ShowToast "現場メモへ書き足しました。見出しの下をご確認ください。", "info"
        Exit Sub
    End If

    ' (5) case_data へ直接保存する。
    If Not StoreArea(caseId, dataKey, body) Then
        modUISheet.WriteNamed countRange, U6_MSG_SAVE_NG
        modUIToast.ShowToast U6_MSG_NO_SAVE, "error"
        Exit Sub
    End If

    ' (7) 保存に成功してからプレビューと状態行を書く(順序を逆にしない)。
    modUINavDraw.RefreshArea caseId, areaKey

    ' (8) 成功トースト(11章§2.2 #7 の逐語)。
    Dim tail As String
    If footerCut Then tail = U6_MSG_FOOTER
    modUIToast.ShowToast labelText & "に " & Format$(Len(body), "#,##0") & _
                         "字 保存しました" & tail & "。次の欄へ進んでください。", "info"
    Exit Sub

Failed:
    modLog.LogError "E0603", U6_SRC & ".PasteIntoArea", "paste_failed:" & areaKey, Err.Number
    modUIToast.ShowToast U6_MSG_NO_SAVE, "error"
End Sub

' 現場メモの枠へ差し込む(枠を丸ごとクリアしない。原則⑤)。
'   いま選んでいる見出しが分からないので【そのほか】の節の末尾へ足す。
'   戻り値 False = 60行に入りきらないので**1行も書かなかった**(M2)。
Private Function AppendToFieldNotes(ByVal body As String) As Boolean
    On Error Resume Next
    Dim cur As String
    cur = ReadFieldNotesArea()

    Dim memoText As String
    Dim othersText As String
    modNavText.SplitFieldNotes cur, memoText, othersText
    If LenB(othersText) > 0 Then
        othersText = othersText & vbLf & body
    Else
        othersText = body
    End If
    AppendToFieldNotes = WriteFieldNotesArea(modNavText.JoinFieldNotes(memoText, othersText))
End Function

' ShowArea - [中身を見る](11章§3.3.4(2))。読むだけ。メモ帳で開く。
Public Sub ShowArea(ByVal areaKey As String)
    On Error GoTo Failed

    Dim a As String, dataKey As String, labelText As String, prevRange As String
    Dim rawRange As String, sentRange As String, countRange As String
    SplitArea AreaLineOf(areaKey), a, dataKey, labelText, prevRange, rawRange, _
              sentRange, countRange
    If LenB(a) = 0 Then Exit Sub

    Dim caseId As String
    caseId = ScreenCaseId()

    Dim body As String
    If StrComp(areaKey, "field_notes", vbBinaryCompare) = 0 Then
        body = ReadFieldNotesArea()
    ElseIf LenB(caseId) > 0 Then
        body = LoadArea(caseId, dataKey)
    End If
    If LenB(body) = 0 Then
        modUIToast.ShowToast labelText & "には、まだ何も入っていません。" & _
                             "[ここに貼る]で貼ってください。", "warn"
        Exit Sub
    End If

    Dim pathText As String
    pathText = WriteTempUtf8(dataKey, body)
    If LenB(pathText) = 0 Then
        modUIToast.ShowToast U6_MSG_NOTEPAD_NG, "warn"
        Exit Sub
    End If

    Shell "notepad.exe """ & pathText & """", 1      ' 1 = vbNormalFocus
    modUIToast.ShowToast "中身をメモ帳で開きました。読むだけの画面です。" & _
                         "直したいときは、直した文章を[ここに貼る]で貼り直してください。", "info"
    Exit Sub

Failed:
    modLog.LogError "E0603", U6_SRC & ".ShowArea", "show_failed:" & areaKey, Err.Number
    modUIToast.ShowToast U6_MSG_NOTEPAD_NG, "warn"
End Sub

' 一時ファイルへ UTF-8 BOM 付きで書き、そのパスを返す(失敗は "")。
Private Function WriteTempUtf8(ByVal dataKey As String, ByVal body As String) As String
    On Error GoTo Failed
    Dim pathText As String
    pathText = Environ$("TEMP") & "\" & U6_TMP_PREFIX & dataKey & "_" & _
               Format$(Now, "yyyymmddhhnnss") & ".txt"

    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                                       ' 2 = adTypeText
    st.Charset = "UTF-8"
    st.Open
    st.WriteText body
    st.SaveToFile pathText, 2                         ' 2 = adSaveCreateOverWrite
    st.Close
    WriteTempUtf8 = pathText
    Exit Function
Failed:
    WriteTempUtf8 = vbNullString
End Function

' 7日より古い一時ファイルを消す(11章§3.3.4(2))。起動時に modBoot から呼ぶ。
Public Sub SweepTempViews()
    On Error Resume Next
    Dim dirText As String
    dirText = Environ$("TEMP") & "\"
    Dim nameText As String
    nameText = Dir$(dirText & U6_TMP_PREFIX & "*.txt")
    Do While LenB(nameText) > 0
        If DateDiff("d", FileDateTime(dirText & nameText), Now) > U6_TMP_KEEP_DAYS Then
            Kill dirText & nameText
        End If
        nameText = Dir$
    Loop
End Sub

' ClearArea - [消す](11章§3.3.4(3))。押す前に必ず止める。
Public Sub ClearArea(ByVal areaKey As String)
    On Error GoTo Failed

    Dim a As String, dataKey As String, labelText As String, prevRange As String
    Dim rawRange As String, sentRange As String, countRange As String
    SplitArea AreaLineOf(areaKey), a, dataKey, labelText, prevRange, rawRange, _
              sentRange, countRange
    If LenB(a) = 0 Then Exit Sub

    Dim caseId As String
    caseId = ScreenCaseId()
    If LenB(caseId) = 0 Then
        modUIToast.ShowToast U6_MSG_MISMATCH, "error"
        Exit Sub
    End If

    Dim n As Long
    If StrComp(areaKey, "field_notes", vbBinaryCompare) = 0 Then
        n = Len(ReadFieldNotesArea())
    Else
        n = Len(LoadArea(caseId, dataKey))
    End If
    If n <= 0 Then Exit Sub

    ' 11章§4.3 の例外(v3.2で追加)。9,000字を1クリックで失わせない。
    If MsgBox(labelText & "に貼った " & Format$(n, "#,##0") & _
              "字 を消します。消すと元に戻せません。消してよろしいですか。", _
              vbYesNo + vbQuestion, "リスク提案ナビ") <> vbYes Then Exit Sub

    If StrComp(areaKey, "field_notes", vbBinaryCompare) = 0 Then
        If Not WriteFieldNotesArea(modNavText.JoinFieldNotes(vbNullString, vbNullString)) Then
            modUIToast.ShowToast U6_MSG_MEMO_ROWS, "error"
            Exit Sub
        End If
    Else
        modCaseStore.SaveData caseId, dataKey, vbNullString
    End If
    modUINavDraw.RefreshArea caseId, areaKey
    modUIToast.ShowToast labelText & "を消しました。もう一度[ここに貼る]で貼り直せます。", "info"
    Exit Sub

Failed:
    modLog.LogError "E0603", U6_SRC & ".ClearArea", "clear_failed:" & areaKey, Err.Number
End Sub

' SaveNav - [貼ったものを保存する](13章§2.11(e))。(1)ci_case_id の3値判定
' (2)会社名の必須検査 (3)直貼り枠の取り込み(はみ出し・結合セルはその欄だけ
' ブロック) (4)現場メモの分解保存 (5)属性欄と調査の深さ (6)描き直してトースト。
Public Sub SaveNav()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_SAVE) Then Exit Sub
    On Error GoTo Done

    ' 裁定書22 m3(B15): 描き切れていない画面から保存しない(全欄ブロック)。
    ' modUIHome.ShowWarning は hm_warning とトーストの両方へ出す(2度出さない)。
    If Not modUINav.DrawOk() Then
        modUIHome.ShowWarning U6_MSG_NO_DRAW
        GoTo Done
    End If

    modUIProgress.ParkFocus

    Dim caseId As String
    Dim isNew As Boolean
    caseId = Trim$(modUISheet.ReadNamed("ci_case_id"))
    If StrComp(caseId, modUICase3.U3_NEW_MARK, vbBinaryCompare) = 0 Then
        isNew = True
        caseId = modUICase3.CreateCaseFromSheet()
        If LenB(caseId) = 0 Then
            modUIToast.ShowToast "会社名か業種が空のため、案件を作れませんでした。" & _
                                 "①の会社名と業種名を入れてから、もう一度押してください。", "error"
            GoTo Done
        End If
    ElseIf Not modCaseStore.IsValidCaseId(caseId) Then
        modLog.LogError "E0302", U6_SRC & ".SaveNav", "case_id_blank"
        modUIHome.ShowWarning U6_MSG_MISMATCH
        GoTo Done
    End If

    If LenB(modUISheet.ReadNamed("ci_company")) = 0 Then
        modUIHome.ShowWarning "会社名が空のため保存しませんでした。" & _
            "①の会社名を入れてから、もう一度[貼ったものを保存する]を押してください。"
        GoTo Done
    End If

    Dim blocked As String
    blocked = modUICase7.ImportDirectPastes(caseId)
    SaveFieldNotes caseId
    SaveAttributes caseId
    If isNew Then modUISheet.WriteNamed "hm_case_id", caseId

    ' 画面の描き直しは modUIProgress.ExitUiLock が必ず通す(2度描かない)。
    If LenB(blocked) > 0 Then
        ' 11章§3.3.7 / v2.5.1 M4: 1欄でもブロックしたら成功案内でそれを隠さない。
        '   裁定書22 m2: 理由(はみ出し / 表の線・画像 / 個人情報)ごとの逐語文を
        '   modUICase7 が組み立てて返す(1文にまとめない)。
        modUIHome.ShowWarning "案件 " & caseId & " を保存しました。ただし次の欄は" & _
            "保存していません。" & vbLf & blocked, "warn"
    Else
        modUIToast.ShowToast "案件 " & caseId & " を保存しました。" & _
                             "次は③の[まとめて作る]を押してください。", "info"
    End If
    modLog.LogUsage "case_input_saved", caseId, "nav"

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' 直貼り枠の取り込み(ImportDirectPastes / JoinExisting / ClearDirectPaste)は
' 30,000字契約により modUICase7 へ移設した(裁定書22 m2)。

' 現場メモを見出しで切り分けて保存する(13章§2.11(d))。
Private Sub SaveFieldNotes(ByVal caseId As String)
    On Error Resume Next
    Dim memoText As String
    Dim othersText As String
    modNavText.SplitFieldNotes ReadFieldNotesArea(), memoText, othersText

    StoreArea caseId, "input_memo", memoText
    StoreArea caseId, "input_field_notes", othersText
    ' 画面の欄を持たない2本は常に空文字で保存する(13章§2.11(a))。
    modCaseStore.SaveData caseId, "input_prev_renewal", vbNullString
    modCaseStore.SaveData caseId, "input_coverage_note", vbNullString
End Sub

' 属性欄を案件一覧へ書き、調査の深さを自動決定する(11章§5 #1)。
Private Sub SaveAttributes(ByVal caseId As String)
    On Error Resume Next

    Dim caseType As String
    caseType = modUICase.EnumEn("case_type", modUISheet.ReadNamed("ci_case_type"))
    If LenB(caseType) = 0 Then caseType = "new"

    Dim pairs As String
    pairs = "case_type" & vbTab & caseType & vbLf & _
            "company" & vbTab & modUISheet.ReadNamed("ci_company") & vbLf & _
            "industry_code" & vbTab & modUISheet.ReadNamed("ci_industry_code") & vbLf & _
            "industry_name" & vbTab & modUISheet.ReadNamed("ci_industry_name") & vbLf & _
            "s4_variant" & vbTab & "proposal"
    modUICase3.SetCaseCells caseId, pairs

    ' 11章§5 #1: 「調べた結果」が非空なら t2_full、空なら t1_quick。
    Dim tierText As String
    If LenB(LoadArea(caseId, "input_dossier")) > 0 Then
        tierText = "t2_full"
    Else
        tierText = "t1_quick"
    End If
    modCaseStore.PromoteTier caseId, tierText
    ' 裁定書22 D13: 表示は「しっかり調査（貼った内容から自動で決まります）」。
    '   利用者が「自分で選ぶ欄」と誤解して探し回っていた(選ぶ欄ではない)。
    '   ラベル本体は enum のまま(補足は丸括弧。modUICase.EnumEn が読み戻せる)。
    modUISheet.WriteNamed "ci_dossier_tier", _
        modUICase.EnumJa("dossier_tier", tierText) & "（貼った内容から自動で決まります）"
End Sub

' 現場メモ枠の読み書き(60行の実体入力枠)。
Public Function ReadFieldNotesArea() As String
    Dim overflow As Boolean
    ReadFieldNotesArea = ReadDirectPaste("ci_area_field_notes", "ci_sent_field_notes", overflow)
End Function

' WriteFieldNotesArea - 現場メモ枠(60行)へ書く。**True=書いた / False=1行も
'   書かなかった**(裁定書22 M2)。行数を超える本文を途中まで書くと、切れた行が
'   そのまま保存されて元の文が戻せなくなる。**入らないなら1行も書かない**。
Public Function WriteFieldNotesArea(ByVal body As String) As Boolean
    On Error Resume Next
    Dim head As Object
    Set head = modUISheet.NamedCell("ci_area_field_notes")
    If head Is Nothing Then Exit Function

    Dim rows As Long
    rows = modUISheet.NamedRows("ci_area_field_notes")
    If rows <= 0 Then Exit Function

    ' 入りきるかを先に純関数で確かめる(層(a)で固定した判定。modNavText)。
    If Not modNavText.FitsInRows(body, rows) Then Exit Function

    Dim ws As Object
    Set ws = head.Worksheet

    Dim lines() As String
    lines = Split(modNavText.NormalizeEol(body), vbLf)

    Dim i As Long
    For i = 1 To rows
        Dim s As String
        s = vbNullString
        If i - 1 <= UBound(lines) Then s = lines(i - 1)
        modUISheet.PutText ws, head.row + i - 1, head.Column, s, U6_SRC & "/field_notes"
    Next i
    WriteFieldNotesArea = True
End Function

' 画面が新規モードなら採番してから案件IDを返す(13章§2.11(e) の書き手(3))。
'   採番できなければ ""(呼び出し側が案内を出す)。
Private Function EnsureCaseId() As String
    Dim v As String
    v = Trim$(modUISheet.ReadNamed("ci_case_id"))
    If modCaseStore.IsValidCaseId(v) Then
        EnsureCaseId = v
        Exit Function
    End If
    If StrComp(v, modUICase3.U3_NEW_MARK, vbBinaryCompare) <> 0 Then Exit Function

    Dim caseId As String
    caseId = modUICase3.CreateCaseFromSheet()
    If LenB(caseId) = 0 Then Exit Function
    modUISheet.WriteNamed "hm_case_id", caseId
    EnsureCaseId = caseId
End Function

' 案件IDを用意できなかったときの案内(原因で文を分ける。原則③の2文)。
Private Function CaseIdFailText() As String
    If StrComp(Trim$(modUISheet.ReadNamed("ci_case_id")), modUICase3.U3_NEW_MARK, _
               vbBinaryCompare) = 0 Then
        CaseIdFailText = "先に①の会社名と業種名を入れてください。" & _
                         "どの会社の話かが決まらないと、貼った文章を保存できません。"
    Else
        CaseIdFailText = U6_MSG_MISMATCH
    End If
End Function

' 画面の案件ID(3値判定で有効なものだけ返す。合わなければ "")。
Private Function ScreenCaseId() As String
    Dim v As String
    v = Trim$(modUISheet.ReadNamed("ci_case_id"))
    If Not modCaseStore.IsValidCaseId(v) Then Exit Function
    ScreenCaseId = v
End Function

' 出る条件(13章§2.11(a))。偽の欄は行ごと非表示にし、中身は消さない。
Public Function AreaHidden(ByVal areaKey As String) As Boolean
    If StrComp(areaKey, "contract", vbBinaryCompare) = 0 Then
        AreaHidden = (StrComp(modUICase.EnumEn("case_type", _
                      modUISheet.ReadNamed("ci_case_type")), "renewal", vbBinaryCompare) <> 0)
    ElseIf StrComp(areaKey, "hearing_answers", vbBinaryCompare) = 0 Then
        AreaHidden = (Val(modUISheet.ReadNamed("hm_round_no")) < 2)
    End If
End Function

' ============================================================================
' OnAction ハンドラ(18本。13章§2.11(b))。実体は共通の PasteIntoArea /
' ShowArea / ClearArea で、ここは「どの欄か」を渡すだけの薄い口である。
' 16章E-11: OnActionで配線される公開Subは先頭で TryEnterUiLock を通す。
' ============================================================================
Public Sub PasteIntoDossier()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_PASTE) Then Exit Sub
    PasteIntoArea "dossier"
    modUIProgress.ExitUiLock
End Sub

Public Sub PasteIntoHp()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_PASTE) Then Exit Sub
    PasteIntoArea "hp"
    modUIProgress.ExitUiLock
End Sub

Public Sub PasteIntoYuho()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_PASTE) Then Exit Sub
    PasteIntoArea "yuho"
    modUIProgress.ExitUiLock
End Sub

Public Sub PasteIntoFieldNotes()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_PASTE) Then Exit Sub
    PasteIntoArea "field_notes"
    modUIProgress.ExitUiLock
End Sub

Public Sub PasteIntoContract()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_PASTE) Then Exit Sub
    PasteIntoArea "contract"
    modUIProgress.ExitUiLock
End Sub

Public Sub PasteIntoHearingAnswers()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_PASTE) Then Exit Sub
    PasteIntoArea "hearing_answers"
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAreaDossier()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_SHOW) Then Exit Sub
    ShowArea "dossier"
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAreaHp()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_SHOW) Then Exit Sub
    ShowArea "hp"
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAreaYuho()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_SHOW) Then Exit Sub
    ShowArea "yuho"
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAreaFieldNotes()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_SHOW) Then Exit Sub
    ShowArea "field_notes"
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAreaContract()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_SHOW) Then Exit Sub
    ShowArea "contract"
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAreaHearingAnswers()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_SHOW) Then Exit Sub
    ShowArea "hearing_answers"
    modUIProgress.ExitUiLock
End Sub

Public Sub ClearAreaDossier()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_CLEAR) Then Exit Sub
    ClearArea "dossier"
    modUIProgress.ExitUiLock
End Sub

Public Sub ClearAreaHp()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_CLEAR) Then Exit Sub
    ClearArea "hp"
    modUIProgress.ExitUiLock
End Sub

Public Sub ClearAreaYuho()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_CLEAR) Then Exit Sub
    ClearArea "yuho"
    modUIProgress.ExitUiLock
End Sub

Public Sub ClearAreaFieldNotes()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_CLEAR) Then Exit Sub
    ClearArea "field_notes"
    modUIProgress.ExitUiLock
End Sub

Public Sub ClearAreaContract()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_CLEAR) Then Exit Sub
    ClearArea "contract"
    modUIProgress.ExitUiLock
End Sub

Public Sub ClearAreaHearingAnswers()
    If Not modUIProgress.TryEnterUiLock(U6_LOCK_CLEAR) Then Exit Sub
    ClearArea "hearing_answers"
    modUIProgress.ExitUiLock
End Sub
