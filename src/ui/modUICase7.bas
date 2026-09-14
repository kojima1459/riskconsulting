Attribute VB_Name = "modUICase7"
Option Explicit

' ============================================================================
' modUICase7 - 区画②の直貼り枠の取り込み(ui層・modUICase6 の分割先)
' ----------------------------------------------------------------------------
' 30,000字契約(12章§2)による modUICase6 の分割先(裁定書22)。分割の軸は
'   modUICase6 = [ここに貼る]／[中身を見る]／[消す]／[貼ったものを保存する]の
'                本体と、6欄の定義表(AreaTable)
'   modUICase7 = **Ctrl+V 直貼り枠(§3.3.4(4))の取り込みと、その失敗の理由分け**
' であり、挙動は分割前と同じ経路を通る(呼び出し口は SaveNav の1本だけ)。
'
' 裁定書22 m2: 直貼りの失敗を **overflow(枠に入りきらない)と merged(表の線・
'   画像が入っている)へ理由別に分け**、11章§3.3.7 のそれぞれの逐語2文を出す。
'   分ける前は両方が「長さを気にせず入ります」の1文へまとめられており、
'   画像を貼った人は何度貼り直しても直らなかった。
'
' 図形・画像の検知(m2): 直貼り枠の**矩形と交差する Shape の名前を集めて数える**。
'   自分たちが置いたボタン・パネル・トーストは接頭辞で除く(自作の図形を
'   「利用者が貼り込んだ画像」と誤検知しない)。
' ============================================================================

Private Const U7_SRC As String = "modUICase7"
Private Const U7_SHEET As String = "ナビ"

' 自分たちが置いた図形の接頭辞(vbLf 区切り)。これらは検知の対象外。
Private Const U7_OWN_PREFIX As String = _
    "btn_" & vbLf & "ci_btn_" & vbLf & "nv_" & vbLf & "ts_" & vbLf & _
    "lbl_" & vbLf & "btncopy_" & vbLf & "gt_"

' 11章§3.3.7 の逐語(理由別の2文ずつ)。〈欄名〉は呼び出し側で前に付ける。
Private Const U7_MSG_OVERFLOW As String = _
    "が枠に入りきりませんでした。[ここに貼る]で貼り直すと、長さを気にせず入ります。"
Private Const U7_MSG_MERGED As String = _
    "の枠の中に、貼り付けで入った表の線や画像があります。[ここに貼る]で貼り直すと入りません。"
Private Const U7_MSG_PII As String = _
    "に個人のお名前らしい記述が見つかったため保存しませんでした。" & _
    "該当の行を消してから、もう一度保存してください。"
' 裁定書38 Z-46: 検知種別が policy_no だけのときは登録は止めず警告のみ出す。
Private Const U7_MSG_POLICY As String = _
    "に契約番号らしき数字列が見つかりました。伏せ字にするか、そのままでよいか確認してください。"

' [中身を見る]の表示シート(裁定書27 W9-B3。詳細は本モジュール末尾の節を参照)。
Private Const U7_BODY_SHEET As String = "中身"
Private Const U7_BODY_CHUNK As Long = 32000
Private Const U7_BODY_ROW0 As Long = 3
Private Const U7_BODY_MAX_ROWS As Long = 2000

' クリップボードの受け皿(裁定書27 W9-B1・裁定書44 A-8c。詳細は本モジュール
'   末尾の節を参照)。読む側(paste_buf)と書く側(copy_buf)は**別の器**にする。
Private Const U7_BUF_SHEET As String = "paste_buf"
Private Const U7_COPY_BUF_SHEET As String = "copy_buf"
Private Const U7_FMT_JA As String = "Unicode テキスト"
Private Const U7_FMT_EN As String = "Unicode Text"
Private Const U7_SHEET_VISIBLE As Long = -1        ' xlSheetVisible
Private Const U7_BUF_MAX_ROWS As Long = 20000
Private Const U7_BUF_MAX_COLS As Long = 50

' ============================================================================
' ImportDirectPastes - 直貼り枠を case_data へ取り込む(SaveNav から1回だけ)。
'   戻り値: ブロックした欄の案内文(理由ごとに1文ずつ。何も無ければ "")。
' ============================================================================
Public Function ImportDirectPastes(ByVal caseId As String) As String
    Dim overflowLabels As String
    Dim mergedLabels As String
    Dim piiLabels As String
    Dim policyLabels As String

    Dim keys() As String
    keys = Split(modUICase6.AreaKeys(), vbLf)

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim areaKey As String
        areaKey = keys(i)
        If LenB(areaKey) > 0 Then
            If StrComp(areaKey, "field_notes", vbBinaryCompare) <> 0 Then
                Dim dataKey As String
                Dim labelText As String
                Dim rawRange As String
                Dim sentRange As String
                dataKey = modUICase6.AreaField(areaKey, 1)
                labelText = modUICase6.AreaField(areaKey, 2)
                rawRange = modUICase6.AreaField(areaKey, 4)
                sentRange = modUICase6.AreaField(areaKey, 5)

                Dim overflow As Boolean
                Dim body As String
                body = modUICase6.ReadDirectPaste(rawRange, sentRange, overflow)

                If overflow Then
                    overflowLabels = AddLabel(overflowLabels, labelText)
                ElseIf MergedOrShapeAt(areaKey) Then
                    mergedLabels = AddLabel(mergedLabels, labelText)
                ElseIf LenB(body) > 0 Then
                    body = modNavText.NormalizeEol(modNavText.StripDrFooter(body))
                    Dim piiKinds As String
                    piiKinds = modPii.KindsOf(body)
                    If LenB(piiKinds) > 0 And piiKinds <> "policy_no" Then
                        piiLabels = AddLabel(piiLabels, labelText)
                        modLog.LogError "E0103", U7_SRC & ".ImportDirectPastes", _
                                        modPii.ScanReport(body, U7_SHEET & ":" & labelText)
                    Else
                        ' 裁定書38 Z-46: policy_no だけの検知は登録したうえで
                        ' 警告(labelを別枠に積む。人名/メール/電話が無ければ通す)。
                        If piiKinds = "policy_no" Then
                            policyLabels = AddLabel(policyLabels, labelText)
                            modLog.LogUsage "pii_policy_warning", caseId, _
                                            modPii.ScanReport(body, U7_SHEET & ":" & labelText)
                        End If
                        If modUICase6.StoreArea(caseId, dataKey, _
                                    JoinExisting(modUICase6.LoadArea(caseId, dataKey), body)) Then
                            ClearDirectPaste rawRange
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ImportDirectPastes = BlockedText(overflowLabels, mergedLabels, piiLabels, policyLabels)
End Function

' 理由ごとの逐語文を1本につなぐ(空の理由は文を作らない)。
'   policyLabels(裁定書38 Z-46)は他の3つと違い**ブロックしていない**(登録済み
'   の欄への警告)。既存3引数の呼び出し元(あれば)は互換のため Optional。
Public Function BlockedText(ByVal overflowLabels As String, _
                            ByVal mergedLabels As String, _
                            ByVal piiLabels As String, _
                            Optional ByVal policyLabels As String = vbNullString) As String
    Dim acc As String
    If LenB(overflowLabels) > 0 Then acc = Add1(acc, overflowLabels & U7_MSG_OVERFLOW)
    If LenB(mergedLabels) > 0 Then acc = Add1(acc, mergedLabels & U7_MSG_MERGED)
    If LenB(piiLabels) > 0 Then acc = Add1(acc, piiLabels & U7_MSG_PII)
    If LenB(policyLabels) > 0 Then acc = Add1(acc, policyLabels & U7_MSG_POLICY)
    BlockedText = acc
End Function

Private Function Add1(ByVal acc As String, ByVal oneText As String) As String
    If LenB(acc) = 0 Then
        Add1 = oneText
    Else
        Add1 = acc & vbLf & oneText
    End If
End Function

Private Function AddLabel(ByVal acc As String, ByVal labelText As String) As String
    If LenB(acc) = 0 Then
        AddLabel = labelText
    Else
        AddLabel = acc & "・" & labelText
    End If
End Function

' ============================================================================
' MergedOrShapeAt - 直貼り枠に結合セル、または図形・画像が入っていないか。
'   True=入っている(その欄だけブロックする)。11章§3.3.7。
' ============================================================================
Public Function MergedOrShapeAt(ByVal areaKey As String) As Boolean
    On Error GoTo Failed

    Dim rawRange As String
    rawRange = modUICase6.AreaField(areaKey, 4)
    If LenB(rawRange) = 0 Then Exit Function

    Dim rng As Object
    Set rng = ThisWorkbook.Names(rawRange).RefersToRange
    If rng.MergeCells <> False Then
        MergedOrShapeAt = True
        Exit Function
    End If

    MergedOrShapeAt = (ShapeHitCount(rng) > 0)
    Exit Function
Failed:
    MergedOrShapeAt = False
End Function

' 枠の矩形と交差する「利用者が貼り込んだ図形」の本数。
'   名前を先に集めてから数える(Shapes を列挙しながら触らない。11章§8.6 禁忌3)。
Public Function ShapeHitCount(ByVal rng As Object) As Long
    On Error GoTo Failed

    Dim ws As Object
    Set ws = rng.Worksheet

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    x1 = rng.Left
    y1 = rng.Top
    x2 = x1 + rng.Width
    y2 = y1 + rng.Height

    Dim names() As String
    Dim cnt As Long
    modUtil.BufInit names, cnt

    Dim shp As Object
    For Each shp In ws.Shapes
        If Not IsOwnShape(shp.Name) Then
            If shp.Left < x2 And (shp.Left + shp.Width) > x1 Then
                If shp.Top < y2 And (shp.Top + shp.Height) > y1 Then
                    modUtil.BufAdd names, cnt, shp.Name
                End If
            End If
        End If
    Next shp

    ShapeHitCount = cnt
    Exit Function
Failed:
    ' 数えられないときは「無い」と言い切らない(fail-closed にはしないが、
    ' 誤って全欄をブロックもしない)。0を返して結合セル判定だけに委ねる。
    ShapeHitCount = 0
End Function

' 自分たちが置いた図形か(接頭辞の完全一致)。
Private Function IsOwnShape(ByVal shapeName As String) As Boolean
    Dim pre() As String
    pre = Split(U7_OWN_PREFIX, vbLf)
    Dim i As Long
    For i = LBound(pre) To UBound(pre)
        If Len(shapeName) >= Len(pre(i)) Then
            If StrComp(Left$(shapeName, Len(pre(i))), pre(i), vbBinaryCompare) = 0 Then
                IsOwnShape = True
                Exit Function
            End If
        End If
    Next i
End Function

' 既存の保管と直貼りぶんを継ぐ(区切りは vbLf。既存が空ならそのまま)。
Public Function JoinExisting(ByVal oldText As String, ByVal addText As String) As String
    If LenB(oldText) = 0 Then
        JoinExisting = addText
    Else
        JoinExisting = oldText & vbLf & addText
    End If
End Function

' 直貼り枠を空へ戻す(取り込んだら枠は空にする。11章§7.2(a))。
Private Sub ClearDirectPaste(ByVal rawRange As String)
    On Error Resume Next
    ThisWorkbook.Names(rawRange).RefersToRange.ClearContents
End Sub

' ============================================================================
' クリップボード(裁定書27 W9-B1。11章§3.3 の挙動は変えない)
' ----------------------------------------------------------------------------
' 撤去したもの: CLSID指定のCOM生成(GetObject の "New" 形式)による MSForms.DataObject の
'   遅延生成。参照設定なしでクリップボードを読める定番手だったが、CLSIDでの
'   COM生成は社内AVのAMSIがマクロ型マルウェアの特徴として重く見る形であり
'   (2026-09-02 実測)、配布物から消す。
'
' 代わりに使うもの: **Excel自身の貼り付けとコピー**。
'   読む: veryHidden の受け皿シート `paste_buf` を作り、A1へ
'         `PasteSpecial Format:="Unicode テキスト"` で貼る。**書式名は数値定数を
'         使わない**(数値のFormatは他の形式を指す)。日本語Excelと英語Excelで
'         名前が違うので "Unicode テキスト" -> "Unicode Text" の順に試す。
'         貼り付いたセルは modNavText.JoinPasteCells で1本のテキストへ戻し、
'         受け皿シートは消す。
'   書く: **別のveryHidden受け皿 `copy_buf`**(裁定書44 A-8c)へ1行1セルで書き、
'         その範囲を `Copy` する。**範囲Copyはクリップボードへの参照**なので、
'         受け皿シートは消さずに veryHidden のまま残す(消すと貼り付け先で
'         空になる)。次にコピーするときだけ、書く**直前**に Clear する
'         (コピー**直後**には触らない)。
'
' なぜ読む側と書く側を分けたか(裁定書44 A-8c。実機NG: T47-W9-01/T47B-W61-15):
'   もとは読み書きとも `paste_buf` 1枚を共用していた。Excel の Copy は
'   遅延レンダリング(貼り付け側が実際に読むまでコピー元セルの中身を確定
'   しない)なので、コピーした直後に**同じセル**を貼り付け側の下ごしらえで
'   Clear すると、確定前のクリップボードの中身ごと消える。ClipPasteText は
'   読む前に必ず `ws.Cells.Clear` するので、直前に ClipCopyText で書いた
'   `paste_buf` をそのまま読もうとすると自分で消したあとを読むことになり、
'   往復が必ず空になっていた(LibreOfficeは遅延レンダリングを実装していない
'   ので層(a)では再現しない・実機だけの壊れ方)。書く器を copy_buf へ分けた
'   ことで、ClipPasteText の Clear は paste_buf だけに閉じ、直前の
'   ClipCopyText が copy_buf へ持たせたクリップボード参照に触らなくなる。
'
' 受け皿シートは実行時生成の作業シートであり、13章§2.9 の `enum_hidden` と
'   同じ扱い(仕様上のシートではないので照合対象外・配布ビルドに焼かない)。
'
' PasteSpecial は**対象シートがアクティブでないと使えない**ため、読むときだけ
'   受け皿を一瞬可視にして戻す。ちらつきは ScreenUpdating を落として抑える。
' ============================================================================

' ClipPasteText - クリップボードのテキストを読む(書式・画像は持ち込まない)。
'   戻り値: 本文。okFlag=False のときは呼び出し側が Ctrl+V の代替枠へ落とす。
Public Function ClipPasteText(ByRef okFlag As Boolean) As String
    Dim prevSheet As Object
    Dim prevUpdate As Boolean
    Dim prevAlerts As Boolean
    Dim ws As Object
    okFlag = False
    On Error GoTo Failed

    Set prevSheet = ThisWorkbook.ActiveSheet
    prevUpdate = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set ws = modUISheet.EnsureHiddenSheet(U7_BUF_SHEET)
    If ws Is Nothing Then GoTo Cleanup
    ws.Cells.Clear
    ws.Visible = U7_SHEET_VISIBLE
    ws.Activate
    ws.Cells(1, 1).Select

    ' 裁定書34 §1.3(b): 会社PCの実走で「書式名を指定した PasteSpecial」が3本とも
    '   通らない端末があった(Excel の版と言語で書式名が変わる)。書式名を要らない
    '   Worksheet.Paste を**先に**試し、駄目なら従来の2書式へ落ちる。
    '   Paste は書式ごと貼るが、ここは使い捨ての受け皿シートで、読み出すのは
    '   ReadPasteBuf の文字だけなので持ち込む書式は捨てられる。
    Dim pasted As Boolean
    pasted = TryPasteWorksheet(ws)
    If Not pasted Then pasted = TryPasteFormat(ws, U7_FMT_JA)
    If Not pasted Then pasted = TryPasteFormat(ws, U7_FMT_EN)
    If Not pasted Then GoTo Cleanup

    Dim bodyText As String
    bodyText = ReadPasteBuf(ws)
    If LenB(bodyText) > 0 Then
        okFlag = True
        ClipPasteText = bodyText
    End If

Cleanup:
    DropPasteBuf ws, prevSheet, prevUpdate, prevAlerts
    Exit Function
Failed:
    okFlag = False
    ClipPasteText = vbNullString
    Resume Cleanup
End Function

' 書式名を指定せずに貼ってみる(裁定書34 §1.3(b))。貼るものが無い・貼れない
'   ときは False を返し、呼び出し側が書式名つきの経路へ落ちる。
Private Function TryPasteWorksheet(ByVal ws As Object) As Boolean
    On Error GoTo Failed
    ws.Paste ws.Cells(1, 1)
    TryPasteWorksheet = True
    Exit Function
Failed:
    TryPasteWorksheet = False
End Function

' 1つの書式名で貼ってみる(名前が通らなければ False)。
Private Function TryPasteFormat(ByVal ws As Object, ByVal formatName As String) As Boolean
    On Error GoTo Failed
    ws.Cells(1, 1).PasteSpecial Format:=formatName
    TryPasteFormat = True
    Exit Function
Failed:
    TryPasteFormat = False
End Function

' 貼り付いた範囲を1本のテキストへ戻す(連結の規則は modNavText が唯一持つ)。
Private Function ReadPasteBuf(ByVal ws As Object) As String
    On Error GoTo Failed
    Dim usedR As Object
    Set usedR = ws.UsedRange
    If usedR Is Nothing Then Exit Function

    Dim rowCount As Long
    Dim colCount As Long
    rowCount = usedR.Rows.count
    colCount = usedR.Columns.count
    If rowCount <= 0 Or colCount <= 0 Then Exit Function
    If rowCount > U7_BUF_MAX_ROWS Then rowCount = U7_BUF_MAX_ROWS
    If colCount > U7_BUF_MAX_COLS Then colCount = U7_BUF_MAX_COLS

    Dim buf() As String
    ReDim buf(0 To (rowCount * colCount) - 1)

    Dim r As Long
    Dim c As Long
    For r = 1 To rowCount
        For c = 1 To colCount
            buf(((r - 1) * colCount) + (c - 1)) = CellTextOf(usedR, r, c)
        Next c
    Next r
    ReadPasteBuf = modNavText.JoinPasteCells(buf, rowCount, colCount)
    Exit Function
Failed:
    ReadPasteBuf = vbNullString
End Function

' 1セルの文字列(エラー値・空セルは "")。
Private Function CellTextOf(ByVal rng As Object, ByVal r As Long, ByVal c As Long) As String
    On Error GoTo Failed
    CellTextOf = CStr(rng.Cells(r, c).Value)
    Exit Function
Failed:
    CellTextOf = vbNullString
End Function

' 受け皿を消して画面を元へ戻す(読み取りの後始末。11章§0.2 タブを増やさない)。
Private Sub DropPasteBuf(ByVal ws As Object, ByVal prevSheet As Object, _
                         ByVal prevUpdate As Boolean, ByVal prevAlerts As Boolean)
    On Error Resume Next
    Application.CutCopyMode = False
    If Not prevSheet Is Nothing Then prevSheet.Activate
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevUpdate
End Sub

' ClipCopyText - テキストをクリップボードへ入れる(11章§3.2 の[コピー])。
'   1行1セルで**書く専用の受け皿 copy_buf**(裁定書44 A-8c)へ書き、その範囲を
'   Copy する。**Copy した直後の copy_buf は消さない**(範囲Copyはクリップ
'   ボードへの参照であり、消すと貼り付け先が空になる。Excelの遅延レンダリング
'   のため、コピー元セルを消すとクリップボードの中身ごと消えるのが実機で
'   起きたNGの原因だった)。ここでの `ws.Cells.Clear` は**次にコピーする直前**
'   の下ごしらえであり、前回コピーぶんの後始末ではない。
'   `"` とタブを含む行があるときは False を返して呼び出し側の
'   「セルを選ぶので Ctrl+C してください」へ落とす: Excelのテキスト形式は
'   その2文字を含むセルを引用符で包み直すため、黙って中身を変えてしまう。
Public Function ClipCopyText(ByVal payloadText As String) As Boolean
    On Error GoTo Failed
    If LenB(payloadText) = 0 Then Exit Function

    Dim bodyText As String
    bodyText = modNavText.NormalizeEol(payloadText)
    If InStr(1, bodyText, """", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, bodyText, vbTab, vbBinaryCompare) > 0 Then Exit Function

    Dim lineList() As String
    lineList = Split(bodyText, vbLf)

    Dim n As Long
    n = UBound(lineList) - LBound(lineList) + 1
    If n <= 0 Or n > U7_BUF_MAX_ROWS Then Exit Function

    Dim ws As Object
    Set ws = modUISheet.EnsureHiddenSheet(U7_COPY_BUF_SHEET)
    If ws Is Nothing Then Exit Function
    ws.Cells.Clear

    Dim i As Long
    For i = LBound(lineList) To UBound(lineList)
        modUtilText.SetCellSafe ws.Cells(i - LBound(lineList) + 1, 1), lineList(i), _
                                U7_SRC & "/clip_copy"
    Next i

    ws.Range(ws.Cells(1, 1), ws.Cells(n, 1)).Copy
    ClipCopyText = True
    Exit Function
Failed:
    ClipCopyText = False
End Function

' CopyBufSheetName / PasteBufSheetName - 裁定書44 A-8c: 書く器と読む器が
'   同じ名前(定数の取り違え)に戻っていないかを純テストが固定するための窓。
'   実体(EnsureHiddenSheet を叩くかどうか)は持たない=どちらもExcelを開かず
'   呼べる(層(a))。
Public Function CopyBufSheetName() As String
    CopyBufSheetName = U7_COPY_BUF_SHEET
End Function

Public Function PasteBufSheetName() As String
    PasteBufSheetName = U7_BUF_SHEET
End Function

' ============================================================================
' [中身を見る]の表示先シート(裁定書27 W9-B3。11章§3.3.5)
' ----------------------------------------------------------------------------
' 撤去したもの: `Shell "notepad.exe ..."` と、そのための %TEMP% への一時ファイル
'   書き出し(ADODB.Stream)。外部プロセスの起動は社内AVが重く見る形であり、
'   一時ファイルは「消し忘れると本文が端末に残る」問題も抱えていた。
'
' 代わりに使うもの: **ブック内のシート「中身」**。保管庫の全文を行分割して
'   流し込み、可視にして見せる。[閉じる]で本文を消して非表示へ戻す
'   (本文をブックに残さない)。読むだけの面なので入力させない。
'
' 1セルの契約は 32,000字(13章§2.2・16章 E-22)。長い本文はその境界で複数セルへ
'   分け、modUtil.SplitForCells が分け方の唯一の持ち主である(切り口が
'   サロゲートペアを割らない)。
'
' シートは実行時生成の作業シート(13章§2.9 の `enum_hidden` と同じ扱い。仕様上の
'   シートではないので照合対象外・配布ビルドには焼かない)。**利用者が既定で
'   見るタブは2枚**(11章§0.2)なので、閉じたら veryHidden へ戻す。
' ============================================================================

' ShowBodySheet - 全文を「中身」シートへ流し込んで見せる(成功で True)。
Public Function ShowBodySheet(ByVal titleText As String, ByVal bodyText As String) As Boolean
    On Error GoTo Failed
    If LenB(bodyText) = 0 Then Exit Function

    Dim ws As Object
    Set ws = modUISheet.EnsureHiddenSheet(U7_BODY_SHEET)
    If ws Is Nothing Then Exit Function
    ws.Cells.Clear

    modUtilText.SetCellSafe ws.Cells(1, 2), titleText, U7_SRC & "/body_title"

    Dim parts() As String
    parts = modUtil.SplitForCells(bodyText, U7_BODY_CHUNK)

    Dim i As Long
    Dim rowNo As Long
    rowNo = U7_BODY_ROW0
    For i = LBound(parts) To UBound(parts)
        If rowNo - U7_BODY_ROW0 >= U7_BODY_MAX_ROWS Then Exit For
        modUtilText.SetCellSafe ws.Cells(rowNo, 2), parts(i), U7_SRC & "/body_text"
        rowNo = rowNo + 1
    Next i

    ' 読むだけの面なので、折り返して上詰めで置く(横スクロールさせない)。
    ws.Columns(2).ColumnWidth = 110
    ws.Range(ws.Cells(U7_BODY_ROW0, 2), ws.Cells(rowNo, 2)).WrapText = True

    modUISheet.EnsureBackButton ws
    modUISheet.EnsureButton ws, "btn_body_close", "閉じる", 1, 4, 100#, _
                            "modUICase7.CloseBodySheet"

    If Not modUISheet.ShowSheet(U7_BODY_SHEET) Then Exit Function
    ShowBodySheet = True
    Exit Function
Failed:
    ShowBodySheet = False
End Function

' CloseBodySheet - [閉じる](図形のOnAction)。本文を消して非表示へ戻し、ナビへ。
Public Sub CloseBodySheet()
    If Not modUIProgress.TryEnterUiLock("中身を閉じる") Then Exit Sub
    On Error GoTo Done
    Dim ws As Object
    Set ws = modUISheet.SheetOf(U7_BODY_SHEET)
    If Not ws Is Nothing Then
        ws.Cells.Clear
        modUISheet.EnsureHiddenSheet U7_BODY_SHEET
    End If
    ' ナビへ戻す。**modUINav.BackToNav は呼ばない**(あちらも同じ関所を取るので
    ' 二重取得で弾かれ、行き止まりになる)。表示だけを直接動かす。
    modUISheet.ShowSheet U7_SHEET
Done:
    modUIProgress.ExitUiLock
End Sub
