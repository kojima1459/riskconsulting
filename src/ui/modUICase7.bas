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

' ============================================================================
' ImportDirectPastes - 直貼り枠を case_data へ取り込む(SaveNav から1回だけ)。
'   戻り値: ブロックした欄の案内文(理由ごとに1文ずつ。何も無ければ "")。
' ============================================================================
Public Function ImportDirectPastes(ByVal caseId As String) As String
    Dim overflowLabels As String
    Dim mergedLabels As String
    Dim piiLabels As String

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
                    If modPii.HasPii(body) Then
                        piiLabels = AddLabel(piiLabels, labelText)
                        modLog.LogError "E0103", U7_SRC & ".ImportDirectPastes", _
                                        modPii.ScanReport(body, U7_SHEET & ":" & labelText)
                    ElseIf modUICase6.StoreArea(caseId, dataKey, _
                                JoinExisting(modUICase6.LoadArea(caseId, dataKey), body)) Then
                        ClearDirectPaste rawRange
                    End If
                End If
            End If
        End If
    Next i

    ImportDirectPastes = BlockedText(overflowLabels, mergedLabels, piiLabels)
End Function

' 理由ごとの逐語文を1本につなぐ(空の理由は文を作らない)。
Public Function BlockedText(ByVal overflowLabels As String, _
                            ByVal mergedLabels As String, _
                            ByVal piiLabels As String) As String
    Dim acc As String
    If LenB(overflowLabels) > 0 Then acc = Add1(acc, overflowLabels & U7_MSG_OVERFLOW)
    If LenB(mergedLabels) > 0 Then acc = Add1(acc, mergedLabels & U7_MSG_MERGED)
    If LenB(piiLabels) > 0 Then acc = Add1(acc, piiLabels & U7_MSG_PII)
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
