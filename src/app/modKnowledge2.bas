Attribute VB_Name = "modKnowledge2"
Option Explicit

' ============================================================================
' modKnowledge2 - ナレッジブックの【純関数】部(絞込・列引き・E-34の純部。T-21)
' ----------------------------------------------------------------------------
' 30,000字契約(12章§2)による modKnowledge の分割先(17章§7 Z-13)。分割の軸は
'   modKnowledge  = シートに触る側(kb_pathの読込・キャッシュ・退避/復元・再送)
'   modKnowledge2 = **シートを触らない純関数**(blk を引数で受けて絞り込む)
' であり、**挙動は分割前と1つも変えていない**(Private -> Public の可視性変更と
' 呼出の修飾だけ)。modKnowledgeFmt が「15章の整形規約」を持つのと同じ切り口で、
' こちらは「13章§3の列と絞込の規約」を持つ。
'
' blk = Range.Value 由来の2次元配列(1行目=見出し)、lastRow = データ最終行。
' Excelトークンを一切持たないので層(a)から直接叩ける(R4の許可は要らない)。
'
' 絞込スペック(SelectRows)は "ind^tgt^act^sts^suf^ref" の位置指定(未使用は空)。
'   ind=業種完全一致列 / tgt=対象業種列(";"区切り・空=指定なし) / act=真偽列(偽を
'   捨てる) / sts=型のstatus列(許可値 KB_SCHEME_STATUS) / suf=行頭名称へ "(値)" を
'   付す列(整形は modKnowledgeFmt。ここでは注入IDとして積むためだけに見る) /
'   ref=";"区切り参照IDも注入IDへ積む列。
' 公開口(14章§6へ登記。呼ぶのは modKnowledge だけ):
'   PickAt / CellAt / CellRaw / AddIdList / ColOf / SelectRows /
'   MissingColsOf / BadRowsOf
' ============================================================================

Private Const KB_BAR As String = "|"
Private Const KB_SPACE As String = " "
Private Const KB_SEMI As String = ";"
' 型ライブラリのS3注入対象 status(13章§3.4)。絞込キー sts の許可値。
Private Const KB_SCHEME_STATUS As String = "proven;adopted"
' リスクユニバース10分類(19章§3)。13章§3.1がLoadKnowledgeの検査項目と定める。
Private Const KB_CATEGORIES As String = "strategy_market;supply_chain;manufacturing_quality;sales_customer;facility_bcp;hr_labor;digital_info;legal_regulatory;finance_counterparty;brand_social"
' E0402 detail へ並べる行番号の最大件数(400字上限の手前で止める)。
Private Const KB_MAX_ROWNOS As Long = 12

' "a|b|c" の idx 番目(1始まり)。範囲外は ""。
Public Function PickAt(ByVal listText As String, ByVal idx As Long) As String
    On Error GoTo Blank0
    Dim parts As Variant
    parts = Split(listText, KB_BAR)
    If idx >= 1 Then
        If idx <= UBound(parts) + 1 Then PickAt = CStr(parts(idx - 1))
    End If
    Exit Function
Blank0:
    PickAt = vbNullString
End Function

' 1セルをテキストへ。改行・タブはスペースへ畳み(1行1件の規約)、外部由来なので
' SanitizeInput を通す(16章 E-04・E-43)。
Public Function CellAt(ByVal blk As Variant, ByVal r As Long, ByVal c As Long) As String
    On Error GoTo Blank0
    If c <= 0 Then Exit Function
    Dim t As String
    t = CStr(blk(r, c))
    If LenB(t) = 0 Then Exit Function
    t = Replace(Replace(Replace(Replace(t, vbCrLf, KB_SPACE), vbCr, KB_SPACE), vbLf, KB_SPACE), vbTab, KB_SPACE)
    CellAt = Trim$(modUtilText.SanitizeInput(t))
    Exit Function
Blank0:
    CellAt = vbNullString
End Function

' 退避用の素読み。タブは区切りに使うのでスペースへ寄せる。
Public Function CellRaw(ByVal blk As Variant, ByVal r As Long, ByVal c As Long) As String
    On Error GoTo Blank0
    CellRaw = Replace(CStr(blk(r, c)), vbTab, KB_SPACE)
    Exit Function
Blank0:
    CellRaw = vbNullString
End Function

Private Sub AppendPart(ByRef acc As String, ByVal sepText As String, ByVal partText As String)
    If LenB(partText) = 0 Then Exit Sub
    If LenB(acc) = 0 Then
        acc = partText
    Else
        acc = acc & sepText & partText
    End If
End Sub

Private Sub AddId(ByRef idsOut As String, ByVal idText As String)
    idsOut = modUtil.AppendIdList(idsOut, idText)
End Sub

Public Sub AddIdList(ByRef idsOut As String, ByVal listText As String)
    If LenB(listText) = 0 Then Exit Sub
    Dim parts As Variant
    Dim i As Long
    parts = Split(listText, KB_SEMI)
    For i = LBound(parts) To UBound(parts)
        AddId idsOut, Trim$(CStr(parts(i)))
    Next i
End Sub

' ";"区切りリストに値が含まれるか(前後空白は無視・大小文字は区別)。
Private Function IsListed(ByVal listText As String, ByVal valueText As String) As Boolean
    Dim v As String
    v = Trim$(valueText)
    If LenB(v) = 0 Then Exit Function
    IsListed = (InStr(1, KB_SEMI & listText & KB_SEMI, KB_SEMI & v & KB_SEMI, vbBinaryCompare) > 0)
End Function

' 業種の完全一致(exact=True)または対象業種リストへの当てはまり(exact=False。
' 空欄=指定なし=全業種)。industryCode が空なら常に True(全業種)。
Private Function IndustryHit(ByVal cellText As String, ByVal industryCode As String, _
                             ByVal exactMatch As Boolean) As Boolean
    Dim code As String, t As String
    code = Trim$(industryCode)
    t = Trim$(cellText)
    If exactMatch Then
        IndustryHit = (LenB(code) = 0)
        If Not IndustryHit Then IndustryHit = (StrComp(t, code, vbTextCompare) = 0)
        Exit Function
    End If
    t = Replace(t, KB_SPACE, vbNullString)
    If LenB(t) = 0 Or LenB(code) = 0 Then
        IndustryHit = True
        Exit Function
    End If
    IndustryHit = (InStr(1, KB_SEMI & t & KB_SEMI, KB_SEMI & code & KB_SEMI, vbTextCompare) > 0)
End Function

' 列名(空可)から列位置を引く。空・不在は0。
Public Function ColOf(ByVal blk As Variant, ByVal colName As String) As Long
    If LenB(Trim$(colName)) > 0 Then ColOf = modUtil.FindHeaderCol(blk, Trim$(colName))
End Function

' 絞込の純部。blk の2行目以降から条件に合う行を最大 maxRows 件選び、
'   「見出し行 + 選ばれた行」だけの2次元配列を selOut へ返す(整形は
'   modKnowledgeFmt の責務)。戻り=選ばれた行数。注入IDもここで積む。
Public Function SelectRows(ByVal blk As Variant, ByVal lastRow As Long, ByVal idCol As String, _
                            ByVal filterSpec As String, ByVal industryCode As String, _
                            ByVal maxRows As Long, ByRef selOut As Variant, _
                            ByRef idsOut As String) As Long
    selOut = Empty
    If Not IsArray(blk) Then Exit Function

    Dim cId As Long, cInd As Long, cTgt As Long, cAct As Long
    Dim cSt As Long, cSuf As Long, cRef As Long
    Dim parts As Variant
    cId = modUtil.FindHeaderCol(blk, idCol)
    parts = Split(filterSpec & "^^^^^", "^")
    cInd = ColOf(blk, CStr(parts(0)))
    cTgt = ColOf(blk, CStr(parts(1)))
    cAct = ColOf(blk, CStr(parts(2)))
    cSt = ColOf(blk, CStr(parts(3)))
    cSuf = ColOf(blk, CStr(parts(4)))
    cRef = ColOf(blk, CStr(parts(5)))

    Dim hits() As Long
    ReDim hits(1 To lastRow + 1)
    Dim idText As String, sufText As String
    Dim r As Long, n As Long
    Dim okAll As Boolean
    For r = 2 To lastRow
        If n >= maxRows Then Exit For
        idText = CellAt(blk, r, cId)
        If LenB(idText) > 0 Then
            okAll = True
            If cInd > 0 Then okAll = IndustryHit(CellAt(blk, r, cInd), industryCode, True)
            If okAll And cTgt > 0 Then okAll = IndustryHit(CellAt(blk, r, cTgt), industryCode, False)
            If okAll And cAct > 0 Then okAll = modConfig.ParseBoolText(CellAt(blk, r, cAct), True)
            If okAll And cSt > 0 Then okAll = IsListed(KB_SCHEME_STATUS, CellAt(blk, r, cSt))
            If okAll Then
                If cSuf > 0 Then
                    sufText = CellAt(blk, r, cSuf)
                    If LenB(sufText) > 0 Then AddId idsOut, sufText
                End If
                AddId idsOut, idText
                If cRef > 0 Then AddIdList idsOut, CellAt(blk, r, cRef)
                n = n + 1
                hits(n) = r
            End If
        End If
    Next r
    If n = 0 Then Exit Function

    Dim cols As Long
    cols = UBound(blk, 2)
    Dim arr() As Variant
    ReDim arr(1 To n + 1, 1 To cols)
    Dim c As Long, k As Long
    For c = 1 To cols
        arr(1, c) = blk(1, c)
    Next c
    For k = 1 To n
        For c = 1 To cols
            arr(k + 1, c) = blk(hits(k), c)
        Next c
    Next k
    selOut = arr
    SelectRows = n
End Function

' 16章 E-34(純部): 必須列のうち見つからなかった列名を ";" 区切りで返す。
Public Function MissingColsOf(ByVal blk As Variant, ByVal requiredCols As String) As String
    If Not IsArray(blk) Then
        MissingColsOf = requiredCols
        Exit Function
    End If
    Dim parts As Variant
    Dim i As Long
    Dim acc As String
    parts = Split(requiredCols, KB_SEMI)
    For i = LBound(parts) To UBound(parts)
        If LenB(Trim$(CStr(parts(i)))) > 0 Then
            If modUtil.FindHeaderCol(blk, CStr(parts(i))) <= 0 Then
                AppendPart acc, KB_SEMI, CStr(parts(i))
            End If
        End If
    Next i
    MissingColsOf = acc
End Function

' 不正行の行番号を "," 区切りで返す。checkKind: dup=ID重複 / cat=categoryが
'   19章§3の10分類外(13章§3.1)。
Public Function BadRowsOf(ByVal blk As Variant, ByVal lastRow As Long, _
                           ByVal colName As String, ByVal checkKind As String) As String
    If Not IsArray(blk) Then Exit Function
    Dim c As Long
    c = modUtil.FindHeaderCol(blk, colName)
    If c <= 0 Then Exit Function

    Dim seen As String, cellText As String, acc As String
    Dim r As Long, n As Long
    Dim isBad As Boolean
    For r = 2 To lastRow
        cellText = CellAt(blk, r, c)
        If LenB(cellText) > 0 Then
            If checkKind = "cat" Then
                isBad = Not IsListed(KB_CATEGORIES, cellText)
            Else
                isBad = IsListed(seen, cellText)
                If Not isBad Then seen = seen & KB_SEMI & cellText
            End If
            If isBad And n < KB_MAX_ROWNOS Then
                AppendPart acc, ",", CStr(r)
                n = n + 1
            End If
        End If
    Next r
    BadRowsOf = acc
End Function
