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

' 裁定書39 R1-03: n-gram 比較に掛ける字数の上限。maxChars<=0 は無制限
'   (既存の呼出と挙動を変えないため)。切るのは**比較に使う複製だけ**で、
'   注入する本文(整形は modKnowledgeFmt)は1字も削らない。
Private Function CapText(ByVal t As String, ByVal maxChars As Long) As String
    If maxChars <= 0 Then
        CapText = t
        Exit Function
    End If
    If Len(t) <= maxChars Then
        CapText = t
        Exit Function
    End If
    CapText = Left$(t, maxChars)
End Function

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
'   totalHits(裁定書38 B-10 / 裁定書39 R1-04 / 裁定書40 P-M2): 打切り前に条件へ
'   合致した総行数(ByRef)。**並べ替え補充が走ったときは「完全一致の該当総数 +
'   実際に補充で採用した行数」**を返す。補充で使った行を分母に数えないと、
'   run_log の `kb_cut:cases=5/2`(「該当2件のうち5件を使用」)という読めない値に
'   なり、`meta.kb_usage`(SEC-14。`cases_total > cases_used` のときだけ出す)が
'   **補充の起きた案件では必ず消える**=最も起きやすい打切りが見えなくなる。
'   逆に**候補を作っただけで足してもいけない**(裁定書40 P-M2): 補充候補は業種
'   完全一致を外した行=実質シートの残り全部なので、1件も採らなかった案件でも
'   分母がシート行数近くまで膨らみ、SEC-14 が常時「該当N件のうちM件を使用」と
'   嘘をつく。数えるのは**採用した行だけ**である。
'   caseChars/rowChars(裁定書39 R1-03): n-gram 比較に掛ける字数上限
'   (案件側 config `kb_rank_case_chars`・行側 `kb_rank_row_chars`)。0=無制限。
'   maxCandRows/candSeenOut(裁定書40 P-M3): 並べ替えに掛ける**候補行数**の上限
'   (config `kb_rank_max_rows`)と、上限を掛ける前の候補総数(ByRef)。0=無制限。
'   上限を超えた分は**シート順で先頭 maxCandRows 行**を採る(順位付けの前に
'   落とすので、KB が育っても並べ替えの手間が増えない)。呼出側は
'   candSeenOut > maxCandRows で「打ち切った」ことを run_log へ残す。
'   caseText/rankCols(同B-10・並べ替え): 完全一致(cInd)の該当数が maxRows に
'   満たないとき、全業種の行から caseText との 2〜3字 n-gram 重なり数
'   (modKnowledgeRank)が高い順に不足分を補う。rankCols は補充候補の本文列を
'   ";" 区切りで指定する(例 "customer_profile;risk_presented")。どちらかが
'   空、または cInd を使わない絞込(schemes 等)では補充は行わない(挙動不変)。
Public Function SelectRows(ByVal blk As Variant, ByVal lastRow As Long, ByVal idCol As String, _
                            ByVal filterSpec As String, ByVal industryCode As String, _
                            ByVal maxRows As Long, ByRef selOut As Variant, _
                            ByRef idsOut As String, ByRef totalHits As Long, _
                            Optional ByVal caseText As String = vbNullString, _
                            Optional ByVal rankCols As String = vbNullString, _
                            Optional ByVal caseChars As Long = 0, _
                            Optional ByVal rowChars As Long = 0, _
                            Optional ByVal maxCandRows As Long = 0, _
                            Optional ByRef candSeenOut As Long = 0) As Long
    selOut = Empty
    totalHits = 0
    candSeenOut = 0
    If Not IsArray(blk) Then Exit Function
    ' 裁定書40 P-m2: データ行が1行も無い(lastRow<2)なら、下の
    '   ReDim hits(1 To lastRow + 1) が lastRow<0 で実行時エラー9 を投げる。
    '   本関数は 14章§6 の公開口であり呼び出し側がエラーを握らないので、
    '   入口で静かに0件で返す(lastRow=0/1 のときの従来の戻り値と同じ)。
    If lastRow < 2 Then Exit Function

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
        idText = CellAt(blk, r, cId)
        If LenB(idText) > 0 Then
            okAll = True
            If cInd > 0 Then okAll = IndustryHit(CellAt(blk, r, cInd), industryCode, True)
            If okAll And cTgt > 0 Then okAll = IndustryHit(CellAt(blk, r, cTgt), industryCode, False)
            If okAll And cAct > 0 Then okAll = modConfig.ParseBoolText(CellAt(blk, r, cAct), True)
            If okAll And cSt > 0 Then okAll = IsListed(KB_SCHEME_STATUS, CellAt(blk, r, cSt))
            If okAll Then
                totalHits = totalHits + 1
                If n < maxRows Then
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
        End If
    Next r

    ' 並べ替え補充(裁定書38 B-10): 完全一致(cInd)がある絞込で、かつ件数が
    ' maxRows に満たないときだけ、全業種の行から関連度上位を補う。
    ' lastRow >= 2(データ行が1行以上ある)を条件に加える(裁定書39 G-1): 無いと
    ' 下の ReDim candRows(1 To lastRow) が lastRow=0 で実行時エラー9 を投げる
    ' (本関数は 14章§6 の公開口であり、呼び出し側がエラーを握らない)。
    If lastRow >= 2 And n < maxRows And cInd > 0 And _
       LenB(caseText) > 0 And LenB(rankCols) > 0 Then
        Dim rcNames As Variant
        rcNames = Split(rankCols, KB_SEMI)
        Dim rCols() As Long
        ReDim rCols(LBound(rcNames) To UBound(rcNames))
        Dim rc As Long
        For rc = LBound(rcNames) To UBound(rcNames)
            rCols(rc) = ColOf(blk, Trim$(CStr(rcNames(rc))))
        Next rc

        Dim candRows() As Long
        Dim candTexts() As String
        Dim candN As Long, candSeen As Long, candCap As Long
        ReDim candRows(1 To lastRow)
        ReDim candTexts(1 To lastRow)
        ' 裁定書40 P-M3(b): 並べ替えに掛ける候補行数の上限(config
        '   kb_rank_max_rows。0=無制限)。超えた分はシート順で切る。
        candCap = maxCandRows
        If candCap <= 0 Or candCap > lastRow Then candCap = lastRow
        Dim already As Boolean, h As Long
        For r = 2 To lastRow
            idText = CellAt(blk, r, cId)
            If LenB(idText) > 0 Then
                already = False
                For h = 1 To n
                    If hits(h) = r Then
                        already = True
                        Exit For
                    End If
                Next h
                If Not already Then
                    okAll = True
                    If cTgt > 0 Then okAll = IndustryHit(CellAt(blk, r, cTgt), industryCode, False)
                    If okAll And cAct > 0 Then okAll = modConfig.ParseBoolText(CellAt(blk, r, cAct), True)
                    If okAll And cSt > 0 Then okAll = IsListed(KB_SCHEME_STATUS, CellAt(blk, r, cSt))
                    If okAll Then
                        candSeen = candSeen + 1
                        ' 上限を超えた候補は**本文を組み立てずに数だけ数える**
                        '   (組立と CapText が候補1行あたりの主費用。裁定書40 P-M3)。
                        If candN < candCap Then
                            candN = candN + 1
                            candRows(candN) = r
                            Dim rowText As String
                            rowText = vbNullString
                            For rc = LBound(rCols) To UBound(rCols)
                                If rCols(rc) > 0 Then rowText = rowText & KB_SPACE & CellAt(blk, r, rCols(rc))
                            Next rc
                            candTexts(candN) = CapText(rowText, rowChars)
                        End If
                    End If
                End If
            End If
        Next r

        ' 裁定書40 P-M3(b): 上限を掛ける前の候補総数を呼出側へ返す
        '   (candSeenOut > 上限 なら呼出側が run_log へ「打ち切った」と残す)。
        candSeenOut = candSeen

        If candN > 0 Then
            Dim candTextsUsed() As String
            ReDim candTextsUsed(1 To candN)
            Dim ci As Long
            For ci = 1 To candN
                candTextsUsed(ci) = candTexts(ci)
            Next ci
            Dim rankText As String
            rankText = CapText(caseText, caseChars)
            Dim order() As Long
            Dim rn As Long
            rn = modKnowledgeRank.RankRows(rankText, candTextsUsed, order)
            Dim need As Long, taken As Long, pickIdx As Long, scoreCheck As Long
            need = maxRows - n
            For ci = 1 To rn
                If taken >= need Then Exit For
                pickIdx = order(ci)
                ' 関連度0の行は補わない(13章§3.1)。点数は「2字+3字」だが、
                '   3字が一致するなら先頭2字も必ず一致するので、**2字が0なら
                '   点数も0**である。1行につき2回だった呼び出しを1回にする
                '   (裁定書40 P-M3 と同型の「ループの中で案件側の索引を作り直す」
                '   を半減させる。呼出は need(<=kb_case_rows)回で止まる)。
                scoreCheck = modKnowledgeRank.NgramOverlap(rankText, candTextsUsed(pickIdx), 2)
                If scoreCheck <= 0 Then Exit For ' 降順なのでここで以降も0
                ' 裁定書39 R1-04 / 裁定書40 P-M2: 分母に数えるのは**採用した行だけ**。
                '   候補を作った時点で足すと、1件も採っていない案件でも
                '   `cases_total > cases_used` が真になり SEC-14 が嘘をつく。
                totalHits = totalHits + 1
                n = n + 1
                hits(n) = candRows(pickIdx)
                idText = CellAt(blk, candRows(pickIdx), cId)
                If cSuf > 0 Then
                    sufText = CellAt(blk, candRows(pickIdx), cSuf)
                    If LenB(sufText) > 0 Then AddId idsOut, sufText
                End If
                AddId idsOut, idText
                If cRef > 0 Then AddIdList idsOut, CellAt(blk, candRows(pickIdx), cRef)
                taken = taken + 1
            Next ci
        End If
    End If

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

' ShouldFallbackToCommonIndustry - 裁定書47 G-5: 業種コード完全一致で0行の
'   とき"00"(共通)へ引き直すべきか。0行かつ既にcommonCodeでなければTrue。
Public Function ShouldFallbackToCommonIndustry(ByVal matchedRows As Long, _
        ByVal industryCode As String, ByVal commonCode As String) As Boolean
    ShouldFallbackToCommonIndustry = (matchedRows <= 0 And industryCode <> commonCode)
End Function
