Attribute VB_Name = "modUICaseFmt"
Option Explicit

' ============================================================================
' modUICaseFmt - セル <-> JSON値 の変換(13章§2.2 セル格納規約。ui層・T-31)
' ----------------------------------------------------------------------------
' 30,000字契約(12章§2)による modUICase の分割先。**規約そのもの**(「; 」区切り
' 配列・整数配列・preventions の `measure(M-0012)` 形式・入れ子オブジェクトの
' 畳み込み・空配列は空セル・日本語ラベルenum)を1箇所に集めたもので、
' modKnowledgeFmt が15章の整形規約を引き受けているのと同じ切り口である。
'
' 本モジュールは**Excelに触れない純関数だけ**で構成する(セルの読み書きは
' modUISheet、ブロックの走査は modUICase2 の責務)。日本語ラベルの変換だけは
' 19章§3の変換表を持つ modUICase へ委ねる(表を2箇所に書かない)。
'
' 列定義の書式(modUICase2 が持つ 13章§2.12-§2.15 の写像表):
'   物理名:種別[:enumグループ][@スキーマ写像]
'   種別 o=順序列(非スキーマ) / s=文字列 / n=整数 / a=「; 」区切り配列 /
'        ia=整数配列 / e=日本語ラベルenum / p=preventions
' ============================================================================

Private Const UF_SEP As String = "; "          ' 13章§2.2 配列のセル格納区切り

' ============================================================================
' 13章§2.12-§2.15 の列定義(物理名:種別[:enumグループ][@スキーマ写像])
' ----------------------------------------------------------------------------
'   種別 o=順序列(非スキーマ) / s=文字列 / n=整数 / a=「; 」区切り配列 /
'        ia=整数配列 / e=日本語ラベルenum / p=preventions
'   @ の右はスキーマ上のパス(省略時は物理名と同じ)。ドット付きは入れ子で、
'   **同じ親を持つ列は隣り合わせに置く**(13章の列順がそうなっている)。
' ============================================================================
Public Function ColsS1Basic() As String
    Dim s As String
    s = s & "company_name:s;business_summary:s;main_products:a;processes:a;"
    s = s & "supply_chain_key_materials:a@supply_chain.key_materials;"
    s = s & "supply_chain_notes:s@supply_chain.notes;"
    s = s & "customers_segments:a@customers.segments;"
    s = s & "customers_channels:a@customers.channels;"
    s = s & "workforce_notes:s;management_notes:s;"
    s = s & "mvv:s@strategy_outlook.mvv;"
    s = s & "aspirations:a@strategy_outlook.aspirations;"
    s = s & "market_context:s@strategy_outlook.market_context;"
    s = s & "input_quality_overall:e:input_quality_overall@input_quality.overall;"
    s = s & "input_quality_advice:s@input_quality.advice;"
    ' v2.6(裁定書25 S3): financials の6項目を fin_ 接頭辞で本ブロックへ収める。
    s = s & "fin_fiscal_year:s@financials.fiscal_year;"
    s = s & "fin_net_assets:s@financials.net_assets;"
    s = s & "fin_sales:s@financials.sales;"
    s = s & "fin_operating_profit:s@financials.operating_profit;"
    s = s & "fin_source:e:financials_source@financials.source;"
    s = s & "fin_note:s@financials.note"
    ColsS1Basic = s
End Function

Public Function ColsS1Locations() As String
    ColsS1Locations = "seq:o;name:s;type:e:location_type;address:s;" & _
                      "hazard_note:s;notes:s"
End Function

Public Function ColsS1Coverage() As String
    ColsS1Coverage = "seq:o;line_name:s;coverage_summary:s;limit_note:s;" & _
                     "special_note:s;certainty:e:certainty"
End Function

Public Function ColsS1Insights() As String
    ColsS1Insights = "seq:o;note:s;tag:e:field_insight_tag"
End Function

Public Function ColsS1Missing() As String
    ColsS1Missing = "seq:o;item:s;why_needed:s"
End Function

Public Function ColsS1Quality() As String
    ColsS1Quality = "aspect:e:input_quality_aspect;status:e:input_quality_status"
End Function

Public Function ColsS1Research() As String
    ColsS1Research = "seq:o;purpose:s;prompt_text:s"
End Function

Public Function ColsS2Gaps() As String
    ColsS2Gaps = "gap_no:n;gap_type:e:gap_type;target:s;description:s;" & _
                 "risk_evidence:s;coverage_evidence:s"
End Function

Public Function ColsS2Risks() As String
    Dim s As String
    s = s & "risk_no:n;category:e:risk_category;risk_name:s;scenario:s;"
    s = s & "status:e:risk_status;frequency:e:frequency;impact:e:impact;"
    s = s & "frequency_score:n;impact_score:n;"
    s = s & "evidence_quote:s@evidence.quote;"
    s = s & "evidence_source:e:evidence_source@evidence.source;"
    s = s & "transferability:e:transferability@insurability.transferability;"
    s = s & "line_note:s@insurability.line_note;"
    s = s & "gap_note:s@insurability.gap_note;"
    s = s & "control_note:s@insurability.control_note;"
    s = s & "loss_scale_note:s;check_points:a;preventions:p"
    ColsS2Risks = s
End Function

Public Function ColsS2Emerging() As String
    Dim s As String
    s = s & "emg_no:o;risk_name:s;category:e:risk_category;horizon:e:horizon;"
    s = s & "scenario:s;evidence_quote:s;evidence_source:e:evidence_source;"
    s = s & "proposal_hint:s"
    ColsS2Emerging = s
End Function

Public Function ColsS3Stories() As String
    Dim s As String
    s = s & "story_no:n;proposal_kind:e:proposal_kind;headline:s;hook_question:s;"
    s = s & "target_risk_nos:ia;target_gap_nos:ia;menu_ids:a;line_ids:a;"
    s = s & "scheme_id:s;pitch:s;similar_case_id:s;expected_objection:s;"
    s = s & "objection_response:s"
    ColsS3Stories = s
End Function

Public Function ColsS3Unmatched() As String
    ColsS3Unmatched = "risk_no:n;risk_name:s;why_unmatched:s"
End Function

Public Function ColsS3DoNot() As String
    ColsS3DoNot = "seq:o;topic:s;reason:s"
End Function

' 13章§2.14 v2.6 追補: 攻めの保険活用(15章 SchemaS3 の growth_ideas)。
Public Function ColsS3Growth() As String
    ColsS3Growth = "title:s;what:s;why:s;insurance_fit:s;effect:n;" & _
                   "difficulty:e:growth_difficulty"
End Function

' 13章§2.14 v2.6: 経営層向けトークスクリプト(データ1行のみ)。
Public Function ColsS3Talk() As String
    ColsS3Talk = "opening:s;flow:a;closing:s;taboo:a"
End Function

Public Function ColsS4Meta() As String
    ColsS4Meta = "file_title:s"
End Function

Public Function ColsS4Slides() As String
    ColsS4Slides = "slide_no:n;title:s;bullets:a;notes:s"
End Function

Public Function ColsS4Hearing() As String
    ColsS4Hearing = "seq:o;question:s;purpose:s"
End Function

' シート上のブロックの並び(部屋の計算に使う。11章のワイヤーの縦並び順)。
Public Function AnchorsOf(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 1
        AnchorsOf = "s1_basic;s1_locations;s1_current_coverage;s1_field_insights;" & _
                    "s1_missing_info;s1_input_quality;s1_research_requests"
    Case 2
        AnchorsOf = "s2_gaps;s2_risks;s2_open_questions;s2_emerging"
    Case 3
        AnchorsOf = "s3_stories;s3_unmatched_risks;s3_do_not_propose;" & _
                    "s3_growth_ideas;s3_talk_script"
    Case 4
        AnchorsOf = "s4_meta;s4_slides;s4_hearing_questions"
    End Select
End Function


' ============================================================================
' 列定義の解釈と値の変換(13章§2.2 セル格納規約)
' ============================================================================
Public Sub SplitCol(ByVal spec As String, ByRef physName As String, _
                     ByRef kindText As String, ByRef extraText As String, _
                     ByRef pathText As String)
    physName = vbNullString
    kindText = vbNullString
    extraText = vbNullString
    pathText = vbNullString

    Dim body As String
    body = Trim$(spec)

    Dim atPos As Long
    atPos = InStr(1, body, "@", vbBinaryCompare)
    If atPos > 0 Then
        pathText = Mid$(body, atPos + 1)
        body = Left$(body, atPos - 1)
    End If

    Dim parts() As String
    parts = Split(body, ":")
    physName = Trim$(parts(LBound(parts)))
    If UBound(parts) - LBound(parts) >= 1 Then kindText = Trim$(parts(LBound(parts) + 1))
    If UBound(parts) - LBound(parts) >= 2 Then extraText = Trim$(parts(LBound(parts) + 2))
    If LenB(pathText) = 0 Then pathText = physName
End Sub

' 1件ぶんのJSONオブジェクト。入れ子(evidence / insurability)は、同じ親を持つ
'   隣り合う列を1つの子オブジェクトへ畳む(13章§2.2 入れ子オブジェクトの規約)。
Public Function RowObjJson(ByVal colSpec As String, ByVal vals As Variant) As String
    Dim specs() As String
    specs = Split(colSpec, ";")

    Dim outText As String
    Dim subText As String
    Dim curParent As String
    Dim i As Long
    Dim physName As String
    Dim kindText As String
    Dim extraText As String
    Dim pathText As String
    Dim parentName As String
    Dim leafName As String

    For i = LBound(specs) To UBound(specs)
        SplitCol specs(i), physName, kindText, extraText, pathText
        If kindText <> "o" Then
            SplitPath pathText, parentName, leafName
            If parentName <> curParent Then
                FlushParent outText, curParent, subText
                curParent = parentName
                subText = vbNullString
            End If
            AppendFrag subText, """" & leafName & """:" & _
                       ValueJson(kindText, extraText, _
                                 CStr(vals(i - LBound(specs) + LBound(vals))))
        End If
    Next i
    FlushParent outText, curParent, subText
    RowObjJson = "{" & outText & "}"
End Function

Private Sub SplitPath(ByVal pathText As String, ByRef parentName As String, _
                      ByRef leafName As String)
    Dim dotPos As Long
    dotPos = InStrRev(pathText, ".")
    If dotPos > 0 Then
        parentName = Left$(pathText, dotPos - 1)
        leafName = Mid$(pathText, dotPos + 1)
    Else
        parentName = vbNullString
        leafName = pathText
    End If
End Sub

Private Sub FlushParent(ByRef outText As String, ByVal parentName As String, _
                        ByVal subText As String)
    If LenB(subText) = 0 Then Exit Sub
    Dim frag As String
    If LenB(parentName) = 0 Then
        frag = subText
    Else
        frag = """" & parentName & """:{" & subText & "}"
    End If
    AppendFrag outText, frag
End Sub

Private Sub AppendFrag(ByRef acc As String, ByVal frag As String)
    If LenB(acc) > 0 Then acc = acc & ","
    acc = acc & frag
End Sub

' 名前付きの1フラグメントを積む(トップレベル用)。
Public Sub AddFrag(ByRef acc As String, ByVal keyName As String, ByVal valueText As String)
    AppendFrag acc, """" & keyName & """:" & valueText
End Sub

' セル値 -> JSON値。種別は列定義の2つ目。
Private Function ValueJson(ByVal kindText As String, ByVal extraText As String, _
                           ByVal cellValue As String) As String
    Select Case kindText
    Case "n"
        ValueJson = NumJson(cellValue)
    Case "a"
        ValueJson = ArrJson(cellValue)
    Case "ia"
        ValueJson = IntArrJson(cellValue)
    Case "e"
        ValueJson = EnumJson(extraText, cellValue)
    Case "p"
        ValueJson = PrevJson(cellValue)
    Case Else
        ValueJson = StrJson(cellValue)
    End Select
End Function

Public Function StrJson(ByVal s As String) As String
    StrJson = """" & modJsonLite.EscapeJsonStr(s) & """"
End Function

' 数値列。数字として読めなければ 0 を書く(型を崩さない。値の妥当性は
'   modValidate が範囲で弾く)。
Private Function NumJson(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If LenB(t) = 0 Then
        NumJson = "0"
        Exit Function
    End If
    If Not IsNumeric(t) Then
        NumJson = "0"
        Exit Function
    End If
    NumJson = CStr(CLng(Val(t)))
End Function

' 「; 」区切り配列。空セルは [](13章§2.2 空配列規約)。
Public Function ArrJson(ByVal s As String) As String
    ArrJson = "[]"
    If LenB(Trim$(s)) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(s, UF_SEP)

    Dim acc As String
    Dim i As Long
    Dim one As String
    For i = LBound(parts) To UBound(parts)
        one = Trim$(parts(i))
        If LenB(one) > 0 Then
            If LenB(acc) > 0 Then acc = acc & ","
            acc = acc & StrJson(one)
        End If
    Next i
    If LenB(acc) = 0 Then Exit Function
    ArrJson = "[" & acc & "]"
End Function

' 整数配列。数値化できない要素は破棄せずそのまま文字列として載せ、
'   modValidate に不合格として弾かせる(13章§2.2)。
Private Function IntArrJson(ByVal s As String) As String
    IntArrJson = "[]"
    If LenB(Trim$(s)) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(s, UF_SEP)

    Dim acc As String
    Dim i As Long
    Dim one As String
    For i = LBound(parts) To UBound(parts)
        one = Trim$(parts(i))
        If LenB(one) > 0 Then
            If LenB(acc) > 0 Then acc = acc & ","
            If IsNumeric(one) Then
                acc = acc & CStr(CLng(Val(one)))
            Else
                acc = acc & StrJson(one)
            End If
        End If
    Next i
    If LenB(acc) = 0 Then Exit Function
    IntArrJson = "[" & acc & "]"
End Function

' 日本語ラベル -> 機械値。表に無いラベルは**原文のまま**載せる
'   (13章§2.2「表に無いラベルは検証不合格」。黙って直さない)。
Public Function EnumJson(ByVal groupName As String, ByVal labelText As String) As String
    Dim v As String
    v = modUICase.EnumEn(groupName, labelText)
    If LenB(v) = 0 Then v = Trim$(labelText)
    EnumJson = StrJson(v)
End Function

' preventions のセル格納形式 -> オブジェクト配列(13章§2.2)。
'   `measure(M-0012)` の括弧内が M-+4桁のときだけ related_menu_id として切り出す。
Private Function PrevJson(ByVal s As String) As String
    PrevJson = "[]"
    If LenB(Trim$(s)) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(s, UF_SEP)

    Dim acc As String
    Dim i As Long
    Dim one As String
    Dim measure1 As String
    Dim menuId As String
    For i = LBound(parts) To UBound(parts)
        one = Trim$(parts(i))
        If LenB(one) > 0 Then
            SplitPrevention one, measure1, menuId
            If LenB(acc) > 0 Then acc = acc & ","
            acc = acc & "{""measure"":" & StrJson(measure1) & _
                  ",""related_menu_id"":" & StrJson(menuId) & "}"
        End If
    Next i
    If LenB(acc) = 0 Then Exit Function
    PrevJson = "[" & acc & "]"
End Function

Private Sub SplitPrevention(ByVal one As String, ByRef measure1 As String, _
                            ByRef menuId As String)
    measure1 = one
    menuId = vbNullString

    If Right$(one, 1) <> ")" Then Exit Sub

    Dim openPos As Long
    openPos = InStrRev(one, "(")
    If openPos <= 1 Then Exit Sub

    Dim inner As String
    inner = Mid$(one, openPos + 1, Len(one) - openPos - 1)
    If Not IsMenuId(inner) Then Exit Sub

    measure1 = Trim$(Left$(one, openPos - 1))
    menuId = inner
End Sub

' `M-` + 数字4桁ちょうどか(13章§1・§2.2)。
Private Function IsMenuId(ByVal s As String) As Boolean
    If Len(s) <> 6 Then Exit Function
    If Left$(s, 2) <> "M-" Then Exit Function
    Dim i As Long
    Dim c As Long
    For i = 3 To 6
        c = AscW(Mid$(s, i, 1))
        If c < 48 Or c > 57 Then Exit Function
    Next i
    IsMenuId = True
End Function

' ============================================================================
' JSON -> セル値
' ============================================================================
Public Function CellFromJson(ByVal kindText As String, ByVal extraText As String, _
                              ByVal pathText As String, ByVal itemJson As String) As String
    Select Case kindText
    Case "a"
        CellFromJson = PathArrCell(itemJson, pathText)
    Case "ia"
        CellFromJson = PathArrCell(itemJson, pathText)
    Case "e"
        CellFromJson = modUICase.EnumJa(extraText, PathStr(itemJson, pathText))
    Case "p"
        CellFromJson = PrevCell(itemJson)
    Case Else
        CellFromJson = PathStr(itemJson, pathText)
    End Select
End Function

' 入れ子パス("evidence.quote")の値。親が無ければ素のキーとして引く。
Private Function PathStr(ByVal jsonText As String, ByVal pathText As String) As String
    Dim parentName As String
    Dim leafName As String
    SplitPath pathText, parentName, leafName
    If LenB(parentName) = 0 Then
        PathStr = modJsonLite.GetStr(jsonText, leafName)
    Else
        PathStr = modJsonLite.GetStr(SubJson(jsonText, parentName), leafName)
    End If
End Function

' 入れ子パスの配列を「; 」で連結する。空配列は空セル(13章§2.2)。
Private Function PathArrCell(ByVal jsonText As String, ByVal pathText As String) As String
    Dim parentName As String
    Dim leafName As String
    SplitPath pathText, parentName, leafName

    Dim src As String
    If LenB(parentName) = 0 Then
        src = jsonText
    Else
        src = SubJson(jsonText, parentName)
    End If

    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(src, leafName)

    Dim acc As String
    Dim i As Long
    For i = 1 To items.count
        If LenB(acc) > 0 Then acc = acc & UF_SEP
        acc = acc & SepSafe(CStr(items(i)))
    Next i
    PathArrCell = acc
End Function

' 13章§2.2(裁定書9 B5): 要素本文に含まれる「;」を全角「，」へ置換する。
'   **要素1本ごとに、連結する前に**通す(連結後にまとめて置換すると区切りの
'   「; 」まで潰れる)。復元はしない(非可逆。各列の header_note に常設注記)。
Private Function SepSafe(ByVal one As String) As String
    SepSafe = Replace(one, ";", ChrW(&HFF0C&))
End Function

' preventions -> `measure(M-0012); measure2` (13章§2.2)。
Private Function PrevCell(ByVal itemJson As String) As String
    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(itemJson, "preventions")

    Dim acc As String
    Dim i As Long
    Dim one As String
    Dim menuId As String
    For i = 1 To items.count
        one = modJsonLite.GetStr(CStr(items(i)), "measure")
        menuId = modJsonLite.GetStr(CStr(items(i)), "related_menu_id")
        If LenB(menuId) > 0 Then one = one & "(" & menuId & ")"
        If LenB(one) > 0 Then
            If LenB(acc) > 0 Then acc = acc & UF_SEP
            acc = acc & SepSafe(one)
        End If
    Next i
    PrevCell = acc
End Function

' キーの値が入れ子オブジェクトのとき、その本文({...})を切り出す。
'   キーが無い・オブジェクトでないときは ""(modJsonLite は配列とオブジェクトの
'   値に "" を返す契約なので、入れ子を1段だけ降りるための最小の補助)。
'   parentName が空なら元のJSONをそのまま返す。
Public Function SubJson(ByVal jsonText As String, ByVal keyName As String) As String
    SubJson = vbNullString
    If LenB(keyName) = 0 Then
        SubJson = jsonText
        Exit Function
    End If

    Dim needle As String
    needle = """" & keyName & """"

    Dim p As Long
    p = InStr(1, jsonText, needle, vbBinaryCompare)
    If p = 0 Then Exit Function

    p = InStr(p + Len(needle), jsonText, ":", vbBinaryCompare)
    If p = 0 Then Exit Function

    Dim n As Long
    n = Len(jsonText)
    Do While p < n
        p = p + 1
        If Mid$(jsonText, p, 1) <> " " Then Exit Do
    Loop
    If Mid$(jsonText, p, 1) <> "{" Then Exit Function

    Dim depth As Long
    Dim i As Long
    Dim inStr1 As Boolean
    Dim ch As String
    depth = 0
    For i = p To n
        ch = Mid$(jsonText, i, 1)
        If inStr1 Then
            If ch = "\" Then
                i = i + 1
            ElseIf ch = """" Then
                inStr1 = False
            End If
        ElseIf ch = """" Then
            inStr1 = True
        ElseIf ch = "{" Then
            depth = depth + 1
        ElseIf ch = "}" Then
            depth = depth - 1
            If depth = 0 Then
                SubJson = Mid$(jsonText, p, i - p + 1)
                Exit Function
            End If
        End If
    Next i
End Function

' ============================================================================
' 13章§2.2 逆シリアライズ規約2(行順=配列順・順序列の昇順)
' ============================================================================

' 順序列の**列定義上の位置**(0始まり)。持たない表は -1。
'   鍵にする列は 13章§2.2 規約2 が挙げる risk_no / gap_no / story_no / slide_no /
'   seq に、13章§2.13 の emg_no(行順の保持だけに使う順序列)を加えた6つ。
Public Function OrderColOf(ByVal colSpec As String) As Long
    OrderColOf = -1

    Dim specs() As String
    specs = Split(colSpec, ";")

    Dim i As Long
    Dim physName As String
    Dim kindText As String
    Dim extraText As String
    Dim pathText As String
    For i = LBound(specs) To UBound(specs)
        SplitCol specs(i), physName, kindText, extraText, pathText
        Select Case physName
        Case "risk_no", "gap_no", "story_no", "slide_no", "emg_no", "seq"
            OrderColOf = i - LBound(specs)
            Exit Function
        End Select
    Next i
End Function

' 順序列の昇順へ並べ替える(件数が小さいので挿入ソート。同値は元の行順を保つ)。
'   items と keys は同じ添字で対応し、先頭 n 件だけを対象にする。
Public Sub SortItems(ByRef items() As String, ByRef keys() As Long, ByVal n As Long)
    Dim i As Long
    Dim j As Long
    Dim keyNo As Long
    Dim keyText As String

    For i = 1 To n - 1
        keyNo = keys(i)
        keyText = items(i)
        j = i - 1
        Do While j >= 0
            If keys(j) <= keyNo Then Exit Do
            keys(j + 1) = keys(j)
            items(j + 1) = items(j)
            j = j - 1
        Loop
        keys(j + 1) = keyNo
        items(j + 1) = keyText
    Next i
End Sub
