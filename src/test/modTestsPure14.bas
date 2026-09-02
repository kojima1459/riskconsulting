Attribute VB_Name = "modTestsPure14"
Option Explicit

' ============================================================================
' modTestsPure14 - W6.1(裁定書22)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 実装コードではなく 11章v3.2.1(§3.1.1 / §3.2 / §3.3.7)と 13章§2.11・
'   §2.19、docs/08 のコードフェンスの実測だけから期待値を導いた。期待値をあとから
'   実装に合わせて緩めることは禁止(17章§1)。
'
' 対象と根拠:
'   W61A FillTemplate   11本  11章§3.2(M1)。8本の雛形について**全穴埋め・全空の
'                             どちらでも `{{` が0件**であること(8本)＋企業規模の
'                             穴の逐語＋コード未入力の穴の逐語＋
'                             「docs/08 の `{{ }}` 集合 ⊆ 置換辞書」
'   W61B AreaTable      3本   13章§2.11(M4)。HandlerName 2本＋表の分解1本
'   W61C StepText/Anchor 2本  11章§3.1.1(M4)。6文と7本目が無いこと／移動先の表
'   W61D StepFor       10本   11章§3.1.1(M4)。優先順位10行を1行1本で固定
'   W61E FitsInRows     2本   11章§3.3.7(M2)。60行ちょうど / 61行
'   計 28本
'
' グループ単位の失敗隔離: modTestsPure.bas と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

Public Sub RunAll()
    On Error GoTo FA
    T_W61A_FillTemplate
WB:
    On Error GoTo FB
    T_W61B_AreaTable
WC:
    On Error GoTo FC
    T_W61C_StepTextAnchor
WD:
    On Error GoTo FD
    T_W61D_StepFor
WE:
    On Error GoTo FE
    T_W61E_FitsInRows
WDone:
    Exit Sub
FA:
    GroupFail "W61A FillTemplate"
    Resume WB
FB:
    GroupFail "W61B AreaTable"
    Resume WC
FC:
    GroupFail "W61C StepText/StepAnchor"
    Resume WD
FD:
    GroupFail "W61D StepFor"
    Resume WE
FE:
    GroupFail "W61E FitsInRows"
    Resume WDone
End Sub

Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), "期待=[" & want & "] 実際=[" & act & "]"
End Sub

Private Sub ChkL(ByVal nm As String, ByVal act As Long, ByVal want As Long)
    modTestRunner.Check nm, (act = want), "期待=" & CStr(want) & " 実際=" & CStr(act)
End Sub

' ============================================================================
' W61A FillTemplate(11章§3.2・裁定書22 M1)
' ----------------------------------------------------------------------------
' 8本の雛形が使う `{{ }}` は docs/08 の `## D-N.` 直下のコードフェンスの実測
' (2026-09-02)。並びは gd_prompt_01..08 と同じ(D-1 / D-3 / D-9 / D-2 / D-4 /
' D-6 / D-7 / D-8)。本文そのものは焼き込みの値源(使い方タブ)にあり、層(a)は
' **穴の集合だけ**を写して「`{{` が残らない」ことを固定する。
' ============================================================================
Private Function TplOf(ByVal n As Long) As String
    Select Case n
    Case 1
        TplOf = "D-1「{{企業名}}」({{本社所在地・業種}})について調査してください。"
    Case 2
        TplOf = "D-3「{{企業名}}」のリスクの兆候を探してください。"
    Case 3
        TplOf = "D-9「{{企業名}}」({{本社所在地}}・証券コード{{コード}})の調達。"
    Case 4
        TplOf = "D-2「{{企業名}}」が属する{{業界名}}業界について調査してください。"
    Case 5
        TplOf = "D-4 {{業種名}}の{{企業規模}}企業に影響しているマクロ環境。"
    Case 6
        TplOf = "D-6「{{企業名}}」の直近{{前回更新からの期間}}の変化。"
    Case 7
        TplOf = "D-7「{{企業名}}」の対象拠点: {{拠点リスト（名称・住所）}}"
    Case 8
        TplOf = "D-8「{{企業名}}」(証券コード{{コード}}・本社{{本社所在地}})の" & _
                "{{直近決算期}}と公式サイト({{公式ドメイン}})。"
    End Select
End Function

' docs/08 が使う `{{ }}` の全集合(vbLf 区切り。11章§3.2 の差込表と同じ12語)。
Private Function DocPlaceholders() As String
    Dim s As String
    s = s & "{{企業名}}" & vbLf
    s = s & "{{本社所在地・業種}}" & vbLf
    s = s & "{{本社所在地}}" & vbLf
    s = s & "{{業界名}}" & vbLf
    s = s & "{{業種名}}" & vbLf
    s = s & "{{コード}}" & vbLf
    s = s & "{{拠点リスト（名称・住所）}}" & vbLf
    s = s & "{{企業規模}}" & vbLf
    s = s & "{{リスク/課題}}" & vbLf
    s = s & "{{前回更新からの期間}}" & vbLf
    s = s & "{{直近決算期}}" & vbLf
    s = s & "{{公式ドメイン}}"
    DocPlaceholders = s
End Function

Private Function FullFill(ByVal tpl As String) As String
    FullFill = modUIResearch.FillTemplate(tpl, "春華堂", "静岡県浜松市中央区", _
                                          "菓子製造業", "9999", "本社・浜松工場")
End Function

Private Function EmptyFill(ByVal tpl As String) As String
    EmptyFill = modUIResearch.FillTemplate(tpl, "", "", "", "", "")
End Function

Private Sub T_W61A_FillTemplate()
    Dim n As Long
    For n = 1 To 8
        Dim okBoth As Boolean
        Dim full As String
        Dim blank As String
        full = FullFill(TplOf(n))
        blank = EmptyFill(TplOf(n))
        okBoth = (InStr(1, full, "{{", vbBinaryCompare) = 0) And _
                 (InStr(1, blank, "{{", vbBinaryCompare) = 0) And _
                 (LenB(full) > 0) And (LenB(blank) > 0)
        modTestRunner.Check "Test_W61A_0" & CStr(n) & _
            "_雛形" & CStr(n) & "本目は全穴埋めでも全空でも{{が残らない_11章3.2", _
            okBoth, "全穴埋め=[" & full & "] 全空=[" & blank & "]"
    Next n

    ' 9: {{企業規模}} は画面に入力口が無いので常に日本語の穴になる(逐語)。
    ChkS "Test_W61A_09_企業規模は常に日本語の穴_11章3.2", _
        FullFill("[{{企業規模}}]"), _
        "[〔ここに会社の規模（従業員数や売上のめやす）を書いてください〕]"

    ' 10: {{コード}} は未入力のときだけ穴になる(逐語)。入力があればその値。
    Dim okCode As Boolean
    okCode = (EmptyFill("[{{コード}}]") = _
              "[〔ここに証券コードを書いてください（上場していなければ消してください）〕]") _
             And (FullFill("[{{コード}}]") = "[9999]")
    modTestRunner.Check "Test_W61A_10_コードは未入力のときだけ日本語の穴_11章3.2", _
        okCode, "空=[" & EmptyFill("[{{コード}}]") & "] 値あり=[" & _
        FullFill("[{{コード}}]") & "]"

    ' 11: docs/08 の `{{ }}` 集合 ⊆ 置換辞書(11章§3.2 の機械検査)。
    Dim keys As String
    keys = vbLf & modUIResearch.PlaceholderKeys() & vbLf
    Dim want() As String
    want = Split(DocPlaceholders(), vbLf)
    Dim i As Long
    Dim missing As String
    For i = LBound(want) To UBound(want)
        If InStr(1, keys, vbLf & want(i) & vbLf, vbBinaryCompare) = 0 Then
            If LenB(missing) > 0 Then missing = missing & "・"
            missing = missing & want(i)
        End If
    Next i
    modTestRunner.Check _
        "Test_W61A_11_docs08の全プレースホルダが置換辞書に載っている_11章3.2", _
        (LenB(missing) = 0), "辞書に無い語=[" & missing & "]"
End Sub

' ============================================================================
' W61B AreaTable / HandlerName(13章§2.11・裁定書22 M4)
' ============================================================================
Private Sub T_W61B_AreaTable()
    ChkS "Test_W61B_01_HandlerNameは欄キーをパスカル化する_13章2.11", _
        modUICase6.HandlerName("PasteInto", "field_notes"), "PasteIntoFieldNotes"
    ChkS "Test_W61B_02_HandlerNameは2語以上でも各語の頭を大文字にする_13章2.11", _
        modUICase6.HandlerName("ShowArea", "hearing_answers"), "ShowAreaHearingAnswers"

    ' 表の分解: 6欄・並び順・data_key・現場メモはプレビューを持たない。
    Dim keys() As String
    keys = Split(modUICase6.AreaKeys(), vbLf)
    Dim okTable As Boolean
    okTable = ((UBound(keys) - LBound(keys) + 1) = 6)
    okTable = okTable And (keys(LBound(keys)) = "dossier")
    okTable = okTable And (keys(LBound(keys) + 3) = "field_notes")
    okTable = okTable And (modUICase6.AreaField("dossier", 1) = "input_dossier")
    okTable = okTable And (modUICase6.AreaField("dossier", 3) = "ci_prev_dossier")
    okTable = okTable And (LenB(modUICase6.AreaField("field_notes", 3)) = 0)
    okTable = okTable And (modUICase6.AreaField("field_notes", 4) = "ci_area_field_notes")
    modTestRunner.Check "Test_W61B_03_AreaTableは6欄に分解できる_13章2.11", _
        okTable, "keys=[" & modUICase6.AreaKeys() & "]"
End Sub

' ============================================================================
' W61C StepText / StepAnchor(11章§3.1.1・裁定書22 M4)
' ============================================================================
Private Sub T_W61C_StepTextAnchor()
    Dim okText As Boolean
    Dim n As Long
    okText = True
    For n = 1 To 6
        If LenB(modUINav.StepText(n)) = 0 Then okText = False
    Next n
    okText = okText And (LenB(modUINav.StepText(0)) = 0)
    okText = okText And (LenB(modUINav.StepText(7)) = 0)
    modTestRunner.Check "Test_W61C_01_STEPの文は6つだけで7つ目を作らない_11章3.1.1", _
        okText, "1本目=[" & modUINav.StepText(1) & "] 7本目=[" & _
        modUINav.StepText(7) & "]"

    Dim okAnchor As Boolean
    okAnchor = (modUINav.StepAnchor(1) = "nv_sec1") And _
               (modUINav.StepAnchor(2) = "nv_sec1") And _
               (modUINav.StepAnchor(3) = "nv_sec2") And _
               (modUINav.StepAnchor(4) = "nv_sec3") And _
               (modUINav.StepAnchor(5) = "nv_sec4") And _
               (modUINav.StepAnchor(6) = "nv_sec4")
    modTestRunner.Check "Test_W61C_02_次への移動先は11章3.1.1の表のとおり", _
        okAnchor, "1..6=[" & modUINav.StepAnchor(1) & "," & modUINav.StepAnchor(2) & _
        "," & modUINav.StepAnchor(3) & "," & modUINav.StepAnchor(4) & "," & _
        modUINav.StepAnchor(5) & "," & modUINav.StepAnchor(6) & "]"
End Sub

' ============================================================================
' W61D StepFor(11章§3.1.1 の優先順位10行・裁定書22 M4)
' ----------------------------------------------------------------------------
' 1行につき1本。**上の行が下の行より必ず優先される**ことを、下の行にも当たる
' 入力を与えて確かめる(順番を入れ替える変異を1本ずつ捕まえる)。
' 引数: kbReady, locked, hasCompany, anyArea, statusText
' ============================================================================
Private Sub T_W61D_StepFor()
    ' 1: ナレッジが読めていない -> 1/6(他の条件が全部そろっていても勝つ)
    ChkL "Test_W61D_01_ナレッジ未読は最優先でSTEP1_11章3.1.1", _
        modUINav.StepFor(False, True, True, True, "s4_done"), 1
    ' 2: 実行中 -> 4/6(会社名が空でも勝つ)
    ChkL "Test_W61D_02_実行中はSTEP4_11章3.1.1", _
        modUINav.StepFor(True, True, False, False, "exported"), 4
    ' 3: 会社名が空 -> 1/6
    ChkL "Test_W61D_03_会社名が空はSTEP1_11章3.1.1", _
        modUINav.StepFor(True, False, False, True, "s4_done"), 1
    ' 4: 会社名あり・②が全部空 -> 2/6
    ChkL "Test_W61D_04_貼付がゼロならSTEP2_11章3.1.1", _
        modUINav.StepFor(True, False, True, False, "s4_done"), 2
    ' 5: error -> 4/6
    ChkL "Test_W61D_05_errorはSTEP4_11章3.1.1", _
        modUINav.StepFor(True, False, True, True, "error"), 4
    ' 6: draft -> 4/6
    ChkL "Test_W61D_06_draftはSTEP4_11章3.1.1", _
        modUINav.StepFor(True, False, True, True, "draft"), 4
    ' 7: s1_done / s2_done / s3_done -> 4/6
    Dim ok7 As Boolean
    ok7 = (modUINav.StepFor(True, False, True, True, "s1_done") = 4) And _
          (modUINav.StepFor(True, False, True, True, "s2_done") = 4) And _
          (modUINav.StepFor(True, False, True, True, "s3_done") = 4)
    modTestRunner.Check "Test_W61D_07_s1からs3の途中はSTEP4_11章3.1.1", ok7, _
        "s1/s2/s3 のいずれかが4以外"
    ' 8: s4_done -> 5/6
    ChkL "Test_W61D_08_s4_doneはSTEP5_11章3.1.1", _
        modUINav.StepFor(True, False, True, True, "s4_done"), 5
    ' 9: exported -> 6/6
    ChkL "Test_W61D_09_exportedはSTEP6_11章3.1.1", _
        modUINav.StepFor(True, False, True, True, "exported"), 6
    ' 10: feedback_done -> 3/6(訪問の答えを②へ貼らせる)
    ChkL "Test_W61D_10_feedback_doneはSTEP3_11章3.1.1", _
        modUINav.StepFor(True, False, True, True, "feedback_done"), 3
End Sub

' ============================================================================
' W61E FitsInRows(11章§3.3.7・裁定書22 M2)
' ============================================================================
Private Function LinesOf(ByVal n As Long) As String
    Dim i As Long
    Dim acc As String
    For i = 1 To n
        If i > 1 Then acc = acc & vbLf
        acc = acc & "行" & CStr(i)
    Next i
    LinesOf = acc
End Function

Private Sub T_W61E_FitsInRows()
    Dim ok1 As Boolean
    ok1 = modNavText.FitsInRows(LinesOf(60), 60) And _
          modNavText.FitsInRows(LinesOf(1), 60) And _
          modNavText.FitsInRows("", 60)
    modTestRunner.Check "Test_W61E_01_60行ちょうどと空は枠に収まる_11章3.3.7", ok1, _
        "60行/1行/空 のいずれかが False"

    Dim ok2 As Boolean
    ok2 = (Not modNavText.FitsInRows(LinesOf(61), 60)) And _
          (Not modNavText.FitsInRows(LinesOf(1), 0)) And _
          modNavText.FitsInRows(Replace$(LinesOf(60), vbLf, vbCrLf), 60)
    modTestRunner.Check "Test_W61E_02_61行は入らず枠0行はfail-closed_11章3.3.7", ok2, _
        "61行=False / rows=0=False / CrLfも1改行として数える の3点"
End Sub
