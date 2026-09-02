Attribute VB_Name = "modTestsPure16"
Option Explicit

' ============================================================================
' modTestsPure16 - W7(裁定書25 センターピン整合)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 実装コードではなく 13章v2.6(§2.1 / §2.11(d) / §3.11)と 15章v2.6
'   (§3 の整形例)だけから期待値を導いた。**期待値をあとから実装に合わせて
'   緩めることは禁止**(17章§1)。整形例は15章§3のフェンスの逐語である。
'
' 対象と根拠:
'   W7D FmtIncidents      4本  13章§3.11・15章§3 の整形例(1行の書式 / 空欄の
'                              項目を丸ごと省く / 0行の既定文言 / 非配列でも続行)
'   W7E focus_line_ids    8本  13章§2.1(NormalizeStoryNos 4本＝正常・空白と重複・
'                              非整数の除去・空／CollectLineIds 4本＝採用storyだけ・
'                              重複の統合・該当なし・空の採用番号)
'   W7F CoverageNoteOf    4本  13章§2.11(d)(節の本文だけを返す / 節が空なら空文字 /
'                              例文行は落とす / 見出しが無ければ空文字)
'   計 16本
'
' 末尾から modTestsPure17.RunAll(W7・17章 T-57 統合の12本)を呼ぶ。
'
' グループ単位の失敗隔離: modTestsPure.bas と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

Public Sub RunAll()
    On Error GoTo FA
    T_W7D_FmtIncidents
WB:
    On Error GoTo FB
    T_W7E_FocusLineIds
WC:
    On Error GoTo FC
    T_W7F_CoverageNoteOf
WD:
    On Error GoTo FD
    modTestsPure17.RunAll
WDone:
    Exit Sub
FA:
    GroupFail "W7D FmtIncidents"
    Resume WB
FB:
    GroupFail "W7E focus_line_ids"
    Resume WC
FC:
    GroupFail "W7F CoverageNoteOf"
    Resume WD
FD:
    GroupFail "modTestsPure17.RunAll"
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

' 13章§3.11 の事故事例シートの列(物理順)。
Private Function HdrInc() As String
    HdrInc = "inc_id" & vbTab & "industry_code" & vbTab & "category" & vbTab & _
             "headline" & vbTab & "cause" & vbTab & "loss_scale" & vbTab & _
             "lesson" & vbTab & "source"
End Function

' 見出し行＋データ行(vbLf区切り・セルは vbTab 区切り)から2次元配列を作る。
Private Function MkRows(ByVal hdrText As String, ByVal dataLines As String) As Variant
    Dim h() As String
    Dim dl() As String
    Dim f() As String
    Dim a() As Variant
    Dim nCol As Long
    Dim nRow As Long
    Dim r As Long
    Dim c As Long

    h = Split(hdrText, vbTab)
    nCol = UBound(h) + 1

    If Len(dataLines) = 0 Then
        ReDim a(1 To 1, 1 To nCol)
        For c = 1 To nCol
            a(1, c) = h(c - 1)
        Next c
        MkRows = a
        Exit Function
    End If

    dl = Split(dataLines, vbLf)
    nRow = UBound(dl) + 2
    ReDim a(1 To nRow, 1 To nCol)
    For c = 1 To nCol
        a(1, c) = h(c - 1)
    Next c
    For r = 2 To nRow
        f = Split(dl(r - 2), vbTab)
        For c = 1 To nCol
            If c - 1 <= UBound(f) Then a(r, c) = f(c - 1)
        Next c
    Next r
    MkRows = a
End Function

' ============================================================================
' W7D FmtIncidents(13章§3.11・15章§3 の整形例)
' ----------------------------------------------------------------------------
'   期待値は15章§3の整形例3行目の書式そのもの:
'     [IC-09-004] カテゴリ:… 見出し:… 原因:… 損害規模:… 教訓:… 出所:…
'   空欄の項目は**その部分ごと省略**する(空の項目名を出さない。§3.11)。
' ============================================================================
Private Sub T_W7D_FmtIncidents()
    Dim rws As Variant
    Dim d1 As String
    Dim e1 As String

    d1 = "IC-09-004" & vbTab & "09" & vbTab & "manufacturing_quality" & vbTab & _
         "菓子工場でのアレルゲン表示誤りによる自主回収" & vbTab & _
         "包装資材の切替時に旧版の表示フィルムが混在した" & vbTab & _
         "回収・廃棄で約8,000万円" & vbTab & _
         "包装資材の切替は現物照合を2名で行う" & vbTab & "公表回収告知"
    e1 = "[IC-09-004] カテゴリ:manufacturing_quality " & _
         "見出し:菓子工場でのアレルゲン表示誤りによる自主回収 " & _
         "原因:包装資材の切替時に旧版の表示フィルムが混在した " & _
         "損害規模:回収・廃棄で約8,000万円 " & _
         "教訓:包装資材の切替は現物照合を2名で行う 出所:公表回収告知"
    rws = MkRows(HdrInc(), d1)
    ChkS "Test_W7D_01_FmtIncidentsの1行の書式_15章§3", _
        modKnowledgeFmt.FmtIncidents(rws), e1

    ' loss_scale が空欄の行は「損害規模:」ごと落とす(§3.11 の空項目省略)。
    d1 = "IC-09-005" & vbTab & "09" & vbTab & "facility_bcp" & vbTab & _
         "浸水で製造ラインが停止" & vbTab & "受変電設備が冠水した" & vbTab & _
         "" & vbTab & "受変電設備の設置階を確認する" & vbTab & "業界紙報道"
    e1 = "[IC-09-005] カテゴリ:facility_bcp 見出し:浸水で製造ラインが停止 " & _
         "原因:受変電設備が冠水した 教訓:受変電設備の設置階を確認する 出所:業界紙報道"
    rws = MkRows(HdrInc(), d1)
    ChkS "Test_W7D_02_FmtIncidentsは空欄の項目を丸ごと省く_13章§3.11", _
        modKnowledgeFmt.FmtIncidents(rws), e1

    rws = MkRows(HdrInc(), "")
    ChkS "Test_W7D_03_FmtIncidentsの0行は業種専用の文言_15章§3", _
        modKnowledgeFmt.FmtIncidents(rws), "(この業種の登録事例はまだありません)"

    ChkS "Test_W7D_04_FmtIncidentsは配列でない入力でも既定文言で続行_16章E09", _
        modKnowledgeFmt.FmtIncidents(Empty), "(この業種の登録事例はまだありません)"
End Sub

' ============================================================================
' W7E focus_line_ids の導出(13章§2.1・裁定書25 S4)
' ----------------------------------------------------------------------------
'   used_proposals(`1;3`) -> adopted_story_nos -> 採用storyの line_ids。
'   実在チェック(ExistingLineIds)はナレッジを読むので層(a)では扱わない。
' ============================================================================
Private Sub T_W7E_FocusLineIds()
    ChkS "Test_W7E_01_NormalizeStoryNosは13章の例1;3をそのまま通す_13章§2.1", _
        modCaseStore3.NormalizeStoryNos("1;3"), "1;3"
    ChkS "Test_W7E_02_NormalizeStoryNosは前後の空白と重複を落とす_13章§2.1", _
        modCaseStore3.NormalizeStoryNos(" 2 ; 1 ;2"), "2;1"
    ChkS "Test_W7E_03_NormalizeStoryNosは整数でない要素を捨てる_13章§2.1", _
        modCaseStore3.NormalizeStoryNos("1;採用;3番;0;2"), "1;2"
    ChkS "Test_W7E_04_NormalizeStoryNosは空入力で空文字_13章§2.1", _
        modCaseStore3.NormalizeStoryNos(""), ""

    Dim s3 As String
    s3 = "{""stories"":[" & _
         "{""story_no"":1,""line_ids"":[""L-03"",""L-07""]}," & _
         "{""story_no"":2,""line_ids"":[""L-11""]}," & _
         "{""story_no"":3,""line_ids"":[""L-07"",""L-21""]}]}"

    ChkS "Test_W7E_05_CollectLineIdsは採用storyの種目だけを集める_13章§2.1", _
        modCaseStore3.CollectLineIds(s3, "1"), "L-03;L-07"
    ChkS "Test_W7E_06_CollectLineIdsは複数storyの重複を1本に畳む_13章§2.1", _
        modCaseStore3.CollectLineIds(s3, "1;3"), "L-03;L-07;L-21"
    ChkS "Test_W7E_07_CollectLineIdsは採用番号が空なら空文字_13章§2.1", _
        modCaseStore3.CollectLineIds(s3, ""), ""
    ChkS "Test_W7E_08_CollectLineIdsは該当storyが無ければ空文字_13章§2.1", _
        modCaseStore3.CollectLineIds(s3, "9"), ""
End Sub

' ============================================================================
' W7F CoverageNoteOf(13章§2.11(d)・裁定書25 S1 の二重保存の派生側)
' ============================================================================
Private Sub T_W7F_CoverageNoteOf()
    Dim memo As String
    memo = "【営業メモ】" & vbLf & "社長はワンマン。" & vbLf & _
           "【前回更新メモ】" & vbLf & "地震は見送り。" & vbLf & _
           "【付保の見立て】" & vbLf & "たぶん火災は他社さん。" & vbLf & _
           "労災上乗せは元請の包括に乗っている模様。"
    ChkS "Test_W7F_01_CoverageNoteOfは付保の見立ての本文だけを返す_13章§2.11d", _
        modNavText.CoverageNoteOf(memo), _
        "たぶん火災は他社さん。" & vbLf & "労災上乗せは元請の包括に乗っている模様。"

    memo = "【営業メモ】" & vbLf & "社長はワンマン。" & vbLf & _
           "【前回更新メモ】" & vbLf & "【付保の見立て】" & vbLf & "【そのほか】" & vbLf & _
           "裏の川が気になる。"
    ChkS "Test_W7F_02_CoverageNoteOfは節が空なら空文字_13章§2.11d", _
        modNavText.CoverageNoteOf(memo), ""

    memo = "【付保の見立て】" & vbLf & _
           "例: たぶん火災は他社さん。労災上乗せは元請の包括に乗っている模様。"
    ChkS "Test_W7F_03_CoverageNoteOfは先置きの例文行を落とす_11章§3.3.3b", _
        modNavText.CoverageNoteOf(memo), ""

    ChkS "Test_W7F_04_CoverageNoteOfは見出しが無ければ空文字_13章§2.11d", _
        modNavText.CoverageNoteOf("なんでも書いた文章。"), ""
End Sub
