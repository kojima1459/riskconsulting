Attribute VB_Name = "modPipeline4"
Option Explicit

' ============================================================================
' modPipeline4 - 15章§0.7 ナレッジ側の切詰め(計画と適用)
' ----------------------------------------------------------------------------
' 12章§2: modPipeline の分割先(30,000字契約)。W7(17章 T-57・裁定書25 S6)で
'   §0.7 の切詰め表が5段から**6段**(事故事例 incidentsText が順2に入る)へ
'   増えたため、TrimPlan(計画)と LoadKbSlots(適用)を本モジュールへ集めた。
'   modPipeline は29,970字で満杯であり、同じ手順が modPipeline / modPipeline2 の
'   2箇所に写経されていたので、この機会に**唯一の実装**へ畳んだ。
'
' 切詰め順(15章§0.7 の表。スロット番号=表の「順」-1):
'   0 成功事例 casesText    (config kb_case_rows・既定5)     下限0行
'   1 事故事例 incidentsText(config kb_incident_rows・既定5) 下限0行
'   2 型ライブラリ schemesText(kb_scheme_rows・既定10)      下限0行
'   3 メニュー menusText    (kb_menu_rows・既定60)           下限5行
'   4 種目 linesText        (全行)                            下限5行
'   5 リスクライブラリ riskLibText(kb_risk_rows・既定20)    下限5行
'
' 責務の分界(答えを2箇所に書かない):
'   行数の数え方 = modPipeline.KbRowCount / Step別の枠 = modPipeline.UsesSlot /
'   予算配分(上限の3割) = modPipeline.BudgetOf。いずれも呼ぶだけで持たない。
'   1行整形の本体は modKnowledge / modKnowledgeFmt。
'
' R4(12章§2): シートには modKnowledge 経由でしか触れない(R4許可は与えない)。
' CP932準拠(15章§0 原則7): 本文・注釈ともに CP932 内の文字だけで書く。
' ============================================================================

' 切詰めスロット数(15章§0.7 の6段)。
Private Const P4_PLAN_N As Long = 6
' スロット別の下限行数(同表「下限」列)。メニュー・種目はS3のID実在制約、
'   リスクライブラリは§0.5 第2層のため5行を割らない。
Private Const P4_FLOORS As String = "0;0;0;5;5;5"
Private Const P4_SEMI As String = ";"
' run_log の detail へ書く対象名(§0.7「記録」)。
Private Const P4_KB_LABELS As String = "cases|incidents|schemes|menus|lines|risklib"
' 0行に切り詰めたスロットの代替文言(15章§6.1)。
Private Const P4_KB_ZERO As String = "なし"

' ============================================================================
' TrimPlan - 15章§0.7 ナレッジ側の切詰めを**計画するだけ**の純関数。
'   counts: 12要素。前半(0..5)=現在の行数、後半(6..11)=現在の文字数。並びは
'     切詰め順 成功事例→事故事例→型→メニュー→種目→リスクライブラリ。
'     6要素以下のときは文字数0とみなす(=切詰め不要)。
'   budgetChars: ナレッジ注入に許される合計文字数(§0.7「上限の3割」)。0以下は
'     上限なしとして現在の行数をそのまま返す。
'   戻り値: 6要素の「注入してよい行数」。1段ずつ順に適用し、そのつど総量を
'     再計算して budgetChars 以下になった時点で止める(1対象あたり半減は1回)。
'     半減は端数切上げ、下限は 0/0/0/5/5/5 行。文字数は行数に比例と見積もる。
'   (T-57 で modKnowledgeFmt から移設。5段→6段。)
' ============================================================================
Public Function TrimPlan(ByRef counts() As Long, ByVal budgetChars As Long) As Long()
    Dim res() As Long
    ReDim res(0 To P4_PLAN_N - 1)

    Dim rowsNow(0 To P4_PLAN_N - 1) As Long
    Dim charsNow(0 To P4_PLAN_N - 1) As Double
    Dim floors As Variant
    floors = Split(P4_FLOORS, P4_SEMI)

    Dim i As Long
    Dim total As Double
    For i = 0 To P4_PLAN_N - 1
        rowsNow(i) = LongAt(counts, i)
        charsNow(i) = CDbl(LongAt(counts, P4_PLAN_N + i))
        res(i) = rowsNow(i)
        total = total + charsNow(i)
    Next i

    If budgetChars <= 0 Then
        TrimPlan = res
        Exit Function
    End If

    Dim newRows As Long
    Dim newChars As Double
    For i = 0 To P4_PLAN_N - 1
        If total <= CDbl(budgetChars) Then Exit For
        newRows = CLng(Fix((rowsNow(i) + 1) / 2))
        If newRows < CLng(Val(CStr(floors(i)))) Then newRows = CLng(Val(CStr(floors(i))))
        If newRows > rowsNow(i) Then newRows = rowsNow(i)
        newChars = 0
        If rowsNow(i) > 0 Then newChars = charsNow(i) * CDbl(newRows) / CDbl(rowsNow(i))
        total = total - charsNow(i) + newChars
        charsNow(i) = newChars
        rowsNow(i) = newRows
        res(i) = newRows
    Next i
    TrimPlan = res
End Function

' ============================================================================
' LoadKbSlots - ナレッジ注入と15章§0.7の切詰めの**適用**(唯一の実装)。
'   txt は本関数が ReDim(0 To 5)して埋める。切り詰めたら注入IDを積み直す
'   (LastInjectedIds は実際に注入したIDのみ。§0.7「記録」)。
'   detailAcc へ "truncated:<対象>=<削った行数>" を積む(黙って削らない)。
'   caseText(裁定書38 B-10): 業種完全一致の該当が上限に満たないときの並べ替え
'   補充に使う案件本文(modPipeline.CaseTextFor)。空なら補充なし。
'   同じ detailAcc へ "kb_cut:cases=U/T;…"(0.7の切詰め**前**・完全一致の
'   該当総数)も、自然装填(config既定行数)を取った直後に1回だけ積む。
' ============================================================================
Public Sub LoadKbSlots(ByRef ctx As TCaseCtx, ByVal stepNo As Long, _
                       ByVal budgetChars As Long, ByRef txt() As String, _
                       ByVal caseText As String, ByRef detailAcc As String)
    Dim counts(0 To 2 * P4_PLAN_N - 1) As Long
    Dim plan As Variant
    Dim i As Long
    Dim trimmed As Boolean

    ReDim txt(0 To P4_PLAN_N - 1)
    modKnowledge.ResetInjectedIds
    FetchKb ctx, stepNo, txt, -1, 0, caseText
    AddNote detailAcc, modKnowledge.LastKbCutNote()
    AddNote detailAcc, modKnowledge.IncidentsFallbackNote() ' 裁定書47 G-5
    For i = 0 To P4_PLAN_N - 1
        counts(i) = modPipeline.KbRowCount(txt(i))
        counts(P4_PLAN_N + i) = Len(txt(i))
    Next i

    plan = TrimPlan(counts, budgetChars)
    For i = 0 To P4_PLAN_N - 1
        If plan(i) < counts(i) Then trimmed = True
    Next i
    If Not trimmed Then Exit Sub

    modKnowledge.ResetInjectedIds
    For i = 0 To P4_PLAN_N - 1
        If counts(i) > 0 Then FetchKb ctx, stepNo, txt, i, plan(i), caseText
        If plan(i) < counts(i) Then
            AddNote detailAcc, "truncated:" & PickAt(P4_KB_LABELS, i) & _
                               "=" & CStr(counts(i) - plan(i))
        End If
    Next i
End Sub

' slotIdx=-1 は全スロットを既定行数(config)で、0以上はそのスロットだけ maxRows 行
'   (0なら既定文言)で取り直す。メニューはStepで別物(12章§3)。
'   caseText(裁定書38 B-10)は成功事例(i=0)の並べ替え補充にだけ使う。
Private Sub FetchKb(ByRef ctx As TCaseCtx, ByVal stepNo As Long, ByRef txt() As String, _
                    ByVal slotIdx As Long, ByVal maxRows As Long, _
                    Optional ByVal caseText As String = vbNullString)
    Dim i As Long, n As Long

    For i = 0 To P4_PLAN_N - 1
        If (slotIdx < 0 Or slotIdx = i) And modPipeline.UsesSlot(stepNo, i) Then
            n = 0
            If slotIdx >= 0 Then n = maxRows
            If slotIdx >= 0 And maxRows <= 0 Then
                txt(i) = P4_KB_ZERO
            ElseIf i = 0 Then
                txt(i) = modKnowledge.CasesFor(ctx.industry_code, n, caseText)
            ElseIf i = 1 Then
                txt(i) = modPipeline3.IncidentsFor(ctx.industry_code, n)
            ElseIf i = 2 Then
                txt(i) = modKnowledge.SchemesFor(ctx.industry_code, n)
            ElseIf i = 3 And stepNo = 2 Then
                txt(i) = modKnowledge.MenusSummaryFor(ctx.industry_code, n)
            ElseIf i = 3 Then
                txt(i) = modKnowledge.MenusFor(ctx.industry_code, n)
            ElseIf i = 4 Then
                txt(i) = modKnowledge.LinesText(n)
            Else
                txt(i) = modKnowledge.RiskLibFor(ctx.industry_code, n)
            End If
        End If
    Next i
End Sub

' === 内部(すべて純関数) ===

' counts の実体が何要素でも落ちない読み取り。範囲外は0(=切詰め不要)。
Private Function LongAt(ByRef arr() As Long, ByVal i As Long) As Long
    Dim lo As Long, hi As Long

    On Error GoTo NG
    lo = LBound(arr)
    hi = UBound(arr)
    If lo + i > hi Then GoTo NG
    LongAt = arr(lo + i)
    Exit Function
NG:
    LongAt = 0
End Function

Private Function PickAt(ByVal listText As String, ByVal idx As Long) As String
    Dim parts As Variant

    parts = Split(listText, "|")
    If idx < 0 Or idx > UBound(parts) Then Exit Function
    PickAt = parts(idx)
End Function

Private Sub AddNote(ByRef acc As String, ByVal noteText As String)
    If LenB(noteText) = 0 Then Exit Sub
    If LenB(acc) > 0 Then acc = acc & ";"
    acc = acc & noteText
End Sub
