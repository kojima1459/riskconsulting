Attribute VB_Name = "modPipeline3"
Option Explicit

' ============================================================================
' modPipeline3 - S1/S2/S3 user の値源組立(15章 v2.6 で増えた3プレースホルダ)
' ----------------------------------------------------------------------------
' 12章§2: modPipeline の分割先(30,000字契約)。W7(17章 T-55・裁定書25)で
'   15章§2 user へ {{financeText}}、§3 user へ {{incidentsText}}、
'   §1.2c へ {{focus_line_ids}} が増えたため、その**値源の解決と1行属性化**を
'   本モジュールが担う(modPipeline は満杯のため新設した)。
'
' 責務:
'   1) 値源の解決(case_data の input_finance / 案件一覧の focus_line_ids /
'      ナレッジの事故事例)。
'   2) 15章§0 原則9 の無害化(SanitizeInput)と、1行属性の改行畳み込み。
'   3) 未提供時の既定文言(15章の各プレースホルダ注記が定める文言)。
'      **例外は事故事例**: 0行のときの文言はKB側(modKnowledgeFmt)が返すので、
'      ここには置かない(T-57 裁定。同じ文言を2箇所に書かない)。
'   4) modPromptsOps.Asm*User への薄い受け渡し(呼出側= modPipeline /
'      modPipeline2 の行を増やさないための包み)。
'
' 本モジュールは 15章の本文を1文字も持たない(本文の正は modPromptsCore /
'   modPromptsBlocks。ここは値だけを組み立てる)。
'
' R4(12章§2): store 経由でシートを読む。テストが叩くのは純関数
'   (FinanceBlockText / IncidentsBlockText / FocusLineIdsAttr)だけである
'   (技術メモ4。他の関数はシートに触れるが実行に到達しない)。
' CP932準拠(15章§0 原則7): 本文・注釈ともに CP932 内の文字だけで書く。
' ============================================================================

' 15章§2 user【決算・財務】の未提供時の値。
Private Const P3_NONE_TEXT As String = "なし"
' 15章§1.2c {{focus_line_ids}} の空時の既定文言。
Private Const P3_FOCUS_NONE As String = "指定なし"

' 直近の原文照合で未照合だった risk_no の一覧(";" 区切り。""=全部照合できた)。
'   **モジュール変数による状態保持**は modPipeline2.LastDeepOutcome と同型で、
'   理由も同じ(RunStep の Boolean 戻り値の契約を変えずに ui・HTML へ渡す口が
'   他に無い)。読む口は LastGroundNote のみ。書くのは GroundHook のみ。
Private mLastGroundNote As String

' 直近のS1で立った警告の集計("V-S1-14:2;V-S1-15:1" 形式。""=指摘なし)。
'   mLastGroundNote と同型・同じ理由(RunStep の戻り値の契約を変えずに HTML へ
'   渡す口が他に無い)。読む口は LastS1Notes のみ。書くのは S1Notes のみ。
Private mLastS1Note As String

' --------------------------------------------------------------------------
' 純関数(15章の各プレースホルダの値づくり)
' --------------------------------------------------------------------------

' {{financeText}} の値。空なら「なし」。SanitizeInput 済みの文字列を受ける
'   前提だが、二重に通しても結果は変わらないためここでも通す(15章§0 原則9)。
Public Function FinanceBlockText(ByVal rawText As String) As String
    Dim t As String

    t = Trim$(modUtilText.SanitizeInput(rawText))
    If LenB(t) = 0 Then t = P3_NONE_TEXT
    FinanceBlockText = t
End Function

' {{incidentsText}} の値。無害化して前後の空白を落とすだけ(15章§0 原則9)。
'   0行のときの文言は modKnowledge.IncidentsFor が返す(T-57 裁定。呼出側は
'   既定文言を持たない)。
Public Function IncidentsBlockText(ByVal rawText As String) As String
    IncidentsBlockText = Trim$(modUtilText.SanitizeInput(rawText))
End Function

' {{focus_line_ids}} の値。1行属性なので改行・タブを空白へ畳む(15章§0 原則9)。
'   空なら「指定なし」。
Public Function FocusLineIdsAttr(ByVal rawIds As String) As String
    Dim t As String

    t = modUtilText.SanitizeInput(rawIds)
    t = Replace(Replace(Replace(Replace(t, vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
    t = Trim$(t)
    If LenB(t) = 0 Then t = P3_FOCUS_NONE
    FocusLineIdsAttr = t
End Function

' --------------------------------------------------------------------------
' MsgE0302 - JSON抽出失敗のエラー文言(裁定書37 C-6 A-3)。単一情報源化。
' --------------------------------------------------------------------------
'   modPipeline.bas / modPipeline2.bas / modPlayOps.bas の3箇所に手書き複製
'   されていた同一文言(伝書鳩20260912 付録A-3)をここへ集約する。文言は不変。
Public Function MsgE0302() As String
    MsgE0302 = "[E0302] 応答からJSONを抽出できませんでした（説明文のみ・括弧の欠落など）"
End Function

' --------------------------------------------------------------------------
' IncidentsFor - 事故事例(13章§3.11)の1行整形テキストの**呼び口1本**
' --------------------------------------------------------------------------
'   実装は班C(17章 T-56 ②)の modKnowledge.IncidentsFor(業種別抽出 ->
'   modKnowledgeFmt.FmtIncidents)。T-57 でスタブから差し替えた(統合)。
'   0行のときの文言(15章§3「(この業種の登録事例はまだありません)」)は
'   **KB側が返す**。呼出側は既定文言を持たない(答えを2箇所に書かない)。
Public Function IncidentsFor(ByVal industryCode As String, ByVal maxRows As Long) As String
    IncidentsFor = modKnowledge.IncidentsFor(industryCode, maxRows)
End Function

' --------------------------------------------------------------------------
' 値源の解決(シート・configを読む)
' --------------------------------------------------------------------------

' case_data の input_finance(13章§2.11(a) 7本目の貼付欄)。
Public Function FinanceTextOf(ByVal caseId As String) As String
    FinanceTextOf = FinanceBlockText(modCaseStore.LoadData(caseId, "input_finance"))
End Function

' 案件一覧の focus_line_ids(13章§2.1。班C が列を足すまでは空= 指定なし)。
Public Function FocusIdsOf(ByVal caseId As String) As String
    FocusIdsOf = FocusLineIdsAttr(modCaseRead.CaseColumnOf(caseId, "focus_line_ids"))
End Function

' 案件一覧の round_no(読めないときは1)。
Public Function RoundNoOf(ByVal caseId As String) As Long
    Dim n As Long

    n = 0
    On Error Resume Next
    n = CLng(Val(modCaseRead.CaseColumnOf(caseId, "round_no")))
    On Error GoTo 0
    If n < 1 Then n = 1
    RoundNoOf = n
End Function

' --------------------------------------------------------------------------
' Asm*User への薄い包み(呼出側の行を増やさないための入口)
' --------------------------------------------------------------------------

' 15章§2 user。pasted は modPipeline の PL_S1_KEYS 順の9欄。
Public Function S1UserText(ByRef ctx As TCaseCtx, ByVal caseId As String, _
                           ByRef pasted() As String) As String
    S1UserText = modPromptsOps.AsmS1User(ctx, pasted(0), pasted(1), pasted(2), pasted(3), _
                                         pasted(4), pasted(5), pasted(6), pasted(7), pasted(8), _
                                         FinanceTextOf(caseId))
End Function

' 15章§3 user。第2ラウンドの絞り込みをここで解決する。{{incidentsText}} は
'   15章§0.7 の切詰め(6段の順2)を通った値を呼出側から受ける(T-57)。
Public Function S2UserText(ByRef ctx As TCaseCtx, ByVal caseId As String, _
                           ByVal s1Json As String, ByVal riskLib As String, _
                           ByVal menus As String, ByVal prevS2Json As String, _
                           ByVal hearingAnswers As String, _
                           ByVal incidents As String) As String
    S2UserText = modPromptsOps.AsmS2User(ctx, s1Json, riskLib, menus, prevS2Json, _
                                         hearingAnswers, IncidentsBlockText(incidents), _
                                         FocusIdsOf(caseId), RoundNoOf(caseId))
End Function

' 15章§4 user。第2ラウンドの絞り込みをここで解決する。
Public Function S3UserText(ByRef ctx As TCaseCtx, ByVal caseId As String, _
                           ByVal s1Summary As String, ByVal s2Json As String, _
                           ByVal menus As String, ByVal lines As String, _
                           ByVal schemes As String, ByVal cases As String) As String
    S3UserText = modPromptsOps.AsmS3User(ctx, s1Summary, s2Json, menus, lines, schemes, _
                                         cases, FocusIdsOf(caseId), RoundNoOf(caseId))
End Function

' ============================================================================
' 裁定書37 B-03 / B-05: 原文照合(modGround)と充足度の run_log 記録
' ----------------------------------------------------------------------------
' modPipeline は30,000字契約でほぼ満杯のため、値源の組立(BuildHaystack)・
'   config の読み・注記の保持(LastGroundNote)をここへ置き、modPipeline からは
'   DefendNotes の**1行**で呼ぶ(B班報告 §3 B-03 の「呼び出し点は1箇所」)。
' **落とさない・修復リトライを起こさない**。run_log の detail に印を残すだけで、
'   検証の戻り値(modValidate の *Core の結果)には一切触れない。
' ============================================================================

' DefendNotes - modPipeline.OneCall が Defend の**直後**に1行で呼ぶ入口。
'   validateOk=False(検証不合格)のときは何もしない(捨てられる出力を測らない)。
'   stepNo=1 -> 充足度(B-05) / stepNo=2 -> 原文照合(B-03)。
Public Sub DefendNotes(ByVal stepNo As Long, ByVal caseId As String, _
                       ByVal stepJson As String, ByVal s1Json As String, _
                       ByVal validateOk As Boolean, ByRef detailAcc As String)
    On Error Resume Next
    If Not validateOk Then Exit Sub
    If stepNo = 1 Then
        P3AddNote detailAcc, SufficiencyNoteOf(stepJson)
        S1Notes caseId, stepJson, detailAcc
        S1Snapshot caseId, stepJson, detailAcc
    ElseIf stepNo = 2 Then
        GroundHook caseId, stepJson, s1Json, detailAcc
    End If
End Sub

' GroundHook - S2の出力を貼付原文と突き合わせ、注記だけを残す(裁定書37 B-03)。
'   config ground_check=FALSE なら何もしない。haystack が空なら検査せず
'   `ground_skipped` を注記する(fail-open。「貼付が空のときに全件未照合で
'   埋めない」= B班テスト観点(5))。
Public Sub GroundHook(ByVal caseId As String, ByVal s2Json As String, _
                      ByVal s1Json As String, ByRef detailAcc As String)
    Dim hay As String, note As String

    mLastGroundNote = vbNullString
    If Not modConfig.GetBool("ground_check", True) Then Exit Sub

    hay = BuildHaystack(caseId, s1Json)
    If LenB(Trim$(hay)) = 0 Then
        P3AddNote detailAcc, "ground_skipped"
        Exit Sub
    End If

    note = modGround.GroundNotes(s2Json, hay, _
                                 modConfig.GetLong("ground_head_chars", modGround.GR_HEAD_DEFAULT))
    mLastGroundNote = note
    ' 0件でも必ず記録する(「検査した」と「検査していない」を区別するため)。
    P3AddNote detailAcc, "ground_unmatched=" & CStr(modGround.NoteCount(note))
End Sub

' LastGroundNote - 直近の未照合 risk_no 一覧(";" 区切り)。
Public Function LastGroundNote() As String
    LastGroundNote = mLastGroundNote
End Function

' ResetGroundNote - 明示リセット口(modPipeline2.ResetDeepOutcome と同じ考え方)。
Public Sub ResetGroundNote()
    mLastGroundNote = vbNullString
End Sub

' BuildHaystack - 照合される「原文」。case_data の input_* 全欄(13章§2.2 の
'   data_key のうち接頭辞 input_ のもの。値源は modCaseStore3.DataKeys 1箇所)と
'   S1の出力JSONを連結する。S1を混ぜるのは、S2の引用が「S1が構造化した値」を
'   写しているのが正常な経路だからである(B班報告 §1)。
Public Function BuildHaystack(ByVal caseId As String, ByVal s1Json As String) As String
    Dim it As Variant, keyText As String, sb As String

    For Each it In Split(modCaseStore3.DataKeys(), ";")
        keyText = Trim$(CStr(it))
        If Left$(keyText, 6) = "input_" Then
            sb = sb & modCaseStore.LoadData(caseId, keyText) & vbLf
        End If
    Next it
    BuildHaystack = sb & s1Json
End Function

' SufficiencyNoteOf - S1の input_quality を run_log 用の1語へ(裁定書37 B-05)。
'   "iq=low;miss=7" の形。overall は 19章§3 の3値、miss は coverage[] のうち
'   status<>ok の観点数。**読めなければ "iq=?"**(黙って mid にしない)。
'   HTML側の警告表示(班3)もこの1本を呼ぶので Public にする(14章§6)。
Public Function SufficiencyNoteOf(ByVal s1Json As String) As String
    Dim ov As String, st As String, it As Variant, n As Long

    ov = LCase$(Trim$(modJsonLite.GetStr(s1Json, "overall")))
    If ov <> "high" And ov <> "mid" And ov <> "low" Then
        SufficiencyNoteOf = "iq=?"
        Exit Function
    End If

    For Each it In modJsonLite.GetArrayItems(s1Json, "coverage")
        st = LCase$(Trim$(modJsonLite.GetStr(CStr(it), "status")))
        If LenB(st) > 0 And st <> "ok" Then n = n + 1
    Next it
    SufficiencyNoteOf = "iq=" & ov & ";miss=" & CStr(n)
End Function

' ============================================================================
' 裁定書38 班A: S1の警告(V-S1-14 / V-S1-15)と、S1再実行時の揺れ(B-14)
' ----------------------------------------------------------------------------
' どちらも**落とさない**。run_log の detail に印を残すだけで、検証の戻り値には
'   一切触れない(modValidate3 の冒頭注釈・15章§2 の注記)。
' ============================================================================

' S1Notes - 15章 V-S1-14 / V-S1-15 を測り、detail へ "s1_warn=..." を足す。
'   照合する原文は**貼付原文だけ**(BuildHaystack の第2引数に s1Json を渡さない。
'   S1の出力を混ぜると捏造URLが自分自身と一致してしまう)。
'   0件のときは何も足さない(注記が増え続けるのを避ける)。
Public Sub S1Notes(ByVal caseId As String, ByVal s1Json As String, _
                   ByRef detailAcc As String)
    Dim note As String

    mLastS1Note = vbNullString
    note = modValidate3.WarnNoteOf(modValidate3.CheckS1Notes(s1Json, BuildHaystack(caseId, vbNullString)))
    mLastS1Note = note
    If LenB(note) > 0 Then P3AddNote detailAcc, "s1_warn=" & note
End Sub

' LastS1Notes - 直近のS1警告の集計(LastGroundNote と同型)。
Public Function LastS1Notes() As String
    LastS1Notes = mLastS1Note
End Function

' ResetS1Notes - 明示リセット口(ResetGroundNote と同じ考え方)。
Public Sub ResetS1Notes()
    mLastS1Note = vbNullString
End Sub

' S1Snapshot - 裁定書38 B-14。**s1_json を上書きする前**に前回分を
'   s1_json_prev(13章§2.2)へ退避し、主要8フィールドの差分件数を
'   detail へ "s1_diff=n" として残す。前回が無ければ何もしない
'   (初回実行を「差分0」と記録すると、揺れが無かったのと区別できなくなる)。
Public Sub S1Snapshot(ByVal caseId As String, ByVal s1Json As String, _
                      ByRef detailAcc As String)
    Dim prevJson As String

    prevJson = modCaseStore.LoadData(caseId, "s1_json")
    If LenB(Trim$(prevJson)) = 0 Then Exit Sub
    modCaseStore.SaveData caseId, "s1_json_prev", prevJson
    P3AddNote detailAcc, "s1_diff=" & CStr(S1DiffCount(prevJson, s1Json))
End Sub

' S1DiffCount - 主要8フィールドのうち値が変わった数(0から8)。純関数。
'   1 company_name / 2 business_summary / 3 strategy_outlook(mvv・市況・
'   aspirations件数) / 4 locations件数 / 5 financials.sales / 6 current_coverage
'   件数 / 7 missing_info件数 / 8 input_quality.overall。
'   **どちらかが空なら 0**(比較していないことを「差分なし」と同じ 0 で表すが、
'   呼出側 S1Snapshot は前回が空のときそもそも記録しない)。
Public Function S1DiffCount(ByVal prevJson As String, ByVal curJson As String) As Long
    Dim i As Long, n As Long
    Dim a(1 To 8) As String, b(1 To 8) As String

    If LenB(Trim$(prevJson)) = 0 Or LenB(Trim$(curJson)) = 0 Then Exit Function
    S1Fields prevJson, a
    S1Fields curJson, b
    For i = 1 To 8
        If a(i) <> b(i) Then n = n + 1
    Next i
    S1DiffCount = n
End Function

' S1DiffCount の比較値を作る(8要素。取り出せない値は空文字のまま比較する)。
Private Sub S1Fields(ByVal s1Json As String, ByRef out() As String)
    out(1) = Trim$(modJsonLite.GetStr(s1Json, "company_name"))
    out(2) = Trim$(modJsonLite.GetStr(s1Json, "business_summary"))
    out(3) = Trim$(modJsonLite.GetStr(s1Json, "mvv")) & vbTab & _
             Trim$(modJsonLite.GetStr(s1Json, "market_context")) & vbTab & _
             CStr(modJsonLite.GetArrayItems(s1Json, "aspirations").Count)
    out(4) = CStr(modJsonLite.GetArrayItems(s1Json, "locations").Count)
    out(5) = Trim$(modJsonLite.GetStr(s1Json, "sales"))
    out(6) = CStr(modJsonLite.GetArrayItems(s1Json, "current_coverage").Count)
    out(7) = CStr(modJsonLite.GetArrayItems(s1Json, "missing_info").Count)
    out(8) = Trim$(modJsonLite.GetStr(s1Json, "overall"))
End Sub

' run_log detail の積み上げ(modPipeline.AddNote と同じ規約=";" 区切り)。
Private Sub P3AddNote(ByRef acc As String, ByVal noteText As String)
    If LenB(noteText) = 0 Then Exit Sub
    If LenB(acc) > 0 Then acc = acc & ";"
    acc = acc & noteText
End Sub
