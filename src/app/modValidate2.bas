Attribute VB_Name = "modValidate2"
Option Explicit

' ============================================================================
' modValidate2 - modValidate の続き(30,000字契約による分割。T-22)
' ----------------------------------------------------------------------------
' 正: modValidate の冒頭コメントに同じ(15章§11 / 各Check節の検証ルール表 /
'   15章§0 原則10 / 19章§3 / 14章§6 / 16章 E-49)。
'
' 分割の理由(12章§2 の30,000字契約): 15章§11 の64ケースを1本の .bas に収める
'   と契約を超える。14章§6 が宣言した公開関数8本の**入口は modValidate 側**に
'   置いたまま、本モジュールには以下の本体だけを移した(依存は modValidate ->
'   modValidate2 の一方向。逆参照はしない)。
'     ・NormalizeCore  : NormalizeLlmJson の正規化エンジン(16章 E-49)
'     ・CheckS4Core    : 15章§5 CheckS4 検証ルール表(6件)
'     ・CheckPFCore    : 15章§6 CheckPF 検証ルール表(7件)
'     ・CheckS2CCore   : 15章§4.5 CheckS2C 検証ルール表(5件)
'     ・CheckS3CCore   : 15章§4.6 CheckS3C 検証ルール表(5件)
'   Core 接尾辞は「14章§6の公開名ではない=分割の内部継ぎ目」を表す。
'
' 純関数のみ。R4準拠(Excelトークン不使用)・CP932準拠・例外を投げない。
' 配列添字の {i} は0始まり。
' ============================================================================

' --- enum(19章§3) ---
Private Const VE_S2C_ISSUE As String = "|missing|generic|weak_evidence|inconsistent|gap_error|insurability_error|"
Private Const VE_S3C_ISSUE As String = "|wont_land|not_executable|wrong_priority|weak_hook|context_mismatch|uw_concern|"
Private Const VE_SURVIVAL As String = "|high|mid|low|"
Private Const VE_APPROACH As String = "|entry_path|benefit_form|underwriter|"
Private Const VE_RELATION As String = "|重複|近接|差分あり|"
Private Const VE_DROP_TYPE As String = "|T1|T2|T3|T4|T5|T6|T7|T8|T9|T10|"
Private Const VE_PATTERN_ID As String = "|P1|P2|P3|P4|P5|P6|P7|P8|P9|P10|P11|P12|P13|P14|P15|"
Private Const VE_QNO As String = "|1|2|3|4|5|"
Private Const VE_GKEY As String = "|a|b|c|d|"
Private Const VE_STORY_NO As String = "|1|2|3|"

' --- 件数のしきい値(15章の各ルール表) ---
Private Const VS4_T1_SLIDES As Long = 5
Private Const VS4_SLIDE_MIN As Long = 5
Private Const VS4_BULLET_MIN As Long = 1
Private Const VS4_BULLET_MAX As Long = 8
Private Const VS4_HQ_MIN As Long = 1
Private Const VS4_HQ_MAX As Long = 10
Private Const VPF_PRINCIPLE_N As Long = 5
Private Const VPF_GRAMMAR_N As Long = 4
Private Const VPF_REWORK_MAX As Long = 3
Private Const VS3C_REACTION_N As Long = 3

' --- 正規化エンジンの再帰上限(壊れた入力で暴走させない) ---
Private Const V_MAX_DEPTH As Long = 40

' --- ID実在検査が実行できないときのエラー文(裁定書7 A-2 の fail-closed。
'     modValidate.V_NOLIST と同一文言。V-S2C-03 の「一覧」は審査対象のS2 JSON) ---
Private Const V_NOLIST As String = "ID実在検査が実行できません(ID一覧未提供)"

' ============================================================================
' NormalizeCore - 配列要素の重複排除(14章§5(2.5)・16章 E-49・E0303)
' ----------------------------------------------------------------------------
'   長文出力の尾部劣化(末尾項目の重複出力)への対策。JSONを走査し、**すべての
'   配列**について要素を modUtilText.NormalizeForHash -> Fnv1a64Hex のハッシュで
'   重複排除する。除去件数を removedCount に返す(呼び出し側が run_log detail へ
'   E0303 として残す。黙って畳んで済ませない)。
'   走査できない入力(JSONとして壊れている)は**原文をそのまま返し**
'   removedCount=0 とする(捏造も部分破壊もしない。14章§5 パターン(4)と同じ作法)。
'   stepName は14章§6の宣言どおり受けるが、重複排除の規則はstep非依存のため
'   分岐には使わない(呼び出し側のログ用に契約を保つ)。
' ============================================================================
Public Function NormalizeCore(ByVal stepName As String, ByVal json As String, _
                              ByRef removedCount As Long) As String
    Dim pos As Long
    Dim outText As String
    Dim work As Long

    removedCount = 0
    NormalizeCore = json
    If LenB(Trim$(json)) = 0 Then Exit Function

    pos = 1
    work = 0
    outText = NormValue(json, pos, work, 0)
    If LenB(outText) = 0 Then Exit Function

    SkipWs json, pos
    If pos <= Len(json) Then Exit Function

    removedCount = work
    NormalizeCore = outText
End Function

' ============================================================================
' CheckS4Core - 骨子生成の検証(15章§5 CheckS4 検証ルール表。6件)
' ----------------------------------------------------------------------------
'   dossierTier : t1_quick / t2_full / t3_sparring(13章・19章§3)。
'   maxSlidesT2 : config `ppt_max_slides_t2`(既定10)。純関数なので外から渡す。
' ============================================================================
Public Function CheckS4Core(ByVal json As String, ByVal dossierTier As String, _
                            ByVal maxSlidesT2 As Long) As String
    Dim r As String
    Dim slidesCol As Collection
    Dim it As Variant
    Dim sj As String
    Dim noText As String
    Dim nCount As Long
    Dim nSlide As Long
    Dim idx As Long
    Dim tier As String

    tier = LCase$(Trim$(dossierTier))
    Set slidesCol = modJsonLite.GetArrayItems(json, "slides")
    nSlide = slidesCol.Count

    ' --- V-S4-01 / V-S4-02: 枚数(ティアで分岐する) ---
    ' ティア(13章 dossier_tier)は s4Json に含まれないので呼び出し側が渡す。
    ' 渡らない(=不明の)ときは、どちらのティアであっても違反になる枚数を見逃さ
    ' ないよう**両方のルールを当てる**。片方のティアを黙って既定に据えると、
    ' もう一方の違反(t1_quick の5枚固定など)が構造的にすり抜ける。
    ' 5枚は両ティアで合法なので、正常な応答はティア不明でも0件のままである。
    If tier <> "t2_full" And tier <> "t3_sparring" Then
        If nSlide <> VS4_T1_SLIDES Then
            Ap r, "[V-S4-01] slides が" & nSlide & "枚です(クイック案件は5枚固定)"
        End If
    End If
    If tier <> "t1_quick" Then
        If nSlide < VS4_SLIDE_MIN Or nSlide > maxSlidesT2 Then
            Ap r, "[V-S4-02] slides が" & nSlide & "枚です(5～" & maxSlidesT2 & "枚)"
        End If
    End If

    idx = 0
    For Each it In slidesCol
        sj = CStr(it)
        idx = idx + 1
        noText = Trim$(modJsonLite.GetStr(sj, "slide_no"))

        ' --- V-S4-03: slide_no は 1..N 各1回の連番 ---
        If modJsonLite.GetLong(sj, "slide_no", 0) <> idx Then
            Ap r, "[V-S4-03] slide_no が1.." & nSlide & "の連番ではありません: " & noText
        End If

        ' --- V-S4-04: bullets の件数 ---
        nCount = modJsonLite.GetArrayItems(sj, "bullets").Count
        If nCount < VS4_BULLET_MIN Or nCount > VS4_BULLET_MAX Then
            Ap r, "[V-S4-04] slide_no " & noText & " の bullets が" & nCount & "点です(1～8点)"
        End If
    Next it

    ' --- V-S4-05: hearing_questions の件数 ---
    nCount = modJsonLite.GetArrayItems(json, "hearing_questions").Count
    If nCount < VS4_HQ_MIN Or nCount > VS4_HQ_MAX Then
        Ap r, "[V-S4-05] hearing_questions が" & nCount & "問です(1～10問)"
    End If

    ' --- V-S4-06: file_title が空 ---
    If LenB(Trim$(modJsonLite.GetStr(json, "file_title"))) = 0 Then
        Ap r, "[V-S4-06] file_title が空です"
    End If

    CheckS4Core = r
End Function

' ============================================================================
' CheckPFCore - プリフライト診断の検証(15章§6 CheckPF 検証ルール表。7件)
' ----------------------------------------------------------------------------
'   refIdsText : PFへ注入した一覧(menusSummary / patternsText / rulesText /
'     researchingText 等を連結したもの。15章§6.1 の1行書式で行頭が "[ID] ")。
'     V-PF-03 の実在判定に使う。空文字は「一覧未提供」= fail-closed で不合格
'     (裁定書7 A-2。V_NOLIST)。PFを結線する側は必ず連結した一覧を渡すこと。
' ============================================================================
Public Function CheckPFCore(ByVal json As String, ByVal refIdsText As String) As String
    Dim r As String
    Dim itemsCol As Collection
    Dim it As Variant
    Dim sj As String
    Dim sVal As String
    Dim seenText As String
    Dim nCount As Long
    Dim idx As Long
    Dim okFlag As Boolean

    ' --- V-PF-01: principle_checks は5件・q_no=1..5各1回 ---
    Set itemsCol = modJsonLite.GetArrayItems(json, "principle_checks")
    nCount = itemsCol.Count
    okFlag = (nCount = VPF_PRINCIPLE_N)
    If okFlag Then
        seenText = "|"
        For Each it In itemsCol
            sVal = Trim$(modJsonLite.GetStr(CStr(it), "q_no"))
            If Not InEnum(sVal, VE_QNO) Then okFlag = False
            If InStr(seenText, "|" & sVal & "|") > 0 Then okFlag = False
            seenText = seenText & sVal & "|"
        Next it
    End If
    If Not okFlag Then
        Ap r, "[V-PF-01] principle_checks が" & nCount & "件です(q_no=1..5各1回)"
    End If

    ' --- V-PF-02: grammar_checks は4件・key=a～d各1回 ---
    Set itemsCol = modJsonLite.GetArrayItems(json, "grammar_checks")
    nCount = itemsCol.Count
    okFlag = (nCount = VPF_GRAMMAR_N)
    If okFlag Then
        seenText = "|"
        For Each it In itemsCol
            sVal = Trim$(modJsonLite.GetStr(CStr(it), "key"))
            If Not InEnum(sVal, VE_GKEY) Then okFlag = False
            If InStr(seenText, "|" & sVal & "|") > 0 Then okFlag = False
            seenText = seenText & sVal & "|"
        Next it
    End If
    If Not okFlag Then
        Ap r, "[V-PF-02] grammar_checks が" & nCount & "件です(key=a～d各1回)"
    End If

    ' --- V-PF-03 / V-PF-07(duplicates[].relation): duplicates ---
    idx = 0
    For Each it In modJsonLite.GetArrayItems(json, "duplicates")
        sj = CStr(it)
        sVal = Trim$(modJsonLite.GetStr(sj, "ref_id"))
        If IsRefIdShaped(sVal) Then
            If Not ListGiven(refIdsText) Then
                Ap r, "[V-PF-03] " & V_NOLIST
            ElseIf Not IdExistsIn(refIdsText, sVal) Then
                Ap r, "[V-PF-03] duplicates[" & idx & "].ref_id " & sVal & " は実在しません"
            End If
        End If
        sVal = modJsonLite.GetStr(sj, "relation")
        If Not InEnum(sVal, VE_RELATION) Then
            Ap r, "[V-PF-07] duplicates[" & idx & "].relation が不正です: " & sVal
        End If
        idx = idx + 1
    Next it

    ' --- V-PF-04 / V-PF-06 / V-PF-07(approach): rework_suggestions ---
    Set itemsCol = modJsonLite.GetArrayItems(json, "rework_suggestions")
    nCount = itemsCol.Count
    If nCount > VPF_REWORK_MAX Then
        Ap r, "[V-PF-04] rework_suggestions が" & nCount & "件です(0～3件)"
    End If
    idx = 0
    For Each it In itemsCol
        sj = CStr(it)
        sVal = Trim$(modJsonLite.GetStr(sj, "pattern_id"))
        If LenB(sVal) > 0 And Not InEnum(sVal, VE_PATTERN_ID) Then
            Ap r, "[V-PF-06] rework_suggestions[" & idx & "].pattern_id が不正です: " & sVal
        End If
        sVal = modJsonLite.GetStr(sj, "approach")
        If Not InEnum(sVal, VE_APPROACH) Then
            Ap r, "[V-PF-07] rework_suggestions[" & idx & "].approach が不正です: " & sVal
        End If
        idx = idx + 1
    Next it

    ' --- V-PF-05: predicted_drop_types は T1～T10 ---
    For Each it In modJsonLite.GetArrayItems(json, "predicted_drop_types")
        sVal = Trim$(CStr(it))
        If Not InEnum(sVal, VE_DROP_TYPE) Then
            Ap r, "[V-PF-05] predicted_drop_types に不正な値があります: " & sVal
        End If
    Next it

    ' --- V-PF-07(survival) ---
    sVal = modJsonLite.GetStr(json, "survival")
    If Not InEnum(sVal, VE_SURVIVAL) Then
        Ap r, "[V-PF-07] survival が不正です: " & sVal
    End If

    CheckPFCore = r
End Function

' ============================================================================
' CheckS2CCore - S2批判パスの検証(15章§4.5 CheckS2C 検証ルール表。5件)
' ----------------------------------------------------------------------------
'   s2Json: 審査対象のS2 JSON(V-S2C-03 の risk_no / gap_no 実在判定)。
'     空文字は「審査対象を渡していない」の意であり「S2に番号が1つも無い」では
'     ないため、V-S2C-03 は**実行できない=不合格**とする(裁定書7 A-2 の
'     fail-closed。V_NOLIST。黙って検査を消さない)。
'   V-S2C-05(issues 0件=合格)はエラー文を持たないため戻り値に現れない
'   (15章§0 原則10。改訂パスのスキップ判定は呼び出し側が issues 件数で行う)。
' ============================================================================
Public Function CheckS2CCore(ByVal json As String, ByVal s2Json As String) As String
    Dim r As String
    Dim it As Variant
    Dim sj As String
    Dim sVal As String
    Dim kind As String
    Dim riskSet As String
    Dim gapSet As String
    Dim idx As Long
    Dim s2Given As Boolean

    s2Given = (LenB(Trim$(s2Json)) > 0)
    riskSet = CollectNos(s2Json, "risks", "risk_no")
    gapSet = CollectNos(s2Json, "gaps", "gap_no")

    idx = 0
    For Each it In modJsonLite.GetArrayItems(json, "issues")
        sj = CStr(it)

        ' --- V-S2C-01: issue_type の enum6値 ---
        sVal = modJsonLite.GetStr(sj, "issue_type")
        If Not InEnum(sVal, VE_S2C_ISSUE) Then
            Ap r, "[V-S2C-01] issues[" & idx & "].issue_type が不正です: " & sVal
        End If

        ' --- V-S2C-02 / V-S2C-03: target の書式と実在 ---
        sVal = Trim$(modJsonLite.GetStr(sj, "target"))
        kind = TargetPrefix(sVal)
        If LenB(kind) = 0 Then
            Ap r, "[V-S2C-02] issues[" & idx & "].target の書式が不正です: " & sVal
        ElseIf kind = "risk_no" Then
            If Not s2Given Then
                Ap r, "[V-S2C-03] " & V_NOLIST
            ElseIf InStr(riskSet, "|" & TargetNumText(sVal) & "|") = 0 Then
                Ap r, "[V-S2C-03] issues[" & idx & "].target " & sVal & " が S2 に存在しません"
            End If
        ElseIf kind = "gap_no" Then
            If Not s2Given Then
                Ap r, "[V-S2C-03] " & V_NOLIST
            ElseIf InStr(gapSet, "|" & TargetNumText(sVal) & "|") = 0 Then
                Ap r, "[V-S2C-03] issues[" & idx & "].target " & sVal & " が S2 に存在しません"
            End If
        End If
        idx = idx + 1
    Next it

    ' --- V-S2C-04: verdict_summary が空 ---
    If LenB(Trim$(modJsonLite.GetStr(json, "verdict_summary"))) = 0 Then
        Ap r, "[V-S2C-04] verdict_summary が空です"
    End If

    CheckS2CCore = r
End Function

' ============================================================================
' CheckS3CCore - S3批判パスの検証(15章§4.6 CheckS3C 検証ルール表。5件)
' ----------------------------------------------------------------------------
'   V-S3C-05(lands全true かつ issues 0件=合格)はエラー文を持たない。
' ============================================================================
Public Function CheckS3CCore(ByVal json As String) As String
    Dim r As String
    Dim reactCol As Collection
    Dim it As Variant
    Dim sj As String
    Dim sVal As String
    Dim noText As String
    Dim seenText As String
    Dim nCount As Long
    Dim idx As Long

    ' --- V-S3C-01: executive_reactions は3件固定 ---
    Set reactCol = modJsonLite.GetArrayItems(json, "executive_reactions")
    nCount = reactCol.Count
    If nCount <> VS3C_REACTION_N Then
        Ap r, "[V-S3C-01] executive_reactions が" & nCount & "件です(3件固定)"
    End If

    ' --- V-S3C-02: story_no は 1..3 各1回(欠落・重複・範囲外) ---
    seenText = "|"
    For Each it In reactCol
        noText = Trim$(modJsonLite.GetStr(CStr(it), "story_no"))
        If (Not InEnum(noText, VE_STORY_NO)) Or InStr(seenText, "|" & noText & "|") > 0 Then
            Ap r, "[V-S3C-02] executive_reactions の story_no が1..3各1回ではありません: " & noText
        End If
        seenText = seenText & noText & "|"
    Next it

    idx = 0
    For Each it In modJsonLite.GetArrayItems(json, "issues")
        sj = CStr(it)

        ' --- V-S3C-03: issue_type の enum6値 ---
        sVal = modJsonLite.GetStr(sj, "issue_type")
        If Not InEnum(sVal, VE_S3C_ISSUE) Then
            Ap r, "[V-S3C-03] issues[" & idx & "].issue_type が不正です: " & sVal
        End If

        ' --- V-S3C-04: target は story_no:N(N=1..3) か overall ---
        sVal = Trim$(modJsonLite.GetStr(sj, "target"))
        If Not IsStoryTarget(sVal) Then
            Ap r, "[V-S3C-04] issues[" & idx & "].target の書式が不正です: " & sVal
        End If
        idx = idx + 1
    Next it

    CheckS3CCore = r
End Function

' ============================================================================
' 内部ヘルパ(本モジュール専用の純関数)
' ============================================================================

Private Sub Ap(ByRef outText As String, ByVal lineText As String)
    If LenB(outText) = 0 Then
        outText = lineText
    Else
        outText = outText & vbLf & lineText
    End If
End Sub

Private Function InEnum(ByVal sVal As String, ByVal enumList As String) As Boolean
    If LenB(sVal) = 0 Then Exit Function
    InEnum = (InStr(enumList, "|" & sVal & "|") > 0)
End Function

' ID一覧が渡されたか(fail-closed の判断点。裁定書7 A-2。modValidate.ListGiven と
' 同じ規約: 空は「未提供」であって「一覧に無い」ではないので、呼び出し側は
' V_NOLIST の不合格を出す。15章§6.1 の "(登録なし)" は空ではない)。
Private Function ListGiven(ByVal listText As String) As Boolean
    ListGiven = (LenB(Trim$(listText)) > 0)
End Function

' 注入テキスト(1行1件・行頭 "[ID] ")に当該IDの行があるか(15章§6.1)。ID実在検査
' (15章§6 V-PF-03)の照合部。一覧が空なら False(未提供は ListGiven で先に分ける)。
Private Function IdExistsIn(ByVal listText As String, ByVal idText As String) As Boolean
    If LenB(idText) = 0 Then Exit Function
    If LenB(listText) = 0 Then Exit Function
    If Left$(listText, Len(idText) + 2) = "[" & idText & "]" Then
        IdExistsIn = True
        Exit Function
    End If
    IdExistsIn = (InStr(listText, vbLf & "[" & idText & "]") > 0)
End Function

Private Function CollectNos(ByVal srcJson As String, ByVal arrKey As String, _
                            ByVal noKey As String) As String
    Dim outText As String
    Dim it As Variant
    outText = "|"
    For Each it In modJsonLite.GetArrayItems(srcJson, arrKey)
        outText = outText & Trim$(modJsonLite.GetStr(CStr(it), noKey)) & "|"
    Next it
    CollectNos = outText
End Function

' V-PF-03: "M-" / "S-" / "K-" で始まるID形式か(それ以外はテーマ名等の自由記述)。
Private Function IsRefIdShaped(ByVal sVal As String) As Boolean
    Dim head As String
    If Len(sVal) < 3 Then Exit Function
    head = Left$(sVal, 2)
    IsRefIdShaped = (head = "M-" Or head = "S-" Or head = "K-")
End Function

' V-S2C-02/03: target の書式判定。"overall" / "risk_no" / "gap_no" / "" (不正)。
Private Function TargetPrefix(ByVal sVal As String) As String
    If sVal = "overall" Then
        TargetPrefix = "overall"
        Exit Function
    End If
    If Left$(sVal, 8) = "risk_no:" Then
        If IsDigitsOnly(Mid$(sVal, 9)) Then TargetPrefix = "risk_no"
        Exit Function
    End If
    If Left$(sVal, 7) = "gap_no:" Then
        If IsDigitsOnly(Mid$(sVal, 8)) Then TargetPrefix = "gap_no"
    End If
End Function

' target の ":" 以降(番号部)。
Private Function TargetNumText(ByVal sVal As String) As String
    Dim p As Long
    p = InStr(sVal, ":")
    If p > 0 Then TargetNumText = Mid$(sVal, p + 1)
End Function

' V-S3C-04: "story_no:N"(N=1..3) か "overall" か。
Private Function IsStoryTarget(ByVal sVal As String) As Boolean
    If sVal = "overall" Then
        IsStoryTarget = True
        Exit Function
    End If
    If Left$(sVal, 9) = "story_no:" Then
        IsStoryTarget = InEnum(Mid$(sVal, 10), VE_STORY_NO)
    End If
End Function

Private Function IsDigitsOnly(ByVal sVal As String) As Boolean
    Dim i As Long
    Dim ch As String
    If LenB(sVal) = 0 Then Exit Function
    For i = 1 To Len(sVal)
        ch = Mid$(sVal, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsDigitsOnly = True
End Function

' ============================================================================
' 正規化エンジン(16章 E-49)
' ----------------------------------------------------------------------------
'   JSONを値単位で読み直し、配列だけ重複要素を落として組み直す。走査に失敗した
'   ときは "" を返し、呼び出し側(NormalizeCore)が原文へ倒す。
' ============================================================================
Private Function NormValue(ByVal s As String, ByRef pos As Long, _
                           ByRef removedCount As Long, ByVal depth As Long) As String
    Dim ch As String
    If depth > V_MAX_DEPTH Then Exit Function
    SkipWs s, pos
    If pos > Len(s) Then Exit Function
    ch = Mid$(s, pos, 1)
    If ch = "{" Then
        NormValue = NormObject(s, pos, removedCount, depth)
    ElseIf ch = "[" Then
        NormValue = NormArray(s, pos, removedCount, depth)
    ElseIf ch = """" Then
        NormValue = ReadStrLit(s, pos)
    Else
        NormValue = ReadScalarTok(s, pos)
    End If
End Function

Private Function NormObject(ByVal s As String, ByRef pos As Long, _
                            ByRef removedCount As Long, ByVal depth As Long) As String
    Dim outText As String
    Dim ch As String
    Dim kText As String
    Dim vText As String
    Dim cnt As Long
    Dim n As Long

    n = Len(s)
    pos = pos + 1
    outText = "{"
    Do
        SkipWs s, pos
        If pos > n Then Exit Function
        ch = Mid$(s, pos, 1)
        If ch = "}" Then
            pos = pos + 1
            NormObject = outText & "}"
            Exit Function
        End If
        If ch = "," Then
            pos = pos + 1
        Else
            If ch <> """" Then Exit Function
            kText = ReadStrLit(s, pos)
            If LenB(kText) = 0 Then Exit Function
            SkipWs s, pos
            If pos > n Then Exit Function
            If Mid$(s, pos, 1) <> ":" Then Exit Function
            pos = pos + 1
            vText = NormValue(s, pos, removedCount, depth + 1)
            If LenB(vText) = 0 Then Exit Function
            If cnt > 0 Then outText = outText & ","
            outText = outText & kText & ":" & vText
            cnt = cnt + 1
        End If
    Loop
End Function

Private Function NormArray(ByVal s As String, ByRef pos As Long, _
                           ByRef removedCount As Long, ByVal depth As Long) As String
    Dim outText As String
    Dim ch As String
    Dim vText As String
    Dim seenHash As String
    Dim hx As String
    Dim cnt As Long
    Dim n As Long

    n = Len(s)
    pos = pos + 1
    outText = "["
    seenHash = "|"
    Do
        SkipWs s, pos
        If pos > n Then Exit Function
        ch = Mid$(s, pos, 1)
        If ch = "]" Then
            pos = pos + 1
            NormArray = outText & "]"
            Exit Function
        End If
        If ch = "," Then
            pos = pos + 1
        Else
            vText = NormValue(s, pos, removedCount, depth + 1)
            If LenB(vText) = 0 Then Exit Function
            hx = modUtilText.Fnv1a64Hex(modUtilText.NormalizeForHash(vText))
            If InStr(seenHash, "|" & hx & "|") > 0 Then
                removedCount = removedCount + 1
            Else
                seenHash = seenHash & hx & "|"
                If cnt > 0 Then outText = outText & ","
                outText = outText & vText
                cnt = cnt + 1
            End If
        End If
    Loop
End Function

' 開き引用符からの文字列リテラルを引用符ごと返す(失敗は "")。
Private Function ReadStrLit(ByVal s As String, ByRef pos As Long) As String
    Dim n As Long
    Dim i As Long
    Dim ch As String
    n = Len(s)
    If Mid$(s, pos, 1) <> """" Then Exit Function
    i = pos + 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "\" Then
            i = i + 2
        ElseIf ch = """" Then
            ReadStrLit = Mid$(s, pos, i - pos + 1)
            pos = i + 1
            Exit Function
        Else
            i = i + 1
        End If
    Loop
End Function

' 数値・true/false/null のトークンを返す(失敗は "")。
Private Function ReadScalarTok(ByVal s As String, ByRef pos As Long) As String
    Dim n As Long
    Dim i As Long
    n = Len(s)
    i = pos
    Do While i <= n
        If InStr(",}] " & vbTab & vbCr & vbLf, Mid$(s, i, 1)) > 0 Then Exit Do
        i = i + 1
    Loop
    If i = pos Then Exit Function
    ReadScalarTok = Mid$(s, pos, i - pos)
    pos = i
End Function

Private Sub SkipWs(ByVal s As String, ByRef pos As Long)
    Dim n As Long
    n = Len(s)
    Do While pos <= n
        If InStr(" " & vbTab & vbCr & vbLf, Mid$(s, pos, 1)) = 0 Then Exit Do
        pos = pos + 1
    Loop
End Sub
