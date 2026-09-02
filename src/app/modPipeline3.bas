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
' 15章§3 user {{incidentsText}} の0行時の既定文言。
Private Const P3_INCIDENTS_ZERO As String = "(この業種の登録事例はまだありません)"
' 15章§1.2c {{focus_line_ids}} の空時の既定文言。
Private Const P3_FOCUS_NONE As String = "指定なし"
' 13章§2.3 config `kb_incident_rows` の既定値(15章§0.7 の切詰め表)。
Private Const P3_INCIDENT_ROWS_DFLT As Long = 5

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

' {{incidentsText}} の値。0行なら15章§3の既定文言。
Public Function IncidentsBlockText(ByVal rawText As String) As String
    Dim t As String

    t = Trim$(modUtilText.SanitizeInput(rawText))
    If LenB(t) = 0 Then t = P3_INCIDENTS_ZERO
    IncidentsBlockText = t
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
' IncidentsFor - 事故事例(13章§3.11)の1行整形テキスト。**W7時点はスタブ**
' --------------------------------------------------------------------------
'   本実装は班C(17章 T-56 ②)が modKnowledge 側に `IncidentsFor(industryCode,
'   maxRows)` として作る。班Bは呼び口だけを1本に固定し、暫定では空文字を返す
'   (= IncidentsBlockText が15章§3の「(この業種の登録事例はまだありません)」へ
'   落とす)。差し替え時はこの関数の中身を
'   `IncidentsFor = modKnowledge.IncidentsFor(industryCode, maxRows)` の1行に
'   置き換えるだけでよい(呼出側は変えない)。
Public Function IncidentsFor(ByVal industryCode As String, ByVal maxRows As Long) As String
    IncidentsFor = vbNullString
End Function

' --------------------------------------------------------------------------
' 値源の解決(シート・configを読む)
' --------------------------------------------------------------------------

' case_data の input_finance(13章§2.11(a) 7本目の貼付欄)。
Public Function FinanceTextOf(ByVal caseId As String) As String
    FinanceTextOf = FinanceBlockText(modCaseStore.LoadData(caseId, "input_finance"))
End Function

' 事故事例の注入テキスト(config `kb_incident_rows` 既定5)。
Public Function IncidentsTextFor(ByVal industryCode As String) As String
    Dim n As Long

    n = modConfig.GetLong("kb_incident_rows", P3_INCIDENT_ROWS_DFLT)
    If n < 0 Then n = 0
    IncidentsTextFor = IncidentsBlockText(IncidentsFor(industryCode, n))
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

' 15章§3 user。{{incidentsText}} と第2ラウンドの絞り込みをここで解決する。
Public Function S2UserText(ByRef ctx As TCaseCtx, ByVal caseId As String, _
                           ByVal s1Json As String, ByVal riskLib As String, _
                           ByVal menus As String, ByVal prevS2Json As String, _
                           ByVal hearingAnswers As String) As String
    S2UserText = modPromptsOps.AsmS2User(ctx, s1Json, riskLib, menus, prevS2Json, _
                                         hearingAnswers, IncidentsTextFor(ctx.industry_code), _
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
