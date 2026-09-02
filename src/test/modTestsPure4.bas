Attribute VB_Name = "modTestsPure4"
Option Explicit

' ============================
' modTestsPure4 - 裁定書6 B/C で14章§6へ公開された純関数の契約テスト
'   (17章§4-1 層(a))G18-G26
' ----------------------------
' 役割: 裁定書6 D-11「テスト独立執筆・第2ラウンド」。**実装を一切読まず**、
'   15章の書式例・13章・11章§4・16章・15章§0.7 だけを根拠に、新しく公開された
'   純関数の入出力を固定する。W2aでは整形・採番・参照優先・状態遷移が Private
'   のままで、区切りを ` | ` から ` / ` へ壊しても誰も気付かなかった。
'   入口は Public Sub RunAll()。
'
' 統合者へ(本ファイル単体では1本も実行されない。3点セットで結線すること):
'   (1) modTestsPure3.RunAll の末尾から modTestsPure4.RunAll を呼ぶ
'       (modTestsPure -> modTestsPure2 -> modTestsPure3 と同じ数珠つなぎ)。
'   (2) build/modules.json へ1件追加(modTestsPure4 / src/test/modTestsPure4.bas
'       / role=test / type=std)。
'   (3) wintest/tests_expected.txt を 184 -> 264 へ更新(本ファイル80本)。
'   run_lo_tests.py の PURE_ALLOWLIST には必要な5モジュールが既にある。
'
' 設計判断(R4準拠): Excelトークン不使用。改行は vbLf 基準。
'
' テスト本数: 81本 = G18 Fmt* 21 / G19 TrimKbLine 4 / G20 TrimPlan 10 /
'   G21 Fill 8 / G22 Asm* 13 / G23 CaseId 6 / G24 CanTransition 7 /
'   G25 ResolveDataKey 10 / G26 BufText 2。期待値の根拠章は各テスト名の末尾。
'
' 前提とする公開契約(14章§6): modKnowledgeFmt.Fmt*(10本)/TrimKbLine、
'   modPipeline4.TrimPlan(T-57 で modKnowledgeFmt から移設・6段化)、
'   modPromptsOps.Fill/AsmS1User/AsmS2User/AsmS4System/AsmS4User、modCaseStore.
'   BuildCaseId/IsValidCaseId/CanTransition/ResolveDataKey、modUtil.Buf*、TCaseCtx
' ============================

' 15章§6.1・§3: 0行/未装填の既定文言。
Private Const KB_NONE As String = "(登録なし)"
' 15章§4: schemes/cases の「無い場合は」。
Private Const KB_NASHI As String = "なし"
' 15章§3: riskLibText の0行既定。
Private Const RL_NONE As String = "(この業種の登録知識はまだありません)"
' 15章§3 {{incidentsText}} の0行時の既定文言(v2.6・裁定書25 S6)。
Private Const IC_NONE As String = "(この業種の登録事例はまだありません)"
' 15章§0.7: 行内切詰め長。
Private Const KB_LINE As Long = 400

Public Sub RunAll()
    On Error GoTo F18
    T_FmtKnowledge
G19:
    On Error GoTo F19
    T_TrimKbLine
G20:
    On Error GoTo F20
    T_TrimPlan
G21:
    On Error GoTo F21
    T_Fill
G22:
    On Error GoTo F22
    T_Assemble
G23:
    On Error GoTo F23
    T_CaseId
G24:
    On Error GoTo F24
    T_Transition
G25:
    On Error GoTo F25
    T_DataKey
G26:
    On Error GoTo F26
    T_BufSep
GDone:
    On Error GoTo 0
    modTestsPure5.RunAll
    Exit Sub

F18:
    GroupFail "G18 FmtKnowledge"
    Resume G19
F19:
    GroupFail "G19 TrimKbLine"
    Resume G20
F20:
    GroupFail "G20 TrimPlan"
    Resume G21
F21:
    GroupFail "G21 Fill"
    Resume G22
F22:
    GroupFail "G22 Assemble"
    Resume G23
F23:
    GroupFail "G23 CaseId"
    Resume G24
F24:
    GroupFail "G24 Transition"
    Resume G25
F25:
    GroupFail "G25 DataKey"
    Resume G26
F26:
    GroupFail "G26 BufSep"
    Resume GDone
End Sub

' ----------------------------
' 共通ヘルパ(modTestsPure3 と同じ作法)
' ----------------------------
Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

' 文字列一致の1本(長い値は先頭120字だけ残す)。
Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), _
        "期待=[" & HeadOf(want) & "] 実際=[" & HeadOf(act) & "]"
End Sub

Private Sub ChkN(ByVal nm As String, ByVal act As Long, ByVal want As Long)
    modTestRunner.Check nm, (act = want), "期待=" & want & " 実際=" & act
End Sub

Private Sub ChkB(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check nm, cond, detail
End Sub

' レポート用の先頭抜粋(SafeLeft に依存させない)。
Private Function HeadOf(ByVal s As String) As String
    If Len(s) <= 120 Then
        HeadOf = s
    Else
        HeadOf = Left(s, 120) & "...(全" & Len(s) & "字)"
    End If
End Function

' ch を n 個ならべる(倍々に伸ばして切る)。
Private Function RepChar(ByVal ch As String, ByVal n As Long) As String
    Dim s As String
    If n <= 0 Or Len(ch) <= 0 Then
        RepChar = ""
        Exit Function
    End If
    s = ch
    Do While Len(s) < n
        s = s & s
    Loop
    RepChar = Left(s, n)
End Function

' MkRows: 14章§6 modKnowledgeFmt が受ける「Range.Value 由来の2次元Variant配列」
'   (1行目=見出し=列名 / 2行目以降=データ行 / 添字 1..n・1..cols)を組む。
'   hdrText=列名を vbTab 連結(13章§3の物理名。各 Hdr* が返す)、dataLines=行を
'   vbLf・列を vbTab 区切り(空文字なら見出しだけ=0件)。欠けた列は空文字。
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
            If c - 1 <= UBound(f) Then
                a(r, c) = f(c - 1)
            Else
                a(r, c) = ""
            End If
        Next c
    Next r
    MkRows = a
End Function

Private Function HdrRisk() As String
    HdrRisk = "risk_lib_id" & vbTab & "industry_code" & vbTab & "category" & vbTab & _
        "risk_name" & vbTab & "typical_scenario" & vbTab & "typical_freq" & vbTab & _
        "typical_impact" & vbTab & "check_points" & vbTab & "priority" & vbTab & "source"
End Function

Private Function HdrMenu() As String
    HdrMenu = "menu_id" & vbTab & "menu_name" & vbTab & "summary" & vbTab & _
        "target_categories" & vbTab & "target_industries" & vbTab & "is_active"
End Function

Private Function HdrLine() As String
    HdrLine = "line_id" & vbTab & "line_name" & vbTab & "market_note"
End Function

Private Function HdrScheme() As String
    HdrScheme = "scheme_id" & vbTab & "scheme_name" & vbTab & "pattern_id" & vbTab & _
        "structure" & vbTab & "conditions" & vbTab & "signals" & vbTab & "example" & _
        vbTab & "status"
End Function

Private Function HdrCase() As String
    HdrCase = "case_lib_id" & vbTab & "industry_code" & vbTab & "customer_profile" & _
        vbTab & "risk_presented" & vbTab & "proposal" & vbTab & "why_it_worked" & _
        vbTab & "quote"
End Function

Private Function HdrPattern() As String
    HdrPattern = "pattern_id" & vbTab & "pattern_name" & vbTab & "structure" & vbTab & _
        "conditions" & vbTab & "examples_public" & vbTab & "internal_refs"
End Function

Private Function HdrRule() As String
    HdrRule = "rule_id" & vbTab & "rule_class" & vbTab & "rule_text" & vbTab & _
        "workaround" & vbTab & "source"
End Function

Private Function HdrMech() As String
    HdrMech = "mech_id" & vbTab & "mech_text" & vbTab & "layer" & vbTab & _
        "target_categories" & vbTab & "source_uid"
End Function

Private Function HdrRt() As String
    HdrRt = "rt_id" & vbTab & "theme_name" & vbTab & "status" & vbTab & "note"
End Function

' ----------------------------
' G18 Fmt*(15章§3の整形例2行 / §4の整形例5行 / §6.1の書式表)
'   「書式例を一字一句そのまま期待値に置く」ことが目的で、サンプル行の各セル値は
'   書式例の該当項目そのもの。共通規約(§6.1)は 1行1件・行頭 `[ID] `・空項目省略。
' ----------------------------
Private Sub T_FmtKnowledge()
    Dim rws As Variant
    Dim d1 As String
    Dim d2 As String
    Dim e1 As String
    Dim e2 As String

    ' --- FmtRiskLib: 15章§3 整形例1行目 ---
    d1 = "RL-09-003" & vbTab & "09" & vbTab & "manufacturing_quality" & vbTab & _
         "アレルゲン表示誤り" & vbTab & "表示の確認漏れで自主回収に至る" & vbTab & _
         "mid" & vbTab & "large" & vbTab & "表示チェック体制;製造ライン分離" & vbTab & _
         "1" & vbTab & "UWマニュアル"
    e1 = "[RL-09-003] カテゴリ:manufacturing_quality リスク:アレルゲン表示誤り " & _
         "典型シナリオ:表示の確認漏れで自主回収に至る 典型頻度:mid 典型影響:large " & _
         "確認点:表示チェック体制;製造ライン分離"
    rws = MkRows(HdrRisk(), d1)
    ChkS "FmtRiskLib_1行の書式_15章§3", modKnowledgeFmt.FmtRiskLib(rws), e1

    rws = MkRows(HdrRisk(), "")
    ChkS "FmtRiskLib_0行は業種未登録の既定文言_15章§3", _
        modKnowledgeFmt.FmtRiskLib(rws), RL_NONE

    ChkS "FmtRiskLib_配列でない入力も既定文言で続行_16章E09", _
        modKnowledgeFmt.FmtRiskLib(Empty), RL_NONE

    ' --- FmtMenusSummary: 15章§3 整形例2行目・§6.1 ---
    d1 = "M-0012" & vbTab & "食品工場リスク診断サービス" & vbTab & _
         "工場の現地診断と改善提案" & vbTab & _
         "manufacturing_quality;supply_chain" & vbTab & "09" & vbTab & "TRUE"
    rws = MkRows(HdrMenu(), d1)
    ChkS "FmtMenusSummary_1行の書式_15章§3", _
        modKnowledgeFmt.FmtMenusSummary(rws), _
        "[M-0012] 食品工場リスク診断サービス | " & _
        "対応カテゴリ:manufacturing_quality;supply_chain"

    rws = MkRows(HdrMenu(), "")
    ChkS "FmtMenusSummary_0行は登録なし_15章§3", _
        modKnowledgeFmt.FmtMenusSummary(rws), KB_NONE

    ' --- FmtMenus: 15章§4 整形例1行目(概要つき) ---
    rws = MkRows(HdrMenu(), d1)
    ChkS "FmtMenus_1行の書式_15章§4", modKnowledgeFmt.FmtMenus(rws), _
        "[M-0012] 食品工場リスク診断サービス | 概要:工場の現地診断と改善提案 | " & _
        "対応カテゴリ:manufacturing_quality;supply_chain"

    ' --- FmtLines: 15章§4 整形例2-3行目(market_note省略) ---
    d1 = "L-03" & vbTab & "生産物賠償責任保険(PL保険)" & vbTab & _
         "再保険料率の上昇で限度額に慎重"
    e1 = "[L-03] 生産物賠償責任保険(PL保険) | 市場環境:再保険料率の上昇で限度額に慎重"
    d2 = "L-07" & vbTab & "企業総合賠償責任保険"
    e2 = "[L-07] 企業総合賠償責任保険"

    rws = MkRows(HdrLine(), d1)
    ChkS "FmtLines_market_noteつきの1行_15章§4", modKnowledgeFmt.FmtLines(rws), e1

    rws = MkRows(HdrLine(), d2)
    ChkS "FmtLines_market_note空欄は市場環境ごと省略_15章§4", _
        modKnowledgeFmt.FmtLines(rws), e2

    rws = MkRows(HdrLine(), d1 & vbLf & d2)
    ChkS "FmtLines_2行の並びと行区切り_15章§4", _
        modKnowledgeFmt.FmtLines(rws), e1 & vbLf & e2

    ' --- FmtSchemes: 15章§4 整形例4行目 ---
    d1 = "S-0004" & vbTab & "見守りヤモリ型" & vbTab & "P2" & vbTab & _
         "検知パートナー×有事補償バンドル" & vbTab & _
         "検知パートナーの実在;現場の受容;精算の設計" & vbTab & _
         "高齢者向け住宅の運営" & vbTab & "社内実証" & vbTab & "proven"
    rws = MkRows(HdrScheme(), d1)
    ChkS "FmtSchemes_1行の書式_15章§4", modKnowledgeFmt.FmtSchemes(rws), _
        "[S-0004] 見守りヤモリ型(P2) | 構造:検知パートナー×有事補償バンドル | " & _
        "成立条件:検知パートナーの実在;現場の受容;精算の設計 | " & _
        "適用シグナル:高齢者向け住宅の運営"

    rws = MkRows(HdrScheme(), "")
    ChkS "FmtSchemes_0行はなし_15章§4", modKnowledgeFmt.FmtSchemes(rws), KB_NASHI

    ' --- FmtCases: 15章§4 整形例5行目 ---
    d1 = "K-0003" & vbTab & "09" & vbTab & "中堅の菓子製造" & vbTab & _
         "表示誤りによる自主回収" & vbTab & "生産物賠償の限度額拡大" & vbTab & _
         "現場の実感に接続できたこと" & vbTab & "回収は他人事ではない"
    rws = MkRows(HdrCase(), d1)
    ChkS "FmtCases_1行の書式_15章§4", modKnowledgeFmt.FmtCases(rws), _
        "[K-0003] 業種:09 顧客像:中堅の菓子製造 提示リスク:表示誤りによる自主回収 " & _
        "提案:生産物賠償の限度額拡大 決め手:現場の実感に接続できたこと"

    rws = MkRows(HdrCase(), "")
    ChkS "FmtCases_0行はなし_15章§4", modKnowledgeFmt.FmtCases(rws), KB_NASHI

    ' --- FmtPatterns: 15章§6.1 ---
    d1 = "P2" & vbTab & "検知×補償バンドル" & vbTab & _
         "検知サービスとセットで残余リスクを保険がカバー" & vbTab & _
         "検知パートナーの実在;検知から引受条件化への接続" & vbTab & _
         "漏水センサー×水濡れ" & vbTab & "見守りヤモリ型"
    rws = MkRows(HdrPattern(), d1)
    ChkS "FmtPatterns_1行の書式_15章§6.1", modKnowledgeFmt.FmtPatterns(rws), _
        "[P2] 検知×補償バンドル | 構造:検知サービスとセットで残余リスクを保険がカバー" & _
        " | 成立条件:検知パートナーの実在;検知から引受条件化への接続 | " & _
        "代表例:漏水センサー×水濡れ | 社内実績:見守りヤモリ型"

    rws = MkRows(HdrPattern(), "")
    ChkS "FmtPatterns_0行は登録なし_15章§6.1", _
        modKnowledgeFmt.FmtPatterns(rws), KB_NONE

    ' --- FmtRules: 15章§6.1 ---
    d1 = "J-03" & vbTab & "adverse_selection" & vbTab & _
         "加入者が予兆を知っている設計は引き受けない" & vbTab & "entry_path" & _
         vbTab & "投稿ボックス分析"
    rws = MkRows(HdrRule(), d1)
    ChkS "FmtRules_1行の書式_15章§6.1", modKnowledgeFmt.FmtRules(rws), _
        "[J-03] class:adverse_selection " & _
        "基準:加入者が予兆を知っている設計は引き受けない | 破り方:entry_path"

    rws = MkRows(HdrRule(), "")
    ChkS "FmtRules_0行は登録なし_15章§6.1", modKnowledgeFmt.FmtRules(rws), KB_NONE

    ' --- FmtMechs: 15章§6.1(適用リスクの源は13章§3.5) ---
    d1 = "MC-0107" & vbTab & "振動センサーで設備異常を予兆検知" & vbTab & "detect" & _
         vbTab & "facility_bcp;manufacturing_quality" & vbTab & "u-0107"
    rws = MkRows(HdrMech(), d1)
    ChkS "FmtMechs_1行の書式_15章§6.1", modKnowledgeFmt.FmtMechs(rws), _
        "[MC-0107] layer:detect 機構:振動センサーで設備異常を予兆検知 | " & _
        "適用リスク:facility_bcp;manufacturing_quality"

    d2 = "MC-0201" & vbTab & "加入経路を管理組合へ寄せる" & vbTab & "entry" & _
         vbTab & "" & vbTab & "u-0201"
    rws = MkRows(HdrMech(), d2)
    ChkS "FmtMechs_適用リスク空欄は項目ごと省略_15章§6.1", _
        modKnowledgeFmt.FmtMechs(rws), _
        "[MC-0201] layer:entry 機構:加入経路を管理組合へ寄せる"

    ChkS "FmtMechs_機構シート未装填でも登録なしで続行_16章E09", _
        modKnowledgeFmt.FmtMechs(Empty), KB_NONE

    ' --- FmtResearching: 0行既定のみ(13章§3.10 に 判定日/関連 の列が無く、
    '     15章§6.1 の書式例の全項目が一意に決まらないため) ---
    rws = MkRows(HdrRt(), "")
    ChkS "FmtResearching_0行は登録なし_15章§6.1", _
        modKnowledgeFmt.FmtResearching(rws), KB_NONE
End Sub

' ----------------------------
' G19 TrimKbLine: 15章§0.7「各行を先頭400字で切り『…』を付す」。400字以内は
'   何も足さず返す(14章§6)ので戻り値は最大401字。
' ----------------------------
Private Sub T_TrimKbLine()
    Dim src As String

    src = RepChar("a", KB_LINE)
    ChkS "TrimKbLine_400字ちょうどは素通し_15章§0.7", _
        modKnowledgeFmt.TrimKbLine(src), src

    src = RepChar("b", KB_LINE + 1)
    ChkN "TrimKbLine_401字は401字になる_15章§0.7", _
        Len(modKnowledgeFmt.TrimKbLine(src)), KB_LINE + 1

    ChkS "TrimKbLine_切詰め後の先頭400字は原文と一致_15章§0.7", _
        Left(modKnowledgeFmt.TrimKbLine(src), KB_LINE), Left(src, KB_LINE)

    src = RepChar("c", 1000)
    ChkS "TrimKbLine_超過行の末尾に三点リーダを付す_15章§0.7", _
        Right(modKnowledgeFmt.TrimKbLine(src), 1), "…"
End Sub

' ----------------------------
' G20 TrimPlan(15章§0.7 ナレッジ側の切詰め。T-57 で6段化・modPipeline4 へ移設)
'   順序は 1成功事例 -> 2事故事例 -> 3型 -> 4メニュー -> 5種目 ->
'   6リスクライブラリ。counts は12要素(0..5=行数 / 6..11=文字数・並びは同順)。
'   1段ずつ適用し総量を再計算し budgetChars 以下で止める。半減は端数切上げ・
'   下限 0/0/0/5/5/5・文字数は行数に比例と見積もる。サンプル(行数
'   5/5/10/60/20/20・文字数各1000=総量6000)の各段の見積りは
'   5600 -> 5200 -> 4700 -> 4200 -> 3700 -> 3200 と割り切れる値を選んである。
' ----------------------------
Private Sub T_TrimPlan()
    Dim c12(0 To 11) As Long
    Dim c6(0 To 5) As Long
    Dim f12(0 To 11) As Long
    Dim i As Long

    c12(0) = 5
    c12(1) = 5
    c12(2) = 10
    c12(3) = 60
    c12(4) = 20
    c12(5) = 20
    For i = 6 To 11
        c12(i) = 1000
    Next i

    ChkS "TrimPlan_予算0は上限なしで現在の行数のまま_14章§6", _
        PlanText(modPipeline4.TrimPlan(c12, 0)), "5,5,10,60,20,20"

    ChkS "TrimPlan_総量が予算以下なら1段も削らない_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c12, 6000)), "5,5,10,60,20,20"

    ChkS "TrimPlan_1段目は成功事例を端数切上げで半減_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c12, 5900)), "3,5,10,60,20,20"

    ChkS "TrimPlan_2段目は事故事例_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c12, 5400)), "3,3,10,60,20,20"

    ChkS "TrimPlan_3段目は型ライブラリ_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c12, 5000)), "3,3,5,60,20,20"

    ChkS "TrimPlan_4段目はメニュー_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c12, 4500)), "3,3,5,30,20,20"

    ChkS "TrimPlan_5段目は種目_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c12, 4000)), "3,3,5,30,10,20"

    ChkS "TrimPlan_6段目はリスクライブラリ_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c12, 3500)), "3,3,5,30,10,10"

    ' 下限 0/0/0/5/5/5。半減の結果が下限と一致する行数を与えてあるので、
    ' 「下限で止める」「下限そのものを返す」のどちらの読みでも同じ値になる。
    f12(0) = 0
    f12(1) = 0
    f12(2) = 0
    f12(3) = 10
    f12(4) = 10
    f12(5) = 10
    f12(6) = 0
    f12(7) = 0
    f12(8) = 0
    f12(9) = 1000
    f12(10) = 1000
    f12(11) = 1000
    ChkS "TrimPlan_下限は0と0と0と5と5と5_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(f12, 1)), "0,0,0,5,5,5"

    ' 要素が6個以下のときは文字数を0とみなす=切詰め不要と判断する(14章§6)。
    For i = 0 To 5
        c6(i) = c12(i)
    Next i
    ChkS "TrimPlan_文字数を伴わない6要素は切詰め不要_14章§6", _
        PlanText(modPipeline4.TrimPlan(c6, 1)), "5,5,10,60,20,20"
End Sub

' TrimPlan の戻り(6要素・0始まり)を "a,b,c,d,e,f" の1行にする。6要素でない場合は
' その事実が期待値との差分として現れるように件数を出す。
Private Function PlanText(ByVal p As Variant) As String
    Dim i As Long
    Dim n As Long
    Dim s As String
    n = 0
    On Error Resume Next
    n = UBound(p) - LBound(p) + 1
    On Error GoTo 0
    If n <> 6 Then
        PlanText = "(要素数=" & n & ")"
        Exit Function
    End If
    s = ""
    For i = 0 To 5
        If i > 0 Then s = s & ","
        s = s & p(LBound(p) + i)
    Next i
    PlanText = s
End Function

' ----------------------------
' G21 Fill(14章§6 組立層)
'   `{{name}}` を対応値へ全置換する唯一の口。値の中に `{{...}}` が含まれていても
'   再帰置換はしない(1巡だけ行い置換後を再走査しない)。unresolved は置換後に
'   なお残っている `{{` の個数。要素数が食い違う場合は短いほうまで。
' ----------------------------
Private Sub T_Fill()
    Dim ns() As String
    Dim vs() As String

    ns = Split("company", ";")
    vs = Split("浜松スイーツファクトリー", ";")
    ChkS "Fill_単一プレースホルダを値へ置換_14章§6", _
        modPromptsOps.Fill("対象企業名: {{company}}", ns, vs), _
        "対象企業名: 浜松スイーツファクトリー"

    ChkS "Fill_同名プレースホルダを全て置換_14章§6", _
        modPromptsOps.Fill("{{company}}/{{company}}", ns, vs), _
        "浜松スイーツファクトリー/浜松スイーツファクトリー"

    ns = Split("a;b", ";")
    vs = Split("A;B", ";")
    ChkS "Fill_複数の名前を添字対応で置換_14章§6", _
        modPromptsOps.Fill("[{{a}}][{{b}}]", ns, vs), "[A][B]"

    ChkS "Fill_namesに無いプレースホルダはそのまま残る_14章§6", _
        modPromptsOps.Fill("[{{a}}][{{zz}}]", ns, vs), "[A][{{zz}}]"

    ChkN "Fill_全て解決したときの未解決数は0_14章§6", _
        FillUn("[{{a}}][{{b}}]", ns, vs), 0

    ChkN "Fill_未解決のプレースホルダ数を帯域外で返す_14章§6", _
        FillUn("[{{a}}][{{yy}}][{{zz}}]", ns, vs), 2

    ' 値がプレースホルダを名乗っても、その値は展開されない(16章E-04と同じ考え方)。
    ns = Split("a;b", ";")
    vs = Split("{{b}};ZZ", ";")
    ChkS "Fill_値に含まれるプレースホルダを再帰置換しない_14章§6", _
        modPromptsOps.Fill("X{{a}}Y", ns, vs), "X{{b}}Y"

    ns = Split("a;b", ";")
    vs = Split("A", ";")
    ChkS "Fill_要素数が食い違う場合は短いほうまで処理_14章§6", _
        modPromptsOps.Fill("[{{a}}][{{b}}]", ns, vs), "[A][{{b}}]"
End Sub

' Fill の unresolved(省略可能な出口)だけを取り出す小道具。
Private Function FillUn(ByVal tpl As String, ByRef ns() As String, _
                        ByRef vs() As String) As Long
    Dim un As Long
    Dim s As String
    un = -1
    s = modPromptsOps.Fill(tpl, ns, vs, un)
    FillUn = un
End Function

' ----------------------------
' G22 modPromptsOps.Asm*(組立層の純部)
'   (a) S4バリアント差替とフォールバック(§5) (b) 更新指示ブロックの条件挿入
'   (§1.2・§10.1(d)「new のとき何も挿入しない=空行を残さない」) (c) BLOCK_CTX の
'   other_insurers 空欄既定(§1.1)。本文の一致検査は prompt_diff.py(T-23)の担当。
' ----------------------------
Private Sub T_Assemble()
    Dim ctxN As TCaseCtx
    Dim ctxR As TCaseCtx
    Dim s As String
    Dim sp As String
    Dim note As String

    ctxN = MakeCtx("new")
    ctxR = MakeCtx("renewal")

    ' --- (a) S4 バリアント差替(15章§5) ---
    note = ""
    sp = modPromptsOps.AsmS4System("proposal", "t1_quick", note)
    ChkB "AsmS4System_proposalは保険提案書の構成指示を差し込む_15章§5", _
        (InStr(sp, "【構成指示: 保険提案書】") > 0), "実際=[" & HeadOf(sp) & "]"
    ChkS "AsmS4System_proposalではfallbackNoteを立てない_15章§5", note, ""
    ChkB "AsmS4System_差替後にバリアントのプレースホルダが残らない_15章§5", _
        (InStr(sp, "{{BLOCK_S4_VARIANT}}") = 0), "実際=[" & HeadOf(sp) & "]"

    note = ""
    s = modPromptsOps.AsmS4System("alliance", "t1_quick", note)
    ChkB "AsmS4System_allianceは協業提案書の構成指示を差し込む_15章§5", _
        (InStr(s, "【構成指示: 協業提案書】") > 0), "実際=[" & HeadOf(s) & "]"

    note = ""
    s = modPromptsOps.AsmS4System("zzz", "t1_quick", note)
    ChkS "AsmS4System_未知のバリアントはproposalと同じ本文_15章§5", s, sp
    ChkS "AsmS4System_未知のバリアントをfallbackNoteで通知_15章§5", _
        note, "s4_variant_fallback:zzz"

    note = ""
    s = modPromptsOps.AsmS4System("", "t1_quick", note)
    ChkS "AsmS4System_空文字も既定へ落としfallbackNoteに残す_15章§5", _
        note, "s4_variant_fallback:"

    ' --- (b) 更新指示ブロックの条件挿入(15章§1.2・§10.1(d)) ---
    s = modPromptsOps.AsmS1User(ctxN, "HP", "なし", "なし", "なし", "なし", _
                                "なし", "なし", "なし", "なし", "なし")
    ChkB "AsmS1User_new案件では更新指示ブロックを入れない_15章§10.1", _
        (InStr(s, "【更新案件の追加指示】") = 0 And _
         InStr(s, "{{BLOCK_RENEWAL_S1") = 0), "実際=[" & HeadOf(s) & "]"
    ChkB "AsmS1User_new案件では消した行の空行を残さない_15章§10.1", _
        (InStr(s, vbLf & vbLf & "■■■企業情報ここから■■■") > 0 And _
         InStr(s, vbLf & vbLf & vbLf & "■■■企業情報ここから■■■") = 0), _
        "実際=[" & HeadOf(s) & "]"

    s = modPromptsOps.AsmS1User(ctxR, "HP", "なし", "なし", "契約サマリ", "なし", _
                                "なし", "なし", "なし", "なし", "なし")
    ChkB "AsmS1User_renewalでは更新指示ブロックを差し込む_15章§1.2", _
        (InStr(s, "【更新案件の追加指示】") > 0 And _
         InStr(s, "current_coverage に契約の構造化を出力すること") > 0), _
        "実際=[" & HeadOf(s) & "]"

    s = modPromptsOps.AsmS2User(ctxN, "{}", RL_NONE, KB_NONE, "なし", "なし", _
                                IC_NONE, "指定なし", 1)
    ChkB "AsmS2User_new案件では更新指示ブロックを入れない_15章§10.1", _
        (InStr(s, "【更新案件の追加指示】") = 0 And _
         InStr(s, "{{BLOCK_RENEWAL_S2") = 0), "実際=[" & HeadOf(s) & "]"

    ' --- (c) BLOCK_CTX の1行属性(15章§1.1) ---
    ChkB "AsmS2User_他社付保が空なら情報なしと埋める_15章§1.1", _
        (InStr(s, "他社付保の状況メモ: 情報なし") > 0), "実際=[" & HeadOf(s) & "]"

    ' t1_quick の slideCountHint は "5"(15章§5。t2以降は config 依存なので見ない)。
    s = modPromptsOps.AsmS4User(ctxN, "{}", "{}", "{}")
    ChkB "AsmS4User_t1_quickのスライド枚数ヒントは5_15章§5", _
        (InStr(s, "商談用の提案書骨子(スライド5枚)") > 0), "実際=[" & HeadOf(s) & "]"
End Sub

' 14章§6 TCaseCtx(app層 modAppTypes)。Asm* は enum を日本語へ変換しないので、
' 埋まる値はここで渡した文字列そのものになる。
Private Function MakeCtx(ByVal caseType As String) As TCaseCtx
    Dim c As TCaseCtx
    c.case_type = caseType
    c.dossier_tier = "t1_quick"
    c.channel = "wholesale"
    c.kanji = "lead"
    c.bid = "no"
    c.reins = "none"
    c.other_insurers = ""
    c.company = "浜松スイーツファクトリー"
    c.industry_code = "09"
    c.industry_name = "食料品製造業"
    MakeCtx = c
End Function

' ----------------------------
' G23 modCaseStore.BuildCaseId / IsValidCaseId(13章§1・16章E-23)
'   `C-YYYYMMDD-NNN`。dayText は8桁ちょうどの数字、seq は 1..999。
'   桁違い・範囲外は "" を返す(例外は投げない)。連番 000 は不正。
' ----------------------------
Private Sub T_CaseId()
    ChkS "BuildCaseId_8桁の日付と連番から採番する_13章§1", _
        modCaseStore.BuildCaseId("20260901", 1), "C-20260901-001"

    ChkS "BuildCaseId_同日連番は999まで採れる_16章E23", _
        modCaseStore.BuildCaseId("20260901", 999), "C-20260901-999"

    ChkS "BuildCaseId_連番0は範囲外で空文字_14章§6", _
        modCaseStore.BuildCaseId("20260901", 0), ""

    ChkS "BuildCaseId_日付が8桁でなければ空文字_14章§6", _
        modCaseStore.BuildCaseId("2026091", 1), ""

    ChkB "IsValidCaseId_13章§1の形はTrue_13章§1", _
        modCaseStore.IsValidCaseId("C-20260901-001"), "C-20260901-001 を不正と判定した"

    ChkB "IsValidCaseId_連番000は不正_14章§6", _
        (Not modCaseStore.IsValidCaseId("C-20260901-000")), _
        "000 を正当と判定した(連番は001から)"
End Sub

' ----------------------------
' G24 CanTransition(11章§4・16章E-12)。許すのは(1)正順の1段進み(2)任意の状態
'   から error へ(3)error から draft..feedback_done の復帰、の3種だけ。
'   自己遷移・段飛ばし・巻き戻しと8値のenum外は False。
' ----------------------------
Private Sub T_Transition()
    Dim st As Variant
    Dim i As Long
    Dim okAll As Boolean
    Dim ng As String

    st = Split("draft;s1_done;s2_done;s3_done;s4_done;exported;feedback_done", ";")

    okAll = True
    ng = ""
    For i = 0 To 5
        If Not modCaseStore.CanTransition(CStr(st(i)), CStr(st(i + 1))) Then
            okAll = False
            ng = ng & st(i) & "->" & st(i + 1) & " "
        End If
    Next i
    ChkB "CanTransition_正順の1段進みを6組すべて許す_11章§4", okAll, "不許可: " & ng

    okAll = True
    ng = ""
    For i = 0 To 6
        If Not modCaseStore.CanTransition(CStr(st(i)), "error") Then
            okAll = False
            ng = ng & st(i) & "->error "
        End If
    Next i
    ChkB "CanTransition_任意の状態からerrorへは許す_16章E12", okAll, "不許可: " & ng

    okAll = True
    ng = ""
    For i = 0 To 6
        If Not modCaseStore.CanTransition("error", CStr(st(i))) Then
            okAll = False
            ng = ng & "error->" & st(i) & " "
        End If
    Next i
    ChkB "CanTransition_errorからは失敗Stepの戻り先へ復帰できる_16章E12", _
        okAll, "不許可: " & ng

    ChkB "CanTransition_段飛ばしは許さない_11章§4", _
        (Not modCaseStore.CanTransition("draft", "s2_done")), _
        "draft->s2_done を許した"

    ChkB "CanTransition_巻き戻しは許さない_11章§4", _
        (Not modCaseStore.CanTransition("s2_done", "s1_done")), _
        "s2_done->s1_done を許した"

    ChkB "CanTransition_自己遷移は許さない_14章§6", _
        (Not modCaseStore.CanTransition("s2_done", "s2_done")), _
        "s2_done->s2_done を許した"

    ChkB "CanTransition_enum8値に無い値は両側でFalse_13章§2.1", _
        (Not modCaseStore.CanTransition("draft", "done")) And _
        (Not modCaseStore.CanTransition("unknown", "draft")), _
        "13章§2.1のenum外を通した"
End Sub

' ----------------------------
' G25 ResolveDataKey(13章§2.2 の参照優先の純核)。N=2,3 は sN_edited >
'   sNr_json > sN_json / N=1,4 は sN_edited > sN_json(hasRevised無視)。
'   どれも無ければ ""。範囲外も ""。下の8本は N=2 の全8組合せ。
' ----------------------------
Private Sub T_DataKey()
    ChkS "ResolveDataKey_S2のTTTはedited優先_13章§2.2", _
        modCaseStore.ResolveDataKey(2, True, True, True), "s2_edited"

    ChkS "ResolveDataKey_S2のTTFはedited優先_13章§2.2", _
        modCaseStore.ResolveDataKey(2, True, True, False), "s2_edited"

    ChkS "ResolveDataKey_S2のTFTはedited優先_13章§2.2", _
        modCaseStore.ResolveDataKey(2, True, False, True), "s2_edited"

    ChkS "ResolveDataKey_S2のTFFはedited単独_13章§2.2", _
        modCaseStore.ResolveDataKey(2, True, False, False), "s2_edited"

    ChkS "ResolveDataKey_S2のFTTは改訂版が生出力に優先_13章§2.2", _
        modCaseStore.ResolveDataKey(2, False, True, True), "s2r_json"

    ChkS "ResolveDataKey_S2のFTFは改訂版単独_13章§2.2", _
        modCaseStore.ResolveDataKey(2, False, True, False), "s2r_json"

    ChkS "ResolveDataKey_S2のFFTは生出力_13章§2.2", _
        modCaseStore.ResolveDataKey(2, False, False, True), "s2_json"

    ChkS "ResolveDataKey_S2のFFFはどれも無いので空文字_14章§6", _
        modCaseStore.ResolveDataKey(2, False, False, False), ""

    ChkS "ResolveDataKey_S1は改訂パスが無いのでhasRevisedを無視_13章§2.2", _
        modCaseStore.ResolveDataKey(1, False, True, True), "s1_json"

    ChkS "ResolveDataKey_範囲外のstepNoは空文字_14章§6", _
        modCaseStore.ResolveDataKey(0, True, True, True), ""
End Sub

' ----------------------------
' G26 BufText の行区切り(14章§6 modUtil節)。「区切りは vbLf = 1行1件の注入
'   テキストを積む道具」。vbNewLine で継ぐと 15章§6.1 の1行1件にCRが混ざる。
' ----------------------------
Private Sub T_BufSep()
    Dim buf() As String
    Dim cnt As Long
    Dim txt As String

    modUtil.BufInit buf, cnt
    modUtil.BufAdd buf, cnt, "[M-0012] 食品工場リスク診断サービス"
    modUtil.BufAdd buf, cnt, "[L-03] 生産物賠償責任保険(PL保険)"
    txt = modUtil.BufText(buf, cnt)

    ChkS "BufText_2行の区切りはvbLf_14章§6modUtil", txt, _
        "[M-0012] 食品工場リスク診断サービス" & vbLf & _
        "[L-03] 生産物賠償責任保険(PL保険)"

    ChkB "BufText_区切りにCRを混ぜない_14章§6modUtil", _
        (InStr(txt, vbCr) = 0), "実際=[" & HeadOf(txt) & "]"
End Sub
