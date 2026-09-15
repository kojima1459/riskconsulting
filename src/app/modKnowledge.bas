Attribute VB_Name = "modKnowledge"
Option Explicit

' ============================================================================
' modKnowledge - ナレッジブックの読込/キャッシュ/整形注入/実在チェック(T-21)
' 正: 14章§6のmodKnowledge節 / 13章§3(列) / 15章§0.7(行数上限) /
'   16章 E-08(退避)・E-09(0行)・E-13(退避再送)・E-34(列検査)。
' 配置: app層(12章§2)。シート名という製品固有語彙を持つため core層 へ置けない。
'   R4のExcelトークン許可10本の1つで、触るのは kb_path のブックと隠しシート2枚。
' 責務の分離(12章§2・裁定書6 項目6): **整形は持たない**。ここは「読む/絞る/
'   行数上限を適用する/注入IDを積む」まで。1行の書式・0行の既定文言は
'   modKnowledgeFmt(純文字列・Excel非依存)の責務(層(a)から検査できるよう
'   整形をここに閉じ込めない。W2aの実害)。
' 分割(12章§2の30,000字契約・17章§7 Z-13): シートを触らない純関数は
'   modKnowledge2(PickAt/CellAt/CellRaw/AddIdList/ColOf/SelectRows/
'   MissingColsOf/BadRowsOf)。本モジュールはシートに触る側(kb_pathの読込・
'   キャッシュ・退避/復元・再送)だけを持つ。絞込スペック("ind^tgt^act^sts^
'   suf^ref")の規約は modKnowledge2 の冒頭が正。
' 注入ID: 整形テキストへ現れるIDを全て積む(行頭[ID]・型行(P2)・パターン行の
'   参照IDも)。17章T-21「ID集合と完全一致」のため。業種コードはIDでない。
' 15章§0.7: 行数上限(kb_*_rows)は本モジュールが適用。総量3割超のときの半減の
'   計画は modPipeline4.TrimPlan、適用も modPipeline4。各注入関数の
'   Optional maxRows がその口(0=config既定。14章§6)。
' ============================================================================

' --- シート索引(13章§3)。KB_SHEETS の並びと対応 ---
Private Const KB_N As Long = 12
Private Const KB_I_RISK As Long = 2
Private Const KB_I_MENU As Long = 3
Private Const KB_I_LINE As Long = 4
Private Const KB_I_CASE As Long = 6
Private Const KB_I_RULE As Long = 7
Private Const KB_I_PAT As Long = 8
Private Const KB_I_SCHEME As Long = 9
Private Const KB_I_MECH As Long = 10
Private Const KB_I_RT As Long = 11
Private Const KB_I_INC As Long = 12
Private Const KB_SHEETS As String = "業種マスタ|リスクライブラリ|メニュー一覧|種目マスタ|メニュー種目対応|成功事例|判断基準|パターンマスタ|型ライブラリ|機構ライブラリ|研究テーマ一覧|事故事例"
Private Const KB_IDCOLS As String = "industry_code|risk_lib_id|menu_id|line_id||case_lib_id|rule_id|pattern_id|scheme_id|mech_id|rt_id|inc_id"
Private Const KB_REQCOLS As String = "industry_code;industry_name|risk_lib_id;industry_code;category;risk_name;typical_scenario;typical_freq;typical_impact;check_points|menu_id;menu_name;summary;target_categories;target_industries;is_active|line_id;line_name;market_note|menu_id;line_id|case_lib_id;industry_code;customer_profile;risk_presented;proposal;why_it_worked|rule_id;rule_class;rule_text;workaround|pattern_id;pattern_name;structure;conditions;examples_public;internal_refs|scheme_id;scheme_name;pattern_id;structure;conditions;signals;status;target_industries|mech_id;mech_text;layer;target_categories|rt_id;theme_name;status;note|inc_id;industry_code;category;headline;cause;lesson;source"
Private Const KB_SHEET_GAP As String = "新サービス候補"
Private Const KB_GAPCOLS As String = "logged_at,case_id,industry_code,unmatched_risk,operator"

' --- 絞込スペック(1行書式そのものは modKnowledgeFmt が持つ) ---
' 裁定書39 R1-03: 並べ替え補充のn-gram比較に掛ける字数上限の既定値。config
'   `kb_rank_case_chars`/`kb_rank_row_chars`が正(13章§2.3。上限が無いと
'   案件本文2万字x全業種の行の比較でExcelが数分固まる)。
Private Const KB_RANK_CASE_CHARS As Long = 3000
Private Const KB_RANK_ROW_CHARS As Long = 2000
' 裁定書40 P-M3(b): 並べ替え候補**行数**の上限。config`kb_rank_max_rows`が正
'   (13章§2.3。超えた分はシート順で切り run_log detail に残す)。
Private Const KB_RANK_MAX_ROWS As Long = 60

Private Const KB_F_MENU As String = "^target_industries^is_active"
Private Const KB_F_SCHEME As String = "^target_industries^^status^pattern_id"

' 退避シート2枚。13章§2に列挙が無いため実行時に作成して隠す(E-08/E-13。concerns)。
Private Const KB_SHEET_SNAP As String = "kb_snapshot"
Private Const KB_SHEET_PEND As String = "kb_pending"
Private Const KB_VERY_HIDDEN As Long = 2

' 見出し行を読む幅(型ライブラリ13列に余裕を見た探索範囲。列番号ではない)。
Private Const KB_SCAN_COLS As Long = 16
' xlUp の数値(Excel組込定数名を書かずLO側で未定義名にしない)。
Private Const KB_DIR_UP As Long = -4162

Private Const KB_SRC As String = "modKnowledge"
' 裁定書47 G-5: 事故事例の00フォールバック(印はIncidentsFallbackNoteが持つ)。
Private Const KB_INDUSTRY_COMMON As String = "00"

' --- キャッシュ(宣言のみ) ---
Private gKbBlocks(1 To KB_N) As Variant
Private gKbRows(1 To KB_N) As Long
Private gKbLoaded As Boolean
Private gInjectedIds As String
Private mIncidentsFallback As Boolean

' --- kb_cut集計(裁定書38 B-10)。4種のみ追跡(cases/incidents/schemes/risks。
'   いずれも industryCode で絞る種別。0=cases 1=incidents 2=schemes 3=risks)。
Private Const KB_CUT_LABELS As String = "cases|incidents|schemes|risks"
Private gKbCutUsed(0 To 3) As Long
Private gKbCutTotal(0 To 3) As Long
Private gKbCutSeen(0 To 3) As Boolean
' 裁定書40 P-M3(b): 並べ替え候補の打切り。gKbRankSeen=上限を掛ける前の候補数、
'   gKbRankCap=掛けた上限。seen > cap のときだけ run_log detail へ
'   "<種別>_rank_cut=<上限>/<候補総数>" を足す(打ち切っていないのに
'   「打ち切った」と書かない。P-M2 と同じ原則)。
Private gKbRankSeen(0 To 3) As Long
Private gKbRankCap(0 To 3) As Long

' === 公開関数(14章§6のmodKnowledge節。ここに無い名前は公開しない) ===

' LoadKnowledge - kb_path を参照専用で開いて全シートを読み、E-34を検査し、隠し
'   シートへ退避する。接続不可(E-08)は前回退避から復元し E0401 を警告記録する。
'   戻り値 True=ナレッジが使える状態(起動は False でも止めない。12章§2.1 手順⑤)。
'   末尾で新サービス候補の退避分も再送する(E-13。手順⑥の公開口が§6に無いため
'   ここで引き取っている)。
Public Function LoadKnowledge() As Boolean
    On Error GoTo Failed
    gKbLoaded = False

    Dim wb As Object
    Set wb = OpenKbBook(True)
    If wb Is Nothing Then
        LogKb "E0401", "LoadKnowledge", "kb_open_failed"
        If RestoreSnapshot() > 0 Then
            gKbLoaded = True
            LoadKnowledge = True
        End If
        FlushPending
        Exit Function
    End If

    ' 裁定書9 B11(16章 E-08): データ行のあるシートが1枚も無い読込は**読込失敗**
    ' として扱い、SaveSnapshot を呼ばない(0行で前回の正常な退避を上書きすると、
    ' 次に接続できない起動で E-08 の退路そのものが失われる)。
    If ReadKbSheets(wb) > 0 Then
        CloseKbBook wb
        gKbLoaded = True
        ValidateKb
        SaveSnapshot
        FlushPending
        LoadKnowledge = True
        Exit Function
    End If

    CloseKbBook wb
    LogKb "E0401", "LoadKnowledge", "kb_zero_rows"
    If RestoreSnapshot() > 0 Then
        gKbLoaded = True
        LoadKnowledge = True
    End If
    FlushPending
    Exit Function

Failed:
    LogKb "E0401", "LoadKnowledge", "unexpected"
    LoadKnowledge = gKbLoaded
End Function

' 注入関数の共通規約(14章§6): Optional maxRows は 15章§0.7 の半減を外から
'   掛ける口。0=config既定(kb_*_rows)/上限を持たない4種は全行。正の値はその行数。

' RiskLibFor - S2用の業種リスク知識(15章§3)。0行は専用文言(E-09)。
Public Function RiskLibFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
    RiskLibFor = Inject(KB_I_RISK, "risk_lib_id", "industry_code", industryCode, _
                        CapCfg(maxRows, "kb_risk_rows", 20), "risk_lib")
End Function

' MenusSummaryFor - S2用のメニュー要約(15章§3)。related_menu_id の候補一覧を与える
'   唯一の口。industryCode="" は全業種(§6.1 {{menusSummary}})。
Public Function MenusSummaryFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
    MenusSummaryFor = Inject(KB_I_MENU, "menu_id", KB_F_MENU, industryCode, _
                             CapCfg(maxRows, "kb_menu_rows", 60), "menus_summary")
End Function

' MenusFor - S3用のメニュー一覧(概要付き。15章§4。S2の要約とは別テキスト)。
Public Function MenusFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
    MenusFor = Inject(KB_I_MENU, "menu_id", KB_F_MENU, industryCode, _
                      CapCfg(maxRows, "kb_menu_rows", 60), "menus")
End Function

' LinesText - S3用の種目一覧(15章§4。既定は全行)。market_note 空欄は項目ごと省略。
Public Function LinesText(Optional ByVal maxRows As Long = 0) As String
    LinesText = Inject(KB_I_LINE, "line_id", vbNullString, vbNullString, _
                       CapAll(maxRows, KB_I_LINE), "lines")
End Function

' CasesFor - S3用の成功事例(15章§4)。無い場合は「なし」(S3 userの見出し規約)。
'   caseText(裁定書38 B-10・任意): 業種コード完全一致の該当が maxRows に
'   満たないとき、全業種の成功事例から customer_profile / risk_presented との
'   n-gram重なりで上位を補う。空文字なら補充なし(挙動不変)。
Public Function CasesFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0, _
                          Optional ByVal caseText As String = vbNullString) As String
    Dim rankCols As String
    If LenB(caseText) > 0 Then rankCols = "customer_profile;risk_presented"
    CasesFor = Inject(KB_I_CASE, "case_lib_id", "industry_code", industryCode, _
                      CapCfg(maxRows, "kb_case_rows", 5), "cases", caseText, rankCols)
End Function

' SchemesFor - S3用の型ライブラリ(status=proven/adopted のみ。15章§4)。
Public Function SchemesFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
    SchemesFor = Inject(KB_I_SCHEME, "scheme_id", KB_F_SCHEME, industryCode, _
                        CapCfg(maxRows, "kb_scheme_rows", 10), "schemes")
End Function

' PatternsText - P1-P15全件(PF/FG用。15章§6.1)。
Public Function PatternsText(Optional ByVal maxRows As Long = 0) As String
    PatternsText = Inject(KB_I_PAT, "pattern_id", "^^^^^internal_refs", vbNullString, _
                          CapAll(maxRows, KB_I_PAT), "patterns")
End Function

' RulesText - 判断基準(PF/FG用。15章§6.1)。
Public Function RulesText(Optional ByVal maxRows As Long = 0) As String
    RulesText = Inject(KB_I_RULE, "rule_id", vbNullString, vbNullString, _
                       CapAll(maxRows, KB_I_RULE), "rules")
End Function

' ResearchingText - 研究テーマ一覧(PF用。15章§6.1)。
Public Function ResearchingText(Optional ByVal maxRows As Long = 0) As String
    ResearchingText = Inject(KB_I_RT, "rt_id", vbNullString, vbNullString, _
                             CapAll(maxRows, KB_I_RT), "researching")
End Function

' IncidentsFor - S2用の業種別事故事例(15章§3・13章§3.11)。maxRowsは0.7の
'   半減口(0=config既定5)。0行は専用文言(E-09)。G-5: 0行なら"00"で引き直す。
Public Function IncidentsFor(ByVal industryCode As String, Optional ByVal maxRows As Long = 0) As String
    Dim capRows As Long, sel As Variant, ids As String, totalHits As Long, n As Long
    capRows = CapCfg(maxRows, "kb_incident_rows", 5)
    n = modKnowledge2.SelectRows(gKbBlocks(KB_I_INC), gKbRows(KB_I_INC), "inc_id", _
        "industry_code", industryCode, capRows, sel, ids, totalHits)
    mIncidentsFallback = modKnowledge2.ShouldFallbackToCommonIndustry(n, industryCode, KB_INDUSTRY_COMMON)
    If mIncidentsFallback Then _
        n = modKnowledge2.SelectRows(gKbBlocks(KB_I_INC), gKbRows(KB_I_INC), "inc_id", _
            "industry_code", KB_INDUSTRY_COMMON, capRows, sel, ids, totalHits)
    IncidentsFor = FormatBy("incidents", sel)
    RecordKbCut "incidents", n, totalHits
    If n <= 0 Then
        modLog.LogUsage "kb_zero_rows", vbNullString, "field=incidents industry=" & industryCode
        Exit Function
    End If
    modKnowledge2.AddIdList gInjectedIds, ids
End Function

' 直近のIncidentsForが"00"へ差替えていたら1語返す(G-5。積む側はmodPipeline4)。
Public Function IncidentsFallbackNote() As String
    If mIncidentsFallback Then IncidentsFallbackNote = "incidents_fallback=00"
End Function

' MechsText - 機構ライブラリ抜粋(商談の予行演習system)。Phase 1.5のため常に「(登録なし)」。
Public Function MechsText(Optional ByVal maxRows As Long = 0) As String
    MechsText = Inject(KB_I_MECH, "mech_id", vbNullString, vbNullString, _
                       CapCfg(maxRows, "kb_mech_rows", 40), "mechs")
End Function

' LastInjectedIds - run_log.injected_kb_ids(13章§2.4・10章FR-10)を埋める累積値。
Public Function LastInjectedIds() As String
    LastInjectedIds = gInjectedIds
End Function

' ResetInjectedIds - Step開始時に modPipeline / modPlayOps が呼ぶ。
Public Sub ResetInjectedIds()
    gInjectedIds = vbNullString
    Dim i As Long
    For i = 0 To 3
        gKbCutUsed(i) = 0
        gKbCutTotal(i) = 0
        gKbCutSeen(i) = False
        gKbRankSeen(i) = 0
        gKbRankCap(i) = 0
    Next i
End Sub

' ============================================================================
' LastKbCutNote - 裁定書38 B-10: run_log detail へ積む
'   "kb_cut:cases=5/23;incidents=…;schemes=…;risks=…" の1行(直近の
'   ResetInjectedIds 以降にInjectした種別だけを載せる。何も注入していなければ
'   ""(記録しない=検討不要)。
' ============================================================================
Public Function LastKbCutNote() As String
    Dim acc As String
    Dim labels As Variant
    labels = Split(KB_CUT_LABELS, "|")
    Dim i As Long
    For i = 0 To 3
        If gKbCutSeen(i) Then
            If LenB(acc) > 0 Then acc = acc & ";"
            acc = acc & CStr(labels(i)) & "=" & CStr(gKbCutUsed(i)) & "/" & CStr(gKbCutTotal(i))
            ' 裁定書40 P-M3(b): 並べ替えの候補行を config kb_rank_max_rows で
            '   切ったときだけ、同じ1語の中へ "<種別>_rank_cut=上限/候補総数" を足す
            '   (切っていないときは何も足さない)。
            If gKbRankCap(i) > 0 And gKbRankSeen(i) > gKbRankCap(i) Then
                acc = acc & ";" & CStr(labels(i)) & "_rank_cut=" & _
                      CStr(gKbRankCap(i)) & "/" & CStr(gKbRankSeen(i))
            End If
        End If
    Next i
    If LenB(acc) > 0 Then LastKbCutNote = "kb_cut:" & acc
End Function

' LastCasesUsed/LastCasesTotal - HTML SEC-14(裁定書38 B-10)向けの単発参照。
'   直近の CasesFor 呼出結果(未実行なら 0/0)。
Public Function LastCasesUsed() As Long
    If gKbCutSeen(0) Then LastCasesUsed = gKbCutUsed(0)
End Function

Public Function LastCasesTotal() As Long
    If gKbCutSeen(0) Then LastCasesTotal = gKbCutTotal(0)
End Function

' 実在チェック5種(16章 E-07。ホワイトリスト=ナレッジブックの当該ID列)。メニュー
'   だけは is_active も見る(無効化したサービスを提案させないため)。
Public Function MenuIdExists(ByVal id As String) As Boolean
    MenuIdExists = IdExistsIn(KB_I_MENU, "menu_id", id, True)
End Function

Public Function LineIdExists(ByVal id As String) As Boolean
    LineIdExists = IdExistsIn(KB_I_LINE, "line_id", id, False)
End Function

Public Function SchemeIdExists(ByVal id As String) As Boolean
    SchemeIdExists = IdExistsIn(KB_I_SCHEME, "scheme_id", id, False)
End Function

Public Function CaseLibIdExists(ByVal id As String) As Boolean
    CaseLibIdExists = IdExistsIn(KB_I_CASE, "case_lib_id", id, False)
End Function

' 5種のうち pattern_id だけは呼出元が0件である(2026-09 時点の実測)。S1～S4の
'   シートに pattern_id 欄が無く(modUICase2 の実在検査は menu/line/scheme/
'   case_lib の4本)、pattern_id が現れる唯一の場所である商談の予行演習(PF)の
'   rework_suggestions[] は modSchemas の固定enum P1～P15 を modValidate2 の
'   V-PF-06 が照合している。16章 E-07 が「実在検査5種」を掲げている以上、
'   ナレッジブックの patterns シートを増やして V-PF-06 を KB ホワイトリストへ
'   切り替えるときの口として5本目を残す(4本だけ残すと E-07 と食い違う)。
' @unused: 16章 E-07 の実在検査5種の5本目。pattern_id はS1～S4の画面に現れず、PF の rework_suggestions は modSchemas の固定enum(P1～P15)を V-PF-06 が見ているため現時点の呼出元は0件
Public Function PatternIdExists(ByVal id As String) As Boolean
    PatternIdExists = IdExistsIn(KB_I_PAT, "pattern_id", id, False)
End Function

' AppendServiceGap - 新サービス候補(13章§3.9)へ1行追記する。ロック・不通時は
'   退避シートへ積み次回起動で再送(16章 E-13。案件処理は成功扱い=例外は投げない)。
Public Sub AppendServiceGap(ByVal caseId As String, ByVal industryCode As String, _
                            ByVal riskDesc As String)
    On Error GoTo Failed
    EnqueuePending modUtil.NowStamp(), caseId, industryCode, riskDesc, KbOperator()
    FlushPending
    Exit Sub
Failed:
    LogKb "E0401", "AppendServiceGap", "queue_failed"
End Sub

' === 内部(Excel I/O) ===

' 公開の注入関数の共通部: 絞込(SelectRows) -> 整形(modKnowledgeFmt) -> 0行処理
'   (16章 E-09)と注入IDの累積。整形の書式は本モジュールが持たない。
'   caseText/rankCols(裁定書38 B-10・任意): SelectRows の並べ替え補充へ
'   そのまま渡す(空なら補充なし=挙動不変)。
Private Function Inject(ByVal idx As Long, ByVal idCol As String, ByVal filterSpec As String, _
                        ByVal industryCode As String, ByVal capRows As Long, _
                        ByVal fieldName As String, _
                        Optional ByVal caseText As String = vbNullString, _
                        Optional ByVal rankCols As String = vbNullString) As String
    Dim ids As String
    Dim sel As Variant
    Dim n As Long
    Dim totalHits As Long
    ' 裁定書39 R1-03: n-gram 比較に掛ける字数上限。並べ替え補充が走るとき
    ' (caseText 非空)だけ読む。0 を渡すと modKnowledge2 側で無制限になる。
    ' 裁定書40 P-M3(b): 並べ替えに掛ける候補**行数**の上限も同じときに読む。
    Dim caseChars As Long, rowChars As Long, maxCandRows As Long, candSeen As Long
    If LenB(caseText) > 0 Then
        caseChars = CapCfg(0, "kb_rank_case_chars", KB_RANK_CASE_CHARS)
        rowChars = CapCfg(0, "kb_rank_row_chars", KB_RANK_ROW_CHARS)
        maxCandRows = CapCfg(0, "kb_rank_max_rows", KB_RANK_MAX_ROWS)
    End If
    n = modKnowledge2.SelectRows(gKbBlocks(idx), gKbRows(idx), idCol, filterSpec, industryCode, _
                   capRows, sel, ids, totalHits, caseText, rankCols, caseChars, rowChars, _
                   maxCandRows, candSeen)
    Inject = FormatBy(fieldName, sel)
    RecordKbCut fieldName, n, totalHits
    RecordRankCut fieldName, maxCandRows, candSeen
    If n <= 0 Then
        ' 0行は未装填文言(modKnowledgeFmt が返す)のまま usage_log へ記録する
        ' (16章 E-09。ナレッジ整備の優先度シグナル)
        modLog.LogUsage "kb_zero_rows", vbNullString, _
                        "field=" & fieldName & " industry=" & industryCode
        Exit Function
    End If
    modKnowledge2.AddIdList gInjectedIds, ids
End Function

' RecordKbCut - 裁定書38 B-10: cases/incidents/schemes/risk_lib の4種だけ
'   (used, totalHits)を記録する(LastKbCutNote が読む)。他の種別は対象外。
Private Sub RecordKbCut(ByVal fieldName As String, ByVal usedN As Long, ByVal totalN As Long)
    Dim i As Long
    Select Case fieldName
        Case "cases": i = 0
        Case "incidents": i = 1
        Case "schemes": i = 2
        Case "risk_lib": i = 3
        Case Else: Exit Sub
    End Select
    gKbCutUsed(i) = usedN
    gKbCutTotal(i) = totalN
    gKbCutSeen(i) = True
End Sub

' RecordRankCut - 裁定書40 P-M3(b): 並べ替え候補の上限と、上限を掛ける前の
'   候補総数を覚える(LastKbCutNote が seen > cap のときだけ1語足す)。
Private Sub RecordRankCut(ByVal fieldName As String, ByVal capRows As Long, _
                          ByVal seenRows As Long)
    Dim i As Long
    Select Case fieldName
        Case "cases": i = 0
        Case "incidents": i = 1
        Case "schemes": i = 2
        Case "risk_lib": i = 3
        Case Else: Exit Sub
    End Select
    gKbRankCap(i) = capRows
    gKbRankSeen(i) = seenRows
End Sub

' 種類名から modKnowledgeFmt の整形関数へ振り分ける(書式の正は15章・実装は
'   modKnowledgeFmt。ここには書式を書かない)。
Private Function FormatBy(ByVal fieldName As String, ByVal rows As Variant) As String
    Select Case fieldName
        Case "risk_lib"
            FormatBy = modKnowledgeFmt.FmtRiskLib(rows)
        Case "menus_summary"
            FormatBy = modKnowledgeFmt.FmtMenusSummary(rows)
        Case "menus"
            FormatBy = modKnowledgeFmt.FmtMenus(rows)
        Case "lines"
            FormatBy = modKnowledgeFmt.FmtLines(rows)
        Case "cases"
            FormatBy = modKnowledgeFmt.FmtCases(rows)
        Case "schemes"
            FormatBy = modKnowledgeFmt.FmtSchemes(rows)
        Case "patterns"
            FormatBy = modKnowledgeFmt.FmtPatterns(rows)
        Case "rules"
            FormatBy = modKnowledgeFmt.FmtRules(rows)
        Case "researching"
            FormatBy = modKnowledgeFmt.FmtResearching(rows)
        Case "mechs"
            FormatBy = modKnowledgeFmt.FmtMechs(rows)
        Case "incidents"
            FormatBy = modKnowledgeFmt.FmtIncidents(rows)
    End Select
End Function

' maxRows>0 はそのまま。0 は config の行数上限(15章§0.7・13章§2.3。0以下は既定値)。
Private Function CapCfg(ByVal maxRows As Long, ByVal cfgKey As String, ByVal dfltRows As Long) As Long
    If maxRows > 0 Then
        CapCfg = maxRows
        Exit Function
    End If
    Dim v As Long
    v = modConfig.GetLong(cfgKey, dfltRows)
    If v <= 0 Then v = dfltRows
    CapCfg = v
End Function

' maxRows>0 はそのまま。0 は「全行」(§0.7が行数上限を持たない4種)。
Private Function CapAll(ByVal maxRows As Long, ByVal idx As Long) As Long
    If maxRows > 0 Then
        CapAll = maxRows
    Else
        CapAll = gKbRows(idx)
    End If
End Function

Private Sub LogKb(ByVal errCode As String, ByVal procName As String, ByVal detail As String)
    modLog.LogError errCode, KB_SRC & "." & procName, detail
End Sub

' ID実在チェックの実体(needActive=True はメニューの is_active も見る)。
Private Function IdExistsIn(ByVal idx As Long, ByVal colName As String, _
                            ByVal idText As String, ByVal needActive As Boolean) As Boolean
    Dim wanted As String
    wanted = Trim$(idText)
    If LenB(wanted) = 0 Or Not gKbLoaded Then Exit Function
    Dim blk As Variant
    blk = gKbBlocks(idx)
    Dim c As Long, cAct As Long
    c = modKnowledge2.ColOf(blk, colName)
    If c <= 0 Then Exit Function
    If needActive Then cAct = modKnowledge2.ColOf(blk, "is_active")
    Dim r As Long
    For r = 2 To gKbRows(idx)
        If StrComp(modKnowledge2.CellAt(blk, r, c), wanted, vbBinaryCompare) = 0 Then
            IdExistsIn = (cAct <= 0)
            If Not IdExistsIn Then IdExistsIn = modConfig.ParseBoolText(modKnowledge2.CellAt(blk, r, cAct), True)
            If IdExistsIn Then Exit Function
        End If
    Next r
End Function

' kb_path のブックを開く。失敗は Nothing。
Private Function OpenKbBook(ByVal readOnlyMode As Boolean) As Object
    Dim prevAlerts As Boolean
    On Error GoTo Failed
    Dim pathText As String
    pathText = Trim$(modConfig.GetStr("kb_path", vbNullString))
    If LenB(pathText) = 0 Then Exit Function
    ' 追補3: 在ることを先に確かめる(無いとExcelの素のダイアログが出る)。
    ' URL経路は Dir$ で判定できないので従来どおり開きに行く。
    If InStr(1, pathText, "://", vbBinaryCompare) = 0 Then
        If LenB(Dir$(pathText)) = 0 Then Exit Function
    End If
    ' 裁定書27 W9-B7(d): 開く前に DisplayAlerts を退避して False にする
    ' (他者ロック・読取専用推奨・リンク更新のモーダルで起動が固まらないように)。
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set OpenKbBook = Application.Workbooks.Open(pathText, 0, readOnlyMode)
    Application.DisplayAlerts = prevAlerts
    Exit Function
Failed:
    RestoreAlerts prevAlerts
    Set OpenKbBook = Nothing
End Function

' ハンドラ稼働中に On Error Resume Next は書けないので、戻しは別Subへ切り出す。
Private Sub RestoreAlerts(ByVal prevAlerts As Boolean)
    On Error Resume Next
    Application.DisplayAlerts = prevAlerts
End Sub

Private Sub CloseKbBook(ByVal wb As Object)
    On Error GoTo Ignore0
    wb.Close False
    Exit Sub
Ignore0:
    Exit Sub
End Sub

' シートを名前で引く。wb=Nothing は本体ブック。createIfMissing=True のときだけ作る。
Private Function SheetOf(ByVal wb As Object, ByVal sheetTitle As String, _
                         ByVal createIfMissing As Boolean) As Object
    On Error GoTo Failed
    Dim bk As Object
    If wb Is Nothing Then
        Set bk = ThisWorkbook
    Else
        Set bk = wb
    End If
    Dim i As Long
    For i = 1 To bk.Worksheets.Count
        If bk.Worksheets(i).Name = sheetTitle Then
            Set SheetOf = bk.Worksheets(i)
            Exit Function
        End If
    Next i
    If Not createIfMissing Then Exit Function

    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = sheetTitle
    ws.Visible = KB_VERY_HIDDEN
    Set SheetOf = ws
    Exit Function
Failed:
    Set SheetOf = Nothing
End Function

Private Function LastRowOfWs(ByVal ws As Object) As Long
    On Error GoTo One1
    LastRowOfWs = ws.Cells(ws.Rows.Count, 1).End(KB_DIR_UP).Row
    If LastRowOfWs < 1 Then LastRowOfWs = 1
    Exit Function
One1:
    LastRowOfWs = 1
End Function

Private Function ReadWsBlock(ByVal ws As Object, ByVal lastRow As Long, _
                             ByVal colCount As Long) As Variant
    On Error GoTo Empty0
    Dim hi As Long
    hi = lastRow
    If hi < 2 Then hi = 2
    ReadWsBlock = ws.Range(ws.Cells(1, 1), ws.Cells(hi, colCount)).Value
    Exit Function
Empty0:
    ReadWsBlock = Empty
End Function

' 全シートをキャッシュへ読む。戻り=データ行のあるシート数。
Private Function ReadKbSheets(ByVal wb As Object) As Long
    Dim i As Long, n As Long, lastRow As Long
    Dim ws As Object
    For i = 1 To KB_N
        Set ws = SheetOf(wb, modKnowledge2.PickAt(KB_SHEETS, i), False)
        gKbBlocks(i) = Empty
        gKbRows(i) = 0
        If Not ws Is Nothing Then
            lastRow = LastRowOfWs(ws)
            gKbBlocks(i) = ReadWsBlock(ws, lastRow, KB_SCAN_COLS)
            gKbRows(i) = lastRow
            If lastRow >= 2 Then n = n + 1
        End If
    Next i
    ReadKbSheets = n
End Function

' 16章 E-34: 必須列欠落・ID重複・category語彙ずれを E0402 で行番号列挙する
'   (利用者向けの案内はui層)。読込は止めない。
Private Sub ValidateKb()
    Dim i As Long
    Dim nameText As String
    For i = 1 To KB_N
        If gKbRows(i) >= 1 Then
            nameText = modKnowledge2.PickAt(KB_SHEETS, i) & ":"
            LogBad "missing_cols:" & nameText, modKnowledge2.MissingColsOf(gKbBlocks(i), modKnowledge2.PickAt(KB_REQCOLS, i))
            LogBad "dup_id_rows:" & nameText, _
                   modKnowledge2.BadRowsOf(gKbBlocks(i), gKbRows(i), modKnowledge2.PickAt(KB_IDCOLS, i), "dup")
        End If
    Next i
    LogBad "bad_category_rows:", modKnowledge2.BadRowsOf(gKbBlocks(KB_I_RISK), gKbRows(KB_I_RISK), "category", "cat")
    ' 13章§3.11: 事故事例の category も同じ10分類(E-34 の検査対象)。
    LogBad "bad_category_rows:事故事例:", modKnowledge2.BadRowsOf(gKbBlocks(KB_I_INC), gKbRows(KB_I_INC), "category", "cat")
End Sub

' 非空のときだけ E0402 を記録する(E-34)。
Private Sub LogBad(ByVal headText As String, ByVal badText As String)
    If LenB(badText) > 0 Then LogKb "E0402", "ValidateKb", headText & badText
End Sub

' 16章 E-08: 隠しシートへの退避。1行 = [シート索引][そのシートの行数][元の1行を
'   vbTabで連結]。索引が同じ行は連続して並ぶので、復元は前から読み戻すだけでよい。
Private Function SaveSnapshot() As Boolean
    On Error GoTo Failed
    Dim ws As Object
    Set ws = SheetOf(Nothing, KB_SHEET_SNAP, True)
    If ws Is Nothing Then Exit Function
    ws.Cells.Clear

    Dim i As Long, r As Long, c As Long, wr As Long
    Dim acc As String
    For i = 1 To KB_N
        For r = 1 To gKbRows(i)
            acc = vbNullString
            For c = 1 To KB_SCAN_COLS
                acc = acc & modKnowledge2.CellRaw(gKbBlocks(i), r, c) & vbTab
            Next c
            wr = wr + 1
            ws.Cells(wr, 1).Value = i            ' SAFE:const 内部生成のLong
            ws.Cells(wr, 2).Value = gKbRows(i)   ' SAFE:const 内部生成のLong
            modUtilText.SetCellSafe ws.Cells(wr, 3), acc, KB_SHEET_SNAP
        Next r
    Next i
    SaveSnapshot = True
    Exit Function
Failed:
    LogKb "E0401", "SaveSnapshot", "snapshot_write_failed"
    SaveSnapshot = False
End Function

' 退避からキャッシュを復元する。戻り=復元できたシート数(E-08)。
Private Function RestoreSnapshot() As Long
    On Error GoTo Failed
    Dim ws As Object
    Set ws = SheetOf(Nothing, KB_SHEET_SNAP, False)
    If ws Is Nothing Then Exit Function
    Dim blk As Variant
    Dim lastRow As Long
    lastRow = LastRowOfWs(ws)
    blk = ReadWsBlock(ws, lastRow, 3)
    If Not IsArray(blk) Then Exit Function

    Dim arr() As Variant
    Dim parts As Variant
    Dim i As Long, r As Long, c As Long, cnt As Long, n As Long
    Dim curIdx As Long, curRow As Long
    For i = 1 To KB_N
        gKbBlocks(i) = Empty
        gKbRows(i) = 0
    Next i
    For r = 1 To lastRow
        i = CLng(Val(modKnowledge2.CellRaw(blk, r, 1)))
        cnt = CLng(Val(modKnowledge2.CellRaw(blk, r, 2)))
        If i <> curIdx Then
            If curIdx > 0 Then gKbBlocks(curIdx) = arr
            curIdx = 0
            If i >= 1 And i <= KB_N And cnt >= 1 Then
                ReDim arr(1 To cnt, 1 To KB_SCAN_COLS)
                gKbRows(i) = cnt
                curIdx = i
                curRow = 0
                n = n + 1
            End If
        End If
        If curIdx > 0 Then
            curRow = curRow + 1
            parts = Split(modKnowledge2.CellRaw(blk, r, 3), vbTab)
            For c = 1 To KB_SCAN_COLS
                If c <= UBound(parts) + 1 Then arr(curRow, c) = parts(c - 1)
            Next c
        End If
    Next r
    If curIdx > 0 Then gKbBlocks(curIdx) = arr
    RestoreSnapshot = n
    Exit Function
Failed:
    LogKb "E0401", "RestoreSnapshot", "snapshot_read_failed"
    RestoreSnapshot = 0
End Function

Private Function KbOperator() As String
    On Error GoTo Blank0
    KbOperator = CStr(Application.UserName)
    Exit Function
Blank0:
    KbOperator = vbNullString
End Function

' 16章 E-13: 新サービス候補のローカル退避(隠しシートの待ち行列。見出し行なし)。
Private Sub EnqueuePending(ByVal stampText As String, ByVal caseId As String, _
                           ByVal industryCode As String, ByVal riskDesc As String, _
                           ByVal operatorName As String)
    On Error GoTo Ignore0
    Dim ws As Object
    Set ws = SheetOf(Nothing, KB_SHEET_PEND, True)
    If ws Is Nothing Then Exit Sub

    Dim vals As Variant
    vals = Array(stampText, caseId, industryCode, riskDesc, operatorName)
    Dim r As Long, c As Long
    r = LastRowOfWs(ws) + 1
    If r = 2 Then
        If LenB(Trim$(CStr(ws.Cells(1, 1).Value))) = 0 Then r = 1
    End If
    For c = 0 To 4
        modUtilText.SetCellSafe ws.Cells(r, c + 1), CStr(vals(c)), KB_SHEET_PEND
    Next c
    Exit Sub
Ignore0:
    Exit Sub
End Sub

' 退避分を 新サービス候補 へ流し込み、書けたら待ち行列を空にする。書けなければ
'   次回起動へ持ち越す(無音。E-13)。
Private Sub FlushPending()
    On Error GoTo Ignore0
    Dim ws As Object
    Set ws = SheetOf(Nothing, KB_SHEET_PEND, False)
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = LastRowOfWs(ws)
    Dim blk As Variant
    blk = ReadWsBlock(ws, lastRow, 5)
    If Not IsArray(blk) Then Exit Sub
    If LenB(Trim$(modKnowledge2.CellRaw(blk, 1, 1))) = 0 Then Exit Sub

    Dim wb As Object
    Set wb = OpenKbBook(False)
    If wb Is Nothing Then Exit Sub
    Dim gapWs As Object
    Set gapWs = SheetOf(wb, KB_SHEET_GAP, False)
    If gapWs Is Nothing Then
        CloseKbBook wb
        Exit Sub
    End If

    Dim hdr As Variant
    Dim cols As Variant
    hdr = ReadWsBlock(gapWs, 2, KB_SCAN_COLS)
    cols = Split(KB_GAPCOLS, ",")
    Dim wr As Long, r As Long, c As Long
    Dim wrote As Long
    wr = LastRowOfWs(gapWs)
    For r = 1 To lastRow
        If LenB(Trim$(modKnowledge2.CellRaw(blk, r, 1))) > 0 Then
            wr = wr + 1
            For c = 0 To 4
                If PutByName(gapWs, hdr, wr, CStr(cols(c)), modKnowledge2.CellRaw(blk, r, c + 1)) Then
                    wrote = wrote + 1
                End If
            Next c
        End If
    Next r
    wb.Save
    CloseKbBook wb
    ' 裁定書9 B19(16章 E-13): 書込成功が0件(見出し不在・列名変更で1セルも
    ' 書けていない)のときは退避キューを消さず次回起動へ持ち越す。
    If wrote > 0 Then ws.Cells.Clear
    Exit Sub
Ignore0:
    Exit Sub
End Sub

' 列名で位置を引いて1セル書く(SetCellSafe。NFR-S7①)。裁定書9 B19: 書けたか
'   どうかを返す(列が無ければ何もせず False。呼び出し側が成功件数を数える)。
Private Function PutByName(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                           ByVal colName As String, ByVal textValue As String) As Boolean
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Function
    modUtilText.SetCellSafe ws.Cells(r, c), textValue, KB_SHEET_GAP & "/" & colName
    PutByName = True
End Function
