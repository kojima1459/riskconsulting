Attribute VB_Name = "modKnowledgeFmt"
Option Explicit

' ============================================================================
' modKnowledgeFmt - ナレッジ注入テキストの整形と切詰め(純関数のみ・T-21)
' ----------------------------------------------------------------------------
' 正: 14章§6のmodKnowledgeFmt節 / 15章§3・§4・§6.1(1行書式) / 15章§0.7(切詰め)
'   / 16章 E-09(0行の既定文言)。
' 位置づけ(12章§2・裁定書6 項目6): modKnowledge から「整形」だけを切り出した
'   app層の純文字列モジュール。**シート・config・ログに一切触れない**ので
'   層(a)のテストから直接叩ける。W2aでは整形が modKnowledge の Private に
'   閉じており、項目区切りを " | " から " / " へ壊してもどのゲートも
'   気づかなかった。整形の実装をここ1箇所に置き、テストの的にする。
'
' 共通の引数 rows: **Range.Value 由来の2次元Variant配列**(1行目=見出し行=列名、
'   2行目以降=データ行)。列は列名で引く(13章冒頭・modUtil.FindHeaderCol)。
'   渡された全データ行を整形する(業種絞込・is_active・status・行数上限は
'   modKnowledge が適用済みで渡す)。配列でない/データ行0件は既定文言を返す。
' 共通の戻り値: 15章の書式の複数行文字列(1行1件・行区切り vbLf・末尾改行なし)。
'   行頭は "[ID] "、項目区切りは " | "、項目内の複数値は ";"、**値が空の項目は
'   項目ごと省略**する(market_note の空欄省略もこの規約の一適用)。
'
' R4準拠(12章§2): Worksheets / Range( / Application. / ThisWorkbook / MsgBox /
'   ActiveSheet に一切触れない。
' CP932準拠(15章§0 原則7): 本文・注釈ともに CP932 内の文字だけで書く。
' ============================================================================

' --- 1行書式(15章§3/§4/§6.1)。"ラベル=列名;..." 形式のスペック ---
'     ラベルが空なら値だけを出す。値が空の項目は項目ごと省略する。
Private Const FM_SP_RISK As String = "カテゴリ=category;リスク=risk_name;典型シナリオ=typical_scenario;典型頻度=typical_freq;典型影響=typical_impact;確認点=check_points"
Private Const FM_SP_CASE As String = "業種=industry_code;顧客像=customer_profile;提示リスク=risk_presented;提案=proposal;決め手=why_it_worked"
Private Const FM_SP_RULE As String = "class=rule_class;基準=rule_text"
Private Const FM_SP_RT As String = "=theme_name;status=status"
Private Const FM_SP_MECH As String = "layer=layer;機構=mech_text"
Private Const FM_SP_MENU As String = "=menu_name"
Private Const FM_SP_LINE As String = "=line_name"
Private Const FM_SP_SCHEME As String = "=scheme_name"
Private Const FM_SP_PAT As String = "=pattern_name"
Private Const FM_PP_MENU2 As String = "概要=summary;対応カテゴリ=target_categories"
Private Const FM_PP_MENU1 As String = "対応カテゴリ=target_categories"
Private Const FM_PP_LINE As String = "市場環境=market_note"
Private Const FM_PP_SCHEME As String = "構造=structure;成立条件=conditions;適用シグナル=signals"
Private Const FM_PP_PAT As String = "構造=structure;成立条件=conditions;代表例=examples_public;社内実績=internal_refs"
Private Const FM_PP_RULE As String = "破り方=workaround"
Private Const FM_PP_RT As String = "メモ=note"
Private Const FM_PP_MECH As String = "適用リスク=target_categories"

' --- 0行・未装填時の値(§6.1の表 / 15章§3のriskLib専用文言 / 15章§4の見出し) ---
Private Const FM_NONE As String = "(登録なし)"
Private Const FM_NONE_RISK As String = "(この業種の登録知識はまだありません)"
Private Const FM_NONE_S3 As String = "なし"

Private Const FM_PIPE As String = " | "
Private Const FM_SPACE As String = " "
Private Const FM_SEMI As String = ";"

' 15章§0.7 最終段: 各行を先頭400字で切り「…」を付す。
Private Const FM_LINE_MAX As Long = 400
Private Const FM_ELLIPSIS As String = "…"

' 15章§0.7 ナレッジ側の切詰め順(成功事例→型→メニュー→種目→リスクライブラリ)
' の下限行数。メニュー・種目・リスクは5行未満にするとS3のID実在制約と
' §0.5第2層が崩れるため0にできない。
Private Const FM_PLAN_N As Long = 5
Private Const FM_FLOORS As String = "0;0;5;5;5"

' === 公開: 種類別の整形(14章§6のmodKnowledgeFmt節。ここに無い名前は公開しない) ===

' FmtRiskLib - 15章§3 riskLibText。0行は業種専用の文言(16章 E-09)。
Public Function FmtRiskLib(ByVal rows As Variant) As String
    FmtRiskLib = RowsText(rows, "risk_lib_id", vbNullString, FM_SP_RISK, vbNullString, FM_NONE_RISK)
End Function

' FmtMenus - 15章§4 menusText(S3用。概要つき)。
Public Function FmtMenus(ByVal rows As Variant) As String
    FmtMenus = RowsText(rows, "menu_id", vbNullString, FM_SP_MENU, FM_PP_MENU2, FM_NONE)
End Function

' FmtMenusSummary - 15章§3・§6.1 menusSummary(S2/PF用。概要を出さない要約版)。
Public Function FmtMenusSummary(ByVal rows As Variant) As String
    FmtMenusSummary = RowsText(rows, "menu_id", vbNullString, FM_SP_MENU, FM_PP_MENU1, FM_NONE)
End Function

' FmtLines - 15章§4 linesText。market_note が空の行は " | 市場環境:" ごと省略。
Public Function FmtLines(ByVal rows As Variant) As String
    FmtLines = RowsText(rows, "line_id", vbNullString, FM_SP_LINE, FM_PP_LINE, FM_NONE)
End Function

' FmtCases - 15章§4 casesText。0行はS3 userの見出し規約どおり「なし」。
Public Function FmtCases(ByVal rows As Variant) As String
    FmtCases = RowsText(rows, "case_lib_id", vbNullString, FM_SP_CASE, vbNullString, FM_NONE_S3)
End Function

' FmtSchemes - 15章§4 schemesText。行頭名称のうしろに "(pattern_id)" を付す。
Public Function FmtSchemes(ByVal rows As Variant) As String
    FmtSchemes = RowsText(rows, "scheme_id", "pattern_id", FM_SP_SCHEME, FM_PP_SCHEME, FM_NONE_S3)
End Function

' FmtPatterns - 15章§6.1 patternsText(P1-P15)。
Public Function FmtPatterns(ByVal rows As Variant) As String
    FmtPatterns = RowsText(rows, "pattern_id", vbNullString, FM_SP_PAT, FM_PP_PAT, FM_NONE)
End Function

' FmtRules - 15章§6.1 rulesText(判断基準)。
Public Function FmtRules(ByVal rows As Variant) As String
    FmtRules = RowsText(rows, "rule_id", vbNullString, FM_SP_RULE, FM_PP_RULE, FM_NONE)
End Function

' FmtResearching - 15章§6.1 researchingText(研究テーマ一覧)。書式例の「判定日」
'   「関連」は13章§3.10に供給元の列が無いため空項目の省略規約に従い出さない。
Public Function FmtResearching(ByVal rows As Variant) As String
    FmtResearching = RowsText(rows, "rt_id", vbNullString, FM_SP_RT, FM_PP_RT, FM_NONE)
End Function

' FmtMechs - 15章§6.1 mechs(機構ライブラリ抜粋)。Phase1は常に0行=「(登録なし)」。
Public Function FmtMechs(ByVal rows As Variant) As String
    FmtMechs = RowsText(rows, "mech_id", vbNullString, FM_SP_MECH, FM_PP_MECH, FM_NONE)
End Function

' ============================================================================
' TrimKbLine - 15章§0.7 最終段「各行を先頭400字で切り『…』を付す」。
'   400字以内は**何も足さずそのまま返す**。超過時は先頭400字(サロゲート安全)へ
'   「…」を付けるので戻り値は最大401字になる。
' ============================================================================
Public Function TrimKbLine(ByVal s As String) As String
    If Len(s) <= FM_LINE_MAX Then
        TrimKbLine = s
        Exit Function
    End If
    TrimKbLine = modUtil.SafeLeft(s, FM_LINE_MAX) & FM_ELLIPSIS
End Function

' ============================================================================
' TrimPlan - 15章§0.7 ナレッジ側の切詰めを**計画するだけ**の純関数。
'   counts: 10要素。前半(0..4)=現在の行数、後半(5..9)=現在の文字数。並びは
'     切詰め順 成功事例→型→メニュー→種目→リスクライブラリ。5要素以下のときは
'     文字数0とみなす(=切詰め不要)。
'   budgetChars: ナレッジ注入に許される合計文字数(§0.7「上限の3割」)。0以下は
'     上限なしとして現在の行数をそのまま返す。
'   戻り値: 5要素の「注入してよい行数」。1段ずつ順に適用し、そのつど総量を
'     再計算して budgetChars 以下になった時点で止める(1対象あたり半減は1回)。
'     半減は端数切上げ、下限は 0/0/5/5/5 行。
' ============================================================================
Public Function TrimPlan(ByRef counts() As Long, ByVal budgetChars As Long) As Long()
    Dim res() As Long
    ReDim res(0 To FM_PLAN_N - 1)

    Dim rowsNow(0 To 4) As Long
    Dim charsNow(0 To 4) As Double
    Dim floors As Variant
    floors = Split(FM_FLOORS, FM_SEMI)

    Dim i As Long
    Dim total As Double
    For i = 0 To FM_PLAN_N - 1
        rowsNow(i) = LongAt(counts, i)
        charsNow(i) = CDbl(LongAt(counts, FM_PLAN_N + i))
        res(i) = rowsNow(i)
        total = total + charsNow(i)
    Next i

    If budgetChars <= 0 Then
        TrimPlan = res
        Exit Function
    End If

    Dim newRows As Long
    Dim newChars As Double
    For i = 0 To FM_PLAN_N - 1
        If total <= CDbl(budgetChars) Then Exit For
        newRows = (rowsNow(i) + 1) \ 2
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

' === 内部(すべて純関数) ===

' 唯一の整形ドライバ。rows の2行目以降を1行1件で積む。
Private Function RowsText(ByVal rows As Variant, ByVal idCol As String, ByVal sufCol As String, _
                          ByVal spaceSpec As String, ByVal pipeSpec As String, _
                          ByVal emptyText As String) As String
    If Not IsArray(rows) Then
        RowsText = emptyText
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = RowCount(rows)
    Dim cId As Long, cSuf As Long
    cId = modUtil.FindHeaderCol(rows, idCol)
    If LenB(Trim$(sufCol)) > 0 Then cSuf = modUtil.FindHeaderCol(rows, sufCol)

    Dim outText As String, idText As String, headText As String, sufText As String
    Dim r As Long, n As Long
    For r = 2 To lastRow
        idText = CellAt(rows, r, cId)
        If LenB(idText) > 0 Then
            headText = BodyOf(rows, r, spaceSpec, FM_SPACE)
            If cSuf > 0 Then
                sufText = CellAt(rows, r, cSuf)
                If LenB(sufText) > 0 Then headText = headText & "(" & sufText & ")"
            End If
            AppendPart outText, vbLf, RowLine(rows, r, idText, headText, pipeSpec)
            n = n + 1
        End If
    Next r

    If n = 0 Then outText = emptyText
    RowsText = outText
End Function

' 1件ぶんの行。"[ID] 見出し | 項目:値 | 項目:値"。
Private Function RowLine(ByVal rows As Variant, ByVal r As Long, ByVal idText As String, _
                         ByVal headText As String, ByVal pipeSpec As String) As String
    Dim lineText As String
    lineText = "[" & idText & "]"
    If LenB(headText) > 0 Then lineText = lineText & FM_SPACE & headText
    AppendPart lineText, FM_PIPE, BodyOf(rows, r, pipeSpec, FM_PIPE)
    RowLine = lineText
End Function

' 書式スペック("ラベル=列名;...")を1本の文字列へ。値が空の項目は丸ごと落とす。
Private Function BodyOf(ByVal rows As Variant, ByVal r As Long, ByVal specText As String, _
                        ByVal sepText As String) As String
    If LenB(specText) = 0 Then Exit Function
    Dim parts As Variant, pair As Variant
    Dim i As Long, acc As String
    parts = Split(specText, FM_SEMI)
    For i = LBound(parts) To UBound(parts)
        pair = Split(CStr(parts(i)), "=")
        If UBound(pair) >= 1 Then
            AppendPart acc, sepText, ItemText(CStr(pair(0)), _
                       CellAt(rows, r, modUtil.FindHeaderCol(rows, CStr(pair(1)))))
        End If
    Next i
    BodyOf = acc
End Function

' 「ラベル:値」。ラベルが空なら値だけ。値が空なら項目ごと省略(空文字を返す)。
Private Function ItemText(ByVal labelText As String, ByVal valueText As String) As String
    If LenB(valueText) = 0 Then Exit Function
    If LenB(labelText) = 0 Then
        ItemText = valueText
    Else
        ItemText = labelText & ":" & valueText
    End If
End Function

' 非空の部品だけを区切りでつなぐ(先頭には区切りを付けない)。
Private Sub AppendPart(ByRef acc As String, ByVal sepText As String, ByVal partText As String)
    If LenB(partText) = 0 Then Exit Sub
    If LenB(acc) = 0 Then
        acc = partText
    Else
        acc = acc & sepText & partText
    End If
End Sub

' 1セルをテキストへ。改行・タブはスペースへ畳み(1行1件の規約)、外部由来なので
' SanitizeInput を通す(16章 E-04・E-43)。列が0(不在)なら空文字。
Private Function CellAt(ByVal rows As Variant, ByVal r As Long, ByVal c As Long) As String
    On Error GoTo Blank0
    If c <= 0 Then Exit Function
    Dim t As String
    t = CStr(rows(r, c))
    If LenB(t) = 0 Then Exit Function
    t = Replace(Replace(Replace(Replace(t, vbCrLf, FM_SPACE), vbCr, FM_SPACE), vbLf, FM_SPACE), vbTab, FM_SPACE)
    CellAt = Trim$(modUtilText.SanitizeInput(t))
    Exit Function
Blank0:
    CellAt = vbNullString
End Function

' 2次元配列の行数(1始まり前提。取れなければ0)。
Private Function RowCount(ByVal rows As Variant) As Long
    On Error GoTo Zero0
    RowCount = UBound(rows, 1)
    Exit Function
Zero0:
    RowCount = 0
End Function

' counts(LBound+i) を安全に読む(範囲外・未初期化は0)。
Private Function LongAt(ByRef arr() As Long, ByVal i As Long) As Long
    On Error GoTo Zero0
    Dim lo As Long, hi As Long
    lo = LBound(arr)
    hi = UBound(arr)
    If lo + i > hi Then Exit Function
    LongAt = arr(lo + i)
    Exit Function
Zero0:
    LongAt = 0
End Function
