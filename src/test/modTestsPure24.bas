Attribute VB_Name = "modTestsPure24"
Option Explicit

' ============================================================================
' modTestsPure24 - W14(裁定書37 班2)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**裁定書37 §1 の B-03 / B-05 / B-06 / D#10#11#12 の文と
'   18章§2・§3.5 だけ**から手で書き出した(17章§1。実装の出力を見てから期待値を
'   合わせない)。
'
' 対象と根拠(全23本):
'   G1 原文照合(modGround。B-03 のテスト観点6つをそのまま置いた)
'     01 引用の**先頭20字より後ろ**を1字変えても照合できる(表記揺れ耐性)
'     02 丸ごと捏造した引用は未照合として検出する
'     03 全角/半角・鍵括弧・空白の違いは吸収して照合できる
'     04 evidence.source が inference / knowledge のものは検査対象外
'     05 貼付原文が空なら**全件未照合で埋めない**(呼出側が ground_skipped)
'     06 20字未満の引用は全長一致(先頭だけの偶然一致を拾わない)
'     07 emerging_risks は risk_no を持たないので E1 / E2 の採番で拾う
'     08 NoteCount は ";" 区切りの件数(run_log の ground_unmatched=n の値)
'     09 NormalizeForMatch は**数字を落とさず**英字を小文字化する
'   G2 充足度(modPipeline3.SufficiencyNoteOf。B-05)
'     10/11/12 overall の3値 x status<>ok の観点数  13 読めなければ iq=?
'   G3 免責と確認フラグ(B-06。18章§3.5)
'     14/15 <noscript> の3項分岐  16 script 側も同じ meta.reviewed_by で分岐
'     17 表紙チップ  18 meta の reviewed_by / reviewed_at / ground_unmatched
'   G4 レポートCSS・辞書(D#10 表フォント / D#12 改ページ / D#11 英語タグ)
'     19/20/21
'   G5 SEC-14 の「原文照合」列(B-03)
'     22
'
' 変異注入(出来レース禁止・裁定書37 §2):
'   (a) modGround.QuoteFound を常に True にすると 02 と 07 が落ちる。
'   (b) modHtmlTemplate5.SecDisclaimerJs の3項分岐を確認済み側で固定すると
'       16 が落ち、modHtmlTemplate1 の noscript を固定すると 14 が落ちる。
'
' グループ単位の失敗隔離: modTestsPure22 と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ============================================================================

' 貼付原文の模擬(DR出力のコピペを想定)。02 の捏造引用はこの中に**無い**。
Private Const W14_HAY As String = _
    "当社は静岡県浜松市に本社を置き、和菓子と洋菓子の製造販売を行っています。" & _
    "2026年3月期の売上高は128億円、営業利益は7億4千万円でした。" & _
    "浜北工場では小麦粉とバターを主要原材料として使用しています。"

' 確認者(B-06)。
Private Const W14_BY As String = "浜松支店 山田"
Private Const W14_AT As String = "2026/09/12 10:05:00"

Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 5
        grpName = "W14-G" & CStr(i)
        On Error Resume Next
        Err.Clear
        RunGroup i
        If Err.Number <> 0 Then
            GroupFail grpName
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Sub

Private Sub RunGroup(ByVal idx As Long)
    Select Case idx
    Case 1: T_Ground
    Case 2: T_Sufficiency
    Case 3: T_Disclaimer
    Case 4: T_ReportCss
    Case 5: T_SourceColumn
    End Select
End Sub

Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

Private Sub ChkB(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check nm, cond, detail
End Sub

Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), "期待=[" & want & "] 実際=[" & act & "]"
End Sub

Private Function Ctn(ByVal hay As String, ByVal needle As String) As Boolean
    Ctn = (InStr(hay, needle) > 0)
End Function

' ============================================================================
' G1 原文照合(modGround)
' ============================================================================
Private Sub T_Ground()
    Dim s2 As String

    ' 01 先頭20字より後ろを1字だけ変えた引用(「128億円」->「129億円」)。
    '    正規化後の先頭20字は原文と一致するので照合できる。
    ChkB "Test_W14_01_引用の先頭20字より後ろの1字違いは照合できる_裁定書37B-03(1)", _
        modGround.QuoteFound("当社は静岡県浜松市に本社を置き、和菓子と洋菓子の製造卸売を" & _
                             "行っています。", W14_HAY, 20), _
        "原文は「製造販売」だが先頭20字が一致するので照合できる"

    ' 02 原文に無い文を丸ごと作った引用は見つからない。
    '    **QuoteFound を常に True にするとここが落ちる**(変異注入(a))。
    ChkB "Test_W14_02_丸ごと捏造した引用は照合できない_裁定書37B-03(2)", _
        (modGround.QuoteFound("従業員数は1200名で東証プライムに上場しています", _
                              W14_HAY, 20) = False), _
        "原文に無い文を「見つかった」にしない"

    ' 03 全角英数・鍵括弧・空白・読点の違いは正規化で吸収する。
    ChkB "Test_W14_03_全角半角と鍵括弧と空白の違いを吸収する_裁定書37B-03(3)", _
        modGround.QuoteFound("「２０２６年３月期の 売上高は１２８億円」", W14_HAY, 20), _
        "全角->半角・記号除去・空白除去のあとで一致させる"

    ' 04 evidence.source が inference / knowledge のものは検査しない。
    s2 = "{""risks"":[" & _
         RiskJson(1, "inference", "存在しない架空の記述です一般論として置いた仮定") & "," & _
         RiskJson(2, "knowledge", "社内ナレッジ由来の架空の記述であり原文には無い") & _
         "],""emerging_risks"":[]}"
    ChkS "Test_W14_04_inferenceとknowledgeは照合対象外_裁定書37B-03(4)", _
        modGround.GroundNotes(s2, W14_HAY, 20), ""

    ' 05 貼付原文が空なら**検査しない**(全件未照合で埋めない= fail-open)。
    s2 = "{""risks"":[" & RiskJson(1, "hp", "原文に無い架空の記述を置いた引用です") & _
         "],""emerging_risks"":[]}"
    ChkS "Test_W14_05_原文が空なら全件未照合で埋めない_裁定書37B-03(5)", _
        modGround.GroundNotes(s2, "", 20), ""
    ChkS "Test_W14_05b_原文があれば同じ入力で未照合を返す_裁定書37B-03(2)", _
        modGround.GroundNotes(s2, W14_HAY, 20), "1"

    ' 06 20字未満の引用は全長で見る(headChars で切って甘くしない)。
    ChkB "Test_W14_06_20字未満の引用は全長で一致を見る_裁定書37B-03", _
        (modGround.QuoteFound("浜北工場では小麦粉", W14_HAY, 20) = True) And _
        (modGround.QuoteFound("浜北工場では大豆", W14_HAY, 20) = False), _
        "短い引用を先頭だけの偶然一致で通さない"

    ' 07 emerging_risks は risk_no を持たない(15章 Schema-S2)ので E1/E2 で返す。
    s2 = "{""risks"":[],""emerging_risks"":[" & _
         EmergJson("hp", "浜北工場では小麦粉とバターを主要原材料として使用") & "," & _
         EmergJson("hp", "海外3拠点で半導体部品の調達を行っていると記載") & "]}"
    ChkS "Test_W14_07_emerging_risksはE1E2の採番で未照合を返す_裁定書37B-03", _
        modGround.GroundNotes(s2, W14_HAY, 20), "E2"

    ' 08 run_log の ground_unmatched=n の n。
    ChkB "Test_W14_08_NoteCountは区切りの件数を返す_裁定書37B-03", _
        (modGround.NoteCount("") = 0) And (modGround.NoteCount("3") = 1) And _
        (modGround.NoteCount("1;4;E2") = 3), _
        "0件/1件/3件"

    ' 09 正規化は数字を残し英字を小さくする(金額の桁が捏造の出どころ)。
    ChkS "Test_W14_09_正規化は数字を残し英字を小文字化し記号を落とす_裁定書37B-03", _
        modGround.NormalizeForMatch("ＡＢＣ－１２３、 「ＸＹＺ」"), "abc123xyz"
End Sub

' risks[] の1件(evidence.quote / evidence.source)。
Private Function RiskJson(ByVal noVal As Long, ByVal srcText As String, _
                          ByVal quoteText As String) As String
    RiskJson = "{""risk_no"":" & CStr(noVal) & ",""risk_name"":""x""," & _
               """evidence"":{""quote"":""" & quoteText & """,""source"":""" & _
               srcText & """}}"
End Function

' emerging_risks[] の1件(evidence_quote / evidence_source)。
Private Function EmergJson(ByVal srcText As String, ByVal quoteText As String) As String
    EmergJson = "{""risk_name"":""y"",""evidence_quote"":""" & quoteText & _
                """,""evidence_source"":""" & srcText & """}"
End Function

' ============================================================================
' G2 充足度(modPipeline3.SufficiencyNoteOf。B-05)
' ============================================================================
Private Sub T_Sufficiency()
    ChkS "Test_W14_10_overall_highで欠落0_裁定書37B-05", _
        modPipeline3.SufficiencyNoteOf(IqJson("high", "ok;ok;ok")), "iq=high;miss=0"

    ChkS "Test_W14_11_overall_midでpartialとmissingを数える_裁定書37B-05", _
        modPipeline3.SufficiencyNoteOf(IqJson("mid", "ok;partial;missing;ok")), _
        "iq=mid;miss=2"

    ChkS "Test_W14_12_overall_lowで全件欠落_裁定書37B-05", _
        modPipeline3.SufficiencyNoteOf(IqJson("low", "missing;missing;partial")), _
        "iq=low;miss=3"

    ChkB "Test_W14_13_読めなければiq疑問符で黙ってmidにしない_裁定書37B-05", _
        (modPipeline3.SufficiencyNoteOf("") = "iq=?") And _
        (modPipeline3.SufficiencyNoteOf("{""input_quality"":{}}") = "iq=?") And _
        (modPipeline3.SufficiencyNoteOf(IqJson("unknown", "ok")) = "iq=?"), _
        "空・overall欠落・enum外はすべて iq=?"
End Sub

' S1の input_quality だけを持つJSON。statusList は ";" 区切りの status 列。
Private Function IqJson(ByVal overallText As String, ByVal statusList As String) As String
    Dim it As Variant, body As String

    For Each it In Split(statusList, ";")
        If LenB(body) > 0 Then body = body & ","
        body = body & "{""aspect"":""profile"",""status"":""" & CStr(it) & """}"
    Next it
    IqJson = "{""company_name"":""x"",""input_quality"":{""coverage"":[" & body & _
             "],""overall"":""" & overallText & """,""advice"":""a""}}"
End Function

' ============================================================================
' G3 免責と確認フラグ(B-06。18章§3.5)
' ============================================================================
Private Sub T_Disclaimer()
    Dim docNo As String, docYes As String
    Dim nsNo As String, nsYes As String
    Dim js As String

    docNo = DocOf("", "")
    docYes = DocOf(W14_BY, W14_AT)
    nsNo = NoScriptOf(docNo)
    nsYes = NoScriptOf(docYes)
    js = modHtmlTemplate5.SecDisclaimerJs()

    ' 14 未確認の書き出し: 既定文だけが出て「確認・編集したものです」は出さない。
    '    **noscript を確認済み側で固定するとここが落ちる**(変異注入(b))。
    ChkB "Test_W14_14_noscriptは未確認ならAI生成担当者確認前を出す_裁定書37B-06", _
        (LenB(nsNo) > 0) And Ctn(nsNo, DiscDefault()) And _
        (Ctn(nsNo, "担当者が確認・編集したものです") = False), _
        "noscript=[" & Left$(nsNo, 120) & "]"

    ' 15 確認済みの書き出し: 確認者名と確認日時が入り、既定文は出ない。
    ChkB "Test_W14_15_noscriptは確認済みなら確認者名と日時を出す_裁定書37B-06", _
        Ctn(nsYes, "担当者が確認・編集したものです") And Ctn(nsYes, W14_BY) And _
        Ctn(nsYes, W14_AT) And (Ctn(nsYes, DiscDefault()) = False), _
        "noscript=[" & Left$(nsYes, 160) & "]"

    ' 16 script 側(SEC-15)も**同じ meta.reviewed_by** で切り替わる。文言が
    '    2箇所に複製されているため、片方だけ直る腐敗をここで止める。
    ChkB "Test_W14_16_SEC15も同じmeta_reviewed_byで3項分岐する_裁定書37B-06", _
        Ctn(js, "var rb=S(m.reviewed_by)") And Ctn(js, "m.reviewed_at") And _
        Ctn(js, "T(d,'p',null,rb?(") And _
        Ctn(js, DiscDefault()) And Ctn(js, "担当者が確認・編集したものです") And _
        Ctn(js, "引受可否・保険料・幹事構成を確約するものではありません"), _
        "SEC-15 の3項分岐と2行目の据え置き"

    ' 17 表紙(SEC-01)のチップ。
    ChkB "Test_W14_17_表紙チップが確認前と確認済で切り替わる_裁定書37B-06", _
        Ctn(modHtmlTemplate2.SecCoverJs(), "'確認済 '") And _
        Ctn(modHtmlTemplate2.SecCoverJs(), "'確認前'"), _
        "renderCover の CHIP"

    ' 18 DATA の meta(18章§2)。未確認は**両方とも空文字**でキーは必ず置く。
    ChkB "Test_W14_18_metaにreviewed_byとreviewed_atとground_unmatchedを置く_18章§2", _
        Ctn(MetaOf("", ""), """reviewed_by"":""""") And _
        Ctn(MetaOf("", ""), """reviewed_at"":""""") And _
        Ctn(MetaOf("", ""), """ground_unmatched"":[]") And _
        Ctn(MetaOf(W14_BY, W14_AT), """reviewed_by"":""" & W14_BY & """"), _
        "meta=[" & MetaOf("", "") & "]"
End Sub

' 18章§3.5 の既定文(裁定書37 B-06 の逐語)。
Private Function DiscDefault() As String
    DiscDefault = "本資料はAIが公開情報等から作成した営業担当者向けの分析資料です" & _
                  "（AI生成・担当者確認前）。お客さまへ提示する前に、担当者が内容を" & _
                  "確認・編集してください。"
End Function

Private Function MetaOf(ByVal reviewedBy As String, ByVal reviewedAt As String) As String
    MetaOf = modExportHtml.BuildMetaJson("C-20260912-001", "甲斐商店", "09", _
                                         "食料品製造業", "renewal", "t2_full", "deep", 1, _
                                         "proposal", "2026/09/12 10:05:00", "2.0.0", _
                                         "standard", "", reviewedBy, reviewedAt, "")
End Function

Private Function DocOf(ByVal reviewedBy As String, ByVal reviewedAt As String) As String
    DocOf = modExportHtml.BuildReportHtml(MetaOf(reviewedBy, reviewedAt), _
                                          "{""company_name"":""甲斐商店""}", "", "", _
                                          "standard")
End Function

' <noscript> ... </noscript> の中身だけを返す(静的HTML側の分岐を見るため)。
Private Function NoScriptOf(ByVal doc As String) As String
    Dim a As Long, b As Long

    a = InStr(doc, "<noscript>")
    If a = 0 Then Exit Function
    a = a + Len("<noscript>")
    b = InStr(a, doc, "</noscript>")
    If b = 0 Then Exit Function
    NoScriptOf = Mid$(doc, a, b - a)
End Function

' ============================================================================
' G4 レポートCSS・辞書(D#10 / D#12 / D#11)
' ============================================================================
Private Sub T_ReportCss()
    Dim css As String, t2 As String

    css = modHtmlTemplate7.CommonCss()
    t2 = modHtmlTemplate2.SecExecJs() & modHtmlTemplate2.SecProfileJs()

    ' 19 表のフォント(画面13px/12px・印刷 11pt/9.5pt と table-layout:fixed)。
    ChkB "Test_W14_19_表フォントと印刷サイズとtable_layout_裁定書37D#10", _
        Ctn(css, "table{border-collapse:collapse;width:100%;min-width:960px;font-size:13px}") And _
        Ctn(css, "font-size:12px}") And Ctn(css, "body{background:#fff;font-size:11pt") And _
        Ctn(css, "table{font-size:9.5pt;min-width:0;table-layout:fixed;width:100%}") And _
        Ctn(css, "word-wrap:break-word"), _
        "画面13px/12px・印刷11pt/9.5pt・fixed・word-wrap"

    ' 20 印刷の改ページ(孤立見出しを作らない・まとまりを割らない)。
    ChkB "Test_W14_20_印刷の改ページ規則_裁定書37D#12", _
        Ctn(css, ".section-head,.sec h2,.sec h3{break-after:avoid;page-break-after:avoid}") And _
        Ctn(css, ".card,.story,.tblwrap,.rublock,.idea{break-inside:avoid;page-break-inside:avoid}"), _
        "見出しは page-break-after:avoid / カード等は page-break-inside:avoid"

    ' 21 英語タグを日本語へ(19章§3に定義のある enum ラベルは触らない)。
    ChkB "Test_W14_21_英語タグを日本語へ_裁定書37D#11", _
        (Ctn(t2, "Public facts") = False) And (Ctn(t2, "Risk implications") = False) And _
        Ctn(t2, "公開情報に基づく事実") And Ctn(t2, "想定される事業影響"), _
        "Public facts / Risk implications を廃止"
End Sub

' ============================================================================
' G5 SEC-14 の「原文照合」列(B-03)
' ============================================================================
Private Sub T_SourceColumn()
    Dim js As String

    js = modHtmlTemplate5.SecSourceJs()
    ChkB "Test_W14_22_SEC14に原文照合の列と原文未照合の値を出す_裁定書37B-03", _
        Ctn(js, "'原文照合'") And Ctn(js, "'原文未照合'") And _
        Ctn(js, "ground_unmatched") And _
        Ctn(js, "['No','リスク名','引用した記述','出所','原文照合']"), _
        "SEC-14 の見出し5列と meta.ground_unmatched の参照"
End Sub
