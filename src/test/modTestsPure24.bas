Attribute VB_Name = "modTestsPure24"
Option Explicit

' ============================================================================
' modTestsPure24 - W14(裁定書37 班2)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**裁定書37 §1 の B-03 / B-05 / B-06 / D#10#11#12 の文と
'   18章§2・§3.5 だけ**から手で書き出した(17章§1。実装の出力を見てから期待値を
'   合わせない)。
'
' 対象と根拠(全36本。W15 Round 2 の G6 10本は裁定書39 §1・裁定書40 §1 だけを、
'   G7 3本は裁定書42 §2 と裁定書43 §2 だけを根拠に追記した):
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
'     22(生成JSの照合。**描画の両方向**は tools/notice_check.py の検査①が
'        持つ。裁定書40 Q-M2 の横展開で SEC-09 と同じ型だったため足した)
'   G6 W15 Round 2(裁定書39 班Q / 裁定書40 班Q2)
'     23/24 R2-04  25/26 R2-05  27 R1-05  28 R1-08(Q-M2で論理へ)  29/30 R1-07
'     31 Q-m1(E-02の帯は実行直後のS1だけ)  32 Q-m3(E-02の帯とdeepを併記)
'   G7 W15 最終是正(裁定書42 §2 確認導線の一本化 / 裁定書43 §2 Y-2)
'     33 見えない文字に NBSP/ZWSP/BOM を含める  34 レポートと提案書の判定が対称
'     35 見えない文字は**範囲**で決まる(範囲の境界の内外。列挙ではない)
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

' 偽の確認日時(裁定書39 R2-05)。確認者名に vbTab を混ぜると <noscript> の免責へ
'   任意の日時を差し込めた。W15_CASE_ID は表紙のフィールドずれ(26)の照合用。
Private Const W15_FAKE_AT As String = "9999/12/31 00:00:00"
Private Const W15_CASE_ID As String = "C-20260912-001"

Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 7
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
    Case 6: T_W15Round2
    Case 7: T_W15Final
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

' ============================================================================
' G6 W15 Round 2(裁定書39 班Q): R2-04 / R2-05 / R1-05 / R1-07 / R1-08
' ----------------------------------------------------------------------------
' 期待値の根拠は**裁定書39 §1 の裁定欄と 16章 E-02 / 18章§2・§3・§4.1 の逐語**
'   だけで、実装の出力を見てから合わせたものは1つも無い。
'   23/24 R2-04 確認者名の空判定(空白類を除いて1文字以上)
'   25/26 R2-05 表紙の vbTab 区切りをフィールドの中身が騙れない
'   27    R1-05 原文(haystack)を1回だけ組み立てて使い回す
'   28    R1-08 SEC-09 にニューリスクの「原文未照合」を出す
'   29/30 R1-07 シート画面にも 16章 E-02 の警告帯を出す
'
' 変異注入(出来レース禁止・裁定書39 §3):
'   (c) modUtilText.HasVisibleText を `LenB(Trim$(s))>0` へ戻すと 23/24 が落ちる。
'   (d) modExportHtml.StripFieldSeps の vbTab 除去をやめると 25/26 が落ちる。
'   (e) modGround.NoteJsonArray を常に "" にすると 28 が落ちる(材料側)。
'       描画側(SEC-09 のカードに印が出る/出ない)の変異は tools/notice_check.py
'       の検査①が受け持つ(ページ内JSは VBA から走らせられないため)。
'   (f) modUICase.IqBannerTextOf を常に "" にすると 29 が落ちる。
'   (g) modUICase2.StepNoticeOf の afterRun ガードを外すと 31 が落ちる。
'   (h) modUICase2.NoticeJoin を secondText だけ返す形に戻すと 32 が落ちる。
' ============================================================================

Private Sub T_W15Round2()
    T_R2_04
    T_R2_05
    T_R1_05
    T_R1_08
    T_R1_07
End Sub

' ---- R2-04 確認者名の空判定 ------------------------------------------------
Private Sub T_R2_04()
    ' 23 空白類(半角/全角/TAB/LF/CR)だけなら「入力あり」にしない。
    '    Trim$ は Chr(32) しか落とさないので、この4つが素通りしていた。
    ChkB "Test_W15_23_HasVisibleTextは空白類だけをFalseにする_裁定書39R2-04", _
        (modUtilText.HasVisibleText("") = False) And _
        (modUtilText.HasVisibleText(" ") = False) And _
        (modUtilText.HasVisibleText("　") = False) And _
        (modUtilText.HasVisibleText(vbTab) = False) And _
        (modUtilText.HasVisibleText(vbLf) = False) And _
        (modUtilText.HasVisibleText(vbCr) = False) And _
        (modUtilText.HasVisibleText(vbTab & "　" & vbLf & vbCr & " ") = False) And _
        (modUtilText.HasVisibleText("山") = True) And _
        (modUtilText.HasVisibleText(vbTab & "山田" & vbLf) = True), _
        "空白類=False / 1文字でも可視文字があれば True"

    ' 24 レポート側の唯一の判断点。空白類だけの確認者名は**未確認**にする。
    ChkB "Test_W15_24_ReviewerOfは空白類だけの確認者名を未確認にする_裁定書39R2-04", _
        (modExportHtml.ReviewerOf(vbTab) = "") And _
        (modExportHtml.ReviewerOf("　") = "") And _
        (modExportHtml.ReviewerOf(vbLf & vbCr) = "") And _
        (modExportHtml.ReviewerOf(" ") = "") And _
        (modExportHtml.ReviewerOf(" 山田 ") = "山田"), _
        "TAB=[" & modExportHtml.ReviewerOf(vbTab) & "] 全角空白=[" & _
        modExportHtml.ReviewerOf("　") & "] 氏名=[" & _
        modExportHtml.ReviewerOf(" 山田 ") & "]"
End Sub

' ---- R2-05 表紙の区切りを騙れない ------------------------------------------
Private Sub T_R2_05()
    Dim ns As String, doc As String

    ' 25 確認者名に vbTab を入れても <noscript> の確認日時は meta.reviewed_at の
    '    ままであること(偽の日時に差し替えられない)。18章§4.1 の3項分岐。
    ns = NoScriptOf(DocOf("山田" & vbTab & W15_FAKE_AT, W14_AT))
    ChkB "Test_W15_25_確認者名のTABで確認日時を偽装できない_裁定書39R2-05", _
        Ctn(ns, W14_AT) And Ctn(ns, "担当者が確認・編集したものです"), _
        "noscript=[" & Left$(ns, 200) & "]"

    ' 26 会社名に vbTab が混じっても表紙の案件IDがずれないこと。
    doc = DocOfCompany("甲斐" & vbTab & "商店")
    ChkB "Test_W15_26_会社名のTABで表紙のフィールドがずれない_裁定書39R2-05", _
        Ctn(doc, "案件ID " & W15_CASE_ID), _
        "案件IDのピルが会社名の後半で上書きされない"
End Sub

' ---- R1-05 原文(haystack)の1回組み立て ------------------------------------
Private Sub T_R1_05()
    ' 27 S2照合用の原文 = 貼付原文 + S1出力(modPipeline3.BuildHaystack の
    '    `sb & s1Json` と同値)。この同値が崩れると使い回しが誤りになる。
    ChkB "Test_W15_27_GroundHaystackは貼付原文の末尾にS1を足したもの_裁定書39R1-05", _
        (modExportHtml.GroundHaystack("原文", "{""a"":1}") = "原文{""a"":1}") And _
        (modExportHtml.GroundHaystack("原文", "") = "原文") And _
        (modExportHtml.GroundHaystack("", "{}") = "{}"), _
        "実際=[" & modExportHtml.GroundHaystack("原文", "{""a"":1}") & "]"
End Sub

' ---- R1-08 SEC-09 の「原文未照合」 ------------------------------------------
Private Sub T_R1_08()
    ' 28 SEC-09 が読む材料= meta.ground_unmatched(18章§2)。**両方向**を固定
    '    する: 未照合が在れば modGround の E採番のまま配列に載り、無ければ
    '    空配列になる。E採番そのもの(出現順の1始まり)は 07 が持つ。
    '    裁定書40 Q-M2: 旧28は生成済みJSの**文字列 grep 3本**だけで、描画側の
    '    判定を反転させる変異(`if(gm[...])` -> `if(!gm[...])`)が全ゲートを
    '    素通りした。描画そのものの両方向(E1が在るカードにだけ印が出て、
    '    無ければ出ない)は**ページ内JSなので VBA からは走らせられない**ため、
    '    実DOMの回帰は `tools/notice_check.py` の検査①が持つ。ここは
    '    「JSへ渡す材料」の論理だけを固定する(2本で1つの契約を挟む)。
    ChkB "Test_W15_28_metaのground_unmatchedはE採番を載せ無ければ空配列_裁定書40Q-M2", _
        Ctn(MetaOfGround("E1"), """ground_unmatched"":[""E1""]") And _
        Ctn(MetaOfGround("E1;E2"), """ground_unmatched"":[""E1"",""E2""]") And _
        Ctn(MetaOfGround(""), """ground_unmatched"":[]"), _
        "meta=[" & MetaOfGround("E1") & "]"
End Sub

' 未照合リスト(";" 区切り)だけを差し替えた meta(28用)。
Private Function MetaOfGround(ByVal groundNote As String) As String
    MetaOfGround = modExportHtml.BuildMetaJson(W15_CASE_ID, "甲斐商店", "09", _
                                               "食料品製造業", "renewal", "t2_full", _
                                               "deep", 1, "proposal", _
                                               "2026/09/12 10:05:00", "2.0.0", _
                                               "standard", "", "", "", groundNote)
End Function

' ---- R1-07 シート画面の iq=low 警告帯 --------------------------------------
Private Sub T_R1_07()
    Dim lowText As String

    ' 29 iq=low のときだけ 16章 E-02(実行後)の1文を返す。
    lowText = modUICase.IqBannerTextOf(IqJson("low", "missing;partial"))
    ChkB "Test_W15_29_シート画面もiq_lowでE02の警告文を出す_裁定書39R1-07", _
        Ctn(lowText, "一般論に近い出力になります") And Ctn(lowText, "助言"), _
        "banner=[" & lowText & "]"

    ' 30 low 以外・読めないときは**何も出さない**(他の警告を消さない)。
    ChkB "Test_W15_30_iq_low以外では警告帯を出さない_裁定書39R1-07", _
        (modUICase.IqBannerTextOf(IqJson("high", "ok")) = "") And _
        (modUICase.IqBannerTextOf(IqJson("mid", "partial")) = "") And _
        (modUICase.IqBannerTextOf("") = ""), _
        "high/mid/読めない はすべて空"

    ' 31 16章 E-02 は「**実行後**」の規定。modUICase2.DrawStep は
    '    RunStepUi / HomeRunAll 以外([シートで編集]・案件切替の描き直し・
    '    企業ファイル取込)からも呼ばれるので、afterRun=False では iq=low でも
    '    **何も出さない**(裁定書40 Q-m1)。S1 以外の段でも出さない。
    ChkB "Test_W15_31_E02の帯は実行直後のS1だけに出す_裁定書40Q-m1", _
        (modUICase2.StepNoticeOf(1, IqJson("low", "missing"), False) = "") And _
        (modUICase2.StepNoticeOf(2, IqJson("low", "missing"), True) = "") And _
        (modUICase2.StepNoticeOf(4, IqJson("low", "missing"), True) = "") And _
        (modUICase2.StepNoticeOf(1, IqJson("high", "ok"), True) = "") And _
        Ctn(modUICase2.StepNoticeOf(1, IqJson("low", "missing"), True), _
            "一般論に近い出力になります"), _
        "描き直し=[" & modUICase2.StepNoticeOf(1, IqJson("low", "missing"), False) & _
        "] 実行直後=[" & modUICase2.StepNoticeOf(1, IqJson("low", "missing"), True) & "]"

    ' 32 シート画面は警告欄(hm_warning)が1枠しかないので、あとから書く1本が
    '    前の1本を消す。deep の警告(E-35/E-36)は E-02 の帯を、部屋あふれの
    '    警告は同じく E-02 の帯を消していた(裁定書40 Q-m3)。片方しか無ければ
    '    その1本、両方あれば vbLf で**併記**する。順序は HTML画面
    '    (modNaviActions.ActRunPipeline)と同じで E-02 の帯が先。
    ChkB "Test_W15_32_E02の帯とdeepの警告を併記して片方を消さない_裁定書40Q-m3", _
        (modUICase2.NoticeJoin("帯", "deep") = "帯" & vbLf & "deep") And _
        (modUICase2.NoticeJoin("", "deep") = "deep") And _
        (modUICase2.NoticeJoin("帯", "") = "帯") And _
        (modUICase2.NoticeJoin("", "") = ""), _
        "両方=[" & modUICase2.NoticeJoin("帯", "deep") & "]"
End Sub

' 会社名だけを差し替えたレポート全文(R2-05 の26用)。
Private Function DocOfCompany(ByVal company As String) As String
    Dim metaJson As String
    metaJson = modExportHtml.BuildMetaJson(W15_CASE_ID, company, "09", _
                                           "食料品製造業", "renewal", "t2_full", "deep", 1, _
                                           "proposal", "2026/09/12 10:05:00", "2.0.0", _
                                           "standard", "", "", "", "")
    DocOfCompany = modExportHtml.BuildReportHtml(metaJson, _
                                                 "{""company_name"":""x""}", "", "", "standard")
End Function

' ============================================================================
' G7 W15 最終是正(裁定書42 §2 確認導線の一本化 / 裁定書43 §2 Y-2)。
' ----------------------------------------------------------------------------
'   33 見えない文字の定義は modUtilText.HasVisibleText が唯一持ち、NBSP
'      (U+00A0)・ZWSP(U+200B)・BOM(U+FEFF)・SOFT HYPHEN(U+00AD)・
'      EN SPACE(U+2002)・ZWNJ(U+200C)・WORD JOINER(U+2060)も見えない文字
'      として数える。
'   34 **レポートと提案書の判定が対称**であること。33 と同じ標本表を使い、
'      全件で「レポートは未確認・提案書は生成しない」を同時に見る。
'      裁定書40 S-m では提案書側だけへ NBSP の前処理を足したため、NBSP だけの
'      確認者名でレポートが「担当者が確認・編集したもの」に切り替わっていた
'      (=「片方だけ直す」型の3回目)。表を1つにすれば非対称は作れない。
'   35 見えない文字が**閉じた範囲の集合**で決まること(裁定書43 §0)。33/34 が
'      使う標本表は**定義ではなく標本**であり、表を長くしても網羅にはならない
'      (裁定書43 までに3回それで漏れた)。35 は範囲の**境界の内外**だけを見る:
'      範囲の内側の端は必ず False、その1つ外は必ず True。1文字ずつは列挙しない。
'   変異注入: modUtilText.IsInvisibleCodeUnit の範囲を1つ削る/端を1つずらすと
'      35 が落ちる(削った範囲が 33/34 の標本を含めば 33/34 も落ちる)。
'      modExportProposal.NeedsReviewMessage に前処理(Replace)を挟み直すと
'      34 が落ちる(レポート側だけが通す形に戻るため)。
' ============================================================================

' 見えない文字の**標本**表(定義ではない)。33 が「1つも可視にならない」ことを、
'   34 が「レポートと提案書で同じ答えになる」ことをこの表で見る。定義そのものは
'   modUtilText.IsInvisibleCodeUnit の範囲で、その網羅は 35 が境界で押さえる。
Private Function BlankKinds() As Variant
    BlankKinds = Array(" ", "　", vbTab, vbLf, vbCr, Chr$(11), Chr$(12), _
                       ChrW$(&HA0&), ChrW$(&HAD&), ChrW$(&H2002&), _
                       ChrW$(&H200B&), ChrW$(&H200C&), ChrW$(&H2060&), _
                       ChrW$(&HFEFF&))
End Function

' 範囲の**内側の端**(すべて見えない文字でなければならない)。U+0000 側は
'   符号なしなので下限の外が存在しない=U+0001 を内側の端として持つ。
Private Function InvisibleEdges() As Variant
    InvisibleEdges = Array(&H1&, &H20&, &H7F&, &HA0&, &HAD&, &H180E&, _
                           &H2000&, &H200F&, &H2028&, &H202F&, &H205F&, _
                           &H2060&, &H3000&, &HFEFF&)
End Function

' 上の範囲の**1つ外**(すべて可視でなければならない=落としすぎの検出)。
Private Function VisibleEdges() As Variant
    VisibleEdges = Array(&H21&, &H7E&, &H80&, &H9F&, &HA1&, &HAC&, &HAE&, _
                         &H180D&, &H180F&, &H1FFF&, &H2010&, &H2027&, _
                         &H2030&, &H205E&, &H2061&, &H2FFF&, &H3001&, _
                         &HFEFE&, &HFF00&)
End Function

Private Sub T_W15Final()
    Dim kinds As Variant
    Dim i As Long
    Dim mixed As String
    Dim okBlank As Boolean, okSym As Boolean
    Dim ngName As String

    kinds = BlankKinds()
    okBlank = True
    mixed = ""
    For i = LBound(kinds) To UBound(kinds)
        mixed = mixed & CStr(kinds(i))
        If modUtilText.HasVisibleText(CStr(kinds(i))) Then
            okBlank = False
            ngName = ngName & "[" & CStr(i) & "]"
        End If
    Next i

    ' 33 空白類だけなら False。可視文字が1つでもあれば True(落としすぎない)。
    ChkB "Test_W15_33_HasVisibleTextはNBSPやZWSPやBOMも見えない文字に数える_裁定書42", _
        okBlank And _
        (modUtilText.HasVisibleText(mixed) = False) And _
        (modUtilText.HasVisibleText(ChrW$(160) & "田" & ChrW$(8203)) = True) And _
        (modUtilText.HasVisibleText(ChrW$(65279) & "山田") = True), _
        "Trueになった空白類=" & ngName & " 混在=" & _
        CStr(modUtilText.HasVisibleText(mixed))

    ' 34 レポート(ReviewerOf)と提案書(NeedsReviewMessage)が**同じ表で対称**。
    okSym = True
    ngName = ""
    For i = LBound(kinds) To UBound(kinds)
        If modExportHtml.ReviewerOf(CStr(kinds(i))) <> "" Then
            okSym = False
            ngName = ngName & "report[" & CStr(i) & "]"
        End If
        If modExportProposal.NeedsReviewMessage(CStr(kinds(i))) = "" Then
            okSym = False
            ngName = ngName & "proposal[" & CStr(i) & "]"
        End If
    Next i
    ChkB "Test_W15_34_確認者名の空白判定がレポートと提案書で対称_裁定書42", _
        okSym And _
        (modExportHtml.ReviewerOf(mixed) = "") And _
        (modExportProposal.NeedsReviewMessage(mixed) <> "") And _
        (modExportHtml.ReviewerOf(ChrW$(160) & "山田") <> "") And _
        (modExportProposal.NeedsReviewMessage(ChrW$(160) & "山田") = ""), _
        "非対称だった標本=" & ngName

    T_W15Ranges
End Sub

' ---- 35 見えない文字は範囲で決まる(境界の内外) ----------------------------
Private Sub T_W15Ranges()
    Dim edges As Variant
    Dim i As Long
    Dim okIn As Boolean, okOut As Boolean
    Dim ngIn As String, ngOut As String

    ' 範囲の内側の端: 1文字だけなら「入力なし」でなければならない。
    okIn = True
    edges = InvisibleEdges()
    For i = LBound(edges) To UBound(edges)
        If modUtilText.HasVisibleText(ChrW$(CLng(edges(i)))) Then
            okIn = False
            ngIn = ngIn & "U+" & Hex$(CLng(edges(i))) & " "
        End If
    Next i

    ' 範囲の1つ外: 1文字でも「入力あり」でなければならない(落としすぎない)。
    okOut = True
    edges = VisibleEdges()
    For i = LBound(edges) To UBound(edges)
        If Not modUtilText.HasVisibleText(ChrW$(CLng(edges(i)))) Then
            okOut = False
            ngOut = ngOut & "U+" & Hex$(CLng(edges(i))) & " "
        End If
    Next i

    ChkB "Test_W15_35_見えない文字は範囲の境界の内外で決まる_裁定書43Y-2", _
        okIn And okOut And _
        (modUtilText.HasVisibleText(ChrW$(&H200C&) & ChrW$(&H2060&) & _
                                    ChrW$(&HAD&) & ChrW$(&H2002&)) = False) And _
        (modUtilText.HasVisibleText(ChrW$(&HAD&) & "山" & ChrW$(&H2060&)) = True), _
        "内側なのに可視=[" & ngIn & "] 外側なのに不可視=[" & ngOut & "]"
End Sub
