Attribute VB_Name = "modTestsPure25"
Option Explicit

' ============================================================================
' modTestsPure25 - W15(裁定書38 班A)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**裁定書38 §1 班A の文と 15章§2(v2.7)・18章§2/§3 だけ**から
'   手で書き出した(17章§1。実装の出力を見てから期待値を合わせない)。
'
' 対象と根拠(全19本):
'   G1 出典の実在照合(modValidate3。V-S1-14)
'     01 実在するURLは警告にならず、捏造したURLだけが警告になる【V-S1-14】
'     02 URL末尾の句読点・閉じ括弧の揺れを吸収して実在と判定する
'     03 貼付原文が空なら**検査しない**(全件を未照合で埋めない= fail-open)
'     04 空文字のURLは警告にする(「URLの無い資料は sources に入れない」)
'   G2 接頭辞・出所の不整合(modValidate3。V-S1-15)
'     05 certainty=assumed なのに「(見立て)」が無ければ警告【V-S1-15】
'     06 「(見立て)」があれば警告にならない(全角括弧も可)
'     07 financials.source<>unknown で4項目すべて「不明」なら警告
'     08 source=unknown なら4項目すべて「不明」でも警告にならない
'   G3 警告の畳み込み(modValidate3.WarnNoteOf)
'     09 件数を "V-S1-14:2,V-S1-15:1" の形へ  10 指摘なしなら空文字
'   G4 S1再実行の揺れ(modPipeline3.S1DiffCount。B-14)
'     11 同一=0  12 1項目違い=1  13 主要8項目すべて違う=8  14 片方が空=0
'   G6 W15 Round2(裁定書39 §1)で追加した6本
'     20 sources 欠落は**不合格ではなく警告**(V-S1-16)へ降格する【R1-09】
'     21 NormalizeLlmJson(s1) が欠落した sources に空配列を補填する【R1-09】
'     22 missing_info[].kind の enum 検査(V-S1-17)【X-1】
'     23 financials の外にある source を読まない(スキーマ順に依存しない)【R1-10】
'     24 financials の外にある fiscal_year 等を読まない【R1-10】
'     25 対象外フィールドの「(見立て)」を接頭辞とみなさない【G-2】
'
'   G5 スキーマ・mock・描画の結線(B-04 / B-11)
'     15 SchemaS1 が sources と missing_info[].kind を required で持つ
'     16 mock の S1応答2本が CheckS1 を**警告も含めて0件**で通る
'     17 mock の MK-S1-RNW に kind=conflict が1件ある(SEC-03/04 の分離表示の素材)
'     18 SEC-14 が sources と meta.s1_warn を読み、<a href> を作らない
'     19 SEC-04 が conflict を分離し、表に「種別」列を出す
'
' 変異注入(出来レース禁止・裁定書38 §2):
'   (a) modValidate3.UrlNotes の InStr 判定を常に「見つかった」にすると 01 と 04
'       が落ちる。
'   (b) modPipeline3.S1DiffCount を常に 0 にすると 12 と 13 が落ちる。
'
' グループ単位の失敗隔離: modTestsPure24 と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ============================================================================

' 貼付原文の模擬(DR出力のコピペを想定)。02 の捏造URLはこの中に**無い**。
Private Const W15_HAY As String = _
    "会社概要は https://example.co.jp/company/ に掲載されています。" & _
    "決算公告(https://example.co.jp/ir/koukoku2025.html)によれば純資産は12億円でした。"

Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 6
        grpName = "W15-G" & CStr(i)
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
    Case 1: T_Sources
    Case 2: T_Prefix
    Case 3: T_WarnNote
    Case 4: T_Diff
    Case 5: T_Wiring
    Case 6: T_Round2Fixes
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

Private Sub ChkN(ByVal nm As String, ByVal act As Long, ByVal want As Long)
    modTestRunner.Check nm, (act = want), "期待=" & CStr(want) & " 実際=" & CStr(act)
End Sub

Private Function Ctn(ByVal hay As String, ByVal needle As String) As Boolean
    Ctn = (InStr(hay, needle) > 0)
End Function

' sources だけを持つS1の断片(V-S1-14 の検査に必要な最小形)。
Private Function SrcJson(ByVal urlText As String) As String
    SrcJson = "{""sources"":[{""label"":""会社概要"",""url"":""" & urlText & _
              """,""aspect"":""profile""}]}"
End Function

' ============================================================================
' G1 出典の実在照合(V-S1-14)
' ============================================================================
Private Sub T_Sources()
    Dim r As String

    ' 01 実在するURLは警告にならず、捏造したURLだけが1件の警告になる。
    '    **InStr 判定を常に「見つかった」にするとここが落ちる**(変異注入(a))。
    r = modValidate3.CheckS1Notes(SrcJson("https://example.co.jp/company/"), W15_HAY)
    ChkS "Test_V-S1-14_出典URLの貼付原文実在(実在は警告なし・捏造は警告)_裁定書38B-04", _
        r & "|" & modValidate3.WarnNoteOf( _
            modValidate3.CheckS1Notes(SrcJson("https://example.co.jp/fake/"), W15_HAY)), _
        "|V-S1-14:1"

    ' 02 末尾の句読点・閉じ括弧はURLの一部ではない(落としてから照合する)。
    r = modValidate3.CheckS1Notes( _
        SrcJson("https://example.co.jp/ir/koukoku2025.html)。"), W15_HAY)
    ChkS "Test_W15_02_URL末尾の句読点と閉じ括弧の揺れを吸収する_裁定書38B-04", r, ""

    ' 03 貼付原文が空なら検査しない(貼付が空のときに全件を未照合で埋めない)。
    r = modValidate3.CheckS1Notes(SrcJson("https://example.co.jp/fake/"), "")
    ChkS "Test_W15_03_貼付原文が空なら出典照合をしない_裁定書38B-04", r, ""

    ' 04 空のURLは「原文に実在するURL」ではない(15章§2 ルール11)。
    ChkB "Test_W15_04_空のURLは警告にする_裁定書38B-04", _
        Ctn(modValidate3.CheckS1Notes(SrcJson(""), W15_HAY), "[V-S1-14] "), _
        "空URLを InStr の戻り1(常に一致)で見逃さない"
End Sub

' ============================================================================
' G2 接頭辞・出所の不整合(V-S1-15)
' ============================================================================
Private Sub T_Prefix()
    Dim bare As String, marked As String, fin As String

    ' 検体は**正規化の後**のS1応答を模す(裁定書40 P-M1 で V-S1-16 も同じ注記
    '   チャネルへ移ったので、sources を欠いた検体だと V-S1-15 の判定に
    '   V-S1-16 の1行が混ざる。実運用では PostNormalize が空配列を補う)。
    bare = "{""sources"":[],""current_coverage"":[{""line_name"":""労働災害総合保険""," & _
           """coverage_summary"":""元請の包括契約に上乗せが乗っている""," & _
           """limit_note"":""不明"",""special_note"":""不明"",""certainty"":""assumed""}]}"
    marked = Replace(bare, """労働災害総合保険""", """労働災害総合保険(見立て)""")

    ' 05 assumed なのに接頭辞が1つも無い要素は警告にする。
    ChkS "Test_V-S1-15_接頭辞と出所の不整合(assumedに見立てが無い)_裁定書38B-12", _
        modValidate3.WarnNoteOf(modValidate3.CheckS1Notes(bare, W15_HAY)), "V-S1-15:1"

    ' 06 接頭辞があれば警告にしない。
    ChkS "Test_W15_06_見立ての接頭辞があれば警告にしない_裁定書38B-12", _
        modValidate3.CheckS1Notes(marked, W15_HAY), ""

    ' 07 出所を名乗っているのに4項目すべて「不明」は食い違い。
    fin = "{""sources"":[],""financials"":{""fiscal_year"":""不明"",""net_assets"":""不明""," & _
          """sales"":""不明"",""operating_profit"":""不明"",""source"":""yuho""," & _
          """note"":""不明""}}"
    ChkB "Test_W15_07_出所を名乗るのに全項目不明なら警告_裁定書38B-12", _
        Ctn(modValidate3.CheckS1Notes(fin, W15_HAY), "[V-S1-15] "), _
        "source=yuho なのに fiscal_year/net_assets/sales/operating_profit が全て不明"

    ' 08 source=unknown なら全項目「不明」が正しい姿(警告にしない)。
    ChkS "Test_W15_08_出所がunknownなら全項目不明でも警告にしない_裁定書38B-12", _
        modValidate3.CheckS1Notes(Replace(fin, """yuho""", """unknown"""), W15_HAY), ""
End Sub

' ============================================================================
' G3 警告の畳み込み(run_log detail / meta.s1_warn の値)
' ============================================================================
Private Sub T_WarnNote()
    Dim notesText As String

    notesText = "[V-S1-14] sources[0].url が貼付原文に見当たりません: a" & vbLf & _
                "[V-S1-14] sources[1].url が貼付原文に見当たりません: b" & vbLf & _
                "[V-S1-15] 接頭辞・出所の不整合があります: c"
    ChkS "Test_W15_09_警告をケースIDごとの件数へ畳む_裁定書38班A", _
        modValidate3.WarnNoteOf(notesText), "V-S1-14:2,V-S1-15:1"

    ChkS "Test_W15_10_指摘が無ければ注記を出さない_裁定書38班A", _
        modValidate3.WarnNoteOf(""), ""
End Sub

' ============================================================================
' G4 S1再実行の揺れ(B-14)
' ============================================================================
Private Sub T_Diff()
    Dim baseJson As String, oneJson As String, allDiff As String

    baseJson = S1Fixture("株式会社甲斐商店", "和菓子の製造販売", "地域に根差す", _
                     "国内市場は横ばい", 1, "85億円", "confirmed", "not_found", "mid")
    oneJson = S1Fixture("株式会社甲斐商店", "和菓子の製造販売", "地域に根差す", _
                    "国内市場は横ばい", 1, "82億円", "confirmed", "not_found", "mid")
    allDiff = S1Fixture("甲斐商店株式会社", "洋菓子の製造卸売", "全国へ広げる", _
                        "国内市場は縮小", 2, "82億円", "assumed", "conflict", "low")

    ' 11 同じ出力を2回並べたら差分は0。
    ChkN "Test_W15_11_同一のS1出力なら差分は0件_裁定書38B-14", _
        modPipeline3.S1DiffCount(baseJson, baseJson), 0

    ' 12 売上高だけが違えば1。**常に0を返す実装ではここが落ちる**(変異注入(b))。
    ChkN "Test_W15_12_1項目だけ違えば差分は1件_裁定書38B-14", _
        modPipeline3.S1DiffCount(baseJson, oneJson), 1

    ' 13 主要8フィールドすべてが違えば8。
    ChkN "Test_W15_13_主要8フィールドすべて違えば差分は8件_裁定書38B-14", _
        modPipeline3.S1DiffCount(baseJson, allDiff), 8

    ' 14 片方が空(前回が無い)なら比較しない=0。
    ChkN "Test_W15_14_片方が空なら比較せず0件_裁定書38B-14", _
        modPipeline3.S1DiffCount("", baseJson), 0
End Sub

' S1DiffCount が見る主要8フィールドだけを持つ最小のS1(他のキーは持たない)。
Private Function S1Fixture(ByVal companyName As String, ByVal summary As String, _
                           ByVal mvv As String, ByVal market As String, _
                           ByVal locN As Long, ByVal sales As String, _
                           ByVal certainty As String, ByVal kind As String, _
                           ByVal overall As String) As String
    Dim s As String, i As Long

    s = "{""company_name"":""" & companyName & """,""business_summary"":""" & summary & """"
    s = s & ",""strategy_outlook"":{""mvv"":""" & mvv & """,""aspirations"":[""EC拡大""]"
    s = s & ",""market_context"":""" & market & """}"
    s = s & ",""locations"":["
    For i = 1 To locN
        If i > 1 Then s = s & ","
        s = s & "{""name"":""拠点" & CStr(i) & """}"
    Next i
    s = s & "]"
    s = s & ",""financials"":{""sales"":""" & sales & """}"
    s = s & ",""current_coverage"":[{""certainty"":""" & certainty & """}"
    If certainty = "assumed" Then s = s & ",{""certainty"":""assumed""}"
    s = s & "]"
    s = s & ",""missing_info"":[{""item"":""A"",""kind"":""" & kind & """}"
    If kind = "conflict" Then s = s & ",{""item"":""B"",""kind"":""conflict""}"
    s = s & "]"
    s = s & ",""input_quality"":{""overall"":""" & overall & """}}"
    S1Fixture = s
End Function

' ============================================================================
' G5 スキーマ・mock・描画の結線
' ============================================================================
Private Sub T_Wiring()
    Dim sch As String, s1new As String, s1rnw As String, js As String

    ' 15 スキーマ(15章§2 v2.7 と同期。prompt_diff が本文一致を見るので、ここは
    '    「required に載っているか」だけを見る)。
    sch = modSchemas.SchemaS1()
    ChkB "Test_W15_15_SchemaS1がsourcesとkindをrequiredで持つ_裁定書38B-04B-11", _
        Ctn(sch, """research_requests"", ""sources""") And _
        Ctn(sch, """item"", ""why_needed"", ""kind""") And _
        Ctn(sch, """hearing_only"""), _
        "ルートの required に sources / missing_info の required に kind"

    ' 16 mock の S1応答は自分の文脈で**警告も含めて**0件(15章§8.1 受入条件1)。
    s1new = modMockLlm.ResponseById("MK-S1-NEW")
    s1rnw = modMockLlm.ResponseById("MK-S1-RNW")
    ChkS "Test_W15_16_mockのS1応答2本がCheckS1を0件で通る_15章8.1", _
        modValidate.CheckS1(s1new, "new") & modValidate.CheckS1(s1rnw, "renewal"), ""

    ' 17 SEC-03/04 の分離表示の素材(kind=conflict)が mock にある。
    ChkB "Test_W15_17_mockのMK-S1-RNWにkind=conflictがある_裁定書38B-11", _
        Ctn(s1rnw, """kind"":""conflict"""), _
        "SEC-03/04 の「資料間で値が食い違っています」を mock で描けること"

    ' 18 SEC-14 が sources と meta.s1_warn を読み、リンクを作らない(18章§4.1)。
    js = modHtmlTemplate5.SecSourceJs()
    ChkB "Test_W15_18_SEC-14が出典とs1_warnを読みリンクを作らない_裁定書38B-04", _
        Ctn(js, "(D.s1||{}).sources") And Ctn(js, "(D.meta||{}).s1_warn") And _
        Ctn(js, "'出典','URL','観点'") And (InStr(js, "href") = 0), _
        "出典表はテキストのみ(<a href> を作らない)"

    ' 19 SEC-04 は conflict を分離し、残りの表に「種別」列を出す。
    js = modHtmlTemplate2.SecSufficiencyJs() & modHtmlTemplate2.SecProfileJs()
    ChkB "Test_W15_19_SEC-03と04がconflictを最上段へ分離する_裁定書38B-11", _
        Ctn(js, "資料間で値が食い違っています") And Ctn(js, "CONFBOX(el,s1);") And _
        Ctn(js, "LB(LMK,mi[k].kind)") And Ctn(js, "'不足している情報','なぜ必要か','種別'"), _
        "最上段の分離表示と、表の種別列"
End Sub

' ============================================================================
' G6 W15 Round2 の是正(裁定書39 §1 R1-09 / R1-10 / G-2 / X-1)
' ----------------------------------------------------------------------------
' 期待値の出典: 裁定書39 §1 の該当行と 15章§11(V-S1-16 / V-S1-17)だけ。
'   実リボンのモデルは 15章§2 ルール11 を落とすことがあり、sources を required の
'   ままにすると S1 が「修復リトライ1回 -> 失敗」で毎回落ちる(mock は必ず返すので
'   ゲートでは絶対に露見しない)。fail-open へ倒す。
' ============================================================================
Private Sub T_Round2Fixes()
    Dim j As String, r As String, norm As String
    Dim removed As Long

    ' 20 sources 欠落は V-S1-16(警告)。**警告は CheckS1 の戻り値に載せない**
    '    (戻り値は modPipeline.Defend の errText = 修復リトライ -> FailStep の
    '    fail-closed 経路。裁定書40 P-M1)。出口は注記チャネル(CheckS1Notes)
    '    1本であり、sources があるときは注記も出ない(両方向で固定する)。
    j = "{""missing_info"":[{""item"":""A"",""why_needed"":""B"",""kind"":""not_found""}]}"
    r = modValidate.CheckS1(j, "new")
    ChkB "Test_V-S1-16_sources欠落は警告チャネルだけに出す_裁定書40P-M1", _
        Not Ctn(r, "[V-S1-16] ") And Not Ctn(r, "必須キー sources がありません") And _
        Ctn(modValidate3.CheckS1Notes(j, W15_HAY), "[V-S1-16] ") And _
        Not Ctn(modValidate3.CheckS1Notes( _
            Replace(j, "{""missing_info""", "{""sources"":[],""missing_info"""), _
            W15_HAY), "[V-S1-16] "), _
        "実際=[" & r & "] 注記=[" & modValidate3.CheckS1Notes(j, W15_HAY) & "]"

    ' 21 補填: NormalizeLlmJson("s1") を通すと空配列の sources が入り、
    '    以降の経路(HTMLレポート SEC-14 の出典表)が「キーが無い」で割れない。
    '    補填後は注記チャネルの V-S1-16 も鳴らない(補填が効かない壊れたJSONの
    '    ときだけ鳴る安全網であること)。
    norm = modValidate.NormalizeLlmJson("s1", j, removed)
    ChkB "Test_W15R2_21_S1の正規化が欠落したsourcesへ空配列を補填する_裁定書39R1-09", _
        (InStr(norm, """sources""") > 0) And _
        (modJsonLite.GetArrayItems(norm, "sources").Count = 0) And _
        Not Ctn(modValidate3.CheckS1Notes(norm, W15_HAY), "[V-S1-16] "), _
        "実際=[" & norm & "]"

    ' 22 missing_info[].kind の enum 検査(警告)。リボン経路はスキーマ強制が
    '    無いので、conflicted のような値が素通りしていた。これも**注記チャネル
    '    だけ**に出す(CheckS1 の戻り値に載せると kind を落とす/間違えるモデルで
    '    案件が二度と S1 を通せない。裁定書40 P-M1)。
    j = "{""missing_info"":[{""item"":""A"",""why_needed"":""B"",""kind"":""conflicted""}]," & _
        """sources"":[]}"
    r = modValidate3.CheckS1Notes(j, W15_HAY)
    ChkB "Test_V-S1-17_kindのenum外は警告チャネルだけに出す_裁定書40P-M1", _
        Ctn(r, "[V-S1-17] missing_info[0].kind が不正です: conflicted") And _
        Not Ctn(modValidate.CheckS1(j, "new"), "[V-S1-17] ") And _
        Not Ctn(modValidate3.CheckS1Notes(Replace(j, "conflicted", "conflict"), W15_HAY), _
                "[V-S1-17] "), _
        "実際=[" & r & "]"

    ' 23 financials の外にある source を読まない(W14 で潰した「スキーマ順に
    '    依存した読み」の再導入。sources[] の要素が financials より前に来ると
    '    V-S1-15 が別の値を見ていた)。
    j = "{""sources"":[{""label"":""会社概要"",""url"":""https://example.co.jp/company/""," & _
        """aspect"":""profile"",""source"":""yuho""}]," & _
        """financials"":{""fiscal_year"":""不明"",""net_assets"":""不明"",""sales"":""不明""," & _
        """operating_profit"":""不明"",""source"":""unknown"",""note"":""不明""}}"
    r = modValidate3.CheckS1Notes(j, W15_HAY)
    ChkS "Test_W15R2_23_financialsの外のsourceを読まない_裁定書39R1-10", r, ""

    ' 24 AllUnknown も financials の中だけを見る(外側に fiscal_year があっても
    '    financials の4項目が全て「不明」なら食い違いとして警告する)。
    j = "{""x"":{""fiscal_year"":""2025年3月期""}," & _
        """financials"":{""fiscal_year"":""不明"",""net_assets"":""不明"",""sales"":""不明""," & _
        """operating_profit"":""不明"",""source"":""yuho"",""note"":""不明""}}"
    ChkB "Test_W15R2_24_financialsの外のfiscal_yearを読まない_裁定書39R1-10", _
        Ctn(modValidate3.CheckS1Notes(j, W15_HAY), "[V-S1-15] "), _
        "financials の4項目が全て不明なのに警告が出ていない"

    ' 25 対象外フィールドの「(見立て)」を接頭辞とみなさない(生JSON片への
    '    InStr をやめ、値を切り出してから判定する)。
    j = "{""current_coverage"":[{""line_name"":""火災保険"",""coverage_summary"":""建物と設備""," & _
        """limit_note"":""不明"",""special_note"":""不明"",""certainty"":""assumed""," & _
        """memo"":""(見立て)による補足""}]}"
    ChkB "Test_W15R2_25_対象外フィールドの見立ては接頭辞とみなさない_裁定書39G-2", _
        Ctn(modValidate3.CheckS1Notes(j, W15_HAY), "[V-S1-15] "), _
        "スキーマ外のフィールドにある「(見立て)」を拾って警告を握りつぶしている"

    ' 26 kind の**欠落**(と空・空白だけ)は不正ではない。v2.7 の新設キーを
    '    モデルが落とすのは常態であり、not_found とみなして黙って続ける
    '    (FillEmptySources と同じ fail-open。裁定書40 P-M1)。**値が enum 外の
    '    ときだけ**警告する、の両方向を1本で押さえる。
    j = "{""missing_info"":[{""item"":""A"",""why_needed"":""B""}],""sources"":[]}"
    ChkB "Test_W15R3_26_kindの欠落と空は警告にしない_裁定書40P-M1", _
        Not Ctn(modValidate3.CheckS1Notes(j, W15_HAY), "[V-S1-17] ") And _
        Not Ctn(modValidate3.CheckS1Notes( _
            Replace(j, """why_needed"":""B""", """why_needed"":""B"",""kind"":"" """), _
            W15_HAY), "[V-S1-17] ") And _
        Ctn(modValidate3.CheckS1Notes( _
            Replace(j, """why_needed"":""B""", """why_needed"":""B"",""kind"":""conflicted"""), _
            W15_HAY), "[V-S1-17] "), _
        "kind の欠落/空を不正扱いしている(または enum 外を見逃している)"

    ' 27 警告の出口が注記チャネル1本であること: V-S1-16/17 も WarnNoteOf が
    '    畳み、run_log の detail と meta.s1_warn(18章§2)へ届く。並びは
    '    modValidate3 の V3_CASE_IDS の順(実行ごとに揺れない)。
    j = "{""missing_info"":[{""item"":""A"",""why_needed"":""B"",""kind"":""conflicted""}]}"
    ChkS "Test_W15R3_27_s1_warnはsources欠落とkind不正も畳む_裁定書40P-M1", _
        modValidate3.WarnNoteOf(modValidate3.CheckS1Notes(j, W15_HAY)), _
        "V-S1-16:1,V-S1-17:1"
End Sub
