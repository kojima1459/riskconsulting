Attribute VB_Name = "modTestsPure10"
Option Explicit

' ============================
' modTestsPure10 - W3 HTMLレポート(T-33/T-35)の純部の契約テスト(17章§4-1 層(a))
' ----------------------------
' テスト本数: 62本 = G84 6 / G85 8 / G86 5 / G87 8 / G88 8 / G89 5 / G90 6 /
'                    G91 5 / G92 4 / G93 4 / G94 3
' 役割: modTestsPure8/9 と同じく**実装を1行も読まず**、18章全文・14章§6の宣言・
'   16章E-46/E-47/NFR-S7・15章スキーマ・19章§3・11章だけを根拠に入出力を固定する。
'   期待値が一意に決まらない項目はテストにせず末尾へ列挙する。
'     G84 HtmlSafe(16章E-47(2))      G85 JsStringSafe(§5.3の適用順4段)
'     G86 危険JSONの実弾             G87 テーマ(§5.1の28変数・§5.2の差替単位)
'     G88 BuildDocumentの全文組立    G89 HeadHtml/BodyShellHtml(§4.4・§5.1・§6)
'     G90 セクション登録表(§4.2)     G91 SEC-09/SEC-16(§3のv1.0裁定)
'     G92 描画規約(§4.1)             G93 固定文(§3.5・§3.4)
'     G94 enum日本語ラベル(19章§3)
' **署名の仮定・要裁定**: 18章§4.4/§5.2は関数名だけを固定し引数を規定していない
'   (14章§6も「本章は宣言を持たない」)。§1.1⑤に逐語で現れる
'   BuildDocument(themeName, dataJson, coverFields) だけを確定として扱い、
'   HeadHtml / BodyShellHtml / RuntimeJs / SectionsJs は末尾(a)(b)の仮定による。
' 結線(統合済み): modTestsPure9.RunAll の末尾から本 RunAll を呼ぶ(数珠つなぎ
'   modTestsPure -> 2 .. -> 9 -> 10。単体では0本)。modules.json(role=test /
'   wave=T-35)・12章§2・PURE_ALLOWLIST へ登録済み。tests_expected 505 -> 567。
' 設計判断(R4): Excelトークン・乱数・時刻を使わず改行は vbLf。素材は架空値で、
'   JSONは単引用符で書き AsJson() で二重引用符へ直す。
' ============================

' ---- 素材の定数 ----
' 18章§5.1 の閉じた一覧28変数(宣言の有無は「名前+:」で照合する。--line は
' --line-height の接頭辞なので、コロンまで含めないと取り違える)。
Private Const THEME_VARS As String = _
    "--page-width:;--page-pad:;--font-sans:;--font-serif:;--font-size:;" & _
    "--line-height:;--paper:;--ink:;--sub:;--mist:;--line:;--ai:;--kaki:;" & _
    "--matsu:;--deep:;--warn:;--warn-line:;--heat-1:;--heat-2:;--heat-3:;" & _
    "--heat-4:;--heat-5:;--tr-cover:;--tr-partial:;--tr-hard:;--iq-ok:;" & _
    "--iq-partial:;--iq-missing:"

' 18章§3の表の並び(SEC-01..09 -> SEC-16 -> SEC-10..15)。slug は section id と
' 目次アンカーになる。
Private Const SEC_SLUGS As String = _
    "cover;exec;profile;sufficiency;riskuniv;riskmap;risks;coverage;newrisk;" & _
    "round-update;story;prevent;limit;hearing;source;disclaimer"

' 19章§3 リスクユニバース10分類の日本語ラベル(15章§0の変換表の記載順)。
Private Const CAT_LABELS As String = _
    "戦略・市場;調達・供給網;製造・品質;販売・顧客;施設・自然災害・BCP;" & _
    "人材・労務;デジタル・情報;法務・規制;財務・取引先;ブランド・社会"

' 18章§4.1・17章 T-46 が出荷前検問で grep するマークアップ注入の4語。
Private Const FORBIDDEN_JS As String = _
    "innerHTML;insertAdjacentHTML;document.write;outerHTML"

' 18章§3.5 の免責4行のうち逐語で与えられている3行。
Private Const DISC1 As String = _
    "本資料はAI支援により作成した骨子を人が確認・編集したものです。"
Private Const DISC2 As String = _
    "記載のリスクは公開情報と当社担当者の見立てに基づく仮説であり、" & _
    "引受可否・保険料・幹事構成を確約するものではありません。"
Private Const DISC3 As String = "保険料の試算は本資料の対象外です（要見積）。"

' 18章§3 SEC-09 の0件時の1行(v1.0で文言が確定。非表示にしない)。
Private Const NOTE_SEC09 As String = _
    "現時点で特筆すべきニューリスクは検出されていません"

' 表紙(§4.1(b))へ渡す会社名。DATA側の meta.company とは別の値にする(DATAは生の
' ままJS文字列へ入るのが正で、同値だと表紙のエスケープ有無を判別できない)。
Private Const COVER_CO As String = "甲斐<&>商店"
Private Const COVER_CO_ESC As String = "甲斐&lt;&amp;&gt;商店"

' BuildDocument の結果は全群で使い回す(純関数なので1回で足りる)。
Private mDoc As String
Private mDocReady As Boolean

' RunAll: グループ隔離実行(未実装/未注入は GroupFail で可視化)。末端。
Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 11
        grpName = "G?" & i
        On Error Resume Next
        Err.Clear
        RunGroup i, grpName
        If Err.Number <> 0 Then
            GroupFail grpName
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Sub

Private Sub RunGroup(ByVal grpNo As Long, ByRef grpName As String)
    Select Case grpNo
    Case 1
        grpName = "G84 HtmlSafe"
        T_HtmlSafe
    Case 2
        grpName = "G85 JsStringSafe"
        T_JsSafe
    Case 3
        grpName = "G86 危険JSONの実弾"
        T_Payload
    Case 4
        grpName = "G87 テーマCSS"
        T_Theme
    Case 5
        grpName = "G88 文書組立"
        T_Document
    Case 6
        grpName = "G89 Head/BodyShell"
        T_Shell
    Case 7
        grpName = "G90 セクション登録表"
        T_Sections
    Case 8
        grpName = "G91 SEC-09/SEC-16"
        T_NewAndRound
    Case 9
        grpName = "G92 描画規約"
        T_Runtime
    Case 10
        grpName = "G93 固定文"
        T_FixedText
    Case 11
        grpName = "G94 enum日本語ラベル"
        T_Labels
    End Select
End Sub

' ---- 共通ヘルパ(modTestsPure3..9 と同じ作法) ----
Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

Private Function HeadOf(ByVal s As String) As String
    If Len(s) <= 120 Then
        HeadOf = s
    Else
        HeadOf = Left$(s, 120) & "...(全" & Len(s) & "字)"
    End If
End Function

Private Function Ctn(ByVal hay As String, ByVal needle As String) As Boolean
    Ctn = (InStr(hay, needle) > 0)
End Function

Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), _
        "期待=[" & HeadOf(want) & "] 実際=[" & HeadOf(act) & "]"
End Sub

Private Sub ChkB(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check nm, cond, detail
End Sub

' 全要素が hay にあること。
Private Sub ChkAll(ByVal nm As String, ByVal hay As String, ByVal listText As String)
    Dim miss As String
    miss = MissingOf(hay, listText)
    modTestRunner.Check nm, (LenB(miss) = 0), "不足=[" & miss & "]"
End Sub

' 要素が hay に1つも無いこと。
Private Sub ChkNone(ByVal nm As String, ByVal hay As String, ByVal listText As String)
    Dim hit As String
    hit = FoundOf(hay, listText)
    modTestRunner.Check nm, (LenB(hit) = 0), "検出=[" & hit & "]"
End Sub

Private Function DQ() As String
    DQ = Chr$(34)
End Function

Private Function AsJson(ByVal s As String) As String
    AsJson = Replace(s, "'", DQ())
End Function

' §4.2の桁揃えに依存せず照合するため空白・改行・タブを畳む。
Private Function Squash(ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    Squash = Replace(t, " ", "")
End Function

Private Function CountOcc(ByVal hay As String, ByVal needle As String) As Long
    Dim p As Long
    Dim n As Long
    If LenB(needle) = 0 Then Exit Function
    p = InStr(hay, needle)
    Do While p > 0
        n = n + 1
        p = InStr(p + Len(needle), hay, needle)
    Loop
    CountOcc = n
End Function

' hay に無い要素を返す。全部あれば ""。
Private Function MissingOf(ByVal hay As String, ByVal listText As String) As String
    Dim parts() As String
    Dim i As Long
    Dim acc As String
    parts = Split(listText, ";")
    For i = 0 To UBound(parts)
        If LenB(parts(i)) > 0 Then
            If InStr(hay, parts(i)) = 0 Then acc = acc & parts(i) & " "
        End If
    Next i
    MissingOf = Trim$(acc)
End Function

' hay にある要素を返す(禁止語の報告用)。0件なら ""。
Private Function FoundOf(ByVal hay As String, ByVal listText As String) As String
    Dim parts() As String
    Dim i As Long
    Dim acc As String
    parts = Split(listText, ";")
    For i = 0 To UBound(parts)
        If LenB(parts(i)) > 0 Then
            If InStr(hay, parts(i)) > 0 Then acc = acc & parts(i) & " "
        End If
    Next i
    FoundOf = Trim$(acc)
End Function

' 各要素がこの順(先頭出現位置が単調増加)で現れるか。
Private Function InOrder(ByVal hay As String, ByVal listText As String) As Boolean
    Dim parts() As String
    Dim i As Long
    Dim p As Long
    Dim q As Long
    parts = Split(listText, ";")
    For i = 0 To UBound(parts)
        If LenB(parts(i)) > 0 Then
            q = InStr(hay, parts(i))
            If q = 0 Or q < p Then Exit Function
            p = q
        End If
    Next i
    InOrder = True
End Function

Private Function Wrap(ByVal listText As String, ByVal pre As String, _
                      ByVal post As String) As String
    Dim parts() As String
    Dim i As Long
    Dim acc As String
    parts = Split(listText, ";")
    For i = 0 To UBound(parts)
        If LenB(parts(i)) > 0 Then acc = acc & pre & parts(i) & post & ";"
    Next i
    Wrap = acc
End Function

Private Function SecIds() As String
    Dim i As Long
    Dim acc As String
    For i = 1 To 16
        If i < 10 Then
            acc = acc & "SEC-0" & i & ";"
        Else
            acc = acc & "SEC-" & i & ";"
        End If
    Next i
    SecIds = acc
End Function

' §5.1の機械検査(3)用。生の16進色(#fff以外)を1件返す。
Private Function RawHexColor(ByVal s As String) As String
    Dim i As Long
    Dim n As Long
    Dim tok As String
    n = Len(s)
    For i = 1 To n - 3
        If Mid$(s, i, 1) = "#" Then
            tok = ""
            If i + 6 <= n Then
                If IsHexRun(Mid$(s, i + 1, 6)) Then tok = Mid$(s, i + 1, 6)
            End If
            If LenB(tok) = 0 Then
                If IsHexRun(Mid$(s, i + 1, 3)) Then tok = Mid$(s, i + 1, 3)
            End If
            If LenB(tok) > 0 Then
                If UCase$(tok) <> "FFF" And UCase$(tok) <> "FFFFFF" Then
                    RawHexColor = tok
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function IsHexRun(ByVal s As String) As Boolean
    Dim i As Long
    Dim c As String
    If LenB(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = UCase$(Mid$(s, i, 1))
        If InStr("0123456789ABCDEF", c) = 0 Then Exit Function
    Next i
    IsHexRun = True
End Function

' :root{...}を1ブロック落とす(共通CSSだけを16進検査に掛けるため)。
Private Function StripRoot(ByVal s As String) As String
    Dim a As Long
    Dim b As Long
    a = InStr(s, ":root{")
    If a = 0 Then
        StripRoot = s
        Exit Function
    End If
    b = InStr(a, s, "}")
    If b = 0 Then
        StripRoot = Left$(s, a - 1)
    Else
        StripRoot = Left$(s, a - 1) & Mid$(s, b + 1)
    End If
End Function

' "\u00XX"(16進の大小はどちらでも可)を判定する。
Private Function IsUEsc(ByVal act As String, ByVal hex2 As String) As Boolean
    If Len(act) <> 6 Then Exit Function
    If Left$(act, 4) <> "\u00" Then Exit Function
    IsUEsc = (UCase$(Mid$(act, 5)) = UCase$(hex2))
End Function

' 素材: 18章§2 の DATA(キー名は15章スキーマのまま)。危険トークンは risk_name=
'   閉じscript+img / scenario=引用符つきonerror / pitch=引用符と改行。構造側の
'   生の改行で§5.3(1)④も通す。
Private Function DataJsonAttack() As String
    Dim s As String
    s = "{'meta':{'case_id':'C-20260901-001'," & vbLf
    s = s & "'company':'浜松スイーツファクトリー株式会社','industry_code':'09'," & vbLf
    s = s & "'industry_name':'食料品製造業','case_type':'renewal'," & vbLf
    s = s & "'dossier_tier':'t2_full','quality_mode':'deep','round_no':2," & vbLf
    s = s & "'s4_variant':'proposal','generated_at':'2026/09/01 14:07:22'," & vbLf
    s = s & "'app_version':'2.4.0','theme':'standard'}," & vbLf
    s = s & "'s1':null," & vbLf
    s = s & "'s2':{'risks':[{'risk_no':1,'category':'digital_info'," & vbLf
    s = s & "'risk_name':'</script><img src=x onerror=alert(1)>'," & vbLf
    s = s & "'scenario':'<img src=x onerror=\'alert(1)\'>'," & vbLf
    s = s & "'status':'new','frequency':'mid','impact':'large'," & vbLf
    s = s & "'frequency_score':3,'impact_score':5," & vbLf
    s = s & "'evidence':{'quote':'A&B','source':'hp'}," & vbLf
    s = s & "'insurability':{'transferability':'hard','line_note':''," & vbLf
    s = s & "'control_note':''},'loss_scale_note':''," & vbLf
    s = s & "'check_points':[],'preventions':[]}]," & vbLf
    s = s & "'gaps':[],'emerging_risks':[],'open_questions':[]}," & vbLf
    s = s & "'s3':{'stories':[{'story_no':1,'proposal_kind':'upsell'," & vbLf
    s = s & "'headline':'物流停止に備える','hook_question':'在庫は何日分ですか'," & vbLf
    s = s & "'target_risk_nos':[1],'target_gap_nos':[],'menu_ids':['M-0012']," & vbLf
    s = s & "'line_ids':['L-03'],'scheme_id':''," & vbLf
    s = s & "'pitch':'担当者は\'やる\'と言った\n次の行へ'," & vbLf
    s = s & "'similar_case_id':'K-0003','expected_objection':''," & vbLf
    s = s & "'objection_response':''}],'unmatched_risks':[]," & vbLf
    s = s & "'do_not_propose':[]}}" & vbLf
    DataJsonAttack = AsJson(s)
End Function

' coverFields(§4.1(b)の3値)。vbTab区切りで [0]=会社名 [1]=案件ID [2]=生成日時
'   (末尾(b)。期待値 COVER_CO / COVER_CO_ESC は動かしていない)。
Private Function CoverFieldsAttack() As String
    CoverFieldsAttack = COVER_CO & vbTab & "C-20260901-001" & vbTab & _
                        "2026/09/01 14:07:22"
End Function

Private Function DocText() As String
    If Not mDocReady Then
        mDoc = modHtmlTemplate1.BuildDocument("standard", DataJsonAttack(), _
                                              CoverFieldsAttack())
        mDocReady = True
    End If
    DocText = mDoc
End Function

' ---- G84 HtmlSafe(16章E-47(2)・NFR-S7③・§5.3(2)) ----
'   5字をエンティティ化。& を後回しにすると &lt; の & が二重化される。
Private Sub T_HtmlSafe()
    Dim ap As String
    ap = modUtilText.HtmlSafe("'")

    ChkS "G84_アンパサンドをエンティティ化する_16章E-47", _
        modUtilText.HtmlSafe("&"), "&amp;"

    ChkB "G84_不等号をエンティティ化する_16章E-47", _
        ((modUtilText.HtmlSafe("<") = "&lt;") And _
         (modUtilText.HtmlSafe(">") = "&gt;")), _
        "lt=[" & modUtilText.HtmlSafe("<") & "] gt=[" & _
        modUtilText.HtmlSafe(">") & "]"

    ChkS "G84_二重引用符をエンティティ化する_16章E-47", _
        modUtilText.HtmlSafe(DQ()), "&quot;"

    ' 実体参照は &#39; と &apos; のどちらでもよい(末尾(d))。生で残らないことだけ。
    ChkB "G84_単引用符が生のまま残らない_16章E-47", _
        ((InStr(ap, "'") = 0) And (Left$(ap, 1) = "&") And _
         (Right$(ap, 1) = ";")), "実際=[" & ap & "]"

    ' 置換順の固定。& を後回しにすると "&amp;amp;lt;" になる。
    ChkS "G84_アンパサンドを最初に置換する_18章§5.3", _
        modUtilText.HtmlSafe("&lt;"), "&amp;lt;"

    ChkB "G84_5字以外は素通しで空文字は空文字_16章E-47", _
        ((modUtilText.HtmlSafe("工場/1-2") = "工場/1-2") And _
         (LenB(modUtilText.HtmlSafe("")) = 0)), _
        "実際=[" & modUtilText.HtmlSafe("工場/1-2") & "]"
End Sub

' ---- G85 JsStringSafe(§5.3(1)の適用順4段) ----
'   ① \ -> \\ ・ " -> \"  ② </ -> <\/  ③ U+2028/U+2029  ④ 制御文字
Private Sub T_JsSafe()
    ChkS "G85_逆斜線を二重化する_18章§5.3", _
        modUtilText.JsStringSafe("\"), "\\"

    ChkS "G85_二重引用符を逆斜線で逃がす_18章§5.3", _
        modUtilText.JsStringSafe(DQ()), "\" & DQ()

    ' ①の内部順(逆斜線が先)。逆だと \" が \\" になり文字列が閉じる。
    ChkS "G85_JSONのエスケープ済み引用符は逆斜線3本になる_18章§5.3", _
        modUtilText.JsStringSafe("\" & DQ()), "\\\" & DQ()

    ' ②が①の後であること。逆順なら "<\\/script>" になる。
    ChkS "G85_閉じscriptタグを無害化する_18章§5.3", _
        modUtilText.JsStringSafe("</script>"), "<\/script>"

    ChkB "G85_置換対象は閉じ記号の対だけで単独記号は変えない_18章§5.3", _
        ((modUtilText.JsStringSafe("<a/b>") = "<a/b>") And _
         (modUtilText.JsStringSafe("a</b") = "a<\/b")), _
        "実際=[" & modUtilText.JsStringSafe("<a/b>") & "]"

    ChkB "G85_行区切りと段落区切りをエスケープする_18章§5.3", _
        ((modUtilText.JsStringSafe(ChrW(&H2028&)) = "\u2028") And _
         (modUtilText.JsStringSafe(ChrW(&H2029&)) = "\u2029")), _
        "u2028=[" & modUtilText.JsStringSafe(ChrW(&H2028&)) & "]"

    ' ④の逆斜線が①で再度倍化されていないこと(長さ6が証拠)。
    ChkB "G85_制御文字をuXXXX形式へ落とす_18章§5.3", _
        (IsUEsc(modUtilText.JsStringSafe(vbLf), "0A") And _
         IsUEsc(modUtilText.JsStringSafe(vbCr), "0D") And _
         IsUEsc(modUtilText.JsStringSafe(vbTab), "09") And _
         IsUEsc(modUtilText.JsStringSafe(Chr$(0)), "00")), _
        "LF=[" & modUtilText.JsStringSafe(vbLf) & "]"

    ChkB "G85_通常文字は素通しで空文字は空文字_18章§5.3", _
        ((modUtilText.JsStringSafe("工場 1-2") = "工場 1-2") And _
         (LenB(modUtilText.JsStringSafe("")) = 0)), _
        "実際=[" & modUtilText.JsStringSafe("工場 1-2") & "]"
End Sub

' ---- G86 実弾(§5.3(1)・16章E-47(1)) ----
'   生の危険トークンが現れないことと、データが消えていない(無害化であって
'   削除ではない)ことを同時に当てる。
Private Sub T_Payload()
    Dim e As String
    e = modUtilText.JsStringSafe(DataJsonAttack())

    ChkB "G86_出力に生の閉じタグ記号が1つも残らない_16章E-47", _
        ((InStr(e, "</") = 0) And (InStr(e, "</script>") = 0)), _
        "残存位置=" & InStr(e, "</") & " 出力頭=[" & HeadOf(e) & "]"

    ChkB "G86_出力に生の改行が1つも残らない_18章§5.3", _
        ((InStr(e, vbLf) = 0) And (InStr(e, vbCr) = 0)), _
        "LF位置=" & InStr(e, vbLf) & " CR位置=" & InStr(e, vbCr)

    ' 逃がし漏れが1つでもあればJS文字列リテラルが閉じる。
    ChkB "G86_全ての二重引用符が逆斜線を伴う_18章§5.3", _
        (CountOcc(e, DQ()) = CountOcc(e, "\" & DQ())), _
        "引用符=" & CountOcc(e, DQ()) & " 逃がし済=" & CountOcc(e, "\" & DQ())

    ChkB "G86_危険トークンは削除ではなく無害化される_18章§5.3", _
        Ctn(e, "<\/script><img src=x onerror=alert(1)>"), _
        "出力頭=[" & HeadOf(e) & "]"

    ' 属性値の引用符は JSON の \" を経て \\\" になる。
    ChkB "G86_属性値の引用符が逆斜線3本を伴う_18章§5.3", _
        Ctn(e, "onerror=\\\" & DQ() & "alert(1)\\\" & DQ()), _
        "onerror位置=" & InStr(e, "onerror")
End Sub

' ---- G87 テーマ(§5.1の閉じた一覧28変数・§5.2の差替単位) ----
Private Sub T_Theme()
    Dim names As String
    Dim nameArr() As String
    Dim css As String
    Dim cssMono As String
    Dim inner As String
    Dim parts() As String
    Dim i As Long
    Dim okAll As Boolean
    Dim wantVals As String

    names = modHtmlTheme.ThemeNames()
    nameArr = Split(names, ";")
    ChkS "G87_テーマ一覧の先頭が既定テーマ_18章§5.2", nameArr(0), "standard"

    ChkB "G87_初期テーマはstandardとmonoの2本_18章§5.2", _
        (Ctn(names, "mono") And (UBound(nameArr) >= 1)), "実際=[" & names & "]"

    css = Squash(modHtmlTheme.ThemeCss("standard"))
    ChkB "G87_戻り値はroot1ブロックだけでセレクタを持たない_18章§5.1", _
        ((Left$(css, 6) = ":root{") And (Right$(css, 1) = "}") And _
         (CountOcc(css, "{") = 1) And (CountOcc(css, "}") = 1)), _
        "頭=[" & HeadOf(css) & "] 中括弧=" & CountOcc(css, "{")

    ' 内側は "--" で始まる宣言だけ(CSS変数以外を書かない)。
    inner = Replace(Replace(css, ":root{", ""), "}", "")
    parts = Split(inner, ";")
    okAll = True
    For i = 0 To UBound(parts)
        If LenB(parts(i)) > 0 Then
            If Left$(parts(i), 2) <> "--" Then okAll = False
        End If
    Next i
    ChkB "G87_内側はCSS変数の宣言だけ_18章§5.1", okAll, _
        "変数以外の宣言がある: [" & HeadOf(inner) & "]"

    ChkB "G87_standardが28変数を過不足なく宣言する_18章§5.1", _
        ((LenB(MissingOf(css, THEME_VARS)) = 0) And (CountOcc(css, "--") = 28)), _
        "不足=[" & MissingOf(css, THEME_VARS) & "] 数=" & CountOcc(css, "--")

    cssMono = Squash(modHtmlTheme.ThemeCss("mono"))
    ChkB "G87_monoも28変数を過不足なく宣言し配色が異なる_18章§5.2", _
        ((LenB(MissingOf(cssMono, THEME_VARS)) = 0) And _
         (CountOcc(cssMono, "--") = 28) And (cssMono <> css)), _
        "不足=[" & MissingOf(cssMono, THEME_VARS) & "] 数=" & CountOcc(cssMono, "--")

    ChkB "G87_未知のテーマ名はstandardへフォールバックする_18章§5.2", _
        (Squash(modHtmlTheme.ThemeCss("no_such_theme")) = css), _
        "実際=[" & HeadOf(Squash(modHtmlTheme.ThemeCss("no_such_theme"))) & "]"

    wantVals = "--PAPER:#FFFFFF;--INK:#24303E;--HEAT-3:#FFF4D6;" & _
               "--TR-HARD:#B4552D;--PAGE-WIDTH:900PX;--FONT-SIZE:14.5PX"
    ChkAll "G87_standardの既定値が18章§5.1の表どおり_18章§5.1", UCase$(css), wantVals
End Sub

' ---- G88 文書組立(§1.1⑤ BuildDocument。全文に対する契約) ----
Private Sub T_Document()
    Dim doc As String
    Dim sq As String
    Dim rest As String
    Dim slugKeys As String
    Dim p As Long
    doc = DocText()
    sq = Squash(doc)
    slugKeys = Wrap(SEC_SLUGS, "slug:'", "'")

    ' §5.3(3): <head> の**最初の要素**が meta charset であること。
    p = InStr(sq, "<head>")
    If p > 0 Then rest = Mid$(sq, p + 6)
    ChkB "G88_headの最初の要素がmetacharsetである_18章§5.3", _
        ((p > 0) And (Left$(rest, 13) = "<metacharset=") And _
         Ctn(Left$(rest, 40), "utf-8")), _
        "head位置=" & p & " 直後=[" & Left$(rest, 40) & "]"

    ChkB "G88_DATAは1本のJS文字列リテラルとJSONparseで埋める_16章E-47", _
        (Ctn(sq, "varDATA=JSON.parse(" & DQ()) And (InStr(sq, "varDATA={") = 0)), _
        "JSONparse位置=" & InStr(sq, "varDATA=JSON.parse(" & DQ())

    ChkB "G88_全文に生の閉じscript付き注入が現れない_16章E-47", _
        (InStr(doc, "</script><img") = 0), _
        "検出位置=" & InStr(doc, "</script><img")

    ChkB "G88_全文に生の引用符つきonerror属性が現れない_16章E-47", _
        (InStr(doc, "onerror=" & DQ() & "alert(1)") = 0), _
        "検出位置=" & InStr(doc, "onerror=" & DQ() & "alert(1)")

    ChkB "G88_表紙の会社名がHtmlSafeを通る_18章§4.1", _
        ((InStr(doc, COVER_CO) = 0) And Ctn(doc, COVER_CO_ESC)), _
        "生=" & InStr(doc, COVER_CO) & " 済=" & InStr(doc, COVER_CO_ESC)

    ChkB "G88_16のセクションIDとslugと目次アンカーを持つ_18章§3", _
        ((LenB(MissingOf(doc, SecIds())) = 0) And _
         (LenB(MissingOf(sq, slugKeys)) = 0) And Ctn(sq, "#sec-")), _
        "ID不足=[" & MissingOf(doc, SecIds()) & "] slug不足=[" & _
        MissingOf(sq, slugKeys) & "]"

    ChkNone "G88_マークアップ注入の4語がテンプレ全文に現れない_18章§4.1", _
        doc, FORBIDDEN_JS

    ChkNone "G88_外部参照を持たない自己完結HTML_18章§1", _
        LCase$(doc), "http://;https://;@import;<link"
End Sub

' ---- G89 HeadHtml / BodyShellHtml(§4.4・§5.3(3)・§5.1・§6) ----
Private Sub T_Shell()
    Dim h As String
    Dim b As String
    Dim sq As String
    Dim wantCss As String
    Dim bad As String
    h = modHtmlTemplate1.HeadHtml("standard", CoverFieldsAttack())
    sq = Squash(h)
    b = modHtmlTemplate1.BodyShellHtml(CoverFieldsAttack())

    ChkB "G89_HeadHtmlがmetacharsetutf8とテーマCSSを持つ_18章§4.4", _
        (Ctn(sq, "<metacharset=") And Ctn(sq, "utf-8") And Ctn(sq, ":root{")), _
        "meta=" & InStr(sq, "<metacharset=") & " root=" & InStr(sq, ":root{")

    ChkAll "G89_A4縦と余白をリテラルで持つ_18章§6", sq, "@page;A4;14mm;12mm"

    wantCss = "break-inside:avoid;break-after:page;@mediaprint;.no-print;" & _
              "print-color-adjust;max-width:var(--page-width);" & _
              "@media(max-width:640px)"
    ChkAll "G89_印刷と画面の両立規約を共通CSSが持つ_18章§6", sq, wantCss

    ' §5.1: 共通CSSは色をリテラルで書かない(例外は #fff)。テーマの :root は
    ' 色を書くのが仕事なので検査対象から外す。
    bad = RawHexColor(StripRoot(sq))
    ChkB "G89_共通CSSにfff以外の生16進色が無い_18章§5.1", (LenB(bad) = 0), _
        "検出=[#" & bad & "]"

    ChkB "G89_noscriptにJS必須案内とAI利用の明示を書く_18章§4.1", _
        (Ctn(LCase$(b), "<noscript") And _
         Ctn(b, "このレポートの表示にはJavaScriptが必要です") And Ctn(b, DISC1)), _
        "noscript=" & InStr(LCase$(b), "<noscript") & " AI=" & InStr(b, DISC1)
End Sub

' ---- G90 セクション登録表(§4.2。登録行7キー固定・例が逐語の正) ----
Private Sub T_Sections()
    Dim sj As String
    Dim keys As String
    Dim slugKeys As String
    sj = Squash(modHtmlTemplate1.SectionsJs())
    keys = "id:';slug:';title:';need:[;empty:';note:';render:"
    slugKeys = Wrap(SEC_SLUGS, "slug:'", "'")

    ChkB "G90_登録配列の宣言と閉じを持つ_18章§4.2", _
        (Ctn(sj, "varSECTIONS=[") And Ctn(sj, "];")), _
        "宣言=" & InStr(sj, "varSECTIONS=[") & " 閉じ=" & InStr(sj, "];")

    ChkAll "G90_16のセクションIDが登録表にある_18章§3", sj, SecIds()

    ChkAll "G90_16のslugが登録表にある_18章§4.2", sj, slugKeys

    ChkAll "G90_登録行のキーは7つに固定_18章§4.2", sj, keys

    ' §4.2 のコード例が示す登録行を逐語で当てる(桁揃えの空白だけを畳む)。
    ChkB "G90_SEC-01の登録行が§4.2の例どおり_18章§4.2", _
        Ctn(sj, "{id:'SEC-01',slug:'cover',title:'',need:['meta']," & _
                "empty:'always',render:renderCover},"), _
        "登録表頭=[" & HeadOf(sj) & "]"

    ChkB "G90_SEC-02とSEC-06の登録行が§4.2の例どおり_18章§4.2", _
        (Ctn(sj, "{id:'SEC-02',slug:'exec',title:'エグゼクティブサマリ'," & _
                 "need:['s1'],empty:'always',render:renderExec},") And _
         Ctn(sj, "{id:'SEC-06',slug:'riskmap',title:'2軸リスクマップ'," & _
                 "need:['s2'],empty:'hide',render:renderRiskMap},")), _
        "02=" & InStr(sj, "id:'SEC-02'") & " 06=" & InStr(sj, "id:'SEC-06'")
End Sub

' ---- G91 SEC-09 / SEC-16(§3・§4.2・19章§1。v1.0で別物と確定) ----
'   SEC-09は空配列でも1行出す(hideにしない)。SEC-16はround_no<2または3つの
'   statusが0件ならhide。
Private Sub T_NewAndRound()
    Dim sj As String
    Dim doc As String
    sj = Squash(modHtmlTemplate1.SectionsJs())
    doc = DocText()

    ChkB "G91_SEC-09の登録行が空配列時の案内文を持つ_18章§3", _
        Ctn(sj, "{id:'SEC-09',slug:'newrisk',title:'ニューリスク',need:['s2']," & _
                "empty:'note',note:'" & NOTE_SEC09 & "',render:renderNewRisk},"), _
        "SEC-09位置=" & InStr(sj, "id:'SEC-09'") & " 文=" & InStr(sj, NOTE_SEC09)

    ChkB "G91_SEC-16の登録行が0件時は非表示_18章§3", _
        Ctn(sj, "{id:'SEC-16',slug:'round-update',title:'訪問で分かったこと'," & _
                "need:['s2'],empty:'hide',render:renderRoundUpdate},"), _
        "SEC-16位置=" & InStr(sj, "id:'SEC-16'")

    ' §3の表の並び: SEC-09 の次が SEC-16、その次が SEC-10(紙面順)。
    ChkB "G91_紙面順はSEC-09の次がSEC-16でその次がSEC-10_18章§3", _
        InOrder(sj, "id:'SEC-09';id:'SEC-16';id:'SEC-10'"), _
        "09=" & InStr(sj, "id:'SEC-09'") & " 16=" & InStr(sj, "id:'SEC-16'") & _
        " 10=" & InStr(sj, "id:'SEC-10'")

    ' v1.0で削除された文言。残っていれば SEC-09 の意味が旧版のままになる。
    ChkB "G91_第2ラウンド以降に表示の文言が残っていない_18章v1.0", _
        (InStr(doc, "第2ラウンド以降に表示") = 0), _
        "検出位置=" & InStr(doc, "第2ラウンド以降に表示")

    ' SEC-09 は emerging_risks、SEC-16 は round_no と risks[].status を読む。
    ChkAll "G91_SEC-09はemergingをSEC-16はround_noとstatusを読む_18章§3", _
        doc, "emerging_risks;round_no;confirmed;rejected"
End Sub

' ---- G92 描画規約(§4.1。T-46の出荷前検問を層(a)で先に回す) ----
Private Sub T_Runtime()
    Dim rj As String
    rj = modHtmlTemplate1.RuntimeJs()

    ChkNone "G92_RuntimeJsにマークアップ注入の4語が現れない_18章§4.1", _
        rj, FORBIDDEN_JS

    ChkAll "G92_描画はcreateElementとtextContentで行う_18章§4.1", _
        rj, "createElement;textContent"

    ChkB "G92_属性はsetAttributeで与える_18章§4.1", _
        Ctn(rj, "setAttribute"), "位置=" & InStr(rj, "setAttribute")

    ChkB "G92_目次生成とemptyの3値判定を持つ_18章§3.6", _
        (Ctn(Squash(rj), "#sec-") And _
         (LenB(MissingOf(rj, "always;hide;note")) = 0)), _
        "不足=[" & MissingOf(rj, "always;hide;note") & "]"
End Sub

' ---- G93 固定文(§3.5の免責4行・§3.4のヒアリング設問文) ----
Private Sub T_FixedText()
    Dim doc As String
    Dim want4 As String
    doc = DocText()
    want4 = "御中;案件ID;リスク提案ナビ;について教えてください;ヒアリングシートを参照"

    ChkB "G93_免責1行目のAI利用の必須表記_16章NFR-S5", _
        Ctn(doc, DISC1), "位置=" & InStr(doc, DISC1)

    ChkB "G93_免責2行目の仮説である旨_18章§3.5", _
        Ctn(doc, DISC2), "位置=" & InStr(doc, DISC2)

    ChkB "G93_免責3行目の保険料試算は対象外_10章FR-43", _
        Ctn(doc, DISC3), "位置=" & InStr(doc, DISC3)

    ChkAll "G93_免責4行目の署名要素とヒアリングの設問文_18章§3.4", doc, want4
End Sub

' ---- G94 enum日本語ラベル(19章§3・15章§0の変換表。§3「生の英字enumを画面に
'   出さない」をテンプレ側の変換表の存在で担保) ----
Private Sub T_Labels()
    Dim doc As String
    Dim want2 As String
    Dim want3 As String
    doc = DocText()
    want2 = "比較的移転しやすい;条件付き・部分的;保険化困難;" & _
            "仮説;確認済み;棄却（記録保持）;新規発見"
    want3 = "既に顕在化;1～3年;3年超;無保険;過小;重複;" & _
            "補償拡大;新種目提案;座組提案"

    ChkB "G94_10分類のラベルが変換表の記載順で全て現れる_19章§3", _
        ((LenB(MissingOf(doc, CAT_LABELS)) = 0) And InOrder(doc, CAT_LABELS)), _
        "不足=[" & MissingOf(doc, CAT_LABELS) & "] 順=" & InOrder(doc, CAT_LABELS)

    ChkAll "G94_移転可能性3値と仮説ライフサイクル4値のラベル_19章§3", doc, want2

    ChkAll "G94_時間軸3値とギャップ3値と提案3値のラベル_19章§3", doc, want3
End Sub

' ---- 意図的に未テスト(期待値が18章・14章§6・19章から一意に定まらないもの) ----
'   甘い期待値を置いて実装を追認しないため、ここへ列挙して空白のまま残す。
'   (a)【統合時に解決・要追認】HeadHtml / BodyShellHtml / RuntimeJs の引数。
'      §4.4は関数名と持ち物だけ、14章§6は「本章は宣言を持たない」。実装側の
'      HeadHtml(themeName, coverFields)/BodyShellHtml(coverFields)/RuntimeJs()
'      へ呼び先を合わせた(執筆時の仮定 titleText から変更。期待値は不変)。
'   (b)【統合時に解決・要追認】BuildDocument 第3引数 coverFields の書式。§1.1⑤は
'      引数名だけ。実装の取り決め=vbTab区切りの3値へ素材を合わせた(執筆時の
'      仮定はJSON。COVER_CO / COVER_CO_ESC の期待値は不変)。
'   (c)【要裁定】JsStringSafe をどちらが呼ぶか。§5.3(1)は「JsStringSafe(dataJson)
'      の結果」、§1.1④は「DATAをJSON文字列として組立」。BuildDocument 側と読んだ。
'   (d) HtmlSafe の単引用符の実体参照(&#39; か &apos; か)。16章E-47は
'      「エンティティ化」としか書かない。生で残らないことだけを当てた。
'   (e) SEC-03..08 / SEC-10..15 の登録行の全文。§4.2の例は5行分だけで、SEC-12
'      の note 本文(該当なし)は逐語で無い。描画関数名も例の5本しか示されない。
'   (f) SEC-16の「round_no<2で0件」の実行時分岐、SEC-13の20問上限と系統別上限
'      (3/5/5/10)、SEC-02の降順、SEC-06の帯5段と帯番号、SEC-05の常に10行。JS側
'      の実行時挙動で層(a)からは観測できない(実挙動はT-33のH検収へ回す)。
'   (g)【欠落・要裁定】ヒアリングシート生成の純部。13章§2.16 は
'      BuildHearingSheet(caseId) が s4_hearing_questions から整形すると書くが、
'      14章§6の宣言はこのBoolean 1本だけで、q_no採番(1..N・最大10)・seq引継ぎ・
'      answer_memo空という純部の関数名が無い。命名権は§6なので1本も書いて
'      いない(§3.4のSEC-13は G93 で当てた)。
'   (h)【欠落・要裁定】modUICase の enum変換表の純関数名。19章§6/17章§4-2は
'      「変換表を modUICase の定数として持ちdiffゼロをテスト」と命じるが、14章§6
'      に modUICase の宣言が1本も無い。18章§3が要求するテンプレ側の日本語
'      ラベルを G94 で当てるに留めた。
'   (i) mock素材(ResponseById)からの組立。11 mock ID の本文は15章§8.1が要点
'      しか書かず、DATA(18章§2)は meta を伴う別構造(値源は13章§2.1)なので、
'      戻り値をそのまま BuildDocument へ渡す手順が章に無い。G86/G88 は15章
'      スキーマのキー名どおりに自給した DATA で当てた。
' ============================
