Attribute VB_Name = "modTestsPure11"
Option Explicit

' ============================
' modTestsPure11 - HTML/JSエスケープの契約テスト(17章§4-1 層(a))
' ----------------------------
' テスト本数: 20本 = G84 HtmlSafe 6 / G85 JsStringSafe 8 / G86 危険JSONの実弾 6
' 役割: 16章 E-47(1)(2)・NFR-S7③・18章§5.3(1)(2) が定める**外部由来テキストの
'   HTML/JS書込口**(modUtilText.HtmlSafe / JsStringSafe)の入出力を固定する。
'   実装を1行も読まず、章の本文だけを根拠に期待値を置く(modTestsPure8..10 と同じ作法)。
' 由来: W3.1 で modTestsPure10 が30,000字契約(12章§2)を超えたため、17章§1
'   「警告帯に入ったモジュールは次に本体へ手を入れる波の前に切り出す」に従い
'   **関数単位で**切り出した(G84/G85/G86 の3群と、その素材 DataJsonAttack)。
'   関数名は変えていない。DataJsonAttack だけは modTestsPure10 の G88(全文組立)も
'   同じ素材を使うため Public にした(攻撃素材を2箇所に写経しない)。
' 結線: modTestsPure10.RunAll の末尾から本 RunAll を呼ぶ(数珠つなぎ
'   modTestsPure -> 2 .. -> 10 -> 11。単体では0本)。
' 設計判断(R4): Excelトークン・乱数・時刻を使わず改行は vbLf。素材は架空値で、
'   JSONは単引用符で書き AsJson() で二重引用符へ直す。
' ============================

' RunAll: グループ隔離実行(未実装/未注入は GroupFail で可視化)。
Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 3
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

    ' 数珠つなぎの継続: W4.1(裁定書9)の回帰テストへ(modTestsPure -> 2 .. -> 11 -> 12)
    On Error Resume Next
    Err.Clear
    modTestsPure12.RunAll
    If Err.Number <> 0 Then
        GroupFail "modTestsPure12.RunAll"
        Err.Clear
    End If
    On Error GoTo 0
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
    End Select
End Sub

' ---- 共通ヘルパ(modTestsPure3..10 と同じ作法) ----
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

Private Function DQ() As String
    DQ = Chr$(34)
End Function

Private Function AsJson(ByVal s As String) As String
    AsJson = Replace(s, "'", DQ())
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

' "\u00XX"(16進の大小はどちらでも可)を判定する。
Private Function IsUEsc(ByVal act As String, ByVal hex2 As String) As Boolean
    If Len(act) <> 6 Then Exit Function
    If Left$(act, 4) <> "\u00" Then Exit Function
    IsUEsc = (UCase$(Mid$(act, 5)) = UCase$(hex2))
End Function

' 素材: 18章§2 の DATA(キー名は15章スキーマのまま)。危険トークンは risk_name=
'   閉じscript+img / scenario=引用符つきonerror / pitch=引用符と改行 /
'   evidence.quote=HTMLコメント開始+開きscript(18章§5.3(1)v1.1が名指しする白紙化経路。
'   `-->` を伴わない形が本命)。構造側の生の改行で§5.3(1)④も通す。
Public Function DataJsonAttack() As String
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
    s = s & "'evidence':{'quote':'<!--<script>A&B','source':'hp'}," & vbLf
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
'   ① \ -> \\ ・ " -> \"  ② すべての < -> \u003C  ③ U+2028/U+2029  ④ 制御文字
Private Sub T_JsSafe()
    ChkS "G85_逆斜線を二重化する_18章§5.3", _
        modUtilText.JsStringSafe("\"), "\\"

    ChkS "G85_二重引用符を逆斜線で逃がす_18章§5.3", _
        modUtilText.JsStringSafe(DQ()), "\" & DQ()

    ' ①の内部順(逆斜線が先)。逆だと \" が \\" になり文字列が閉じる。
    ChkS "G85_JSONのエスケープ済み引用符は逆斜線3本になる_18章§5.3", _
        modUtilText.JsStringSafe("\" & DQ()), "\\\" & DQ()

    ' ②が①の後であること(逆なら \u003C の逆斜線が二重化される)。v1.1で
    ' 「</ -> <\/」から「すべての < -> \u003C」へ改訂した(18章§5.3(1))。
    ChkS "G85_閉じscriptタグを無害化する_18章§5.3", _
        modUtilText.JsStringSafe("</script>"), "\u003C/script>"

    ' v1.1: 単独の < も対象。生の < を1文字も残さないことがページ白紙化
    ' (script data double escaped)を構造的に塞ぐ唯一の条件(18章§5.3(1))。
    ChkB "G85_単独の不等号も含め生の記号が1つも残らない_18章§5.3", _
        ((modUtilText.JsStringSafe("<a/b>") = "\u003Ca/b>") And _
         (modUtilText.JsStringSafe("a</b") = "a\u003C/b") And _
         (InStr(modUtilText.JsStringSafe("<!--<script>"), "<") = 0)), _
        "実際=[" & modUtilText.JsStringSafe("<!--<script>") & "]"

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
'   削除ではない)ことを同時に当てる。v1.1の受入条件は「生の < が1文字も無い」。
Private Sub T_Payload()
    Dim e As String
    e = modUtilText.JsStringSafe(DataJsonAttack())

    ChkB "G86_出力に生の不等号が1文字も残らない_16章E-47", _
        ((InStr(e, "<") = 0) And (InStr(e, "</script>") = 0)), _
        "残存位置=" & InStr(e, "<") & " 出力頭=[" & HeadOf(e) & "]"

    ChkB "G86_出力に生の改行が1つも残らない_18章§5.3", _
        ((InStr(e, vbLf) = 0) And (InStr(e, vbCr) = 0)), _
        "LF位置=" & InStr(e, vbLf) & " CR位置=" & InStr(e, vbCr)

    ' 逃がし漏れが1つでもあればJS文字列リテラルが閉じる。
    ChkB "G86_全ての二重引用符が逆斜線を伴う_18章§5.3", _
        (CountOcc(e, DQ()) = CountOcc(e, "\" & DQ())), _
        "引用符=" & CountOcc(e, DQ()) & " 逃がし済=" & CountOcc(e, "\" & DQ())

    ChkB "G86_危険トークンは削除ではなく無害化される_18章§5.3", _
        Ctn(e, "\u003C/script>\u003Cimg src=x onerror=alert(1)>"), _
        "出力頭=[" & HeadOf(e) & "]"

    ' 18章§5.3(1)v1.1が名指しする白紙化経路。`-->` を伴わない <!--<script> が
    ' 生のまま残ると script data double escaped 状態へ入りDATAの </script> が
    ' 終端として働かなくなる。削除ではなく \u003C へ落ちていることまで当てる。
    ChkB "G86_HTMLコメント開始と開きscriptも無害化される_18章§5.3", _
        (Ctn(e, "\u003C!--\u003Cscript>") And (InStr(e, "<!--") = 0)), _
        "生の注入位置=" & InStr(e, "<!--")

    ' 属性値の引用符は JSON の \" を経て \\\" になる。
    ChkB "G86_属性値の引用符が逆斜線3本を伴う_18章§5.3", _
        Ctn(e, "onerror=\\\" & DQ() & "alert(1)\\\" & DQ()), _
        "onerror位置=" & InStr(e, "onerror")
End Sub
