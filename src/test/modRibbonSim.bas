Attribute VB_Name = "modRibbonSim"
Option Explicit

' ==============================================================================
' modRibbonSim - リボンちゃんの応答抽出(parseText / ExtractText / UnEscapeJSON)の
'                純層模擬(裁定書33 C-3・16章 E-63)
' ------------------------------------------------------------------------------
' 役割:
'   社内AIアドイン「リボンちゃん ver202606」が **200応答から本文を切り出す手順**
'   を、相手側のソース(log.bas の parseText 58行目〜・ExtractText 184行目〜、
'   GPT.bas の UnEscapeJSON 449行目〜)から**逐語で写した**模擬である。
'   目的は「当方が返させるJSONが、相手側の終点規則で途中から切られないこと」を
'   Excel非依存の層(a)で毎回検査することにある(15章§8・17章§4-1)。
'
'   **相手側の実装を写すことが目的**なので、当方の modJsonLite の同種の関数
'   (UnescapeJsonStr など)を流用してはならない。流用すると「当方の実装どうしで
'   つじつまが合っている」だけになり、相手の癖(3置換・終点候補の最も手前・
'   不正エスケープの素通し)を1つも検査できない。
'
'   ニセリボンちゃん(`wintest/mock_ribbon/modMockRibbon.bas`)とは別物である。
'   あちらは 12引数 Application.Run の**配管**を実機で確かめる .xlam スタブで、
'   応答本文の加工はしない(裁定書33 C-3 で役割を分けた)。
'
' 相手側の確定事実(14章§2 の 7.・16章 E-63):
'   (1) parseText はHTTPボディ**全体**へ3置換を掛ける
'       `": "`->`":"` / `content_filter_results`->`content_filter_result` /
'       `prompt_filter_results`->`prompt_filter_result`
'   (2) ExtractText の始点候補は `content":"` ;;; `text":"`(**最初に見つかった
'       候補**)。終点候補は `","` ;;; `"},` ;;; `"`+LF(**最も手前の位置**)。
'       どちらか一方でも見つからなければ定型文を返す。
'   (3) 切り出した文字列を UnEscapeJSON で戻す(不正エスケープはそのまま)。
'   帰結: モデル本文の中に `"},` があると、ボディ上は `\"},\"` なので `"},` が
'   部分一致して**そこで切れる**。`","` は本文中では `\",\"` になるので当たらず、
'   LF は本文中では `\n` になるので当たらない。**危ないのは `"},` だけ**である。
'
' R4準拠(12章§2): Excelトークン(Worksheets/Range(/Application./ThisWorkbook/
'   MsgBox/ActiveSheet)に触れない純関数だけで書く。config も読まない。
'
' 移植元: リボンちゃん ver202606(相手側)。当方に対応物なし。
' ==============================================================================

' 相手側の定型文の先頭句(log.bas ExtractText 216行目〜)。始点か終点が見つから
' なかったときだけ返る。**先頭句だけを定数にし、続く長文はその場で連結する**
' (16章 E-54 の判定は先頭一致で行うため、値源はこの1本にそろえる)。
Private Const RS_NG_HEAD As String = "レスポンスから当該テキストを抽出できません"
Private Const RS_NG_TAIL As String = "ChatGPTの仕様が変更となった可能性がありますので、AIリボンをダウンロードしたホームページの情報をご確認ください"

' 候補の区切り(相手側の ";;;")。
Private Const RS_SEP As String = ";;;"

' ==============================================================================
' SimWrapBody - モデル本文を Azure OpenAI 応答風のHTTPボディへ包む(裁定書33 C-3)
' ------------------------------------------------------------------------------
'   {"choices":[{"message":{"content":"<JSONエスケープ済み>","role":"assistant"},
'    "finish_reason":"stop"}],"usage":{"total_tokens":1}}
'   エスケープは `\` `"` LF CR TAB を JSON 規約(`\\` `\"` `\n` `\r` `\t`)で行う。
'   **ここで LF が `\n` になる**ことが、終点候補 `"`+LF が本文に当たらない理由。
' ==============================================================================
Public Function SimWrapBody(ByVal content As String) As String
    SimWrapBody = "{""choices"":[{""message"":{""content"":""" & _
                  EscapeForBody(content) & _
                  """,""role"":""assistant""},""finish_reason"":""stop""}]," & _
                  """usage"":{""total_tokens"":1}}"
End Function

' ==============================================================================
' SimParseText - 相手側 parseText(results="200")の逐語移植(裁定書33 C-3)
' ------------------------------------------------------------------------------
'   3置換 -> 始点/終点の探索 -> Mid -> UnEscape の順。相手側の変数名(str1/str2)
'   と候補の並び順をそのまま残す(写しであることを読み手に分かるようにする)。
' ==============================================================================
Public Function SimParseText(ByVal body As String) As String
    Dim rsps As String
    Dim str1 As String
    Dim str2 As String

    rsps = body
    rsps = Replace(rsps, ": ", ":")
    rsps = Replace(rsps, "content_filter_results", "content_filter_result")
    rsps = Replace(rsps, "prompt_filter_results", "prompt_filter_result")

    str1 = "content" & Chr$(34) & ":" & Chr$(34)
    str1 = str1 & RS_SEP & "text" & Chr$(34) & ":" & Chr$(34)
    str2 = Chr$(34) & "," & Chr$(34)
    str2 = str2 & RS_SEP & Chr$(34) & "},"
    str2 = str2 & RS_SEP & Chr$(34) & Chr$(10)

    SimParseText = ExtractTextSim(rsps, str1, str2)
End Function

' ==============================================================================
' SimRoundTrip - 本文を包んで切り出す1往復(裁定書33 C-3)
' ------------------------------------------------------------------------------
'   戻りが元の本文と一致すれば「リボンを通しても切られない」。一致しなければ
'   どこかで切られている(または相手側の3置換で変わっている)。
' ==============================================================================
Public Function SimRoundTrip(ByVal content As String) As String
    SimRoundTrip = SimParseText(SimWrapBody(content))
End Function

' ==============================================================================
' EscapeForBody - HTTPボディへ載せるためのJSON文字列エスケープ(裁定書33 C-3)
' ------------------------------------------------------------------------------
'   `\` `"` LF CR TAB の5種だけを規約どおりに変換し、他は素通しする(相手側は
'   受け取ったボディを解釈するだけなので、送り手の側は最小の規約で足りる)。
'   文字列連結は要素数分の配列 + Join で行う(VBAの `s = s & ch` は長文で二乗の
'   時間になるため。UnEscapeJSON 側の作法にそろえた)。
' ==============================================================================
Private Function EscapeForBody(ByVal s As String) As String
    Dim ln As Long
    Dim i As Long
    Dim ch As String
    Dim tmp() As String

    ln = Len(s)
    If ln = 0 Then Exit Function

    ReDim tmp(0 To ln - 1)
    For i = 1 To ln
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "\"
                tmp(i - 1) = "\\"
            Case Chr$(34)
                tmp(i - 1) = "\" & Chr$(34)
            Case vbLf
                tmp(i - 1) = "\n"
            Case vbCr
                tmp(i - 1) = "\r"
            Case vbTab
                tmp(i - 1) = "\t"
            Case Else
                tmp(i - 1) = ch
        End Select
    Next i

    EscapeForBody = Join(tmp, vbNullString)
End Function

' ==============================================================================
' ExtractTextSim - 相手側 ExtractText の逐語移植(log.bas 184行目〜)
' ------------------------------------------------------------------------------
'   始点: 候補を先頭から順に探し、**最初に見つかった候補**を採る(見つかった時点で
'         str1 をその候補へ置き換える=相手側と同じ)。
'   終点: 候補それぞれを `p1 + Len(str1)` 以降で探し、**最も手前の位置**を採る。
'   p1 か p2 が 0 なら定型文。そうでなければ Mid で切り出して UnEscape する。
'   相手側の Optional s(既定1)は本製品の呼び出しでは常に既定なので固定した。
' ==============================================================================
Private Function ExtractTextSim(ByVal tgtText As String, ByVal str1 As String, _
                                ByVal str2 As String) As String
    Dim p1 As Long
    Dim p2 As Long
    Dim arrStr As Variant
    Dim i As Long
    Dim pn As Long
    Dim pTemp As Long

    ' 始点
    arrStr = Split(str1, RS_SEP)
    pn = UBound(arrStr)
    For i = 0 To pn
        p1 = InStr(1, tgtText, CStr(arrStr(i)))
        If p1 > 0 Then
            str1 = CStr(arrStr(i))
            Exit For
        End If
    Next i

    ' 末尾
    arrStr = Split(str2, RS_SEP)
    pn = UBound(arrStr)
    For i = 0 To pn
        pTemp = InStr(p1 + Len(str1), tgtText, CStr(arrStr(i)))
        If pTemp > 0 Then
            If p2 = 0 Then
                p2 = pTemp
            ElseIf pTemp < p2 Then
                p2 = pTemp
            End If
        End If
    Next i

    ' 抽出
    If p1 = 0 Or p2 = 0 Then
        ExtractTextSim = RS_NG_HEAD & RS_NG_TAIL
    Else
        p1 = p1 + Len(str1)
        p2 = p2 - p1
        ExtractTextSim = UnEscapeJsonSim(Mid$(tgtText, p1, p2))
    End If
End Function

' ==============================================================================
' UnEscapeJsonSim - 相手側 UnEscapeJSON の逐語移植(GPT.bas 449行目〜)
' ------------------------------------------------------------------------------
'   `\\` `\"` `\/` `\b` `\f` `\n` `\r` `\t` `\uXXXX` を実体へ戻し、**不正な
'   エスケープはバックスラッシュごとそのまま**返す(相手側の Case Else)。
'   相手側は長さを `LenB(s) \ 2` で数えるが、これはVBAの String が2バイト/字で
'   あることに由来する `Len(s)` と同値であり、LibreOffice でも同じ本数になる
'   `Len(s)` を採った(**写しからの唯一の差**。挙動は変えていない)。
'   `vbBack` / `vbFormFeed` は LibreOffice Basic に無いため `Chr$(8)` /
'   `Chr$(12)`(同じ実体)で書く。
' ==============================================================================
Private Function UnEscapeJsonSim(ByVal s As String) As String
    Dim ln As Long
    Dim i As Long
    Dim idx As Long
    Dim tmp() As String
    Dim ch As String
    Dim code4 As String

    ln = Len(s)
    If ln = 0 Then Exit Function

    ReDim tmp(0 To ln - 1)

    For i = 1 To ln
        ch = Mid$(s, i, 1)
        If ch <> "\" Then
            tmp(idx) = ch
        Else
            i = i + 1
            ch = Mid$(s, i, 1)
            Select Case ch
                Case "\"
                    tmp(idx) = "\"
                Case Chr$(34)
                    tmp(idx) = Chr$(34)
                Case "/"
                    tmp(idx) = "/"
                Case "b"
                    tmp(idx) = Chr$(8)
                Case "f"
                    tmp(idx) = Chr$(12)
                Case "n"
                    tmp(idx) = vbLf
                Case "r"
                    tmp(idx) = vbCr
                Case "t"
                    tmp(idx) = vbTab
                Case "u"
                    code4 = Mid$(s, i + 1, 4)
                    tmp(idx) = ChrW$(CLng("&H" & code4))
                    i = i + 4
                Case Else
                    ' 不正エスケープはそのまま(相手側と同じ)。
                    tmp(idx) = "\" & ch
            End Select
        End If
        idx = idx + 1
    Next i

    UnEscapeJsonSim = Join(tmp, vbNullString)
End Function
