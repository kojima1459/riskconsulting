Attribute VB_Name = "modJsonLite"
Option Explicit

' ============================================================================
' modJsonLite - LLM応答からのJSON抽出と、必要キーだけの取り出し
' ----------------------------------------------------------------------------
' 役割(14章§5 JSON防衛線の(1)と(2)):
'   (1) ExtractJsonBlock : 説明文・コードフェンス・全角記号が混じった生応答から
'       【最外の1本のJSON】だけを切り出す。切り出せなければ "" を返す。
'   (2) GetStr / GetLong / GetBoolJ / GetArrayItems : 切り出した本文から
'       【対象スキーマに必要なキーだけ】を取り出す。汎用JSONパーサ(木を組む
'       実装)は作らない。作ると壊れた入力で例外を投げる口が増えるうえ、
'       30,000字契約を1モジュールで食い潰す。
'
' 受入基準の正は 14章§5 の「ExtractJsonBlock 入力パターン表」(7本)。
'   (1)前後に説明文 / (2)コードフェンス / (3)フェンス閉じ忘れ / (4)末尾途切れ /
'   (5)JSON2連結 / (6)全角波括弧・全角引用符 / (7)値内の生改行とエスケープ引用符。
'   **(4)は補完しない**(欠落を捏造せず "" を返し、上位の修復リトライへ渡す)。
'
' 外部ライブラリ(JsonConverter等)は使わない。正規表現オブジェクトも使わない
'   (CreateObjectはLibreOffice実行テストで挙動が揃わず、PoCで作法を固めた
'   「純文字列走査」だけで7パターンすべてを満たせるため)。
'
' 移植元: PoC「マイ本棚AI」 src/core/modGatewayDirect.bas の EscapeJsonStr
'   (制御文字を \uXXXX へ落とす方針とサロゲート安全性)。抽出・キー取り出しは
'   14章§5の表に合わせた新設。
'
' R4準拠: Excelトークン(Worksheets/Range/Application/ThisWorkbook/MsgBox/
'   ActiveSheet)には触れない。全関数が純ロジックであり、17章 T-11 の方針どおり
'   Excel非依存のテスト層から全数試験できる。
'
' エラー規約(14章§6): 本モジュールは外部由来テキストを【返さない】=解析結果
'   だけを返すため、失敗は "" / 既定値 / 0件コレクションで表す。例外は投げない。
' ============================================================================

' コードフェンスの目印(3連バッククォート)。行頭がこれで始まる行は丸ごと捨てる。
Private Const FENCE_MARK As String = "```"

' 走査で「空白」とみなす文字(半角空白・タブ・CR・LF)。
Private Const WS_CHARS As String = " " & vbTab & vbCr & vbLf

' Long の範囲(GetLong で範囲外を既定値へ倒すため。CLngの実行時エラーを避ける)。
Private Const LONG_MAX_D As Double = 2147483647#
Private Const LONG_MIN_D As Double = -2147483648#

' 直近の ExtractJsonBlock が行った全角->半角の置換件数(14章§5 前処理P0)。
' 上位(modPipeline / modPlayOps)が run_log detail へ `fw_normalized=n` として
' 記録するための帯域外の値。ExtractJsonBlock の呼び出しごとに0へ戻す。
Private gFwNormalized As Long

' 直近の ExtractJsonBlock が破棄した2本目以降のJSONの有無(14章§5 パターン(5))。
' 上位が run_log detail へ `extra_json=1` として記録する。0 または 1。
Private gExtraJson As Long

' ============================================================================
' ExtractJsonBlock - 生応答から最外の1本のJSONを切り出す(14章§5(1))。
' ----------------------------------------------------------------------------
'   手順: 前処理P0(全角記号の一律半角化) -> フェンス行の除去 -> 最初の "{" から
'   対応が閉じる "}" までの切り出し。閉じなければ "" (補完しない)。
'   2本目以降が続いていた場合は1本目だけを返し、gExtraJson=1 を立てる。
'
'   フェンス行の除去は「行頭(前後空白を除く)が3連バッククォートの行」を対象と
'   する。閉じフェンスが無くても開始フェンスだけ消えるのでパターン(3)を満たす。
'   値の本文中にフェンス行が現れると巻き添えでJSONが壊れるが、その場合も
'   推測補完はせず "" を返して修復リトライへ落とす(捏造しないほうを採る)。
' ============================================================================
Public Function ExtractJsonBlock(ByVal raw As String) As String
    gFwNormalized = 0
    gExtraJson = 0
    ExtractJsonBlock = vbNullString
    If LenB(raw) = 0 Then Exit Function

    Dim txt As String
    txt = NormalizeFullWidth(raw)
    txt = StripFenceLines(txt)

    Dim startPos As Long
    startPos = InStr(1, txt, "{", vbBinaryCompare)
    If startPos = 0 Then Exit Function

    Dim endPos As Long
    endPos = MatchBracePos(txt, startPos)
    If endPos = 0 Then Exit Function

    ExtractJsonBlock = Mid$(txt, startPos, endPos - startPos + 1)

    If InStr(endPos + 1, txt, "{", vbBinaryCompare) > 0 Then gExtraJson = 1
End Function

' ============================================================================
' LastFwNormalized / LastExtraJson - 直近の ExtractJsonBlock の副次情報。
' ----------------------------------------------------------------------------
'   14章§5 は置換件数と破棄件数の run_log 記録を要求するが、§6 の
'   ExtractJsonBlock シグネチャには出力引数が無い。シグネチャを変えずに値を
'   運ぶため、モジュール変数＋読み出し関数の形で帯域外に出す(§6の表は
'   closed ではないため公開関数の追加は契約違反にならない)。
' ============================================================================
Public Function LastFwNormalized() As Long
    LastFwNormalized = gFwNormalized
End Function

Public Function LastExtraJson() As Long
    LastExtraJson = gExtraJson
End Function

' ============================================================================
' GetStr - キーの値を文字列として取り出す(14章§5(2))。
' ----------------------------------------------------------------------------
'   ・文字列値はエスケープを解いて返す。値の途中の生CR/LFは改行1個、生タブは
'     タブ1個として扱う(§5 パターン(7)の実体: 走査時に "\n" / "\t" へ寄せてから
'     エスケープ解除する)。
'   ・値の終端は「直後に , } ] のいずれか(空白を挟んでよい)か文末が続く "」で
'     判定する。`\"` は終端とみなさない(§5 パターン(7))。
'   ・数値・真偽値はトークン文字列(例 "12" / "true")をそのまま返す。null は ""。
'   ・値がオブジェクト/配列のときは "" を返す(配列は GetArrayItems を使う)。
'   ・同名キーが入れ子にもある場合は【最初に現れたもの】を返す。スキーマごとに
'     必要なキーだけを引く用途に限る(汎用パーサではない)。
' ============================================================================
Public Function GetStr(ByVal json As String, ByVal key As String) As String
    GetStr = vbNullString

    Dim valPos As Long
    valPos = FindKeyValuePos(json, key)
    If valPos = 0 Then Exit Function

    Dim ch As String
    ch = Mid$(json, valPos, 1)

    If ch = """" Then
        Dim endPos As Long
        GetStr = ReadStringValue(json, valPos, endPos)
        Exit Function
    End If
    If ch = "{" Or ch = "[" Then Exit Function

    Dim tok As String
    tok = ReadScalarToken(json, valPos)
    If LCase$(tok) = "null" Then Exit Function
    GetStr = tok
End Function

' ============================================================================
' GetLong - キーの値を Long として取り出す。取れなければ dflt。
' ----------------------------------------------------------------------------
'   数値以外・空・Longの範囲外はすべて dflt。小数はゼロ方向へ切り捨てる。
'   Val は小数点を "." で解釈する(ロケール非依存)ため CDbl ではなく Val を使う。
' ============================================================================
Public Function GetLong(ByVal json As String, ByVal key As String, _
                        ByVal dflt As Long) As Long
    GetLong = dflt

    Dim tok As String
    tok = Trim$(GetStr(json, key))
    If LenB(tok) = 0 Then Exit Function
    If Not IsNumberToken(tok) Then Exit Function

    Dim d As Double
    d = Val(tok)
    If d > LONG_MAX_D Then Exit Function
    If d < LONG_MIN_D Then Exit Function
    GetLong = CLng(Fix(d))
End Function

' ============================================================================
' GetBoolJ - キーの値を真偽値として取り出す。取れなければ dflt。
' ----------------------------------------------------------------------------
'   関数名の末尾 J は、VBAの Bool 変換関数群との紛れを避けるための 14章§6 の
'   命名(GetBool ではなく GetBoolJ が契約名)。
'   LLMは true/false のかわりに "1"/"0"/"yes"/"no" を返すことがあるため受ける。
' ============================================================================
Public Function GetBoolJ(ByVal json As String, ByVal key As String, _
                         ByVal dflt As Boolean) As Boolean
    GetBoolJ = dflt

    Dim v As String
    v = LCase$(Trim$(GetStr(json, key)))
    Select Case v
        Case "true", "1", "yes"
            GetBoolJ = True
        Case "false", "0", "no"
            GetBoolJ = False
    End Select
End Function

' ============================================================================
' GetArrayItems - キーの配列要素を Collection で返す(要素はすべて String)。
' ----------------------------------------------------------------------------
'   ・文字列要素はエスケープを解いた本文、オブジェクト/配列要素は【その要素の
'     JSON本文そのまま】(呼び出し側がさらに GetStr で引ける)、数値・真偽値は
'     トークン文字列を入れる。
'   ・キーが無い/配列でない/空配列のときは 0件のコレクションを返す。
'     Nothing は返さない(呼び出し側の Is Nothing 判定漏れで落ちないため)。
' ============================================================================
Public Function GetArrayItems(ByVal json As String, ByVal key As String) As Collection
    Dim items As Collection
    Set items = New Collection
    Set GetArrayItems = items

    Dim valPos As Long
    valPos = FindKeyValuePos(json, key)
    If valPos = 0 Then Exit Function
    If Mid$(json, valPos, 1) <> "[" Then Exit Function

    Dim n As Long
    n = Len(json)

    Dim depth As Long
    Dim itemStart As Long
    Dim i As Long
    Dim q As Long
    Dim ch As String

    depth = 1
    i = valPos + 1
    itemStart = SkipWs(json, i)

    Do While i <= n
        ch = Mid$(json, i, 1)
        If ch = """" Then
            q = StringEndPos(json, i, False)
            If q = 0 Then Exit Do
            i = q + 1
        ElseIf ch = "{" Or ch = "[" Then
            depth = depth + 1
            i = i + 1
        ElseIf ch = "}" Or ch = "]" Then
            depth = depth - 1
            If depth <= 0 Then
                AddArrayItem json, itemStart, i - 1, items
                Exit Do
            End If
            i = i + 1
        ElseIf ch = "," And depth = 1 Then
            AddArrayItem json, itemStart, i - 1, items
            itemStart = SkipWs(json, i + 1)
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End Function

' ============================================================================
' EscapeJsonStr - 文字列をJSONの文字列リテラル本体へ落とす(囲みの " は付けない)。
' ----------------------------------------------------------------------------
'   移植元PoCと同じ方針: \ と " と改行・復帰・タブを短縮エスケープへ、
'   それ以外の制御文字(コード32未満)を \uXXXX へ。サロゲートは対で残るので
'   1コードユニットずつ通しても壊れない。
' ============================================================================
Public Function EscapeJsonStr(ByVal s As String) As String
    EscapeJsonStr = vbNullString
    Dim n As Long
    n = Len(s)
    If n = 0 Then Exit Function

    Dim parts() As String
    ReDim parts(1 To n)

    Dim i As Long
    Dim ch As String
    Dim code As Long
    For i = 1 To n
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "\"
                parts(i) = "\\"
            Case """"
                parts(i) = "\" & """"
            Case vbLf
                parts(i) = "\n"
            Case vbCr
                parts(i) = "\r"
            Case vbTab
                parts(i) = "\t"
            Case Else
                code = AscW(ch)
                If code < 0 Then code = code + 65536
                If code < 32 Then
                    parts(i) = "\u" & Right$("000" & Hex$(code), 4)
                Else
                    parts(i) = ch
                End If
        End Select
    Next i

    EscapeJsonStr = Join(parts, vbNullString)
End Function

' ============================================================================
' UnescapeJsonStr - JSONの文字列リテラル本体をVBA文字列へ戻す。
' ----------------------------------------------------------------------------
'   \" \\ \/ \b \f \n \r \t \uXXXX を解く。未知のエスケープは壊さずそのまま
'   残す(捏造しない方針の一貫)。\uXXXX のサロゲートは1つずつ ChrW で積むと
'   UTF-16の対として正しく復元される。
' ============================================================================
Public Function UnescapeJsonStr(ByVal s As String) As String
    If InStr(1, s, "\", vbBinaryCompare) = 0 Then
        UnescapeJsonStr = s
        Exit Function
    End If

    Dim n As Long
    n = Len(s)
    Dim parts() As String
    ReDim parts(1 To n)

    Dim i As Long
    Dim ch As String
    Dim nx As String
    Dim code As Long
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch <> "\" Or i = n Then
            parts(i) = ch
            i = i + 1
        Else
            nx = Mid$(s, i + 1, 1)
            Select Case nx
                Case """"
                    parts(i) = """"
                Case "\"
                    parts(i) = "\"
                Case "/"
                    parts(i) = "/"
                Case "b"
                    parts(i) = Chr$(8)
                Case "f"
                    parts(i) = Chr$(12)
                Case "n"
                    parts(i) = vbLf
                Case "r"
                    parts(i) = vbCr
                Case "t"
                    parts(i) = vbTab
                Case "u"
                    If i + 5 <= n And IsHex4(Mid$(s, i + 2, 4)) Then
                        code = CLng("&H" & Mid$(s, i + 2, 4))
                        If code > 32767 Then code = code - 65536
                        parts(i) = ChrW(code)
                        i = i + 4
                    Else
                        parts(i) = "\u"
                    End If
                Case Else
                    parts(i) = "\" & nx
            End Select
            i = i + 2
        End If
    Loop

    UnescapeJsonStr = Join(parts, vbNullString)
End Function

' ============================================================================
' NormalizeFullWidth - 前処理P0(14章§5 パターン(6)の実体)。
' ----------------------------------------------------------------------------
'   構造記号を一律で半角化する。値の日本語本文中の全角コロン・全角読点まで
'   巻き込むが、JSON全体が壊れて修復リトライに落ちる損失のほうが大きいと
'   §5が裁定しているため一律置換を採る。置換件数は gFwNormalized に積む。
'   全角文字はソースへ直接書かず ChrW で指定する(U+FF02 はCP932に無いため
'   リテラルで書くとVBEでの保持に耐えない)。
' ============================================================================
Private Function NormalizeFullWidth(ByVal s As String) As String
    Dim t As String
    t = s
    t = ReplaceCount(t, ChrW(&HFF5B&), "{")
    t = ReplaceCount(t, ChrW(&HFF5D&), "}")
    t = ReplaceCount(t, ChrW(&HFF3B&), "[")
    t = ReplaceCount(t, ChrW(&HFF3D&), "]")
    t = ReplaceCount(t, ChrW(&H201C&), """")
    t = ReplaceCount(t, ChrW(&H201D&), """")
    t = ReplaceCount(t, ChrW(&HFF02&), """")
    t = ReplaceCount(t, ChrW(&HFF1A&), ":")
    t = ReplaceCount(t, ChrW(&HFF0C&), ",")
    NormalizeFullWidth = t
End Function

' 置換しつつ件数を gFwNormalized へ積む小道具。
Private Function ReplaceCount(ByVal s As String, ByVal fromText As String, _
                              ByVal toText As String) As String
    Dim p As Long
    p = InStr(1, s, fromText, vbBinaryCompare)
    If p = 0 Then
        ReplaceCount = s
        Exit Function
    End If

    Dim cnt As Long
    cnt = 0
    Do While p > 0
        cnt = cnt + 1
        p = InStr(p + Len(fromText), s, fromText, vbBinaryCompare)
    Loop

    gFwNormalized = gFwNormalized + cnt
    ReplaceCount = Replace(s, fromText, toText)
End Function

' ============================================================================
' StripFenceLines - コードフェンス行を落とす(14章§5 パターン(2)(3))。
' ----------------------------------------------------------------------------
'   行の前後空白を除いた先頭が3連バッククォートなら、その行を丸ごと捨てる。
'   CRLF は vbLf で分割しても各行末に CR が残るだけなので、連結し直せば
'   もとの改行が保たれる(値内の生改行を壊さない = パターン(7)と両立する)。
' ============================================================================
Private Function StripFenceLines(ByVal s As String) As String
    If InStr(1, s, FENCE_MARK, vbBinaryCompare) = 0 Then
        StripFenceLines = s
        Exit Function
    End If

    Dim rows() As String
    rows = Split(s, vbLf)

    Dim keep() As String
    ReDim keep(0 To UBound(rows))

    Dim cnt As Long
    Dim i As Long
    cnt = 0
    For i = LBound(rows) To UBound(rows)
        If Left$(LTrim$(rows(i)), Len(FENCE_MARK)) <> FENCE_MARK Then
            keep(cnt) = rows(i)
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        StripFenceLines = vbNullString
        Exit Function
    End If

    ReDim Preserve keep(0 To cnt - 1)
    StripFenceLines = Join(keep, vbLf)
End Function

' ============================================================================
' MatchBracePos - openPos の "{" に対応する "}" の位置。閉じなければ 0。
' ----------------------------------------------------------------------------
'   文字列リテラルの中の括弧・引用符は数えない(バックスラッシュのエスケープを
'   見る)。値の中の生CR/LFは中身として素通しするのでパターン(7)でも成功する。
'   パターン(4)(末尾途切れ)はここが 0 を返し、上位が "" を返す = 補完しない。
' ============================================================================
Private Function MatchBracePos(ByVal s As String, ByVal openPos As Long) As Long
    MatchBracePos = 0

    Dim n As Long
    n = Len(s)

    Dim depth As Long
    Dim inQuote As Boolean
    Dim esc As Boolean
    Dim i As Long
    Dim ch As String

    depth = 0
    inQuote = False
    esc = False

    For i = openPos To n
        ch = Mid$(s, i, 1)
        If inQuote Then
            If esc Then
                esc = False
            ElseIf ch = "\" Then
                esc = True
            ElseIf ch = """" Then
                inQuote = False
            End If
        Else
            If ch = """" Then
                inQuote = True
            ElseIf ch = "{" Or ch = "[" Then
                depth = depth + 1
            ElseIf ch = "}" Or ch = "]" Then
                depth = depth - 1
                If depth <= 0 Then
                    If ch = "}" Then MatchBracePos = i
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ============================================================================
' FindKeyValuePos - キー名に対応する値の先頭位置。見つからなければ 0。
' ----------------------------------------------------------------------------
'   文字列を1本ずつ拾い、直後(空白可)が ":" のものだけをキーとみなす。値として
'   現れた文字列は ":" が続かないので自然に読み飛ばされる。
' ============================================================================
Private Function FindKeyValuePos(ByVal s As String, ByVal keyName As String) As Long
    FindKeyValuePos = 0

    Dim n As Long
    n = Len(s)

    Dim i As Long
    Dim q As Long
    Dim j As Long
    Dim ch As String
    Dim keyRaw As String

    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = """" Then
            q = StringEndPos(s, i, False)
            If q = 0 Then Exit Do
            keyRaw = Mid$(s, i + 1, q - i - 1)
            j = SkipWs(s, q + 1)
            If j <= n Then
                If Mid$(s, j, 1) = ":" Then
                    If keyRaw = keyName Then
                        FindKeyValuePos = SkipWs(s, j + 1)
                        Exit Function
                    ElseIf UnescapeJsonStr(keyRaw) = keyName Then
                        FindKeyValuePos = SkipWs(s, j + 1)
                        Exit Function
                    End If
                End If
            End If
            i = q + 1
        Else
            i = i + 1
        End If
    Loop
End Function

' ============================================================================
' StringEndPos - quotePos の開き引用符に対応する閉じ引用符の位置。無ければ 0。
' ----------------------------------------------------------------------------
'   strictEnd=True のときだけ 14章§5 パターン(7)の終端規則を使う:
'   「直後に , } ] のいずれか(空白を挟んでよい)か文末が続く "」だけを終端と
'   みなす。エスケープされていない生の " が値の途中に混ざっていても、その手前で
'   切らずに読み切れる。キー側は直後が ":" なのでこの規則を使えない
'   (strictEnd=False = 最初のエスケープされていない " が終端)。
' ============================================================================
Private Function StringEndPos(ByVal s As String, ByVal quotePos As Long, _
                              ByVal strictEnd As Boolean) As Long
    StringEndPos = 0

    Dim n As Long
    n = Len(s)

    Dim i As Long
    Dim nxt As Long
    Dim ch As String

    i = quotePos + 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "\" Then
            i = i + 2
        ElseIf ch = """" Then
            If Not strictEnd Then
                StringEndPos = i
                Exit Function
            End If
            nxt = SkipWs(s, i + 1)
            If nxt > n Then
                StringEndPos = i
                Exit Function
            End If
            ch = Mid$(s, nxt, 1)
            If ch = "," Or ch = "}" Or ch = "]" Then
                StringEndPos = i
                Exit Function
            End If
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End Function

' ============================================================================
' ReadStringValue - 開き引用符の位置から文字列値を読み、エスケープを解いて返す。
'   endPos には閉じ引用符の位置を返す(閉じていなければ 0)。
' ============================================================================
Private Function ReadStringValue(ByVal s As String, ByVal quotePos As Long, _
                                 ByRef endPos As Long) As String
    ReadStringValue = vbNullString
    endPos = 0

    Dim q As Long
    q = StringEndPos(s, quotePos, True)
    If q = 0 Then Exit Function
    endPos = q

    Dim body As String
    body = Mid$(s, quotePos + 1, q - quotePos - 1)
    ReadStringValue = UnescapeJsonStr(NormalizeRawControls(body))
End Function

' ============================================================================
' NormalizeRawControls - 値の中の生の制御文字をJSONのエスケープ表記へ寄せる。
' ----------------------------------------------------------------------------
'   14章§5 パターン(7)の実体: 生のCR/LFは "\n" へ、生のタブは "\t" へ。
'   CR+LF の2文字は改行1個として1本の "\n" にまとめる(2行に増やさない)。
'   ここで表記へ寄せてから UnescapeJsonStr を通すので、最終的な戻り値では
'   もとどおり1個の改行/タブに戻る。
' ============================================================================
Private Function NormalizeRawControls(ByVal s As String) As String
    Dim n As Long
    n = Len(s)
    If n = 0 Then
        NormalizeRawControls = vbNullString
        Exit Function
    End If
    If InStr(1, s, vbCr, vbBinaryCompare) = 0 _
       And InStr(1, s, vbLf, vbBinaryCompare) = 0 _
       And InStr(1, s, vbTab, vbBinaryCompare) = 0 Then
        NormalizeRawControls = s
        Exit Function
    End If

    Dim parts() As String
    ReDim parts(1 To n)

    Dim i As Long
    Dim ch As String
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = vbCr Then
            parts(i) = "\n"
            If i < n Then
                If Mid$(s, i + 1, 1) = vbLf Then i = i + 1
            End If
        ElseIf ch = vbLf Then
            parts(i) = "\n"
        ElseIf ch = vbTab Then
            parts(i) = "\t"
        Else
            parts(i) = ch
        End If
        i = i + 1
    Loop

    NormalizeRawControls = Join(parts, vbNullString)
End Function

' ============================================================================
' ReadScalarToken - 数値・真偽値・null のトークンを , } ] 空白 の手前まで読む。
' ============================================================================
Private Function ReadScalarToken(ByVal s As String, ByVal valPos As Long) As String
    Dim n As Long
    n = Len(s)

    Dim i As Long
    Dim ch As String
    For i = valPos To n
        ch = Mid$(s, i, 1)
        If InStr(1, WS_CHARS, ch, vbBinaryCompare) > 0 Then Exit For
        If ch = "," Or ch = "}" Or ch = "]" Then Exit For
    Next i

    ReadScalarToken = Mid$(s, valPos, i - valPos)
End Function

' ============================================================================
' AddArrayItem - 配列の1要素(startPos..endPos)をコレクションへ積む。
'   空要素([] や末尾カンマ)は積まない。
' ============================================================================
Private Sub AddArrayItem(ByVal s As String, ByVal startPos As Long, _
                         ByVal endPos As Long, ByVal items As Collection)
    If startPos <= 0 Then Exit Sub
    If endPos < startPos Then Exit Sub

    Dim t As String
    t = Trim$(Mid$(s, startPos, endPos - startPos + 1))
    If LenB(t) = 0 Then Exit Sub

    If Len(t) >= 2 And Left$(t, 1) = """" And Right$(t, 1) = """" Then
        items.Add UnescapeJsonStr(NormalizeRawControls(Mid$(t, 2, Len(t) - 2)))
    Else
        items.Add t
    End If
End Sub

' ============================================================================
' SkipWs - pos 以降で最初の非空白の位置。全部空白なら Len(s)+1。
' ============================================================================
Private Function SkipWs(ByVal s As String, ByVal pos As Long) As Long
    Dim n As Long
    n = Len(s)

    Dim i As Long
    i = pos
    If i < 1 Then i = 1
    Do While i <= n
        If InStr(1, WS_CHARS, Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Do
        i = i + 1
    Loop
    SkipWs = i
End Function

' ============================================================================
' IsNumberToken - JSONの数値トークンとして読めるか(桁が1つ以上あること)。
'   記号の位置までは厳密に見ない。ここへ来るのは値の切り出し済みトークンだけで、
'   誤って通しても Val が 0 相当に落ちるだけだから。
' ============================================================================
Private Function IsNumberToken(ByVal s As String) As Boolean
    IsNumberToken = False

    Dim n As Long
    n = Len(s)
    If n = 0 Then Exit Function

    Dim i As Long
    Dim ch As String
    Dim numCount As Long
    numCount = 0
    For i = 1 To n
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            numCount = numCount + 1
        ElseIf ch = "-" Or ch = "+" Or ch = "." Or ch = "e" Or ch = "E" Then
            ' 符号・小数点・指数は許す
        Else
            Exit Function
        End If
    Next i

    IsNumberToken = (numCount > 0)
End Function

' ============================================================================
' IsHex4 - 4桁の16進数字か(\uXXXX の妥当性判定)。
' ============================================================================
Private Function IsHex4(ByVal s As String) As Boolean
    IsHex4 = False
    If Len(s) <> 4 Then Exit Function

    Dim i As Long
    Dim ch As String
    For i = 1 To 4
        ch = UCase$(Mid$(s, i, 1))
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "F")) Then Exit Function
    Next i

    IsHex4 = True
End Function
