Attribute VB_Name = "modGround"
Option Explicit

' ============================================================================
' modGround - evidence.quote の「原文に実在するか」の照合(裁定書37 B-03/C-4)
' ----------------------------------------------------------------------------
' 12章§2: app層の**純文字列モジュール**(R4)。Worksheets / Range( /
'   Application. / ThisWorkbook / MsgBox / ActiveSheet に触れない。案件データも
'   config も読まない(受け取るのは引数だけ)。値源の解決と注記の保持は
'   modPipeline3、呼び出しは modPipeline.Defend の後ろ1行が行う。
'
' 存在理由(B班報告 §1・§3 B-03): S1 が原文を構造化したあと、**後段は誰も原文を
'   見ない**。modValidate:268 の V-S2-04 は `LenB()=0`(空チェック)だけで、値が
'   貼付原文に実在するかを見る仕組みが製品全体で0件だった。ここはその1点を
'   埋めるだけの照合器であり、**落とさない・修復リトライを起こさない**
'   (伝書鳩3-4「落とすとリトライ地獄」)。結果は run_log の detail と
'   HTMLレポート SEC-14 の「原文照合」列にだけ出す。SEC-04 の充足度バッジは
'   別の事実なので動かさない(伝書鳩1-5「混ぜない」)。
'
' 照合の考え方(表記揺れに強く、捏造に弱く):
'   ・正規化(NormalizeForMatch)で、全角英数記号を半角へ寄せ、空白・改行・
'     句読点・鍵括弧・中黒・カンマ・ピリオド・長音を落とし、英字を小文字化する。
'     **数字は落とさない**(金額・年月の桁が肝であり、捏造はそこに出るため)。
'   ・引用の先頭 headChars 字(既定20)だけを見る。LLMは引用の末尾を伸ばしたり
'     要約したりするが、先頭は原文どおりのことが多い。
'   ・20字未満の引用は全長で一致を見る(短い引用で当たりを甘くしない)。
'
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ============================================================================

' 先頭何字を見るか(config ground_head_chars の既定。13章§2.3)。
Public Const GR_HEAD_DEFAULT As Long = 20

' 未照合の risk_no を並べるときの区切り(modPipeline2.LastDeepOutcome と同型の
'   1本文字列で ui へ運ぶため)。
Private Const GR_SEP As String = ";"

' 照合の対象**外**にする evidence.source(19章§3の source enum のうち、入力に
'   由来しないと自称するもの)。これ以外(hp / yuho / memo / contract /
'   prev_renewal)は「入力に書いてある」と名乗っているので照合する。
Private Const GR_SKIP_SOURCES As String = ";inference;knowledge;"

' ============================================================================
' NormalizeForMatch - 照合用の正規化(純関数)
' ----------------------------------------------------------------------------
'   (1) 全角英数記号(U+FF01..U+FF5E)を半角へ寄せ、全角空白(U+3000)は半角空白へ。
'   (2) 空白・タブ・改行、句読点・鍵括弧・中黒・カンマ・ピリオド・長音、および
'       意味を持たない記号類を落とす。
'   (3) 英大文字を小文字へ。**数字・かな・漢字はそのまま残す**。
' ============================================================================
Public Function NormalizeForMatch(ByVal t As String) As String
    Dim i As Long, n As Long, code As Long
    Dim dropSet As String
    Dim outArr() As String, outCount As Long

    n = Len(t)
    If n = 0 Then Exit Function
    dropSet = DropChars()
    ReDim outArr(1 To n)
    outCount = 0

    For i = 1 To n
        code = AscW(Mid$(t, i, 1))
        If code < 0 Then code = code + 65536
        ' 全角英数記号 -> 半角(U+FF01..U+FF5E は U+0021..U+007E と 0xFEE0 差)
        If code >= 65281 And code <= 65374 Then code = code - 65248
        ' 全角空白 -> 半角空白(次の行で落ちる)
        If code = 12288 Then code = 32
        ' 英大文字 -> 小文字
        If code >= 65 And code <= 90 Then code = code + 32
        If InStr(1, dropSet, ChrW$(code), vbBinaryCompare) = 0 Then
            outCount = outCount + 1
            outArr(outCount) = ChrW$(code)
        End If
    Next i

    If outCount = 0 Then Exit Function
    ReDim Preserve outArr(1 To outCount)
    NormalizeForMatch = Join(outArr, vbNullString)
End Function

' 落とす文字の集合。Const にすると vbTab 等を畳めない環境があるため関数で持つ
'   (値源はここ1箇所)。半角へ寄せたあとの文字だけを並べればよい。
Private Function DropChars() As String
    DropChars = " " & vbTab & vbCr & vbLf & _
                "!""#$%&'()*+,-./:;<=>?@[\]^_`{|}~" & _
                "、。・「」『』【】〔〕ー"
End Function

' ============================================================================
' QuoteFound - 引用が原文に実在するか(純関数)
' ----------------------------------------------------------------------------
'   正規化した quote の先頭 headChars 字を、正規化した haystack から探す。
'   ・quote が正規化後に空 -> False(照合できないものを「見つかった」にしない)
'   ・haystack が正規化後に空 -> False(**呼出側が先に「検査しない」を選ぶ**。
'     B-03 テスト観点(5)。ここで True を返すと捏造が素通りする)
'   ・headChars が1未満 -> 既定20。quote が headChars 未満なら全長で見る
' ============================================================================
Public Function QuoteFound(ByVal quote As String, ByVal haystack As String, _
                           ByVal headChars As Long) As Boolean
    QuoteFound = FoundIn(quote, NormalizeForMatch(haystack), headChars)
End Function

' 正規化済みの haystack を受ける内側(GroundNotes が原文を1回だけ正規化して
'   使い回すためのもの。数万字の再正規化をリスク件数ぶん繰り返さない)。
Private Function FoundIn(ByVal quote As String, ByVal hayNorm As String, _
                         ByVal headChars As Long) As Boolean
    Dim q As String, nHead As Long

    If LenB(hayNorm) = 0 Then Exit Function
    q = NormalizeForMatch(quote)
    If LenB(q) = 0 Then Exit Function

    nHead = headChars
    If nHead < 1 Then nHead = GR_HEAD_DEFAULT
    If Len(q) < nHead Then nHead = Len(q)
    FoundIn = (InStr(1, hayNorm, Left$(q, nHead), vbBinaryCompare) > 0)
End Function

' ============================================================================
' GroundNotes - S2出力の未照合リスト(純関数)
' ----------------------------------------------------------------------------
'   risks[](evidence.quote / evidence.source)と emerging_risks[]
'   (evidence_quote / evidence_source)を走査し、source が inference /
'   knowledge **以外**のものだけ照合する。見つからなかったものを ";" 区切りで
'   返す(全部見つかれば "")。
'   ・risks[] は risk_no をそのまま並べる。
'   ・emerging_risks[] は 15章 Schema-S2 に risk_no が**無い**ため、配列の
'     出現順で "E1" "E2" と採番する(HTML側の突き合わせもこの番号で行う)。
'   ・source が空のものは照合しない(enum違反は V-S2-03 の担当。ここでは
'     「入力由来と名乗っていない」として扱う)。
'   ・haystack が空なら**何も返さない**(呼出側が ground_skipped を記録する)。
' ============================================================================
Public Function GroundNotes(ByVal s2Json As String, ByVal haystack As String, _
                            ByVal headChars As Long) As String
    Dim hayNorm As String, res As String
    Dim it As Variant, rj As String, noText As String
    Dim idx As Long

    hayNorm = NormalizeForMatch(haystack)
    If LenB(hayNorm) = 0 Then Exit Function

    For Each it In modJsonLite.GetArrayItems(s2Json, "risks")
        rj = CStr(it)
        If Checkable(modJsonLite.GetStr(rj, "source")) Then
            If Not FoundIn(modJsonLite.GetStr(rj, "quote"), hayNorm, headChars) Then
                noText = Trim$(modJsonLite.GetStr(rj, "risk_no"))
                If LenB(noText) = 0 Then noText = "?"
                res = AddNo(res, noText)
            End If
        End If
    Next it

    idx = 0
    For Each it In modJsonLite.GetArrayItems(s2Json, "emerging_risks")
        rj = CStr(it)
        idx = idx + 1
        If Checkable(modJsonLite.GetStr(rj, "evidence_source")) Then
            If Not FoundIn(modJsonLite.GetStr(rj, "evidence_quote"), hayNorm, headChars) Then
                res = AddNo(res, "E" & CStr(idx))
            End If
        End If
    Next it

    GroundNotes = res
End Function

' 照合の対象か(空・inference・knowledge は対象外)。
Private Function Checkable(ByVal sourceText As String) As Boolean
    Dim s As String

    s = LCase$(Trim$(sourceText))
    If LenB(s) = 0 Then Exit Function
    Checkable = (InStr(1, GR_SKIP_SOURCES, GR_SEP & s & GR_SEP, vbBinaryCompare) = 0)
End Function

' ============================================================================
' NoteCount - ";" 区切りの未照合リストの件数(run_log の ground_unmatched=n)。
' ============================================================================
Public Function NoteCount(ByVal noteText As String) As Long
    Dim it As Variant, n As Long

    If LenB(Trim$(noteText)) = 0 Then Exit Function
    For Each it In Split(noteText, GR_SEP)
        If LenB(Trim$(CStr(it))) > 0 Then n = n + 1
    Next it
    NoteCount = n
End Function

' ============================================================================
' NoteJsonArray - 未照合リストを18章§2 meta の `ground_unmatched` 用のJSON配列
'   本文へ変換する(角括弧は呼出側が付ける)。要素は文字列。
' ============================================================================
Public Function NoteJsonArray(ByVal noteText As String) As String
    Dim it As Variant, res As String, one As String

    For Each it In Split(noteText, GR_SEP)
        one = Trim$(CStr(it))
        If LenB(one) > 0 Then
            If LenB(res) > 0 Then res = res & ","
            res = res & """" & modJsonLite.EscapeJsonStr(one) & """"
        End If
    Next it
    NoteJsonArray = res
End Function

' ";" 区切りの積み上げ(重複は先に出たものを残す)。
Private Function AddNo(ByVal acc As String, ByVal oneText As String) As String
    AddNo = acc
    If LenB(oneText) = 0 Then Exit Function
    If InStr(1, GR_SEP & acc & GR_SEP, GR_SEP & oneText & GR_SEP, vbBinaryCompare) > 0 Then Exit Function
    If LenB(AddNo) > 0 Then AddNo = AddNo & GR_SEP
    AddNo = AddNo & oneText
End Function
