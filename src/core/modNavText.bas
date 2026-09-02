Attribute VB_Name = "modNavText"
Option Explicit

' ============================================================================
' modNavText - 貼付テキストの純変換(core層・純関数のみ・層(a)テスト対象)
' ----------------------------------------------------------------------------
' 11章§7.2(a) が求める4本の純関数の実装:
'   StripDrFooter   既知フッターの除去(11章§3.3.5)
'   PreviewLines    プレビュー5行の組み立て(11章§3.3.2)
'   SplitFieldNotes 現場メモの見出し切り分け(11章§3.3.6)
'   JoinFieldNotes  その逆(可逆性を層(a)で検査する)
'
' なぜ core なのか: どれも Excel を1つも触らない文字列変換であり、ui層へ置くと
'   層(a)からテストできない。**画面が無くても正しさを確かめられる形**にしておく。
'
' R4(12章§4): core層なので Excel トークンを1つも書かない。
' 12章§4: core層に製品固有の語彙(製品名・シート名)を書かない。
'
' 【申し送り】フッター語の一覧の正は 15章§2.0 である(11章§3.3.5)。15章は本波の
'   担当範囲外(並行作業中)のため、いまは本モジュールの NT_FOOTER_WORDS が実体を
'   持つ。15章§2.0 へ登記されたら、この定数の横に「15章§2.0が正」と1行足し、
'   語が増えたときは15章だけを直せばよい状態にすること。
' ============================================================================

' 既知フッターの語(行頭一致で判定する。前後の空白は除いて比べる)。
' 社内ディープリサーチのアプリが返答の末尾へ付けてくる表示であり、そのまま
' AIへ渡すと会社の内部コードや他の案件の質問文が企業分析の材料に混ざる。
' 15章§2.0が正。
Private Const NT_FOOTER_WORDS As String = _
    "役職コード:" & vbLf & "役職コード：" & vbLf & _
    "部課コード:" & vbLf & "部課コード：" & vbLf & _
    "ご利用にあたって" & vbLf & "【履歴一覧】" & vbLf & "履歴一覧"

' 本文の先頭から何割より後ろでの一致だけを採るか(11章§3.3.5 規則5)。
' 本文そのものが「ご利用にあたって」で始まる文書だったときに全部消さない保険。
Private Const NT_FOOTER_MIN_RATIO As Double = 0.5

' 現場メモの4見出し(11章§3.3.3(b))。順序はこの並びが正で、入力の並びが違っても
' 組み直すときはこの順にする(2箇所で違う並びを作らない)。
Private Const NT_HEAD_SALES As String = "【営業メモ】"
Private Const NT_HEAD_PREV As String = "【前回更新メモ】"
Private Const NT_HEAD_COVER As String = "【付保の見立て】"
Private Const NT_HEAD_OTHER As String = "【そのほか】"

' 先置きした例文の目印。この語で始まる行は保存しない(11章§3.3.3(b))。
Private Const NT_EXAMPLE_MARK As String = "例: "

' プレビュー1行の上限字数(11章§3.3.2)。
Public Const NT_PREVIEW_COLS As Long = 120

' ============================================================================
' StripDrFooter - 末尾の既知フッターを切り落とす(11章§3.3.5)。
' ----------------------------------------------------------------------------
'   1. 行に分ける
'   2. 行頭が既知の語に一致する行の位置を全部拾う
'   3. **本文の先頭から50%より後ろ**の一致だけを候補にする(規則5の保険)
'   4. 候補のうち**いちばん先頭に近いもの**から末尾までを落とす(落とす量を最大化)
'   5. 落としたあと末尾の空行を落とす
'   6. これ以外の整形は一切しない(「情報なし」等はそのまま残す)
' ============================================================================
Public Function StripDrFooter(ByVal bodyText As String) As String
    StripDrFooter = bodyText
    If LenB(bodyText) = 0 Then Exit Function

    Dim norm As String
    norm = NormalizeEol(bodyText)

    Dim lines() As String
    lines = Split(norm, vbLf)

    Dim words() As String
    words = Split(NT_FOOTER_WORDS, vbLf)

    Dim minLimit As Double
    minLimit = CDbl(Len(norm)) * NT_FOOTER_MIN_RATIO

    Dim pos As Long
    Dim i As Long
    Dim cutAt As Long
    cutAt = -1
    pos = 0
    For i = LBound(lines) To UBound(lines)
        If CDbl(pos) >= minLimit Then
            If IsFooterLine(lines(i), words) Then
                cutAt = i
                Exit For
            End If
        End If
        pos = pos + Len(lines(i)) + 1
    Next i

    If cutAt < 0 Then
        ' 一致なし = 無変更(改行の揃えもしない。整形は呼び出し側の役目)。
        StripDrFooter = bodyText
        Exit Function
    End If

    Dim kept As String
    If cutAt = LBound(lines) Then
        kept = vbNullString
    Else
        Dim keptArr() As String
        ReDim keptArr(0 To cutAt - LBound(lines) - 1)
        For i = LBound(lines) To cutAt - 1
            keptArr(i - LBound(lines)) = lines(i)
        Next i
        kept = Join(keptArr, vbLf)
    End If

    StripDrFooter = TrimTrailingBlankLines(kept)
End Function

' 行頭が既知の語のいずれかに一致するか(前後の空白は除いて比べる)。
Private Function IsFooterLine(ByVal lineText As String, ByRef words() As String) As Boolean
    Dim s As String
    s = Trim$(lineText)
    If LenB(s) = 0 Then Exit Function

    Dim i As Long
    For i = LBound(words) To UBound(words)
        Dim w As String
        w = words(i)
        If LenB(w) > 0 Then
            If Len(s) >= Len(w) Then
                If StrComp(Left$(s, Len(w)), w, vbBinaryCompare) = 0 Then
                    IsFooterLine = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ============================================================================
' PreviewLines - 先頭 n 行を vbLf 連結で返す(11章§3.3.2)。
'   各行は NT_PREVIEW_COLS 字で切り、切ったら末尾へ省略記号を付ける。
'   **表示専用**であり、ここから本文を復元する経路は作らない。
' ============================================================================
Public Function PreviewLines(ByVal bodyText As String, ByVal maxLines As Long) As String
    If LenB(bodyText) = 0 Then Exit Function
    If maxLines <= 0 Then Exit Function

    Dim lines() As String
    lines = Split(NormalizeEol(bodyText), vbLf)

    Dim n As Long
    n = UBound(lines) - LBound(lines) + 1
    If n > maxLines Then n = maxLines

    Dim outArr() As String
    ReDim outArr(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        outArr(i) = modUIGeom.ClipToChars(lines(LBound(lines) + i), NT_PREVIEW_COLS)
    Next i
    PreviewLines = Join(outArr, vbLf)
End Function

' ============================================================================
' SplitFieldNotes - 現場メモを2本の保存先へ切り分ける(11章§3.3.6)。
' ----------------------------------------------------------------------------
'   memoText   = 【営業メモ】【前回更新メモ】【付保の見立て】の3節を、
'                **見出しごと** vbLf で連結して1本にしたもの(節が空でも見出しは残す)
'   othersText = 【そのほか】の節と、どの見出しにも属さない先頭部
'   ・「例: 」で始まる行はどの節でも落とす(先置きの例文をAIへ渡さない)
'   ・見出しが1つも見つからないときは全文を memoText へ入れ、先頭へ
'     【営業メモ】を1行足す(利用者が見出しを消しても壊れない)
'   ・見出しの判定は**前後の空白を除いた行の完全一致**(本文中に見出し語が
'     現れても見出しとして扱わない)
'   戻り値: 見出しが1つでも見つかったか(呼び出し側の分岐用)
' ============================================================================
Public Function SplitFieldNotes(ByVal bodyText As String, ByRef memoText As String, _
                                ByRef othersText As String) As Boolean
    memoText = vbNullString
    othersText = vbNullString

    Dim salesBody As String
    Dim prevBody As String
    Dim coverBody As String
    Dim otherBody As String
    Dim leadBody As String

    Dim lines() As String
    lines = Split(NormalizeEol(bodyText), vbLf)

    Dim cur As String
    Dim seen As Boolean
    Dim i As Long
    cur = vbNullString
    For i = LBound(lines) To UBound(lines)
        Dim s As String
        s = Trim$(lines(i))
        If StrComp(s, NT_HEAD_SALES, vbBinaryCompare) = 0 Then
            cur = NT_HEAD_SALES
            seen = True
        ElseIf StrComp(s, NT_HEAD_PREV, vbBinaryCompare) = 0 Then
            cur = NT_HEAD_PREV
            seen = True
        ElseIf StrComp(s, NT_HEAD_COVER, vbBinaryCompare) = 0 Then
            cur = NT_HEAD_COVER
            seen = True
        ElseIf StrComp(s, NT_HEAD_OTHER, vbBinaryCompare) = 0 Then
            cur = NT_HEAD_OTHER
            seen = True
        ElseIf IsExampleLine(lines(i)) Then
            ' 先置きの例文。どの節でも保存しない。
        Else
            Select Case cur
            Case NT_HEAD_SALES
                salesBody = AppendLine(salesBody, lines(i))
            Case NT_HEAD_PREV
                prevBody = AppendLine(prevBody, lines(i))
            Case NT_HEAD_COVER
                coverBody = AppendLine(coverBody, lines(i))
            Case NT_HEAD_OTHER
                otherBody = AppendLine(otherBody, lines(i))
            Case Else
                leadBody = AppendLine(leadBody, lines(i))
            End Select
        End If
    Next i

    If Not seen Then
        memoText = NT_HEAD_SALES
        Dim lead1 As String
        lead1 = TrimTrailingBlankLines(leadBody)
        If LenB(lead1) > 0 Then memoText = memoText & vbLf & lead1
        SplitFieldNotes = False
        Exit Function
    End If

    memoText = SectionText(NT_HEAD_SALES, salesBody) & vbLf & _
               SectionText(NT_HEAD_PREV, prevBody) & vbLf & _
               SectionText(NT_HEAD_COVER, coverBody)

    Dim others1 As String
    others1 = TrimTrailingBlankLines(leadBody)
    Dim others2 As String
    others2 = TrimTrailingBlankLines(otherBody)
    If LenB(others1) > 0 And LenB(others2) > 0 Then
        othersText = others1 & vbLf & others2
    ElseIf LenB(others1) > 0 Then
        othersText = others1
    Else
        othersText = others2
    End If
    SplitFieldNotes = True
End Function

' ============================================================================
' CoverageNoteOf - 【付保の見立て】節の本文だけを取り出す(13章§2.11(d) v2.6)。
' ----------------------------------------------------------------------------
'   入力は SplitFieldNotes が返した memoText(3節を見出しごと連結した1本)、
'   または現場メモの原文のどちらでもよい(見出しの判定規約は同じ)。
'   戻り値は**見出し行を除いた本文**で、節が空なら空文字。
'   ・見出しの判定は前後の空白を除いた行の完全一致(SplitFieldNotes と同じ)
'   ・「例: 」で始まる行は落とす(先置きの例文をAIへ渡さない)
'   ・次の見出しが現れたらそこで打ち切る
'   これは input_coverage_note への**派生(読み取り専用の写し)**を作るための
'   関数であり、input_memo からは何も落とさない(二重保存。裁定書25 S1)。
' ============================================================================
Public Function CoverageNoteOf(ByVal bodyText As String) As String
    Dim lines() As String
    lines = Split(NormalizeEol(bodyText), vbLf)

    Dim inSection As Boolean
    Dim acc As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim s As String
        s = Trim$(lines(i))
        If StrComp(s, NT_HEAD_COVER, vbBinaryCompare) = 0 Then
            inSection = True
        ElseIf IsHeadLine(s) Then
            inSection = False
        ElseIf inSection Then
            If Not IsExampleLine(lines(i)) Then acc = AppendLine(acc, lines(i))
        End If
    Next i
    CoverageNoteOf = TrimTrailingBlankLines(acc)
End Function

' 4見出しのいずれかの行か(前後の空白を除いた完全一致)。
Private Function IsHeadLine(ByVal s As String) As Boolean
    IsHeadLine = (StrComp(s, NT_HEAD_SALES, vbBinaryCompare) = 0 _
                  Or StrComp(s, NT_HEAD_PREV, vbBinaryCompare) = 0 _
                  Or StrComp(s, NT_HEAD_COVER, vbBinaryCompare) = 0 _
                  Or StrComp(s, NT_HEAD_OTHER, vbBinaryCompare) = 0)
End Function

' ============================================================================
' FitsInRows - 本文が rows 行の枠に収まるか(11章§3.3.7 の現場メモ60行の判定)。
' ----------------------------------------------------------------------------
'   True=収まる(1行1セルで rows 行以内) / False=はみ出す。
'   ・rows <= 0 は「枠が無い」ので False(fail-closed。書ける保証が無い)
'   ・空文字は 0 行として True(空を書くのは常に安全)
'   ・改行は vbCrLf / vbCr / vbLf のいずれも1つの改行として数える
'   呼び出し側は False のとき **1行も書かない**(途中まで書いて切れた枠を作らない)。
' ============================================================================
Public Function FitsInRows(ByVal bodyText As String, ByVal rows As Long) As Boolean
    If rows <= 0 Then Exit Function
    If LenB(bodyText) = 0 Then
        FitsInRows = True
        Exit Function
    End If

    Dim lines() As String
    lines = Split(NormalizeEol(bodyText), vbLf)
    FitsInRows = ((UBound(lines) - LBound(lines) + 1) <= rows)
End Function

' ============================================================================
' JoinFieldNotes - SplitFieldNotes の逆(画面の枠へ戻すときの1本化)。
'   memoText は3見出しを含む正規形であることを前提にし、末尾へ【そのほか】節を
'   足す。SplitFieldNotes(JoinFieldNotes(m, o)) が (m, o) に戻ることが契約。
' ============================================================================
Public Function JoinFieldNotes(ByVal memoText As String, ByVal othersText As String) As String
    Dim outText As String
    outText = memoText
    If LenB(outText) = 0 Then outText = SectionText(NT_HEAD_SALES, vbNullString) & vbLf & _
                                        SectionText(NT_HEAD_PREV, vbNullString) & vbLf & _
                                        SectionText(NT_HEAD_COVER, vbNullString)
    outText = outText & vbLf & NT_HEAD_OTHER
    If LenB(othersText) > 0 Then outText = outText & vbLf & othersText
    JoinFieldNotes = outText
End Function

' ============================================================================
' 内部
' ============================================================================

' 見出し1行＋本文(本文が空でも見出しは残す)。
Private Function SectionText(ByVal headText As String, ByVal bodyText As String) As String
    Dim b As String
    b = TrimTrailingBlankLines(bodyText)
    If LenB(b) = 0 Then
        SectionText = headText
    Else
        SectionText = headText & vbLf & b
    End If
End Function

' 「例: 」で始まる行か(前後の空白を除いた行頭からの一致)。
Private Function IsExampleLine(ByVal lineText As String) As Boolean
    Dim s As String
    s = Trim$(lineText)
    If Len(s) < Len(NT_EXAMPLE_MARK) Then Exit Function
    IsExampleLine = (StrComp(Left$(s, Len(NT_EXAMPLE_MARK)), NT_EXAMPLE_MARK, _
                             vbBinaryCompare) = 0)
End Function

' 改行を vbLf へ揃える(vbCrLf / vbCr / vbLf のいずれも1つの改行として扱う)。
Public Function NormalizeEol(ByVal bodyText As String) As String
    NormalizeEol = Replace$(Replace$(bodyText, vbCrLf, vbLf), vbCr, vbLf)
End Function

' 1行を足す(初回は改行を入れない)。
Private Function AppendLine(ByVal acc As String, ByVal lineText As String) As String
    If LenB(acc) = 0 Then
        AppendLine = lineText
    Else
        AppendLine = acc & vbLf & lineText
    End If
End Function

' 末尾の空行を落とす(先頭・途中の空行は残す)。
Public Function TrimTrailingBlankLines(ByVal bodyText As String) As String
    If LenB(bodyText) = 0 Then Exit Function

    Dim lines() As String
    lines = Split(NormalizeEol(bodyText), vbLf)

    Dim last As Long
    last = UBound(lines)
    Do While last >= LBound(lines)
        If LenB(Trim$(lines(last))) > 0 Then Exit Do
        last = last - 1
    Loop
    If last < LBound(lines) Then Exit Function

    Dim outArr() As String
    ReDim outArr(0 To last - LBound(lines))
    Dim i As Long
    For i = LBound(lines) To last
        outArr(i - LBound(lines)) = lines(i)
    Next i
    TrimTrailingBlankLines = Join(outArr, vbLf)
End Function
