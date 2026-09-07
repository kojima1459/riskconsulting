Attribute VB_Name = "modRibbonWire"
Option Explicit

' ==============================================================================
' modRibbonWire - 経路側で本文が途中から切られた疑いの判定(裁定書33 C-4・16章 E-63)
' ------------------------------------------------------------------------------
' 役割:
'   受け取った生応答が「JSONとして始まっているのに閉じていない」形かどうかだけを
'   見る純関数を1本持つ。**新しいエラーコードは作らない**。既存の JSON 抽出失敗
'   (16章 E-06)の記録へ `ribbon_cut?` の1語を添えるために使う目印であり、
'   **断定ではなく疑いの印**である(切れ方が同じでも原因はモデル側の打ち切りかも
'   しれない。原因を1つに決めない)。
'
'   判定(3条件のAND):
'     (1) 前後の空白を落とした先頭が `{` または `[` である
'     (2) `{`/`[` と `}`/`]` の対応が閉じていない(文字列リテラルの中の括弧は
'         数えない。数え方は modJsonLite.ExtractJsonBlock の MatchBracePos と同じ
'         =バックスラッシュのエスケープを見る)
'     (3) 末尾の1字が `}` `]` `"` のどれでもない
'   空文字は False(切られたのではなく「空で返った」= 別の事象。16章 E-16)。
'
'   (3) を入れているのは、末尾がきちんと閉じ記号や引用符で終わっているものを
'   「切られた」と呼ばないため(モデルが構造を間違えただけの応答を、経路の
'   せいにしない)。
'
' R4準拠(12章§2): Excelトークンに触れない。core層なので製品固有の語彙
'   (製品名・シート名・Step名)も持たない。
' ==============================================================================

' ==============================================================================
' LooksRibbonCut - 本文が途中で切られた疑いがあるか(裁定書33 C-4)
' ==============================================================================
Public Function LooksRibbonCut(ByVal raw As String) As Boolean
    Dim t As String
    Dim head As String
    Dim tail As String

    LooksRibbonCut = False

    t = Trim$(raw)
    If LenB(t) = 0 Then Exit Function

    head = Left$(t, 1)
    If head <> "{" And head <> "[" Then Exit Function

    If BracketsClosed(t) Then Exit Function

    tail = Right$(t, 1)
    If tail = "}" Or tail = "]" Or tail = """" Then Exit Function

    LooksRibbonCut = True
End Function

' ==============================================================================
' CutNote - err_log の detail 末尾へ添える1語(裁定書33 C-4)
' ------------------------------------------------------------------------------
'   切られた疑いがあれば " ribbon_cut?"、無ければ ""。detail は3箇所(通常Step・
'   入念パイプ・プリフライト)で組み立てるので、**語そのものはここ1箇所に置く**
'   (同じ語を3つのモジュールへ書き写すと、片方だけ直る形になるため)。
'   末尾の `?` は「疑い」の印であり、断定ではない(16章 E-63)。
' ==============================================================================
Public Function CutNote(ByVal raw As String) As String
    If LooksRibbonCut(raw) Then CutNote = " ribbon_cut?"
End Function

' ==============================================================================
' BracketsClosed - 括弧の対応が閉じているか(文字列リテラルの中は数えない)
' ------------------------------------------------------------------------------
'   modJsonLite の MatchBracePos と同じ数え方。走査の終わりで深さが0かつ文字列の
'   中で終わっていなければ「閉じている」。深さが負になった時点でも閉じていると
'   みなす(閉じ過ぎは「切られた」ではない=別の壊れ方)。
' ==============================================================================
Private Function BracketsClosed(ByVal s As String) As Boolean
    Dim n As Long
    Dim i As Long
    Dim ch As String
    Dim depth As Long
    Dim inQuote As Boolean
    Dim esc As Boolean

    n = Len(s)
    depth = 0
    inQuote = False
    esc = False

    For i = 1 To n
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
                    BracketsClosed = True
                    Exit Function
                End If
            End If
        End If
    Next i

    BracketsClosed = (depth = 0 And Not inQuote)
End Function
