Attribute VB_Name = "modUIGeom"
Option Explicit

' ============================================================================
' modUIGeom - 画面の幾何計算(core層・純関数のみ・層(a)テスト対象)
' ----------------------------------------------------------------------------
' なぜ core に置くのか:
'   帯やボタンの並び・カードの高さ・表示時間は、Excelのオブジェクトを1つも
'   触らずに決まる「算数」である。ui層に置くと層(a)からテストできないため、
'   **Excel非依存の部分だけ**を core へ切り出す(11章§8.6 の流用表: notebook の
'   modChrome.bas:28 SumSpan / :48 FlowLeft / :189 TextSpan / :217 ClipToWidth /
'   :301 PillWidth と、modChrome.bas:534 ToastHeightFor / :564 ToastWaitMsFor を
'   本製品の作法へ改変して移植した)。
'
' R4(12章§4): core層なので Excel トークンを1つも書かない。
' 12章§4: core層に製品固有の語彙(製品名・シート名)を書かない。
' 単位は pt(ポイント)。文字幅は等幅ではないので「全角=fontPt・半角=fontPt/2」の
'   近似で足りる(ボタン幅の決定は多少太くても困らない。足りないより余る側へ倒す)。
' ============================================================================

' 進捗ドットの丸(CP932にある記号なのでリテラルで書いてよい。11章§7.1)。
Private Const UG_DOT_ON As String = "●"
Private Const UG_DOT_OFF As String = "○"

' 文字幅の近似。半角は全角の半分とみなす。
Private Const UG_HALF_RATIO As Double = 0.5
' 省略記号(1文字ぶんの幅を持つ)。
Private Const UG_ELLIPSIS As String = "…"

' ============================================================================
' SumSpan - 「幅1;幅2;...」の合計に、区切りの余白を (件数-1) 個ぶん足した幅。
'   空文字は 0。数値でない項は 0 として数える(壊れた指定で例外にしない)。
' ============================================================================
Public Function SumSpan(ByVal widthsCsv As String, ByVal gapPt As Double) As Double
    If LenB(widthsCsv) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(widthsCsv, ";")

    Dim i As Long
    Dim total As Double
    Dim n As Long
    For i = LBound(parts) To UBound(parts)
        If LenB(Trim$(parts(i))) > 0 Then
            total = total + Val(parts(i))
            n = n + 1
        End If
    Next i
    If n > 1 Then total = total + gapPt * CDbl(n - 1)
    SumSpan = total
End Function

' ============================================================================
' FlowLeft - 直前の部品の右端から、次の部品の左端を求める(左から右へ流す)。
'   直前が無いとき(prevRight <= 0)は originLeft をそのまま返す。
' ============================================================================
Public Function FlowLeft(ByVal originLeft As Double, ByVal prevRight As Double, _
                         ByVal gapPt As Double) As Double
    If prevRight <= 0# Then
        FlowLeft = originLeft
        Exit Function
    End If
    FlowLeft = prevRight + gapPt
    If FlowLeft < originLeft Then FlowLeft = originLeft
End Function

' ============================================================================
' TextSpan - 文字列の見かけの幅(pt)。全角1字=fontPt、半角1字=fontPt/2。
'   改行を含む場合は**いちばん長い行**の幅を返す(折り返さない前提の実測)。
' ============================================================================
Public Function TextSpan(ByVal bodyText As String, ByVal fontPt As Double) As Double
    If LenB(bodyText) = 0 Then Exit Function

    Dim lines() As String
    lines = Split(Replace$(Replace$(bodyText, vbCrLf, vbLf), vbCr, vbLf), vbLf)

    Dim i As Long
    Dim best As Double
    For i = LBound(lines) To UBound(lines)
        Dim w As Double
        w = LineSpan(lines(i), fontPt)
        If w > best Then best = w
    Next i
    TextSpan = best
End Function

' 1行ぶんの幅。半角(コード<128)は半分で数える。
Private Function LineSpan(ByVal lineText As String, ByVal fontPt As Double) As Double
    Dim i As Long
    Dim w As Double
    For i = 1 To Len(lineText)
        If AscW(Mid$(lineText, i, 1)) < 128 Then
            w = w + fontPt * UG_HALF_RATIO
        Else
            w = w + fontPt
        End If
    Next i
    LineSpan = w
End Function

' ============================================================================
' ClipToWidth - 幅に収まるところまで切り、切ったら末尾へ省略記号を付ける。
'   maxPt 以下ならそのまま返す。maxPt が省略記号1つぶんにも満たないときは
'   空文字を返す(記号だけの表示は意味が無い)。
' ============================================================================
Public Function ClipToWidth(ByVal bodyText As String, ByVal fontPt As Double, _
                            ByVal maxPt As Double) As String
    If LenB(bodyText) = 0 Then Exit Function
    If maxPt <= 0# Then Exit Function
    If TextSpan(bodyText, fontPt) <= maxPt Then
        ClipToWidth = bodyText
        Exit Function
    End If

    Dim room As Double
    room = maxPt - fontPt                      ' 省略記号1字ぶんを空けておく
    If room <= 0# Then Exit Function

    Dim i As Long
    Dim w As Double
    Dim outText As String
    For i = 1 To Len(bodyText)
        Dim ch As String
        ch = Mid$(bodyText, i, 1)
        If AscW(ch) < 128 Then
            w = w + fontPt * UG_HALF_RATIO
        Else
            w = w + fontPt
        End If
        If w > room Then Exit For
        outText = outText & ch
    Next i
    ClipToWidth = outText & UG_ELLIPSIS
End Function

' ============================================================================
' ClipToChars - 文字数で切って省略記号を付ける(プレビュー行の120字切りに使う)。
'   maxChars 以下ならそのまま返す。maxChars <= 0 は空文字。
' ============================================================================
Public Function ClipToChars(ByVal bodyText As String, ByVal maxChars As Long) As String
    If LenB(bodyText) = 0 Then Exit Function
    If maxChars <= 0 Then Exit Function
    If Len(bodyText) <= maxChars Then
        ClipToChars = bodyText
        Exit Function
    End If
    ClipToChars = Left$(bodyText, maxChars) & UG_ELLIPSIS
End Function

' ============================================================================
' PillWidth - 文字を包む丸ボタン(ピル)の幅。左右の余白を足し、下限で丸める。
' ============================================================================
Public Function PillWidth(ByVal caption As String, ByVal fontPt As Double, _
                          ByVal padPt As Double, ByVal minPt As Double) As Double
    Dim w As Double
    w = TextSpan(caption, fontPt) + padPt * 2#
    If w < minPt Then w = minPt
    PillWidth = w
End Function

' ============================================================================
' StepDots - 進捗ドットの文字列(11章§3.1.1 の `nv_step_dots`)。
'   total 個のうち current 個までを塗る。範囲外は丸めて必ず total 文字を返す。
' ============================================================================
Public Function StepDots(ByVal current As Long, ByVal total As Long) As String
    Dim n As Long
    n = total
    If n <= 0 Then Exit Function

    Dim k As Long
    k = current
    If k < 0 Then k = 0
    If k > n Then k = n

    Dim i As Long
    Dim outText As String
    For i = 1 To n
        If i <= k Then
            outText = outText & UG_DOT_ON
        Else
            outText = outText & UG_DOT_OFF
        End If
    Next i
    StepDots = outText
End Function

' ============================================================================
' CardHeightFor - 文字量からカード・トーストの高さ(pt)を見積もる。
'   colChars = 1行に入る全角字数。minPt / maxPt で丸める。
'   (notebook modChrome.bas:534 ToastHeightFor の改変移植)
' ============================================================================
Public Function CardHeightFor(ByVal bodyText As String, ByVal colChars As Long, _
                              ByVal linePt As Double, ByVal padPt As Double, _
                              ByVal minPt As Double, ByVal maxPt As Double) As Double
    Dim rows As Long
    rows = LineCountFor(bodyText, colChars)

    Dim h As Double
    h = CDbl(rows) * linePt + padPt
    If h < minPt Then h = minPt
    If maxPt > 0# And h > maxPt Then h = maxPt
    CardHeightFor = h
End Function

' 折り返しを含めた行数。空文字でも1行と数える(高さ0のカードを作らない)。
Public Function LineCountFor(ByVal bodyText As String, ByVal colChars As Long) As Long
    Dim cols As Long
    cols = colChars
    If cols <= 0 Then cols = 1

    If LenB(bodyText) = 0 Then
        LineCountFor = 1
        Exit Function
    End If

    Dim lines() As String
    lines = Split(Replace$(Replace$(bodyText, vbCrLf, vbLf), vbCr, vbLf), vbLf)

    Dim i As Long
    Dim rows As Long
    For i = LBound(lines) To UBound(lines)
        Dim n As Long
        n = (Len(lines(i)) + cols - 1) \ cols
        If n < 1 Then n = 1
        rows = rows + n
    Next i
    LineCountFor = rows
End Function

' ============================================================================
' CardWaitMsFor - 文字量に応じた表示時間(ミリ秒)。短文は短く、長文は長く。
'   (notebook modChrome.bas:564 ToastWaitMsFor の改変移植)
' ============================================================================
Public Function CardWaitMsFor(ByVal bodyText As String, ByVal minMs As Long, _
                              ByVal maxMs As Long, ByVal msPerChar As Long) As Long
    Dim ms As Long
    ms = minMs + Len(bodyText) * msPerChar
    If ms < minMs Then ms = minMs
    If maxMs > 0 And ms > maxMs Then ms = maxMs
    CardWaitMsFor = ms
End Function

' ============================================================================
' ToastSecondsFor - トースト1枚の表示秒数(11章§4.1・裁定書22 m9)。
' ----------------------------------------------------------------------------
'   秒 = 3 + 字数 / 20 を 3秒～9秒でクリップする。短い1文は3秒で消え、
'   2文の長い案内は最長9秒まで残る(読み切る前に消えたという苦情を止める)。
'   固定秒(6秒)をやめた理由: 文の長さが2倍違っても同じ時間で消えていた。
' ============================================================================
Public Function ToastSecondsFor(ByVal bodyText As String) As Double
    Dim sec As Double
    sec = 3# + CDbl(Len(bodyText)) / 20#
    If sec < 3# Then sec = 3#
    If sec > 9# Then sec = 9#
    ToastSecondsFor = sec
End Function
