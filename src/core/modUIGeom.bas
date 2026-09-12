Attribute VB_Name = "modUIGeom"
Option Explicit

' ============================================================================
' modUIGeom - 画面の幾何計算(core層・純関数のみ・層(a)テスト対象)
' ----------------------------------------------------------------------------
' なぜ core に置くのか:
'   帯やボタンの並び・カードの高さ・表示時間は、Excelのオブジェクトを1つも
'   触らずに決まる「算数」である。ui層に置くと層(a)からテストできないため、
'   **Excel非依存の部分だけ**を core へ切り出す(11章§8.6 の流用表: notebook の
'   modChrome.bas:48 FlowLeft と modChrome.bas:534 ToastHeightFor を本製品の
'   作法へ改変して移植した)。
'   **W15 Round3(裁定書41 §1)**: 移植したが本製品では一度も呼ばれなかった
'   SumSpan(:28) / ClipToWidth(:217) / PillWidth(:301) / CardWaitMsFor(:564 の
'   ToastWaitMsFor 由来)の4本と、その4本だけが使っていた TextSpan(:189) および
'   その下請け LineSpan を撤去した。ボタン幅は modUISheet が配置表の `幅pt` を
'   そのまま使い、文字切りは ClipToChars(字数)で足りている。
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

' コーチ帯の1行目で「STEP n/6」と進捗ドットを隔てる空白(裁定書26 A の逐語)。
Private Const UG_BAND_GAP As String = "  "

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
' CoachBandText - コーチ帯の図形の中に書く文字(裁定書26 A)。
' ----------------------------------------------------------------------------
' なぜ図形の中に持たせるのか: 図形は色や重ね順に関わらず**常にセルの上**に
'   描かれる。帯の地(不透明グラデーション)を敷くと、その下のセルへ書いた
'   STEP番号・進捗ドット・次の一手が実機(Windows Excel)で1文字も見えない
'   (11章§8.5 の禁忌)。文字を図形の中に持たせれば必ず見える。
' 組立て(裁定書26 A の逐語):
'   1行目 = "STEP <n>/<count>" + 空白2つ + 進捗ドット(StepDots と同一の値)
'   2行目 = 次の一手(hm_next_action の逐語。1字も変えない)
' セル(nv_step_no / nv_step_dots / hm_next_action)への書込みは従来どおり
'   維持する(層(b)・純層の読取値源であり、ここはその写しである)。
' ============================================================================
Public Function CoachBandText(ByVal stepNo As Long, ByVal stepCount As Long, _
                              ByVal actionText As String) As String
    CoachBandText = "STEP " & CStr(stepNo) & "/" & CStr(stepCount) & _
                    UG_BAND_GAP & StepDots(stepNo, stepCount) & vbLf & actionText
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
        n = CLng(Fix((Len(lines(i)) + cols - 1) / cols))
        If n < 1 Then n = 1
        rows = rows + n
    Next i
    LineCountFor = rows
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
