Attribute VB_Name = "modUIGuide"
Option Explicit

' ============================================================================
' modUIGuide - 初回ガイドツアーと操作ガイドのボタン(ui層・裁定書14 裁定6)
' ----------------------------------------------------------------------------
' 役割:
'   はじめてブックを開いた人に「①調べる→②貼る→③まとめて作る→④⑤出す」の3枚の
'   カードを順番に見せる。完了・スキップは config `guide_tour_done`="1" として
'   残し、以後は出さない。操作ガイドシートの[ツアーをもう一度見る]から何度でも
'   見直せる。あわせて操作ガイドシートの2つの図形ボタン([テストを実行]・
'   [ツアーをもう一度見る])を生やす(gd_ アンカーの上に置く)。
'
' 防衛設計(PoC「マイ本棚AI」modTour の流儀をそのまま採る):
'   1. エントリは StartTourIfFirstRun の1本(modBoot の画面用意手順の後の1行)。
'      既存モジュールの中身は書き換えない。
'   2. サーキットブレーカー: 全Publicエントリの最上部で On Error Resume Next。
'      本機能の失敗は「カードが出ない・消えない」に留め、本体機能へ波及させない。
'   3. クリックハンドラは modUIProgress.TryEnterUiLock / ExitUiLock で連打を防ぐ
'      (16章 E-11。ui層のハンドラの作法と同じ)。
'   4. Shape名は全て "gt_" 接頭辞。既存の btn_ / lbl_ / btncopy_ と衝突しない。
'      削除は**名前を配列へ集めてから**行う(列挙中に削除すると添字がずれる)。
'   5. **Hyperlinks.Add は使わない**。図形にハイパーリンクを付けるとOnActionが
'      効かなくなる(実機で確定した知見)。画面遷移は OnAction のみで行う。
'
' フラグの既定値(13章§2.3): guide_tour_done = "0"。"1" 以外の値(空・壊れた値)は
'   すべて「まだ見ていない」と解釈して表示側へ倒す(見せ損なうより出す)。
' ============================================================================

Private Const UG_SHEET As String = "HOME"
Private Const UG_GUIDE_SHEET As String = "操作ガイド"
Private Const UG_PREFIX As String = "gt_"
Private Const UG_FLAG As String = "guide_tour_done"
Private Const UG_STEPS As Long = 3
Private Const UG_LOCK_NAME As String = "はじめの案内"

' 13章§2.10 の名前付きレンジ(カードの位置の基準。セル番地は書かない)。
Private Const UG_ANCHOR As String = "hm_case_id"
' 操作ガイドのボタンアンカー(13章§2.18)。
Private Const UG_BTN_TEST As String = "gd_btn_test"
Private Const UG_BTN_TOUR As String = "gd_btn_tour"

' Excel組み込み定数の数値(modUISheet と同じ流儀で名前を書かない)。
Private Const UG_SHAPE_ROUNDED As Long = 5      ' msoShapeRoundedRectangle
Private Const UG_SHAPE_TEXTBOX As Long = 1      ' msoTextOrientationHorizontal
Private Const UG_PLACEMENT_FREE As Long = 3     ' xlFreeFloating
Private Const UG_ALIGN_LEFT As Long = 1         ' xlLeft
Private Const UG_ALIGN_CENTER As Long = 2       ' xlCenter

' 配色(裁定書14 裁定7と同じ値。RGB(r,g,b) = r + g*256 + b*65536)。
Private Const UG_COLOR_SURFACE As Long = 16777215&   ' 白
Private Const UG_COLOR_ACCENT As Long = 5990145&     ' 濃緑(RGB 1,103,91)
Private Const UG_COLOR_TEXT As Long = 2562065&       ' 文字(RGB 17,24,39)
Private Const UG_FONT As String = "Yu Gothic UI"

' カードの大きさ(実機で1枚目と2枚目が動くと目線が迷子になるため3枚とも同位置)。
' 高さは本文の実測(TextFrame2.AutoSize)で確定させ、UG_CARD_H を下限にする
' (文言を短くしても間延びせず、伸ばしても見切れない。裁定書17 H3(b))。
Private Const UG_CARD_W As Double = 540#
Private Const UG_CARD_H As Double = 150#
Private Const UG_BODY_TOP As Double = 52#
Private Const UG_BODY_PAD As Double = 62#
Private Const UG_AUTOSIZE_ON As Long = 1        ' msoAutoSizeShapeToFitText
Private Const UG_WRAP_ON As Long = -1           ' msoTrue

' 現在の段(1..3)。永続しない画面制御変数(14章§6の状態保持の例外に当たらない)。
Private gStep As Long

' ============================================================================
' StartTourIfFirstRun - 起動時の唯一のエントリ(modBoot から1行で呼ぶ)。
'   すでに見た人には何もしない。まだの人には1枚目を描く。
' ============================================================================
Public Sub StartTourIfFirstRun()
    On Error Resume Next

    NoticeKbMissingOnce

    If TourDone() Then Exit Sub
    ClearTour
    DrawCard 1
End Sub

' ============================================================================
' RestartTour - 操作ガイドの[ツアーをもう一度見る]の OnAction。
' ============================================================================
Public Sub RestartTour()
    On Error Resume Next
    If Not modUIProgress.TryEnterUiLock(UG_LOCK_NAME) Then Exit Sub
    modUISheet.ShowSheet UG_SHEET
    ClearTour
    DrawCard 1
    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' OnTourNext - [次へ] / 最終段の[はじめる]。
' ============================================================================
Public Sub OnTourNext()
    On Error Resume Next
    If Not modUIProgress.TryEnterUiLock(UG_LOCK_NAME) Then Exit Sub

    Dim nextStep As Long
    nextStep = gStep + 1
    If nextStep > UG_STEPS Then
        ClearTour
        MarkDone
    Else
        DrawCard nextStep
    End If

    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' OnTourSkip - [スキップ]。完了扱いにして閉じる(再視聴は操作ガイドから)。
' ============================================================================
Public Sub OnTourSkip()
    On Error Resume Next
    If Not modUIProgress.TryEnterUiLock(UG_LOCK_NAME) Then Exit Sub
    ClearTour
    MarkDone
    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' EnsureGuideButtons - 操作ガイドシートの図形ボタン(modUIHome.EnsureScreens から
'   他の Ensure*Buttons と同じ並びで呼ぶ)。アンカーが無ければ何もしない。
' ============================================================================
Public Sub EnsureGuideButtons()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UG_GUIDE_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    Dim c As Long

    r = modUISheet.BlockRow(UG_BTN_TEST)
    c = modUISheet.BlockCol(UG_BTN_TEST)
    If r > 0 And c > 0 Then
        modUISheet.EnsureButtonEx ws, "btn_gd_test", "テストを実行", r, c, 132#, _
                                  "modTestsRunnerUi.RunAllTestsFromBook", "plain"
    End If

    r = modUISheet.BlockRow(UG_BTN_TOUR)
    c = modUISheet.BlockCol(UG_BTN_TOUR)
    If r > 0 And c > 0 Then
        modUISheet.EnsureButtonEx ws, "btn_gd_tour", "ツアーをもう一度見る", r, c, 168#, _
                                  "modUIGuide.RestartTour", "plain"
    End If
End Sub

' ============================================================================
' ClearTour - gt_ の図形を全部消す(名前を集めてから消す)。
' ============================================================================
Public Sub ClearTour()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UG_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim shapeNames() As String
    ReDim shapeNames(0 To ws.Shapes.count)

    Dim n As Long
    Dim i As Long
    n = 0
    For i = 1 To ws.Shapes.count
        If InStr(1, ws.Shapes(i).Name, UG_PREFIX, vbBinaryCompare) = 1 Then
            shapeNames(n) = ws.Shapes(i).Name
            n = n + 1
        End If
    Next i

    For i = 0 To n - 1
        modUISheet.DropShape ws, shapeNames(i)
    Next i
End Sub

' ============================================================================
' 内部
' ============================================================================

' ツアー済みか。"1" 以外はすべて未完了として扱う(表示側へ倒す)。
Private Function TourDone() As Boolean
    TourDone = (StrComp(Trim$(modConfig.GetStr(UG_FLAG, "0")), "1", vbBinaryCompare) = 0)
End Function

Private Sub MarkDone()
    On Error Resume Next
    modConfig.SetValue UG_FLAG, "1"
End Sub

' 段 n(1..3)のカードを描く。先に前の段を消してから置く。
Private Sub DrawCard(ByVal n As Long)
    On Error Resume Next

    ClearTour

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UG_SHEET)
    If ws Is Nothing Then Exit Sub

    gStep = n

    Dim cardL As Double
    Dim cardT As Double
    cardL = CardLeft(ws)
    cardT = CardTop(ws)

    Dim card As Object
    Set card = ws.Shapes.AddShape(UG_SHAPE_ROUNDED, cardL, cardT, UG_CARD_W, UG_CARD_H)
    If card Is Nothing Then Exit Sub
    card.Name = UG_PREFIX & "card"
    card.Placement = UG_PLACEMENT_FREE
    card.Fill.ForeColor.RGB = UG_COLOR_SURFACE
    card.Line.Visible = True
    card.Line.Weight = 1.25
    card.Line.ForeColor.RGB = UG_COLOR_ACCENT
    card.Adjustments(1) = 0.08

    PutText ws, UG_PREFIX & "step", CStr(n) & " / " & CStr(UG_STEPS), _
            cardL + UG_CARD_W - 80#, cardT + 14#, 60#, 18#, 9#, False, UG_ALIGN_CENTER, False
    PutText ws, UG_PREFIX & "title", TitleOf(n), _
            cardL + 20#, cardT + 16#, UG_CARD_W - 110#, 26#, 13#, True, UG_ALIGN_LEFT, False
    ' 本文だけ AutoSize=1 で伸ばし、その実測でカードの高さを決める(見切れ根絶)。
    PutText ws, UG_PREFIX & "body", BodyOf(n), _
            cardL + 20#, cardT + UG_BODY_TOP, UG_CARD_W - 40#, 40#, 11#, False, _
            UG_ALIGN_LEFT, True

    Dim cardH As Double
    cardH = UG_BODY_TOP + ShapeHeight(ws, UG_PREFIX & "body") + UG_BODY_PAD
    If cardH < UG_CARD_H Then cardH = UG_CARD_H
    card.Height = cardH

    ' ボタン2つ(左=スキップ・右=次へ/はじめる)。**カード確定後の下端**を基準に
    ' 置く(本文が伸びてもボタンがカードから外れない)。Hyperlinks.Add は使わない。
    Dim btnT As Double
    btnT = cardT + cardH - 46#
    PutButton ws, UG_PREFIX & "skip", "スキップ", cardL + 20#, btnT, 108#, _
              "modUIGuide.OnTourSkip", False
    PutButton ws, UG_PREFIX & "next", NextLabelOf(n), _
              cardL + UG_CARD_W - 168#, btnT, 148#, "modUIGuide.OnTourNext", True
End Sub

' カードの左端(HOMEの表示領域の中央寄せ。取れないときは既定値)。
Private Function CardLeft(ByVal ws As Object) As Double
    On Error Resume Next
    Dim x As Double
    x = 24#
    Dim r As Long
    r = modUISheet.BlockRow(UG_ANCHOR)
    If r > 0 Then x = modUISheet.CellLeft(ws, r, 2)
    If x < 8# Then x = 8#
    CardLeft = x
End Function

Private Function CardTop(ByVal ws As Object) As Double
    On Error Resume Next
    Dim y As Double
    y = 24#
    Dim r As Long
    r = modUISheet.BlockRow(UG_ANCHOR)
    If r > 0 Then y = ws.Cells(r, 1).Top + 24#
    If y < 8# Then y = 8#
    CardTop = y
End Function

' 文言(専門用語を使わず、押す場所を番号と色で示す。docs/26 のトーン)。
' 裁定書17 H3(b): 実機でカード②の文字が見切れたため、3枚とも**元の半分以下**へ
'   短くした(読ませる量を減らすほど見切れは起きにくい)。字数の目安は H6 の
'   統一表現をそのまま使う。
Private Function TitleOf(ByVal n As Long) As String
    Select Case n
    Case 1
        TitleOf = "① 調べる → ② 貼る"
    Case 2
        TitleOf = "③ まとめて作る"
    Case Else
        TitleOf = "④⑤ 出す"
    End Select
End Function

Private Function BodyOf(ByVal n As Long) As String
    Select Case n
    Case 1
        BodyOf = _
            "[① 調べる指示文を出す]の文を1本ずつ投げ、返答を貼ります。" & vbLf & _
            "1つの欄は約32,000字(A4で約20枚)まで。超えたら「続き1」へ。"
    Case 2
        BodyOf = _
            "③ まとめて作る を押して待ちます。画面が白くなっても処理は続いています。" & vbLf & _
            "終わるとS1～S4に下書きが入ります。読んで直すのが人の仕事。"
    Case Else
        BodyOf = _
            "[④ レポートを出す]でお客様に見せるレポートが出ます。" & vbLf & _
            "[⑤ ヒアリングシートを出す]で聞くことが紙1枚で出ます。" & vbLf & _
            "困ったときは「操作ガイド」タブへ。"
    End Select
End Function

Private Function NextLabelOf(ByVal n As Long) As String
    If n >= UG_STEPS Then
        NextLabelOf = "はじめる"
    Else
        NextLabelOf = "次へ"
    End If
End Function

' カード上の文字(枠なし・塗りなしのテキスト図形)。
Private Sub PutText(ByVal ws As Object, ByVal shapeKey As String, ByVal bodyText As String, _
                    ByVal leftPt As Double, ByVal topPt As Double, _
                    ByVal widthPt As Double, ByVal heightPt As Double, _
                    ByVal fontSize As Double, ByVal boldText As Boolean, _
                    ByVal alignMode As Long, ByVal autoFit As Boolean)
    On Error Resume Next

    Dim shp As Object
    Set shp = ws.Shapes.AddTextbox(UG_SHAPE_TEXTBOX, leftPt, topPt, widthPt, heightPt)
    If shp Is Nothing Then Exit Sub

    shp.Name = shapeKey
    shp.Placement = UG_PLACEMENT_FREE
    shp.Line.Visible = False
    shp.Fill.Visible = False
    shp.TextFrame.Characters.Text = bodyText
    shp.TextFrame.HorizontalAlignment = alignMode
    shp.TextFrame.Characters.Font.Name = UG_FONT
    shp.TextFrame.Characters.Font.Size = fontSize
    shp.TextFrame.Characters.Font.Bold = boldText
    shp.TextFrame.Characters.Font.Color = UG_COLOR_TEXT

    ' 折返しを効かせたうえで高さだけを文字量に合わせる(幅は固定のまま)。
    ' AutoSizeが効かない環境では初期高さのまま残る(下限として働く)。
    If autoFit Then
        shp.TextFrame2.WordWrap = UG_WRAP_ON
        shp.TextFrame2.AutoSize = UG_AUTOSIZE_ON
    End If
End Sub

' 図形の高さ(取れなければ0)。カードの高さを本文の実測で決めるための読み口。
Private Function ShapeHeight(ByVal ws As Object, ByVal shapeKey As String) As Double
    On Error Resume Next
    ShapeHeight = ws.Shapes(shapeKey).Height
End Function

' カード上のボタン(位置をpt直指定するため modUISheet.EnsureButtonEx は使わない)。
Private Sub PutButton(ByVal ws As Object, ByVal shapeKey As String, ByVal caption As String, _
                      ByVal leftPt As Double, ByVal topPt As Double, ByVal widthPt As Double, _
                      ByVal onActionName As String, ByVal primaryKind As Boolean)
    On Error Resume Next

    Dim shp As Object
    Set shp = ws.Shapes.AddShape(UG_SHAPE_ROUNDED, leftPt, topPt, widthPt, 26#)
    If shp Is Nothing Then Exit Sub

    shp.Name = shapeKey
    shp.Placement = UG_PLACEMENT_FREE
    shp.Adjustments(1) = 0.35
    shp.Line.Visible = True
    shp.Line.ForeColor.RGB = UG_COLOR_ACCENT
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.HorizontalAlignment = UG_ALIGN_CENTER
    shp.TextFrame.Characters.Font.Name = UG_FONT
    shp.TextFrame.Characters.Font.Size = 10#
    If primaryKind Then
        shp.Fill.ForeColor.RGB = UG_COLOR_ACCENT
        shp.TextFrame.Characters.Font.Color = UG_COLOR_SURFACE
        shp.TextFrame.Characters.Font.Bold = True
    Else
        shp.Fill.ForeColor.RGB = UG_COLOR_SURFACE
        shp.TextFrame.Characters.Font.Color = UG_COLOR_TEXT
    End If
    shp.OnAction = onActionName
End Sub

' ============================================================================
' ナレッジブックが見つからないときの案内(裁定書14 追補3)
' ----------------------------------------------------------------------------
' 実機では config `kb_path` の先にファイルが無いと、Excelの素のダイアログ
' 「(ファイル名)を開くことができませんでした。」が出ていた。開く側は
' 開く前に存在を確かめて素のダイアログを止めるようにし、利用者への案内は
' ui層の本Subが1回だけ出す(HOMEのナレッジ読込状態と E0401 の記録は従来どおり)。
' 起動の1回だけ呼ばれる経路(StartTourIfFirstRun)に置き、テストからは呼ばない。
' ============================================================================
Private Sub NoticeKbMissingOnce()
    On Error Resume Next

    Dim pathText As String
    pathText = Trim$(modConfig.GetStr("kb_path", vbNullString))
    If LenB(pathText) = 0 Then Exit Sub
    If InStr(1, pathText, "://", vbBinaryCompare) > 0 Then Exit Sub
    If LenB(Dir$(pathText)) > 0 Then Exit Sub

    MsgBox "ナレッジブックが見つかりません。管理者にご連絡ください" & _
           "（操作ガイドの④ 困ったとき をご覧ください）。", vbExclamation, _
           "リスク提案ナビ"
End Sub
