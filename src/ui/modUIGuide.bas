Attribute VB_Name = "modUIGuide"
Option Explicit

' ============================================================================
' modUIGuide - 初回ガイドツアーと使い方タブのボタン(ui層・裁定書14 裁定6)
' ----------------------------------------------------------------------------
' 役割:
'   はじめてブックを開いた人に「①調べる→②貼る→③まとめて作る→④⑤出す」の3枚の
'   カードを順番に見せる。完了・スキップは config `guide_tour_done`="1" として
'   残し、以後は出さない。使い方シートの[ツアーをもう一度見る]から何度でも
'   見直せる。あわせて使い方シートの図形ボタン([テストを実行]・[記録を見る]・
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

Private Const UG_SHEET As String = "ナビ"
Private Const UG_GUIDE_SHEET As String = "使い方"
Private Const UG_PREFIX As String = "gt_"
Private Const UG_FLAG As String = "guide_tour_done"
Private Const UG_STEPS As Long = 4
Private Const UG_LOCK_NAME As String = "はじめの案内"
' 13章§2.3 kb_path の既定値に入る「配置前のプレースホルダ」の目印(modBoot の
' BOOT_KB_PLACEHOLDER と同値。W9.2)。この形は Dir$ に掛けない。
Private Const UG_KB_PLACEHOLDER As String = "\\...\"

' 13章§2.10 の名前付きレンジ(カードの位置の基準。セル番地は書かない)。
Private Const UG_ANCHOR As String = "nv_sec1"
' 使い方タブのボタンアンカー(13章§2.18)。
Private Const UG_BTN_TEST As String = "gd_btn_test"
Private Const UG_BTN_TOUR As String = "gd_btn_tour"
' v3.2(11章§3.6): ⑤困ったときの[記録を見る]と⑦上級の[表示する]5本。
Private Const UG_BTN_LOGS As String = "gd_btn_logs"
Private Const UG_BTN_ADV As String = "gd_btn_adv"
' 裁定書22: ⑦上級の**動作**ボタン3本のアンカー(13章§2.18)。1件 =
'   "図形名;キャプション;OnAction;幅pt" を vbLf 区切り。並びは build_rpn.py の
'   GUIDE_ADVANCED_ACTIONS と同順であり、tools/caption_check.py が逐語照合する。
Private Const UG_BTN_ADV_ACT As String = "gd_btn_adv_act"
' フッター(裁定書26 D)。⑦上級の動作ボタン3行の**2行下**へ置く。
Private Const UG_BTN_FOOTER As String = "btn_gd_footer"
Private Const UG_ADV_ACT_ROWS As Long = 3
Private Const UG_FOOTER_GAP_ROWS As Long = 2
Private Const UG_ROW_ADV_ACT As String = _
    "btn_gd_round2;第2ラウンドを始める;modUIHome2.HomeFreezeRound;180" & vbLf & _
    "btn_gd_cfsave;企業ファイルへ保存;modUIHome2.HomeCompanySave;180" & vbLf & _
    "btn_gd_cfopen;企業ファイルを開く;modUIHome2.HomeCompanyOpen;180"

' 記録3枚(13章§2.9)。まとめて可視にし err_log へ移る。
Private Const UG_LOG_SHEETS As String = "err_log" & vbLf & "run_log" & vbLf & "usage_log"
' 上級5枚(11章§3.6⑦)。押した1枚だけを可視にする。並び順は使い方タブと同じ。
Private Const UG_ADV_SHEETS As String = "壁打ち" & vbLf & "受信箱" & vbLf & _
    "フィードバック" & vbLf & "判断台帳" & vbLf & "案件一覧"

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
' 上級の動作ボタン3本の配置表の読み出し口(caption_check の値源)。
Public Function AdvActionRow() As String
    AdvActionRow = UG_ROW_ADV_ACT
End Function

' ============================================================================
Public Sub StartTourIfFirstRun()
    On Error Resume Next

    NoticeKbMissingOnce

    If TourDone() Then Exit Sub
    ClearTour
    DrawCard 1
End Sub

' ============================================================================
' RestartTour - 使い方タブの[ツアーをもう一度見る]の OnAction。
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
' EnsureGuideButtons - 使い方シートの図形ボタン(modUIHome.EnsureScreens から
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

    ' ⑤困ったときの[記録を見る](11章§4.4・13章§2.9)。
    r = modUISheet.BlockRow(UG_BTN_LOGS)
    c = modUISheet.BlockCol(UG_BTN_LOGS)
    If r > 0 And c > 0 Then
        modUISheet.EnsureButtonEx ws, "btn_gd_logs", "記録を見る", r, c, 132#, _
                                  "modUIGuide.ShowLogs", "plain"
    End If

    ' ⑦上級の[表示する]5本。config `ui_advanced` が FALSE のときは**作らない**
    '   (描いて隠すのではない。11章§7.2(c)。孤児図形と誤クリックの両方を消す)。
    If Not modConfig.GetBool("ui_advanced", True) Then Exit Sub
    r = modUISheet.BlockRow(UG_BTN_ADV)
    c = modUISheet.BlockCol(UG_BTN_ADV)
    If r <= 0 Or c <= 0 Then Exit Sub

    Dim names() As String
    names = Split(UG_ADV_SHEETS, vbLf)
    Dim i As Long
    For i = LBound(names) To UBound(names)
        ' 個別の機能フラグ(11章。feature_inbox/feature_judgelogがFALSEの間は
        '   その1行だけ図形を作らない。ui_advancedとは別に、その機能だけ隠せる)。
        If (names(i) = "受信箱" And Not modConfig.GetBool("feature_inbox", True)) Or _
           (names(i) = "判断台帳" And Not modConfig.GetBool("feature_judgelog", True)) Then
            ' skip: 導線を作らない
        Else
            modUISheet.EnsureButtonEx ws, "btn_gd_adv" & CStr(i + 1), "表示する", _
                                      r + i, c, 100#, _
                                      "modUIGuide.ShowAdvanced" & CStr(i + 1), "plain"
        End If
    Next i

    ' ⑦上級の動作ボタン3本(裁定書22)。ui_advanced が FALSE のときは上で
    '   Exit Sub 済みなので、ここも作られない(導線をまとめて消す)。
    r = modUISheet.BlockRow(UG_BTN_ADV_ACT)
    c = modUISheet.BlockCol(UG_BTN_ADV_ACT)
    If r <= 0 Or c <= 0 Then Exit Sub

    Dim specs() As String
    specs = Split(UG_ROW_ADV_ACT, vbLf)
    Dim k As Long
    For k = LBound(specs) To UBound(specs)
        Dim f() As String
        f = Split(specs(k), ";")
        If UBound(f) - LBound(f) >= 3 Then
            modUISheet.EnsureButtonEx ws, f(0), f(1), r + k, c, Val(f(3)), f(2), "plain"
        End If
    Next k
End Sub

' ============================================================================
' EnsureFooterButton - 使い方タブの最下部のフッター(裁定書26 D)。
'   最終ブロック(⑦上級の動作ボタン3行)の2行下へ、ナビと同じ1本を置く。
'   キャプションとOnActionの値源は modUINav.NavRowFooter()(2箇所で別の名前を
'   持たない)。図形名だけ使い方タブ用に btn_gd_footer とする。
' ============================================================================
Public Sub EnsureFooterButton()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UG_GUIDE_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    Dim c As Long
    r = modUISheet.BlockRow(UG_BTN_ADV_ACT)
    c = modUISheet.BlockCol(UG_BTN_ADV_ACT)
    If r <= 0 Or c <= 0 Then Exit Sub

    Dim flds() As String
    flds = Split(modUINav.NavRowFooter(), ";")
    If UBound(flds) - LBound(flds) < 3 Then Exit Sub

    modUISheet.EnsureFooterButton ws, UG_BTN_FOOTER, flds(1), _
                                  r + UG_ADV_ACT_ROWS + UG_FOOTER_GAP_ROWS, c, _
                                  Val(flds(3)), flds(2)
End Sub

' ============================================================================
' OpenGuide - ナビの帯の[使い方を開く](13章§2.10(f))。
' ============================================================================
Public Sub OpenGuide()
    If Not modUIProgress.TryEnterUiLock("使い方を開く") Then Exit Sub
    On Error GoTo Done
    modUISheet.ShowSheet UG_GUIDE_SHEET
    modUIToast.ShowToast "「使い方」を開きました。上の目次から、知りたいところへ飛べます。", "info"
Done:
    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' ShowLogs - 使い方タブ⑤の[記録を見る](13章§2.9)。3枚まとめて可視にし
'   err_log へ移る。**一度出したシートは、そのブックでは出したままにする**。
' ============================================================================
Public Sub ShowLogs()
    If Not modUIProgress.TryEnterUiLock("記録の表示") Then Exit Sub
    On Error GoTo Done

    Dim names() As String
    names = Split(UG_LOG_SHEETS, vbLf)
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim ws As Object
        Set ws = modUISheet.SheetOf(names(i))
        If Not ws Is Nothing Then
            ws.Visible = -1                       ' xlSheetVisible
            modUISheet.EnsureBackButton ws
        End If
    Next i

    modUISheet.ShowSheet names(LBound(names))
    modUIToast.ShowToast "記録のタブを3枚出しました。err_log のいちばん下の行を" & _
                         "コピーして、開発担当へ送ってください。", "info"
Done:
    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' ShowAdvanced - 使い方タブ⑦の[表示する](11章§3.6⑦)。押した**1枚だけ**を
'   可視にする。上級側の公開Subは flag に関係なく動く(flag は導線の有無だけ)。
' ============================================================================
Public Sub ShowAdvanced(ByVal sheetKey As String)
    On Error GoTo Failed

    Dim ws As Object
    Set ws = modUISheet.SheetOf(sheetKey)
    If ws Is Nothing Then
        modUIToast.ShowToast "そのタブを出せませんでした。使い方タブの[記録を見る]を" & _
                             "押して、いちばん下の行を開発担当へ送ってください。", "error"
        Exit Sub
    End If
    ws.Visible = -1                               ' xlSheetVisible
    modUISheet.EnsureBackButton ws
    modUISheet.ShowSheet sheetKey
    modUIToast.ShowToast "「" & sheetKey & "」を出しました。" & _
                         "ナビへ戻るときは1行目の[ナビへ戻る]を押してください。", "info"
    Exit Sub
Failed:
    modLog.LogError "E0603", "modUIGuide.ShowAdvanced", "show_failed:" & sheetKey, Err.Number
End Sub

' 上級の[表示する]5本(OnAction の口。実体は共通の ShowAdvanced)。
Private Function AdvSheetOf(ByVal n As Long) As String
    Dim names() As String
    names = Split(UG_ADV_SHEETS, vbLf)
    If n < 1 Or n > UBound(names) - LBound(names) + 1 Then Exit Function
    AdvSheetOf = names(LBound(names) + n - 1)
End Function

Public Sub ShowAdvanced1()
    If Not modUIProgress.TryEnterUiLock("上級タブの表示") Then Exit Sub
    ShowAdvanced AdvSheetOf(1)
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAdvanced2()
    If Not modUIProgress.TryEnterUiLock("上級タブの表示") Then Exit Sub
    ShowAdvanced AdvSheetOf(2)
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAdvanced3()
    If Not modUIProgress.TryEnterUiLock("上級タブの表示") Then Exit Sub
    ShowAdvanced AdvSheetOf(3)
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAdvanced4()
    If Not modUIProgress.TryEnterUiLock("上級タブの表示") Then Exit Sub
    ShowAdvanced AdvSheetOf(4)
    modUIProgress.ExitUiLock
End Sub

Public Sub ShowAdvanced5()
    If Not modUIProgress.TryEnterUiLock("上級タブの表示") Then Exit Sub
    ShowAdvanced AdvSheetOf(5)
    modUIProgress.ExitUiLock
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
        TitleOf = "1 会社のこと"
    Case 2
        TitleOf = "2 貼る"
    Case 3
        TitleOf = "3 作る"
    Case Else
        TitleOf = "4 出す"
    End Select
End Function

Private Function BodyOf(ByVal n As Long) As String
    ' 11章§3.6 末尾: 文面はナビのコーチ帯の6文から STEP 1・3・4・6 を流用し、
    ' **1字も別の文を作らない**(値源は modUINav.StepText の1本だけ)。
    Select Case n
    Case 1
        BodyOf = modUINav.StepText(1)
    Case 2
        BodyOf = modUINav.StepText(3)
    Case 3
        BodyOf = modUINav.StepText(4)
    Case Else
        BodyOf = modUINav.StepText(6)
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
    ' 配置前の既定プレースホルダ(13章§2.3 kb_path の既定値)を Dir$ に掛けない
    ' (W9.2)。"\\...\" は不正なUNCで、Mac の実Excel では Dir$ が実行時エラー5
    ' (プロシージャの呼び出し、または引数が無効です)を投げる。
    If InStr(1, pathText, UG_KB_PLACEHOLDER, vbBinaryCompare) > 0 Then Exit Sub
    If LenB(Dir$(pathText)) > 0 Then Exit Sub

    MsgBox "ナレッジブックが見つかりません。管理者にご連絡ください" & _
           "（操作ガイドの4 困ったとき をご覧ください）。", vbExclamation, _
           "リスク提案ナビ"
End Sub
