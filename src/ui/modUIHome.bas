Attribute VB_Name = "modUIHome"
Option Explicit

' ============================================================================
' modUIHome - HOMEの描画・状態表示・図形ボタンの配置表(ui層・T-30)
' ----------------------------------------------------------------------------
' 11章 HOME のワイヤーと 13章§2.10 の名前付きレンジが正。ボタンは図形＋OnAction
' (11章§5)で、**OnActionで配線される全ハンドラは先頭で
' modUIProgress.TryEnterUiLock を通す**(16章 E-11。多重クリック・案件切替の関所)。
' 取得できなかったときはダイアログを出さず進捗欄に案内するだけにする(E-50(c))。
'
' 30,000字契約(12章§2)による分割(17章§7 Z-13)。本モジュールはHOMEの**画面**
' だけを持ち、OnActionで配線される**ボタンハンドラ群は modUIHome2** にある
' (移設のみで挙動は変えていない)。ハンドラ側が使う画面の口は本モジュールの
' 公開関数(RefreshHome / SelectedCaseId / ShowWarning / DrawAllSteps /
' WriteRoundNo)で渡し、名前付きレンジ名の定数は2箇所に持たない(14章§6へ登記)。
'
' S1～S4シートの[SNから再実行][SN+1へ進む]は modUIHome2 のハンドラを指す
' (同じ動作の実装を2箇所に置かない。配線は modUICase2.EnsureStepButtons)。
' ============================================================================

Private Const UH_SRC As String = "modUIHome"
Private Const UH_SHEET As String = "HOME"

' 13章§2.10 の名前付きレンジ。
Private Const UH_CASE_ID As String = "hm_case_id"
Private Const UH_CASE_LABEL As String = "hm_case_label"
Private Const UH_STATUS As String = "hm_status"
Private Const UH_ROUND As String = "hm_round_no"
Private Const UH_KB As String = "hm_kb_status"
Private Const UH_BANNER As String = "hm_transport_banner"
Private Const UH_WARNING As String = "hm_warning"
Private Const UH_INBOX_UND As String = "hm_inbox_undiagnosed"
Private Const UH_INBOX_DIA As String = "hm_inbox_diagnosed"
Private Const UH_INBOX_HOLD As String = "hm_inbox_hold"

' 16章 E-27: transport が ribbon 以外のときHOMEに常時出す赤帯。
Private Const UH_BANNER_TEXT As String = "デモ／開発経路で動作中です（本番の判断に使わないでください）"

' 画面制御用のモジュール変数(裁定書9 B1)。**永続でない画面制御変数**であり、
' 14章§6の「状態保持の例外」への登録は要らない。直前に RefreshHome が描いた
' 対象案件IDを覚えておき、変わっていたらS1～S4を描き直す(前の案件の画面が残った
' まま次の案件の sN_edited として確定するクロス案件汚染を、描画側でも塞ぐ)。
Private gShownCaseId As String

' ボタンの間隔と探索範囲。
Private Const UH_BTN_GAP As Double = 8#
Private Const UH_BTN_COL_FIRST As Long = 4
Private Const UH_BTN_COL_LAST As Long = 60

' 1行ぶんの並び。1件 = "図形名;キャプション;OnAction;幅pt" を vbLf 区切り。
' 主要動線は**5本**を2列×3行(1段目①②・2段目③④・3段目⑤は左)で置く。幅200pt・
' 高さ30pt(高さは kind="primary" が modUISheet 側で決める)。横一列の4本は実機で
' 見切れた(裁定書17 H3(a)＋司令塔追補: W6のボタン名を先取りし2度変えない)。
Private Const UH_ROW_MAIN1 As String = _
    "btn_hm_step1;① 調べる指示文を出す;modUIToast.ShowResearchPrompts;200" & vbLf & _
    "btn_hm_step2;② 案件を作って貼る;modUIHome2.HomeNewCase;200"
Private Const UH_ROW_MAIN2 As String = _
    "btn_hm_step3;③ まとめて作る;modUIHome2.HomeRunAll;200" & vbLf & _
    "btn_hm_step4;④ レポートを出す;modUIHome2.HomeExportHtml;200"
Private Const UH_ROW_MAIN3 As String = _
    "btn_hm_step5;⑤ ヒアリングシートを出す;modUIHome2.HomeBuildHearing;200"
Private Const UH_ROW_SUB1 As String = _
    "btn_hm_caseinput;案件入力を開く;modUIHome2.HomeOpenCaseInput;112" & vbLf & _
    "btn_hm_cfopen;企業ファイルを開く;modUIHome2.HomeCompanyOpen;124" & vbLf & _
    "btn_hm_cfsave;企業ファイルへ保存;modUIHome2.HomeCompanySave;124" & vbLf & _
    "btn_round_freeze;第2ラウンド開始;modUIHome2.HomeFreezeRound;116"
Private Const UH_ROW_SUB2 As String = _
    "btn_hm_s1;S1;modUIHome2.HomeRunS1;44" & vbLf & _
    "btn_hm_s2;S2;modUIHome2.HomeRunS2;44" & vbLf & _
    "btn_hm_s3;S3;modUIHome2.HomeRunS3;44" & vbLf & _
    "btn_hm_s4;S4;modUIHome2.HomeRunS4;44" & vbLf & _
    "btn_hm_pf;プリフライト診断;modUIHome2.HomePreflightAll;116"
Private Const UH_ROW_SUB3 As String = _
    "btn_hm_inbox;受信箱を開く;modUIHome2.HomeOpenInbox;100" & vbLf & _
    "btn_hm_fb;商談の記録;modUIHome2.HomeOpenFeedback;92" & vbLf & _
    "btn_hm_judge;判断台帳;modUIHome2.HomeOpenJudgeLog;92" & vbLf & _
    "btn_hm_sparring;壁打ち;modUIHome2.HomeOpenSparring;80"
Private Const UH_ROW_SUB4 As String = _
    "btn_hm_kb;ナレッジ再読込;modUIHome2.HomeReloadKnowledge;108"


' ============================================================================
' 起動時の画面用意(11章§5の図形ボタンとenum入力規則)。modBoot から呼ぶ。
' ============================================================================
Public Sub EnsureScreens()
    On Error Resume Next

    EnsureHomeButtons
    modUICase2.EnsureStepButtons
    modUICase3.EnsureCaseButtons
    modUICase4.EnsureRecordButtons
    modUIInbox.EnsureInboxButtons
    modUISparring.EnsureSparringButtons
    ' 操作ガイドの[テストを実行][ツアーをもう一度見る](裁定書14 裁定5/6)
    modUIGuide.EnsureGuideButtons

    RefreshHome
End Sub

' ============================================================================
' HOMEの図形ボタン(裁定書14 裁定7＋追補1・裁定書17 H3(a): 番号つき動線)
' ----------------------------------------------------------------------------
' 上段に主要動線**5本**(①調べる指示文を出す/②案件を作って貼る/③まとめて作る/
' ④レポートを出す/⑤ヒアリングシートを出す)を**2列×3行**で置き、残り14本は
' 「くわしい操作」区画へ縦に並べる。横一列の4本は実機で右端が画面外へ出た
' (見切れ)ため折り返す(裁定書17 H3(a)＋司令塔追補)。
'
' 重なりを構造的に起こさない置き方(実機でボタンが重なった件=追補1):
'   縦 = 1行に置くのは1段ぶんだけにし、行高は modUISheet.EnsureButtonEx が
'        ボタン高＋余白まで広げる(隣接行のボタンと重ならない)。
'   横 = 列アンカーの**実測の左端**(modUISheet.CellLeft)を読み、直前のボタンの
'        右端＋UH_BTN_GAP より右にある最初の列だけをアンカーにする(FitCol)。
'        列幅を仮定しないので、帳票の列幅を変えても重ならない。
' セル番地はコードに書かない(13章§2.9)。行の基準は名前付きレンジから引く。
' ============================================================================

Private Sub EnsureHomeButtons()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UH_SHEET)
    If ws Is Nothing Then Exit Sub

    ' 旧版の図形名(btn_hm_newcase 等)が残っていても消えるように、置き直す前に
    ' btn_hm_ の図形をまとめて落とす(この後で全部作り直す)。
    modUISheet.DropShapesByPrefix ws, "btn_hm_"

    ' 上段: 主要動線5本(①→⑤の順に押す)を2列×3行で置く。
    Dim mainRow As Long
    mainRow = RowOfNamed(UH_CASE_ID)
    PlaceButtonRow ws, mainRow, UH_ROW_MAIN1, "primary"
    PlaceButtonRow ws, mainRow + 1, UH_ROW_MAIN2, "primary"
    PlaceButtonRow ws, mainRow + 2, UH_ROW_MAIN3, "primary"

    ' 下段: くわしい操作。受信箱の件数欄の下を起点に1行ずつ下へ並べる。
    Dim r As Long
    r = RowOfNamed(UH_INBOX_HOLD) + 2
    modUISheet.EnsureLabel ws, "lbl_hm_more", "くわしい操作", r, UH_BTN_COL_FIRST, 160#, 0#
    PlaceButtonRow ws, r + 1, UH_ROW_SUB1, "plain"
    PlaceButtonRow ws, r + 2, UH_ROW_SUB2, "plain"
    PlaceButtonRow ws, r + 3, UH_ROW_SUB3, "plain"
    PlaceButtonRow ws, r + 4, UH_ROW_SUB4, "plain"
End Sub

' 1行ぶんのボタンを左から右へ、重ならない位置に置く。
Private Sub PlaceButtonRow(ByVal ws As Object, ByVal rowNo As Long, _
                           ByVal specText As String, ByVal kind As String)
    If rowNo <= 0 Then Exit Sub

    Dim specs() As String
    Dim flds() As String
    Dim i As Long
    Dim col As Long
    Dim widthPt As Double
    Dim minLeft As Double

    specs = Split(specText, vbLf)
    col = UH_BTN_COL_FIRST
    minLeft = modUISheet.CellLeft(ws, rowNo, UH_BTN_COL_FIRST)
    If minLeft < 0 Then Exit Sub

    For i = LBound(specs) To UBound(specs)
        flds = Split(specs(i), ";")
        If UBound(flds) - LBound(flds) >= 3 Then
            widthPt = Val(flds(3))
            col = FitCol(ws, rowNo, col, minLeft)
            If col <= 0 Then Exit Sub
            modUISheet.EnsureButtonEx ws, flds(0), flds(1), rowNo, col, widthPt, _
                                      flds(2), kind
            minLeft = modUISheet.CellLeft(ws, rowNo, col) + widthPt + UH_BTN_GAP
        End If
    Next i
End Sub

' 左端が minLeft 以上になる最初のアンカー列(0=範囲内に無い)。
'   直前のボタンの右端＋余白を minLeft に渡すことで、横の重なりが起きない。
Private Function FitCol(ByVal ws As Object, ByVal rowNo As Long, _
                        ByVal fromCol As Long, ByVal minLeft As Double) As Long
    Dim c As Long
    Dim x As Double
    For c = fromCol To UH_BTN_COL_LAST
        x = modUISheet.CellLeft(ws, rowNo, c)
        If x < 0 Then Exit Function
        If x >= minLeft Then
            FitCol = c
            Exit Function
        End If
    Next c
End Function

' 名前付きレンジの行番号(不在なら1)。
Private Function RowOfNamed(ByVal rangeName As String) As Long
    Dim cell As Object
    Set cell = modUISheet.NamedCell(rangeName)
    If cell Is Nothing Then
        RowOfNamed = 1
        Exit Function
    End If
    On Error GoTo One1
    RowOfNamed = cell.row
    Exit Function
One1:
    RowOfNamed = 1
End Function

' ============================================================================
' RefreshHome - 案件の表示・ナレッジ状態・赤帯・受信箱件数(13章§2.10)
' ============================================================================
Public Sub RefreshHome()
    On Error Resume Next

    ' 16章 E-27: ribbon 以外は常時赤帯。ribbon のときは空へ戻す。
    Dim route As String
    route = modGatewayRPN.ResolveTransport(modConfig.GetBool("mock_llm", False), _
                                           modConfig.GetStr("llm_transport", "ribbon"))
    If route = "ribbon" Then
        modUISheet.WriteNamed UH_BANNER, vbNullString
        modUISheet.MarkNamed UH_BANNER, False
    Else
        modUISheet.WriteNamed UH_BANNER, UH_BANNER_TEXT & "（transport=" & route & "）"
        modUISheet.MarkNamed UH_BANNER, True
    End If

    RefreshInboxCounts

    Dim caseId As String
    caseId = SelectedCaseId()

    ' 裁定書9 B1(13章§2.12): hm_case_id が前回の描画時から変わっていたら、
    ' S1～S4を必ず描き直す。前の案件の画面が残ったまま[S2]等を押すと、その
    ' 画面が今の案件の sN_edited として確定してしまう(クロス案件汚染)。
    If StrComp(caseId, gShownCaseId, vbBinaryCompare) <> 0 Then
        gShownCaseId = caseId
        DrawAllSteps caseId
    End If

    If LenB(caseId) = 0 Then
        modUISheet.WriteNamed UH_CASE_LABEL, vbNullString
        modUISheet.WriteNamed UH_STATUS, vbNullString
        modUISheet.WriteNamed UH_ROUND, vbNullString
        modUISheet.WriteNamed UH_KB, KbStatusText("00")
        Exit Sub
    End If

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String
    Dim s4Variant As String
    Dim tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        modUISheet.WriteNamed UH_CASE_LABEL, vbNullString
        Exit Sub
    End If

    modUISheet.WriteNamed UH_CASE_LABEL, caseId & " " & ctx.company & _
        "（" & modUICase.EnumJa("case_type", ctx.case_type) & "）"
    modUISheet.WriteNamed UH_STATUS, _
        modUICase.EnumJa("case_status", modUICase3.CaseCellText(caseId, "status"))
    modUISheet.WriteNamed UH_ROUND, CStr(roundNo)
    modUISheet.WriteNamed UH_KB, KbStatusText(ctx.industry_code)
End Sub

' 受信箱の件数(13章§2.10 hm_inbox_*)。受信箱シートは modUIInbox が数える。
Private Sub RefreshInboxCounts()
    On Error Resume Next
    modUISheet.WriteNamed UH_INBOX_UND, CStr(modUIInbox.CountByStatus("undiagnosed"))
    modUISheet.WriteNamed UH_INBOX_DIA, CStr(modUIInbox.CountByStatus("diagnosed"))
    modUISheet.WriteNamed UH_INBOX_HOLD, CStr(modUIInbox.CountByStatus("conditional_hold"))
End Sub

' ナレッジ読込状態(13章§2.10)。注入テキストの行数から件数を出す
' (modKnowledge は件数の公開口を持たないため、実際に注入される行数で表す)。
Private Function KbStatusText(ByVal industryCode As String) As String
    On Error GoTo Unknown1

    Dim menus As Long
    Dim schemes As Long
    Dim cases1 As Long
    menus = modPipeline.KbRowCount(modKnowledge.MenusFor(industryCode, 0))
    schemes = modPipeline.KbRowCount(modKnowledge.SchemesFor(industryCode, 0))
    cases1 = modPipeline.KbRowCount(modKnowledge.CasesFor(industryCode, 0))

    If menus + schemes + cases1 <= 0 Then
        KbStatusText = "ナレッジブック.xlsx を本体と同じフォルダに置いて" & _
                       "[ナレッジ再読込]を押してください（探した場所: " & _
                       modConfig.GetStr("kb_path", vbNullString) & "）。" & _
                       "置いてあるのに読めないときは、そのファイルを右クリック→" & _
                       "プロパティ→[許可する]にチェックを入れて開き直してください。"
        Exit Function
    End If

    KbStatusText = modBoot.KbAutoNote() & "読込OK（メニュー" & CStr(menus) & _
                   "/型" & CStr(schemes) & "/事例" & CStr(cases1) & "）"
    Exit Function
Unknown1:
    KbStatusText = "ナレッジの状態を取得できませんでした"
End Function

' 選択中の案件ID(13章§2.10 hm_case_id)。書式が不正なら ""。
Public Function SelectedCaseId() As String
    Dim v As String
    v = modUISheet.ReadNamed(UH_CASE_ID)
    If Not modCaseStore.IsValidCaseId(v) Then Exit Function
    SelectedCaseId = v
End Function

' 警告欄への表示(13章§2.10 hm_warning)。空文字で消す。
' 裁定書17 H2/H4: 同じ文言をトーストでも出し(セルは見られていなかった)、
'   失敗系(既定 kind="error")だけ末尾に err_log の送り方を足す。文言の加工と
'   トーストの実装は modUIToast が持つ(本モジュールに残量が無いため)。
'   kind: "error"=失敗(既定・err_log案内あり) / "warn"=注意 / "info"=成功案内。
Public Sub ShowWarning(ByVal messageText As String, _
                       Optional ByVal kind As String = "error")
    modUISheet.WriteNamed UH_WARNING, modUIToast.WarnLine(messageText, kind)
    modUIToast.ShowToast messageText, kind
End Sub

' 一括実行後の4シート描画。
Public Sub DrawAllSteps(ByVal caseId As String)
    Dim n As Long
    For n = 1 To 4
        modUICase2.DrawStep caseId, n
    Next n
End Sub

' 第2ラウンド開始(modUIHome2.HomeFreezeRound)が新しい round_no を書く口。
'   名前付きレンジ名(hm_round_no)の定数を2モジュールに持たないための1本
'   (14章§6へ登記。値の決定は modCaseStore.FreezeRound が唯一持つ)。
Public Sub WriteRoundNo(ByVal roundNo As Long)
    modUISheet.WriteNamed UH_ROUND, CStr(roundNo)
End Sub
