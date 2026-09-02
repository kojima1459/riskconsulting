Attribute VB_Name = "modUIHome"
Option Explicit

' ============================================================================
' modUIHome - HOMEの描画・プレイ起動・進捗・赤帯(ui層・T-30)
' ----------------------------------------------------------------------------
' 11章 HOME のワイヤーと 13章§2.10 の名前付きレンジが正。ボタンは図形＋OnAction
' (11章§5)で、**OnActionで配線される全ハンドラは先頭で
' modUIProgress.TryEnterUiLock を通す**(16章 E-11。多重クリック・案件切替の関所)。
' 取得できなかったときはダイアログを出さず進捗欄に案内するだけにする(E-50(c))。
'
' 実行の流れ(1ハンドラ=1アクション):
'   TryEnterUiLock -> ParkFocus(E-51: 編集モードのまま処理へ入らない)
'   -> SetStage(リボン呼出の前に進捗を書き切る。E-50(a))
'   -> app層(modPipeline / modExportHtml / modCompanyFile / modPlayOps)
'   -> 画面更新 -> ExitUiLock(ScreenUpdating の復帰を内包)
'
' S1～S4シートの[SNから再実行][SN+1へ進む]も本モジュールのハンドラを指す
' (同じ動作の実装を2箇所に置かない。配線は modUICase2.EnsureStepButtons)。
' ============================================================================

Private Const UH_SRC As String = "modUIHome"
Private Const UH_SHEET As String = "HOME"

' 13章§2.10 の名前付きレンジ。
Private Const UH_CASE_ID As String = "hm_case_id"
Private Const UH_CASE_LABEL As String = "hm_case_label"
Private Const UH_STATUS As String = "hm_status"
Private Const UH_ROUND As String = "hm_round_no"
Private Const UH_QUALITY As String = "hm_quality_mode"
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

' 18章§1.1(1): S1未実行時のレポート出力は**エラーコードを立てず**中止する。
Private Const UH_MSG_NEED_S1 As String = "先にStep1を実行してください。"
Private Const UH_MSG_NO_CASE As String = "対象案件が選ばれていません（HOMEの対象案件IDをご確認ください）。"

' 企業ドシエファイルの保存先。13章§2.3 に専用キーが無いため、出力の共通
' フォルダ(html_out_dir)を使う(17章の裁定事項として申し送り)。
Private Const UH_DIR_KEY As String = "html_out_dir"
Private Const UH_DIR_DEFAULT As String = "%USERPROFILE%\Documents\RPN出力"

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
    "btn_hm_step2;② 案件を作って貼る;modUIHome.HomeNewCase;200"
Private Const UH_ROW_MAIN2 As String = _
    "btn_hm_step3;③ まとめて作る;modUIHome.HomeRunAll;200" & vbLf & _
    "btn_hm_step4;④ レポートを出す;modUIHome.HomeExportHtml;200"
Private Const UH_ROW_MAIN3 As String = _
    "btn_hm_step5;⑤ ヒアリングシートを出す;modUIHome.HomeBuildHearing;200"
Private Const UH_ROW_SUB1 As String = _
    "btn_hm_caseinput;案件入力を開く;modUIHome.HomeOpenCaseInput;112" & vbLf & _
    "btn_hm_cfopen;企業ファイルを開く;modUIHome.HomeCompanyOpen;124" & vbLf & _
    "btn_hm_cfsave;企業ファイルへ保存;modUIHome.HomeCompanySave;124" & vbLf & _
    "btn_round_freeze;第2ラウンド開始;modUIHome.HomeFreezeRound;116"
Private Const UH_ROW_SUB2 As String = _
    "btn_hm_s1;S1;modUIHome.HomeRunS1;44" & vbLf & _
    "btn_hm_s2;S2;modUIHome.HomeRunS2;44" & vbLf & _
    "btn_hm_s3;S3;modUIHome.HomeRunS3;44" & vbLf & _
    "btn_hm_s4;S4;modUIHome.HomeRunS4;44" & vbLf & _
    "btn_hm_pf;プリフライト診断;modUIHome.HomePreflightAll;116"
Private Const UH_ROW_SUB3 As String = _
    "btn_hm_inbox;受信箱を開く;modUIHome.HomeOpenInbox;100" & vbLf & _
    "btn_hm_fb;商談の記録;modUIHome.HomeOpenFeedback;92" & vbLf & _
    "btn_hm_judge;判断台帳;modUIHome.HomeOpenJudgeLog;92" & vbLf & _
    "btn_hm_sparring;壁打ち;modUIHome.HomeOpenSparring;80"
Private Const UH_ROW_SUB4 As String = _
    "btn_hm_kb;ナレッジ再読込;modUIHome.HomeReloadKnowledge;108"


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
Private Function SelectedCaseId() As String
    Dim v As String
    v = modUISheet.ReadNamed(UH_CASE_ID)
    If Not modCaseStore.IsValidCaseId(v) Then Exit Function
    SelectedCaseId = v
End Function

' HOMEの品質上書き(13章§2.10 hm_quality_mode)。日本語ラベル -> 機械値。
'   空(未選択)は "" のまま返し、config・ティア連動の解決へ委ねる。
Private Function QualityOverride() As String
    QualityOverride = modUICase.EnumEn("quality_mode", modUISheet.ReadNamed(UH_QUALITY))
End Function

' 警告欄への表示(13章§2.10 hm_warning)。空文字で消す。
' 裁定書17 H2/H4: 同じ文言をトーストでも出し(セルは見られていなかった)、
'   失敗系(既定 kind="error")だけ末尾に err_log の送り方を足す。文言の加工と
'   トーストの実装は modUIToast が持つ(本モジュールに残量が無いため)。
'   kind: "error"=失敗(既定・err_log案内あり) / "warn"=注意 / "info"=成功案内。
Private Sub ShowWarning(ByVal messageText As String, _
                        Optional ByVal kind As String = "error")
    modUISheet.WriteNamed UH_WARNING, modUIToast.WarnLine(messageText, kind)
    modUIToast.ShowToast messageText, kind
End Sub

' 16章 E-35/E-36(裁定書9 B9・14章§6 N1): 入念モードで「審査を省略した」
'   「改訂を破棄した」ときの逐語警告をHOMEへ出す。文言の値源は
'   modPipeline2.DeepWarningOf だけであり(ui層は文言を持たない)、直近の結末は
'   modPipeline2.LastDeepOutcome から取る。**実行の冒頭で掛ける
'   ShowWarning vbNullString より後**に呼ぶこと(先に呼ぶと消える)。
Private Sub ShowDeepWarning()
    On Error Resume Next

    Dim outcome As String
    outcome = modPipeline2.LastDeepOutcome()
    If LenB(outcome) = 0 Then Exit Sub

    Dim warnText As String
    warnText = modPipeline2.DeepWarningOf(outcome)
    If LenB(warnText) = 0 Then Exit Sub

    ShowWarning warnText, "warn"
End Sub

' ============================================================================
' OnActionハンドラ(11章 HOMEワイヤー)
' ----------------------------------------------------------------------------
' すべて先頭で TryEnterUiLock を取り、終了時に必ず ExitUiLock を通す。
' ============================================================================

Public Sub HomeRunAll()
    If Not modUIProgress.TryEnterUiLock("一括実行") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    ShowWarning vbNullString
    ' 裁定書10 M1/N9: deep outcome のリセットは**実行の開始時に1回だけ**。
    ' 一括実行の中では消さない(Step2/3 の結末を Step4 が消さないため)。
    modPipeline2.ResetDeepOutcome

    Dim caseId As String
    caseId = SelectedCaseId()
    If LenB(caseId) = 0 Then
        ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    modUIProgress.SetStage "一括実行（Step1～Step4）", WaitSec()
    If modPipeline.RunAll(caseId, QualityOverride()) Then
        DrawAllSteps caseId
        modUIToast.ShowNext 2
        ShowDeepWarning
    Else
        ShowWarning "実行が完了しませんでした。err_log をご確認ください。"
    End If
    RefreshHome

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Public Sub HomeRunS1()
    If Not modUIProgress.TryEnterUiLock("Step1") Then Exit Sub
    RunStepUi 1
End Sub

Public Sub HomeRunS2()
    If Not modUIProgress.TryEnterUiLock("Step2") Then Exit Sub
    RunStepUi 2
End Sub

Public Sub HomeRunS3()
    If Not modUIProgress.TryEnterUiLock("Step3") Then Exit Sub
    RunStepUi 3
End Sub

Public Sub HomeRunS4()
    If Not modUIProgress.TryEnterUiLock("Step4") Then Exit Sub
    RunStepUi 4
End Sub

' 1Stepの実行(ロックは呼び出し元のハンドラが取得済み)。
'   (1) 上流シートの編集を sN-1_edited へ確定(11章「編集は次へ進むと下流に反映」)
'   (2) 下流の成果物があれば確認ダイアログ＋InvalidateDownstream(16章 E-10)
'   (3) SetStage -> modPipeline.RunStep -> 描画
Private Sub RunStepUi(ByVal stepNo As Long)
    On Error GoTo Done

    modUIProgress.ParkFocus
    ShowWarning vbNullString
    ' 裁定書10 M1/N9: deep outcome のリセットは実行の開始時に1回だけ。
    modPipeline2.ResetDeepOutcome

    Dim caseId As String
    caseId = SelectedCaseId()
    If LenB(caseId) = 0 Then
        ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    SaveUpstreamEdits caseId, stepNo
    If Not ConfirmInvalidate(caseId, stepNo) Then GoTo Done

    modUIProgress.SetStage "Step" & CStr(stepNo) & " を実行中", WaitSec()
    If modPipeline.RunStep(caseId, stepNo, QualityOverride()) Then
        modUICase2.DrawStep caseId, stepNo
        modUISheet.ShowSheet modUICase2.SheetNameOf(stepNo)
        ShowDeepWarning
    Else
        ShowWarning "Step" & CStr(stepNo) & " が完了しませんでした。err_log をご確認ください。"
    End If
    RefreshHome

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' 上流Stepのシート編集を確定する(11章「編集はSN+1へ進むと下流に反映」)。
'   13章§2.2 規約4: 不合格なら sN_edited を保存せず、下流は従前の参照優先
'   (sNr_json / sN_json)を使い続ける。**実行そのものは止めない**ので戻り値を
'   持たない(止めるべきかを呼び出し側が判断できるかのような形にしない)。
Private Sub SaveUpstreamEdits(ByVal caseId As String, ByVal stepNo As Long)
    If stepNo <= 1 Then Exit Sub

    Dim upstream As Long
    upstream = stepNo - 1
    If LenB(Trim$(modCaseStore.ResolveStepJson(caseId, upstream))) = 0 Then Exit Sub

    Dim errText As String
    If modUICase2.SaveEditedStep(caseId, upstream, errText) Then Exit Sub

    If LenB(errText) = 0 Then Exit Sub
    ShowWarning "Step" & CStr(upstream) & " の編集内容が検証に通らなかったため、" & _
                "編集前の内容のまま実行します: " & modUtil.SafeLeft(errText, 300), "warn"
End Sub

' 16章 E-10: 下流の成果物があるなら確認してから無効化する。
'   利用者が「いいえ」を選んだら実行しない(戻り値 False)。
Private Function ConfirmInvalidate(ByVal caseId As String, ByVal stepNo As Long) As Boolean
    ConfirmInvalidate = True
    If stepNo >= 4 Then Exit Function

    Dim n As Long
    Dim hasDownstream As Boolean
    For n = stepNo + 1 To 4
        If LenB(Trim$(modCaseStore.LoadData(caseId, "s" & CStr(n) & "_json"))) > 0 Then
            hasDownstream = True
        End If
    Next n
    If Not hasDownstream Then Exit Function

    Dim answer As Long
    answer = MsgBox("Step" & CStr(stepNo) & " を実行すると、Step" & CStr(stepNo + 1) & _
                    "以降の結果は破棄されます。続けますか？", vbYesNo + vbQuestion, _
                    "下流の結果を無効化します")
    If answer <> vbYes Then
        ConfirmInvalidate = False
        Exit Function
    End If

    modCaseStore.InvalidateDownstream caseId, stepNo
End Function

' 一括実行後の4シート描画。
Private Sub DrawAllSteps(ByVal caseId As String)
    Dim n As Long
    For n = 1 To 4
        modUICase2.DrawStep caseId, n
    Next n
End Sub

' 進捗表示に出す最大待ち時間(config llm_wait_sec)。
Private Function WaitSec() As Long
    WaitSec = modConfig.GetLong("llm_wait_sec", 1200)
End Function

' ============================================================================
' 新規案件・画面遷移
' ============================================================================
Public Sub HomeNewCase()
    If Not modUIProgress.TryEnterUiLock("新規案件") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    ShowWarning vbNullString

    ' 裁定書12 V1(13章§2.11): ここでは採番しない。案件入力を**新規モード**で
    ' 開くだけにし、ci_case_id へ固定マーカー「(新規)」を書く。採番は案件入力の
    ' [保存して戻る](modUICase3.CaseSave の3値判定)が企業名・業種を読んで行う。
    ' 空欄のまま採番して幽霊案件が積まれるのを防ぎ、かつ「表示が空のまま保存」
    ' を新規採番へ倒さないという保証を、ブックに残るセル1つで成り立たせる。
    '
    ' 裁定書13 W1(13章§2.11): マーカーを書く**前に画面を全クリアする**。前の案件
    ' を描いた画面のまま新規モードへ入ると、その画面の貼付内容がそのまま新しい
    ' 案件の case_data として確定する(切り詰まった描画のあとでも同じ)。
    ' **新規モードは空画面から始まる**。
    modUICase4.ClearCaseInput
    modUISheet.WriteNamed "ci_case_id", modUICase3.U3_NEW_MARK
    modUISheet.ShowSheet "案件入力"
    modUIToast.ShowNext 1

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Public Sub HomeOpenCaseInput()
    If Not modUIProgress.TryEnterUiLock("案件入力を開く") Then Exit Sub
    On Error GoTo Done
    modUICase3.DrawCaseInput SelectedCaseId()
    modUISheet.ShowSheet "案件入力"
Done:
    modUIProgress.ExitUiLock
End Sub

Public Sub HomeOpenInbox()
    If Not modUIProgress.TryEnterUiLock("受信箱を開く") Then Exit Sub
    On Error GoTo Done
    modUIInbox.RefreshInbox
    modUISheet.ShowSheet "受信箱"
Done:
    modUIProgress.ExitUiLock
End Sub

Public Sub HomeOpenFeedback()
    If Not modUIProgress.TryEnterUiLock("商談の記録") Then Exit Sub
    On Error GoTo Done
    modUISheet.ShowSheet "フィードバック"
Done:
    modUIProgress.ExitUiLock
End Sub

Public Sub HomeOpenJudgeLog()
    If Not modUIProgress.TryEnterUiLock("判断台帳") Then Exit Sub
    On Error GoTo Done
    modUISheet.ShowSheet "判断台帳"
Done:
    modUIProgress.ExitUiLock
End Sub

Public Sub HomeOpenSparring()
    If Not modUIProgress.TryEnterUiLock("壁打ち") Then Exit Sub
    On Error GoTo Done
    modUISparring.OpenSparring SelectedCaseId()
Done:
    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' 第2ラウンド開始(11章 HOMEワイヤー・裁定書9 A-2/N7)
' ----------------------------------------------------------------------------
' 現在のS2を前ラウンド(s2_prev_json)として確定し round_no を+1する。退避の
' 実装は modCaseStore.FreezeRound が唯一持ち(退避できなければ round_no を
' 進めない)、本ハンドラは確認ダイアログと画面の更新だけを行う。
' ============================================================================
Public Sub HomeFreezeRound()
    If Not modUIProgress.TryEnterUiLock("第2ラウンド開始") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    ShowWarning vbNullString

    Dim caseId As String
    caseId = SelectedCaseId()
    If LenB(caseId) = 0 Then
        ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    Dim answer As Long
    answer = MsgBox("現在のS2を前ラウンドとして確定し、第2ラウンドを開始します。" & _
                    "よろしいですか", vbYesNo + vbQuestion, "第2ラウンド開始")
    If answer <> vbYes Then GoTo Done

    Dim newRound As Long
    newRound = modCaseStore.FreezeRound(caseId)
    If newRound <= 0 Then
        ShowWarning "第2ラウンドを開始できませんでした（前ラウンドの退避に失敗しました）。" & _
                    "err_log をご確認ください。"
        GoTo Done
    End If

    modUISheet.WriteNamed UH_ROUND, CStr(newRound)
    RefreshHome
    ShowWarning "第" & CStr(newRound) & "ラウンドを開始しました（前ラウンドのS2を退避しました）。", "info"

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' ============================================================================
' 出力(11章 HOME「出力: [リスクレポートHTML] [ヒアリングシート]」)
' ============================================================================

' 18章§1.1(1): S1が空なら**エラーコードを立てず**案内して中止する。
Public Sub HomeExportHtml()
    If Not modUIProgress.TryEnterUiLock("リスクレポートHTML") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    ShowWarning vbNullString

    Dim caseId As String
    caseId = SelectedCaseId()
    If LenB(caseId) = 0 Then
        ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    ' S1未実行ガード(18章§1.1(1))。データ未整備は障害ではないので E-code を
    ' 立てず、hm_warning で案内して生成そのものを呼ばない。
    If LenB(Trim$(modCaseStore.ResolveStepJson(caseId, 1))) = 0 Then
        ShowWarning UH_MSG_NEED_S1, "warn"
        GoTo Done
    End If

    modUIProgress.SetStage "リスクレポートHTMLを生成中", WaitSec()

    Dim outPath As String
    Dim errText As String
    errText = modExportHtml.GenerateHtmlReport(caseId, outPath)
    If LenB(errText) > 0 Then
        ShowWarning errText
    Else
        ' 13章§2.1・11章§4(裁定書9 B16(b)): HTMLレポートの生成成功が
        ' case_status の exported を立てる唯一の点。SetStatus は遷移検査を
        ' 通さない(14章§6の注記)が、CS_STATUSES_ABOVE_S4 の降格抑止は効く。
        modCaseStore.SetStatus caseId, "exported"
        ShowWarning "リスクレポートを出力しました: " & outPath, "info"
        modUIToast.ShowNext 3
    End If
    RefreshHome

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Public Sub HomeBuildHearing()
    If Not modUIProgress.TryEnterUiLock("ヒアリングシート") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    ShowWarning vbNullString

    Dim caseId As String
    caseId = SelectedCaseId()
    If LenB(caseId) = 0 Then
        ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    ' S4シートの編集をヒアリングシートへ反映する(11章 S4ワイヤー)。
    Dim errText As String
    If LenB(Trim$(modCaseStore.ResolveStepJson(caseId, 4))) > 0 Then
        modUICase2.SaveEditedStep caseId, 4, errText
    End If

    ' 裁定書9 B12(13章§2.16): 生成は answer_memo を含む全行を空へ戻す。
    ' 訪問後に手書きの回答が入っている状態で押されたら、必ず確認を挟む
    ' (VBAの書込は Undo できない。16章 E-10 の下流無効化と同じ作法)。
    ' 裁定書10 M5: AnswerMemoCount は案件を問わず**常に数える**(旧仕様の
    ' 「hs_case_id 不一致なら0」は、別案件の回答が残っている場面でちょうど
    ' 確認を素通りさせる fail-open だった)。別案件かどうかの判定はここで行い、
    ' 文言だけを変える(ヒアリングシートはブックに1枚しかない)。
    Dim memoRows As Long
    memoRows = modExportHearing.AnswerMemoCount(caseId)
    If memoRows > 0 Then
        Dim sheetCase As String
        sheetCase = Trim$(modUISheet.ReadNamed("hs_case_id"))

        Dim memoText As String
        If LenB(sheetCase) > 0 And StrComp(sheetCase, caseId, vbBinaryCompare) <> 0 Then
            memoText = "別案件（" & sheetCase & "）の手書き回答が" & CStr(memoRows) & _
                       "行残っています。作り直すとこの回答は消えます。"
        Else
            memoText = "ヒアリングシートに手書きの回答が" & CStr(memoRows) & _
                       "行あります。作り直すとこの回答は消えます。"
        End If

        Dim answer As Long
        answer = MsgBox(memoText & vbLf & _
                        "続けますか？", vbYesNo + vbExclamation, "回答の上書き確認")
        If answer <> vbYes Then
            ShowWarning "ヒアリングシートは作り直しませんでした（手書きの回答を残しました）。", "warn"
            GoTo Done
        End If
    End If

    If modExportHearing.BuildHearingSheet(caseId) Then
        modUISheet.ShowSheet "ヒアリングシート"
        modUIToast.ShowNext 4
    Else
        ShowWarning "ヒアリングシートを生成できませんでした（先にStep4を実行してください）。"
    End If

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' ============================================================================
' 企業ドシエファイル(13章§2.8・FR-45・16章 E-05(7))
' ============================================================================
Public Sub HomeCompanySave()
    If Not modUIProgress.TryEnterUiLock("企業ファイルへ保存") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    ShowWarning vbNullString

    Dim caseId As String
    caseId = SelectedCaseId()
    If LenB(caseId) = 0 Then
        ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    ' 16章 E-05(7): 検知があれば「確認した」を選ばせてから書き出す。
    Dim confirmedAt As String
    Dim piiText As String
    piiText = modCompanyFile.ScanCaseForPii(caseId)
    If LenB(piiText) > 0 Then
        Dim answer As Long
        answer = MsgBox("個人情報らしき記述を検知しました（" & _
                        modUtil.SafeLeft(piiText, 200) & "）。" & vbLf & _
                        "内容を確認しましたか？「はい」で書き出します。", _
                        vbYesNo + vbExclamation, "共有前の確認")
        If answer <> vbYes Then
            ShowWarning "個人情報の確認が済んでいないため書き出しませんでした。", "warn"
            GoTo Done
        End If
        confirmedAt = modUtil.NowStamp()
    End If

    Dim pathText As String
    pathText = modCompanyFile.ExportCompanyFile(caseId, OutDir(), confirmedAt)
    If LenB(pathText) = 0 Then
        ShowWarning "企業ファイルへ書き出せませんでした。err_log をご確認ください。"
    Else
        ShowWarning "企業ファイルへ保存しました: " & pathText, "info"
    End If

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Public Sub HomeCompanyOpen()
    If Not modUIProgress.TryEnterUiLock("企業ファイルを開く") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    ShowWarning vbNullString

    Dim caseId As String
    caseId = SelectedCaseId()
    If LenB(caseId) = 0 Then
        ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    Dim pathText As String
    pathText = AskCompanyFile(caseId)
    If LenB(pathText) = 0 Then GoTo Done

    If modCompanyFile.ImportCompanyFile(pathText, caseId) Then
        DrawAllSteps caseId
        RefreshHome
        ShowWarning "企業ファイルから前ラウンドの内容を取り込みました。", "info"
    Else
        ShowWarning "企業ファイルを取り込めませんでした。ファイルをご確認ください。"
    End If

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' 既定パスを初期値にしたファイル選択。キャンセルは ""。
Private Function AskCompanyFile(ByVal caseId As String) As String
    On Error GoTo Cancelled

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String
    Dim s4Variant As String
    Dim tierText As String
    Dim initPath As String
    If modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        initPath = modCompanyFile.CompanyFilePath(ctx.company, caseId, OutDir())
    End If

    ' 既定の置き場所を題名で示す(GetOpenFilename は初期パスを引数に取れない)。
    Dim picked As Variant
    picked = Application.GetOpenFilename("Excelブック (*.xlsx),*.xlsx", 1, _
                                         "企業ファイル " & initPath)
    If VarType(picked) = vbBoolean Then Exit Function
    AskCompanyFile = CStr(picked)
    Exit Function
Cancelled:
    AskCompanyFile = vbNullString
End Function

Private Function OutDir() As String
    OutDir = modConfig.GetStr(UH_DIR_KEY, UH_DIR_DEFAULT)
End Function

' ============================================================================
' ナレッジ再読込・プリフライト一括診断
' ============================================================================
Public Sub HomeReloadKnowledge()
    If Not modUIProgress.TryEnterUiLock("ナレッジ再読込") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    modUIProgress.SetStage "ナレッジブックを読み込み中", WaitSec()

    If Not modKnowledge.LoadKnowledge() Then
        ShowWarning "ナレッジブックへ接続できませんでした（前回の知識で続行します）。"
    Else
        ShowWarning vbNullString
    End If
    RefreshHome

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Public Sub HomePreflightAll()
    If Not modUIProgress.TryEnterUiLock("プリフライト診断") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    modUIProgress.SetStage "未診断の投函をプリフライト診断中", WaitSec()

    Dim n As Long
    n = modPlayOps.RunPreflightAll()
    ShowWarning "プリフライト診断: " & CStr(n) & "件を診断しました。", "info"

    modUIInbox.RefreshInbox
    RefreshHome

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub
