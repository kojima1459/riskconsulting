Attribute VB_Name = "modUIHome"
Option Explicit

' ============================================================================
' modUIHome - ナビの状態表示と画面用意の入口(ui層・T-30)
' ----------------------------------------------------------------------------
' 13章§2.10 の名前付きレンジ(hm_*)が正。**v3.2 で HOME シートは廃止**され、
' hm_* は名前を変えずにナビシートのセルへ付け替えられた(11章§8.1)。
' modUISheet.NamedCell はブック・スコープの ThisWorkbook.Names で解決するため、
' 本モジュールの読み書きは**シート名を1文字も知らずに**そのまま動く。
' ボタンの配置表は modUINav が持つ(HOME の主要動線5本+くわしい操作14本は廃止)。
'
' 30,000字契約(12章§2)による分割(17章§7 Z-13)。本モジュールは**画面**だけを
' 持ち、OnActionで配線される**ボタンハンドラ群は modUIHome2** にある。
' ハンドラ側が使う画面の口は本モジュールの公開関数(RefreshHome /
' SelectedCaseId / ShowWarning / DrawAllSteps / WriteRoundNo)で渡す。
'
' S1～S4シートの[SNから再実行][SN+1へ進む]は modUIHome2 のハンドラを指す
' (同じ動作の実装を2箇所に置かない。配線は modUICase2.EnsureStepButtons)。
' ============================================================================

Private Const UH_SRC As String = "modUIHome"
Private Const UH_SHEET As String = "ナビ"

' 13章§2.10 の名前付きレンジ。
Private Const UH_CASE_ID As String = "hm_case_id"
Private Const UH_CASE_LABEL As String = "hm_case_label"
Private Const UH_STATUS As String = "hm_status"
Private Const UH_ROUND As String = "hm_round_no"
Private Const UH_KB As String = "hm_kb_status"
Private Const UH_BANNER As String = "hm_transport_banner"
Private Const UH_WARNING As String = "hm_warning"
' 16章 E-27: transport が ribbon 以外のときナビの区画④に常時出す赤帯。
Private Const UH_BANNER_TEXT As String = "デモ／開発経路で動作中です（本番の判断に使わないでください）"

' 画面制御用のモジュール変数(裁定書9 B1)。**永続でない画面制御変数**であり、
' 14章§6の「状態保持の例外」への登録は要らない。直前に RefreshHome が描いた
' 対象案件IDを覚えておき、変わっていたらS1～S4を描き直す(前の案件の画面が残った
' まま次の案件の sN_edited として確定するクロス案件汚染を、描画側でも塞ぐ)。
Private gShownCaseId As String

' ============================================================================
' 起動時の画面用意(11章§5の図形ボタンとenum入力規則)。modBoot から呼ぶ。
' ============================================================================
Public Sub EnsureScreens()
    On Error Resume Next

    ' v3.2: 利用者が触る画面はナビ1枚(11章§0.2)。ナビの図形は modUINav が置く。
    modUIResearch.EnsureCopyButtons
    modUICase2.EnsureStepButtons
    modUICase4.EnsureRecordButtons
    modUIInbox.EnsureInboxButtons
    modUISparring.EnsureSparringButtons
    ' 使い方タブの[テストを実行][ツアーをもう一度見る][記録を見る][表示する]
    modUIGuide.EnsureGuideButtons
    ' 使い方タブの最下部のフッター(裁定書26 D)。ui_advanced=FALSE のときも
    ' 置くので EnsureGuideButtons とは別に呼ぶ。
    modUIGuide.EnsureFooterButton

    RefreshHome

    ' W9.2 N3: On Error Resume Next が握りつぶした失敗を記録だけは残す
    ' (起動経路の最外周の網。黙って画面が半分だけ出た状態を作らない)。
    If Err.Number <> 0 Then
        modLog.LogError "E0603", UH_SRC & ".EnsureScreens", "ensure_screens_failed", Err.Number
        Err.Clear
    End If
End Sub

' ============================================================================
' RefreshHome - 案件の表示・ナレッジ状態・赤帯(13章§2.10)。書き終えたら
'   ナビを描き直す(状態が変われば帯の1文と強調枠も変わるため)。
' ============================================================================
Public Sub RefreshHome()
    ' W9.2 N8: 起動シーケンス(EnsureScreens)から呼ばれる。ここが未捕捉の
    ' 実行時エラーを外へ出すと、起動直後に生ダイアログが出る。必ず受け止めて
    ' err_log へ残す(画面が古いままでも業務は止めない)。
    On Error GoTo Failed

    RefreshState
    ' 状態が変われば帯の1文・進捗ドット・強調枠の位置も変わる。**状態を書いた
    ' あとに必ず描き直す**ことで、「画面と状態が食い違ったまま」を作らない。
    ' (DrawNav は本モジュールを SelectedCaseId でしか呼ばないので再帰しない)
    modUINav.DrawNav
    Exit Sub

Failed:
    ' ハンドラ稼働中は On Error Resume Next が効かないので、記録は別Subへ。
    LogRefreshFailure Err.Number
End Sub

' RefreshHome の Failed: から呼ぶ記録専用(W9.2 N8)。ハンドラの外なので網が張れる。
Private Sub LogRefreshFailure(ByVal errNo As Long)
    On Error Resume Next
    modLog.LogError "E0603", UH_SRC & ".RefreshHome", "refresh_home_failed", errNo
End Sub

' 状態表示の書き込みだけを行う(描画は呼び出し側の RefreshHome が続けて行う)。
Private Sub RefreshState()
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
                       "[ナレッジを読み直す]を押してください（探した場所: " & _
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
