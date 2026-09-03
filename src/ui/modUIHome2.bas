Attribute VB_Name = "modUIHome2"
Option Explicit

' ============================================================================
' modUIHome2 - HOMEのOnActionハンドラ群(ui層・T-30。modUIHome の分割先)
' ----------------------------------------------------------------------------
' 30,000字契約(12章§2)による modUIHome の分割先(17章§7 Z-13)。分割の軸は
'   modUIHome  = HOMEの**画面**(図形ボタンの配置表・RefreshHome・状態表示・
'                KbStatusText・ShowWarning・DrawAllSteps)
'   modUIHome2 = HOMEの**動作**(OnActionで配線されるボタンハンドラ)
' であり、**挙動は分割前と1つも変えていない**(移設のみ)。
'
' 実行の流れ(1ハンドラ=1アクション)は分割前と同じ:
'   TryEnterUiLock -> ParkFocus(E-51) -> SetStage(E-50(a))
'   -> app層(modPipeline / modExportHtml / modCompanyFile / modPlayOps)
'   -> 画面更新 -> ExitUiLock
' OnActionで配線される全ハンドラは先頭で modUIProgress.TryEnterUiLock を通す
' (16章 E-11。多重クリック・案件切替の関所)。
'
' 画面側の口は modUIHome の公開関数から借りる(同じ実装を2箇所に持たない。
' 14章§6へ登記): RefreshHome / SelectedCaseId / ShowWarning / DrawAllSteps /
' WriteRoundNo。名前付きレンジ定数は借りず、公開関数の引数で受け渡す。
'
' S1～S4シートの[SNから再実行][SN+1へ進む]も本モジュールのハンドラを指す
' (配線は modUICase2.EnsureStepButtons)。
' ============================================================================

' 13章§2.10 の名前付きレンジ(本モジュールが読むぶん)。
Private Const UH_QUALITY As String = "hm_quality_mode"

' 18章§1.1(1): S1未実行時のレポート出力は**エラーコードを立てず**中止する。
Private Const UH_MSG_NEED_S1 As String = "先に③の[まとめて作る]を押してください。会社の理解の下書きがまだありません。"
Private Const UH_MSG_NO_CASE As String = "案件が選ばれていません。いちばん上の帯で、案件を選んでください。"

' 企業ドシエファイルの保存先。13章§2.3 に専用キーが無いため、出力の共通
' フォルダ(html_out_dir)を使う(17章の裁定事項として申し送り)。
Private Const UH_DIR_KEY As String = "html_out_dir"
' 裁定書27 W9-C2: 保存先の正は config data_dir(既定は会社のOneDrive)。
' html_out_dir に値が入っていればそちらを優先する(分けたい管理者向け)。
Private Const UH_DATA_DIR_KEY As String = "data_dir"
Private Const UH_DIR_DEFAULT As String = "%OneDriveCommercial%\リスク提案ナビ\データ"

' HOMEの品質上書き(13章§2.10 hm_quality_mode)。日本語ラベル -> 機械値。
'   空(未選択)は "" のまま返し、config・ティア連動の解決へ委ねる。
Private Function QualityOverride() As String
    QualityOverride = modUICase.EnumEn("quality_mode", modUISheet.ReadNamed(UH_QUALITY))
End Function

' 進捗表示に出す最大待ち時間(config llm_wait_sec)。
Private Function WaitSec() As Long
    WaitSec = modConfig.GetLong("llm_wait_sec", 1200)
End Function

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

    modUIHome.ShowWarning warnText, "warn"
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
    modUIHome.ShowWarning vbNullString
    ' 裁定書10 M1/N9: deep outcome のリセットは**実行の開始時に1回だけ**。
    ' 一括実行の中では消さない(Step2/3 の結末を Step4 が消さないため)。
    modPipeline2.ResetDeepOutcome

    Dim caseId As String
    caseId = modUIHome.SelectedCaseId()
    If LenB(caseId) = 0 Then
        modUIHome.ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    ' 裁定書22 M3: 段ごとに SetStage を出す(進捗が「4枚まとめて」の1行のまま
    '   10～20分止まって見えるのを止める)。modPipeline.RunAll の中身は
    '   「For i=1 To 4: RunStep -> 失敗で Exit / DoEvents」だけであり、前後処理を
    '   1つも持たない(ResetDeepOutcome は裁定書10 M1 で呼び出し側=本ハンドラの
    '   冒頭へ移してあり、E-35/E-36 の警告収集は ShowDeepWarning、run_log は
    '   RunStep の中の modGatewayRPN が書く)。したがって本ループは RunAll と
    '   **等価**である(app層は触らない)。
    Dim stepNo As Long
    Dim okAll As Boolean
    okAll = True
    For stepNo = 1 To 4
        modUIProgress.SetStage StageNameOf(stepNo), WaitSec(), stepNo, 4
        If Not modPipeline.RunStep(caseId, stepNo, QualityOverride()) Then
            okAll = False
            Exit For
        End If
        DoEvents
    Next stepNo

    If okAll Then
        modUIHome.DrawAllSteps caseId
        modUIToast.ShowNext 2
        ShowDeepWarning
    Else
        modUIHome.ShowWarning "実行が完了しませんでした。err_log をご確認ください。"
    End If
    modUIHome.RefreshHome

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' 段の名前(11章§4.2(4): ステータスバー・カード・進捗4行・トーストの4箇所で
'   同じ段名を使う。4箇所で違う呼び方をしない)。
Public Function StageNameOf(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 1
        StageNameOf = "① 会社の理解を作っています"
    Case 2
        StageNameOf = "② リスクの洗い出しを作っています"
    Case 3
        StageNameOf = "③ 提案の候補を作っています"
    Case 4
        StageNameOf = "④ 提案の骨子を作っています"
    Case Else
        StageNameOf = "下書きを作っています"
    End Select
End Function

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
    modUIHome.ShowWarning vbNullString
    ' 裁定書10 M1/N9: deep outcome のリセットは実行の開始時に1回だけ。
    modPipeline2.ResetDeepOutcome

    Dim caseId As String
    caseId = modUIHome.SelectedCaseId()
    If LenB(caseId) = 0 Then
        modUIHome.ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    SaveUpstreamEdits caseId, stepNo
    If Not ConfirmInvalidate(caseId, stepNo) Then GoTo Done

    modUIProgress.SetStage StageNameOf(stepNo), WaitSec(), stepNo, 4
    If modPipeline.RunStep(caseId, stepNo, QualityOverride()) Then
        modUICase2.DrawStep caseId, stepNo
        modUISheet.ShowSheet modUICase2.SheetNameOf(stepNo)
        ShowDeepWarning
    Else
        modUIHome.ShowWarning "Step" & CStr(stepNo) & " が完了しませんでした。err_log をご確認ください。"
    End If
    modUIHome.RefreshHome

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
    modUIHome.ShowWarning "Step" & CStr(upstream) & " の編集内容が検証に通らなかったため、" & _
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


' ============================================================================
' 新規案件・画面遷移
' ============================================================================
Public Sub HomeNewCase()
    If Not modUIProgress.TryEnterUiLock("新規案件") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    modUIHome.ShowWarning vbNullString

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
    modUINavDraw.ResetForNewCase vbNullString
    modUISheet.WriteNamed "ci_case_id", modUICase3.U3_NEW_MARK
    modUISheet.ShowSheet "ナビ"
    modUIToast.ShowNext 1

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' v3.2: 「案件入力」シートは廃止され、区画②はナビの中にある(11章§1.1)。
'   本ハンドラは残るが、開くのはナビであり、描き直しは modUINav が行う。
Public Sub HomeOpenCaseInput()
    If Not modUIProgress.TryEnterUiLock("貼る欄を開く") Then Exit Sub
    On Error GoTo Done
    modUISheet.ShowSheet "ナビ"
    modUINav.DrawNav
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
    modUISparring.OpenSparring modUIHome.SelectedCaseId()
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
    modUIHome.ShowWarning vbNullString

    Dim caseId As String
    caseId = modUIHome.SelectedCaseId()
    If LenB(caseId) = 0 Then
        modUIHome.ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    Dim answer As Long
    answer = MsgBox("現在のS2を前ラウンドとして確定し、第2ラウンドを開始します。" & _
                    "よろしいですか", vbYesNo + vbQuestion, "第2ラウンド開始")
    If answer <> vbYes Then GoTo Done

    Dim newRound As Long
    newRound = modCaseStore.FreezeRound(caseId)
    If newRound <= 0 Then
        modUIHome.ShowWarning "第2ラウンドを開始できませんでした（前ラウンドの退避に失敗しました）。" & _
                    "err_log をご確認ください。"
        GoTo Done
    End If

    modUIHome.WriteRoundNo newRound
    modUIHome.RefreshHome
    modUIHome.ShowWarning "第" & CStr(newRound) & "ラウンドを開始しました（前ラウンドのS2を退避しました）。", "info"

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
    modUIHome.ShowWarning vbNullString

    Dim caseId As String
    caseId = modUIHome.SelectedCaseId()
    If LenB(caseId) = 0 Then
        modUIHome.ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    ' S1未実行ガード(18章§1.1(1))。データ未整備は障害ではないので E-code を
    ' 立てず、hm_warning で案内して生成そのものを呼ばない。
    If LenB(Trim$(modCaseStore.ResolveStepJson(caseId, 1))) = 0 Then
        modUIHome.ShowWarning UH_MSG_NEED_S1, "warn"
        GoTo Done
    End If

    modUIProgress.SetStage "レポートを作っています", WaitSec(), 1, 1

    Dim outPath As String
    Dim errText As String
    errText = modExportHtml.GenerateHtmlReport(caseId, outPath)
    If LenB(errText) > 0 Then
        modUIHome.ShowWarning errText
    Else
        ' 13章§2.1・11章§4(裁定書9 B16(b)): HTMLレポートの生成成功が
        ' case_status の exported を立てる唯一の点。SetStatus は遷移検査を
        ' 通さない(14章§6の注記)が、CS_STATUSES_ABOVE_S4 の降格抑止は効く。
        modCaseStore.SetStatus caseId, "exported"
        modUIHome.ShowWarning "リスクレポートを出力しました: " & outPath, "info"
        modUIToast.ShowNext 3
    End If
    modUIHome.RefreshHome

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Public Sub HomeBuildHearing()
    If Not modUIProgress.TryEnterUiLock("ヒアリングシート") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    modUIHome.ShowWarning vbNullString

    Dim caseId As String
    caseId = modUIHome.SelectedCaseId()
    If LenB(caseId) = 0 Then
        modUIHome.ShowWarning UH_MSG_NO_CASE, "warn"
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
            modUIHome.ShowWarning "ヒアリングシートは作り直しませんでした（手書きの回答を残しました）。", "warn"
            GoTo Done
        End If
    End If

    If modExportHearing.BuildHearingSheet(caseId) Then
        ' 13章§2.9(v3.2): 既定は非表示。作成に成功したときだけ可視にし、
        ' 1行目へ[ナビへ戻る]を置く(行き止まりを作らない。11章§3.5)。
        Dim hsWs As Object
        Set hsWs = modUISheet.SheetOf("ヒアリングシート")
        If Not hsWs Is Nothing Then modUISheet.EnsureBackButton hsWs
        modUISheet.ShowSheet "ヒアリングシート"
        modUIToast.ShowNext 4
    Else
        modUIHome.ShowWarning "ヒアリングシートを作れませんでした。先に③の[まとめて作る]を押してください。"
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
    modUIHome.ShowWarning vbNullString

    Dim caseId As String
    caseId = modUIHome.SelectedCaseId()
    If LenB(caseId) = 0 Then
        modUIHome.ShowWarning UH_MSG_NO_CASE, "warn"
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
            modUIHome.ShowWarning "個人情報の確認が済んでいないため書き出しませんでした。", "warn"
            GoTo Done
        End If
        confirmedAt = modUtil.NowStamp()
    End If

    Dim pathText As String
    pathText = modCompanyFile.ExportCompanyFile(caseId, OutDir(), confirmedAt)
    If LenB(pathText) = 0 Then
        modUIHome.ShowWarning "企業ファイルへ書き出せませんでした。err_log をご確認ください。"
    Else
        modUIHome.ShowWarning "企業ファイルへ保存しました: " & pathText, "info"
    End If

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Public Sub HomeCompanyOpen()
    If Not modUIProgress.TryEnterUiLock("企業ファイルを開く") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    modUIHome.ShowWarning vbNullString

    Dim caseId As String
    caseId = modUIHome.SelectedCaseId()
    If LenB(caseId) = 0 Then
        modUIHome.ShowWarning UH_MSG_NO_CASE, "warn"
        GoTo Done
    End If

    Dim pathText As String
    pathText = AskCompanyFile(caseId)
    If LenB(pathText) = 0 Then GoTo Done

    If modCompanyFile.ImportCompanyFile(pathText, caseId) Then
        modUIHome.DrawAllSteps caseId
        modUIHome.RefreshHome
        modUIHome.ShowWarning "企業ファイルから前ラウンドの内容を取り込みました。", "info"
    Else
        modUIHome.ShowWarning "企業ファイルを取り込めませんでした。ファイルをご確認ください。"
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

' 企業ファイルの書出先(裁定書27 W9-C2)。html_out_dir が空なら data_dir に従い、
'   OneDrive が無ければ Documents へ落ちる(落ちたことは NoticeDataDir が出す)。
Private Function OutDir() As String
    Dim raw As String
    raw = Trim$(modConfig.GetStr(UH_DIR_KEY, vbNullString))
    If LenB(raw) = 0 Then raw = Trim$(modConfig.GetStr(UH_DATA_DIR_KEY, UH_DIR_DEFAULT))
    If LenB(raw) = 0 Then raw = UH_DIR_DEFAULT

    Dim dirText As String
    dirText = modUtil.ResolveDataDir(raw)
    If LenB(dirText) = 0 Then dirText = raw
    OutDir = dirText
End Function

' ============================================================================
' ナレッジ再読込・プリフライト一括診断
' ============================================================================
Public Sub HomeReloadKnowledge()
    If Not modUIProgress.TryEnterUiLock("ナレッジ再読込") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    modUIProgress.SetStage "社内ナレッジを読み込んでいます", WaitSec(), 1, 1

    If Not modKnowledge.LoadKnowledge() Then
        modUIHome.ShowWarning "ナレッジブックへ接続できませんでした（前回の知識で続行します）。"
    Else
        modUIHome.ShowWarning vbNullString
    End If
    modUIHome.RefreshHome

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

Public Sub HomePreflightAll()
    If Not modUIProgress.TryEnterUiLock("プリフライト診断") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus
    modUIProgress.SetStage "投函をまとめて診断しています", WaitSec(), 1, 1

    Dim n As Long
    n = modPlayOps.RunPreflightAll()
    modUIHome.ShowWarning "プリフライト診断: " & CStr(n) & "件を診断しました。", "info"

    modUIInbox.RefreshInbox
    modUIHome.RefreshHome

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub
