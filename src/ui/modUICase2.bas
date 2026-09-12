Attribute VB_Name = "modUICase2"
Option Explicit

' ============================================================================
' modUICase2 - S1～S4シートの描画と逆シリアライズ(ui層・T-31)
' ----------------------------------------------------------------------------
' 30,000字契約(12章§2)による modUICase の分割先。**呼んでよいのは modUICase /
' modUIHome** であり、変換表(19章§3)は modUICase から引く(表を2箇所に持たない)。
'
' 本モジュールは「書く側(描画)と確定(SaveEditedStep)」を持ち、「読む側(逆シリア
' ライズ)とブロックの幾何」は 30,000字契約により modUICase5 へ分けた(裁定書9 W4.1
' の司令塔裁定)。読み書きで同じ列定義・同じ部屋の数え方を使う不変条件は、
' modUICase5.ColIndexes / ColCount / RoomOf を**両側が同じ1本を呼ぶ**ことで保つ
' (数え方を2箇所に持たない)。
'
' 本モジュールが持つ規約:
'   ・13章§2.12-§2.15 の列定義(Cols*)と 13章§2.2 のセル格納規約の変換関数は
'     modUICaseFmt が唯一持ち、本モジュールはそれを**書き出しと逆シリアライズの
'     両方で同じように読む**(片方だけ直すことができない)。
'   ・13章§2.2 逆シリアライズ規約5「空行が表の終端」・規約2「順序列の昇順」・
'     規約3「順序列はJSONへ書き戻さない」。
'   ・保存前に必ず modValidate.CheckSN を通し、不合格なら sN_edited を保存しない
'     (規約4)。下流は従前の参照優先(sNr_json / sN_json)を使い続ける。
'
' ブロックの行数(部屋)は**次のブロックのアンカー行から動的に決める**。13章§2.9が
' 「行番号を仮定しない」と定めているため、確保行数を定数で持たない(持つと台帳の
' reserve_rows を変えた瞬間に隣のブロックの見出しを消す)。
' ============================================================================

Private Const U2_SRC As String = "modUICase2"
Private Const U2_WARN As String = "hm_warning"    ' 13章§2.10 HOMEの警告欄
Private Const U2_MSG_MISMATCH As String = "画面の案件と保存先が一致しません。再描画してください"

' 画面制御用のモジュール変数(裁定書9 B6。永続でない**描画状態のフラグ**であり、
' 14章§6の「状態保持の例外」への登録は要らない)。
'   gDrawStep    = いま DrawStep が描いている Step 番号(DrawArrBlock への文脈)
'   gTruncStep() = 直近の描画で部屋あふれ(切捨て)が起きた Step
'   gTruncNote   = 切捨ての内訳(利用者向け文言の材料)
'   gStepNotice  = 直近の実行で実際に画面へ出した 16章 E-02 の警告帯(裁定書40
'                  Q-m3)。あとから出る deep モードの警告が**上書きで消さない**
'                  ために、書いた本文をそのまま覚えておく。modPipeline2 の
'                  LastDeepOutcome / ResetDeepOutcome と同じ形(実行の開始時に
'                  呼出側が ResetStepNotice で1回だけ消す)。
'   gDeepNotice  = 入念モードの警告(16章 E-35/E-36)の本文(裁定書41 §2)。
'                  これも同じ hm_warning の1枠を使うので、gStepNotice と同じ
'                  ように覚えて WriteWarnCell が3本まとめて書く。
Private gDrawStep As Long
Private gTruncStep(1 To 4) As Boolean
Private gTruncNote As String
Private gStepNotice As String
Private gDeepNotice As String
Private Const U2_MAX_ROOM As Long = 200        ' 最終ブロックの部屋の上限
Private Const U2_HDR_WIDTH As Long = 24        ' 見出し行の探索幅(列番号ではない)
Private Const U2_FIRST_COL As Long = 1         ' ブロックの左端列(build/sheets_main.json)

Public Function SheetNameOf(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 1
        SheetNameOf = "S1_企業プロファイル"
    Case 2
        SheetNameOf = "S2_リスク仮説"
    Case 3
        SheetNameOf = "S3_提案"
    Case 4
        SheetNameOf = "S4_骨子"
    End Select
End Function

' ============================================================================
' SerializeStep - シート -> JSON(13章§2.2 逆シリアライズ規約)
'   組めなければ ""(捏造しない)。
' ============================================================================
Public Function SerializeStep(ByVal stepNo As Long) As String
    On Error GoTo Failed
    SerializeStep = modUICase5.SerializeBody(stepNo)
    Exit Function
Failed:
    modLog.LogError "E0302", U2_SRC & ".SerializeStep", _
                    "serialize_failed:s" & CStr(stepNo), Err.Number
    SerializeStep = vbNullString
End Function


' ============================================================================
' SaveEditedStep - 13章§2.2 逆シリアライズ規約4
'   シート -> JSON -> CheckSN -> 合格なら sN_edited を保存。
'   不合格は保存せず errText に不合格理由(呼び出し側が画面へ出す)。
' ============================================================================
Public Function SaveEditedStep(ByVal caseId As String, ByVal stepNo As Long, _
                               ByRef errText As String) As Boolean
    On Error GoTo Failed
    errText = vbNullString

    If stepNo < 1 Or stepNo > 4 Then Exit Function
    If Not modCaseStore.IsValidCaseId(caseId) Then
        errText = "案件IDが不正です。"
        Exit Function
    End If

    ' 裁定書9 B6: 部屋あふれで表示しきれていない行があるあいだは保存しない
    ' (切り詰められた画面が参照優先の最上位に居座るのを防ぐ)。
    ' 裁定書10 m3: このフラグは揮発性であり**保存ブロックの根拠は下の案件ID
    ' 一致ガード(切捨て時は DrawStep が sN_case_id を空のままにする)**が担う。
    ' ここは「なぜ止まったのか」を具体的な内訳で伝えるための文言専用であり、
    ' 一致ガードより先に置いて、より具体的な案内が出るようにする。
    If gTruncStep(stepNo) Then
        modUISheet.WriteNamed U2_WARN, "表示しきれていない行があるため、Step" & _
            CStr(stepNo) & " の編集は保存しませんでした（" & gTruncNote & "）。"
        Exit Function
    End If

    ' 裁定書9 B1: 画面が別案件の内容のまま sN_edited を確定させない。読取専用の
    ' 案件ID表示セル(13章§2.12 sN_case_id)と引数が一致しなければ1セルも保存しない。
    ' 呼び出し側で二重に文言を出さないよう errText は空のままにし、画面へ直接出す。
    If Trim$(modUISheet.ReadNamed(CaseIdCellOf(stepNo))) <> caseId Then
        modUISheet.WriteNamed U2_WARN, U2_MSG_MISMATCH
        modLog.LogError "E0302", U2_SRC & ".SaveEditedStep", _
                        "case_id_mismatch:s" & CStr(stepNo)
        Exit Function
    End If

    Dim jsonText As String
    jsonText = SerializeStep(stepNo)
    If LenB(jsonText) = 0 Then
        errText = "画面から内容を読み取れませんでした（ブロックの見出しをご確認ください）。"
        Exit Function
    End If

    errText = ValidateEdited(caseId, stepNo, jsonText)
    If LenB(errText) > 0 Then
        modLog.LogError "E0302", U2_SRC & ".SaveEditedStep", _
                        "edited_rejected:s" & CStr(stepNo)
        Exit Function
    End If

    If Not modCaseStore.SaveData(caseId, "s" & CStr(stepNo) & "_edited", jsonText) Then
        errText = "編集内容を保存できませんでした。"
        Exit Function
    End If

    modLog.LogUsage "sheet_edited_saved", caseId, "step=s" & CStr(stepNo)

    ' 裁定書9 B14(16章NFR-S7 設計原則(2)): 一次のfail-closed(注入テキスト照合)を
    ' 残したまま、ナレッジブック本体との**二次照合**を後段で行う。不一致は保存を
    ' ブロックせず hm_warning へ出す(注入の切詰めやKB更新でも差分は生じうるため)。
    Dim unknownIds As String
    unknownIds = UnknownKbIds(stepNo, jsonText)
    If LenB(unknownIds) > 0 Then
        modUISheet.WriteNamed U2_WARN, "ナレッジブックに見当たらないIDがあります（" & _
            modUtil.SafeLeft(unknownIds, 200) & "）。内容をご確認ください。"
    End If

    SaveEditedStep = True
    Exit Function

Failed:
    errText = "編集内容の保存中にエラーが発生しました。"
    modLog.LogError "E0302", U2_SRC & ".SaveEditedStep", "save_failed", Err.Number
    SaveEditedStep = False
End Function

' 編集JSONの検証。modPipeline の防衛線(3)と同じ Check* を同じ材料で呼ぶ
' (ID実在検査は fail-closed なので一覧テキストを必ず渡す。16章E-07)。
Private Function ValidateEdited(ByVal caseId As String, ByVal stepNo As Long, _
                                ByVal jsonText As String) As String
    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String
    Dim s4Variant As String
    Dim tierText As String

    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        ValidateEdited = "案件一覧からこの案件を読めませんでした。"
        Exit Function
    End If

    Select Case stepNo
    Case 1
        ValidateEdited = modValidate.CheckS1(jsonText, ctx.case_type, _
            (LenB(Trim$(modCaseStore.LoadData(caseId, "input_field_notes"))) > 0))
    Case 2
        ValidateEdited = modValidate.CheckS2(jsonText, ctx.case_type, _
            modKnowledge.MenusSummaryFor(ctx.industry_code, 0), _
            modCaseStore.LoadData(caseId, "s2_prev_json"))
    Case 3
        ValidateEdited = modValidate.CheckS3(jsonText, _
            modCaseStore.ResolveStepJson(caseId, 2), _
            modKnowledge.MenusFor(ctx.industry_code, 0), _
            modKnowledge.LinesText(0), _
            modKnowledge.SchemesFor(ctx.industry_code, 0), _
            modKnowledge.CasesFor(ctx.industry_code, 0), ctx.case_type)
    Case 4
        ValidateEdited = modValidate.CheckS4(jsonText, tierText, _
            modConfig.GetLong("ppt_max_slides_t2", 10))
    End Select
End Function

' 二次照合(裁定書9 B14)。JSONに現れるナレッジIDのうち modKnowledge 本体に実在
'   しないものを ";" 区切りで返す(全て実在なら "")。走査表は
'   `配列キー|要素内のキー(空=要素そのものがID)|種別` の vbLf 区切り。
Private Function UnknownKbIds(ByVal stepNo As Long, ByVal jsonText As String) As String
    On Error GoTo Skip1

    Dim spec As String
    If stepNo = 2 Then
        spec = "risks>preventions|related_menu_id|m"
    ElseIf stepNo = 3 Then
        spec = "stories|menu_ids|m" & vbLf & "stories|line_ids|l" & vbLf & _
               "stories|scheme_id|s" & vbLf & "stories|similar_case_id|c"
    Else
        Exit Function
    End If

    Dim buf() As String
    Dim cnt As Long
    modUtil.BufInit buf, cnt

    Dim rows1() As String
    rows1 = Split(spec, vbLf)
    Dim i As Long
    For i = LBound(rows1) To UBound(rows1)
        ScanIds jsonText, rows1(i), buf, cnt
    Next i

    UnknownKbIds = Replace(modUtil.BufText(buf, cnt), vbLf, ";")
    Exit Function
Skip1:
    UnknownKbIds = vbNullString
End Function

' 走査表1行ぶん。外側配列(">"で入れ子を1段)の各要素からIDを取り出して照合する。
Private Sub ScanIds(ByVal jsonText As String, ByVal rowText As String, _
                    ByRef buf() As String, ByRef cnt As Long)
    Dim flds() As String
    flds = Split(rowText, "|")
    If UBound(flds) - LBound(flds) <> 2 Then Exit Sub

    Dim outerKey As String
    Dim innerKey As String
    Dim keyName As String
    Dim kindText As String
    outerKey = flds(LBound(flds))
    keyName = flds(LBound(flds) + 1)
    kindText = flds(LBound(flds) + 2)

    Dim p As Long
    p = InStr(1, outerKey, ">", vbBinaryCompare)
    If p > 0 Then
        innerKey = Mid$(outerKey, p + 1)
        outerKey = Left$(outerKey, p - 1)
    End If

    Dim outer1 As Collection
    Set outer1 = modJsonLite.GetArrayItems(jsonText, outerKey)

    Dim i As Long
    For i = 1 To outer1.count
        If LenB(innerKey) > 0 Then
            Dim inner1 As Collection
            Dim j As Long
            Set inner1 = modJsonLite.GetArrayItems(CStr(outer1(i)), innerKey)
            For j = 1 To inner1.count
                AddIfMissing modJsonLite.GetStr(CStr(inner1(j)), keyName), kindText, buf, cnt
            Next j
        Else
            Dim ids As Collection
            Dim k As Long
            Set ids = modJsonLite.GetArrayItems(CStr(outer1(i)), keyName)
            If ids.count > 0 Then
                For k = 1 To ids.count
                    AddIfMissing CStr(ids(k)), kindText, buf, cnt
                Next k
            Else
                AddIfMissing modJsonLite.GetStr(CStr(outer1(i)), keyName), kindText, buf, cnt
            End If
        End If
    Next i
End Sub

' 種別ごとの実在検査(14章§6 modKnowledge の5本のうち画面に現れる4本)。
Private Sub AddIfMissing(ByVal idText As String, ByVal kindText As String, _
                         ByRef buf() As String, ByRef cnt As Long)
    Dim id1 As String
    id1 = Trim$(idText)
    If LenB(id1) = 0 Then Exit Sub

    Dim okId As Boolean
    Select Case kindText
    Case "m"
        okId = modKnowledge.MenuIdExists(id1)
    Case "l"
        okId = modKnowledge.LineIdExists(id1)
    Case "s"
        okId = modKnowledge.SchemeIdExists(id1)
    Case Else
        okId = modKnowledge.CaseLibIdExists(id1)
    End Select
    If Not okId Then modUtil.BufAdd buf, cnt, id1
End Sub

' 13章§2.12: S1～S4の読取専用の案件ID表示セル(名前付きレンジ)。
Private Function CaseIdCellOf(ByVal stepNo As Long) As String
    CaseIdCellOf = "s" & CStr(stepNo) & "_case_id"
End Function

' ============================================================================
' DrawStep - JSON -> シート(参照優先の解決は modCaseStore.ResolveStepJson)
' ----------------------------------------------------------------------------
' afterRun(裁定書40 Q-m1): **この描画が実行の直後かどうか**。既定は False。
'   16章 E-02 の警告帯は「実行後」の規定なので、実行の直後に描き直す呼び口
'   (modUIHome2.RunStepUi / HomeRunAll -> modUIHome.DrawAllSteps)だけが True を
'   渡す。それ以外の呼び口は**実行していない**ので False のまま置く:
'     ・modNaviActions の open_step_sheet([シートで編集]で開くだけ)
'     ・modUIHome.RefreshHome の案件切替による描き直し
'     ・modUIHome2 の企業ファイル取込のあとの描き直し
'   既定を False にしてあるので、呼び口を増やしたときに**黙って警告帯が出る**
'   ことはない(出したい側が明示する)。
' ============================================================================
Public Function DrawStep(ByVal caseId As String, ByVal stepNo As Long, _
                         Optional ByVal afterRun As Boolean = False) As Boolean
    On Error GoTo Failed

    If stepNo < 1 Or stepNo > 4 Then Exit Function

    ' 描き切るまでは「どの案件のものでもない」画面として扱う(裁定書9 B1)。
    ' 案件IDが不正・案件一覧に無いときもここで空へ戻すので、案件を切り替えた
    ' 直後に前の案件の画面が「今の案件のもの」として保存されることはない。
    modUISheet.WriteNamed CaseIdCellOf(stepNo), vbNullString
    If Not modCaseStore.IsValidCaseId(caseId) Then Exit Function

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String
    Dim s4Variant As String
    Dim tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        Exit Function
    End If

    ' 裁定書9 B1/B6: この描画で「どの案件を」「切り詰めずに」描けたかを記録する。
    ' 案件ID表示セルは**描き切ったときだけ**書く。描けなかった(内容が無い・
    ' 例外)ときに書いてしまうと、前の案件の行が残ったまま表示case_idだけが
    ' 新しい案件になり、SaveEditedStep の一致検査をすり抜けてしまう。
    gDrawStep = stepNo
    gTruncStep(stepNo) = False
    gTruncNote = vbNullString
    ' 裁定書40 Q-m1: 実行**以外**の描画([シートで編集]・案件切替・取込)が
    ' 起きた時点で「実行直後」は終わる。前の実行で出した 16章 E-02 の帯と
    ' 入念モードの警告をここで手放し、以後の警告へ混ぜない(裁定書41 §2)。
    If Not afterRun Then
        gStepNotice = vbNullString
        gDeepNotice = vbNullString
    End If

    Dim jsonText As String
    jsonText = modCaseStore.ResolveStepJson(caseId, stepNo)
    If LenB(Trim$(jsonText)) = 0 Then
        modLog.LogUsage "sheet_draw_skipped", caseId, "step=s" & CStr(stepNo) & ";empty"
        Exit Function
    End If

    ' 16章E-31: 画面表示の直前に実名へ戻す({{POLICY_NO}} は復元表が無いので残す)。
    jsonText = modUICase.RestoreNames(jsonText, ctx.company)

    Select Case stepNo
    Case 1
        DrawS1 jsonText
    Case 2
        DrawS2 jsonText
    Case 3
        DrawS3 jsonText
    Case 4
        DrawS4 jsonText
    End Select

    ' 裁定書39 R1-07(a) / 裁定書40 Q-m1: シート画面にも 16章 E-02(実行後)の
    ' 警告帯を出す。**出すかどうかと文言は StepNoticeOf の1本**が持ち(純関数)、
    ' ここは書く場所を知っているだけ。RunStepUi は S1 成功後にここを必ず通る。
    ' 一括実行では S2〜S4 の描画で部屋あふれの警告が、そのあと入念モードの
    ' 警告(16章 E-35/E-36)が同じ hm_warning へ来るが、WriteWarnCell が3本とも
    ' 併記するので帯は消えない(裁定書40 Q-m3・裁定書41 §2)。
    ShowStepNotice StepNoticeOf(stepNo, jsonText, afterRun)

    ' ここまで来たら画面はこの案件の内容で描けている(13章§2.12)。ただし
    ' 裁定書10 m3: **部屋あふれ(切捨て)が起きた描画では案件ID表示セルを書かず
    ' 空のまま**にする。gTruncStep は未捕捉エラーや再コンパイルで False へ戻る
    ' 揮発性のフラグなので、それだけを保存ブロックの根拠にすると切り詰まった
    ' 画面がそのまま保存できてしまう。セルを空に保てば B1 の一致ガード
    ' (SaveEditedStep 冒頭)が**ブックに残る状態として**保存を止め続ける。
    ' 再描画して切捨てが無ければ、その描画で通常どおり書かれる。
    If Not gTruncStep(stepNo) Then
        modUISheet.WriteNamed CaseIdCellOf(stepNo), caseId
    End If

    modLog.LogUsage "sheet_drawn", caseId, "step=s" & CStr(stepNo)
    DrawStep = True
    Exit Function

Failed:
    ' 途中で落ちた画面は「どの案件のものでもない」。表示case_idを空にして、
    ' 中途半端な画面が sN_edited として確定する経路を閉じる(裁定書9 B1)。
    modUISheet.WriteNamed CaseIdCellOf(stepNo), vbNullString
    modLog.LogError "E0603", U2_SRC & ".DrawStep", "draw_failed:s" & CStr(stepNo), Err.Number
    DrawStep = False
End Function

Private Sub DrawS1(ByVal jsonText As String)
    DrawSingleRow "s1_basic", modUICaseFmt.ColsS1Basic(), jsonText
    DrawArrBlock "s1_locations", modUICaseFmt.ColsS1Locations(), jsonText, "locations"
    DrawArrBlock "s1_current_coverage", modUICaseFmt.ColsS1Coverage(), jsonText, "current_coverage"
    DrawArrBlock "s1_field_insights", modUICaseFmt.ColsS1Insights(), jsonText, "field_insights"
    DrawArrBlock "s1_missing_info", modUICaseFmt.ColsS1Missing(), jsonText, "missing_info"
    DrawArrBlock "s1_input_quality", modUICaseFmt.ColsS1Quality(), _
                 modUICaseFmt.SubJson(jsonText, "input_quality"), "coverage"
    DrawArrBlock "s1_research_requests", modUICaseFmt.ColsS1Research(), jsonText, "research_requests"
    DrawArrBlock "s1_sources", modUICaseFmt.ColsS1Sources(), jsonText, "sources"
    DrawQualityBanner jsonText
    DrawResearchButtons
End Sub

' ============================================================================
' StepNoticeOf - 1段を描いた直後に画面へ出す注記(16章 E-02)。**純関数**。
' ----------------------------------------------------------------------------
'   3つの条件が全部そろったときだけ本文を返し、それ以外は ""(=何も書かない。
'   他の警告を消さない):
'     (1) afterRun = True。E-02 は「実行後」の規定であり、[シートで編集]や
'         案件切替の描き直しでは出さない(裁定書40 Q-m1)。
'     (2) stepNo = 1。E-02(実行後)が見るのは S1 の input_quality だけ。
'     (3) iq=low。判断は modPipeline3.SufficiencyNoteOf、文言は
'         modUICase.IqBannerTextOf の1本(HTML画面と同じ逐語)。
'   純関数なので、純層テストがこの真理値表をそのまま固定できる(裁定書40 Q-m2)。
' ============================================================================
Public Function StepNoticeOf(ByVal stepNo As Long, ByVal jsonText As String, _
                             ByVal afterRun As Boolean) As String
    If Not afterRun Then Exit Function
    If stepNo <> 1 Then Exit Function
    StepNoticeOf = modUICase.IqBannerTextOf(jsonText)
End Function

' 注記を HOME の hm_warning とトーストへ出す。空なら**何も書かない**。
'   書いた本文は gStepNotice に覚えておく。hm_warning は1枠しかないので、
'   放っておくと「あとから書いた1本が前の1本を消す」が必ず起きる
'   (裁定書40 Q-m3)。
'   **実行にともなって出る3本**(16章 E-02 の帯・部屋あふれ・入念モードの
'   E-35/E-36)は、消し合わないよう WriteWarnCell の1本を通して NoticeJoin で
'   併記する。SaveEditedStep の3本(:103 / :112 / :148)は**保存を止めた理由を
'   単独で伝える場面**なので、いまも直に WriteNamed して1本だけを出す
'   (描画していないので併記すべき相手が無い。裁定書41 §2 でここを事実どおり
'   書き直した。以前の注記は「この画面から出る警告はすべて WriteWarnCell を
'   通す」と書いていたが、実物はそうなっていなかった)。
Private Sub ShowStepNotice(ByVal noticeText As String)
    On Error Resume Next
    If LenB(noticeText) = 0 Then Exit Sub
    gStepNotice = noticeText
    WriteWarnCell
    modUIToast.ShowToast noticeText, "warn"
End Sub

' ============================================================================
' ShowDeepNotice - 入念モードの警告(16章 E-35/E-36)を hm_warning へ出す唯一の
'   口(裁定書41 §2)。modUIHome2.ShowDeepWarning から呼ぶ。
' ----------------------------------------------------------------------------
'   以前は modUIHome2 が modUIHome.ShowWarning へ直行し、NoticeJoin で
'   E-02 の帯だけを併記していた。切捨ての警告(TruncWarnText。Private)は
'   外から取れないので**併記できず**、一括実行で S1〜S4 のどれかが部屋あふれ
'   した回は「編集を保存すると残りが失われるため、この画面の保存は行いません。」
'   が deep の警告に上書きされて消えていた(利用者は切り捨てに気付けない)。
'   本文を覚えて WriteWarnCell に3本まとめて書かせれば、どれも消えない。
'   トーストは ShowStepNotice と同じく**その回に出す本文だけ**を出す。
' ============================================================================
Public Sub ShowDeepNotice(ByVal deepText As String)
    On Error Resume Next
    If LenB(deepText) = 0 Then Exit Sub
    gDeepNotice = deepText
    WriteWarnCell
    modUIToast.ShowToast deepText, "warn"
End Sub

' ============================================================================
' NoticeJoin - 1枠しかない警告欄へ2本を**併記**する(裁定書40 Q-m3)。**純関数**。
' ----------------------------------------------------------------------------
'   どちらか片方しか無ければその1本を、両方あれば vbLf で連結して返す。
'   順序は「実行後の注記(16章 E-02)が先」。HTML画面の
'   modNaviActions.ActRunPipeline が同じ2本を同じ順で連結しているので、
'   予備経路(シート画面)だけ挙動を変えない。
'   警告を出す側が各々 WriteNamed すると必ず上書きが起きるので、**併記の仕方は
'   この1本**が持つ(WriteWarnCell が3本を2回に分けてここへ通す)。
' ============================================================================
Public Function NoticeJoin(ByVal firstText As String, _
                           ByVal secondText As String) As String
    If LenB(firstText) = 0 Then
        NoticeJoin = secondText
        Exit Function
    End If
    If LenB(secondText) = 0 Then
        NoticeJoin = firstText
        Exit Function
    End If
    NoticeJoin = firstText & vbLf & secondText
End Function

' 部屋あふれ(切捨て)の警告文。切捨てが無ければ ""。**逐語の値源はここ1本**
'   (NoteTruncation と WriteWarnCell が同じ文を2箇所に持たない)。
Private Function TruncWarnText() As String
    If LenB(gTruncNote) = 0 Then Exit Function
    TruncWarnText = gTruncNote & _
        "。編集を保存すると残りが失われるため、この画面の保存は行いません。"
End Function

' 実行にともなう警告3本をまとめて hm_warning へ書く唯一の口(裁定書40 Q-m3・
'   裁定書41 §2)。16章 E-02 の帯(実行直後のS1)・部屋あふれ・入念モードの
'   E-35/E-36 は**どれも消さない**。一括実行では S1 の帯 -> S2〜S4 の切捨て ->
'   deep の警告の順に来るので、放っておくと最後の1本しか残らなかった。
'   並びは HTML画面の modNaviActions.ActRunPipeline と同じ「E-02 が先」。
Private Sub WriteWarnCell()
    On Error Resume Next
    Dim bodyText As String
    bodyText = NoticeJoin(NoticeJoin(gStepNotice, TruncWarnText()), gDeepNotice)
    If LenB(bodyText) = 0 Then Exit Sub
    modUISheet.WriteNamed U2_WARN, modUIToast.WarnLine(bodyText, "warn")
End Sub

' LastStepNotice(裁定書40 Q-m3)は**削除した**(裁定書41 §2)。唯一の呼び口だった
'   modUIHome2.ShowDeepWarning が、警告本文を取り出して自分で書く形をやめ、
'   ShowDeepNotice へ本文を渡す形へ変わったため、外へ出す口が要らなくなった
'   (src 全体・docs・tools を grep して他の参照が無いことを確かめた)。

' ResetStepNotice - 明示リセット口(modPipeline2.ResetDeepOutcome と同じ考え方)。
'   **実行の開始時に1回だけ**呼ぶ(前回の実行の帯と deep 警告を今回の警告に
'   混ぜない。裁定書41 §2 で gDeepNotice も同時に手放すようにした)。
Public Sub ResetStepNotice()
    gStepNotice = vbNullString
    gDeepNotice = vbNullString
End Sub

Private Sub DrawS2(ByVal jsonText As String)
    DrawArrBlock "s2_gaps", modUICaseFmt.ColsS2Gaps(), jsonText, "gaps"
    DrawArrBlock "s2_risks", modUICaseFmt.ColsS2Risks(), jsonText, "risks"
    DrawScalarArr "s2_open_questions", "question", jsonText, "open_questions"
    DrawArrBlock "s2_emerging", modUICaseFmt.ColsS2Emerging(), jsonText, "emerging_risks"
End Sub

Private Sub DrawS3(ByVal jsonText As String)
    DrawArrBlock "s3_stories", modUICaseFmt.ColsS3Stories(), jsonText, "stories"
    DrawArrBlock "s3_unmatched_risks", modUICaseFmt.ColsS3Unmatched(), jsonText, "unmatched_risks"
    DrawArrBlock "s3_do_not_propose", modUICaseFmt.ColsS3DoNot(), jsonText, "do_not_propose"
    ' 13章§2.14 v2.6: 攻めの保険活用(配列)と経営層への話し方(単数オブジェクト)。
    DrawArrBlock "s3_growth_ideas", modUICaseFmt.ColsS3Growth(), jsonText, "growth_ideas"
    DrawSingleRow "s3_talk_script", modUICaseFmt.ColsS3Talk(), _
                  modUICaseFmt.SubJson(jsonText, "talk_script")
End Sub

Private Sub DrawS4(ByVal jsonText As String)
    DrawSingleRow "s4_meta", modUICaseFmt.ColsS4Meta(), jsonText
    DrawArrBlock "s4_slides", modUICaseFmt.ColsS4Slides(), jsonText, "slides"
    DrawArrBlock "s4_hearing_questions", modUICaseFmt.ColsS4Hearing(), jsonText, "hearing_questions"
End Sub

' 11章: 実行後はS1の入力充足度診断がS1シート上部に表示される。13章§2.12は
'   この行に列も名前付きレンジも与えていないので、データ面を汚さないよう
'   ラベル図形で描く(表示専用)。
Private Sub DrawQualityBanner(ByVal jsonText As String)
    On Error Resume Next
    Dim ws As Object
    Set ws = modUISheet.SheetOf(SheetNameOf(1))
    If ws Is Nothing Then Exit Sub

    Dim iq As String
    iq = modUICaseFmt.SubJson(jsonText, "input_quality")

    Dim overallJa As String
    overallJa = modUICase.EnumJa("input_quality_overall", modJsonLite.GetStr(iq, "overall"))

    Dim bannerText As String
    bannerText = "充足度: " & overallJa & " ｜ 助言: " & modJsonLite.GetStr(iq, "advice")

    modUISheet.EnsureLabel ws, "lbl_s1_quality", bannerText, 1, 3, 460#, 16#
End Sub

' ============================================================================
' ブロック単位の読み書き
' ============================================================================


' 単一行ブロックへ書く。
Private Sub DrawSingleRow(ByVal anchorName As String, ByVal colSpec As String, _
                          ByVal jsonText As String)
    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Sub

    Dim cols As Variant
    cols = modUICase5.ColIndexes(ws, headerRow, colSpec)
    If IsEmpty(cols) Then Exit Sub

    WriteRow ws, headerRow + 1, cols, colSpec, jsonText, 1, anchorName
End Sub

' 配列ブロックへ書く(部屋を超える分は書かない。書けなかった件数は記録する)。
Private Sub DrawArrBlock(ByVal anchorName As String, ByVal colSpec As String, _
                         ByVal jsonText As String, ByVal arrayKey As String)
    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Sub

    Dim cols As Variant
    cols = modUICase5.ColIndexes(ws, headerRow, colSpec)
    If IsEmpty(cols) Then Exit Sub

    Dim room As Long
    room = modUICase5.RoomOf(anchorName)
    modUISheet.ClearBlock ws, headerRow, U2_FIRST_COL, modUICase5.ColCount(colSpec), room

    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(jsonText, arrayKey)

    Dim i As Long
    For i = 1 To items.count
        If i > room Then Exit For
        WriteRow ws, headerRow + i, cols, colSpec, CStr(items(i)), i, anchorName
    Next i

    If items.count > room Then
        modLog.LogError "E0604", U2_SRC & ".DrawArrBlock", _
                        anchorName & ":rows=" & CStr(items.count) & ";room=" & CStr(room)
        NoteTruncation anchorName, items.count, room
    End If
End Sub

' 裁定書9 B6: 部屋あふれを画面へ出し、当該Stepの保存をブロックする印を立てる。
Private Sub NoteTruncation(ByVal anchorName As String, ByVal total1 As Long, ByVal room As Long)
    On Error Resume Next

    If gDrawStep >= 1 And gDrawStep <= 4 Then gTruncStep(gDrawStep) = True

    Dim one As String
    one = anchorName & ": " & CStr(total1) & "件のうち" & CStr(room) & "件しか表示できていません"
    If LenB(gTruncNote) > 0 Then gTruncNote = gTruncNote & " ／ "
    gTruncNote = gTruncNote & one

    ' 裁定書40 Q-m3: ここで直に WriteNamed すると、直前に出した 16章 E-02 の
    ' 帯(一括実行のS1)を消してしまう。併記は WriteWarnCell の1本に任せる。
    WriteWarnCell
End Sub

' 文字列配列ブロックへ書く。
Private Sub DrawScalarArr(ByVal anchorName As String, ByVal colName As String, _
                          ByVal jsonText As String, ByVal arrayKey As String)
    Dim ws As Object
    Set ws = modUISheet.BlockSheet(anchorName)
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow(anchorName)
    If headerRow <= 0 Then Exit Sub

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, U2_FIRST_COL, U2_HDR_WIDTH)

    Dim seqCol As Long
    Dim colNo As Long
    seqCol = modUISheet.ColOf(hdr, U2_FIRST_COL, "seq")
    colNo = modUISheet.ColOf(hdr, U2_FIRST_COL, colName)
    If colNo <= 0 Then Exit Sub

    Dim room As Long
    room = modUICase5.RoomOf(anchorName)
    modUISheet.ClearBlock ws, headerRow, U2_FIRST_COL, 2, room

    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(jsonText, arrayKey)

    Dim i As Long
    For i = 1 To items.count
        If i > room Then Exit For
        If seqCol > 0 Then
            modUISheet.PutText ws, headerRow + i, seqCol, CStr(i), anchorName & "/seq"
        End If
        modUISheet.PutText ws, headerRow + i, colNo, CStr(items(i)), _
                           anchorName & "/" & colName
    Next i
End Sub

' 1行ぶんを書く(順序列は行番号、他はスキーマ写像から取り出す)。
Private Sub WriteRow(ByVal ws As Object, ByVal rowNo As Long, ByVal cols As Variant, _
                     ByVal colSpec As String, ByVal itemJson As String, _
                     ByVal ordinal As Long, ByVal whereNote As String)
    Dim specs() As String
    specs = Split(colSpec, ";")

    Dim i As Long
    Dim physName As String
    Dim kindText As String
    Dim extraText As String
    Dim pathText As String
    Dim cellValue As String

    For i = LBound(specs) To UBound(specs)
        modUICaseFmt.SplitCol specs(i), physName, kindText, extraText, pathText
        If kindText = "o" Then
            cellValue = CStr(ordinal)
        Else
            cellValue = modUICaseFmt.CellFromJson(kindText, extraText, pathText, itemJson)
        End If
        modUISheet.PutText ws, rowNo, CLng(cols(i - LBound(specs))), cellValue, _
                           whereNote & "/" & physName
    Next i
End Sub


' ============================================================================
' S1～S4シートの図形ボタン(11章§5・11章§2の各ワイヤー)
' ----------------------------------------------------------------------------
' 実行系のハンドラは modUIHome2 が持つ(HOMEの[S1][S2][S3][S4]と同じ動作を
' 各シートのボタンからも起こすだけなので、実装を2箇所に置かない)。
' ============================================================================
Public Sub EnsureStepButtons()
    On Error Resume Next

    Dim ws As Object

    Set ws = modUISheet.SheetOf(SheetNameOf(1))
    modUISheet.EnsureButton ws, "btn_s1_rerun", "S1から再実行", 1, 1, 84#, "modUIHome2.HomeRunS1"
    modUISheet.EnsureButton ws, "btn_s1_next", "S2へ進む", 1, 2, 84#, "modUIHome2.HomeRunS2"

    Set ws = modUISheet.SheetOf(SheetNameOf(2))
    modUISheet.EnsureButton ws, "btn_s2_rerun", "S2から再実行", 1, 1, 84#, "modUIHome2.HomeRunS2"
    modUISheet.EnsureButton ws, "btn_s2_next", "S3へ進む", 1, 2, 84#, "modUIHome2.HomeRunS3"

    Set ws = modUISheet.SheetOf(SheetNameOf(3))
    modUISheet.EnsureButton ws, "btn_s3_rerun", "S3から再実行", 1, 1, 84#, "modUIHome2.HomeRunS3"
    modUISheet.EnsureButton ws, "btn_s3_next", "S4へ進む", 1, 2, 84#, "modUIHome2.HomeRunS4"

    Set ws = modUISheet.SheetOf(SheetNameOf(4))
    modUISheet.EnsureButton ws, "btn_s4_rerun", "S4から再実行", 1, 1, 84#, "modUIHome2.HomeRunS4"
    modUISheet.EnsureButton ws, "btn_s4_hearing", "ヒアリングシート生成", 1, 2, 128#, _
                            "modUIHome2.HomeBuildHearing"
End Sub

' S1の追加収集ブロックの各行に[コピー]ボタンを置く(11章。案件入力側と同じ動作)。
Private Sub DrawResearchButtons()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.BlockSheet("s1_research_requests")
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = modUISheet.BlockRow("s1_research_requests")
    If headerRow <= 0 Then Exit Sub

    modUISheet.DropShapesByPrefix ws, "btncopy_"

    Dim room As Long
    room = modUICase5.RoomOf("s1_research_requests")

    Dim lastRow As Long
    lastRow = modUISheet.BlockLastRow(ws, headerRow, U2_FIRST_COL, 3, room)

    Dim r As Long
    For r = headerRow + 1 To lastRow
        modUISheet.EnsureButton ws, "btncopy_" & CStr(r), "コピー", r, 4, 52#, _
                                "modUICase4.CopyResearchRow"
    Next r
End Sub
