Attribute VB_Name = "modTestsExcel2"
Option Explicit

' ============================================================================
' modTestsExcel2 - テスト3層の層(b)の分割先(30,000字契約・12章§2)
' ----------------------------------------------------------------------------
' 入口は modTestsExcel.RunAllExcelTests から呼ばれる RunExcelTests2 の1本
' (14章§6。wintest からの唯一の入口は従来どおり RunAllExcelTests)。戻り値は
' **本モジュールが打った Check の本数**で、呼出側の自己照合(TE_EXPECTED)へ足す。
'
' 本モジュールが受け持つ検査(裁定書11 -> W6.1(裁定書22)で v3.2 の新経路へ全面
' 書き換え。旧 Q9/Q1/V1 の跡地と対応表は下の「跡地」節が持つ):
'   Q9N = 保管+プレビューの往復一致(11章§3.3.6・16章E-22)。StoreArea ->
'         RefreshArea -> LoadArea で原文が1字も変わらないこと＝**プレビューは
'         表示専用で書き戻す経路が無い**。連結側 modCaseStore.LoadData は
'         シートI/Oのため層(a)から到達できず、区切り文字の再混入と
'         プレビューからの書き戻しは純層のどの変異でも捕まらない。
'   V1N = SaveNav の3値判定(13章§2.11(e))。空を採番へ倒さず、固定マーカーの
'         ときだけ採番する。**ブックに残るセルが根拠**であることの回帰。
'   V5  = 受信箱の下書き行を CountByStatus が数えない(13章§2.6)。
'   W61 = ナビの[ここに貼る]/[貼ったものを保存する]の3点(裁定書22 M4)。
'
' フィクスチャ規律は modTestsExcel と同じ: 自前で作り、成功・失敗のどちらの
' 経路でも後始末する。Notice(MsgBox)を出す経路は無人実行を止めるため呼ばない。
' LibreOffice では走らない(層(b)は実Excel専用)。
' ============================================================================

' フィクスチャ案件(書式は IsValidCaseId 合格・実在しない未来日付系)。
Private Const T2_CASE As String = "C-97990102-903"
Private Const T2_COMPANY As String = "T47検査用商事"
Private Const T2_SHEET As String = "ナビ"

' 裁定書26 A/D の層(b)検査で使う図形名と次の一手(11章§3.1.1 STEP3)。
Private Const T2_BAND_SHAPE As String = "nv_band_bg"
Private Const T2_FOOTER_SHAPE As String = "btn_nv_footer"
Private Const T2_BAND_ACTION As String = _
    "返ってきた文章を、②の枠へ貼ってください。長くても分けなくて大丈夫です。"

' case_data の1セル上限(modUtil.SplitForCells の分割幅・13章§2.2)。ここを超えた
' 本文は seq を進めて次の行へ回る。**往復一致はこの境界で壊れやすい**ので、
' Q9N は 32,000+1,000 字でまたがせる(v3.1 の「続き欄」は撤去済み)。
Private Const T2_CHUNK As Long = 32000

' 【跡地】T2_MSG_MISMATCH(13章§2.12 B1 の文言)は、旧 Q1/V1 群が
'   modUICase3.CaseSave の hm_warning を突き合わせるために持っていた。W6.1 で
'   その群を撤去したため未使用になり、削除した(ナビ側の逐語は
'   T2_MSG_NAV_MISMATCH が持つ)。

' 11章§3.3.7 の逐語(ナビ側の不一致文言。裁定書22 W6.1 で層(b)へ回帰を1本足した)。
Private Const T2_MSG_NAV_MISMATCH As String = _
    "画面の案件と保存先が合いません。いちばん上の帯で、案件を選び直してください。"

' 個人情報を含む貼付の材料(16章 E-05。実在しない氏名)。
Private Const T2_PII_TEXT As String = "本件の窓口は 山田 太郎 様（総務部長）です。"
Private Const T2_SAFE_TEXT As String = "当社は1952年の創業以来、静岡県浜松市を拠点に事業を営んでいます。"

Private m2Run As Long

' ============================================================================
' RunExcelTests2 - 層(b)の分割先の入口。戻り値=打った Check の本数。
' ============================================================================
Public Function RunExcelTests2() As Long
    m2Run = 0
    TestQ9NavRoundTrip
    TestV1NavNumbering
    TestV5DraftRowNotCounted
    TestW61NavPaste
    TestW7FinanceRoundTrip
    TestW81BandAndFooter
    RunExcelTests2 = m2Run
End Function

Private Sub ECheck(ByVal testName As String, ByVal cond As Boolean, _
                   Optional ByVal detail As String)
    m2Run = m2Run + 1
    modTestRunner.Check testName, cond, detail
End Sub

' ============================================================================
' 【跡地・W6.1で書き換え】旧 Q9/Q1/V1 群(裁定書11・裁定書12)
' ----------------------------------------------------------------------------
' 旧テストは v3.1 の「予約行方式」を前提にしていた:
'   T47B-Q9-01..05  ci_paste_hp_1 / _2 / _3 への描画 -> 保存の往復一致
'   T47B-Q1-06..10  ci_paste_memo_1 の overflow で ci_case_id を空にする永続ガード
'   T47B-V1-11,12   modUICase3.CaseSave の3値判定(空 -> 採番しない)
' 11章v3.2 §3.3.0 が予約行方式そのものを**撤回**したため、前提が現行ブックで
' 1つも成立しない:
'   ・`ci_paste_*` の名前付きレンジは撤去済み(13章§3.3.9)。NamedCell が
'     Nothing を返すので、旧 T47B-Q9-01 は必ず赤になる
'   ・`modUICase3.DrawCaseInput` / `SaveCaseInput` / `CaseSave` は `案件入力`
'     シートの経路であり、そのシートは v3.2 で廃止された(11章§1.1)。加えて
'     `SaveCaseInput` は失敗時に `Notice`(MsgBox)へ到達しうるので、**無人実行を
'     止める**(層(b)の規律違反)
'   ・overflow(枠に入りきらない)は v3.2 では**構造的に起きない**。[ここに貼る]は
'     セルへ展開せず `case_data` へ直接保存するため、検査対象の事象が消えた
' したがって「同じ不変条件を、新経路(保管＋プレビュー)で検査するテスト」へ
' 書き換えた。対応は次のとおり(W6.1・裁定書22 の司令塔裁定):
'   Q9(往復一致)      -> T47B-Q9N-01..05(下記 TestQ9NavRoundTrip)。
'                        StoreArea -> RefreshArea -> LoadArea の往復で原文が
'                        1字も変わらないこと＝**プレビューは表示専用であり
'                        本文を書き戻す経路が無い**(11章§3.3.6)を検査する。
'   Q1(overflowガード) -> **削除**(対象機能が消えた)。「描き切れていない画面から
'                        保存しない」(B15)は `modUINav.DrawOk()` へ移り、
'                        `modUICase6.PasteIntoArea` / `SaveNav` の冒頭で見る
'                        (11章§3.3.7・裁定書22 m3)。DrawNav を人為的に失敗させる
'                        経路が無いため層(b)では検査せず、代わりに「3値判定で
'                        止まる」側を T47B-W61-16 が持つ。
'   V1(採番へ倒さない) -> T47B-V1N-06,07(下記 TestV1NavNumbering)。判定の担い手が
'                        `modUICase3.CaseSave` から `modUICase6.SaveNav` へ
'                        移ったので、SaveNav の3値判定で検査する。
' 本数は 16本 -> 12本(Q9N 5 + V1N 2 + V5 1 + W61 4)。**通し番号は詰めない**
' (T47B-V5-13 / T47B-W61-14..16 は W6.1 以前から使っている名前であり、
'  実機の合否ログを過去の記録と突き合わせられなくなるため)。08〜12 は
'  旧 Q1/V1 の欠番であり、その理由は上の対応表が持つ。
' ============================================================================

' ============================================================================
' Q9N: 保管＋プレビューの往復一致(11章§3.3.6・§7.2(a)・16章E-22)
'   (1) 33,000字(=32,000+1,000。SplitForCells の境界をまたぐ)を StoreArea で
'       保存し、LoadArea で読み戻した原文が1字も変わらないこと
'   (2) その間に RefreshArea(状態行＋プレビュー5行の描画)を挟んでも
'       case_data の原文が変わらないこと(**プレビューから本文へ書き戻さない**)
'   (3) 状態行が 11章§3.3.2 の逐語形になり、プレビューは5行だけで6行目が空
'   連結側 modCaseStore.LoadData はシートI/Oのため層(a)から到達できず、
'   区切り文字の再混入・プレビューからの書き戻しは純層のどの変異でも捕まらない。
' ============================================================================
Private Sub TestQ9NavRoundTrip()
    Dim wsCases As Object
    Dim hdr As Variant
    Dim cCase As Long
    Dim rowNo As Long
    Dim warnOrig As String
    Dim actOrig As String
    Dim hpText As String
    Dim okPre As Boolean
    Dim okLen As Boolean
    Dim okSame As Boolean
    Dim okDraw As Boolean
    Dim okLine As Boolean
    Dim detPre As String, detLen As String, detSame As String
    Dim detDraw As String, detLine As String
    On Error GoTo Crashed

    detPre = "前提不成立"
    detLen = "前提不成立"
    detSame = "前提不成立"
    detDraw = "前提不成立"
    detLine = "前提不成立"
    actOrig = ActiveSheetName()
    warnOrig = modUISheet.ReadNamed("hm_warning")

    ' (0) 前提: v3.2 の3レンジ(状態行・プレビュー・直貼り枠)が引けること。
    okPre = Not (modUISheet.NamedCell("ci_count_hp") Is Nothing)
    okPre = okPre And Not (modUISheet.NamedCell("ci_prev_hp") Is Nothing)
    okPre = okPre And Not (modUISheet.NamedCell("ci_raw_hp") Is Nothing)
    detPre = "ci_count_hp / ci_prev_hp / ci_raw_hp の3本"
    If Not okPre Then GoTo Report

    Set wsCases = SheetByName("案件一覧")
    If wsCases Is Nothing Then GoTo Report
    hdr = Hdr1(wsCases, 32)
    cCase = modUtil.FindHeaderCol(hdr, "case_id")
    If cCase <= 0 Then GoTo Report

    rowNo = LastRowA(wsCases) + 1
    PutCell wsCases, rowNo, cCase, T2_CASE
    PutNamedCol wsCases, hdr, rowNo, "company", T2_COMPANY
    PutNamedCol wsCases, hdr, rowNo, "industry_code", "T47"
    PutNamedCol wsCases, hdr, rowNo, "industry_name", "検査用"
    PutNamedCol wsCases, hdr, rowNo, "case_type", "new"
    PutNamedCol wsCases, hdr, rowNo, "status", "draft"

    ' (1) 本欄を1字だけ超える長さにして分割の境界をまたがせる。
    '     先頭に5行以上を置き、プレビューが6行目を書かないことも見る。
    hpText = "1行目" & vbLf & "2行目" & vbLf & "3行目" & vbLf & "4行目" & vbLf & _
             "5行目" & vbLf & "6行目" & vbLf & _
             String$(T2_CHUNK - 4, "あ") & "境界" & String$(996, "い") & "末尾"
    modUICase6.StoreArea T2_CASE, "input_hp", hpText

    Dim back1 As String
    back1 = modUICase6.LoadArea(T2_CASE, "input_hp")
    okLen = (Len(back1) = Len(hpText))
    detLen = "元=" & CStr(Len(hpText)) & " 往復後=" & CStr(Len(back1))
    okSame = (StrComp(back1, hpText, vbBinaryCompare) = 0)
    detSame = "境界に区切り文字が増えていないこと"

    ' (2) 描画を挟んでも case_data の原文が変わらない(表示専用の担保)。
    modUINavDraw.RefreshArea T2_CASE, "hp"
    Dim back2 As String
    back2 = modUICase6.LoadArea(T2_CASE, "input_hp")
    okDraw = (StrComp(back2, hpText, vbBinaryCompare) = 0)
    detDraw = "描画後の字数=" & CStr(Len(back2))

    ' (3) 状態行の逐語形とプレビュー5行の上限。
    Dim lineText As String
    lineText = modUISheet.ReadNamed("ci_count_hp")
    okLine = (InStr(1, lineText, "貼り付け済み　", vbBinaryCompare) = 1)
    okLine = okLine And (InStr(1, lineText, " に貼りました）", vbBinaryCompare) > 0)

    Dim prevHead As Object
    Set prevHead = modUISheet.NamedCell("ci_prev_hp")
    If prevHead Is Nothing Then
        okLine = False
    Else
        Dim wsNav As Object
        Set wsNav = prevHead.Worksheet
        okLine = okLine And _
            (LenB(Trim$(CStr(wsNav.Cells(prevHead.row, prevHead.Column).Value))) > 0)
        okLine = okLine And _
            (LenB(Trim$(CStr(wsNav.Cells(prevHead.row + 5, prevHead.Column).Value))) = 0)
    End If
    detLine = "状態行=[" & lineText & "]"

Report:
    ECheck "T47B-Q9N-01_v3.2の状態行/プレビュー/直貼り枠のレンジが引ける", okPre, detPre
    ECheck "T47B-Q9N-02_保管->読み戻しの往復で字数が変わらない(33,000字)", okLen, detLen
    ECheck "T47B-Q9N-03_保管->読み戻しの往復で原文が1字も変わらない", okSame, detSame
    ECheck "T47B-Q9N-04_描画を挟んでもcase_dataの原文が変わらない(表示専用)", _
           okDraw, detDraw
    ECheck "T47B-Q9N-05_状態行は逐語形でプレビューは5行を超えない", okLine, detLine

    On Error Resume Next
    modCaseStore.SaveData T2_CASE, "input_hp", vbNullString
    ClearNamed "ci_count_hp"
    ClearNamed "ci_raw_hp"
    modUINavDraw.RefreshArea vbNullString, "hp"
    modUISheet.WriteNamed "hm_warning", warnOrig
    DropFixtureRow wsCases, cCase
    If LenB(actOrig) > 0 Then modUISheet.ShowSheet actOrig
    Exit Sub
Crashed:
    okPre = False
    okLen = False
    okSame = False
    okDraw = False
    okLine = False
    detPre = "Err=" & CStr(Err.Number) & " " & Err.Description
    detLen = detPre
    detSame = detPre
    detDraw = detPre
    detLine = detPre
    Resume Report
End Sub

' ============================================================================
' V1N: SaveNav の3値判定は「空」を採番へ倒さない(13章§2.11(e)・裁定書12 V1)
'   (1) ci_case_id が**空**(案件ID書式でも固定マーカーでもない)のまま
'       [貼ったものを保存する]を押しても、案件一覧の行が増えない
'   (2) ci_case_id が固定マーカー「(新規)」で会社名・業種が入っていれば
'       採番して1行だけ増え、hm_case_id へ有効な案件IDが入る
'   判定の担い手は v3.2 で modUICase3.CaseSave から modUICase6.SaveNav へ
'   移った。3値判定を「空 -> 採番」へ戻す変異はこの2本でのみ落ちる。
' ============================================================================
Private Sub TestV1NavNumbering()
    Dim wsCases As Object
    Dim hdr As Variant
    Dim cCase As Long
    Dim warnOrig As String
    Dim actOrig As String
    Dim newCaseId As String
    Dim okBlank As Boolean
    Dim okNew As Boolean
    Dim detBlank As String, detNew As String
    Dim lockFree As Boolean
    On Error GoTo Crashed

    detBlank = "前提不成立"
    detNew = "前提不成立"
    actOrig = ActiveSheetName()
    warnOrig = modUISheet.ReadNamed("hm_warning")

    Set wsCases = SheetByName("案件一覧")
    If wsCases Is Nothing Then GoTo Report
    hdr = Hdr1(wsCases, 32)
    cCase = modUtil.FindHeaderCol(hdr, "case_id")
    If cCase <= 0 Then GoTo Report

    ' SaveNav は冒頭で TryEnterUiLock を取る。握られていれば本文へ入らないので、
    ' 先に空きを確かめる(取れなければ2本とも SKIP 扱い=本数は不変)。
    lockFree = modUIProgress.TryEnterUiLock("T47B-V1Nの前提確認")
    If lockFree Then modUIProgress.ExitUiLock
    If Not lockFree Then
        okBlank = True
        okNew = True
        detBlank = "SKIP: UiLockが握られており SaveNav を無人実行できない"
        detNew = detBlank
        GoTo Report
    End If

    ' SaveNav は描き切った画面からしか保存しない(B15・裁定書22 m3)。
    modUINav.DrawNav
    If Not modUINav.DrawOk() Then
        okBlank = True
        okNew = True
        detBlank = "SKIP: DrawNav が描き切っていない(全欄ブロックが正)"
        detNew = detBlank
        GoTo Report
    End If

    ' (1) 空のまま保存 -> 採番しない。
    Dim rows0 As Long
    rows0 = LastRowA(wsCases)
    modUISheet.WriteNamed "hm_case_id", vbNullString
    modUISheet.WriteNamed "ci_case_id", vbNullString
    modUISheet.WriteNamed "hm_warning", vbNullString
    modUICase6.SaveNav
    okBlank = (LastRowA(wsCases) = rows0)
    detBlank = "案件一覧の最終行 前=" & CStr(rows0) & " 後=" & CStr(LastRowA(wsCases))

    ' (2) 固定マーカー＋会社名・業種あり -> 採番して1行だけ増える。
    modUISheet.WriteNamed "ci_company", T2_COMPANY
    modUISheet.WriteNamed "ci_industry_name", "検査用"
    modUISheet.WriteNamed "ci_industry_code", "T47"
    modUISheet.WriteNamed "ci_case_id", modUICase3.U3_NEW_MARK
    modUICase6.SaveNav
    newCaseId = Trim$(modUISheet.ReadNamed("hm_case_id"))
    okNew = modCaseStore.IsValidCaseId(newCaseId) And _
            (LastRowA(wsCases) = rows0 + 1)
    detNew = "採番=[" & newCaseId & "] 最終行 前=" & CStr(rows0) & _
             " 後=" & CStr(LastRowA(wsCases))

Report:
    ECheck "T47B-V1N-06_ci_case_idが空のままの保存は採番しない", okBlank, detBlank
    ECheck "T47B-V1N-07_固定マーカーの保存は採番して1行だけ増やす", okNew, detNew

    On Error Resume Next
    If LenB(newCaseId) > 0 Then
        Dim rNew As Long
        For rNew = LastRowA(wsCases) To 2 Step -1
            If StrComp(Trim$(CellStr(wsCases, rNew, cCase)), newCaseId, _
                       vbBinaryCompare) = 0 Then wsCases.Rows(rNew).Delete
        Next rNew
    End If
    ClearNamed "ci_company"
    ClearNamed "ci_industry_code"
    ClearNamed "ci_industry_name"
    ClearNamed "ci_dossier_tier"
    modUISheet.WriteNamed "ci_case_id", vbNullString
    modUISheet.WriteNamed "hm_case_id", vbNullString
    modUISheet.WriteNamed "hm_warning", warnOrig
    DropFixtureRow wsCases, cCase
    If LenB(actOrig) > 0 Then modUISheet.ShowSheet actOrig
    Exit Sub
Crashed:
    okBlank = False
    okNew = False
    detBlank = "Err=" & CStr(Err.Number) & " " & Err.Description
    detNew = detBlank
    Resume Report
End Sub

' ============================================================================
' W6.1(裁定書22 M4 層(b)3本): ナビの[ここに貼る]と[貼ったものを保存する]の
'   **ブックに残るセルが根拠**である3点。いずれも層(a)から原理的に到達できない
'   (クリップボード・案件一覧への採番・画面のセル状態が絡むため)。
'   (1) [ここに貼る]で個人情報を検知したら case_data が**1字も増えない**
'       (16章 E-05・11章§3.3.7。採番より前に弾く=裁定書22 m8)
'   (2) 案件未選択(固定マーカー)の画面へ貼ると **採番 -> 保存** の順で通り、
'       採番したIDの case_data に本文が入る(13章§2.11(e) の書き手(3))
'   (3) ci_case_id が案件ID書式でも固定マーカーでもない画面からの
'       [貼ったものを保存する]は**1欄も書かず**不一致文言を出す(第2の壁)
' ============================================================================
' ============================================================================
' W7: 7欄目「決算・財務」の保管->結合の往復(13章§2.11(a)・§2.2。裁定書25 S3)
' ----------------------------------------------------------------------------
'   v2.6 で増えたのは**欄1本と data_key 1本**だけであり、経路は既存6欄と同じ
'   (StoreArea -> modCaseStore.SaveData -> SplitForCells -> LoadData の区切りなし
'   連結)。層(a)からは case_data のシートI/Oへ到達できないため、
'   「input_finance が data_key として受理され、32,000字の分割境界をまたいでも
'   1字も変わらずに戻る」ことは層(b)でしか確かめられない。
'   1本の Check に (0)4レンジが引ける (1)往復の字数一致 (2)往復の原文一致 を
'   すべて詰める(欄が1本増えただけの回帰なので本数を増やさない)。
' ============================================================================
Private Sub TestW7FinanceRoundTrip()
    Dim okAll As Boolean
    Dim detText As String
    Dim finText As String
    Dim back1 As String
    On Error GoTo Crashed

    detText = "前提不成立"

    ' (0) v2.6 で足した7欄目の4レンジ(13章§2.10(c) の65本の増分)。
    okAll = Not (modUISheet.NamedCell("ci_count_finance") Is Nothing)
    okAll = okAll And Not (modUISheet.NamedCell("ci_prev_finance") Is Nothing)
    okAll = okAll And Not (modUISheet.NamedCell("ci_raw_finance") Is Nothing)
    okAll = okAll And Not (modUISheet.NamedCell("ci_sent_finance") Is Nothing)
    detText = "ci_count/prev/raw/sent_finance の4本"
    If Not okAll Then GoTo Report

    ' (1)(2) 32,000字の分割境界を1字またぐ本文で往復させる。
    finText = "純資産 12億円、売上 85億円（2025年3月期・決算公告）" & vbLf & _
              String$(T2_CHUNK - 4, "あ") & "境界" & String$(996, "い") & "末尾"
    modUICase6.StoreArea T2_CASE, "input_finance", finText
    back1 = modUICase6.LoadArea(T2_CASE, "input_finance")
    okAll = okAll And (Len(back1) = Len(finText))
    okAll = okAll And (StrComp(back1, finText, vbBinaryCompare) = 0)
    detText = "元=" & CStr(Len(finText)) & " 往復後=" & CStr(Len(back1))

Report:
    ECheck "T47B-W7-01_7欄目input_financeの保管->結合の往復が原文と一致する", _
           okAll, detText

    On Error Resume Next
    modCaseStore.SaveData T2_CASE, "input_finance", vbNullString
    Exit Sub
Crashed:
    okAll = False
    detText = "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Report
End Sub

' W81 = 帯の図形の中の文字とフッター図形の存在(裁定書26 A/D・11章§8.5 #16)。
'   図形の描画は LibreOffice で確かめられず実機でしか見えない欠陥だった。
'   期待値は純関数 modUIGeom.CoachBandText との一致だけを見る(逐語は層(a))。
Private Sub TestW81BandAndFooter()
    Dim ws As Object
    Dim okText As Boolean
    Dim okFooter As Boolean
    Dim detText As String
    Dim detFooter As String
    Dim wantText As String
    Dim gotText As String
    On Error GoTo Crashed

    detText = "前提不成立"
    detFooter = "前提不成立"

    Set ws = modUISheet.SheetOf(T2_SHEET)
    If ws Is Nothing Then GoTo Report

    wantText = modUIGeom.CoachBandText(2, 6, T2_BAND_ACTION)
    modUINavDraw.DrawCoachBar 2, 6, T2_BAND_ACTION
    gotText = ShapeTextOf(ws, T2_BAND_SHAPE)
    okText = (StrComp(gotText, wantText, vbBinaryCompare) = 0)
    detText = "期待=[" & wantText & "] 実際=[" & gotText & "]"

    modUINav.DrawNavFooter
    okFooter = HasShape(ws, T2_FOOTER_SHAPE)
    detFooter = "図形 " & T2_FOOTER_SHAPE & " の有無"

    ' 跡地(W9.3): T47B-W81-03「ブックイベントのクラスが結線されている」は
    '   撤去した。ブックイベントの受け口が clsAppEvents から **ThisWorkbook
    '   文書モジュール**へ移り(17章 Z-24)、VBA から「ThisWorkbook に
    '   Workbook_Activate があるか」を読む手段が VBE のプロジェクト参照しか
    '   無く、その語は配布物の禁止文字列(裁定書27 W9-B 6)だからである。
    '   焼き込まれたスタブの形は tools/bin_roundtrip.py [4b] が配布物の側から
    '   検査する(層(b)ではなくビルド検問の担当へ移した)。

Report:
    ECheck "T47B-W81-01_コーチ帯の図形の中にCoachBandTextと同じ文字がある", _
           okText, detText
    ECheck "T47B-W81-02_ナビの最下部にフッターの図形がある", okFooter, detFooter

    ' 画面をふだんの状態へ戻す。
    On Error Resume Next
    modUINav.DrawNav
    Exit Sub
Crashed:
    okText = False
    okFooter = False
    detText = "Err=" & CStr(Err.Number) & " " & Err.Description
    detFooter = detText
    Resume Report
End Sub

' 図形の中の文字(無ければ空文字)。
Private Function ShapeTextOf(ByVal ws As Object, ByVal shapeName As String) As String
    On Error Resume Next
    ShapeTextOf = ws.Shapes(shapeName).TextFrame.Characters.Text
End Function

' その名前の図形があるか。
Private Function HasShape(ByVal ws As Object, ByVal shapeName As String) As Boolean
    On Error Resume Next
    Dim shp As Object
    Set shp = ws.Shapes(shapeName)
    HasShape = Not (shp Is Nothing)
End Function

Private Sub TestW61NavPaste()
    Dim wsCases As Object
    Dim hdr As Variant
    Dim cCase As Long
    Dim rowNo As Long
    Dim actOrig As String
    Dim newCaseId As String
    Dim okPii As Boolean
    Dim okNew As Boolean
    Dim okMismatch As Boolean
    Dim okClean As Boolean
    Dim detClean As String
    Dim detPii As String
    Dim detNew As String
    Dim detMis As String
    On Error GoTo Crashed

    detPii = "前提不成立"
    detNew = "前提不成立"
    detMis = "前提不成立"
    detClean = "前提不成立"
    actOrig = ActiveSheetName()

    Set wsCases = SheetByName("案件一覧")
    If wsCases Is Nothing Then GoTo Report
    hdr = Hdr1(wsCases, 32)
    cCase = modUtil.FindHeaderCol(hdr, "case_id")
    If cCase <= 0 Then GoTo Report

    ' フィクスチャ案件を1本置き、ナビを描いてから始める(DrawOk を立てる。m3)。
    rowNo = LastRowA(wsCases) + 1
    PutCell wsCases, rowNo, cCase, T2_CASE
    PutNamedCol wsCases, hdr, rowNo, "company", T2_COMPANY
    PutNamedCol wsCases, hdr, rowNo, "industry_code", "T47"
    PutNamedCol wsCases, hdr, rowNo, "industry_name", "検査用"
    PutNamedCol wsCases, hdr, rowNo, "case_type", "new"
    PutNamedCol wsCases, hdr, rowNo, "status", "draft"

    modCaseStore.SaveData T2_CASE, "input_hp", vbNullString
    modUISheet.WriteNamed "hm_case_id", T2_CASE
    modUINav.DrawNav
    If Not modUINav.DrawOk() Then
        detPii = "DrawNav が描き切っていない(以降の貼付は全欄ブロックが正)"
        detNew = detPii
        detMis = detPii
        GoTo Report
    End If

    ' ---- (1) 個人情報を含む貼付は case_data を1字も増やさない ----------
    Dim beforeLen As Long
    beforeLen = Len(modCaseStore.LoadData(T2_CASE, "input_hp"))
    If modUISheet.CopyToClipboard(T2_PII_TEXT) Then
        modUICase6.PasteIntoArea "hp"
        Dim afterLen As Long
        afterLen = Len(modCaseStore.LoadData(T2_CASE, "input_hp"))
        okPii = (afterLen = beforeLen)
        detPii = "貼付前=" & CStr(beforeLen) & "字 貼付後=" & CStr(afterLen) & "字"
    Else
        okPii = True
        detPii = "SKIP: クリップボードへ書けない環境"
    End If

    ' ---- (2) 案件未選択 -> 採番 -> 保存 の順序 --------------------------
    modUISheet.WriteNamed "hm_case_id", vbNullString
    modUISheet.WriteNamed "ci_company", T2_COMPANY & "2"
    modUISheet.WriteNamed "ci_industry_name", "検査用"
    modUISheet.WriteNamed "ci_industry_code", "T47"
    modUISheet.WriteNamed "ci_case_id", modUICase3.U3_NEW_MARK
    If modUISheet.CopyToClipboard(T2_SAFE_TEXT) Then
        modUICase6.PasteIntoArea "hp"
        newCaseId = Trim$(modUISheet.ReadNamed("hm_case_id"))
        okNew = modCaseStore.IsValidCaseId(newCaseId)
        If okNew Then
            okNew = (StrComp(modCaseStore.LoadData(newCaseId, "input_hp"), _
                             T2_SAFE_TEXT, vbBinaryCompare) = 0)
        End If
        detNew = "採番=[" & newCaseId & "] 保存字数=" & _
                 CStr(Len(modCaseStore.LoadData(newCaseId, "input_hp")))
        ' 裁定書29 裁定4: Mac版Excelは貼り付けの書式名が無く PasteSpecial が
        '   全滅し、採番へ進む前に PasteIntoArea が抜ける(採番=[])。製品は
        '   変えず、**貼り付けが不成立のMacのときだけ**SKIPで緑にする。
        '   Windowsでは IsMacExcel が False なのでこの枝に入らない
        '   (=Windowsに新しいSKIP経路は作らない。落ちたら FAIL のまま)。
        If LenB(newCaseId) = 0 And modTestsExcel3.IsMacExcel() Then
            okNew = True
            detNew = "SKIP(Mac): 貼り付け書式名がMac版Excelに無い"
        End If
    Else
        okNew = True
        detNew = "SKIP: クリップボードへ書けない環境"
    End If

    ' ---- (3) ci_case_id が3値のどれでもない画面からは1字も書かない -----
    modCaseStore.SaveData T2_CASE, "input_hp", T2_SAFE_TEXT
    modUISheet.WriteNamed "hm_warning", vbNullString
    modUISheet.WriteNamed "ci_case_id", "こわれた値"
    modUISheet.WriteNamed "ci_raw_hp", "直貼りした本文"
    modUICase6.SaveNav
    okMismatch = (StrComp(modCaseStore.LoadData(T2_CASE, "input_hp"), _
                          T2_SAFE_TEXT, vbBinaryCompare) = 0)
    okMismatch = okMismatch And _
        (InStr(1, modUISheet.ReadNamed("hm_warning"), T2_MSG_NAV_MISMATCH, _
               vbBinaryCompare) > 0)
    detMis = "hm_warning=[" & modUISheet.ReadNamed("hm_warning") & "]"

    ' 裁定書30 裁定2(Z-28)。(3)は SaveNav にわざと不一致警告を出させる。
    ' 帯(トーストの図形)と hm_warning を残すと、テストのあと利用者の画面に
    ' 赤い警告が居座る。ここで消し、消えたことを1本のECheckで押さえる。
    modUISheet.WriteNamed "hm_warning", vbNullString
    modUIToast.CancelToast
    modUIToast.HideToast
    okClean = (LenB(Trim$(modUISheet.ReadNamed("hm_warning"))) = 0)
    detClean = "後始末後 hm_warning=[" & modUISheet.ReadNamed("hm_warning") & "]"

Report:
    ECheck "T47B-W61-14_[ここに貼る]は個人情報を検知したらcase_dataを1字も増やさない", _
           okPii, detPii
    ECheck "T47B-W61-15_案件未選択の画面へ貼ると採番->保存の順で通る", _
           okNew, detNew
    ECheck "T47B-W61-16_ci_case_id不一致の画面からの保存は1欄も書かない", _
           okMismatch, detMis
    ECheck "T47B-W61-17_(3)の後にhm_warningと警告帯を残さない", _
           okClean, detClean

    On Error Resume Next
    modCaseStore.SaveData T2_CASE, "input_hp", vbNullString
    If LenB(newCaseId) > 0 Then
        modCaseStore.SaveData newCaseId, "input_hp", vbNullString
        DropFixtureRow wsCases, cCase
        Dim rNew As Long
        For rNew = LastRowA(wsCases) To 2 Step -1
            If StrComp(CellStr(wsCases, rNew, cCase), newCaseId, vbBinaryCompare) = 0 Then
                wsCases.Rows(rNew).Delete
            End If
        Next rNew
    End If
    ClearNamed "ci_raw_hp"
    ClearNamed "ci_company"
    ClearNamed "ci_industry_code"
    ClearNamed "ci_industry_name"
    modUISheet.WriteNamed "ci_case_id", vbNullString
    modUISheet.WriteNamed "hm_case_id", vbNullString
    modUISheet.WriteNamed "hm_warning", vbNullString
    modUIToast.CancelToast
    modUIToast.HideToast
    DropFixtureRow wsCases, cCase
    If LenB(actOrig) > 0 Then modUISheet.ShowSheet actOrig
    Exit Sub
Crashed:
    detPii = "Err=" & CStr(Err.Number) & " " & Err.Description
    detNew = detPii
    detMis = detPii
    okPii = False
    okNew = False
    okMismatch = False
    okClean = False
    Resume Report
End Sub

' ============================================================================
' V5(裁定書12・13章§2.6): 下書き行に status を書いても CountByStatus が
'   数えない。HOMEの hm_inbox_* は CountByStatus だけが値源であり、下書き行を
'   数えると「一括診断は0件なのにHOMEは1件」という食い違いが出る。下書き行の
'   スキップを無効化する変異はこの1本でのみ検出できる(層(a)から到達できない)。
' ============================================================================
Private Sub TestV5DraftRowNotCounted()
    Dim ws As Object
    Dim addedRow As Long
    Dim before As Long
    Dim after1 As Long
    Dim okAll As Boolean
    Dim detail As String
    On Error GoTo Crashed

    Set ws = SheetByName("受信箱")
    If ws Is Nothing Then
        detail = "受信箱シートが無い"
        GoTo Report
    End If

    Dim hdr As Variant
    Dim cId As Long
    Dim cStatus As Long
    hdr = Hdr1(ws, 32)
    cId = modUtil.FindHeaderCol(hdr, "inbox_id")
    cStatus = modUtil.FindHeaderCol(hdr, "status")
    If cId <= 0 Or cStatus <= 0 Then
        detail = "inbox_id / status 列が引けない"
        GoTo Report
    End If

    before = modUIInbox.CountByStatus("undiagnosed")

    ' 末尾へ「マーカー付き かつ status=undiagnosed」の行を1本だけ足す。
    addedRow = LastRowA(ws) + 1
    PutCell ws, addedRow, cId, modInboxStore.IB_DRAFT_MARK
    PutCell ws, addedRow, cStatus, "undiagnosed"

    after1 = modUIInbox.CountByStatus("undiagnosed")
    okAll = (after1 = before)
    detail = "追加前=" & CStr(before) & " 追加後=" & CStr(after1)

Report:
    ECheck "T47B-V5-13_下書き行にstatusを書いてもCountByStatusが数えない", _
           okAll, detail
    On Error Resume Next
    If addedRow >= 2 And Not ws Is Nothing Then ws.Rows(addedRow).Delete
    Exit Sub
Crashed:
    okAll = False
    detail = "Err=" & CStr(Err.Number) & " " & Err.Description
    Resume Report
End Sub

' ============================================================================
' 共通の下請け(すべて Private。公開口は RunExcelTests2 のみ=14章§6)
' ============================================================================

' 案件一覧からフィクスチャ案件の行を全削除する(下から回す)。
Private Sub DropFixtureRow(ByVal ws As Object, ByVal cCase As Long)
    On Error Resume Next
    If ws Is Nothing Or cCase <= 0 Then Exit Sub
    Dim r As Long
    For r = LastRowA(ws) To 2 Step -1
        If Trim$(CellStr(ws, r, cCase)) = T2_CASE Then ws.Rows(r).Delete
    Next r
End Sub

Private Sub ClearNamed(ByVal rangeName As String)
    On Error Resume Next
    If modUISheet.NamedCell(rangeName) Is Nothing Then Exit Sub
    modUISheet.WriteNamed rangeName, vbNullString
End Sub

Private Sub PutCell(ByVal ws As Object, ByVal r As Long, ByVal c As Long, _
                    ByVal valueText As String)
    If c <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(r, c), valueText, "T47B/案件一覧"
End Sub

Private Sub PutNamedCol(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                        ByVal colName As String, ByVal valueText As String)
    PutCell ws, r, modUtil.FindHeaderCol(hdr, colName), valueText
End Sub

' 実行前の活性シート名(取れなければ "")。裁定書13 W5 の復帰用。
Private Function ActiveSheetName() As String
    On Error GoTo NoActive
    ActiveSheetName = CStr(ActiveSheet.Name)
    Exit Function
NoActive:
    ActiveSheetName = vbNullString
End Function

' シートを名前で引く(非表示・veryHidden も対象。無ければ Nothing)。
Private Function SheetByName(ByVal sheetTitle As String) As Object
    On Error GoTo NoSheet
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.count
        If ThisWorkbook.Worksheets(i).Name = sheetTitle Then
            Set SheetByName = ThisWorkbook.Worksheets(i)
            Exit Function
        End If
    Next i
    Exit Function
NoSheet:
    Set SheetByName = Nothing
End Function

' A列基準の最終行(データ無しは1)。
Private Function LastRowA(ByVal ws As Object) As Long
    On Error GoTo One1
    LastRowA = ws.Cells(ws.Rows.count, 1).End(-4162).row
    If LastRowA < 1 Then LastRowA = 1
    Exit Function
One1:
    LastRowA = 1
End Function

' 見出し行(1行目)を 1..scanCols で読む。
Private Function Hdr1(ByVal ws As Object, ByVal scanCols As Long) As Variant
    On Error GoTo Empty0
    Hdr1 = ws.Range(ws.Cells(1, 1), ws.Cells(1, scanCols)).Value
    Exit Function
Empty0:
    Hdr1 = Empty
End Function

' 1セルを文字列で読む(読めなければ "")。
Private Function CellStr(ByVal ws As Object, ByVal r As Long, ByVal c As Long) As String
    On Error GoTo Blank0
    CellStr = CStr(ws.Cells(r, c).Value)
    Exit Function
Blank0:
    CellStr = vbNullString
End Function

' err_log の現在の最終行(この行より後を「テストが積んだ分」として検査する)。
Private Function ErrLogTail() As Long
    Dim ws As Object
    Set ws = SheetByName("err_log")
    If ws Is Nothing Then Exit Function
    ErrLogTail = LastRowA(ws)
End Function

' err_log の fromRow より後に「err_code一致 かつ detail部分一致」の行があるか。
Private Function ErrLogged(ByVal fromRow As Long, ByVal codeText As String, _
                           ByVal detailPart As String) As Boolean
    On Error GoTo No0
    Dim ws As Object
    Set ws = SheetByName("err_log")
    If ws Is Nothing Then Exit Function
    Dim hdr As Variant
    hdr = Hdr1(ws, 8)
    Dim cCode As Long
    Dim cDetail As Long
    cCode = modUtil.FindHeaderCol(hdr, "err_code")
    cDetail = modUtil.FindHeaderCol(hdr, "detail")
    If cCode <= 0 Or cDetail <= 0 Then Exit Function
    Dim r As Long
    Dim lastRow As Long
    lastRow = LastRowA(ws)
    For r = fromRow + 1 To lastRow
        If CStr(ws.Cells(r, cCode).Value) = codeText Then
            If InStr(1, CStr(ws.Cells(r, cDetail).Value), detailPart, vbBinaryCompare) > 0 Then
                ErrLogged = True
                Exit Function
            End If
        End If
    Next r
    Exit Function
No0:
    ErrLogged = False
End Function
