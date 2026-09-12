Attribute VB_Name = "modUINav"
Option Explicit

' ============================================================================
' modUINav - ナビの状態とOnActionハンドラ(ui層・T-49)
' ----------------------------------------------------------------------------
' 11章v3.2 §3.1 / §3.1.1 / §3.1.2 と 13章§2.10 が正。描画は modUINavDraw。
'
' 本モジュールが持つのは次の3つだけである:
'   (1) 図形ボタンの**配置表**(区画ごとに1本の定数。tools/caption_check.py が
'       13章§2.10(f)・使い方タブの早見表・初回ツアーと逐語照合する唯一の値源)
'   (2) **いまどのSTEPか**の自動決定(11章§3.1.1 の優先順位10行。利用者は選べない)
'   (3) OnAction で配線されるハンドラ(NavPrev / NavNext / ShowDrafts / BackToNav)
'
' 原則⑥(11章§0.1): ナビの最上段は「次にすることを名指しする1行」で始まり、
'   ウィンドウ枠固定でスクロールしても消えない(枠固定はビルドが焼く)。
' ============================================================================

Private Const UN_SRC As String = "modUINav"
Private Const UN_SHEET As String = "ナビ"
Private Const UN_STEP_COUNT As Long = 6

' 画面制御用のモジュール変数(**永続でない**画面制御の状態。14章§6の
'   「状態保持の例外」への登録は要らない)。
'   gShownStep = 直前に描いたSTEP番号([次へ]の移動先を決めるのに使う)
'   gMoreOpen  = 区画①の[＋ もっと調べる]が開いているか
'   gDrawOk    = 直近の DrawNav が描き切ったか(B15: 描けなかった画面から保存しない)
Private gShownStep As Long
Private gMoreOpen As Boolean
Private gDrawOk As Boolean
'   gShownCaseId = 直前に描いた案件ID(変わったら画面の枠を入れ替える)
'   gDrewOnce    = 一度でも描いたか(起動直後の1回目は必ず枠を用意する)
Private gShownCaseId As String
Private gDrewOnce As Boolean

' ============================================================================
' フッター(裁定書26 D)。ナビと使い方の最下部に1本ずつ置く同じボタンで、
'   押すと部のポータルをブラウザで開く。
' ----------------------------------------------------------------------------
' キャプションの先頭の丸C(U+00A9)は**CP932に無い**ため、ソースへ直接書くと
'   VBEへの注入時に "?" へ化ける(tools/vba_lint.py の CP932検査が ERROR に
'   する)。lint の指示どおり ChrW() で組み立て、定数には残りの語だけを置く。
'   13章§2.10 のフッター行との逐語照合は tools/caption_check.py (G) が行う。
' ============================================================================
Private Const UN_FOOTER_TEXT As String = " リスクコンサルティング支援部"
Private Const UN_FOOTER_CHAR As Long = 169      ' U+00A9(丸C)
Private Const UN_FOOTER_NAME As String = "btn_nv_footer"
Private Const UN_FOOTER_ACTION As String = "modUINav.OpenPortal"
Private Const UN_FOOTER_WIDTH As String = "240"
' ナビの最終ブロック(区画④の最終行)と、そこから何行下へ置くか。
Private Const UN_FOOTER_ANCHOR As String = "hm_transport_banner"
Private Const UN_FOOTER_GAP_ROWS As Long = 2
Private Const UN_FOOTER_COL As Long = 2
' ポータルのURL(13章§2.3 `portal_url`)。config が空のときの既定。
Private Const UN_PORTAL_KEY As String = "portal_url"
Private Const UN_PORTAL_DEFAULT As String = _
    "http://www.portal.s1.ms-ad-ins.co.jp/loader/hp/OpenContents/" & _
    "A201203280048/toppage.html"

' ============================================================================
' 図形ボタンの配置表(13章§2.10(f))
' ----------------------------------------------------------------------------
' 1件 = "図形名;キャプション;OnAction;幅pt" を vbLf 区切り。
' **繰り返しキャプションのボタンは載せない**(区画①の[コピー]8本と区画②の18本。
'   同じ語が複数出ると逐語照合の突合が壊れるため、生成規則の側で正を持つ)。
' OnAction の修飾名は modUIHome2. である(Z-13 分割でハンドラ18本が移った)。
'   modUIHome. と書くと押しても何も起きない(最も起きやすい取り違え)。
' ============================================================================
Private Const UN_ROW_COACH As String = _
    "btn_nv_prev;← 戻る;modUINav.NavPrev;96" & vbLf & _
    "btn_nv_next;次へ →;modUINav.NavNext;96" & vbLf & _
    "btn_nv_kb;ナレッジを読み直す;modUIHome2.HomeReloadKnowledge;140" & vbLf & _
    "btn_nv_guide;使い方を開く;modUIGuide.OpenGuide;116"
' 区画①の見出し行の直下(裁定書26 C)。調べる場所への導線2本。
Private Const UN_ROW_SEC1B As String = _
    "btn_nv_dr_full;調査ページを開く;modUIResearch.OpenDrFull;168" & vbLf & _
    "btn_nv_dr_quick;クイック調査を開く;modUIResearch.OpenDrQuick;168"
Private Const UN_ROW_SEC1 As String = _
    "btn_nv_more;＋ もっと調べる（あと5本）;modUIResearch.ToggleMore;200"
Private Const UN_ROW_SEC2 As String = _
    "btn_ci_save;貼ったものを保存する;modUICase6.SaveNav;200"
Private Const UN_ROW_SEC3 As String = _
    "btn_hm_step3;まとめて作る;modUIHome2.HomeRunAll;200"
Private Const UN_ROW_SEC4 As String = _
    "btn_hm_step4;レポートを出す;modUIHome2.HomeExportHtml;200" & vbLf & _
    "btn_hm_step5;ヒアリングシートを出す;modUIHome2.HomeBuildHearing;200" & vbLf & _
    "btn_nv_drafts;下書きを見る;modUINav.ShowDrafts;200"

Public Function NavRowCoach() As String
    NavRowCoach = UN_ROW_COACH
End Function

Public Function NavRowSec1() As String
    NavRowSec1 = UN_ROW_SEC1
End Function

Public Function NavRowSec1B() As String
    NavRowSec1B = UN_ROW_SEC1B
End Function

Public Function NavRowSec2() As String
    NavRowSec2 = UN_ROW_SEC2
End Function

Public Function NavRowSec3() As String
    NavRowSec3 = UN_ROW_SEC3
End Function

Public Function NavRowSec4() As String
    NavRowSec4 = UN_ROW_SEC4
End Function

Public Function FooterCaption() As String
    FooterCaption = ChrW(UN_FOOTER_CHAR) & UN_FOOTER_TEXT
End Function

' 「図形名;キャプション;OnAction;幅pt」1本(配置表と同じ書式)。
Public Function NavRowFooter() As String
    NavRowFooter = UN_FOOTER_NAME & ";" & FooterCaption() & ";" & _
                   UN_FOOTER_ACTION & ";" & UN_FOOTER_WIDTH
End Function

' DrawNavFooter - ナビ最下部へフッターを置く(区画④の最終行の2行下)。
'   図形名は btn_nv_ 接頭辞なので、描き直しのたび DropNavShapes が落とす。
Public Sub DrawNavFooter()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UN_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = modUISheet.BlockRow(UN_FOOTER_ANCHOR)
    If r <= 0 Then Exit Sub

    Dim flds() As String
    flds = Split(NavRowFooter(), ";")
    If UBound(flds) - LBound(flds) < 3 Then Exit Sub

    modUISheet.EnsureFooterButton ws, flds(0), flds(1), r + UN_FOOTER_GAP_ROWS, _
                                  UN_FOOTER_COL, Val(flds(3)), flds(2)
End Sub

' OpenPortal - フッターのOnAction。部のポータルを既定ブラウザで開く。
'   Hyperlinks.Add は使わない(OnActionを殺す禁忌。11章§8.6)。開けなかった
'   ときはURLを逐語でトーストに出し、利用者が手で開けるようにする。
Public Sub OpenPortal()
    If Not modUIProgress.TryEnterUiLock("ポータルを開く") Then Exit Sub
    On Error GoTo Failed

    Dim url As String
    url = modConfig.GetStr(UN_PORTAL_KEY, UN_PORTAL_DEFAULT)
    If LenB(Trim$(url)) = 0 Then url = UN_PORTAL_DEFAULT

    ThisWorkbook.FollowHyperlink url
    modUIToast.ShowToast "ブラウザで開きました。", "info"
    modUIProgress.ExitUiLock
    Exit Sub

Failed:
    Err.Clear
    modUIToast.ShowToast "ブラウザで " & url & " を開いてください。", "warn"
    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' STEPの文(11章§3.1.1 の逐語表)。**7つ目を作らない**。
' ============================================================================
Public Function StepText(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 1
        StepText = "会社名と本社の場所を入れてください。" & _
                   "入れると、社内のディープリサーチに貼る文がすぐ下に出ます。"
    Case 2
        StepText = "1本目の[コピー]を押して、社内のディープリサーチに貼ってください。" & _
                   "1本ずつ・各10分です。"
    Case 3
        StepText = "返ってきた文章を、2の枠へ貼ってください。長くても分けなくて大丈夫です。"
    Case 4
        StepText = "3の[まとめて作る]を押してください。10～20分で下書きが4枚そろいます。"
    Case 5
        StepText = "4の[下書きを見る]を押して、違うところを手で直してください。"
    Case 6
        StepText = "4の[レポートを出す]を押すと、お客様に見せるレポートができます。"
    End Select
End Function

' [次へ]の移動先(11章§3.1.1 の表)。
Public Function StepAnchor(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 1, 2
        StepAnchor = "nv_sec1"
    Case 3
        StepAnchor = "nv_sec2"
    Case 4
        StepAnchor = "nv_sec3"
    Case Else
        StepAnchor = "nv_sec4"
    End Select
End Function

' ============================================================================
' 優先順位10行の純関数(11章§3.1.1・裁定書22 M4)
' ----------------------------------------------------------------------------
' CurrentStep は「シートを読む」と「10行を上から評価する」の2つを1本の中で
' やっていたため、優先順位の入れ替わりを層(a)から一度も確かめられなかった。
' **判定だけを純関数へ割り出す**(Excelを1つも触らない):
'   StepRuleOf   5つの状態 -> 当たった行の番号(1..10。どれにも当たらなければ11)
'   StepFor      同じ5つ   -> STEP番号(1..6)
'   StepActionOf 行の番号  -> hm_next_action の逐語文
' CurrentStep はシートから5つを読んで、この3本を呼ぶだけの薄い口にする。
' ============================================================================
Public Function StepRuleOf(ByVal kbReady As Boolean, ByVal locked As Boolean, _
                           ByVal hasCompany As Boolean, ByVal anyArea As Boolean, _
                           ByVal statusText As String) As Long
    If Not kbReady Then
        StepRuleOf = 1
    ElseIf locked Then
        StepRuleOf = 2
    ElseIf Not hasCompany Then
        StepRuleOf = 3
    ElseIf Not anyArea Then
        StepRuleOf = 4
    ElseIf StrComp(statusText, "error", vbBinaryCompare) = 0 Then
        StepRuleOf = 5
    ElseIf StrComp(statusText, "draft", vbBinaryCompare) = 0 Then
        StepRuleOf = 6
    ElseIf StrComp(statusText, "s1_done", vbBinaryCompare) = 0 _
           Or StrComp(statusText, "s2_done", vbBinaryCompare) = 0 _
           Or StrComp(statusText, "s3_done", vbBinaryCompare) = 0 Then
        StepRuleOf = 7
    ElseIf StrComp(statusText, "s4_done", vbBinaryCompare) = 0 Then
        StepRuleOf = 8
    ElseIf StrComp(statusText, "exported", vbBinaryCompare) = 0 Then
        StepRuleOf = 9
    ElseIf StrComp(statusText, "feedback_done", vbBinaryCompare) = 0 Then
        StepRuleOf = 10
    Else
        StepRuleOf = 11
    End If
End Function

' StepFor - 5つの状態から STEP番号(1..6)を決める(利用者は選べない)。
Public Function StepFor(ByVal kbReady As Boolean, ByVal locked As Boolean, _
                        ByVal hasCompany As Boolean, ByVal anyArea As Boolean, _
                        ByVal statusText As String) As Long
    StepFor = StepNoOfRule(StepRuleOf(kbReady, locked, hasCompany, anyArea, statusText))
End Function

' 行の番号 -> STEP番号(11章§3.1.1 の表の3列目)。
Private Function StepNoOfRule(ByVal ruleNo As Long) As Long
    Select Case ruleNo
    Case 1, 3
        StepNoOfRule = 1
    Case 4
        StepNoOfRule = 2
    Case 10
        StepNoOfRule = 3
    Case 8
        StepNoOfRule = 5
    Case 9
        StepNoOfRule = 6
    Case Else
        StepNoOfRule = 4
    End Select
End Function

' 行の番号 -> hm_next_action の逐語文(11章§3.1.1 の表の4列目)。
'   上表の文を使う行は StepText を引く(1字も別の文を作らない)。
Public Function StepActionOf(ByVal ruleNo As Long) As String
    Select Case ruleNo
    Case 1
        StepActionOf = "ナレッジブック.xlsx を、この本体と同じフォルダに置いて" & _
                       "[ナレッジを読み直す]を押してください。" & _
                       "置いてあるのに出ないときは、ファイル名が違わないかご確認ください。"
    Case 2
        StepActionOf = "いま作っています。終わるまでボタンを押さずにお待ちください。"
    Case 3
        StepActionOf = StepText(1)
    Case 4
        StepActionOf = StepText(2)
    Case 5
        StepActionOf = "前回が途中で止まりました。もう一度3の[まとめて作る]を" & _
                       "押してください。直らないときは、使い方タブの[記録を見る]を押して、" & _
                       "いちばん下の行を開発担当へ送ってください。"
    Case 7
        StepActionOf = "途中まで出来ています。もう一度3の[まとめて作る]を押すと、" & _
                       "最後まで作ります。"
    Case 8
        StepActionOf = StepText(5)
    Case 9
        StepActionOf = "4の[ヒアリングシートを出す]を押して、訪問に持っていく紙を" & _
                       "印刷してください。"
    Case 10
        StepActionOf = "訪問おつかれさまでした。聞いてきたことを2の「ヒアリング回答」へ" & _
                       "貼ると、提案が深まります。"
    Case Else
        StepActionOf = StepText(4)
    End Select
End Function

' ============================================================================
' CurrentStep - いまどのSTEPかを状態から決める(11章§3.1.1 の優先順位10行)。
' ----------------------------------------------------------------------------
' シートから5つの状態を読み、判定は上の純関数へ委ねる。**利用者はSTEPを
' 選べない**([← 戻る][次へ →]は画面をその区画へ動かすだけで、番号は書き換え
' ない)。
' ============================================================================
Public Function CurrentStep(ByRef actionText As String) As Long
    Dim caseId As String
    caseId = modUIHome.SelectedCaseId()

    Dim statusText As String
    If LenB(caseId) > 0 Then statusText = modUICase3.CaseCellText(caseId, "status")

    ' 状態は**1回ずつだけ**読む(2度読むと読んだ間に変わりうる)。
    Dim ruleNo As Long
    ruleNo = StepRuleOf(KnowledgeReady(), modUIProgress.IsUiLocked(), _
                        (LenB(modUISheet.ReadNamed("ci_company")) > 0), _
                        AnyAreaFilled(caseId), statusText)
    actionText = StepActionOf(ruleNo)
    CurrentStep = StepNoOfRule(ruleNo)
End Function

' 社内ナレッジを読めているか(11章§4.5 の3状態のうち「読めた」だけを True)。
Private Function KnowledgeReady() As Boolean
    On Error Resume Next
    KnowledgeReady = (modPipeline.KbRowCount(modKnowledge.MenusFor(vbNullString, 0)) > 0)
End Function

' ②の6欄のどれかに中身があるか。
'   現場メモだけは**枠に見出しと例文が先置きされている**ので、字数で判定すると
'   常に「入っている」になってしまう(STEPが2/6へ進まなくなる)。ひな型と同じ
'   中身なら「まだ書いていない」と数える。
Private Function AnyAreaFilled(ByVal caseId As String) As Boolean
    On Error Resume Next
    Dim keys() As String
    keys = Split(modUICase6.AreaKeys(), vbLf)
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        If StrComp(keys(i), "field_notes", vbBinaryCompare) = 0 Then
            If FieldNotesWritten() Then
                AnyAreaFilled = True
                Exit Function
            End If
        ElseIf LenB(modUICase6.AreaBody(caseId, keys(i))) > 0 Then
            AnyAreaFilled = True
            Exit Function
        End If
    Next i
End Function

' FieldNotesWritten - 現場メモの枠に、先置きのひな型より先の中身があるか。
'   ひな型と**同じ規約(SplitFieldNotes)を通してから**比べるので、例文の有無や
'   見出しの並びの違いに引きずられない(「例: 」の行はどちらでも落ちる)。
Public Function FieldNotesWritten() As Boolean
    On Error Resume Next
    Dim m1 As String, o1 As String
    Dim m2 As String, o2 As String
    modNavText.SplitFieldNotes modUICase6.ReadFieldNotesArea(), m1, o1
    modNavText.SplitFieldNotes modUINavDraw.FieldNotesTemplate(), m2, o2
    FieldNotesWritten = (StrComp(m1 & vbLf & o1, m2 & vbLf & o2, vbBinaryCompare) <> 0)
End Function

' ============================================================================
' DrawNav - ナビ1枚を描き直す(11章§3.1)。
' ----------------------------------------------------------------------------
' 骨格は notebook/src/ui/modUI.bas:619 Repaint の Resume-cleanup 方式を採った
' (失敗しても必ず ScreenUpdating=True へ到達する)。ただし DisplayGridlines など
' 「窓の設定を変える」部分は移していない(他人のブックへ影響しうるため。§8.6)。
' 描き切ったときだけ gDrawOk を立てる(B15: 描けなかった画面から保存しない)。
' ============================================================================
Public Sub DrawNav()
    gDrawOk = False
    On Error GoTo Cleanup

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UN_SHEET)
    If ws Is Nothing Then GoTo Cleanup

    Application.ScreenUpdating = False

    ' 描き直しの前に、ナビの図形の孤児を1回だけ落とす(接頭辞ごとにまとめて)。
    ' **各描画関数の中で落とさない**(同じ接頭辞を複数の描画が使うため、
    '  あとから描くものが先に描いたものを消してしまう)。
    modUINavDraw.DropNavShapes

    Dim caseId As String
    caseId = modUIHome.SelectedCaseId()

    ' 13章§2.11(e) の書き手(1): 描き切ったときだけ案件IDを書く。案件が変わって
    ' いたら、その前に画面の枠を空へ戻す(前の案件の画面がそのまま次の案件として
    ' 確定するのを防ぐ。裁定書13 W1 と同じ趣旨。**case_data は触らない**)。
    If (Not gDrewOnce) Or StrComp(caseId, gShownCaseId, vbBinaryCompare) <> 0 Then
        gShownCaseId = caseId
        gDrewOnce = True
        modUINavDraw.ResetForNewCase caseId
    End If
    modUISheet.WriteNamed "ci_case_id", vbNullString

    ' 出る条件が偽の欄を隠し、補助5本の開閉を戻す(中身は消さない)。
    modUINavDraw.ApplyAreaVisibility
    modUINavDraw.ShowMoreRows gMoreOpen

    ' 区画①の調べる文8本をその場で書き直し、[コピー]を置き直す。
    modUIResearch.BuildPrompts
    modUIResearch.EnsureCopyButtons

    ' 区画②の状態行・プレビュー・3ボタン。
    modUINavDraw.RefreshAreas caseId
    modUINavDraw.EnsureAreaButtons caseId

    ' 区画のパネルとボタン。
    modUINavDraw.DrawSections

    ' コーチ帯(最後に描く。ここまでの状態を読んでSTEPを決めるため)。
    Dim actionText As String
    Dim stepNo As Long
    stepNo = CurrentStep(actionText)
    gShownStep = stepNo
    modUINavDraw.DrawCoachBar stepNo, UN_STEP_COUNT, actionText
    modUINavDraw.MoveFocusFrame StepAnchor(stepNo)

    ' 最下部のフッター(裁定書26 D)。
    DrawNavFooter

    ' 13章§2.11(e): 案件が選ばれていなければ**新規モードの固定マーカー**を書く
    ' (空のままにすると[貼ったものを保存する]が1欄も書かずに止まり、新しい案件を
    '  作る道が無くなる)。IsValidCaseId は案件ID書式だけを通すので、この値を
    ' 有効なIDと誤認しない。
    If LenB(caseId) > 0 Then
        modUISheet.WriteNamed "ci_case_id", caseId
    Else
        modUISheet.WriteNamed "ci_case_id", modUICase3.U3_NEW_MARK
    End If

    gDrawOk = True

Cleanup:
    If Err.Number <> 0 Then
        modLog.LogError "E0101", UN_SRC & ".DrawNav", "draw_failed", Err.Number
        Err.Clear
    End If
    ' W9.2 N3: ハンドラの中(この Cleanup)で起きた失敗は同じハンドラでは受けられず
    ' 呼び出し元へ投げてしまう。画面更新の復帰と記録は**別Sub**へ切り出す。
    RestoreScreenUpdating
End Sub

' DrawNav の Cleanup: から呼ぶ後始末(W9.2 N3)。ハンドラの外なので網が張れる。
Private Sub RestoreScreenUpdating()
    On Error Resume Next
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        modLog.LogError "E0603", UN_SRC & ".DrawNav", "draw_nav_tail", Err.Number
        Err.Clear
    End If
End Sub

' 直近の DrawNav が描き切ったか(呼び出し側の保存ブロック判定に使う。B15)。
Public Function DrawOk() As Boolean
    DrawOk = gDrawOk
End Function

' 補助5本の開閉状態(modUIResearch.ToggleMore が読み書きする)。
Public Function MoreOpen() As Boolean
    MoreOpen = gMoreOpen
End Function

Public Sub SetMoreOpen(ByVal openIt As Boolean)
    gMoreOpen = openIt
End Sub

' ============================================================================
' OnAction ハンドラ(13章§2.10(f))
' ============================================================================

' [← 戻る] 1つ上の区画へ画面を動かす(STEP番号は変えない)。
Public Sub NavPrev()
    If Not modUIProgress.TryEnterUiLock("画面の移動") Then Exit Sub
    On Error GoTo Done
    GotoSection PrevAnchor()
    modUIToast.ShowToast "1つ上へ動きました。", "info"
Done:
    modUIProgress.ExitUiLock
End Sub

' [次へ →] いまのSTEPの区画へ画面を動かし、強調枠をそこへ移す。
Public Sub NavNext()
    If Not modUIProgress.TryEnterUiLock("画面の移動") Then Exit Sub
    On Error GoTo Done
    Dim anchorName As String
    anchorName = StepAnchor(gShownStep)
    GotoSection anchorName
    modUINavDraw.MoveFocusFrame anchorName
    modUIToast.ShowToast "いまやることの区画へ動きました。黄色い枠の中をご覧ください。", "info"
Done:
    modUIProgress.ExitUiLock
End Sub

' 1つ上の区画のアンカー名(いちばん上なら区画①のまま)。
Private Function PrevAnchor() As String
    Select Case StepAnchor(gShownStep)
    Case "nv_sec4"
        PrevAnchor = "nv_sec3"
    Case "nv_sec3"
        PrevAnchor = "nv_sec2"
    Case "nv_sec2"
        PrevAnchor = "nv_sec1"
    Case Else
        PrevAnchor = "nv_sec1"
    End Select
End Function

' 区画へ画面を動かす(Hyperlinks.Add は使わない。11章§7.1)。
Private Sub GotoSection(ByVal anchorName As String)
    On Error Resume Next
    Dim cell As Object
    Set cell = modUISheet.NamedCell(anchorName)
    If cell Is Nothing Then Exit Sub
    modUISheet.ShowSheet UN_SHEET
    Application.Goto cell, True
End Sub

' [下書きを見る] 下書き4枚をまとめて可視にし、①へ移る(11章§0.2・§3.4)。
Public Sub ShowDrafts()
    If Not modUIProgress.TryEnterUiLock("下書きの表示") Then Exit Sub
    On Error GoTo Done

    Dim names() As String
    names = Split(DraftSheets(), vbLf)

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim ws As Object
        Set ws = modUISheet.SheetOf(names(i))
        If Not ws Is Nothing Then
            ws.Visible = -1                      ' xlSheetVisible
            modUISheet.EnsureBackButton ws
        End If
    Next i

    modUISheet.ShowSheet names(LBound(names))
    modUIToast.ShowToast "下書きのタブを4枚出しました。" & _
                         "「AIの下書き1」から順に読んでください。", "info"
Done:
    modUIProgress.ExitUiLock
End Sub

' 下書き4枚のシート名(改名は第3弾。いまは旧名のまま)。
Private Function DraftSheets() As String
    DraftSheets = "S1_企業プロファイル" & vbLf & "S2_リスク仮説" & vbLf & _
                  "S3_提案" & vbLf & "S4_骨子"
End Function

' [ナビへ戻る] 可視にしたシートの1行目に置くボタン(11章§3.4)。行き止まりを作らない。
Public Sub BackToNav()
    If Not modUIProgress.TryEnterUiLock("ナビへ戻る") Then Exit Sub
    On Error GoTo Done
    modUISheet.ShowSheet UN_SHEET
    ' 描き直しは ExitUiLock が必ず通す(2度描かない)。
    modUIToast.ShowToast "ナビへ戻りました。", "info"
Done:
    modUIProgress.ExitUiLock
End Sub
