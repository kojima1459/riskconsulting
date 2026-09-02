Attribute VB_Name = "modUIResearch"
Option Explicit

' ============================================================================
' modUIResearch - ナビ区画①「会社のこと」の調べる文8本(ui層・T-49)
' ----------------------------------------------------------------------------
' 11章v3.2 §3.2 と 13章§2.19 が正。
'
' 値の在処は1箇所(11章§3.2・13章§2.18):
'   8本の雛形は **VBAソースに埋め込まない**。使い方タブの非表示行
'   (`gd_prompt_01`～`gd_prompt_08`。ビルドが docs/08 から逐語で焼く)を読んで
'   `{{ }}` を置換する。これにより「使い方タブに書いてある文」と「ナビが出す文」
'   が**構造的に同一**になる(2箇所で違う文を出さない)。
'
' [文を作る]ボタンは置かない。画面を描いたとき(DrawNav)に書き直すだけにする
'   (常駐タイマーは作らない。11章§7.2(d))。
' ============================================================================

Private Const UR_SRC As String = "modUIResearch"
Private Const UR_SHEET As String = "ナビ"
Private Const UR_LOCK_COPY As String = "調べる文のコピー"
Private Const UR_LOCK_MORE As String = "もっと調べるの開閉"

' 次に押すべき枠の太枠の色(RGB(255,184,0) 橙。強調枠と同じ色で意味をそろえる)。
Private Const UR_COLOR_NEXT As Long = 47359&

' 標準3本(常に見えている)と補助5本([＋ もっと調べる]で開く)。
Private Const UR_STD_COUNT As Long = 3
Private Const UR_ALL_COUNT As Long = 8

' 11章§2.2 #3 の逐語(失敗の2文)。
Private Const UR_MSG_COPY_NG As String = _
    "文を写せませんでした。枠の中の文をマウスで選んで、Ctrl+C で写してください。"
Private Const UR_MSG_NO_COMPANY As String = _
    "先に①の会社名を入れてください。会社名が空のままでは、調べる文を写せません。"

' 8本の見出し(11章§3.2 の逐語)。
Private Function TitleOfPrompt(ByVal n As Long) As String
    Select Case n
    Case 1
        TitleOfPrompt = "1本目 会社の基本"
    Case 2
        TitleOfPrompt = "2本目 リスクの兆候"
    Case 3
        TitleOfPrompt = "3本目 調達・仕入れの構造"
    Case 4
        TitleOfPrompt = "業界と競合"
    Case 5
        TitleOfPrompt = "世の中の動きとの関係"
    Case 6
        TitleOfPrompt = "前回の更新からの変化"
    Case 7
        TitleOfPrompt = "拠点の災害リスク"
    Case 8
        TitleOfPrompt = "決算のハイライト"
    End Select
End Function

' ============================================================================
' BuildPrompts - 8本の本文を組み立てて dr_body_01..08 へ書く(書込専用)。
'   埋まらない `{{ }}` は必ず `〔ここに～を書いてください〕` の日本語の穴へ
'   置き換える(利用者は `{{ }}` の意味を知らない。11章§3.2)。
' ============================================================================
Public Sub BuildPrompts()
    On Error Resume Next

    Dim company As String
    Dim address As String
    Dim industry As String
    Dim secCode As String
    Dim sites As String
    company = modUISheet.ReadNamed("ci_company")
    address = modUISheet.ReadNamed("dr_address")
    industry = modUISheet.ReadNamed("ci_industry_name")
    secCode = modUISheet.ReadNamed("dr_sec_code")
    sites = modUISheet.ReadNamed("dr_sites")

    Dim n As Long
    For n = 1 To UR_ALL_COUNT
        modUISheet.WriteNamed BodyRangeOf(n), _
            FillTemplate(TemplateOf(n), company, address, industry, secCode, sites)
    Next n
End Sub

' 雛形の在処(使い方タブの非表示行)。読めなければ空。
Private Function TemplateOf(ByVal n As Long) As String
    TemplateOf = modUISheet.ReadNamed(PromptRangeOf(n))
End Function

Private Function PromptRangeOf(ByVal n As Long) As String
    PromptRangeOf = "gd_prompt_" & Format$(n, "00")
End Function

Private Function BodyRangeOf(ByVal n As Long) As String
    BodyRangeOf = "dr_body_" & Format$(n, "00")
End Function

' `{{ }}` の差し込みと、埋まらない穴の日本語化(11章§3.2 の表が正)。
Public Function FillTemplate(ByVal tpl As String, ByVal company As String, _
                             ByVal address As String, ByVal industry As String, _
                             ByVal secCode As String, ByVal sites As String) As String
    Dim s As String
    s = tpl
    If LenB(s) = 0 Then Exit Function

    Dim placeText As String
    placeText = address
    If LenB(industry) > 0 Then
        If LenB(placeText) > 0 Then
            placeText = placeText & "・" & industry
        Else
            placeText = industry
        End If
    End If

    s = Replace$(s, "{{企業名}}", company)
    s = Replace$(s, "{{本社所在地・業種}}", placeText)
    s = Replace$(s, "{{本社所在地}}", address)
    s = Replace$(s, "{{業界名}}", industry)
    s = Replace$(s, "{{業種名}}", industry)
    s = Replace$(s, "{{コード}}", secCode)
    s = Replace$(s, "{{拠点リスト（名称・住所）}}", HoleOr(sites, "拠点の名前と住所"))
    s = Replace$(s, "{{リスク/課題}}", Hole("気になっていること"))
    s = Replace$(s, "{{前回更新からの期間}}", Hole("前回の更新からの期間"))
    s = Replace$(s, "{{直近決算期}}", Hole("決算期を書いてください（例: 2026年3月期）"))
    s = Replace$(s, "{{公式ドメイン}}", Hole("会社の公式サイトのURL"))
    FillTemplate = s
End Function

' 日本語の穴。`{{ }}` を画面に出さないための置換文。
Private Function Hole(ByVal what As String) As String
    If InStr(1, what, "ください", vbBinaryCompare) > 0 Then
        Hole = "〔ここに" & what & "〕"
    Else
        Hole = "〔ここに" & what & "を書いてください〕"
    End If
End Function

' 値があればそれを、無ければ日本語の穴を返す。
Private Function HoleOr(ByVal valueText As String, ByVal what As String) As String
    If LenB(Trim$(valueText)) > 0 Then
        HoleOr = valueText
    Else
        HoleOr = Hole(what)
    End If
End Function

' ============================================================================
' EnsureCopyButtons - 8本の[コピー](11章§3.1.2)。
'   閉じている間の補助5本には図形を作らない(隠れた図形を残さない)。
'   ボタンは小さめ(modUISheet.EnsureButtonSm)。説明は AlternativeText へ入れる。
' ============================================================================
Public Sub EnsureCopyButtons()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(UR_SHEET)
    If ws Is Nothing Then Exit Sub

    ' 孤児の一括削除は modUINavDraw.DropNavShapes が描画の最初に行う。ここでは
    ' 閉じているときに余る4本目以降だけを落とす(隠れた図形を残さない)。
    Dim n2 As Long
    For n2 = 1 To UR_ALL_COUNT
        modUISheet.DropShape ws, "btn_nv_copy" & CStr(n2)
    Next n2

    Dim last As Long
    last = UR_STD_COUNT
    If modUINav.MoreOpen() Then last = UR_ALL_COUNT

    Dim n As Long
    For n = 1 To last
        Dim rowNo As Long
        rowNo = modUISheet.BlockRow(BodyRangeOf(n))
        If rowNo > 0 Then
            modUISheet.EnsureButtonSm ws, "btn_nv_copy" & CStr(n), "コピー", rowNo, 4, _
                                      72#, "modUIResearch.CopyPrompt" & CStr(n), _
                                      TitleOfPrompt(n) & "を写します"
        End If
    Next n
End Sub

' ============================================================================
' CopyPromptN - 1本ぶんを写す(11章§3.2)。
'   経路は既存の modUICase4.CopyResearchRow と同系(DataObject の遅延バインド)。
'   押した本数を覚えて、次に押すべき枠を dr_copied_seq に出す(config へ書かない
'   =ブックを閉じれば消える。これが「1本ずつ・直列」を体で覚えさせる仕掛け)。
' ============================================================================
Public Sub CopyPrompt(ByVal n As Long)
    On Error GoTo Failed

    If LenB(modUISheet.ReadNamed("ci_company")) = 0 Then
        modUIToast.ShowToast UR_MSG_NO_COMPANY, "warn"
        Exit Sub
    End If

    Dim body As String
    body = modUISheet.ReadNamed(BodyRangeOf(n))
    If LenB(body) = 0 Then
        modUIToast.ShowToast UR_MSG_COPY_NG, "error"
        Exit Sub
    End If

    If Not modUISheet.CopyToClipboard(body) Then
        modUIToast.ShowToast UR_MSG_COPY_NG, "error"
        Exit Sub
    End If

    modUISheet.WriteNamed "dr_copied_seq", "つぎは " & CStr(n + 1) & "本目です"
    MarkNextFrame n + 1
    modUIToast.ShowToast CopiedTextOf(n), "info"
    Exit Sub

Failed:
    modLog.LogError "E0603", UR_SRC & ".CopyPrompt", "copy_failed:" & CStr(n), Err.Number
    modUIToast.ShowToast UR_MSG_COPY_NG, "error"
End Sub

' 写したあとのトースト(11章§2.2 #3-#5 の逐語)。
Private Function CopiedTextOf(ByVal n As Long) As String
    Select Case n
    Case 1
        CopiedTextOf = "1本目（会社の基本）を写しました。" & _
                       "社内のディープリサーチに貼って投げてください。" & _
                       "返ってくるまで10分ほどかかります。"
    Case 2
        CopiedTextOf = "2本目（リスクの兆候）を写しました。" & _
                       "1本目が返ってから投げてください。同時には投げられません。"
    Case 3
        CopiedTextOf = "3本目（調達・仕入れ）を写しました。これで標準の3本は終わりです。" & _
                       "もっと調べたいときは[＋ もっと調べる（あと5本）]を押してください。"
    Case Else
        CopiedTextOf = TitleOfPrompt(n) & "を写しました。" & _
                       "前の1本が返ってから投げてください。同時には投げられません。"
    End Select
End Function

' MarkNextFrame - 次に押すべき枠を太枠で示す(11章§3.2)。
'   1本目を押したら2本目の枠が太くなる。**これが「1本ずつ・直列」を体で覚え
'   させる仕掛け**である。配置は画面上の表示だけで config には書かない
'   (ブックを閉じれば消える)。
Private Sub MarkNextFrame(ByVal nextNo As Long)
    On Error Resume Next

    Dim n As Long
    For n = 1 To UR_ALL_COUNT
        Dim cell As Object
        Set cell = modUISheet.NamedCell(BodyRangeOf(n))
        If Not cell Is Nothing Then
            If n = nextNo Then
                cell.Borders.LineStyle = 1              ' xlContinuous
                cell.Borders.Weight = 4                 ' xlThick
                cell.Borders.Color = UR_COLOR_NEXT
            Else
                cell.Borders.LineStyle = -4142          ' xlLineStyleNone
            End If
        End If
    Next n
End Sub

' ============================================================================
' ToggleMore - 補助5本の表示/非表示(11章§3.2)。既定は閉じている。
' ============================================================================
Public Sub ToggleMore()
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_MORE) Then Exit Sub
    On Error GoTo Done

    modUINav.SetMoreOpen Not modUINav.MoreOpen()
    modUINavDraw.ShowMoreRows modUINav.MoreOpen()
    EnsureCopyButtons

    If modUINav.MoreOpen() Then
        modUIToast.ShowToast "調べる文を、あと5本ぶん出しました。" & _
                             "必要なものだけ写して投げてください。", "info"
    Else
        modUIToast.ShowToast "あと5本ぶんを閉じました。もう一度押すと出ます。", "info"
    End If
Done:
    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' OnAction ハンドラ(8本)。実体は共通の CopyPrompt(n)。
' 16章E-11: OnActionで配線される公開Subは先頭で TryEnterUiLock を通す。
' ============================================================================
Public Sub CopyPrompt1()
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_COPY) Then Exit Sub
    CopyPrompt 1
    modUIProgress.ExitUiLock
End Sub

Public Sub CopyPrompt2()
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_COPY) Then Exit Sub
    CopyPrompt 2
    modUIProgress.ExitUiLock
End Sub

Public Sub CopyPrompt3()
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_COPY) Then Exit Sub
    CopyPrompt 3
    modUIProgress.ExitUiLock
End Sub

Public Sub CopyPrompt4()
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_COPY) Then Exit Sub
    CopyPrompt 4
    modUIProgress.ExitUiLock
End Sub

Public Sub CopyPrompt5()
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_COPY) Then Exit Sub
    CopyPrompt 5
    modUIProgress.ExitUiLock
End Sub

Public Sub CopyPrompt6()
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_COPY) Then Exit Sub
    CopyPrompt 6
    modUIProgress.ExitUiLock
End Sub

Public Sub CopyPrompt7()
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_COPY) Then Exit Sub
    CopyPrompt 7
    modUIProgress.ExitUiLock
End Sub

Public Sub CopyPrompt8()
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_COPY) Then Exit Sub
    CopyPrompt 8
    modUIProgress.ExitUiLock
End Sub
