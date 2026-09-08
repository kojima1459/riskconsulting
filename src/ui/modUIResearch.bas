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

' ============================================================================
' 調べる場所(社内ディープリサーチ)のURL(裁定書26 C・13章§2.3)
' ----------------------------------------------------------------------------
' config の3キーが正で、空・欠落のときだけ下の既定へ倒す(**設定を消した
' だけで導線が死なない**ようにする)。既定値の逐語は13章§2.3 と同じ。
' 種別: "menu"=入口メニュー / "quick"=急ぎのとき / "full"=しっかり調査。
' ============================================================================
Private Const UR_DR_KIND_MENU As String = "menu"
Private Const UR_DR_KIND_QUICK As String = "quick"
Private Const UR_DR_KIND_FULL As String = "full"
Private Const UR_DR_URL_MENU As String = "https://app.hdtech.jp/research/menu"
Private Const UR_DR_URL_QUICK As String = "https://app.hdtech.jp/research/quick-search"
Private Const UR_DR_URL_FULL As String = "https://app.hdtech.jp/research/instructions"
Private Const UR_DR_OPEN_KEY As String = "dr_open_after_copy"
Private Const UR_LOCK_DR As String = "調査ページを開く"

' DrUrlDefaultOf - 種別ごとの既定URL(純関数・層(a)テスト対象)。
Public Function DrUrlDefaultOf(ByVal kind As String) As String
    Select Case kind
    Case UR_DR_KIND_MENU
        DrUrlDefaultOf = UR_DR_URL_MENU
    Case UR_DR_KIND_QUICK
        DrUrlDefaultOf = UR_DR_URL_QUICK
    Case UR_DR_KIND_FULL
        DrUrlDefaultOf = UR_DR_URL_FULL
    End Select
End Function

' DrUrlOf - config の値(cfgText)が空なら既定へ倒す(純関数・層(a)テスト対象)。
Public Function DrUrlOf(ByVal kind As String, ByVal cfgText As String) As String
    Dim s As String
    s = Trim$(cfgText)
    If LenB(s) = 0 Then
        DrUrlOf = DrUrlDefaultOf(kind)
    Else
        DrUrlOf = s
    End If
End Function

' 種別ごとの config キー名。
Private Function DrUrlKeyOf(ByVal kind As String) As String
    DrUrlKeyOf = "dr_url_" & kind
End Function

' いま使うURL(config を読んで DrUrlOf へ通す)。
Private Function DrUrlNow(ByVal kind As String) As String
    DrUrlNow = DrUrlOf(kind, modConfig.GetStr(DrUrlKeyOf(kind), vbNullString))
End Function

' 既定ブラウザで開く(Hyperlinks.Add は使わない。11章§8.6 禁忌1)。
'   開けたら True。EDR等で開けないことがあるので、呼び出し側は必ず
'   「開けなかったとき」の案内を出す。
'   裁定書34 §1.2: HTML画面(modNaviActions の open_url / open_report)も同じ
'   1本を通すため Public にした(URLの開き方をもう1つ作らない)。
Public Function OpenUrl(ByVal url As String) As Boolean
    On Error GoTo Failed
    ThisWorkbook.FollowHyperlink url
    OpenUrl = True
    Exit Function
Failed:
    Err.Clear
    OpenUrl = False
End Function

' ============================================================================
' OpenDrFull / OpenDrQuick - 区画①の見出し行の直下の2本(裁定書26 C)。
' ============================================================================
Public Sub OpenDrFull()
    OpenDrPage UR_DR_KIND_FULL
End Sub

Public Sub OpenDrQuick()
    OpenDrPage UR_DR_KIND_QUICK
End Sub

Private Sub OpenDrPage(ByVal kind As String)
    If Not modUIProgress.TryEnterUiLock(UR_LOCK_DR) Then Exit Sub
    On Error Resume Next
    Dim url As String
    url = DrUrlNow(kind)
    If OpenUrl(url) Then
        modUIToast.ShowToast "ブラウザで調査ページを開きました。", "info"
    Else
        modUIToast.ShowToast "ブラウザで " & url & " を開いてください。", "warn"
    End If
    modUIProgress.ExitUiLock
End Sub

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

' ============================================================================
' PlaceholderTable - `{{ }}` の差し込み表(11章§3.2 / 13章§2.19 が正)。
' ----------------------------------------------------------------------------
' 1行 = "プレースホルダ|種別|穴の文言" を vbLf 区切り。**置換辞書はここ1本**で
' あり、FillTemplate はこの表を上から順に当てるだけにする(表と実装が2箇所に
' 分かれると、docs/08 が語を増やしたときに片方だけ直る)。
'   種別 v_company  = 会社名をそのまま
'        v_place    = 本社の場所と業種名を「・」でつないだもの
'        v_address  = 本社の場所をそのまま
'        v_industry = 業種名をそのまま
'        h_seccode  = 証券コード。空なら穴
'        h_sites    = 調べたい拠点。空なら穴
'        x          = 画面に入力口が無い。常に穴
' 並びの規約: `{{本社所在地・業種}}` は `{{本社所在地}}` より**前**に置く
' (前方一致で食い違わないための保険。互いに部分文字列ではないが、表の並びで
'  意図を示しておく)。
' ============================================================================
Public Function PlaceholderTable() As String
    Dim s As String
    s = s & "{{企業名}}|v_company|" & vbLf
    s = s & "{{本社所在地・業種}}|v_place|" & vbLf
    s = s & "{{本社所在地}}|v_address|" & vbLf
    s = s & "{{業界名}}|v_industry|" & vbLf
    s = s & "{{業種名}}|v_industry|" & vbLf
    s = s & "{{コード}}|h_seccode|証券コードを書いてください（上場していなければ消してください）" & vbLf
    s = s & "{{拠点リスト（名称・住所）}}|h_sites|拠点の名前と住所" & vbLf
    s = s & "{{企業規模}}|x|会社の規模（従業員数や売上のめやす）" & vbLf
    s = s & "{{リスク/課題}}|x|気になっていること" & vbLf
    s = s & "{{前回更新からの期間}}|x|前回の更新からの期間" & vbLf
    s = s & "{{直近決算期}}|x|決算期を書いてください（例: 2026年3月期）" & vbLf
    s = s & "{{公式ドメイン}}|x|会社の公式サイトのURL"
    PlaceholderTable = s
End Function

' PlaceholderKeys - 置換辞書の見出し語だけを vbLf 区切りで返す。
'   層(a)が「docs/08 が使う `{{ }}` の集合 ⊆ 置換辞書」を機械で確かめる口。
Public Function PlaceholderKeys() As String
    Dim lines() As String
    lines = Split(PlaceholderTable(), vbLf)
    Dim i As Long
    Dim acc As String
    For i = LBound(lines) To UBound(lines)
        Dim f() As String
        f = Split(lines(i), "|")
        If LenB(acc) > 0 Then acc = acc & vbLf
        acc = acc & f(0)
    Next i
    PlaceholderKeys = acc
End Function

' `{{ }}` の差し込みと、埋まらない穴の日本語化(11章§3.2 の表が正)。
'   **戻り値に `{{` が1つも残らないことが契約**である(利用者は `{{ }}` の
'   意味を知らない)。層(a)が8本ぶんの雛形で全穴埋め・全空の2系を固定する。
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

    Dim lines() As String
    lines = Split(PlaceholderTable(), vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim f() As String
        f = Split(lines(i), "|")
        If UBound(f) - LBound(f) >= 2 Then
            Dim rep As String
            Select Case f(1)
            Case "v_company"
                rep = company
            Case "v_place"
                rep = placeText
            Case "v_address"
                rep = address
            Case "v_industry"
                rep = industry
            Case "h_seccode"
                rep = HoleOr(secCode, f(2))
            Case "h_sites"
                rep = HoleOr(sites, f(2))
            Case Else
                rep = Hole(f(2))
            End Select
            s = Replace$(s, f(0), rep)
        End If
    Next i
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

    ' 写した直後に調査ページを開く(裁定書26 C)。開けなかった(EDR等)ときは
    ' URLを逐語で案内する。**写せたこと自体は取り消さない**。
    If modConfig.GetBool(UR_DR_OPEN_KEY, True) Then
        Dim url As String
        url = DrUrlNow(UR_DR_KIND_FULL)
        If Not OpenUrl(url) Then
            modUIToast.ShowToast "コピーしました。ブラウザで " & url & _
                                 " を開いて貼り付けてください。", "warn"
        End If
    End If
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
