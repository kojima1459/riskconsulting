Attribute VB_Name = "modHtmlTheme"
Option Explicit

' ============================================================================
' modHtmlTheme - HTMLレポートのテーマCSS(18章§5.1・§5.2)
' ----------------------------------------------------------------------------
' 責務: テーマ名に対応する CSS変数ブロック ":root{ ... }" **だけ**を返す。
'   セレクタ・レイアウト宣言・CSS変数以外の宣言を1つも書かない(18章§1の表)。
'
' なぜ1モジュールに閉じるのか(18章§5.2):
'   テーマを増やす作業を「ThemeCss の Select Case に分岐を1本足し、§5.1の
'   28変数を書く」だけで完結させるため。modExportHtml にも modHtmlTemplate1..n
'   にも触れずに配色を差し替えられる状態が、本章が守ろうとしている構造そのもの。
'
' 変数の閉じた一覧(18章§5.1。28個。これ以外を定義しない・これ以外を参照しない):
'   寸法・書体(6) --page-width / --page-pad / --font-sans / --font-serif /
'                  --font-size / --line-height
'   地色と文字(5) --paper / --ink / --sub / --mist / --line
'   強調(4)       --ai / --kaki / --matsu / --deep
'   注意面(2)     --warn / --warn-line
'   リスクマップの帯(5) --heat-1 .. --heat-5(§3.3の帯番号1から5に対応)
'   移転可能性(3) --tr-cover / --tr-partial / --tr-hard
'   入力充足度(3) --iq-ok / --iq-partial / --iq-missing
'
' 未知のテーマ名は standard へフォールバックする(§5.2)。**黙って戻さない**ため
'   の run_log 記録は呼び出し側(modExportHtml)が ThemeNames() と突き合わせて
'   行う。純文字列モジュールである本モジュールはログにも config にも触れない。
'
' R4準拠(12章§2・18章§1): Worksheets / Range( / Application. / ThisWorkbook /
'   MsgBox / ActiveSheet には一切触れない。案件データも参照しない。
' CP932準拠: 本文・注釈ともに CP932 内の文字だけで書く(絵文字不可)。
' ============================================================================

' --------------------------------------------------------------------------
' ThemeNames - 利用できるテーマ名の一覧(";"区切り・先頭が既定テーマ)。
'   テーマ名の一覧をconfigに書かない(18章§5.2。本関数が唯一の一覧)。
' --------------------------------------------------------------------------
Public Function ThemeNames() As String
    ThemeNames = "standard;mono"
End Function

' --------------------------------------------------------------------------
' ThemeCss - テーマ名に対応する ":root{ ... }" 1ブロックを返す(18章§5.1)。
'   戻り値は必ず ":root{" で始まり "}" で終わり、内側は "--" で始まる宣言だけ。
' --------------------------------------------------------------------------
Public Function ThemeCss(ByVal themeName As String) As String
    Select Case LCase$(Trim$(themeName))
        Case "mono"
            ThemeCss = MonoCss()
        Case Else
            ThemeCss = StandardCss()
    End Select
End Function

' --------------------------------------------------------------------------
' standard - §5.1の既定値。カラー画面・カラー印刷向け(白地・落ち着いた配色)。
' --------------------------------------------------------------------------
Private Function StandardCss() As String
    Dim s As String
    s = s & ":root{" & vbLf
    s = s & "  --page-width: 900px;" & vbLf
    s = s & "  --page-pad: 24px;" & vbLf
    s = s & "  --font-sans: ""Noto Sans JP"",""Yu Gothic"",""Hiragino Kaku Gothic ProN"",""Meiryo"",sans-serif;" & vbLf
    s = s & "  --font-serif: ""Shippori Mincho"",""Yu Mincho"",""Hiragino Mincho ProN"",serif;" & vbLf
    s = s & "  --font-size: 14.5px;" & vbLf
    s = s & "  --line-height: 1.85;" & vbLf
    s = s & "  --paper: #FFFFFF;" & vbLf
    s = s & "  --ink: #24303E;" & vbLf
    s = s & "  --sub: #5A6B7E;" & vbLf
    s = s & "  --mist: #EFF3F8;" & vbLf
    s = s & "  --line: #D8E0E9;" & vbLf
    s = s & "  --ai: #2B5B8F;" & vbLf
    s = s & "  --kaki: #B4552D;" & vbLf
    s = s & "  --matsu: #2F7A54;" & vbLf
    s = s & "  --deep: #7A5C9E;" & vbLf
    s = s & "  --warn: #FBF3E6;" & vbLf
    s = s & "  --warn-line: #E8D5B5;" & vbLf
    s = s & "  --heat-1: #E8F5EE;" & vbLf
    s = s & "  --heat-2: #F2F7E9;" & vbLf
    s = s & "  --heat-3: #FFF4D6;" & vbLf
    s = s & "  --heat-4: #FFE9D3;" & vbLf
    s = s & "  --heat-5: #FBDCD9;" & vbLf
    s = s & "  --tr-cover: #2F7A54;" & vbLf
    s = s & "  --tr-partial: #B08A2E;" & vbLf
    s = s & "  --tr-hard: #B4552D;" & vbLf
    s = s & "  --iq-ok: #2F7A54;" & vbLf
    s = s & "  --iq-partial: #B08A2E;" & vbLf
    s = s & "  --iq-missing: #A9B4C0;" & vbLf
    s = s & "}"
    StandardCss = s
End Function

' --------------------------------------------------------------------------
' mono - 白黒印刷・FAX配布向け(18章§5.2)。強調4色を --ink と --sub の濃淡に、
'   heat 5段を白から薄灰の5段に置き換える。§3.3の帯番号があるため mono でも
'   重篤度は読める(色に意味を運ばせない設計)。
' --------------------------------------------------------------------------
Private Function MonoCss() As String
    Dim s As String
    s = s & ":root{" & vbLf
    s = s & "  --page-width: 900px;" & vbLf
    s = s & "  --page-pad: 24px;" & vbLf
    s = s & "  --font-sans: ""Noto Sans JP"",""Yu Gothic"",""Hiragino Kaku Gothic ProN"",""Meiryo"",sans-serif;" & vbLf
    s = s & "  --font-serif: ""Shippori Mincho"",""Yu Mincho"",""Hiragino Mincho ProN"",serif;" & vbLf
    s = s & "  --font-size: 14.5px;" & vbLf
    s = s & "  --line-height: 1.85;" & vbLf
    s = s & "  --paper: #FFFFFF;" & vbLf
    s = s & "  --ink: #1F1F1F;" & vbLf
    s = s & "  --sub: #5A5A5A;" & vbLf
    s = s & "  --mist: #F2F2F2;" & vbLf
    s = s & "  --line: #C8C8C8;" & vbLf
    s = s & "  --ai: #1F1F1F;" & vbLf
    s = s & "  --kaki: #3C3C3C;" & vbLf
    s = s & "  --matsu: #5A5A5A;" & vbLf
    s = s & "  --deep: #4A4A4A;" & vbLf
    s = s & "  --warn: #F2F2F2;" & vbLf
    s = s & "  --warn-line: #B0B0B0;" & vbLf
    s = s & "  --heat-1: #FFFFFF;" & vbLf
    s = s & "  --heat-2: #F2F2F2;" & vbLf
    s = s & "  --heat-3: #E4E4E4;" & vbLf
    s = s & "  --heat-4: #D6D6D6;" & vbLf
    s = s & "  --heat-5: #C4C4C4;" & vbLf
    s = s & "  --tr-cover: #6E6E6E;" & vbLf
    s = s & "  --tr-partial: #4A4A4A;" & vbLf
    s = s & "  --tr-hard: #1F1F1F;" & vbLf
    s = s & "  --iq-ok: #1F1F1F;" & vbLf
    s = s & "  --iq-partial: #6E6E6E;" & vbLf
    s = s & "  --iq-missing: #B0B0B0;" & vbLf
    s = s & "}"
    MonoCss = s
End Function
