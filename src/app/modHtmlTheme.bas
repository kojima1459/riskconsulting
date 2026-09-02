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
'   39変数を書く」だけで完結させるため。modExportHtml にも modHtmlTemplate1..n
'   にも触れずに配色を差し替えられる状態が、本章が守ろうとしている構造そのもの。
'
' 変数の閉じた一覧(18章§5.1。v1.2で28->39。旧 --ai は --brand へ改称):
'   寸法・書体(6) --page-width / --page-pad / --font-sans / --font-serif /
'                  --font-size / --line-height
'   地色と文字(6) --bg / --paper / --ink / --sub / --mist / --line
'   主色と強調(7) --brand / --brand2 / --accent / --navy / --kaki / --matsu / --deep
'   面色と影(7)   --soft-brand / --soft-red / --soft-amber / --soft-green /
'                  --soft-blue / --soft-purple / --shadow
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
    ThemeNames = "standard;mono;ds"
End Function

' --------------------------------------------------------------------------
' ThemeCss - テーマ名に対応する ":root{ ... }" 1ブロックを返す(18章§5.1)。
'   戻り値は必ず ":root{" で始まり "}" で終わり、内側は "--" で始まる宣言だけ。
' --------------------------------------------------------------------------
Public Function ThemeCss(ByVal themeName As String) As String
    Select Case LCase$(Trim$(themeName))
        Case "mono"
            ThemeCss = MonoCss()
        Case "ds"
            ThemeCss = DsCss()
        Case Else
            ThemeCss = StandardCss()
    End Select
End Function

' --------------------------------------------------------------------------
' 全テーマ共通の寸法・書体(§5.1の6変数)。値そのものは各テーマが決めてよいが、
'   本製品の3テーマは同じ紙面寸法を使うのでここへ集約する(3箇所に写経しない)。
' --------------------------------------------------------------------------
Private Function MetricsCss() As String
    Dim s As String
    s = s & "  --page-width: 1180px;" & vbLf
    s = s & "  --page-pad: 22px;" & vbLf
    s = s & "  --font-sans: -apple-system,BlinkMacSystemFont,""Segoe UI"",""Hiragino Kaku Gothic ProN"",""Yu Gothic"",Meiryo,sans-serif;" & vbLf
    s = s & "  --font-serif: ""Shippori Mincho"",""Yu Mincho"",""Hiragino Mincho ProN"",serif;" & vbLf
    s = s & "  --font-size: 14px;" & vbLf
    s = s & "  --line-height: 1.7;" & vbLf
    MetricsCss = s
End Function

' --------------------------------------------------------------------------
' standard - §5.1の既定値。体裁の正である見本(docs/design/出力見本_春華堂
'   統合提案_v0.1.html)の :root を出発点にした臙脂の配色。
' --------------------------------------------------------------------------
Private Function StandardCss() As String
    Dim s As String
    s = s & ":root{" & vbLf
    s = s & MetricsCss() ' SAFE:html
    s = s & "  --bg: #F5F6F8;" & vbLf
    s = s & "  --paper: #FFFFFF;" & vbLf
    s = s & "  --ink: #1D2433;" & vbLf
    s = s & "  --sub: #667085;" & vbLf
    s = s & "  --mist: #F8F9FB;" & vbLf
    s = s & "  --line: #E6E8EC;" & vbLf
    s = s & "  --brand: #A5312F;" & vbLf
    s = s & "  --brand2: #6F1E1E;" & vbLf
    s = s & "  --accent: #B88A44;" & vbLf
    s = s & "  --navy: #27364A;" & vbLf
    s = s & "  --kaki: #B4552D;" & vbLf
    s = s & "  --matsu: #2F7A54;" & vbLf
    s = s & "  --deep: #6F42C1;" & vbLf
    s = s & "  --soft-brand: #F7EEED;" & vbLf
    s = s & "  --soft-red: #FFF1F0;" & vbLf
    s = s & "  --soft-amber: #FFF7E8;" & vbLf
    s = s & "  --soft-green: #EDF8F2;" & vbLf
    s = s & "  --soft-blue: #EEF4FF;" & vbLf
    s = s & "  --soft-purple: #F5F1FF;" & vbLf
    s = s & "  --shadow: 0 12px 32px rgba(16,24,40,.07);" & vbLf
    s = s & "  --warn: #FBF3E6;" & vbLf
    s = s & "  --warn-line: #E8D5B5;" & vbLf
    s = s & "  --heat-1: #E9F5EE;" & vbLf
    s = s & "  --heat-2: #F2F7E9;" & vbLf
    s = s & "  --heat-3: #FFF6DB;" & vbLf
    s = s & "  --heat-4: #FFE8D3;" & vbLf
    s = s & "  --heat-5: #FFD9D7;" & vbLf
    s = s & "  --tr-cover: #2F7A54;" & vbLf
    s = s & "  --tr-partial: #B08A2E;" & vbLf
    s = s & "  --tr-hard: #B42318;" & vbLf
    s = s & "  --iq-ok: #2F7A54;" & vbLf
    s = s & "  --iq-partial: #B08A2E;" & vbLf
    s = s & "  --iq-missing: #A9B4C0;" & vbLf
    s = s & "}"
    StandardCss = s
End Function

' --------------------------------------------------------------------------
' mono - 白黒印刷・FAX配布向け(18章§5.2)。主色・強調色を --ink と --sub の
'   濃淡に、heat 5段を白から薄灰の5段に置き換える。§3.3の帯番号と§3.7の
'   「効きめ {n}/5」の併記があるため mono でも重篤度・効きめは読める
'   (色に意味を運ばせない設計)。
' --------------------------------------------------------------------------
Private Function MonoCss() As String
    Dim s As String
    s = s & ":root{" & vbLf
    s = s & MetricsCss() ' SAFE:html
    s = s & "  --bg: #F6F6F6;" & vbLf
    s = s & "  --paper: #FFFFFF;" & vbLf
    s = s & "  --ink: #1F1F1F;" & vbLf
    s = s & "  --sub: #5A5A5A;" & vbLf
    s = s & "  --mist: #F2F2F2;" & vbLf
    s = s & "  --line: #C8C8C8;" & vbLf
    s = s & "  --brand: #1F1F1F;" & vbLf
    s = s & "  --brand2: #000000;" & vbLf
    s = s & "  --accent: #4A4A4A;" & vbLf
    s = s & "  --navy: #2E2E2E;" & vbLf
    s = s & "  --kaki: #3C3C3C;" & vbLf
    s = s & "  --matsu: #5A5A5A;" & vbLf
    s = s & "  --deep: #4A4A4A;" & vbLf
    s = s & "  --soft-brand: #EDEDED;" & vbLf
    s = s & "  --soft-red: #E4E4E4;" & vbLf
    s = s & "  --soft-amber: #EDEDED;" & vbLf
    s = s & "  --soft-green: #F4F4F4;" & vbLf
    s = s & "  --soft-blue: #F4F4F4;" & vbLf
    s = s & "  --soft-purple: #EDEDED;" & vbLf
    s = s & "  --shadow: 0 0 0 1px rgba(0,0,0,.10);" & vbLf
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

' --------------------------------------------------------------------------
' ds - DS版(docs/design/出力見本_春華堂統合提案_DS版_v0.1.html)。白地に
'   紺青の帯。構成は1文字も変えず配色だけが変わることの実証(18章§5.2・
'   11章§3.8.2 #1)。
' --------------------------------------------------------------------------
Private Function DsCss() As String
    Dim s As String
    s = s & ":root{" & vbLf
    s = s & MetricsCss() ' SAFE:html
    s = s & "  --bg: #F4F6FA;" & vbLf
    s = s & "  --paper: #FFFFFF;" & vbLf
    s = s & "  --ink: #16202E;" & vbLf
    s = s & "  --sub: #5B6B80;" & vbLf
    s = s & "  --mist: #F5F8FC;" & vbLf
    s = s & "  --line: #DCE3EC;" & vbLf
    s = s & "  --brand: #1B4E8C;" & vbLf
    s = s & "  --brand2: #0F2E55;" & vbLf
    s = s & "  --accent: #A07C36;" & vbLf
    s = s & "  --navy: #16324F;" & vbLf
    s = s & "  --kaki: #B4552D;" & vbLf
    s = s & "  --matsu: #2F7A54;" & vbLf
    s = s & "  --deep: #5B49A8;" & vbLf
    s = s & "  --soft-brand: #EAF1FA;" & vbLf
    s = s & "  --soft-red: #FFF1F0;" & vbLf
    s = s & "  --soft-amber: #FFF7E8;" & vbLf
    s = s & "  --soft-green: #EDF8F2;" & vbLf
    s = s & "  --soft-blue: #EEF4FF;" & vbLf
    s = s & "  --soft-purple: #F1EFFB;" & vbLf
    s = s & "  --shadow: 0 10px 26px rgba(16,32,60,.08);" & vbLf
    s = s & "  --warn: #FBF3E6;" & vbLf
    s = s & "  --warn-line: #E8D5B5;" & vbLf
    s = s & "  --heat-1: #E9F5EE;" & vbLf
    s = s & "  --heat-2: #F1F6E9;" & vbLf
    s = s & "  --heat-3: #FFF6DB;" & vbLf
    s = s & "  --heat-4: #FFE8D3;" & vbLf
    s = s & "  --heat-5: #FFD9D7;" & vbLf
    s = s & "  --tr-cover: #2F7A54;" & vbLf
    s = s & "  --tr-partial: #B08A2E;" & vbLf
    s = s & "  --tr-hard: #B42318;" & vbLf
    s = s & "  --iq-ok: #2F7A54;" & vbLf
    s = s & "  --iq-partial: #B08A2E;" & vbLf
    s = s & "  --iq-missing: #A9B4C0;" & vbLf
    s = s & "}"
    DsCss = s
End Function
