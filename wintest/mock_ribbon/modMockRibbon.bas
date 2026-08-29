Attribute VB_Name = "modMockRibbon"
Option Explicit

' ============================================================================
' modMockRibbon - ニセリボンちゃん(Windows実機テスト用スタブアドイン)
' ----------------------------------------------------------------------------
' 目的(12章§2 移植対応表・14章§4(b)・17章 T-14b):
'   本物の社内AIリボンが無い個人Windows PCでも、「リスク提案ナビ」の
'   **12引数 Application.Run 配管そのもの**(引数の型と順序・アドイン名の
'   部分一致検出・LimitCheck の真偽解釈・toolN の中身)を本物のWindows Excelで
'   検証するための、PoC確定台帳(RIBBON_API_CONFIRMED.md §1)と
'   **同一シグネチャ**のスタブ実装。
'
' 移植元: PoC「マイ本棚AI」 wintest/mock_ribbon/modMockRibbon.bas。
'   ChatGPT / LimitCheck の**シグネチャは1文字も変えていない**。
'   RPN に呼び出し経路が存在しない関数(GetEmbeddings / CosineSimilarityN2 /
'   ChatGPTV / IsImageInCB / Base64FromCB / Base64FromFile / ConvToJpeg /
'   OpenWordMark / OpenMemo / CellMarkDown)は移植していない。
'   これらは PoC のRAG・画像取込・Word連携のための口で、RPN の
'   modGatewayRPN は ChatGPT と LimitCheck しか呼ばない(12章§2「RAG検索系関数は
'   移植しない」)。スタブに余計な口を残すと、検証したい配管が増えるだけでなく、
'   万一会社PCへ持ち込まれたときの副作用が増える。
'
' 使い方(wintest/README.md 参照): このモジュールを空のブックにインポートし、
'   「リボンちゃん(検証用).xlam」という名前でExcelアドインとして保存して有効化する。
'   **ファイル名に「リボンちゃん」を含めること**(RPNのアドイン検出は
'   config `ribbon_addin_name`(既定「リボンちゃん」)の部分一致のため。14章§2)。
'
' 注意:
'   ・AIの回答品質は検証できない(全て決め打ちのダミー応答)。
'   ・**本体内のmockトランスポート modMockLlm とは別物**(14章§4)。
'     決定的なJSON応答は modMockLlm が返す。こちらは配管の検証専用。
'   ・本物のリボンちゃんと同居させないこと(会社PCには入れない)。
' ============================================================================

Private Const MOCK_TAG As String = "(ニセリボン応答)"

' --- 0) 起動制御 (台帳§2 D3) -----------------------------------------------

Public Function LimitCheck() As Boolean
    ' True=続行不可 の解釈。スタブは常に続行可を返す。
    ' RPN側は True でも起動を止めない(12章§2.1 手順(7)・14章§2)。
    LimitCheck = False
End Function

' --- 1) ChatGPT (台帳§1 #1 + 互換後方追加のeffort/verbosity) ---------------
' 引数の順序と型は確定シグネチャそのもの。ここを変えると T-14b の検証が
' 「配管の検証」ではなくなるため、1つも足さず・削らず・並べ替えないこと。

Public Function ChatGPT(ByVal Text As String, _
        Optional ByVal roleSystem As String, _
        Optional ByVal Temperature As Double = 0.4, _
        Optional ByVal MaxTokens As Long = 4096, _
        Optional ByVal Wait As Long = 120, _
        Optional ByVal optModel As String, _
        Optional ByVal prevU As String, _
        Optional ByVal prevA As String, _
        Optional ByVal toolN As String, _
        Optional ByVal reasoning_effort As String, _
        Optional ByVal verbosity As String) As String

    Dim s As String

    ' 受領した引数をそのまま応答に反映する(=呼び出し側から「型と順序どおりに
    ' 届いたか」を機械的に確認できるようにする。T-14b の(1)と(4))。
    s = MOCK_TAG & " model=" & optModel & " tool=" & toolN & vbLf
    s = s & "temp=" & Temperature & " maxTokens=" & MaxTokens & " wait=" & Wait & vbLf
    s = s & "systemLen=" & Len(roleSystem) & " userLen=" & Len(Text) & vbLf
    s = s & "effort=" & reasoning_effort & " verbosity=" & verbosity & vbLf

    ' 会話履歴が渡ってきたことを応答に反映する(PL-04壁打ちの prevU/prevA 配管。
    ' 履歴は「新しい順」に ";;;" 区切りで連結して渡す規約。14章§6)
    If LenB(prevU) > 0 Then
        s = s & "前回の質問「" & FirstToken(prevU) & "」を踏まえた続きの回答です。" & vbLf
    End If
    If LenB(prevA) > 0 Then
        s = s & "前回の回答の先頭「" & FirstToken(prevA) & "」を受領しました。" & vbLf
    End If

    s = s & "ここに回答本文が入ります(ニセリボンは品質を検証しません)。"

    ChatGPT = s
End Function

' prevU / prevA(新しい順・";;;"区切り)の先頭要素を返す
Private Function FirstToken(ByVal joined As String) As String
    Dim p As Long
    p = InStr(joined, ";;;")
    If p > 0 Then
        FirstToken = Left$(joined, p - 1)
    Else
        FirstToken = joined
    End If
    If Len(FirstToken) > 40 Then FirstToken = Left$(FirstToken, 40) & "..."
End Function
