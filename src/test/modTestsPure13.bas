Attribute VB_Name = "modTestsPure13"
Option Explicit

' ============================================================================
' modTestsPure13 - W6第1弾(1画面ナビ)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 実装コードではなく 11章v3.2(§3.3.2 / §3.3.5 / §3.3.6 / §7.2(a))と
'   13章§2.10・§2.11 だけから期待値を導いた。期待値をあとから実装に合わせて
'   緩めることは禁止(17章§1)。
'
' 対象と根拠(11章§8.3 が指定した最低本数):
'   W6A  StripDrFooter    最低7本   11章§3.3.5(4語 / 複数一致 / 50%規則 / 一致なし)
'   W6B  PreviewLines     最低4本   11章§3.3.2(5行未満 / 5行超 / 120字超 / 空文字)
'   W6C  SplitFieldNotes  最低7本   11章§3.3.6(4見出し / 見出しなし / 順序違い /
'                                    空節 / 本文中の見出し語 / 末尾改行 / 例文除去)
'   計 18本(612 -> 630。7 + 4 + 7)
'
' グループ単位の失敗隔離: modTestsPure.bas と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

Public Sub RunAll()
    On Error GoTo FA
    T_W6A_StripDrFooter
WB:
    On Error GoTo FB
    T_W6B_PreviewLines
WC:
    On Error GoTo FC
    T_W6C_SplitFieldNotes
WDone:
    Exit Sub
FA:
    GroupFail "W6A StripDrFooter"
    Resume WB
FB:
    GroupFail "W6B PreviewLines"
    Resume WC
FC:
    GroupFail "W6C SplitFieldNotes"
    Resume WDone
End Sub

Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), "期待=[" & want & "] 実際=[" & act & "]"
End Sub

' 本文の半分より後ろに一致を置くための「十分に長い本文」。
Private Function LongBody() As String
    LongBody = "1. 主要拠点の所在地と自然災害リスク" & vbLf & _
               "本社・浜松工場は震度6強が想定される区域に位置している。" & vbLf & _
               "2. 調達先の集中" & vbLf & _
               "主要な原材料は2社に集中しており、代替の確保に時間がかかる。" & vbLf & _
               "3. 情報なし" & vbLf & _
               "（この項目に関する情報は見つかりませんでした）"
End Function

' ============================================================================
' W6A StripDrFooter(11章§3.3.5)
' ============================================================================
Private Sub T_W6A_StripDrFooter()
    Dim body As String
    body = LongBody()

    ' 1-4: 4種の語それぞれで、その行以降が末尾まで落ちる。
    ChkS "Test_W6A_01_役職コードの行以降を落とす_11章3.3.5", _
        modNavText.StripDrFooter(body & vbLf & "役職コード: A123" & vbLf & "部長"), body
    ChkS "Test_W6A_02_部課コードの行以降を落とす_11章3.3.5", _
        modNavText.StripDrFooter(body & vbLf & "部課コード：0456" & vbLf & "営業第一部"), body
    ChkS "Test_W6A_03_ご利用にあたっての行以降を落とす_11章3.3.5", _
        modNavText.StripDrFooter(body & vbLf & "ご利用にあたって" & vbLf & "本結果は参考です"), body
    ChkS "Test_W6A_04_履歴一覧の行以降を落とす_11章3.3.5", _
        modNavText.StripDrFooter(body & vbLf & "【履歴一覧】" & vbLf & "2026-09-01 質問A"), body

    ' 5: 複数見つかったら**いちばん先頭に近いもの**を採る(落とす量を最大にする)。
    ChkS "Test_W6A_05_複数一致は先頭に近いものを採る_11章3.3.5", _
        modNavText.StripDrFooter(body & vbLf & "ご利用にあたって" & vbLf & _
                                 "注意書き" & vbLf & "役職コード: A123"), body

    ' 6: 先頭50%より前でしか一致しないときは切らない(本文を全部消さない保険)。
    Dim early As String
    early = "ご利用にあたって" & vbLf & body
    ChkS "Test_W6A_06_先頭50%より前の一致は切らない_11章3.3.5", _
        modNavText.StripDrFooter(early), early

    ' 7: 一致が無ければ1字も変えない(「情報なし」等はそのまま残す)。
    ChkS "Test_W6A_07_一致なしは無変更_11章3.3.5", _
        modNavText.StripDrFooter(body), body

    ' (規則4「切り落としたあと末尾の空行を落とす」は 01-05 の期待値が
    '  そのまま含意している: 本文と footer のあいだの改行が残れば不一致になる)
End Sub

' ============================================================================
' W6B PreviewLines(11章§3.3.2)
' ============================================================================
Private Sub T_W6B_PreviewLines()
    ' 1: 5行に満たないときはある行だけを返す(足りない行を作らない)。
    ChkS "Test_W6B_01_5行未満はある行だけ_11章3.3.2", _
        modNavText.PreviewLines("あ" & vbLf & "い" & vbLf & "う", 5), _
        "あ" & vbLf & "い" & vbLf & "う"

    ' 2: 5行を超えたら先頭5行だけ(6行目以降は書かない)。
    ChkS "Test_W6B_02_5行超は先頭5行だけ_11章3.3.2", _
        modNavText.PreviewLines("1" & vbLf & "2" & vbLf & "3" & vbLf & "4" & vbLf & _
                                "5" & vbLf & "6", 5), _
        "1" & vbLf & "2" & vbLf & "3" & vbLf & "4" & vbLf & "5"

    ' 3: 1行が120字を超えたら先頭120字で切り、末尾へ省略記号を付ける。
    Dim long1 As String
    long1 = String$(130, "x")
    ChkS "Test_W6B_03_120字超は切って省略記号_11章3.3.2", _
        modNavText.PreviewLines(long1, 5), String$(120, "x") & "…"

    ' 4: 空文字は空文字(「まだ貼っていません」の判定は呼び出し側が持つ)。
    ChkS "Test_W6B_04_空文字は空文字_11章3.3.2", _
        modNavText.PreviewLines(vbNullString, 5), vbNullString
End Sub

' ============================================================================
' W6C SplitFieldNotes / JoinFieldNotes(11章§3.3.6)
' ============================================================================
Private Sub T_W6C_SplitFieldNotes()
    Dim memoText As String
    Dim othersText As String
    Dim src As String

    ' 1: 4見出しがそろっているとき memo は3節を見出しごと連結したもの。
    '    **入力は JoinFieldNotes で組み立てる**ので、この1本が
    '    JoinFieldNotes -> SplitFieldNotes の可逆性も同時に見ている(11章§8.3)。
    Dim memoIn As String
    memoIn = "【営業メモ】" & vbLf & "社長はワンマン。" & vbLf & _
             "【前回更新メモ】" & vbLf & "火災だけ更新。" & vbLf & _
             "【付保の見立て】" & vbLf & "火災は他社。"
    src = modNavText.JoinFieldNotes(memoIn, "裏の川が気になる。")
    modNavText.SplitFieldNotes src, memoText, othersText
    ChkS "Test_W6C_01_4見出しそろい_JoinしてSplitすると戻る_11章3.3.6", _
        memoText & "//" & othersText, memoIn & "//" & "裏の川が気になる。"

    ' 2: 見出しが1つも無いときは全文を memo へ入れ、先頭へ【営業メモ】を足す。
    modNavText.SplitFieldNotes "なんでも書いた文章。", memoText, othersText
    ChkS "Test_W6C_02_見出しなしは全文をmemoへ_11章3.3.6", memoText, _
        "【営業メモ】" & vbLf & "なんでも書いた文章。"

    ' 3: 見出しの順序が違っても、組み直しは決められた順(営業→前回→付保)。
    src = "【付保の見立て】" & vbLf & "C" & vbLf & _
          "【営業メモ】" & vbLf & "A" & vbLf & _
          "【前回更新メモ】" & vbLf & "B"
    modNavText.SplitFieldNotes src, memoText, othersText
    ChkS "Test_W6C_03_見出しの順序違いは正規順へ_11章3.3.6", memoText, _
        "【営業メモ】" & vbLf & "A" & vbLf & _
        "【前回更新メモ】" & vbLf & "B" & vbLf & _
        "【付保の見立て】" & vbLf & "C"

    ' 4: 節が空でも見出しは残す。
    src = "【営業メモ】" & vbLf & "【前回更新メモ】" & vbLf & "【付保の見立て】" & vbLf & "【そのほか】"
    modNavText.SplitFieldNotes src, memoText, othersText
    ChkS "Test_W6C_04_空節でも見出しは残す_11章3.3.6", memoText, _
        "【営業メモ】" & vbLf & "【前回更新メモ】" & vbLf & "【付保の見立て】"

    ' 5: 見出し語が本文中に現れても見出しにしない(行の完全一致で判定する)。
    src = "【営業メモ】" & vbLf & "これは【営業メモ】の続きです。" & vbLf & "【そのほか】" & vbLf & "X"
    modNavText.SplitFieldNotes src, memoText, othersText
    ChkS "Test_W6C_05_本文中の見出し語は見出しにしない_11章3.3.6", memoText, _
        "【営業メモ】" & vbLf & "これは【営業メモ】の続きです。" & vbLf & _
        "【前回更新メモ】" & vbLf & "【付保の見立て】"

    ' 6: 末尾の改行は落とす(見えない空行を保存しない)。
    src = "【営業メモ】" & vbLf & "A" & vbLf & vbLf & "【そのほか】" & vbLf & "X" & vbLf & vbLf
    modNavText.SplitFieldNotes src, memoText, othersText
    ChkS "Test_W6C_06_末尾改行は落とす_11章3.3.6", othersText, "X"

    ' 7: 「例: 」で始まる行はどの節でも保存しない(先置きの例文がAIへ渡らない)。
    src = "【営業メモ】" & vbLf & "  例: 社長はワンマンで決裁は即断。" & vbLf & "実際のメモ。" & vbLf & _
          "【そのほか】" & vbLf & "例: 工場の裏の川が気になる。"
    modNavText.SplitFieldNotes src, memoText, othersText
    ChkS "Test_W6C_07_例文の行はどの節でも保存しない_11章3.3.3", _
        memoText & "//" & othersText, _
        "【営業メモ】" & vbLf & "実際のメモ。" & vbLf & _
        "【前回更新メモ】" & vbLf & "【付保の見立て】" & "//"
End Sub
