Attribute VB_Name = "modTestsPure3"
Option Explicit

' ============================================================================
' modTestsPure3 - modUtil(14章§6 modUtil節)の契約テスト(17章§4-1 層(a))G12-G17
' ----------------------------------------------------------------------------
' 役割:
'   **本ファイルが叩くのは modUtil の公開関数だけ**である。14章§6の modUtil節
'   (裁定書6 項目8で契約化)が定める入出力を、実装コードを読まずに固定する。
'   modUtil は case_data の分割保存・注入IDの累積・行バッファといった上位機能が
'   共通に使う道具であり、ここが壊れると広範囲が壊れる。
'
'   **重要(裁定書6 項目10)**: 本ファイルは modCaseStore / modKnowledge /
'   modKnowledgeFmt の**振る舞いを検査しない**。それらの本体(採番・状態遷移・
'   参照優先・15章の整形書式・§0.7の切詰め)は 14章§6 が公開した純関数
'   (modCaseStore.BuildCaseId / IsValidCaseId / CanTransition / ResolveDataKey、
'   modKnowledgeFmt.Fmt* / TrimKbLine / TrimPlan)に対する別ファイルのテストが
'   受け持つ。ここでそれらを名乗ると「押さえてあるように見えて誰も見ていない」
'   状態を作る(W2aで実際に起きた)。テスト名・根拠コメントは modUtil契約に限る。
'
'   入口は Public Sub RunAll()。modTestsPure2.RunAll の末尾から呼ばれる。
'
' 設計判断(R4準拠): Worksheets / Range / Application / ThisWorkbook / MsgBox /
'   ActiveSheet には一切触れない。改行は vbLf 基準。
'
' 本ファイルのテスト本数: 69本(G12 40 / G13 8 / G14 4 / G15 6 / G16 6 / G17 5)
'
' グループ / 本数 / 根拠(すべて14章§6 modUtil節):
'   G12 SplitForCells / JoinCellChunks の往復  40本 (+13章§2.2の32,000字単位)
'   G13 SplitKeepNonEmpty の分割とTrim          8本 (+13章§2.2のセル格納規約)
'   G14 AppendIdList の重複排除と ";" 区切り    4本
'   G15 ClampLong の上下限                      6本
'   G16 SafeLeft の先頭切詰め                   6本
'   G17 BufInit / BufAdd / BufText の行バッファ 5本 (区切りは vbLf)
'
' 本ファイルが前提とする公開契約(すべて modUtil。14章§6 modUtil節):
'   modUtil.SplitForCells / JoinCellChunks / SplitKeepNonEmpty / AppendIdList
'   modUtil.ClampLong / SafeLeft / BufInit / BufAdd / BufText
'   modTestRunner.Check
' ============================================================================

' 13章§2.2: case_data の content は最大32,000字。分割保存の単位。
Private Const CH_MAX As Long = 32000
' 16章E-22: セルの物理上限は32,767字(32,000字契約はこの内側に取ってある)。
Private Const CELL_MAX As Long = 32767
' SafeLeft へ渡す切り出し長のサンプル値(2,000)。
Private Const CUT_2000 As Long = 2000
' SafeLeft へ渡す切り出し長のサンプル値(400)。本ファイルは SafeLeft が「先頭n字で
' 切る」ことだけを見る。ナレッジ行の400字切詰めの実装は modKnowledgeFmt.TrimKbLine。
Private Const CUT_400 As Long = 400

Public Sub RunAll()
    On Error GoTo F12
    T_ChunkStore
G13:
    On Error GoTo F13
    T_CellList
G14:
    On Error GoTo F14
    T_InjectedIds
G15:
    On Error GoTo F15
    T_KbTrimFloor
G16:
    On Error GoTo F16
    T_HeadTrim
G17:
    On Error GoTo F17
    T_LineBuffer
GDone:
    On Error GoTo 0
    modTestsPure4.RunAll
    Exit Sub

F12:
    GroupFail "G12 ChunkStore"
    Resume G13
F13:
    GroupFail "G13 CellList"
    Resume G14
F14:
    GroupFail "G14 InjectedIds"
    Resume G15
F15:
    GroupFail "G15 KbTrimFloor"
    Resume G16
F16:
    GroupFail "G16 HeadTrim"
    Resume G17
F17:
    GroupFail "G17 LineBuffer"
    Resume GDone
End Sub

' ----------------------------------------------------------------------------
' 共通ヘルパ(modTestsPure2 と同じ作法)
' ----------------------------------------------------------------------------
Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

' 文字列一致の1本。長い期待値はレポートが膨らむので先頭120字だけ残す。
Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), _
        "期待=[" & HeadOf(want) & "] 実際=[" & HeadOf(act) & "]"
End Sub

' 数値一致の1本。
Private Sub ChkN(ByVal nm As String, ByVal act As Long, ByVal want As Long)
    modTestRunner.Check nm, (act = want), "期待=" & want & " 実際=" & act
End Sub

' 真偽の1本。
Private Sub ChkB(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check nm, cond, detail
End Sub

' レポート用の先頭抜粋(テスト対象の SafeLeft には依存させない)。
Private Function HeadOf(ByVal s As String) As String
    If Len(s) <= 120 Then
        HeadOf = s
    Else
        HeadOf = Left(s, 120) & "...(全" & Len(s) & "字)"
    End If
End Function

' 配列の要素数。LBound がいくつでも、未初期化でも1件の数値として扱えるようにする。
Private Function ArrN(ByRef a() As String) As Long
    Dim n As Long
    n = -1
    On Error Resume Next
    n = UBound(a) - LBound(a) + 1
    On Error GoTo 0
    If n < 0 Then n = 0
    ArrN = n
End Function

' 配列の i 番目(0始まりの相対位置)。範囲外は空文字にして失敗として見せる。
Private Function ArrAt(ByRef a() As String, ByVal i As Long) As String
    Dim r As String
    r = ""
    On Error Resume Next
    r = a(LBound(a) + i)
    On Error GoTo 0
    ArrAt = r
End Function

' ch を n 個ならべた文字列。倍々に伸ばしてから切るので O(log n) 回の連結で済む
' (1文字ずつの連結は40,000字級で実用にならない)。
Private Function RepChar(ByVal ch As String, ByVal n As Long) As String
    Dim s As String
    If n <= 0 Or Len(ch) <= 0 Then
        RepChar = ""
        Exit Function
    End If
    s = ch
    Do While Len(s) < n
        s = s & s
    Loop
    RepChar = Left(s, n)
End Function

' ----------------------------------------------------------------------------
' G12 case_data の32,000字分割・結合
'   根拠: 13章§2.2「seq=分割連番。contentは最大32,000字。結合は modCaseStore が
'   透過処理」/ 16章E-22「case_dataは32,000字で分割保存」/ 17章 T-20 DoD
'   「40,000字往復一致」。modCaseStore はシートI/Oを伴うため、層(a)では
'   その土台である modUtil の分割・結合の契約を固定する。
' ----------------------------------------------------------------------------
Private Sub T_ChunkStore()
    Dim src As String
    Dim parts() As String
    Dim i As Long
    Dim total As Long
    Dim okLen As Boolean

    ' --- 最小入力 ---
    src = "A"
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_1字は1チャンク_13章§2.2", ArrN(parts), 1
    ChkS "SplitForCells_1字チャンクの中身が一致_13章§2.2", ArrAt(parts, 0), "A"
    ChkS "JoinCellChunks_1字の往復一致_13章§2.2", modUtil.JoinCellChunks(parts), src

    ' --- 32,000字ちょうどの前後(分割境界) ---
    src = RepChar("a", CH_MAX - 1)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_31999字は1チャンク_13章§2.2", ArrN(parts), 1

    src = RepChar("a", CH_MAX)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_32000字ちょうどは1チャンク_13章§2.2", ArrN(parts), 1
    ChkN "SplitForCells_32000字ちょうどのチャンク長_13章§2.2", Len(ArrAt(parts, 0)), CH_MAX
    ChkS "JoinCellChunks_32000字ちょうどの往復一致_16章E22", _
        modUtil.JoinCellChunks(parts), src

    src = RepChar("a", CH_MAX + 1)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_32001字は2チャンク_13章§2.2", ArrN(parts), 2
    ChkN "SplitForCells_32001字の1本目は32000字_13章§2.2", Len(ArrAt(parts, 0)), CH_MAX
    ChkN "SplitForCells_32001字の2本目は1字_13章§2.2", Len(ArrAt(parts, 1)), 1
    ChkS "JoinCellChunks_32001字の往復一致_16章E22", _
        modUtil.JoinCellChunks(parts), src

    ' --- セル物理上限32,767字(16章E-22。32,000字契約はこの内側) ---
    src = RepChar("a", CELL_MAX)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_32767字は2チャンク_16章E22", ArrN(parts), 2
    ChkS "JoinCellChunks_32767字の往復一致_16章E22", _
        modUtil.JoinCellChunks(parts), src

    ' --- 2倍・2倍+1 ---
    src = RepChar("a", CH_MAX * 2)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_64000字は2チャンク_13章§2.2", ArrN(parts), 2

    src = RepChar("a", CH_MAX * 2 + 1)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_64001字は3チャンク_13章§2.2", ArrN(parts), 3

    ' --- 17章 T-20 DoD の40,000字往復 ---
    src = RepChar("x", 40000)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_40000字は2チャンク_T20DoD", ArrN(parts), 2
    ChkS "JoinCellChunks_40000字の往復一致_T20DoD", _
        modUtil.JoinCellChunks(parts), src

    total = 0
    okLen = True
    For i = 0 To ArrN(parts) - 1
        total = total + Len(ArrAt(parts, i))
        If Len(ArrAt(parts, i)) > CH_MAX Then okLen = False
    Next i
    ChkN "SplitForCells_40000字の総長が原文と一致_13章§2.2", total, 40000
    ChkB "SplitForCells_各チャンクが32000字以下_16章E22", okLen, _
        "1本でも32,000字を超えたらセル分割の前提が崩れる"

    ' --- 3チャンク(端数あり / ちょうど) ---
    src = RepChar("y", 80000)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_80000字は3チャンク_13章§2.2", ArrN(parts), 3
    ChkN "SplitForCells_80000字の3本目は16000字_13章§2.2", Len(ArrAt(parts, 2)), 16000
    ChkS "JoinCellChunks_80000字の往復一致_13章§2.2", _
        modUtil.JoinCellChunks(parts), src

    src = RepChar("z", CH_MAX * 3)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_96000字は3チャンク_13章§2.2", ArrN(parts), 3
    ChkS "JoinCellChunks_96000字の往復一致_13章§2.2", _
        modUtil.JoinCellChunks(parts), src

    ' --- seq の並び = 原文の並び(13章§2.2「seq=分割連番」) ---
    src = RepChar("p", 70000)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkS "SplitForCells_seq1は原文の先頭区間_13章§2.2seq", _
        ArrAt(parts, 0), Mid(src, 1, CH_MAX)
    ChkS "SplitForCells_seq2は原文の第2区間_13章§2.2seq", _
        ArrAt(parts, 1), Mid(src, CH_MAX + 1, CH_MAX)
    ChkS "SplitForCells_seq3は原文の残余_13章§2.2seq", _
        ArrAt(parts, 2), Mid(src, CH_MAX * 2 + 1)

    ' --- 境界の直前・直後の1字が欠落も重複もしない ---
    src = RepChar("a", CH_MAX - 1) & "X" & "Y" & RepChar("b", 2000)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkS "SplitForCells_境界直前の文字が1本目の末尾_13章§2.2", _
        Right(ArrAt(parts, 0), 1), "X"
    ChkS "SplitForCells_境界直後の文字が2本目の先頭_13章§2.2", _
        Left(ArrAt(parts, 1), 1), "Y"
    ChkS "JoinCellChunks_境界前後の文字を欠落なく復元_13章§2.2", _
        modUtil.JoinCellChunks(parts), src

    ' --- 改行(13章§2.2「セル内改行(vbLf)をそのまま保持」) ---
    src = RepChar("a", CH_MAX - 1) & vbLf & RepChar("b", 8000)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_境界に改行がある場合も2チャンク_13章§2.2", ArrN(parts), 2
    ChkS "JoinCellChunks_境界をまたぐ改行を保持_13章§2.2", _
        modUtil.JoinCellChunks(parts), src

    ' --- 日本語(分割は文字数基準であってバイト数基準ではない) ---
    src = RepChar("あ", 40000)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkN "SplitForCells_日本語40000字は2チャンク_13章§2.2", ArrN(parts), 2
    ChkN "SplitForCells_日本語の1本目は32000字_13章§2.2", Len(ArrAt(parts, 0)), CH_MAX
    ChkS "JoinCellChunks_日本語40000字の往復一致_13章§2.2", _
        modUtil.JoinCellChunks(parts), src

    ' --- sN_json の実相(境界記号・二重引用符・タブ・全角空白を含む本文) ---
    src = RepChar("a", 20000) & "■■■" & "{""risks"": [1, 2]}" & vbTab & _
          "　" & RepChar("b", 20000)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkS "JoinCellChunks_境界記号とJSON本文の往復一致_13章§2.2", _
        modUtil.JoinCellChunks(parts), src

    ' --- 単一チャンクの結合はその値そのもの ---
    src = RepChar("s", 100)
    parts = modUtil.SplitForCells(src, CH_MAX)
    ChkS "JoinCellChunks_単一チャンクはその値を返す_13章§2.2", _
        modUtil.JoinCellChunks(parts), src

    ' --- chunkLen は引数で決まる(32,000は呼び出し側の契約値) ---
    src = RepChar("q", 40000)
    parts = modUtil.SplitForCells(src, 1000)
    ChkN "SplitForCells_chunkLen1000は40チャンク_13章§2.2", ArrN(parts), 40
    ChkN "SplitForCells_chunkLen1000の1本目は1000字_13章§2.2", Len(ArrAt(parts, 0)), 1000
    ChkS "JoinCellChunks_chunkLen1000の往復一致_13章§2.2", _
        modUtil.JoinCellChunks(parts), src
End Sub

' ----------------------------------------------------------------------------
' G13 SplitKeepNonEmpty(14章§6 modUtil節)
'   契約: 「区切って空要素を捨てる。各要素は Trim する。0件・sep が空のときは
'   0要素」。区切り "; " は 13章§2.2 のセル格納規約が使う実際の値なので、
'   代表値としてこれを渡している(検査対象は modUtil の分割そのもの)。
' ----------------------------------------------------------------------------
Private Sub T_CellList()
    Dim items() As String

    items = modUtil.SplitKeepNonEmpty("", "; ")
    ChkN "SplitKeepNonEmpty_空文字は要素0件_14章§6modUtil", ArrN(items), 0

    items = modUtil.SplitKeepNonEmpty("冷凍食品; 惣菜; 調味料", "; ")
    ChkN "SplitKeepNonEmpty_3要素の件数_14章§6modUtil", ArrN(items), 3
    ChkS "SplitKeepNonEmpty_1要素目_14章§6modUtil", ArrAt(items, 0), "冷凍食品"
    ChkS "SplitKeepNonEmpty_2要素目_14章§6modUtil", ArrAt(items, 1), "惣菜"
    ChkS "SplitKeepNonEmpty_3要素目_14章§6modUtil", ArrAt(items, 2), "調味料"

    items = modUtil.SplitKeepNonEmpty("M-0012", "; ")
    ChkN "SplitKeepNonEmpty_単一要素は1件_14章§6modUtil", ArrN(items), 1

    items = modUtil.SplitKeepNonEmpty("2; 5", "; ")
    ChkN "SplitKeepNonEmpty_数字並びも2要素_14章§6modUtil", ArrN(items), 2
    ChkS "SplitKeepNonEmpty_数字並びの2要素目_14章§6modUtil", ArrAt(items, 1), "5"
End Sub

' ----------------------------------------------------------------------------
' G14 注入IDの累積
'   根拠: 14章§6 modKnowledge.LastInjectedIds「直近の ResetInjectedIds 以降に
'   各注入関数が実際に使ったナレッジIDの累積(";"区切り)」/ 17章 T-21 DoD
'   「整形テキストに現れるIDの集合と LastInjectedIds() の返す ";" 区切りID列が
'   完全一致」。LastInjectedIds 本体はナレッジシート読込を伴うため層(a)対象外。
'   区切り文字の契約だけを modUtil.AppendIdList で固定する。
' ----------------------------------------------------------------------------
Private Sub T_InjectedIds()
    Dim acc As String

    acc = modUtil.AppendIdList("", "M-0012")
    ChkS "注入ID累積_空リストへの追加はそのIDだけ_14章§6", acc, "M-0012"

    acc = modUtil.AppendIdList(acc, "L-03")
    ChkB "注入ID累積_1件目が残る_14章§6", (InStr(acc, "M-0012") > 0), "実際=[" & acc & "]"
    ChkB "注入ID累積_2件目が含まれる_14章§6", (InStr(acc, "L-03") > 0), "実際=[" & acc & "]"
    ChkB "注入ID累積_区切りはセミコロン_14章§6", (InStr(acc, ";") > 0), "実際=[" & acc & "]"
End Sub

' ----------------------------------------------------------------------------
' G15 ClampLong(14章§6 modUtil節)
'   契約: 「値を [minV, maxV] へ収める。minV > maxV の指定は minV を優先」。
'   下の 5 / 20 / 60 は代表値であって、ナレッジ切詰めの下限・上限そのものでは
'   ない(15章§0.7 の計画は modKnowledgeFmt.TrimPlan が実装する)。
' ----------------------------------------------------------------------------
Private Sub T_KbTrimFloor()
    ChkN "ClampLong_下限未満は下限へ_14章§6modUtil", _
        modUtil.ClampLong(2, 5, 60), 5
    ChkN "ClampLong_下限未満は下限へ2_14章§6modUtil", _
        modUtil.ClampLong(3, 5, 60), 5
    ChkN "ClampLong_下限未満は下限へ3_14章§6modUtil", _
        modUtil.ClampLong(1, 5, 20), 5
    ChkN "ClampLong_下限0のときは0を許す_14章§6modUtil", _
        modUtil.ClampLong(0, 0, 5), 0
    ChkN "ClampLong_上限超過は上限へ_14章§6modUtil", _
        modUtil.ClampLong(80, 5, 60), 60
    ChkN "ClampLong_範囲内はそのまま_14章§6modUtil", _
        modUtil.ClampLong(10, 5, 60), 10
End Sub

' ----------------------------------------------------------------------------
' G16 SafeLeft(14章§6 modUtil節)
'   契約: 「先頭n字で切る。n <= 0 は ""。末尾に単独の高位サロゲートを残さない」。
'   下の 2,000 / 400 は代表的な切り出し長であって、表示セルの2,000字(16章E-22)
'   やナレッジ行の400字(15章§0.7 -> modKnowledgeFmt.TrimKbLine)の実装を
'   検査するものではない。
' ----------------------------------------------------------------------------
Private Sub T_HeadTrim()
    Dim src As String

    src = RepChar("a", 3000)
    ChkN "SafeLeft_2000字超は2000字_14章§6modUtil", _
        Len(modUtil.SafeLeft(src, CUT_2000)), CUT_2000

    src = RepChar("a", CUT_2000)
    ChkN "SafeLeft_ちょうどの長さは切らない_14章§6modUtil", _
        Len(modUtil.SafeLeft(src, CUT_2000)), CUT_2000

    ChkS "SafeLeft_短い本文は素通し_14章§6modUtil", _
        modUtil.SafeLeft("あいうえお", CUT_2000), "あいうえお"

    ChkS "SafeLeft_空文字は空文字_14章§6modUtil", _
        modUtil.SafeLeft("", CUT_2000), ""

    src = RepChar("b", 900)
    ChkN "SafeLeft_400字で切る_14章§6modUtil", _
        Len(modUtil.SafeLeft(src, CUT_400)), CUT_400

    ChkS "SafeLeft_切らない長さでは先頭が保たれる_14章§6modUtil", _
        Left(modUtil.SafeLeft("[M-0012] 食品工場リスク診断サービス", CUT_400), 8), _
        "[M-0012]"
End Sub

' ----------------------------------------------------------------------------
' G17 BufInit / BufAdd / BufText(14章§6 modUtil節)
'   契約: 「行バッファ(16章 E-26)。BufText の区切りは vbLf。itemCount <= 0 は ""」。
'   ここで見るのは modUtil の積み上げと連結だけであり、ナレッジ注入テキストの
'   書式(15章§6.1「1行1件・行頭 [ID] 」)は modKnowledgeFmt が実装する。
' ----------------------------------------------------------------------------
Private Sub T_LineBuffer()
    Dim buf() As String
    Dim cnt As Long
    Dim txt As String
    Dim p1 As Long
    Dim p2 As Long
    Dim p3 As Long

    modUtil.BufInit buf, cnt
    ChkN "BufInit_初期化直後の件数は0_14章§6modUtil", cnt, 0
    ChkS "BufText_初期化直後は空文字_14章§6modUtil", modUtil.BufText(buf, cnt), ""

    modUtil.BufAdd buf, cnt, "[M-0012] 食品工場リスク診断サービス"
    modUtil.BufAdd buf, cnt, "[L-03] 生産物賠償責任保険(PL保険)"
    modUtil.BufAdd buf, cnt, "[S-0004] 見守りヤモリ型(P2)"
    ChkN "BufAdd_追加件数が一致_14章§6modUtil", cnt, 3

    txt = modUtil.BufText(buf, cnt)
    p1 = InStr(txt, "[M-0012]")
    p2 = InStr(txt, "[L-03]")
    p3 = InStr(txt, "[S-0004]")
    ChkB "BufText_追加した全行を含む_14章§6modUtil", _
        (p1 > 0 And p2 > 0 And p3 > 0), "実際=[" & HeadOf(txt) & "]"
    ChkB "BufText_追加順が保たれる_14章§6modUtil", _
        (p1 < p2 And p2 < p3), "実際=[" & HeadOf(txt) & "]"
End Sub
