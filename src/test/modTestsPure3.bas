Attribute VB_Name = "modTestsPure3"
Option Explicit

' ============================================================================
' modTestsPure3 - W2a(T-20 modCaseStore / T-21 modKnowledge)の純ロジックテスト
'                 (17章§4-1 層(a))G12-G17
' ----------------------------------------------------------------------------
' 役割:
'   W2a が土台にする「純ロジックの契約」を、実装コードを一切読まずに
'   仕様(13章§2.2 / 15章§0.7・§3・§4・§6.1 / 16章 E-09/E-22 / 14章§6 /
'   17章 T-20・T-21 の DoD)だけを根拠に固定する。
'
'   入口は Public Sub RunAll()。modTestsPure2 と同じく、ランナーからは直接
'   呼ばれない前提で書いてある(結線は統合者の担当。下の「統合者へ」を参照)。
'
' 統合者へ(本ファイル単体では1本も実行されない。3点セットで結線すること):
'   (1) modTestsPure2.RunAll の末尾から modTestsPure3.RunAll を呼ぶ
'       (modTestsPure.RunAll -> modTestsPure2.RunAll と同じ数珠つなぎ)。
'   (2) build/modules.json へ 1件追加(name=modTestsPure3 /
'       path=src/test/modTestsPure3.bas / role=test / type=std / wave=T-20)。
'   (3) wintest/tests_expected.txt を 115 -> 184 へ更新。
'   tools/run_lo_tests.py の PURE_ALLOWLIST には modTestsPure3 が既にある。
'
' 設計判断(R4準拠): Worksheets / Range / Application / ThisWorkbook / MsgBox /
'   ActiveSheet には一切触れない。改行は vbLf 基準。
'   **modCaseStore / modKnowledge を参照しない**: 両モジュールの14章§6の公開
'   関数はすべてシートI/Oを伴い、tools/run_lo_tests.py の PURE_ALLOWLIST にも
'   載っていない(=層(a)へ注入されない)。したがって採番・状態遷移・
'   ResolveStepJson・FreezeRound・整形・実在チェック・切詰めの「本体」は
'   層(b)(wintest)の担当であり、本ファイルはそれらが依存する純プリミティブの
'   契約だけを固定する。積み残しは報告書の concerns に列挙した。
'
' 本ファイルのテスト本数: 69本(G12 40 / G13 8 / G14 4 / G15 6 / G16 6 / G17 5)
'   ※ wintest/tests_expected.txt を +69 すること(現行115本 -> 184本)。
'
' グループ / 本数 / 根拠章:
'   G12 case_data の32,000字分割・結合   40本 13章§2.2 / 16章E-22 / T-20 DoD
'   G13 セル格納規約「; 」の分割          8本 13章§2.2 セル格納規約
'   G14 注入IDの累積(LastInjectedIds)     4本 14章§6 / T-21 DoD
'   G15 ナレッジ切詰めの下限・上限        6本 15章§0.7
'   G16 先頭切詰め(表示2,000字/行400字)   6本 16章E-22 / 15章§0.7
'   G17 注入テキストの行バッファ          5本 15章§3・§4・§6.1(1行1件)
'
' 本ファイルが前提とする公開契約(すべて既存の modUtil。W1 T-10 の成果物):
'   modUtil.SplitForCells / JoinCellChunks   ' T-10 DoD「分割結合ユーティリティ緑」
'   modUtil.SplitKeepNonEmpty / AppendIdList
'   modUtil.ClampLong / SafeLeft
'   modUtil.BufInit / BufAdd / BufText
'   modTestRunner.Check
' ============================================================================

' 13章§2.2: case_data の content は最大32,000字。分割保存の単位。
Private Const CH_MAX As Long = 32000
' 16章E-22: セルの物理上限は32,767字(32,000字契約はこの内側に取ってある)。
Private Const CELL_MAX As Long = 32767
' 16章E-22 表示側: 表示セルは先頭2,000字+注記。
Private Const DISP_MAX As Long = 2000
' 15章§0.7: 5段の切詰めでなお超過するときは各行を先頭400字で切る。
Private Const KB_LINE_MAX As Long = 400

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
' G13 セル格納規約「; 」の分割
'   根拠: 13章§2.2 セル格納規約「要素を『; 』(半角セミコロン+半角スペース)で
'   連結した1セル」「『; 』で分割し前後空白を除去。空セルは []」。
'   逆シリアライズ本体(modUICase)はシートI/Oのため層(a)対象外。ここでは
'   その土台の分割プリミティブだけを固定する。
' ----------------------------------------------------------------------------
Private Sub T_CellList()
    Dim items() As String

    items = modUtil.SplitKeepNonEmpty("", "; ")
    ChkN "セル格納_空セルは要素0件_13章§2.2空配列規約", ArrN(items), 0

    items = modUtil.SplitKeepNonEmpty("冷凍食品; 惣菜; 調味料", "; ")
    ChkN "セル格納_文字列配列3要素の件数_13章§2.2", ArrN(items), 3
    ChkS "セル格納_文字列配列の1要素目_13章§2.2", ArrAt(items, 0), "冷凍食品"
    ChkS "セル格納_文字列配列の2要素目_13章§2.2", ArrAt(items, 1), "惣菜"
    ChkS "セル格納_文字列配列の3要素目_13章§2.2", ArrAt(items, 2), "調味料"

    items = modUtil.SplitKeepNonEmpty("M-0012", "; ")
    ChkN "セル格納_単一要素は1件_13章§2.2", ArrN(items), 1

    ' 整数配列(target_risk_nos / target_gap_nos)も同じ「; 」連結(例 `2; 5`)。
    items = modUtil.SplitKeepNonEmpty("2; 5", "; ")
    ChkN "セル格納_整数配列2要素の件数_13章§2.2", ArrN(items), 2
    ChkS "セル格納_整数配列の2要素目_13章§2.2", ArrAt(items, 1), "5"
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
' G15 ナレッジ切詰めの下限・上限
'   根拠: 15章§0.7 の切詰め表。成功事例・型ライブラリは0行まで可、メニュー・
'   種目・リスクライブラリは5行が下限、各行数上限は config
'   (kb_menu_rows 既定60 / kb_risk_rows 既定20 / kb_case_rows 既定5)。
'   段階適用のループ本体は modKnowledge / modPipeline 側(シートI/O)のため
'   層(a)対象外。行数を範囲へ収めるプリミティブだけを固定する。
' ----------------------------------------------------------------------------
Private Sub T_KbTrimFloor()
    ChkN "切詰め下限_メニューは5行を下回らない_15章§0.7", _
        modUtil.ClampLong(2, 5, 60), 5
    ChkN "切詰め下限_種目は5行を下回らない_15章§0.7", _
        modUtil.ClampLong(3, 5, 60), 5
    ChkN "切詰め下限_リスクライブラリは5行を下回らない_15章§0.7", _
        modUtil.ClampLong(1, 5, 20), 5
    ChkN "切詰め下限_成功事例は0行まで可_15章§0.7", _
        modUtil.ClampLong(0, 0, 5), 0
    ChkN "切詰め上限_kb_menu_rowsの60行を超えない_15章§0.7", _
        modUtil.ClampLong(80, 5, 60), 60
    ChkN "切詰め_範囲内の行数はそのまま_15章§0.7", _
        modUtil.ClampLong(10, 5, 60), 10
End Sub

' ----------------------------------------------------------------------------
' G16 先頭切詰め
'   根拠: 16章E-22 表示側「表示セルは先頭2,000字+注記」/ 15章§0.7
'   「5まで適用してなお超過する場合は、各行を先頭400字で切り『...』を付す」。
' ----------------------------------------------------------------------------
Private Sub T_HeadTrim()
    Dim src As String

    src = RepChar("a", 3000)
    ChkN "表示切詰め_2000字超は2000字_16章E22", _
        Len(modUtil.SafeLeft(src, DISP_MAX)), DISP_MAX

    src = RepChar("a", DISP_MAX)
    ChkN "表示切詰め_2000字ちょうどは切らない_16章E22", _
        Len(modUtil.SafeLeft(src, DISP_MAX)), DISP_MAX

    ChkS "表示切詰め_短い本文は素通し_16章E22", _
        modUtil.SafeLeft("あいうえお", DISP_MAX), "あいうえお"

    ChkS "表示切詰め_空文字は空文字_16章E22", _
        modUtil.SafeLeft("", DISP_MAX), ""

    src = RepChar("b", 900)
    ChkN "ナレッジ行切詰め_400字で切る_15章§0.7", _
        Len(modUtil.SafeLeft(src, KB_LINE_MAX)), KB_LINE_MAX

    ChkS "ナレッジ行切詰め_行頭のIDが保たれる_15章§0.7", _
        Left(modUtil.SafeLeft("[M-0012] 食品工場リスク診断サービス", KB_LINE_MAX), 8), _
        "[M-0012]"
End Sub

' ----------------------------------------------------------------------------
' G17 注入テキストの行バッファ
'   根拠: 15章§6.1 共通規約「1行1件、行頭は `[ID] `」/ 15章§3・§4 の整形例。
'   整形本体(RiskLibFor / MenusFor / LinesText 等)はナレッジシート読込を伴う
'   ため層(a)対象外。複数行を順序どおり積む土台だけを固定する。
'   区切り文字(vbLf か否か)は14章§6にも15章にも規定が無いため、ここでは
'   含有と順序だけを見る(concerns 参照)。
' ----------------------------------------------------------------------------
Private Sub T_LineBuffer()
    Dim buf() As String
    Dim cnt As Long
    Dim txt As String
    Dim p1 As Long
    Dim p2 As Long
    Dim p3 As Long

    modUtil.BufInit buf, cnt
    ChkN "注入行バッファ_初期化直後の件数は0_15章§6.1", cnt, 0
    ChkS "注入行バッファ_初期化直後は空文字_15章§6.1", modUtil.BufText(buf, cnt), ""

    modUtil.BufAdd buf, cnt, "[M-0012] 食品工場リスク診断サービス"
    modUtil.BufAdd buf, cnt, "[L-03] 生産物賠償責任保険(PL保険)"
    modUtil.BufAdd buf, cnt, "[S-0004] 見守りヤモリ型(P2)"
    ChkN "注入行バッファ_追加件数が一致_15章§6.1", cnt, 3

    txt = modUtil.BufText(buf, cnt)
    p1 = InStr(txt, "[M-0012]")
    p2 = InStr(txt, "[L-03]")
    p3 = InStr(txt, "[S-0004]")
    ChkB "注入行バッファ_追加した全行を含む_15章§6.1", _
        (p1 > 0 And p2 > 0 And p3 > 0), "実際=[" & HeadOf(txt) & "]"
    ChkB "注入行バッファ_追加順が保たれる_15章§6.1", _
        (p1 < p2 And p2 < p3), "実際=[" & HeadOf(txt) & "]"
End Sub
