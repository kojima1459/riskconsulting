Attribute VB_Name = "modUtil"
Option Explicit

' ============================================================================
' modUtil - 汎用ユーティリティ(文字列以外の道具・時刻・配列・バッファ)
' ----------------------------------------------------------------------------
' 役割:
'   どの層からも使う小さな道具を1箇所に集める。文字列の【無害化・整形】は
'   modUtilText の責務なので、こちらには置かない(答えを2箇所に書かない)。
'
' 移植元: PoC「マイ本棚AI」 src/core/modUtil.bas。
'   移植したのは SafeLeft / SplitKeepNonEmpty / NowStamp と、そこから
'   派生したセル分割・バッファ連結のみ。RPNで使わない機能
'   (ベクトル演算 VectorToCsv/DotProduct/L2Normalize、ページ結合
'   JoinPagedText/SplitPagedText、難読化 DeobfuscateSecret、COMエラー説明
'   DescribeComError、進捗文言 EtaText/ProgressText 等)は 12章§2の
'   「必要関数のみ」に従い非移植とした。とくに DeobfuscateSecret は
'   16章NFR-S2が「難読化してブックに埋め込む方式は使わない」と明記した
'   ため、機構ごと持ち込まない。
'   Fnv1a64Hex / NormalizeForHash は 14章§6が modUtilText の関数として
'   宣言しているため、PoCの modUtil からは移さず modUtilText 側へ置いた。
'
' R4準拠: Excelトークン(Worksheets/Range/Application/ThisWorkbook/MsgBox/
'   ActiveSheet)には触れない。ヘッダ探索(13章§6)も「セルを読む」のではなく
'   「読み終えた見出し行の配列を受け取る」形にして純関数に保つ。
' ============================================================================

' 1セルに入れてよい最大字数(16章 E-22)。Excelの物理上限は32,767字だが、
' 「'」前置ぶんと余白を見て 32,000 で切る(13章§2.2・16章 NFR-S7①と同値)。
Private Const CELL_CHUNK_CHARS As Long = 32000

' Timerが日跨ぎで0へ戻ったときの補正量(1日=86,400,000ミリ秒)。
Private Const MS_PER_DAY_UTIL As Double = 86400000#

' 注入ナレッジID等を1セルに連結するときの区切り(13章§2.4 injected_kb_ids)。
Private Const ID_LIST_SEP As String = ";"

' ============================================================================
' SafeLeft - 先頭n字で切り詰める。ただし末尾に単独の高位サロゲートを残さない。
' ----------------------------------------------------------------------------
'   VBAの Left$ はUTF-16コードユニット単位で切るため、サロゲートペアの
'   途中で切ると【不正な半端文字】が残る。これをそのままセルやログへ書くと
'   Excelの保存時に化け、ファイルを開き直すまで気付けない。切り口が高位
'   サロゲート(&HD800-&HDBFF)なら1文字余分に落とす。
' ============================================================================
Public Function SafeLeft(ByVal s As String, ByVal n As Long) As String
    Dim lim As Long
    lim = n
    If lim < 0 Then lim = 0
    If Len(s) <= lim Then
        SafeLeft = s
        Exit Function
    End If
    SafeLeft = Left$(s, lim)
    If lim = 0 Then Exit Function

    Dim code As Long
    code = AscW(Right$(SafeLeft, 1))
    If code < 0 Then code = code + 65536
    If code >= &HD800& And code <= &HDBFF& Then
        SafeLeft = Left$(SafeLeft, lim - 1)
    End If
End Function

' ============================================================================
' SplitKeepNonEmpty - 区切って空要素を捨てる。0件なら0要素配列を返す。
' ----------------------------------------------------------------------------
'   ReDim x(0 To -1) は実機Excel VBAで実行時エラー9になるため、0件は
'   Split(vbNullString) で表現する(この式だけが実機・LO双方で0要素配列)。
' ============================================================================
Public Function SplitKeepNonEmpty(ByVal s As String, ByVal sep As String) As String()
    If LenB(sep) = 0 Then
        SplitKeepNonEmpty = Split(vbNullString)
        Exit Function
    End If

    Dim raw() As String
    raw = Split(s, sep)
    Dim maxN As Long
    maxN = UBound(raw) - LBound(raw) + 1
    If maxN <= 0 Then
        SplitKeepNonEmpty = Split(vbNullString)
        Exit Function
    End If

    Dim outArr() As String
    ReDim outArr(0 To maxN - 1)
    Dim cnt As Long
    cnt = 0
    Dim i As Long
    For i = LBound(raw) To UBound(raw)
        If LenB(Trim$(raw(i))) > 0 Then
            outArr(cnt) = Trim$(raw(i))
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        SplitKeepNonEmpty = Split(vbNullString)
    Else
        ReDim Preserve outArr(0 To cnt - 1)
        SplitKeepNonEmpty = outArr
    End If
End Function

' ============================================================================
' SplitForCells / JoinCellChunks - 1セル上限を超える本文の分割保存と結合
'   (13章§2.2 case_data・16章 E-22)。
' ----------------------------------------------------------------------------
'   分割は SafeLeft と同じサロゲート安全性を持たせる(切れ目でペアを割らない)。
'   このため各断片は chunkLen ちょうどではなく chunkLen-1 になることがある。
'   結合は単純連結であり、【分割→結合で元の文字列に一字も欠けず戻る】ことが
'   契約(往復テストの対象)。空文字は「1件の空断片」ではなく0件を返す。
' ============================================================================
Public Function SplitForCells(ByVal s As String, ByVal chunkLen As Long) As String()
    Dim lim As Long
    lim = chunkLen
    If lim <= 0 Then lim = CELL_CHUNK_CHARS
    If LenB(s) = 0 Then
        SplitForCells = Split(vbNullString)
        Exit Function
    End If

    ' 断片数の上限見積り(必ず実際の件数以上になる)。
    Dim maxN As Long
    maxN = (Len(s) \ 1) + 1
    If lim > 1 Then maxN = (Len(s) \ (lim - 1)) + 2

    Dim outArr() As String
    ReDim outArr(0 To maxN - 1)
    Dim cnt As Long
    cnt = 0

    Dim rest As String
    rest = s
    Do While Len(rest) > 0
        Dim piece As String
        piece = SafeLeft(rest, lim)
        If Len(piece) = 0 Then
            ' 進まない(理論上は起きない)ので無限ループを避けて打切る。
            Exit Do
        End If
        If cnt > UBound(outArr) Then
            ReDim Preserve outArr(0 To cnt)
        End If
        outArr(cnt) = piece
        cnt = cnt + 1
        rest = Mid$(rest, Len(piece) + 1)
    Loop

    If cnt = 0 Then
        SplitForCells = Split(vbNullString)
    Else
        ReDim Preserve outArr(0 To cnt - 1)
        SplitForCells = outArr
    End If
End Function

Public Function JoinCellChunks(ByRef parts() As String) As String
    Dim lo As Long, hi As Long
    If Not ArrayBounds(parts, lo, hi) Then
        JoinCellChunks = vbNullString
        Exit Function
    End If
    JoinCellChunks = Join(parts, vbNullString)
End Function

' ============================================================================
' バッファ連結(16章 E-26: 32bitメモリ制約)
' ----------------------------------------------------------------------------
'   `s = s & piece` を数千回繰り返すと、そのたびに新しい文字列を確保し直す
'   ため 32bit Excel では実メモリを使い切って「メモリが不足しています」で
'   落ちる。断片を配列へ溜めて最後に一度だけ Join する方式に統一する。
'
'   使い方:
'     Dim buf() As String, n As Long
'     modUtil.BufInit buf, n
'     modUtil.BufAdd buf, n, "..."
'     result = modUtil.BufText(buf, n)
' ============================================================================
Public Sub BufInit(ByRef buf() As String, ByRef itemCount As Long)
    ReDim buf(0 To 63)
    itemCount = 0
End Sub

Public Sub BufAdd(ByRef buf() As String, ByRef itemCount As Long, ByVal s As String)
    Dim lo As Long, hi As Long
    If Not ArrayBounds(buf, lo, hi) Then
        ReDim buf(0 To 63)
        hi = 63
    End If
    If itemCount > hi Then
        ReDim Preserve buf(0 To (hi + 1) * 2 - 1)
    End If
    buf(itemCount) = s
    itemCount = itemCount + 1
End Sub

Public Function BufText(ByRef buf() As String, ByVal itemCount As Long) As String
    If itemCount <= 0 Then
        BufText = vbNullString
        Exit Function
    End If
    Dim tmp() As String
    ReDim tmp(0 To itemCount - 1)
    Dim i As Long
    For i = 0 To itemCount - 1
        tmp(i) = buf(i)
    Next i
    BufText = Join(tmp, vbNullString)
End Function

' ============================================================================
' FindHeaderCol - 見出し行(1行ぶんの値の配列)から列名の位置を探す。
' ----------------------------------------------------------------------------
'   13章§6「VBAアクセスは列名ベース(列番号ハードコード禁止)」の実体。
'   R4のため本モジュールはセルを読めないので、呼び出し側(store系)が
'   見出し行を配列で読み込み、その配列を渡す形にしてある。
'   戻り値は【配列の添字ではなく1始まりの列位置】(Cells(r, col) に直接
'   渡せる値)。見つからなければ0。比較は前後空白を無視した大小文字非依存。
' ============================================================================
Public Function FindHeaderCol(ByVal headerRow As Variant, ByVal headerName As String) As Long
    On Error GoTo NotFound

    Dim wanted As String
    wanted = LCase$(Trim$(headerName))
    If LenB(wanted) = 0 Then Exit Function

    Dim dims As Long
    dims = VariantDimCount(headerRow)
    Dim c As Long, lo1 As Long, hi1 As Long

    If dims = 2 Then
        ' Range.Value 由来(1始まりの2次元配列。行は1行ぶんのみを想定)
        lo1 = LBound(headerRow, 2)
        hi1 = UBound(headerRow, 2)
        For c = lo1 To hi1
            If LCase$(Trim$(CStr(headerRow(LBound(headerRow, 1), c)))) = wanted Then
                FindHeaderCol = c - lo1 + 1
                Exit Function
            End If
        Next c
    ElseIf dims = 1 Then
        ' Array(...) 由来(1次元。0始まり/1始まりのどちらでもよい)
        lo1 = LBound(headerRow)
        hi1 = UBound(headerRow)
        For c = lo1 To hi1
            If LCase$(Trim$(CStr(headerRow(c)))) = wanted Then
                FindHeaderCol = c - lo1 + 1
                Exit Function
            End If
        Next c
    End If
    Exit Function
NotFound:
    FindHeaderCol = 0
End Function

' 配列の次元数(1 or 2)。配列でなければ0。
Private Function VariantDimCount(ByVal arr As Variant) As Long
    Dim n As Long
    n = 0
    If Not IsArray(arr) Then
        VariantDimCount = 0
        Exit Function
    End If
    On Error GoTo Done
    Dim probe As Long
    probe = UBound(arr, 1)
    n = 1
    probe = UBound(arr, 2)
    n = 2
Done:
    VariantDimCount = n
End Function

' ============================================================================
' AppendIdList - ";"区切りのID列へ、重複しないよう1件足す。
' ----------------------------------------------------------------------------
'   run_log.injected_kb_ids(13章§2.4)と modKnowledge.LastInjectedIds の
'   蓄積に使う。空IDは足さない。既にある同一IDは足さない(順序は保つ)。
' ============================================================================
Public Function AppendIdList(ByVal listText As String, ByVal idText As String) As String
    Dim one As String
    one = Trim$(idText)
    If LenB(one) = 0 Then
        AppendIdList = listText
        Exit Function
    End If
    If LenB(listText) = 0 Then
        AppendIdList = one
        Exit Function
    End If

    Dim items() As String
    items = SplitKeepNonEmpty(listText, ID_LIST_SEP)
    Dim lo As Long, hi As Long
    If ArrayBounds(items, lo, hi) Then
        Dim i As Long
        For i = lo To hi
            If StrComp(items(i), one, vbTextCompare) = 0 Then
                AppendIdList = listText
                Exit Function
            End If
        Next i
    End If
    AppendIdList = listText & ID_LIST_SEP & one
End Function

' ============================================================================
' ClampLong - 値を [minV, maxV] へ収める。minV > maxV の指定は minV を優先。
' ============================================================================
Public Function ClampLong(ByVal v As Long, ByVal minV As Long, ByVal maxV As Long) As Long
    Dim r As Long
    r = v
    If r > maxV Then r = maxV
    If r < minV Then r = minV
    ClampLong = r
End Function

' ============================================================================
' ElapsedMsSince - Timer基準の経過ミリ秒(日跨ぎ補正つき)。
' ----------------------------------------------------------------------------
'   VBAの Timer は「その日の午前0時からの秒数」なので、日付が変わると0へ
'   戻り、素の (Timer - t0) が大きな負値になる。latency_ms が負になると
'   run_log の集計が壊れるため、負なら1日ぶんを足して実経過へ戻す。
'   移植元(PoC modUtilText.ElapsedMsSince)と同じ扱いだが、RPNでは
'   「文字列以外の道具」を modUtil に寄せる方針のためこちらへ置いた。
' ============================================================================
Public Function ElapsedMsSince(ByVal t0 As Double) As Double
    Dim ms As Double
    ms = (Timer - t0) * 1000#
    If ms < 0 Then ms = ms + MS_PER_DAY_UTIL   ' 日跨ぎ(Timerが0へ戻った)
    If ms < 0 Then ms = 0                      ' それでも負なら0(時計の巻戻し)
    ElapsedMsSince = ms
End Function

' ============================================================================
' NowStamp - "yyyy-mm-dd hh:nn:ss"。ログの logged_at / run_at はこれで書く。
'   Format$ を使わないのは、和暦カレンダー設定の端末で年が元号年になり、
'   ログの並び順と突合が壊れるため(正は modUtilText.IsoDateTime)。
' ============================================================================
Public Function NowStamp() As String
    NowStamp = modUtilText.IsoDateTime(Now)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------
' 未初期化配列に UBound を掛けると実行時エラー9になる。境界の取得は必ず
' ここを通し、未初期化なら False を返す(呼び出し側で If 1本にする)。
Private Function ArrayBounds(ByRef arr() As String, ByRef lo As Long, ByRef hi As Long) As Boolean
    On Error GoTo NotReady
    lo = LBound(arr)
    hi = UBound(arr)
    ArrayBounds = (hi >= lo)
    Exit Function
NotReady:
    lo = 0
    hi = -1
    ArrayBounds = False
End Function
