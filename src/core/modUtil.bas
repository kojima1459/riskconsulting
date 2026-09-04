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

' ファイル属性のディレクトリビット(vbDirectory)。定数名を書かず数値で持つ
' (LibreOffice側の構文チェックで未定義名にしないための既存の流儀と同じ)。
' 本体と同じフォルダのヒント(modBoot が起動手順(2)で1回だけ入れる)。
' core層は ThisWorkbook を参照できない(12章§4)ため、値は外から預かる。
Private mBookDir As String

Private Const ATTR_DIRECTORY As Long = 16

' 保存先(data_dir)の解決で使う環境変数名と最後の逃げ場(裁定書27 W9-C2)。
Private Const DD_VAR_COMMERCIAL As String = "%OneDriveCommercial%"
Private Const DD_VAR_ONEDRIVE As String = "%OneDrive%"
Private Const DD_VAR_PROFILE As String = "%USERPROFILE%"
Private Const DD_LAST_RESORT_TAIL As String = "\Documents\RPN出力"

' 裁定書28: ランチャー(.bat)が本体と同じフォルダへ書く「data_dir の値そのもの」。
'   環境変数に依存しない最優先の値源で、中身は1行のフルパス。
Public Const DATA_DIR_POINTER_FILE As String = "data_dir.txt"

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
    maxN = Len(s) + 1
    If lim > 1 Then maxN = CLng(Fix(Len(s) / (lim - 1))) + 2

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
' 行バッファ連結(16章 E-26: 32bitメモリ制約)
' ----------------------------------------------------------------------------
'   `s = s & piece` を数千回繰り返すと、そのたびに新しい文字列を確保し直す
'   ため 32bit Excel では実メモリを使い切って「メモリが不足しています」で
'   落ちる。断片を配列へ溜めて最後に一度だけ Join する方式に統一する。
'
'   **BufText の区切りは vbLf**(14章§6。裁定書6 項目8)。用途は「1行1件の
'   注入テキストを順に積む」であり、15章§6.1 の共通規約(1行1件)を満たす
'   最小の道具として契約する。区切り無しで断片を継ぎたい場合は
'   JoinCellChunks(単純連結)を使う(答えを2箇所に書かないため BufText へ
'   区切りの分岐を持たせない)。
'
'   使い方:
'     Dim buf() As String, n As Long
'     modUtil.BufInit buf, n
'     modUtil.BufAdd buf, n, "[M-0012] ..."
'     result = modUtil.BufText(buf, n)   ' 行はvbLfで連結される
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
    BufText = Join(tmp, vbLf)
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

' ============================================================================
' ファイルとフォルダ(裁定書27 W9-B2)
' ----------------------------------------------------------------------------
' なぜ core の汎用道具に置くのか:
'   HTMLレポート(app層 modExportHtml)と[中身を見る]・企業ファイル(ui/app層)が
'   同じ「UTF-8で書く」「無ければフォルダを作る」を必要とする。従来は
'   ADODB.Stream と Scripting.FileSystemObject を各所で生成していたが、この
'   2つのCOM生成は社内AVのAMSIがマクロ型マルウェアの特徴として重く見るため
'   配布物から消す(裁定書27 W9-B2)。代替は **純VBA**(`Open For Binary` と
'   `MkDir`)で、符号化は modUtilText.Utf8Bytes が唯一持つ。
'
' R4(12章§4): Excelトークンは1つも使わない。ファイル操作はVBAの組み込み文で
'   あってExcelのオブジェクトモデルではないため、core層に置いてよい。
' ============================================================================

' フォルダが在るか(ファイルは False)。
Public Function FolderExists(ByVal dirText As String) As Boolean
    On Error GoTo Failed
    If LenB(dirText) = 0 Then Exit Function
    Dim attrVal As Long
    attrVal = GetAttr(dirText)
    FolderExists = ((attrVal And ATTR_DIRECTORY) = ATTR_DIRECTORY)
    Exit Function
Failed:
    FolderExists = False
End Function

' ファイルが在るか(フォルダは False)。
Public Function FileExistsAt(ByVal pathText As String) As Boolean
    On Error GoTo Failed
    If LenB(pathText) = 0 Then Exit Function
    Dim attrVal As Long
    attrVal = GetAttr(pathText)
    FileExistsAt = ((attrVal And ATTR_DIRECTORY) <> ATTR_DIRECTORY)
    Exit Function
Failed:
    FileExistsAt = False
End Function

' 無ければ作る(途中の階層もまとめて作る)。作れたら True。
'   UNC(\\server\share)とURL(OneDriveの "https://...")は階層を作りに行かず、
'   在るかどうかだけを見る(作成の権限が無い場所で例外を積まないため)。
Public Function EnsureFolder(ByVal dirText As String) As Boolean
    On Error GoTo Failed
    Dim t As String
    t = TrimTrailingSep(dirText)
    If LenB(t) = 0 Then Exit Function
    If FolderExists(t) Then
        EnsureFolder = True
        Exit Function
    End If
    If InStr(1, t, "://", vbBinaryCompare) > 0 Then Exit Function

    ' 分割・連結の規則は純関数 SplitPathParts が唯一持つ(W9.2)。区切り文字を
    ' "\" 決め打ちにしない(Mac の実Excel は "/")。
    Dim stepsText As String
    stepsText = SplitPathParts(t, PathSep())
    If LenB(stepsText) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(stepsText, vbLf)

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Not FolderExists(parts(i)) Then MkDir parts(i)
    Next i
    EnsureFolder = FolderExists(t)
    Exit Function
Failed:
    EnsureFolder = False
End Function

' ============================================================================
' PathSep - このプラットフォームのパス区切り("\" か "/")。
'   core層は Application. を参照できない規約(12章§2の層規約)のため
'   Application.PathSeparator は使えない。**CurDir$ の先頭で判定する**
'   (Mac の実Excel は "/Users/..." を返す)。取れなければ Windows 既定の "\"。
'   **環境依存の分岐なので純層では検査できない**。検査できる形の「分割と連結」は
'   下の SplitPathParts が持ち、純層はそちらを見る(W9.2)。
' ============================================================================
Public Function PathSep() As String
    On Error GoTo FallbackSep
    If Left$(CurDir$, 1) = "/" Then
        PathSep = "/"
    Else
        PathSep = "\"
    End If
    Exit Function
FallbackSep:
    PathSep = "\"
End Function

' ============================================================================
' SplitPathParts - EnsureFolder が MkDir する「親から順の一覧」を作る純関数。
'   区切り sep で分割し、先頭要素(ドライブ名など。単独では作らない)から
'   1つずつ足した経路を vbLf 区切りで返す。空の要素は飛ばす。
'   **区切り2つで始まる経路(UNC \\server\share)は空文字を返す**
'   (共有名の途中まで MkDir できないため、EnsureFolder はそこで諦める)。
' ============================================================================
Public Function SplitPathParts(ByVal t As String, ByVal sep As String) As String
    If LenB(t) = 0 Then Exit Function
    If LenB(sep) = 0 Then Exit Function
    If Left$(t, 2) = sep & sep Then Exit Function

    Dim parts() As String
    parts = Split(t, sep)
    If UBound(parts) < LBound(parts) Then Exit Function

    Dim built As String
    Dim out As String
    Dim i As Long
    built = parts(LBound(parts))
    For i = LBound(parts) + 1 To UBound(parts)
        If LenB(parts(i)) > 0 Then
            built = built & sep & parts(i)
            If LenB(out) > 0 Then out = out & vbLf
            out = out & built
        End If
    Next i
    SplitPathParts = out
End Function

' 末尾の "\" と "/" を落とす(パスの連結を1箇所に保つための小道具)。
Public Function TrimTrailingSep(ByVal pathText As String) As String
    Dim t As String
    t = Trim$(pathText)
    Do While Len(t) > 0
        If Right$(t, 1) = "\" Or Right$(t, 1) = "/" Then
            t = Left$(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    TrimTrailingSep = t
End Function

' UTF-8 でファイルへ書く(withBom=True で EF BB BF を先頭に置く)。
'   既存ファイルは消してから作り直す(Binary は上書きで前の残骸が残るため)。
Public Function WriteUtf8File(ByVal pathText As String, ByVal bodyText As String, _
                              ByVal withBom As Boolean) As Boolean
    Dim fileNo As Long
    On Error GoTo Failed
    If LenB(pathText) = 0 Then Exit Function

    Dim n As Long
    n = modUtilText.Utf8Len(bodyText, withBom)

    If FileExistsAt(pathText) Then Kill pathText

    Dim buf() As Byte
    fileNo = FreeFile
    Open pathText For Binary Access Write As #fileNo
    If n > 0 Then
        buf = modUtilText.Utf8Bytes(bodyText, withBom)
        Put #fileNo, 1, buf
    End If
    Close #fileNo
    WriteUtf8File = True
    Exit Function
Failed:
    CloseQuiet fileNo
    WriteUtf8File = False
End Function

' UTF-8 のファイルを読む(裁定書28。設定.txt。復号は modUtilText.Utf8Text)。
'   読めなければ空文字(読めない設定ファイルで起動を止めない)。
Public Function ReadUtf8File(ByVal pathText As String) As String
    Dim fileNo As Long
    On Error GoTo Failed
    If LenB(pathText) = 0 Then Exit Function
    If Not FileExistsAt(pathText) Then Exit Function

    Dim n As Long
    Dim buf() As Byte
    fileNo = FreeFile
    Open pathText For Binary Access Read As #fileNo
    n = LOF(fileNo)
    If n > 0 Then
        ReDim buf(0 To n - 1)
        Get #fileNo, 1, buf
    End If
    Close #fileNo
    If n <= 0 Then Exit Function
    ReadUtf8File = modUtilText.Utf8Text(buf, n)
    Exit Function
Failed:
    CloseQuiet fileNo
    ReadUtf8File = vbNullString
End Function

' 開いたままのファイル番号を黙って閉じる(ハンドラ稼働中に On Error Resume Next
'   を書けないため、後始末は別Subへ切り出す)。
Private Sub CloseQuiet(ByVal fileNo As Long)
    On Error Resume Next
    If fileNo > 0 Then Close #fileNo
End Sub

' ============================================================================
' 保存先(data_dir)の解決(裁定書27 W9-C2)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   会社PCの `D:` はシャットダウンで消える。`%USERPROFILE%\Documents` が残るか
'   はOneDriveのリダイレクト設定次第で、未測定である(裁定書27 事実)。企業
'   ファイル・HTMLレポート・ヒアリングシートを既定で **OneDrive(会社)** の下
'   へ置き、そこが無いときだけ Documents へ落とす。落ちたことは黙らせず、
'   ナビのお知らせで警告する(11章§8.6 禁忌: 黙って別の場所へ書かない)。
'
' 解決の順(裁定書28 で(0)を追加):
'   (0) **本体と同じフォルダの data_dir.txt の1行目**(ランチャーが書く。環境
'       変数に依存しない実測値なので最優先。裁定書28「裁定の確定」3)
'   (1) config `data_dir` を展開したもの(既定 %OneDriveCommercial%\...)
'   (2) (1)が %OneDriveCommercial% を含むときだけ、それを %OneDrive% に
'       読み替えたもの(個人用OneDriveしか無い端末の救済)
'   (3) %USERPROFILE%\Documents\RPN出力
'   環境変数が空の候補は最初から並べない(展開できない "%" を残さない)。
'
' 候補の並べ方は純関数 DataDirCandidates が唯一持ち(層(a)でテストする)、
' 実在確認とフォルダ作成だけを ResolveDataDir が行う(modBoot.ResolveKbPath と
' 同じ「純部+実在確認」の切り分け)。
' ============================================================================

' 候補の並び(vbLf 区切り。純関数)。env は Environ$ の値をそのまま渡す。
Public Function DataDirCandidates(ByVal pointerRaw As String, _
                                  ByVal configRaw As String, _
                                  ByVal envCommercial As String, _
                                  ByVal envOneDrive As String, _
                                  ByVal envUserProfile As String) As String
    Dim outText As String
    Dim raw As String

    ' (0) 裁定書28: 本体と同じフォルダの data_dir.txt の値。ランチャーが
    '     書いた実測値なので、config よりも環境変数よりも先に置く。空なら
    '     何も足さない(=従来の順のまま)。
    outText = AppendCandidate(outText, TrimTrailingSep(pointerRaw))

    raw = TrimTrailingSep(configRaw)

    If LenB(raw) > 0 Then
        outText = AppendCandidate(outText, _
            ExpandDirVars(raw, envCommercial, envOneDrive, envUserProfile))
        If InStr(1, raw, DD_VAR_COMMERCIAL, vbTextCompare) > 0 Then
            outText = AppendCandidate(outText, ExpandDirVars( _
                Replace(raw, DD_VAR_COMMERCIAL, DD_VAR_ONEDRIVE, 1, -1, vbTextCompare), _
                envCommercial, envOneDrive, envUserProfile))
        End If
    End If

    If LenB(envUserProfile) > 0 Then
        outText = AppendCandidate(outText, _
            TrimTrailingSep(envUserProfile) & DD_LAST_RESORT_TAIL)
    End If
    DataDirCandidates = outText
End Function

' 候補を1つ足す(空・"%"が残っているもの・既出は足さない)。
Private Function AppendCandidate(ByVal listText As String, ByVal candidate As String) As String
    AppendCandidate = listText
    If LenB(candidate) = 0 Then Exit Function
    If InStr(1, candidate, "%", vbBinaryCompare) > 0 Then Exit Function
    Dim probe As String
    probe = vbLf & listText & vbLf
    If InStr(1, probe, vbLf & candidate & vbLf, vbTextCompare) > 0 Then Exit Function
    If LenB(listText) = 0 Then
        AppendCandidate = candidate
    Else
        AppendCandidate = listText & vbLf & candidate
    End If
End Function

' 3つの環境変数だけを展開する(未知の "%..%" は残し、候補から外す材料にする)。
Private Function ExpandDirVars(ByVal pathText As String, ByVal envCommercial As String, _
                               ByVal envOneDrive As String, _
                               ByVal envUserProfile As String) As String
    Dim t As String
    t = pathText
    If LenB(envCommercial) > 0 Then
        t = Replace(t, DD_VAR_COMMERCIAL, TrimTrailingSep(envCommercial), 1, -1, vbTextCompare)
    End If
    If LenB(envOneDrive) > 0 Then
        t = Replace(t, DD_VAR_ONEDRIVE, TrimTrailingSep(envOneDrive), 1, -1, vbTextCompare)
    End If
    If LenB(envUserProfile) > 0 Then
        t = Replace(t, DD_VAR_PROFILE, TrimTrailingSep(envUserProfile), 1, -1, vbTextCompare)
    End If
    ExpandDirVars = TrimTrailingSep(t)
End Function

' 解決した保存先がOneDriveの下かどうか(純関数)。ナビの警告の要否はこれで決める。
Public Function IsUnderOneDrive(ByVal dirText As String, ByVal envCommercial As String, _
                                ByVal envOneDrive As String) As Boolean
    If LenB(dirText) = 0 Then Exit Function
    If LenB(envCommercial) > 0 Then
        If InStr(1, dirText, TrimTrailingSep(envCommercial), vbTextCompare) = 1 Then
            IsUnderOneDrive = True
            Exit Function
        End If
    End If
    If LenB(envOneDrive) > 0 Then
        If InStr(1, dirText, TrimTrailingSep(envOneDrive), vbTextCompare) = 1 Then
            IsUnderOneDrive = True
        End If
    End If
End Function

' 実在確認つきの解決。使える(作れた)最初の候補を返す。どれも駄目なら ""。
Public Function ResolveDataDir(ByVal configRaw As String, _
                               Optional ByVal bookDir As String = vbNullString) As String
    Dim listText As String
    Dim baseDir As String
    baseDir = bookDir
    If LenB(baseDir) = 0 Then baseDir = BookDirHint()

    listText = DataDirCandidates(ReadDataDirPointer(baseDir), configRaw, _
                                 Environ$("OneDriveCommercial"), _
                                 Environ$("OneDrive"), Environ$("USERPROFILE"))
    If LenB(listText) = 0 Then Exit Function

    Dim cands() As String
    cands = Split(listText, vbLf)

    Dim i As Long
    For i = LBound(cands) To UBound(cands)
        If EnsureFolder(cands(i)) Then
            ResolveDataDir = cands(i)
            Exit Function
        End If
    Next i
End Function

' 解決した保存先がOneDriveの下でないとき True(ナビのお知らせに warn を出す)。
Public Function DataDirNotOneDrive(ByVal resolvedDir As String) As Boolean
    If LenB(resolvedDir) = 0 Then Exit Function
    DataDirNotOneDrive = Not IsUnderOneDrive(resolvedDir, Environ$("OneDriveCommercial"), _
                                             Environ$("OneDrive"))
End Function

' ============================================================================
' data_dir ポインタ(裁定書28「裁定の確定」3)
' ----------------------------------------------------------------------------
' ランチャー(.bat)は本体を D: へ写したあと、`%DST%\data_dir.txt` へ
' `%SRC%データ`(=OneDrive の配布フォルダの下の「データ」)を1行だけ書く。
' 本体は環境変数を一切見ずにこの1行を読めばよい(会社アカウントの OneDrive の
' 環境変数名が端末ごとに違っても壊れない)。
'
' 読み方の規約: **1行目だけ**を読み、前後の空白と改行を落とし、フォルダとして
'   実在するときだけ値を返す(実在しない値を最優先の候補に置くと、その先の
'   候補まで到達しない)。読めなければ空文字=「ポインタ無し」。
' 符号化: ランチャーの `echo` が書くので端末の既定コードページ(CP932)である。
'   `Open For Input` の既定と一致するため、ここでは UTF-8 の復号を通さない。
' ============================================================================

Public Sub SetBookDir(ByVal dirText As String)
    mBookDir = TrimTrailingSep(dirText)
End Sub

Public Function BookDirHint() As String
    BookDirHint = mBookDir
End Function

' bookDir\data_dir.txt の1行目(実在するフォルダのときだけ返す)。
Public Function ReadDataDirPointer(ByVal bookDir As String) As String
    Dim fileNo As Long
    On Error GoTo Failed
    If LenB(bookDir) = 0 Then Exit Function

    Dim pathText As String
    pathText = TrimTrailingSep(bookDir) & PathSep() & DATA_DIR_POINTER_FILE
    If Not FileExistsAt(pathText) Then Exit Function

    Dim lineText As String
    fileNo = FreeFile
    Open pathText For Input As #fileNo
    If Not EOF(fileNo) Then Line Input #fileNo, lineText
    Close #fileNo

    lineText = TrimTrailingSep(Trim$(Replace(Replace(lineText, vbCr, " "), vbLf, " ")))
    If LenB(lineText) = 0 Then Exit Function
    If Not FolderExists(lineText) Then Exit Function
    ReadDataDirPointer = lineText
    Exit Function
Failed:
    CloseQuiet fileNo
    ReadDataDirPointer = vbNullString
End Function

' 親フォルダ(純関数)。区切り sep で最後の1段を落とす。段が1つしか無いとき
'   (= "C:" や "d" だけ)は空文字を返す(親を騙って同じ場所を返さない)。
Public Function ParentDirOf(ByVal pathText As String, ByVal sep As String) As String
    Dim t As String
    t = TrimTrailingSep(pathText)
    If LenB(t) = 0 Then Exit Function
    If LenB(sep) = 0 Then Exit Function

    Dim pos As Long
    pos = InStrRev(t, sep)
    If pos <= 1 Then Exit Function
    ParentDirOf = TrimTrailingSep(Left$(t, pos - 1))
End Function

' ============================================================================
' FileCandidatesIn - ナレッジブックの探索順(純関数。裁定書28)
' ----------------------------------------------------------------------------
'   (1) 本体と同じフォルダ(=ランチャーが写した D: 側)
'   (2) data_dir の親(=OneDrive の配布フォルダ。本体だけ D: に写っていて
'       ナレッジブックが写せていない端末の救済)
'   URL形式("://" を含む)の基点は区切りを "/" にする(OneDrive同期フォルダ)。
'   空の基点・同じ場所の重複は並べない。返り値は vbLf 区切り。
' ============================================================================
Public Function FileCandidatesIn(ByVal bookDir As String, ByVal dataDir As String, _
                             ByVal sepText As String, ByVal fileName As String) As String
    Dim outText As String
    If LenB(fileName) = 0 Then Exit Function

    outText = FcAppend(outText, bookDir, sepText, fileName)
    outText = FcAppend(outText, ParentDirOf(dataDir, FcSepOf(dataDir, sepText)), _
                       sepText, fileName)
    FileCandidatesIn = outText
End Function

' 基点1つぶんの連結(空・重複は足さない)。
Private Function FcAppend(ByVal listText As String, ByVal baseDir As String, _
                          ByVal sepText As String, ByVal fileName As String) As String
    FcAppend = listText
    Dim b As String
    b = TrimTrailingSep(baseDir)
    If LenB(b) = 0 Then Exit Function

    Dim candidate As String
    candidate = b & FcSepOf(b, sepText) & fileName

    Dim probe As String
    probe = vbLf & listText & vbLf
    If InStr(1, probe, vbLf & candidate & vbLf, vbTextCompare) > 0 Then Exit Function
    If LenB(listText) = 0 Then
        FcAppend = candidate
    Else
        FcAppend = listText & vbLf & candidate
    End If
End Function

' URL形式の基点は "/"、それ以外は渡された区切り(既定は "\")。
Private Function FcSepOf(ByVal baseDir As String, ByVal sepText As String) As String
    If InStr(1, baseDir, "://", vbBinaryCompare) > 0 Then
        FcSepOf = "/"
    ElseIf LenB(sepText) = 0 Then
        FcSepOf = "\"
    Else
        FcSepOf = sepText
    End If
End Function
