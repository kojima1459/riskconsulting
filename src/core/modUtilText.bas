Attribute VB_Name = "modUtilText"
Option Explicit

' ============================================================================
' modUtilText - 外部由来テキストの無害化と、文字列の正規化・ハッシュ
' ----------------------------------------------------------------------------
' 役割(16章 NFR-S7 の中心):
'   利用者の貼付テキスト・LLM出力・受信箱の本文・取込ファイル名といった
'   【外部由来テキスト】が製品の外へ出る口は3つあり、そのすべてを本モジュール
'   の関数に強制する。
'     (1) セル   -> SetCellSafe            (先頭式記号の無害化・NUL除去・切詰め)
'     (2) CSV    -> SanitizeForCell と同一ガードを通す
'     (3) HTML/JS-> HtmlSafe / JsStringSafe
'   防御関数が1本あっても素通り経路が残っていた実例(姉妹PJ B6BE7監査)への
'   対処として、tools/vba_lint.py が「本モジュール以外でのセル直接書込」と
'   「HTML連結の素通し」をERRORで落とす。ここが唯一の口である。
'
' 移植元: PoC「マイ本棚AI」 src/core/modUtilText.bas / modUtil.bas。
'   移植したのは SanitizeForCell / IsoDate系 / Fnv1a64Hex / NormalizeForHash。
'   PoC固有の機能(ADODBによるファイル読込・文字コード自動判定・進捗文言
'   組立・作業用Excel起動コマンド組立・掲示板行の直列化)は RPN に対応する
'   責務が無いため非移植。SetCellSafe / SanitizeInput / SanitizeFileName /
'   HtmlSafe / JsStringSafe は 16章・18章の要求に基づく新設。
'
' R4準拠: Excelトークン(Worksheets/Range/Application/ThisWorkbook/MsgBox/
'   ActiveSheet)は使わない。SetCellSafe の書込先は【後期バインドの Object】
'   として受け取り、`.Value` にだけ触れる。これで core層のまま
'   「セル書込の唯一の口」を成立させられる。
' ============================================================================

' 1セルへ書ける最大字数(16章 E-22。物理上限32,767字の手前で止める)。
Private Const CELL_MAX_CHARS As Long = 32000

' 15章のデータ境界記号。貼付テキスト本文がこれを騙るのを防ぐ(16章 E-04)。
Private Const BOUNDARY_MARK As String = "■■■"
Private Const BOUNDARY_ALT As String = "[境界記号]"

' Windowsのファイル名禁止文字(13章§2.8 手順1)。
Private Const FILENAME_BAD As String = "\/:*?""<>|"

' ファイル名の切詰め長(13章§2.8 手順3)と、最終パス長の上限(手順5)。
Private Const FILENAME_MAX_CHARS As Long = 32
Private Const FULLPATH_MAX_CHARS As Long = 240

' FNV-1a 64bit の定数(offset_basis と prime = 2^40 + 435)。
' LibreOffice Basic は「宣言より前の行での参照」を解決できないため、
' 使用箇所より前=モジュール先頭に置く(PoCがLO実行テストで踏んだ罠)。
Private Const FNV_OFFSET_HI As Long = &HCBF29CE4&
Private Const FNV_OFFSET_LO As Long = &H84222325&
Private Const FNV_PRIME_HI As Long = &H100
Private Const FNV_PRIME_LO As Long = &H1B3

' SetCellSafe の再入ガード。ガード発火時に modLog へ記録するが、modLog は
' セルを SetCellSafe で書くため、記録処理の中の書込が再びここへ戻ってくる。
' 記録が記録を呼ぶ入れ子を1段で止める(記録の欠落より暴走を避ける)。
Private gTxtInGuardLog As Boolean

' ============================================================================
' SetCellSafe - 外部由来テキストをセルへ書き込む唯一の口(16章 NFR-S7①)。
' ----------------------------------------------------------------------------
'   target   : 書込先のセル(Range)。Excel型に依存しないよう Object で受ける。
'   rawText  : 外部由来のテキスト(貼付・LLM出力・受信箱body・ファイル名)。
'   whereNote: 発火時に err_log の detail へ書く【箇所】(例 "run_log/detail")。
'              NFR-S3により本文は記録しないので、箇所だけをここで受け取る。
'   戻り値   : 実際にセルへ書いた文字列(呼び出し側の検証・テスト用)。
'
'   本関数は失敗を握りつぶさない。書込に失敗したら例外はそのまま呼び出し側の
'   エラーハンドラへ抜ける(「書けなかったのに成功に見える」を作らない)。
' ============================================================================
Public Function SetCellSafe(ByVal target As Object, ByVal rawText As String, _
                            Optional ByVal whereNote As String = "") As String
    Dim outText As String
    outText = SanitizeForCell(rawText)

    ' 実際に切り詰めが起きたかは「NUL除去後の長さ」で判定する(NULだけで
    ' 上限を超えていた文字列を打切り扱いにしない)。
    Dim srcLen As Long
    srcLen = Len(Replace(rawText, vbNullChar, vbNullString))

    ' 発火の記録(16章 E-46 / E-22)。detail には箇所だけを書く(NFR-S3)。
    If Not gTxtInGuardLog Then
        gTxtInGuardLog = True
        If StartsWithFormulaChar(rawText) Then
            modLog.LogError "E0606", "SetCellSafe", "formula_guard:" & whereNote
        End If
        If srcLen > CELL_MAX_CHARS Then
            modLog.LogError "E0604", "SetCellSafe", _
                "cell_truncated:" & whereNote & " len=" & CStr(srcLen)
        End If
        gTxtInGuardLog = False
    End If

    target.Value = outText
    SetCellSafe = outText
End Function

' ============================================================================
' SanitizeForCell - セル/CSVへ書く直前の無害化(純関数。SetCellSafeの中身)。
' ----------------------------------------------------------------------------
'   16章 E-46 / NFR-S7①の3点を、この順で行う。
'     (1) NULバイト(Chr(0))を除去する
'         セルの値にNULが混ざるとExcelは以降の文字を黙って捨てることがあり、
'         「保存はできたのに中身が切れている」という気付けない欠損になる。
'     (2) 32,000字で切り詰める(サロゲートペアを割らない modUtil.SafeLeft)
'     (3) 先頭が = + - @ ならアポストロフィを前置してテキストへ固定する
'         セルは実行環境であり、外部由来テキストの先頭記号は数式として
'         評価される(DDE・外部参照へ繋がる)。
'   (3)を最後に置くのは、切詰めで先頭が変わることはない一方、前置してから
'   切ると「'」ぶんだけ本文が1字失われるため。前置後の長さは最大32,001字で、
'   セルの物理上限32,767字を超えない。
' ============================================================================
Public Function SanitizeForCell(ByVal s As String) As String
    If LenB(s) = 0 Then
        SanitizeForCell = s
        Exit Function
    End If

    Dim t As String
    t = Replace(s, vbNullChar, vbNullString)
    t = modUtil.SafeLeft(t, CELL_MAX_CHARS)

    If StartsWithFormulaChar(t) Then t = "'" & t
    SanitizeForCell = t
End Function

' 先頭が数式記号(= + - @)か。空文字はFalse。
Private Function StartsWithFormulaChar(ByVal s As String) As Boolean
    If LenB(s) = 0 Then Exit Function
    Dim c As String
    c = Left$(s, 1)
    StartsWithFormulaChar = (c = "=" Or c = "+" Or c = "-" Or c = "@")
End Function

' ============================================================================
' SanitizeInput - 外部由来テキストの無害化(16章 E-04)。
' ----------------------------------------------------------------------------
'   (1) 制御文字を除去する。ただしタブとLFは本文の構造なので残す。改行は
'       先にLFへ均す(CRLF/CRをLFへ。表現を1つにしないと字数とハッシュが揺れる)。
'   (2) 私用領域(U+E000-U+F8FF)の文字を除去する。フォント依存の外字であり、
'       LLMへ送ってもレポートへ載せても意味を持たないうえ、CP932往復で
'       別の文字へ化ける。
'   (3) 本文中の "■■■" を "[境界記号]" へ置換する。15章のデータ境界記号を
'       貼付テキストが騙るのを防ぐ多層防御の1枚(E-43)。
'
'   removedCount / markerCount は【件数だけ】を返す省略可能な出口。
'   16章 E-04 が「除去・置換の件数はrun_logのdetailに件数のみ記録し本文は
'   残さない」を求めるが、14章§6の宣言は引数を持たないため、後方互換な
'   Optional ByRef として足した(引数なしの呼び出しは宣言どおり動く)。
' ============================================================================
Public Function SanitizeInput(ByVal s As String, _
                              Optional ByRef removedCount As Long = 0, _
                              Optional ByRef markerCount As Long = 0) As String
    removedCount = 0
    markerCount = 0
    If LenB(s) = 0 Then
        SanitizeInput = s
        Exit Function
    End If

    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)

    Dim n As Long
    n = Len(t)
    Dim buf() As String
    ReDim buf(1 To n)
    Dim kept As Long
    kept = 0

    Dim i As Long
    For i = 1 To n
        Dim ch As String
        ch = Mid$(t, i, 1)
        Dim code As Long
        code = AscW(ch)
        If code < 0 Then code = code + 65536

        Dim drop As Boolean
        drop = False
        If code < 32 Then
            If code <> 9 And code <> 10 Then drop = True
        ElseIf code = 127 Then
            drop = True
        ElseIf code >= &HE000& And code <= &HF8FF& Then
            drop = True
        End If

        If drop Then
            removedCount = removedCount + 1
        Else
            kept = kept + 1
            buf(kept) = ch
        End If
    Next i

    Dim cleaned As String
    If kept > 0 Then
        ReDim Preserve buf(1 To kept)
        cleaned = Join(buf, vbNullString)
    End If

    markerCount = CountOccurrences(cleaned, BOUNDARY_MARK)
    If markerCount > 0 Then
        cleaned = Replace(cleaned, BOUNDARY_MARK, BOUNDARY_ALT)
    End If
    SanitizeInput = cleaned
End Function

Private Function CountOccurrences(ByVal hay As String, ByVal needle As String) As Long
    If LenB(hay) = 0 Or LenB(needle) = 0 Then Exit Function
    Dim n As Long
    n = 0
    Dim pos As Long
    pos = InStr(1, hay, needle, vbBinaryCompare)
    Do While pos > 0
        n = n + 1
        pos = InStr(pos + Len(needle), hay, needle, vbBinaryCompare)
    Loop
    CountOccurrences = n
End Function

' ============================================================================
' SanitizeFileName - ファイル名の安全化(13章§2.8 手順1-3)。
' ----------------------------------------------------------------------------
'   1. 禁止文字 \ / : * ? " < > | と制御文字(Chr(0)-Chr(31))を "_" へ置換
'   2. 前後の空白を除去し、末尾のピリオドを除去(Windowsが黙って落とすため)
'   3. 先頭から32字で切り詰める
'
'   手順4(case_id由来8桁の付与)と手順5(最終パス240字超の退避)は、案件IDや
'   出力先ディレクトリという【呼び出し側にしかない情報】を必要とするため、
'   本関数ではなく BuildFileNameSafe が担う。企業ドシエ・report_path・
'   ppt_path の3経路は BuildFileNameSafe を呼ぶこと(手順1-3だけでは
'   同名社2件の衝突とMAX_PATH超過が防げない)。
'
'   引数名は14章§6・13章§2.8の宣言では name だが、VBAの組み込み名 Name と
'   衝突して実機でコンパイルエラーになるため rawName とした(型・順序・
'   個数は宣言どおり)。
' ============================================================================
Public Function SanitizeFileName(ByVal rawName As String) As String
    Dim t As String
    t = rawName

    Dim i As Long
    For i = 1 To Len(FILENAME_BAD)
        t = Replace(t, Mid$(FILENAME_BAD, i, 1), "_")
    Next i
    For i = 0 To 31
        t = Replace(t, ChrW(i), "_")
    Next i

    t = Trim$(t)
    Do While Len(t) > 0
        If Right$(t, 1) <> "." Then Exit Do
        t = Left$(t, Len(t) - 1)
    Loop
    t = Trim$(t)

    SanitizeFileName = modUtil.SafeLeft(t, FILENAME_MAX_CHARS)
End Function

' ============================================================================
' BuildFileNameSafe - 13章§2.8のファイル名生成規則(手順1-5)の全体。
' ----------------------------------------------------------------------------
'   company  : 案件入力の自由記述(生のままファイル名にしない対象)
'   caseId   : 案件ID。先頭8桁のfnv1a64を衝突回避キーとして付ける(手順4)
'   tailPart : 用途を表す末尾(例 "リスクドシエ" / "リスクレポート_20260901")
'   dirPath  : 出力先ディレクトリ(手順5の長さ判定に使う。判定不要なら "")
'   ext      : 拡張子("." を含む。例 ".html"。長さ判定にのみ使う)
'   戻り値   : 拡張子を除いたファイル名(呼び出し側が ext を付けて使う)
'
'   手順5では company 部を丸ごと Fnv1a64Hex の16桁へ置換して再構成する。
'   短縮ではなく置換なのは、途中で切ると別会社が同名になり得るため。
' ============================================================================
Public Function BuildFileNameSafe(ByVal company As String, ByVal caseId As String, _
                                  ByVal tailPart As String, ByVal dirPath As String, _
                                  ByVal ext As String) As String
    Dim head As String
    head = SanitizeFileName(company)
    If LenB(head) = 0 Then head = "no_name"

    Dim key8 As String
    key8 = Left$(Fnv1a64Hex(caseId), 8)

    Dim tail As String
    tail = SanitizeFileName(tailPart)

    Dim nameText As String
    nameText = JoinParts(head, key8, tail)

    ' 手順5: ディレクトリを含めた最終パスが240字を超えるなら company 部を退避
    Dim fullLen As Long
    fullLen = Len(dirPath) + 1 + Len(nameText) + Len(ext)
    If fullLen > FULLPATH_MAX_CHARS Then
        nameText = JoinParts(Fnv1a64Hex(company), key8, tail)
    End If

    BuildFileNameSafe = nameText
End Function

Private Function JoinParts(ByVal a As String, ByVal b As String, ByVal c As String) As String
    Dim t As String
    t = a
    If LenB(b) > 0 Then t = t & "_" & b
    If LenB(c) > 0 Then t = t & "_" & c
    JoinParts = t
End Function

' ============================================================================
' HtmlSafe - HTML本文・属性値へ差し込む値のエンティティ化(16章 E-47・18章§5.3)。
' ----------------------------------------------------------------------------
'   & を最初に置換する。後回しにすると、先に入れた &lt; の & がもう一度
'   置換されて &amp;lt; になる(二重エスケープ)。
' ============================================================================
Public Function HtmlSafe(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    t = Replace(t, "'", "&#39;")
    HtmlSafe = t
End Function

' ============================================================================
' JsStringSafe - JS の二重引用符文字列リテラルへ入れるための変換。
' ----------------------------------------------------------------------------
'   適用順の正は18章§5.3。この順を守らないと二重エスケープになる。
'     1. \ -> \\ ・ " -> \"
'     2. </ -> <\/        (本文中の </script> でブロックが閉じるのを防ぐ)
'     3. U+2028 -> \u2028 ・ U+2029 -> \u2029 (JSでは行終端子として扱われる)
'     4. その他の制御文字(U+0000-U+001F) -> \u00XX
'   1を最初に置くのは、後段で足す "\" が再エスケープされないようにするため。
'   4は改行・タブも含めて \u00XX にする(JSON.parse が
'   元の文字へ戻すので情報は落ちない)。
' ============================================================================
Public Function JsStringSafe(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, "</", "<\/")
    t = Replace(t, ChrW(&H2028&), "\u2028")
    t = Replace(t, ChrW(&H2029&), "\u2029")

    Dim i As Long
    For i = 0 To 31
        If InStr(1, t, ChrW(i), vbBinaryCompare) > 0 Then
            t = Replace(t, ChrW(i), "\u00" & Right$("0" & Hex$(i), 2))
        End If
    Next i
    JsStringSafe = t
End Function

' ============================================================================
' NormalizeForHash - 重複判定用の正規化(16章 E-49 の前段)。
' ----------------------------------------------------------------------------
'   改行をLFへ統一し、半角空白・タブの連続を1個へ圧縮し、前後を Trim する。
'   LF自体は圧縮しない(段落の境界情報を壊すと別内容が同一視される)。
'   尾部劣化(同名項目の重複出力)の検出は、この正規化のあとの fnv1a64 が
'   一致するかで行う。
' ============================================================================
Public Function NormalizeForHash(ByVal s As String) As String
    Dim t As String
    t = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    Dim n As Long
    n = Len(t)
    If n = 0 Then
        NormalizeForHash = vbNullString
        Exit Function
    End If

    Dim outArr() As String
    ReDim outArr(1 To n)
    Dim outCount As Long
    outCount = 0
    Dim prevWasSpace As Boolean
    prevWasSpace = False

    Dim i As Long
    For i = 1 To n
        Dim c As String
        c = Mid$(t, i, 1)
        If c = " " Or c = vbTab Then
            If Not prevWasSpace Then
                outCount = outCount + 1
                outArr(outCount) = " "
            End If
            prevWasSpace = True
        Else
            outCount = outCount + 1
            outArr(outCount) = c
            prevWasSpace = False
        End If
    Next i

    Dim joined As String
    If outCount > 0 Then
        ReDim Preserve outArr(1 To outCount)
        joined = Join(outArr, vbNullString)
    End If
    NormalizeForHash = Trim$(joined)
End Function

' ============================================================================
' Fnv1a64Hex - FNV-1a 64bit ハッシュを16進16桁(小文字)で返す。
' ----------------------------------------------------------------------------
'   offset_basis = 0xCBF29CE484222325 / prime = 0x100000001B3
'   文字列を先頭から1文字ずつUTF-16コードユニットとして取り出し、
'   下位バイト→上位バイトの順に投入する:
'     hash = hash Xor byte : hash = (hash * prime) Mod 2^64
'   VBAには符号なし64bit整数が無いので、上位32bit/下位32bitのLong2本で持ち、
'   乗算は16bitずつに割ってDoubleで行う(Doubleが誤差なく表せる整数の上限
'   2^53 を中間値が超えない構成)。32bit Excel前提のため LongLong は使わない。
' ============================================================================
Public Function Fnv1a64Hex(ByVal s As String) As String
    Dim hHi As Long, hLo As Long
    hHi = FNV_OFFSET_HI
    hLo = FNV_OFFSET_LO

    Dim n As Long
    n = Len(s)
    Dim i As Long
    For i = 1 To n
        Dim code As Long
        code = AscW(Mid$(s, i, 1))
        If code < 0 Then code = code + 65536

        Dim byteLo As Long, byteHi As Long
        byteLo = code And &HFF
        byteHi = (code \ 256) And &HFF

        hLo = hLo Xor byteLo
        MulU64ByPrime hHi, hLo
        hLo = hLo Xor byteHi
        MulU64ByPrime hHi, hLo
    Next i

    Fnv1a64Hex = U32ToHex8(hHi) & U32ToHex8(hLo)
End Function

' ============================================================================
' Iso日時 - 日付を文字列にする唯一の口。
' ----------------------------------------------------------------------------
'   Format$(d, "yyyy-mm-dd") は和暦カレンダー設定の端末で年が元号年になる。
'   ログの並び順・集計キー・突合が黙って壊れるため、日付を「文字列として
'   書く・比べる」箇所は Format$ を使わず必ずここを通す。
' ============================================================================
Public Function IsoDate(ByVal d As Date) As String
    IsoDate = Pad0(Year(d), 4) & "-" & Pad0(Month(d), 2) & "-" & Pad0(Day(d), 2)
End Function

Public Function IsoDateTime(ByVal dt As Date) As String
    IsoDateTime = IsoDate(dt) & " " & Pad0(Hour(dt), 2) & ":" & _
                  Pad0(Minute(dt), 2) & ":" & Pad0(Second(dt), 2)
End Function

Public Function IsoDateCompact(ByVal d As Date) As String
    IsoDateCompact = Pad0(Year(d), 4) & Pad0(Month(d), 2) & Pad0(Day(d), 2)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------
Private Function Pad0(ByVal n As Long, ByVal digits As Long) As String
    Pad0 = Right$(String$(digits, "0") & CStr(n), digits)
End Function

' Longのビットパターン(0-2^32-1相当)をDoubleの数値へ。
Private Function U32ToDouble(ByVal bits As Long) As Double
    If bits < 0 Then
        U32ToDouble = CDbl(bits) + 4294967296#
    Else
        U32ToDouble = CDbl(bits)
    End If
End Function

' [0, 2^32) のDouble値を、その値を表すLongのビットパターンへ。
Private Function DoubleToU32Bits(ByVal d As Double) As Long
    Dim v As Double
    v = d - Int(d / 4294967296#) * 4294967296#
    If v >= 2147483648# Then
        DoubleToU32Bits = CLng(v - 4294967296#)
    Else
        DoubleToU32Bits = CLng(v)
    End If
End Function

' 32bit x 32bit -> 64bit(hiOut:loOut) の符号なし乗算(オーバーフローなし)。
Private Sub MulU32(ByVal aBits As Long, ByVal bBits As Long, _
                   ByRef hiOut As Long, ByRef loOut As Long)
    Dim a As Double, b As Double
    a = U32ToDouble(aBits)
    b = U32ToDouble(bBits)

    Dim aHi As Double, aLo As Double, bHi As Double, bLo As Double
    aHi = Int(a / 65536#)
    aLo = a - aHi * 65536#
    bHi = Int(b / 65536#)
    bLo = b - bHi * 65536#

    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double
    t0 = aLo * bLo
    t1 = aHi * bLo
    t2 = aLo * bHi
    t3 = aHi * bHi

    Dim digit0 As Double, digit1 As Double, digit2 As Double, digit3 As Double
    Dim carry As Double, sum1 As Double, sum2 As Double

    digit0 = t0 - Int(t0 / 65536#) * 65536#
    carry = Int(t0 / 65536#)

    sum1 = carry + (t1 - Int(t1 / 65536#) * 65536#) + (t2 - Int(t2 / 65536#) * 65536#)
    digit1 = sum1 - Int(sum1 / 65536#) * 65536#
    carry = Int(sum1 / 65536#) + Int(t1 / 65536#) + Int(t2 / 65536#)

    sum2 = carry + (t3 - Int(t3 / 65536#) * 65536#)
    digit2 = sum2 - Int(sum2 / 65536#) * 65536#
    carry = Int(sum2 / 65536#) + Int(t3 / 65536#)

    digit3 = carry - Int(carry / 65536#) * 65536#   ' 64bit超過分は破棄

    loOut = DoubleToU32Bits(digit1 * 65536# + digit0)
    hiOut = DoubleToU32Bits(digit3 * 65536# + digit2)
End Sub

' hash(hHi:hLo) = hash * FNV_PRIME Mod 2^64 を破壊的に計算する。
'   prime = PHi*2^32 + PLo (PHi=256, PLo=435)であることを使うと、
'   HHi*PHi*2^64 の項は mod 2^64 で必ず0になるため計算不要。
Private Sub MulU64ByPrime(ByRef hHi As Long, ByRef hLo As Long)
    Dim hiA As Long, loA As Long
    MulU32 hLo, FNV_PRIME_LO, hiA, loA

    Dim hiB As Long, loB As Long
    MulU32 hHi, FNV_PRIME_LO, hiB, loB

    Dim hiC As Long, loC As Long
    MulU32 hLo, FNV_PRIME_HI, hiC, loC

    Dim sumHi As Double
    sumHi = U32ToDouble(hiA) + U32ToDouble(loB) + U32ToDouble(loC)

    hLo = loA
    hHi = DoubleToU32Bits(sumHi)
End Sub

Private Function U32ToHex8(ByVal bits As Long) As String
    U32ToHex8 = LCase$(Right$("00000000" & Hex$(bits), 8))
End Function
