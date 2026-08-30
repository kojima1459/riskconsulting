Attribute VB_Name = "modCaseRead"
Option Explicit

' ============================================================================
' modCaseRead - 案件一覧の【読取専用】API(app層・裁定書7 B-7)
' ----------------------------------------------------------------------------
' 正: 14章§6(ReadCaseCtx の宣言) / 13章§2.1(案件一覧の列) / 12章§2(層と責務) /
'   19章§3(enum)。
' 責務: 案件一覧1行から「Stepを実行するのに要る文脈」を1回の読取で取り出す。
'   書込は一切しない(状態遷移・採番・case_data は modCaseStore が唯一の口)。
'   modPipeline / modCompanyFile は R4 でシートに触れないため、案件一覧の値は
'   本モジュール経由でしか手に入らない(裁定書7 B-7 の「死に経路の解消」)。
' R4(12章§4): Excelトークン許可モジュールに追加。触るのは【案件一覧】1枚だけで、
'   Range への書込を一切持たない(NFR-S7 の書込口は modUtilText.SetCellSafe のまま)。
' 列アクセスは列名ベース(13章冒頭。列番号のハードコード禁止=FindHeaderCol)。
' 例外を投げない(14章§6)。読めなければ False を返し、呼び出し側は fail-closed で
'   中止する(既定値で走らせるとAI利用枠を誤った前提で消費するため)。
' ============================================================================

Private Const CR_SRC As String = "modCaseRead"
Private Const CR_SHEET_CASES As String = "案件一覧"

' xlUp の数値(組込定数名を書かず LO の構文チェックで未定義名にしない)。
Private Const CR_DIR_UP As Long = -4162

' 見出し行を読む幅(案件一覧23列に余裕を見た探索範囲。列番号ではない)。
Private Const CR_SCAN_COLS As Long = 32

' 13章§2.1 の既定値(列が空のときに補う値。19章§3のenum)。
Private Const CR_TIER_DEFAULT As String = "t1_quick"
Private Const CR_VARIANT_DEFAULT As String = "proposal"

' ReadCaseCtx - 案件一覧1行を読む唯一の口(14章§6)。
'   ctx        : 13章§2.1 の10列(TCaseCtx。14章§6の定義)
'   roundNo    : 現ラウンド番号(空・不正は1)
'   qualityMode: config `quality_mode`(13章§2.3。空=ティア連動は
'                modPipeline.ResolveQualityMode が解決する)。案件一覧に列は無い
'   s4Variant  : 案件一覧 s4_variant(空は proposal)
'   dossierTier: 案件一覧 dossier_tier(空は t1_quick)。ctx.dossier_tier と同値
'   戻り値     : True=1行を読めた / False=読めない(シート・見出し・行が無い)
Public Function ReadCaseCtx(ByVal caseId As String, ByRef ctx As TCaseCtx, _
                            ByRef roundNo As Long, ByRef qualityMode As String, _
                            ByRef s4Variant As String, ByRef dossierTier As String) As Boolean
    On Error GoTo Failed

    roundNo = 0
    qualityMode = vbNullString
    s4Variant = vbNullString
    dossierTier = vbNullString

    If Not modCaseStore.IsValidCaseId(caseId) Then
        modLog.LogError "E0101", CR_SRC & ".ReadCaseCtx", "invalid_case_id"
        Exit Function
    End If

    Dim ws As Object
    Set ws = SheetOf(CR_SHEET_CASES)
    If ws Is Nothing Then
        modLog.LogError "E0603", CR_SRC & ".ReadCaseCtx", "sheet_missing"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = LastRowOf(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = ReadBlock(ws, lastRow)

    Dim cCase As Long
    cCase = modUtil.FindHeaderCol(blk, "case_id")
    If cCase <= 0 Then
        modLog.LogError "E0603", CR_SRC & ".ReadCaseCtx", "header_missing:case_id"
        Exit Function
    End If

    Dim r As Long
    r = RowOfCase(blk, lastRow, cCase, Trim$(caseId))
    If r < 2 Then
        modLog.LogError "E0101", CR_SRC & ".ReadCaseCtx", "case_not_found"
        Exit Function
    End If

    ctx.case_type = ValueOf(blk, r, "case_type")
    ctx.dossier_tier = OrDefault(ValueOf(blk, r, "dossier_tier"), CR_TIER_DEFAULT)
    ctx.channel = ValueOf(blk, r, "channel")
    ctx.kanji = ValueOf(blk, r, "kanji")
    ctx.bid = ValueOf(blk, r, "bid")
    ctx.reins = ValueOf(blk, r, "reins")
    ctx.other_insurers = ValueOf(blk, r, "other_insurers")
    ctx.company = ValueOf(blk, r, "company")
    ctx.industry_code = ValueOf(blk, r, "industry_code")
    ctx.industry_name = ValueOf(blk, r, "industry_name")

    dossierTier = ctx.dossier_tier
    s4Variant = OrDefault(ValueOf(blk, r, "s4_variant"), CR_VARIANT_DEFAULT)
    roundNo = LongOf(ValueOf(blk, r, "round_no"))
    If roundNo < 1 Then roundNo = 1
    qualityMode = modConfig.GetStr("quality_mode", vbNullString)

    ReadCaseCtx = True
    Exit Function

Failed:
    modLog.LogError "E0603", CR_SRC & ".ReadCaseCtx", "read_failed", Err.Number
    ReadCaseCtx = False
End Function

' --- 内部ヘルパー(読取のみ) ---

' 列名で1セルを引く(13章冒頭。列が無ければ "")。
Private Function ValueOf(ByVal blk As Variant, ByVal r As Long, _
                         ByVal headerName As String) As String
    On Error GoTo Empty0
    Dim c As Long
    c = modUtil.FindHeaderCol(blk, headerName)
    If c <= 0 Then Exit Function
    ValueOf = Trim$(CStr(blk(r, c)))
    Exit Function
Empty0:
    ValueOf = vbNullString
End Function

' 空なら既定値(13章§2.1 の「既定」列)。
Private Function OrDefault(ByVal s As String, ByVal defaultText As String) As String
    If LenB(s) = 0 Then
        OrDefault = defaultText
    Else
        OrDefault = s
    End If
End Function

' 数値列の読み(数値化できないものは0)。
Private Function LongOf(ByVal s As String) As Long
    On Error GoTo Zero0
    If LenB(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    LongOf = CLng(Val(s))
    Exit Function
Zero0:
    LongOf = 0
End Function

' 名前でシートを取る。無ければ Nothing。
Private Function SheetOf(ByVal sheetName As String) As Object
    On Error GoTo NoSheet
    Set SheetOf = ThisWorkbook.Worksheets(sheetName)
    Exit Function
NoSheet:
    Set SheetOf = Nothing
End Function

' A列基準の最終行。データが無ければ1(見出し行)。
Private Function LastRowOf(ByVal ws As Object) As Long
    Dim n As Long
    On Error GoTo One1
    n = ws.Cells(ws.Rows.count, 1).End(CR_DIR_UP).row
    If n < 1 Then n = 1
    LastRowOf = n
    Exit Function
One1:
    LastRowOf = 1
End Function

' 見出し行を含む矩形を一度だけ読む(1セルだけの Range は2次元配列にならない)。
Private Function ReadBlock(ByVal ws As Object, ByVal lastRow As Long) As Variant
    On Error GoTo Empty1
    Dim hi As Long
    hi = lastRow
    If hi < 2 Then hi = 2
    ReadBlock = ws.Range(ws.Cells(1, 1), ws.Cells(hi, CR_SCAN_COLS)).Value
    Exit Function
Empty1:
    ReadBlock = Empty
End Function

' 読み込み済みブロックから case_id 一致行を探す(見つからなければ0)。
Private Function RowOfCase(ByVal blk As Variant, ByVal lastRow As Long, _
                           ByVal caseCol As Long, ByVal caseId As String) As Long
    On Error GoTo NotFound
    Dim n As Long
    For n = 2 To lastRow
        If Trim$(CStr(blk(n, caseCol))) = caseId Then
            RowOfCase = n
            Exit Function
        End If
    Next n
    Exit Function
NotFound:
    RowOfCase = 0
End Function
