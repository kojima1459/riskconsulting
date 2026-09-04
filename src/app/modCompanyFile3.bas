Attribute VB_Name = "modCompanyFile3"
Option Explicit

' ============================================================================
' modCompanyFile3 - 企業ファイルの自動保存とスキーマ拡張(app層・T-59)
' ----------------------------------------------------------------------------
' なぜ3本目か(30,000字契約・12章§2):
'   modCompanyFile(「何を書くか」)は 29,431 字で満杯、modCompanyFile2
'   (「どこへどう書くか」)は下位I/Oだけを持つ約束である。裁定書28(W10)が
'   足した3つの責務
'     (1) 企業ファイルのスキーマ拡張(案件一覧の全列・case_data の全 data_key・
'         S1〜S4 の編集後JSON・商談の記録・判断台帳・第2ラウンド情報)
'     (2) 無言の自動保存(貼付の保管・S1〜S4 の完了・記録の追記・ラウンド確定)
'     (3) 起動時再構成のための「見出しだけ読む」口
'   はどちらにも収まらないため本モジュールを新設した。分け方は
'   modValidate/modValidate2・modCaseStore2/3 と同じ「責務の境目で割る」。
'
' 13章§2.1 の位置づけ(裁定書28):
'   **蓄積の正は企業ファイル**であり、本体xlsm の案件一覧・case_data は
'   「作業用キャッシュ」である(本体は毎朝消えてよい)。本モジュールの
'   ExportExtensions が、その「正」を1社1ファイルへ書き切る役を担う。
'
' R4(12章§4): 触るのは【企業ファイル側のシート】だけである。本体ブックの
'   値は必ず modCaseRead / modCaseStore 経由で読む(シートを直接読まない)。
'   セル書込は全て modCompanyFile2.SheetPutText/SheetPutNum(=SetCellSafe)。
'
' 16章 E-05(7)との関係: 自動保存は**無言**でなければならないので、PII検知が
'   あるときは「確認した」を求められない。よって AutoSaveCase は書き出さずに
'   "skipped" を返す(利用者は手動の[企業ファイルへ保存]で確認して書き出す)。
'   検知を握り潰して書き出す経路は作らない(fail-closed)。
' ============================================================================

Private Const CF3_SRC As String = "modCompanyFile3"

' 企業ファイルのスキーマ版(13章§2.8 dossier_meta.schema_version)。app_version
' とは別に持つ。**列を足したら major を上げず minor を上げる**(前方互換=古い
' ファイルは列が欠けているだけで読める)。
Private Const CF3_SCHEMA As String = "3.0.0"
' 版が書かれていない企業ファイル(W10 より前に作られたもの)の扱い。
Private Const CF3_SCHEMA_LEGACY As String = "1.0.0"

' 裁定書28 で足した3シート(見出しの定義は modCompanyFile2 が持つ)。
Private Const CF3_SHEET_CASE As String = "dossier_case"
Private Const CF3_SHEET_DATA As String = "dossier_data"
Private Const CF3_SHEET_JUDGE As String = "dossier_judge"
Private Const CF3_SHEET_META As String = "dossier_meta"
Private Const CF3_SHEET_FACTS As String = "dossier_facts"

' 13章§2.1 案件一覧の全列(物理順。build/sheets_main.json と同じ並び)。
Private Const CF3_CASE_COLS As String = _
    "case_id;case_type;dossier_tier;parent_case_id;company;industry_code;" & _
    "industry_name;channel;kanji;bid;reins;other_insurers;status;created_at;" & _
    "updated_at;owner;adopted_story_nos;focus_line_ids;ppt_path;report_path;" & _
    "note;s4_variant;round_no;last_ok_step;failed_step"

' 起動時再構成が案件一覧へ組み直す列(13章§2.1)。見出しから読める分だけ。
Private Const CF3_REBUILD_COLS As String = _
    "case_id;company;industry_code;case_type;status;updated_at"

' 企業フォルダ(data_dir\企業)。1社1ファイルの置き場(裁定書28)。
Private Const CF3_DIR_TAIL As String = "企業"

' ファイル名(13章§2.8)。会社名は SanitizeFileName で無害化し、8桁は company 由来。
Private Const CF3_TAIL As String = "_企業カルテ"
Private Const CF3_EXT As String = ".xlsx"

Private Const CF3_SEP As String = ";"
Private Const CF3_KV As String = "="
Private Const CF3_CHUNK_CHARS As Long = 32000

' AutoSaveCase の戻り値(ui層が表示を選ぶための3値。文言は ui が持つ)。
Public Const CF3_SAVED As String = "saved"
Public Const CF3_SKIPPED As String = "skipped"
Public Const CF3_FAILED As String = "failed"

' ============================================================================
' 純関数(層(a)テスト対象。Excelもシートも触らない)
' ============================================================================

' SchemaVersionCurrent - いま書き出す企業ファイルのスキーマ版。
Public Function SchemaVersionCurrent() As String
    SchemaVersionCurrent = CF3_SCHEMA
End Function

' SchemaVersionOf - 企業ファイルの schema_version 列の値を正規化する。
'   空(=W10 より前のファイル)は 1.0.0 とみなす。前後の空白は落とす。
Public Function SchemaVersionOf(ByVal rawText As String) As String
    Dim t As String
    t = Trim$(rawText)
    If LenB(t) = 0 Then
        SchemaVersionOf = CF3_SCHEMA_LEGACY
    Else
        SchemaVersionOf = t
    End If
End Function

' IsSchemaReadable - その版の企業ファイルを読んでよいか(前方互換の判定)。
'   **major が現行以下なら読む**(列が欠けていれば空として読む=13章§2.8)。
'   major が現行より新しいファイルは、こちらが知らない規約で書かれている
'   可能性があるため読まない(壊した内容で本体を上書きしない=fail-closed)。
Public Function IsSchemaReadable(ByVal verText As String) As Boolean
    IsSchemaReadable = (MajorOf(SchemaVersionOf(verText)) <= MajorOf(CF3_SCHEMA))
End Function

' CompanyFileNameOf - 企業ファイルのファイル名(13章§2.8 の命名テンプレート)。
'   `<Sanitize(company)>_<8桁>_企業カルテ.xlsx`。8桁は company 由来
'   (`Left$(Fnv1a64Hex(NormalizeForHash(company)), 8)` = dossier_meta.company_id
'   と同じ式)。会社名が空なら "" を返す(名前の無いファイルを作らない)。
Public Function CompanyFileNameOf(ByVal company As String) As String
    Dim baseText As String
    baseText = modUtilText.SanitizeFileName(company)
    If LenB(baseText) = 0 Then Exit Function

    CompanyFileNameOf = baseText & "_" & _
        Left$(modUtilText.Fnv1a64Hex(modUtilText.NormalizeForHash(company)), 8) & _
        CF3_TAIL & CF3_EXT
End Function

' CompanyDirOf - data_dir から企業フォルダのパスを組む(裁定書28 の構成)。
'   末尾の区切りは落としてから足す(重ねない)。空なら "" を返す。
Public Function CompanyDirOf(ByVal dataDir As String) As String
    Dim t As String
    t = modUtil.TrimTrailingSep(dataDir)
    If LenB(t) = 0 Then Exit Function
    CompanyDirOf = t & "\" & CF3_DIR_TAIL
End Function

' HeaderToCaseRow - 企業ファイルの見出し(「key=value」の行並び)から案件一覧の
'   1行を組む(起動時再構成の写像。13章§2.1)。
'   ・戻り値は「列名<TAB>値」を vbLf でつないだもの(列の並びは CF3_REBUILD_COLS)
'   ・**欠けている key は空値の行として出す**(古い企業ファイルでも落ちない)
'   ・case_id が読めなければ "" を返す(主キーの無い行を案件一覧へ入れない)
Public Function HeaderToCaseRow(ByVal headerText As String) As String
    If LenB(HeaderValueOf(headerText, "case_id")) = 0 Then Exit Function

    Dim cols() As String
    cols = modUtil.SplitKeepNonEmpty(CF3_REBUILD_COLS, CF3_SEP)

    Dim acc As String
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        If LenB(acc) > 0 Then acc = acc & vbLf
        acc = acc & cols(i) & vbTab & HeaderValueOf(headerText, cols(i))
    Next i
    HeaderToCaseRow = acc
End Function

' HeaderValueOf - 「key=value」の行並びから1つの値を引く(無ければ "")。
'   値の側に "=" が含まれていてもよい(最初の "=" だけで割る)。
Public Function HeaderValueOf(ByVal headerText As String, ByVal keyName As String) As String
    Dim lines() As String
    lines = Split(headerText, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim p As Long
        p = InStr(1, lines(i), CF3_KV, vbBinaryCompare)
        If p > 1 Then
            If StrComp(Trim$(Left$(lines(i), p - 1)), keyName, vbBinaryCompare) = 0 Then
                HeaderValueOf = Trim$(Mid$(lines(i), p + 1))
                Exit Function
            End If
        End If
    Next i
End Function

' PickNewerHeader - 同じ case_id の見出しが2つあるとき、新しいほうを選ぶ
'   (裁定書28「同じ case_id は企業ファイルを正とし新しい方で上書きする」)。
'   比較は updated_at の**文字列**で行う(13章§2.1 の時刻は
'   `yyyy-mm-dd hh:mm:ss` 固定幅なので辞書順=時刻順)。片方が空なら他方を返す。
Public Function PickNewerHeader(ByVal aText As String, ByVal bText As String) As String
    If LenB(aText) = 0 Then
        PickNewerHeader = bText
        Exit Function
    End If
    If LenB(bText) = 0 Then
        PickNewerHeader = aText
        Exit Function
    End If

    If HeaderValueOf(bText, "updated_at") > HeaderValueOf(aText, "updated_at") Then
        PickNewerHeader = bText
    Else
        PickNewerHeader = aText
    End If
End Function

' ============================================================================
' IsFileNewer - その企業ファイルの見出しが、本体の案件一覧より新しいか
'   (起動時再構成の「新しい方で上書き」の判定。裁定書28)。
'   本体に行が無ければ True(=企業ファイルから復元する)。
'   同時刻は False(本体を残す)。判定そのものは純関数 PickNewerHeader が持つ。
' ============================================================================
Public Function IsFileNewer(ByVal headerText As String, ByVal caseId As String) As Boolean
    Dim mineText As String
    mineText = Trim$(modCaseRead.CaseColumnOf(caseId, "updated_at"))
    If LenB(mineText) = 0 Then
        IsFileNewer = True
        Exit Function
    End If
    IsFileNewer = (PickNewerHeader(KvLine("updated_at", mineText), headerText) = headerText)
End Function

' ============================================================================
' 企業フォルダ(data_dir\企業)。実在確認とフォルダ作成つき。
'   data_dir の解決そのものは modUtil.ResolveDataDir が唯一持つ(呼ぶだけ)。
' ============================================================================
Public Function CompanyDir() As String
    On Error GoTo Failed

    Dim dirText As String
    dirText = CompanyDirOf(modUtil.ResolveDataDir(modConfig.GetStr("data_dir", vbNullString)))
    If LenB(dirText) = 0 Then Exit Function
    If Not modUtil.EnsureFolder(dirText) Then Exit Function

    CompanyDir = dirText
    Exit Function

Failed:
    CompanyDir = vbNullString
End Function

' ============================================================================
' AutoSaveCase - 選択中の案件を企業ファイルへ**無言で**保存する(裁定書28)。
' ----------------------------------------------------------------------------
'   戻り値: CF3_SAVED   書き出した
'           CF3_SKIPPED 書き出さなかった(案件未選択・PII未確認・保存先不明)
'           CF3_FAILED  書き出そうとして失敗した(ui層が toast と err_log を出す)
'
'   **PII検知があるときは書き出さない**(16章 E-05(7))。自動保存は無言なので
'   「確認した」を選ばせられず、確認なしの書き出しは E-05(7) 違反になるため。
'   利用者へは何も出さない(手動の[企業ファイルへ保存]で確認して書き出す)。
' ============================================================================
Public Function AutoSaveCase(ByVal caseId As String) As String
    On Error GoTo Failed

    AutoSaveCase = CF3_SKIPPED
    If Not modCaseStore.IsValidCaseId(caseId) Then Exit Function

    Dim dirText As String
    dirText = CompanyDir()
    If LenB(dirText) = 0 Then Exit Function

    If LenB(modCompanyFile.ScanCaseForPii(caseId)) > 0 Then
        modLog.LogUsage "dossier_autosave_skipped", caseId, "pii_unconfirmed"
        Exit Function
    End If

    Dim pathText As String
    pathText = modCompanyFile.ExportCompanyFile(caseId, dirText, vbNullString)
    If LenB(pathText) = 0 Then
        AutoSaveCase = CF3_FAILED
        Exit Function
    End If

    modLog.LogUsage "dossier_autosave", caseId, "ok"
    AutoSaveCase = CF3_SAVED
    Exit Function

Failed:
    modLog.LogError "E0603", CF3_SRC & ".AutoSaveCase", "autosave_failed", Err.Number
    AutoSaveCase = CF3_FAILED
End Function

' ============================================================================
' ExportExtensions - 裁定書28 が足した4つの中身を書く(modCompanyFile.
'   ExportCompanyFile が5シートを書いたあとに1回だけ呼ぶ)。
'     dossier_case  案件一覧の全列(+ schema_version)
'     dossier_data  case_data の全 data_key(S1〜S4 の編集後JSONを含む)
'     dossier_facts 商談の記録(13章§2.5。見出しだけだった枠へ行を入れる)
'     dossier_judge 判断台帳(13章§2.7)
'   同じ case_id の行は先に消して積み直す(再保存が二重行にならない)。
' ============================================================================
Public Sub ExportExtensions(ByVal wb As Object, ByVal caseId As String, ByVal roundNo As Long)
    On Error GoTo Failed
    If wb Is Nothing Then Exit Sub

    WriteCaseRow wb, caseId
    WriteCaseData wb, caseId
    WriteRecordSheet wb, caseId, roundNo, CF3_SHEET_FACTS, "フィードバック", "case_id"
    WriteRecordSheet wb, caseId, roundNo, CF3_SHEET_JUDGE, "判断台帳", "case_ref"
    Exit Sub

Failed:
    modLog.LogError "E0603", CF3_SRC & ".ExportExtensions", "write_failed", Err.Number
End Sub

' 案件一覧の全列を1行で書く(13章§2.1 の全列 + schema_version)。
Private Sub WriteCaseRow(ByVal wb As Object, ByVal caseId As String)
    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, CF3_SHEET_CASE)
    If ws Is Nothing Then Exit Sub

    modCompanyFile2.SheetDropCase ws, caseId

    Dim hdr As Variant
    hdr = modCompanyFile2.SheetBlock(ws, 2)
    Dim wr As Long
    wr = modCompanyFile2.SheetLastRow(ws) + 1

    Dim cols() As String
    cols = modUtil.SplitKeepNonEmpty(CF3_CASE_COLS, CF3_SEP)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        modCompanyFile2.SheetPutText ws, hdr, wr, cols(i), _
            modCaseRead.CaseColumnOf(caseId, cols(i))
    Next i
    modCompanyFile2.SheetPutText ws, hdr, wr, "schema_version", CF3_SCHEMA
End Sub

' case_data の全 data_key を縦持ちで書く(13章§2.2 と同じ 32,000字分割)。
Private Sub WriteCaseData(ByVal wb As Object, ByVal caseId As String)
    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, CF3_SHEET_DATA)
    If ws Is Nothing Then Exit Sub

    modCompanyFile2.SheetDropCase ws, caseId

    Dim hdr As Variant
    hdr = modCompanyFile2.SheetBlock(ws, 2)
    Dim stampText As String
    stampText = modUtil.NowStamp()

    Dim keys() As String
    keys = modUtil.SplitKeepNonEmpty(modCaseStore3.DataKeys(), CF3_SEP)

    Dim k As Long
    For k = LBound(keys) To UBound(keys)
        Dim bodyText As String
        bodyText = modCaseStore.LoadData(caseId, keys(k))
        If LenB(bodyText) > 0 Then
            Dim parts() As String
            parts = modUtil.SplitForCells(bodyText, CF3_CHUNK_CHARS)
            Dim i As Long
            For i = LBound(parts) To UBound(parts)
                Dim wr As Long
                wr = modCompanyFile2.SheetLastRow(ws) + 1
                modCompanyFile2.SheetPutText ws, hdr, wr, "case_id", caseId
                modCompanyFile2.SheetPutText ws, hdr, wr, "data_key", keys(k)
                modCompanyFile2.SheetPutNum ws, hdr, wr, "seq", i - LBound(parts) + 1
                modCompanyFile2.SheetPutText ws, hdr, wr, "content", parts(i)
                modCompanyFile2.SheetPutText ws, hdr, wr, "saved_at", stampText
            Next i
        End If
    Next k
End Sub

' 商談の記録・判断台帳を、本体シートの見出し名のまま企業ファイルへ写す。
'   本体の該当行の抽出は modCaseStore2 の読取プリミティブを通す(本モジュールが
'   本体ブックのシートを直接読まない=12章§2の責務分割)。
Private Sub WriteRecordSheet(ByVal wb As Object, ByVal caseId As String, _
                             ByVal roundNo As Long, ByVal sheetTitle As String, _
                             ByVal srcTitle As String, ByVal keyCol As String)
    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, sheetTitle)
    If ws Is Nothing Then Exit Sub

    modCompanyFile2.SheetDropCase ws, caseId

    Dim rowsText As String
    rowsText = modCaseStore3.RecordRowsOf(srcTitle, keyCol, caseId)
    If LenB(rowsText) = 0 Then Exit Sub

    Dim hdr As Variant
    hdr = modCompanyFile2.SheetBlock(ws, 2)

    Dim lines() As String
    lines = Split(rowsText, vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If LenB(lines(i)) > 0 Then
            Dim wr As Long
            wr = modCompanyFile2.SheetLastRow(ws) + 1
            modCompanyFile2.SheetPutText ws, hdr, wr, "case_id", caseId
            modCompanyFile2.SheetPutNum ws, hdr, wr, "round_no", roundNo
            PutPairs ws, hdr, wr, lines(i)
        End If
    Next i
End Sub

'  modCaseStore3.RecordRowsOf が返す1行ぶんを書き込む。区切りは同関数の規約
'  (欄=vbVerticalTab / 列名と値=vbFormFeed)。値に「;」を含む列があるため
'  「;」は使えない(理由は modCaseStore3.RecordRowsOf のコメントが持つ)。
Private Sub PutPairs(ByVal ws As Object, ByVal hdr As Variant, ByVal wr As Long, _
                     ByVal lineText As String)
    Dim pairs() As String
    pairs = Split(lineText, vbVerticalTab)
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        Dim p As Long
        p = InStr(1, pairs(i), vbFormFeed, vbBinaryCompare)
        If p > 1 Then
            modCompanyFile2.SheetPutText ws, hdr, wr, Left$(pairs(i), p - 1), _
                Mid$(pairs(i), p + 1)
        End If
    Next i
End Sub

' ============================================================================
' ImportExtensions - 企業ファイルの dossier_data を案件へ戻す(遅延読込の本体)。
'   modCompanyFile.ImportCompanyFile([企業ファイルを開く])から呼ぶ。
'   **空いている枠にだけ書く**(取込が作業中の入力を黙って上書きしない。
'   17章 T-42 観点(1)。modCompanyFile.RestoreIfEmpty と同じ規律)。
'   戻り値 = 書き戻した data_key の本数。
' ============================================================================
Public Function ImportExtensions(ByVal wb As Object, ByVal caseId As String) As Long
    On Error GoTo Failed
    If wb Is Nothing Then Exit Function

    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, CF3_SHEET_DATA)
    If ws Is Nothing Then Exit Function

    Dim keys() As String
    keys = modUtil.SplitKeepNonEmpty(modCaseStore3.DataKeys(), CF3_SEP)

    Dim n As Long
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim bodyText As String
        bodyText = modCompanyFile2.SheetJoinByKey(ws, caseId, keys(i))
        If LenB(bodyText) > 0 Then
            If LenB(modCaseStore.LoadData(caseId, keys(i))) = 0 Then
                If modCaseStore.SaveData(caseId, keys(i), bodyText) Then n = n + 1
            End If
        End If
    Next i
    ImportExtensions = n
    Exit Function

Failed:
    modLog.LogError "E0603", CF3_SRC & ".ImportExtensions", "restore_failed", Err.Number
    ImportExtensions = 0
End Function

' ============================================================================
' ReadFileHeader - 起動時再構成のために「見出しだけ」読む(裁定書28)。
'   dossier_meta(会社名・業種・最終更新・schema_version)と dossier_case
'   (案件種別・ステータス・case_id)から「key=value」の行並びを作る。
'   case_data 等の本体は読まない(遅延読込=案件を選んだときに[開く]経路で読む)。
'   読めない・版が新しすぎるファイルは "" を返す。
' ============================================================================
Public Function ReadFileHeader(ByVal filePath As String) As String
    On Error GoTo Failed

    Dim wb As Object
    Set wb = modCompanyFile2.DossierOpen(filePath, True)
    If wb Is Nothing Then Exit Function

    Dim acc As String
    acc = KvLine("company", CellOfFirstRow(wb, CF3_SHEET_META, "company"))
    acc = acc & KvLine("industry_code", CellOfFirstRow(wb, CF3_SHEET_META, "industry_code"))
    acc = acc & KvLine("updated_at", CellOfFirstRow(wb, CF3_SHEET_META, "updated_at"))

    Dim verText As String
    verText = SchemaVersionOf(CellOfFirstRow(wb, CF3_SHEET_META, "schema_version"))
    acc = acc & KvLine("schema_version", verText)

    acc = acc & KvLine("case_id", CellOfFirstRow(wb, CF3_SHEET_CASE, "case_id"))
    acc = acc & KvLine("case_type", CellOfFirstRow(wb, CF3_SHEET_CASE, "case_type"))
    acc = acc & KvLine("status", CellOfFirstRow(wb, CF3_SHEET_CASE, "status"))

    modCompanyFile2.DossierClose wb
    Set wb = Nothing

    If Not IsSchemaReadable(verText) Then Exit Function
    ReadFileHeader = acc
    Exit Function

Failed:
    modLog.LogError "E0603", CF3_SRC & ".ReadFileHeader", "header_read_failed", Err.Number
    ReadFileHeader = vbNullString
End Function

' 企業ファイルの1枚目のデータ行(2行目)の1セルを列名で読む。
Private Function CellOfFirstRow(ByVal wb As Object, ByVal sheetTitle As String, _
                                ByVal colName As String) As String
    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, sheetTitle)
    If ws Is Nothing Then Exit Function

    Dim blk As Variant
    blk = modCompanyFile2.SheetBlock(ws, modCompanyFile2.SheetLastRow(ws))
    CellOfFirstRow = Trim$(modCompanyFile2.BlockText(blk, 2, modUtil.FindHeaderCol(blk, colName)))
End Function

Private Function KvLine(ByVal keyName As String, ByVal valueText As String) As String
    KvLine = keyName & CF3_KV & valueText & vbLf
End Function

' ============================================================================
' 往復一致の照合(層(b)の企業ファイル往復テストが使う口。17章§4-1)
' ----------------------------------------------------------------------------
' CaseFingerprint  : 本体(案件一覧 + case_data)側の全項目を1本の文字列にする
' FileFingerprint  : 企業ファイル(dossier_case + dossier_data)側を同じ形にする
' 同じ並べ方で作るので、**列を1つ落とすと必ず食い違う**(変異注入の検出点)。
' ============================================================================
Public Function CaseFingerprint(ByVal caseId As String) As String
    Dim acc As String
    Dim cols() As String
    cols = modUtil.SplitKeepNonEmpty(CF3_CASE_COLS, CF3_SEP)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        acc = acc & cols(i) & CF3_KV & modCaseRead.CaseColumnOf(caseId, cols(i)) & vbLf
    Next i

    Dim keys() As String
    keys = modUtil.SplitKeepNonEmpty(modCaseStore3.DataKeys(), CF3_SEP)
    For i = LBound(keys) To UBound(keys)
        acc = acc & keys(i) & CF3_KV & modCaseStore.LoadData(caseId, keys(i)) & vbLf
    Next i
    CaseFingerprint = acc
End Function

Public Function FileFingerprint(ByVal filePath As String, ByVal caseId As String) As String
    On Error GoTo Failed

    Dim wb As Object
    Set wb = modCompanyFile2.DossierOpen(filePath, True)
    If wb Is Nothing Then Exit Function

    Dim wsCase As Object
    Set wsCase = modCompanyFile2.DossierSheet(wb, CF3_SHEET_CASE)
    Dim blk As Variant
    blk = modCompanyFile2.SheetBlock(wsCase, modCompanyFile2.SheetLastRow(wsCase))

    Dim acc As String
    Dim cols() As String
    cols = modUtil.SplitKeepNonEmpty(CF3_CASE_COLS, CF3_SEP)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        acc = acc & cols(i) & CF3_KV & _
              Trim$(modCompanyFile2.BlockText(blk, 2, modUtil.FindHeaderCol(blk, cols(i)))) & vbLf
    Next i

    Dim wsData As Object
    Set wsData = modCompanyFile2.DossierSheet(wb, CF3_SHEET_DATA)
    Dim keys() As String
    keys = modUtil.SplitKeepNonEmpty(modCaseStore3.DataKeys(), CF3_SEP)
    For i = LBound(keys) To UBound(keys)
        acc = acc & keys(i) & CF3_KV & _
              modCompanyFile2.SheetJoinByKey(wsData, caseId, keys(i)) & vbLf
    Next i

    modCompanyFile2.DossierClose wb
    Set wb = Nothing
    FileFingerprint = acc
    Exit Function

Failed:
    FileFingerprint = vbNullString
End Function

' ----------------------------------------------------------------------------
' 小道具
' ----------------------------------------------------------------------------

' 版文字列の major(先頭の数字)。読めなければ 0。
Private Function MajorOf(ByVal verText As String) As Long
    On Error GoTo Zero0
    Dim p As Long
    p = InStr(1, verText, ".", vbBinaryCompare)
    Dim headText As String
    If p > 1 Then
        headText = Left$(verText, p - 1)
    Else
        headText = verText
    End If
    If Not IsNumeric(headText) Then Exit Function
    MajorOf = CLng(headText)
    Exit Function
Zero0:
    MajorOf = 0
End Function
