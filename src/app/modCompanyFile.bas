Attribute VB_Name = "modCompanyFile"
Option Explicit

' ============================================================================
' modCompanyFile - 企業ドシエファイル(1社1.xlsx)の書出・取込(app層・T-29)
' ----------------------------------------------------------------------------
' 責務(10章FR-45・13章§2.8・11章§2・16章 E-05(7)):
'   ツール本体(xlsm)は「エンジン」、企業ファイルは「蓄積」。担当15社なら15
'   ファイルを営業が持ち、面談・提案のたびに追記して育てる。本モジュールが
'   その1ファイルに対する I/O の唯一の口。
'     ExportCompanyFile : 現ラウンドを追記書き出し(HOMEの[保存])
'     ImportCompanyFile : 前ラウンドの s1/s2 と notes を案件へ復元(HOMEの[開く])
'     ScanCaseForPii    : 書出前・共有前のPII走査(結果は dossier_meta へ記録)
'     CompanyFilePath   : 13章§2.8のファイル名規則で決まる保存先の絶対パス
'
' 2本に分かれている理由: 1本だと30,000字契約(12章§2)を超えたため、下位の
'   ブック・シートI/Oを modCompanyFile2 へ出した(modValidate/modValidate2 と
'   同じ分け方)。本モジュールは「何を書くか」だけを持つ。
'
' PoC modPack / modPackExport からの転用骨格(12章§2 移植対応表): (1)マクロ無し
'   .xlsx への新規書出と再オープン追記 (2)2段検証=書いて閉じたあと開き直して
'   照合 (3)書出前のPII走査。pack_meta / pack_chunks の2シートは13章§2.8の
'   5シート構成へ差し替えた。
'
' R4(12章§4): Excelトークン許可10本の1つ。触るのは【企業ファイル側のシート】
'   だけで、本体ブックのシートには一切触らない。案件データは必ず modCaseStore
'   経由(唯一の口)。セル書込は全て modUtilText.SetCellSafe(16章 NFR-S7(1))。
'
' 案件一覧の値(company / industry_code / dossier_tier / round_no)は
'   modCaseRead.ReadCaseCtx(14章§6・裁定書7 B-7)で読む。本モジュールが本体
'   ブックのシートを直接読むと「そのシートの唯一の入出力口は store」という
'   12章§2の責務分割が崩れるため、読取は必ずこの1本を通す。
'   dossier_facts(13章§2.5 フィードバック由来)は5シートの1枚として見出しつきで
'   作るが、行の投入は供給元APIの宣言待ちで保留(見出しだけ先に作る)。
'
' 16章 E-05(7)(共有前は確認必須): ExportCompanyFile は書き出す中身を modPii に
'   通し、検知があるのに confirmedAt(利用者が「確認した」を選んだ日時)が空なら
'   【書き出さない】。走査結果は検知の有無に関わらず dossier_meta の
'   pii_scan_result へ、確認日時は pii_confirmed_at へ記録する。err_log の
'   detail は modPii が返す「種別と箇所」だけで本文を含まない(NFR-S3)。
' ============================================================================

Private Const CF_SRC As String = "modCompanyFile"

' 13章§2.8の5シート。1行目ヘッダの定義と作成は modCompanyFile2 が持つ。
Private Const CF_SHEET_META As String = "dossier_meta"
Private Const CF_SHEET_PROFILE As String = "dossier_profile"
Private Const CF_SHEET_ROUNDS As String = "dossier_rounds"
Private Const CF_SHEET_NOTES As String = "dossier_notes"

' dossier_notes に積む case_data のキー(13章§2.8「現場メモ・ヒアリング回答の
' 原文累積」)。取込時に書き戻してよいキーの許可リストも兼ねる。
Private Const CF_NOTE_KEYS As String = "input_field_notes;input_hearing_answers"
Private Const CF_NOTE_TAG As String = "field_insights"

' ファイル名(13章§2.8)。1社1ファイルなので `_2` の連番は付けない。
Private Const CF_TAIL As String = "リスクドシエ"
Private Const CF_EXT As String = ".xlsx"

' 1セルに入れる最大字数(case_data と同方式。16章 E-22)。
Private Const CF_CHUNK_CHARS As Long = 32000

Private Const CF_SEP As String = ";"
Private Const CF_SCHEMA_FALLBACK As String = "2.0.0"

' ============================================================================
' CompanyFilePath - 企業ドシエファイルの絶対パス(13章§2.8のファイル名規則)。
'   company は自由記述なので生では使わず必ず BuildFileNameSafe(禁止文字置換・
'   32字切詰め・8桁ハッシュの付与・240字超のハッシュ退避)を通す。
'   裁定書9 B7(13章§2.8 手順4): 企業ドシエだけは8桁を case_id 由来ではなく
'   **company 由来**(Left$(Fnv1a64Hex(NormalizeForHash(company)), 8)=
'   dossier_meta.company_id と同じ式)にする。case_id 由来だと同じ会社の2件目の
'   案件が必ず別ファイルになり「1社1ファイル・追記して育てる」(FR-45)が
'   成立しない。BuildFileNameSafe は第2引数の Fnv1a64Hex 先頭8桁を使うため、
'   NormalizeForHash(company) を渡すことで仕様の式と一致させる。
'   caseId 引数はシグネチャ(14章§6)維持のため残す(ファイル名には使わない)。
' ============================================================================
Public Function CompanyFilePath(ByVal company As String, ByVal caseId As String, _
                                ByVal dirPath As String) As String
    Dim dirText As String
    dirText = TrimTrailingSep(dirPath)
    If LenB(dirText) = 0 Then Exit Function

    Dim baseName As String
    baseName = modUtilText.BuildFileNameSafe(company, _
                   modUtilText.NormalizeForHash(company), CF_TAIL, dirText, CF_EXT)
    If LenB(baseName) = 0 Then Exit Function

    CompanyFilePath = dirText & "\" & baseName & CF_EXT
End Function

' ============================================================================
' ScanCaseForPii - 書き出す予定の中身をまとめてPII走査する(16章 E-05(7))。
'   戻り値 "" = 検知なし。非空 = 「箇所名|種別@文字位置」の列挙(本文を含まない
'   =NFR-S3)。ui層はこれを利用者へ提示し「確認した」を選ばせてから
'   ExportCompanyFile へ確認日時を渡す。
' ============================================================================
Public Function ScanCaseForPii(ByVal caseId As String) As String
    Dim buf() As String
    Dim cnt As Long
    modUtil.BufInit buf, cnt

    Dim n As Long
    For n = 1 To 3
        AddReport buf, cnt, modCaseStore.ResolveStepJson(caseId, n), _
                  CF_SHEET_ROUNDS & "/s" & CStr(n) & "_json"
    Next n

    Dim keys() As String
    keys = modUtil.SplitKeepNonEmpty(CF_NOTE_KEYS, CF_SEP)
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        AddReport buf, cnt, modCaseStore.LoadData(caseId, keys(i)), _
                  CF_SHEET_NOTES & "/" & keys(i)
    Next i

    ScanCaseForPii = modUtil.BufText(buf, cnt)
End Function

' ============================================================================
' ExportCompanyFile - 現ラウンドを企業ファイルへ追記書き出しする(FR-45)。
' ----------------------------------------------------------------------------
'   caseId      : 案件ID。company / industry_code / dossier_tier / round_no は
'                 modCaseRead.ReadCaseCtx(14章§6・裁定書7 B-7)で読む。案件一覧の
'                 読取APIが無かったため ui から値を貰う設計にしていたのを結線した
'                 (読めなければ fail-closed。空の企業名でファイルを作らない)。
'   dirPath     : 保存先ディレクトリ
'   confirmedAt : PII検知に対し利用者が「確認した」を選んだ日時(未選択は "")
'   戻り値      : 書き出したファイルの絶対パス。失敗・中止は ""
'
'   同じ (case_id, round_no) の行は先に消して積み直す(再保存が二重行にならな
'   い)。別ラウンドの行は消さない=これが「追記して育てる」の実体。書き終えたら
'   閉じて開き直し s1-s3 JSON と notes を読み直してハッシュ照合する(2段検証)。
' ============================================================================
Public Function ExportCompanyFile(ByVal caseId As String, ByVal dirPath As String, _
                                  ByVal confirmedAt As String) As String
    On Error GoTo Failed

    If Not modCaseStore.IsValidCaseId(caseId) Then
        modLog.LogError "E0101", CF_SRC & ".ExportCompanyFile", "invalid_case_id"
        Exit Function
    End If

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String, s4Variant As String, tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, _
                                   tierText) Then
        modLog.LogUsage "dossier_export_blocked", caseId, "case_ctx_unreadable"
        Exit Function
    End If

    Dim pathText As String
    pathText = CompanyFilePath(ctx.company, caseId, dirPath)
    If LenB(pathText) = 0 Then
        modLog.LogError "E0101", CF_SRC & ".ExportCompanyFile", "empty_out_dir"
        Exit Function
    End If

    ' 16章 E-05(7): 走査 -> 検知があるのに未確認なら書き出さない。
    Dim piiText As String
    piiText = ScanCaseForPii(caseId)
    If LenB(piiText) > 0 Then
        modLog.LogError "E0103", CF_SRC & ".ExportCompanyFile", piiText
        If LenB(Trim$(confirmedAt)) = 0 Then
            modLog.LogUsage "dossier_export_blocked", caseId, "pii_unconfirmed"
            Exit Function
        End If
    End If

    Dim wb As Object
    Set wb = modCompanyFile2.DossierOpenOrCreate(pathText)
    If wb Is Nothing Then
        modLog.LogError "E0603", CF_SRC & ".ExportCompanyFile", "open_failed"
        Exit Function
    End If

    WriteMeta wb, ctx, piiText, confirmedAt, caseId
    WriteProfile wb, caseId, roundNo
    WriteRounds wb, caseId, ctx, roundNo
    WriteNotes wb, caseId, roundNo

    Dim expectHash As String
    expectHash = PayloadHash(caseId)

    ' 裁定書9 B8(14章§6): 保存の失敗を成功として返さない。False なら
    ' VerifyRoundTrip へ進まず ""(失敗)を返す(ui層が失敗を表示する)。
    If Not modCompanyFile2.DossierSaveAndClose(wb, pathText) Then
        Set wb = Nothing
        modLog.LogError "E0603", CF_SRC & ".ExportCompanyFile", _
                        "save_failed:round=" & CStr(roundNo)
        Exit Function
    End If
    Set wb = Nothing

    ' 裁定書9 B8: VerifyRoundTrip の前に対象ブックが閉じていることを確認する。
    ' 同一プロセスで開いたままのブックを Workbooks.Open が返すと、ディスクでは
    ' なくメモリ上の内容と突合して合格してしまう(2段検証の無効化)。
    If BookStillOpen(pathText) Then
        modLog.LogError "E0603", CF_SRC & ".ExportCompanyFile", "book_still_open"
        Exit Function
    End If

    If Not VerifyRoundTrip(pathText, caseId, roundNo, expectHash) Then
        modLog.LogError "E0603", CF_SRC & ".ExportCompanyFile", _
                        "verify_mismatch:round=" & CStr(roundNo)
        Exit Function
    End If

    modLog.LogUsage "dossier_export", caseId, "round=" & CStr(roundNo)
    ExportCompanyFile = pathText
    Exit Function

Failed:
    modLog.LogError "E0603", CF_SRC & ".ExportCompanyFile", "write_failed", Err.Number
    ExportCompanyFile = vbNullString
End Function

' ============================================================================
' ImportCompanyFile - 企業ファイルの最新ラウンドを案件へ復元する(HOMEの[開く])。
' ----------------------------------------------------------------------------
'   復元先(13章§2.2の data_key):
'     S1 JSON -> s1_json      (前回までの企業理解。S1再実行で上書きされる)
'     S2 JSON -> s2_prev_json (15章 S2 user の {{prevS2Json}} の唯一の供給元=
'                              S2の status 更新〔confirmed/rejected/new〕の材料)
'     notes   -> 元の data_key(input_field_notes / input_hearing_answers)
'   いずれも【その枠が空のときだけ】書く。取込が作業中の入力を黙って上書きする
'   とデータ喪失経路になる(17章 T-42 観点(1))。
' ============================================================================
Public Function ImportCompanyFile(ByVal filePath As String, ByVal caseId As String) As Boolean
    On Error GoTo Failed

    If Not modCaseStore.IsValidCaseId(caseId) Then Exit Function

    Dim wb As Object
    Set wb = modCompanyFile2.DossierOpen(filePath, True)
    If wb Is Nothing Then
        modLog.LogError "E0603", CF_SRC & ".ImportCompanyFile", "open_failed"
        Exit Function
    End If

    Dim roundNo As Long
    roundNo = modCompanyFile2.SheetMaxRound(modCompanyFile2.DossierSheet(wb, CF_SHEET_ROUNDS))

    Dim restored As Long
    If roundNo > 0 Then
        If RestoreIfEmpty(caseId, "s1_json", ReadRoundJson(wb, roundNo, "s1_json")) Then restored = restored + 1
        If RestoreIfEmpty(caseId, "s2_prev_json", ReadRoundJson(wb, roundNo, "s2_json")) Then restored = restored + 1
        restored = restored + RestoreNotes(wb, caseId, roundNo)
    End If

    modCompanyFile2.DossierClose wb
    Set wb = Nothing

    modLog.LogUsage "dossier_import", caseId, "round=" & CStr(roundNo) & " restored=" & CStr(restored)
    ImportCompanyFile = (restored > 0)
    Exit Function

Failed:
    modLog.LogError "E0603", CF_SRC & ".ImportCompanyFile", "read_failed", Err.Number
    ImportCompanyFile = False
End Function

' ----------------------------------------------------------------------------
' 書き出し(シートごと)
' ----------------------------------------------------------------------------

' dossier_meta は1社1行。created_at は既存値を残し updated_at だけ更新する。
Private Sub WriteMeta(ByVal wb As Object, ByRef ctx As TCaseCtx, _
                      ByVal piiText As String, ByVal confirmedAt As String, _
                      ByVal caseId As String)
    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, CF_SHEET_META)
    If ws Is Nothing Then Exit Sub

    Dim hdr As Variant
    hdr = modCompanyFile2.SheetBlock(ws, modCompanyFile2.SheetLastRow(ws))

    Dim stampText As String
    stampText = modUtil.NowStamp()

    Dim createdText As String
    createdText = modCompanyFile2.BlockText(hdr, 2, modUtil.FindHeaderCol(hdr, "created_at"))
    If LenB(createdText) = 0 Then createdText = stampText

    modCompanyFile2.SheetPutText ws, hdr, 2, "company_id", _
            modUtilText.Fnv1a64Hex(modUtilText.NormalizeForHash(ctx.company))
    modCompanyFile2.SheetPutText ws, hdr, 2, "company", ctx.company
    modCompanyFile2.SheetPutText ws, hdr, 2, "industry_code", ctx.industry_code
    modCompanyFile2.SheetPutText ws, hdr, 2, "schema_version", modConfig.GetStr("app_version", CF_SCHEMA_FALLBACK)
    modCompanyFile2.SheetPutText ws, hdr, 2, "created_at", createdText
    modCompanyFile2.SheetPutText ws, hdr, 2, "updated_at", stampText
    modCompanyFile2.SheetPutText ws, hdr, 2, "pii_scan_result", PiiResultText(piiText)
    modCompanyFile2.SheetPutText ws, hdr, 2, "pii_confirmed_at", Trim$(confirmedAt)
    modCompanyFile2.SheetPutText ws, hdr, 2, "owner", OwnerName()

    If LenB(piiText) > 0 Then
        modLog.LogUsage "dossier_pii_recorded", caseId, _
                        "confirmed=" & CStr(LenB(Trim$(confirmedAt)) > 0)
    End If
End Sub

' dossier_profile は「最新の1本」だけを持つ。前の内容は消して積み直す。
Private Sub WriteProfile(ByVal wb As Object, ByVal caseId As String, ByVal roundNo As Long)
    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, CF_SHEET_PROFILE)
    If ws Is Nothing Then Exit Sub

    modCompanyFile2.SheetClearRows ws

    Dim s1Text As String
    s1Text = modCaseStore.ResolveStepJson(caseId, 1)
    If LenB(s1Text) = 0 Then Exit Sub

    Dim parts() As String
    parts = modUtil.SplitForCells(s1Text, CF_CHUNK_CHARS)

    Dim hdr As Variant
    hdr = modCompanyFile2.SheetBlock(ws, 2)
    Dim stampText As String
    stampText = modUtil.NowStamp()

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim wr As Long
        wr = 2 + (i - LBound(parts))
        modCompanyFile2.SheetPutText ws, hdr, wr, "case_id", caseId
        modCompanyFile2.SheetPutNum ws, hdr, wr, "round_no", roundNo
        modCompanyFile2.SheetPutNum ws, hdr, wr, "seq", i - LBound(parts) + 1
        modCompanyFile2.SheetPutText ws, hdr, wr, "s1_json", parts(i)
        modCompanyFile2.SheetPutText ws, hdr, wr, "updated_at", stampText
    Next i

    ' strategy_outlook(13章§2.12)は「開いた瞬間に前回の理解が戻る」ための
    ' 見出しなので先頭行にだけ展開する。正はあくまで s1_json 側。
    modCompanyFile2.SheetPutText ws, hdr, 2, "mvv", modJsonLite.GetStr(s1Text, "mvv")
    modCompanyFile2.SheetPutText ws, hdr, 2, "aspirations", ArrayText(s1Text, "aspirations")
    modCompanyFile2.SheetPutText ws, hdr, 2, "market_context", modJsonLite.GetStr(s1Text, "market_context")
End Sub

' dossier_rounds は1ラウンド=1行(32,000字超のときだけ seq で行が増える)。
Private Sub WriteRounds(ByVal wb As Object, ByVal caseId As String, _
                        ByRef ctx As TCaseCtx, ByVal roundNo As Long)
    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, CF_SHEET_ROUNDS)
    If ws Is Nothing Then Exit Sub

    modCompanyFile2.SheetDropRound ws, caseId, roundNo

    Dim s1Parts() As String, s2Parts() As String, s3Parts() As String
    s1Parts = modUtil.SplitForCells(modCaseStore.ResolveStepJson(caseId, 1), CF_CHUNK_CHARS)
    s2Parts = modUtil.SplitForCells(modCaseStore.ResolveStepJson(caseId, 2), CF_CHUNK_CHARS)
    s3Parts = modUtil.SplitForCells(modCaseStore.ResolveStepJson(caseId, 3), CF_CHUNK_CHARS)

    Dim rowCount As Long
    rowCount = PartCount(s1Parts)
    If PartCount(s2Parts) > rowCount Then rowCount = PartCount(s2Parts)
    If PartCount(s3Parts) > rowCount Then rowCount = PartCount(s3Parts)
    If rowCount < 1 Then rowCount = 1

    Dim hdr As Variant
    hdr = modCompanyFile2.SheetBlock(ws, 2)
    Dim lastRow As Long
    lastRow = modCompanyFile2.SheetLastRow(ws)
    If lastRow < 1 Then lastRow = 1

    Dim stampText As String
    stampText = modUtil.NowStamp()

    Dim i As Long
    For i = 1 To rowCount
        Dim wr As Long
        wr = lastRow + i
        modCompanyFile2.SheetPutNum ws, hdr, wr, "round_no", roundNo
        modCompanyFile2.SheetPutText ws, hdr, wr, "case_id", caseId
        modCompanyFile2.SheetPutText ws, hdr, wr, "executed_at", stampText
        modCompanyFile2.SheetPutText ws, hdr, wr, "dossier_tier", ctx.dossier_tier
        modCompanyFile2.SheetPutNum ws, hdr, wr, "seq", i
        modCompanyFile2.SheetPutText ws, hdr, wr, "s1_json", PartAt(s1Parts, i)
        modCompanyFile2.SheetPutText ws, hdr, wr, "s2_json", PartAt(s2Parts, i)
        modCompanyFile2.SheetPutText ws, hdr, wr, "s3_json", PartAt(s3Parts, i)
    Next i
End Sub

' dossier_notes は現場メモ・ヒアリング回答の原文累積(field_insights タグつき)。
Private Sub WriteNotes(ByVal wb As Object, ByVal caseId As String, ByVal roundNo As Long)
    Dim ws As Object
    Set ws = modCompanyFile2.DossierSheet(wb, CF_SHEET_NOTES)
    If ws Is Nothing Then Exit Sub

    modCompanyFile2.SheetDropRound ws, caseId, roundNo

    Dim hdr As Variant
    hdr = modCompanyFile2.SheetBlock(ws, 2)
    Dim stampText As String
    stampText = modUtil.NowStamp()

    Dim keys() As String
    keys = modUtil.SplitKeepNonEmpty(CF_NOTE_KEYS, CF_SEP)

    Dim k As Long
    For k = LBound(keys) To UBound(keys)
        Dim bodyText As String
        bodyText = modCaseStore.LoadData(caseId, keys(k))
        If LenB(bodyText) > 0 Then
            Dim parts() As String
            parts = modUtil.SplitForCells(bodyText, CF_CHUNK_CHARS)
            Dim i As Long
            For i = LBound(parts) To UBound(parts)
                Dim wr As Long
                wr = modCompanyFile2.SheetLastRow(ws) + 1
                modCompanyFile2.SheetPutNum ws, hdr, wr, "round_no", roundNo
                modCompanyFile2.SheetPutText ws, hdr, wr, "case_id", caseId
                modCompanyFile2.SheetPutText ws, hdr, wr, "note_kind", keys(k)
                modCompanyFile2.SheetPutText ws, hdr, wr, "tag", CF_NOTE_TAG
                modCompanyFile2.SheetPutNum ws, hdr, wr, "seq", i - LBound(parts) + 1
                modCompanyFile2.SheetPutText ws, hdr, wr, "content", parts(i)
                modCompanyFile2.SheetPutText ws, hdr, wr, "saved_at", stampText
            Next i
        End If
    Next k
End Sub

' ----------------------------------------------------------------------------
' 2段検証と取込の読み側
' ----------------------------------------------------------------------------

' 書き出したはずの中身を、閉じたファイルを開き直して読み比べる(PoC modPack の
' 2段検証)。「書けたつもりで空ファイルを配る」事故をここ1箇所で止める。
Private Function VerifyRoundTrip(ByVal pathText As String, ByVal caseId As String, _
                                 ByVal roundNo As Long, ByVal expectHash As String) As Boolean
    On Error GoTo Failed

    Dim wb As Object
    Set wb = modCompanyFile2.DossierOpen(pathText, True)
    If wb Is Nothing Then Exit Function

    Dim actual As String
    actual = ReadRoundJson(wb, roundNo, "s1_json") & vbLf & _
             ReadRoundJson(wb, roundNo, "s2_json") & vbLf & _
             ReadRoundJson(wb, roundNo, "s3_json") & vbLf & _
             NotesText(vbNullString, wb, roundNo)

    modCompanyFile2.DossierClose wb
    Set wb = Nothing

    VerifyRoundTrip = (modUtilText.Fnv1a64Hex(actual) = expectHash)
    Exit Function

Failed:
    VerifyRoundTrip = False
End Function

' 照合の期待値。VerifyRoundTrip が組み立てる文字列と同じ順・同じ区切りで作る。
Private Function PayloadHash(ByVal caseId As String) As String
    Dim acc As String
    acc = modCaseStore.ResolveStepJson(caseId, 1) & vbLf & _
          modCaseStore.ResolveStepJson(caseId, 2) & vbLf & _
          modCaseStore.ResolveStepJson(caseId, 3) & vbLf & _
          NotesText(caseId, Nothing, 0)
    PayloadHash = modUtilText.Fnv1a64Hex(acc)
End Function

' notes 群を1本の文字列にする。wb=Nothing なら case_data 側(書く前の期待値)、
' wb を渡せば企業ファイル側(書いたあとの実際値)を読む。同じ関数で作るので
' 「並べ方が違うせいで照合が落ちる」ずれが構造的に起こらない。
Private Function NotesText(ByVal caseId As String, ByVal wb As Object, _
                           ByVal roundNo As Long) As String
    Dim keys() As String
    keys = modUtil.SplitKeepNonEmpty(CF_NOTE_KEYS, CF_SEP)
    Dim acc As String
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then acc = acc & vbLf
        acc = acc & keys(i) & "="
        If wb Is Nothing Then
            acc = acc & modCaseStore.LoadData(caseId, keys(i))
        Else
            acc = acc & ReadNoteOf(wb, roundNo, keys(i))
        End If
    Next i
    NotesText = acc
End Function

Private Function ReadRoundJson(ByVal wb As Object, ByVal roundNo As Long, _
                               ByVal colName As String) As String
    ReadRoundJson = modCompanyFile2.SheetJoinBySeq( _
        modCompanyFile2.DossierSheet(wb, CF_SHEET_ROUNDS), roundNo, vbNullString, colName)
End Function

Private Function ReadNoteOf(ByVal wb As Object, ByVal roundNo As Long, _
                            ByVal noteKind As String) As String
    ReadNoteOf = modCompanyFile2.SheetJoinBySeq( _
        modCompanyFile2.DossierSheet(wb, CF_SHEET_NOTES), roundNo, noteKind, "content")
End Function

' 空いている枠にだけ書き戻す(取込が作業中の入力を上書きしない)。
Private Function RestoreIfEmpty(ByVal caseId As String, ByVal dataKey As String, _
                                ByVal bodyText As String) As Boolean
    If LenB(bodyText) = 0 Then Exit Function
    If LenB(modCaseStore.LoadData(caseId, dataKey)) > 0 Then Exit Function
    RestoreIfEmpty = modCaseStore.SaveData(caseId, dataKey, bodyText)
End Function

Private Function RestoreNotes(ByVal wb As Object, ByVal caseId As String, _
                              ByVal roundNo As Long) As Long
    Dim keys() As String
    keys = modUtil.SplitKeepNonEmpty(CF_NOTE_KEYS, CF_SEP)
    Dim i As Long
    Dim n As Long
    For i = LBound(keys) To UBound(keys)
        If RestoreIfEmpty(caseId, keys(i), ReadNoteOf(wb, roundNo, keys(i))) Then n = n + 1
    Next i
    RestoreNotes = n
End Function

' ----------------------------------------------------------------------------
' 小道具
' ----------------------------------------------------------------------------

Private Sub AddReport(ByRef buf() As String, ByRef cnt As Long, _
                      ByVal bodyText As String, ByVal whereNote As String)
    If LenB(bodyText) = 0 Then Exit Sub
    Dim reportText As String
    reportText = modPii.ScanReport(bodyText, whereNote)
    If LenB(reportText) > 0 Then modUtil.BufAdd buf, cnt, reportText
End Sub

' 裁定書9 B8・裁定書10 m6: 当該ブックが同一プロセスで開いたままかを
'   FullName(フルパス)で調べる。ファイル名のみの比較だと別フォルダの同名
'   ブックを「開いている」と誤認し、企業ドシエ書出が失敗扱いになるため。
'   判定に失敗したときは True(=開いている扱い)へ倒し、無効な2段検証で
'   合格を出す方向へは倒さない(fail-closed)。
Private Function BookStillOpen(ByVal pathText As String) As Boolean
    On Error GoTo Unknown0
    Dim wantPath As String
    wantPath = Trim$(pathText)
    If LenB(wantPath) = 0 Then Exit Function

    ' 裁定書11 Q8(両立案): フルパス一致=True。**ファイル名が一致してフルパスが
    ' 一致しないときもTrue**へ倒す(UNC とマップドライブ・8.3短縮名など同じブック
    ' でも表記が違いうるため、ここでFalseを返すと開いたままのブックに対して
    ' VerifyRoundTrip が走り、B8が塞いだ2段検証の無効化が再現する)。
    ' ファイル名まで違うときだけ False(別フォルダの同名ブックの誤検知は避ける)。
    Dim wantName As String
    wantName = BaseNameOf(wantPath)

    Dim i As Long
    For i = 1 To Application.Workbooks.Count
        If StrComp(Application.Workbooks(i).FullName, wantPath, vbTextCompare) = 0 Then
            BookStillOpen = True
            Exit Function
        End If
        If LenB(wantName) > 0 Then
            If StrComp(Application.Workbooks(i).Name, wantName, vbTextCompare) = 0 Then
                BookStillOpen = True
                Exit Function
            End If
        End If
    Next i
    Exit Function
Unknown0:
    BookStillOpen = True
End Function

' パスの末尾(ファイル名)。区切りが無ければ全体。
Private Function BaseNameOf(ByVal pathText As String) As String
    Dim p As Long
    p = InStrRev(pathText, "\")
    Dim q As Long
    q = InStrRev(pathText, "/")
    If q > p Then p = q
    If p <= 0 Then
        BaseNameOf = pathText
    Else
        BaseNameOf = Mid$(pathText, p + 1)
    End If
End Function

' Application.UserName(13章§2.1 owner と同じ扱い)。取れない環境では ""。
Private Function OwnerName() As String
    On Error GoTo NoName
    OwnerName = CStr(Application.UserName)
    Exit Function
NoName:
    OwnerName = vbNullString
End Function

' dossier_meta に残す走査結果。検知ゼロも「走査した事実」として残す。
Private Function PiiResultText(ByVal piiText As String) As String
    If LenB(piiText) = 0 Then
        PiiResultText = "clean/" & modUtil.NowStamp()
        Exit Function
    End If
    PiiResultText = modUtil.SafeLeft("detected/" & Replace(piiText, vbLf, " "), 400)
End Function

' JSONの文字列配列を「; 」連結の1セル形式へ(13章§2.2 セル格納規約)。
Private Function ArrayText(ByVal jsonText As String, ByVal keyName As String) As String
    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(jsonText, keyName)
    Dim acc As String
    Dim i As Long
    For i = 1 To items.Count
        If i > 1 Then acc = acc & "; "
        acc = acc & CStr(items(i))
    Next i
    ArrayText = acc
End Function

Private Function PartCount(ByRef parts() As String) As Long
    On Error GoTo Zero0
    PartCount = UBound(parts) - LBound(parts) + 1
    Exit Function
Zero0:
    PartCount = 0
End Function

Private Function PartAt(ByRef parts() As String, ByVal idx As Long) As String
    On Error GoTo Empty0
    Dim k As Long
    k = LBound(parts) + idx - 1
    If k < LBound(parts) Or k > UBound(parts) Then Exit Function
    PartAt = parts(k)
    Exit Function
Empty0:
    PartAt = vbNullString
End Function

Private Function TrimTrailingSep(ByVal dirPath As String) As String
    Dim t As String
    t = Trim$(dirPath)
    Do While Len(t) > 0
        If Right$(t, 1) <> "\" And Right$(t, 1) <> "/" Then Exit Do
        t = Left$(t, Len(t) - 1)
    Loop
    TrimTrailingSep = t
End Function
