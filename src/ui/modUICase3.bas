Attribute VB_Name = "modUICase3"
Option Explicit

' ============================================================================
' modUICase3 - 案件入力(ui層・T-31)
' ----------------------------------------------------------------------------
' 30,000字契約(12章§2)による modUICase の分割先。変換表(19章§3)は modUICase
' から引き、シート操作は modUISheet を通す。
'
' 本モジュールが担う規約:
'   ・13章§2.11 の貼付欄9欄と連番「続き」欄。組立は**末尾番号の昇順に vbLf で
'     連結**して1本にする(16章 E-22。空欄はスキップ)。
'   ・16章 E-01 必須入力の欠落(企業名/業種/HP、renewal時は現契約)=E0101 で実行
'     させない。E-02 入力が薄い(HP<200字)は警告のみで続行可。
'   ・16章 E-04 外部由来テキストは modUtilText.SanitizeInput を通してから保存。
'   ・16章 E-05(1) 貼付欄のPII検知は**保存をブロック**。フィードバック・判断台帳
'     (E-05(4)(5))は modUICase4 が受け持つ。
'   ・16章 E-31 匿名化(ci_anonymize)。置換規則の実装は modUICase が持つ。
'
' 案件一覧への書込について(申し送り):
'   14章§6 は案件一覧の**属性列**(channel / kanji / bid / reins / industry_name /
'   other_insurers / dossier_tier / s4_variant)の書込口を宣言していないが、
'   13章§2.1 は人が画面で埋めると明記しているため ui層が当該行へ直接書く
'   (§6の裁定待ち)。**状態遷移・採番・case_data は modCaseStore が唯一の口**。
' ============================================================================

Private Const U3_SRC As String = "modUICase3"
Private Const U3_SHEET As String = "ナビ"
Private Const U3_CASES As String = "案件一覧"
Private Const U3_SCAN_COLS As Long = 32
' 13章§2.11(裁定書10 M3): 続き欄への分割幅。JoinField が**区切り文字なし**で
' 連結し SplitForCells と完全に可逆なので、余白を引かず1セル上限ちょうど。
Private Const U3_CHUNK As Long = 32000
Private Const U3_THIN_HP As Long = 200        ' 16章 E-02: HPが薄いと判断する字数
' 裁定書12 V1(13章§2.11): 新規モードの固定マーカー。HOMEの[＋新規案件]がこの
' 値を ci_case_id へ書き、CaseSave はこの値のときだけ採番する(採番の根拠を
' 揮発変数から**ブックに残るセル**へ移す)。IsValidCaseId は案件ID書式だけを
' 通すのでこの値を有効IDと誤認しない。
Public Const U3_NEW_MARK As String = "(新規)"
' 表示case_idと保存先が食い違うときの逐語文言(13章§2.12 B1と同作法)。
Private Const U3_MSG_MISMATCH As String = _
    "画面の案件と保存先が一致しません。再描画してください"
Private Const U3_RESEARCH_ROOM As Long = 20   ' 追加収集ブロックの表示上限行

' 画面制御用のモジュール変数(裁定書9 B15/B22)。**永続でない画面制御の状態**で
' あり、14章§6の「状態保持の例外」への登録は要らない(役割はここに明記する)。
'   gDrawState  = 直近の DrawCaseInput の結末(0=未実行 / 1=描き切った / 2=失敗)
'   gDrawCaseId = その DrawCaseInput が対象にした案件ID
'   gOverBuf    = 続き欄に収まりきらなかった欄の data_key(";"区切り)
Private Const U3_ST_NONE As Long = 0
Private Const U3_ST_OK As Long = 1
Private Const U3_ST_FAILED As Long = 2
Private gDrawState As Long
Private gDrawCaseId As String
Private gOverBuf As String

' ============================================================================
' 13章§2.11 貼付欄の定義(data_key|本欄|続き欄(;区切り)|字数欄|画面ラベル|必須)
'   必須 0=任意 / 1=常に必須 / 2=renewal時のみ必須
' ============================================================================
Private Function PasteTable() As String
    Dim s As String
    s = s & "input_hp|ci_paste_hp_1|ci_paste_hp_2;ci_paste_hp_3|ci_count_hp|HPテキスト|1" & vbLf
    s = s & "input_yuho|ci_paste_yuho_1|ci_paste_yuho_2;ci_paste_yuho_3|ci_count_yuho|有報リスク章|0" & vbLf
    s = s & "input_memo|ci_paste_memo_1||ci_count_memo|営業メモ|0" & vbLf
    s = s & "input_contract|ci_paste_contract_1|ci_paste_contract_2;ci_paste_contract_3|ci_count_contract|現契約サマリ|2" & vbLf
    s = s & "input_prev_renewal|ci_paste_prev_renewal_1||ci_count_prev_renewal|前回更新メモ|0" & vbLf
    s = s & "input_dossier|ci_paste_dossier_1|ci_paste_dossier_2;ci_paste_dossier_3|ci_count_dossier|追加ドシエ|0" & vbLf
    s = s & "input_field_notes|ci_paste_field_notes_1||ci_count_field_notes|現場メモ|0" & vbLf
    s = s & "input_coverage_note|ci_paste_coverage_note_1||ci_count_coverage_note|付保の見立て|0" & vbLf
    s = s & "input_hearing_answers|ci_paste_hearing_answers_1||ci_count_hearing_answers|ヒアリング回答|0"
    PasteTable = s
End Function

' 案件入力の属性欄 -> 案件一覧の列(名前付きレンジ|列名|enumグループ)。
'   enumグループが空の欄は自由記述(変換しない)。
Private Function AttrTable() As String
    Dim s As String
    s = s & "ci_case_type|case_type|case_type" & vbLf
    s = s & "ci_company|company|" & vbLf
    s = s & "ci_industry_code|industry_code|" & vbLf
    s = s & "ci_industry_name|industry_name|" & vbLf
    s = s & "ci_dossier_tier|dossier_tier|dossier_tier" & vbLf
    s = s & "ci_channel|channel|channel" & vbLf
    s = s & "ci_kanji|kanji|kanji" & vbLf
    s = s & "ci_bid|bid|bid" & vbLf
    s = s & "ci_reins|reins|reins" & vbLf
    s = s & "ci_other_insurers|other_insurers|" & vbLf
    s = s & "ci_s4_variant|s4_variant|s4_variant"
    AttrTable = s
End Function

' ============================================================================
' 図形ボタン(11章 案件入力ワイヤー)
' ============================================================================
Public Sub EnsureCaseButtons()
    On Error Resume Next

    Dim ws As Object
    Set ws = modUISheet.SheetOf(U3_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = RowOfNamed("ci_case_id")
    modUISheet.EnsureButton ws, "btn_ci_save", "保存して戻る", r, 4, 100#, _
                            "modUICase3.CaseSave"
    modUISheet.EnsureButton ws, "btn_ci_count", "字数を数える", r, 5, 100#, _
                            "modUICase3.CaseCountChars"

    r = RowOfNamed("ci_paste_contract_1")
    modUISheet.EnsureButton ws, "btn_ci_anon", "匿名化の切替", r, 4, 100#, _
                            "modUICase3.CaseToggleAnonymize"

End Sub

Private Function RowOfNamed(ByVal rangeName As String) As Long
    Dim cell As Object
    Set cell = modUISheet.NamedCell(rangeName)
    If cell Is Nothing Then
        RowOfNamed = 1
        Exit Function
    End If
    On Error GoTo One1
    RowOfNamed = cell.row
    Exit Function
One1:
    RowOfNamed = 1
End Function

' ============================================================================
' 案件一覧の直接アクセス(本モジュール冒頭の申し送りを参照)
' ============================================================================

' 案件一覧の1セルを列名で読む。見つからなければ ""。
Public Function CaseCellText(ByVal caseId As String, ByVal colName As String) As String
    On Error GoTo NoValue

    Dim ws As Object
    Dim rowNo As Long
    Dim hdr As Variant
    If Not LocateCase(caseId, ws, hdr, rowNo) Then Exit Function

    Dim colNo As Long
    colNo = modUtil.FindHeaderCol(hdr, colName)
    If colNo <= 0 Then Exit Function

    CaseCellText = Trim$(CStr(ws.Cells(rowNo, colNo).Value))
    Exit Function
NoValue:
    CaseCellText = vbNullString
End Function

' 案件一覧の複数セルを1回の行探索で書く。pairsText は1行1件で
'   `列名 <TAB> 値`(vbLf区切り)。書けた件数を返す。
Public Function SetCaseCells(ByVal caseId As String, ByVal pairsText As String) As Long
    On Error GoTo Failed

    Dim ws As Object
    Dim rowNo As Long
    Dim hdr As Variant
    If Not LocateCase(caseId, ws, hdr, rowNo) Then Exit Function

    Dim lines() As String
    lines = Split(pairsText, vbLf)

    Dim i As Long
    Dim tabPos As Long
    Dim colName As String
    Dim valueText As String
    Dim colNo As Long
    Dim done As Long
    For i = LBound(lines) To UBound(lines)
        tabPos = InStr(1, lines(i), vbTab, vbBinaryCompare)
        If tabPos > 1 Then
            colName = Left$(lines(i), tabPos - 1)
            valueText = Mid$(lines(i), tabPos + 1)
            colNo = modUtil.FindHeaderCol(hdr, colName)
            If colNo > 0 Then
                modUISheet.PutText ws, rowNo, colNo, valueText, U3_CASES & "/" & colName
                done = done + 1
            End If
        End If
    Next i

    ' 13章§2.1: 属性を変えたら更新時刻も動かす。
    colNo = modUtil.FindHeaderCol(hdr, "updated_at")
    If colNo > 0 Then
        modUISheet.PutText ws, rowNo, colNo, modUtil.NowStamp(), U3_CASES & "/updated_at"
    End If

    SetCaseCells = done
    Exit Function
Failed:
    modLog.LogError "E0603", U3_SRC & ".SetCaseCells", "write_failed", Err.Number
    SetCaseCells = 0
End Function

Private Function LocateCase(ByVal caseId As String, ByRef ws As Object, _
                            ByRef hdr As Variant, ByRef rowNo As Long) As Boolean
    rowNo = 0
    If Not modCaseStore.IsValidCaseId(caseId) Then Exit Function

    Set ws = modUISheet.SheetOf(U3_CASES)
    If ws Is Nothing Then Exit Function

    hdr = modUISheet.HeaderOf(ws, 1, 1, U3_SCAN_COLS)
    If IsEmpty(hdr) Then Exit Function

    Dim colNo As Long
    colNo = modUtil.FindHeaderCol(hdr, "case_id")
    If colNo <= 0 Then Exit Function

    Dim lastRow As Long
    lastRow = modUISheet.LastRowOf(ws)

    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, colNo).Value)) = caseId Then
            rowNo = r
            LocateCase = True
            Exit Function
        End If
    Next r
End Function

' ============================================================================
' 新規案件(HOMEの[＋新規案件])
' ----------------------------------------------------------------------------
'   採番は modCaseStore.NewCase が唯一の口。企業名・業種は案件入力の欄から
'   受け取る(空なら作らない=16章 E-01)。
' ============================================================================
Public Function CreateCaseFromSheet() As String
    On Error GoTo Failed

    Dim company As String
    Dim industryCode As String
    Dim caseTypeJa As String
    company = modUISheet.ReadNamed("ci_company")
    industryCode = modUISheet.ReadNamed("ci_industry_code")
    caseTypeJa = modUISheet.ReadNamed("ci_case_type")

    If LenB(company) = 0 Or LenB(industryCode) = 0 Then
        modLog.LogError "E0101", U3_SRC & ".CreateCaseFromSheet", "company_or_industry_empty"
        Exit Function
    End If

    Dim caseType As String
    caseType = modUICase.EnumEn("case_type", caseTypeJa)
    If LenB(caseType) = 0 Then caseType = "new"

    Dim caseId As String
    caseId = modCaseStore.NewCase(company, industryCode, caseType)
    If LenB(caseId) = 0 Then Exit Function

    modUISheet.WriteNamed "ci_case_id", caseId
    CreateCaseFromSheet = caseId
    Exit Function

Failed:
    modLog.LogError "E0101", U3_SRC & ".CreateCaseFromSheet", "new_case_failed", Err.Number
    CreateCaseFromSheet = vbNullString
End Function

' ============================================================================
' 案件入力の描画(案件 -> 画面)
' ============================================================================
Public Sub DrawCaseInput(ByVal caseId As String)
    On Error Resume Next

    ' 裁定書9 B15: 描き切るまでは「失敗」として扱う。途中で抜けた画面(案件を
    ' 読めない・名前付きレンジが無い等)を保存すると、空欄がそのまま case_data
    ' へ書かれて貼付元の原文が消える。最後まで到達したときだけ成功印を立てる。
    gDrawCaseId = caseId
    gDrawState = U3_ST_FAILED
    gOverBuf = vbNullString

    ' 裁定書11 Q1/裁定書12 V1: 表示case_id は**描き切ったときだけ**書く。ここで
    ' 空へ戻せば、途中で抜けた画面・切捨てが起きた画面は「どの案件のものでもない」
    ' 状態でブックに残り、CaseSave の3値判定が保存を止める(新規モードの
    ' マーカーもここで消えるので、描画した画面が採番へ倒れることはない)。
    modUISheet.WriteNamed "ci_case_id", vbNullString

    If Not modCaseStore.IsValidCaseId(caseId) Then Exit Sub

    Dim ctx As TCaseCtx
    Dim roundNo As Long
    Dim qualityMode As String
    Dim s4Variant As String
    Dim tierText As String
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, roundNo, qualityMode, s4Variant, tierText) Then
        Exit Sub
    End If

    ' 属性欄(機械値 -> 日本語ラベル)。
    Dim rows1() As String
    rows1 = Split(AttrTable(), vbLf)
    Dim i As Long
    Dim rangeName As String
    Dim colName As String
    Dim groupName As String
    Dim rawText As String
    For i = LBound(rows1) To UBound(rows1)
        SplitAttr rows1(i), rangeName, colName, groupName
        If LenB(rangeName) > 0 Then
            rawText = CaseCellText(caseId, colName)
            If LenB(groupName) > 0 Then
                modUISheet.WriteNamed rangeName, modUICase.EnumJa(groupName, rawText)
            Else
                modUISheet.WriteNamed rangeName, rawText
            End If
        End If
    Next i

    ' 貼付欄(case_data -> 本欄・続き欄。16章 E-31 の復元を通す)。
    Dim specs() As String
    specs = Split(PasteTable(), vbLf)
    Dim dataKey As String
    Dim baseRange As String
    Dim contRanges As String
    Dim countRange As String
    Dim labelText As String
    Dim requiredKind As Long
    For i = LBound(specs) To UBound(specs)
        SplitPaste specs(i), dataKey, baseRange, contRanges, countRange, labelText, requiredKind
        If LenB(baseRange) > 0 Then
            DrawPasteField caseId, ctx.company, dataKey, baseRange, contRanges
        End If
    Next i

    DrawResearchBlock caseId
    CountChars

    ' 裁定書11 Q1: 続き欄に収まらなかった欄がある描画では表示case_idを空のまま
    ' にする。gOverBuf は揮発性(未捕捉エラー・再コンパイルで消える)なので、
    ' 保存ブロックの根拠は**ブックに残るこのセル**が持つ(13章§2.11)。
    If LenB(gOverBuf) = 0 Then modUISheet.WriteNamed "ci_case_id", caseId

    gDrawState = U3_ST_OK
End Sub

' 1欄ぶんを本欄・続き欄へ分けて書き戻す(13章§2.11)。
Private Sub DrawPasteField(ByVal caseId As String, ByVal company As String, _
                           ByVal dataKey As String, ByVal baseRange As String, _
                           ByVal contRanges As String)
    Dim content As String
    content = modCaseStore.LoadData(caseId, dataKey)
    content = modUICase.RestoreNames(content, company)

    Dim targets() As String
    targets = Split(baseRange & ";" & contRanges, ";")

    Dim parts() As String
    parts = modUtil.SplitForCells(content, U3_CHUNK)

    Dim slots As Long
    Dim i As Long
    Dim v As String
    For i = LBound(targets) To UBound(targets)
        v = vbNullString
        If LenB(targets(i)) > 0 Then
            slots = slots + 1
            If i - LBound(targets) + LBound(parts) <= UBound(parts) Then
                If LenB(content) > 0 Then v = parts(i - LBound(targets) + LBound(parts))
            End If
            modUISheet.WriteNamed targets(i), v
        End If
    Next i

    ' 裁定書9 B22: 断片数が枠数を超えたら、あふれた分は画面に出ていない。
    ' そのまま保存すると開くたびに末尾が削れていくので、当該欄の保存を止める。
    If LenB(content) > 0 Then
        If UBound(parts) - LBound(parts) + 1 > slots Then NoteOverflow dataKey
    End If
End Sub

' 続き欄に収まらなかった欄を覚え、画面へ出す(裁定書9 B22)。
Private Sub NoteOverflow(ByVal dataKey As String)
    On Error Resume Next
    If InStr(1, ";" & gOverBuf & ";", ";" & dataKey & ";", vbBinaryCompare) = 0 Then
        If LenB(gOverBuf) > 0 Then gOverBuf = gOverBuf & ";"
        gOverBuf = gOverBuf & dataKey
    End If
    modUISheet.WriteNamed "hm_warning", "この欄は続き欄に収まりません（" & gOverBuf & _
        "）。表示しきれていない分が失われるため、この欄は保存しません。"
End Sub

' 当該欄が「続き欄に収まらなかった」印を持つか(裁定書9 B22)。
Private Function IsOverflowed(ByVal dataKey As String) As Boolean
    IsOverflowed = (InStr(1, ";" & gOverBuf & ";", ";" & dataKey & ";", vbBinaryCompare) > 0)
End Function

' 11章 追加収集ブロック(S1の research_requests を2列N行＋各行にコピーボタン)。
Private Sub DrawResearchBlock(ByVal caseId As String)
    On Error Resume Next

    Dim anchor As Object
    Set anchor = modUISheet.NamedCell("ci_research_anchor")
    If anchor Is Nothing Then Exit Sub

    Dim ws As Object
    Set ws = modUISheet.SheetOf(U3_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim r0 As Long
    Dim c0 As Long
    r0 = anchor.row
    c0 = anchor.Column

    modUISheet.DropShapesByPrefix ws, "btncopy_"
    modUISheet.ClearBlock ws, r0 - 1, c0, 2, U3_RESEARCH_ROOM

    Dim s1Json As String
    s1Json = modCaseStore.ResolveStepJson(caseId, 1)
    If LenB(Trim$(s1Json)) = 0 Then Exit Sub

    Dim items As Collection
    Set items = modJsonLite.GetArrayItems(s1Json, "research_requests")

    Dim i As Long
    For i = 1 To items.count
        If i > U3_RESEARCH_ROOM Then Exit For
        modUISheet.PutText ws, r0 + i - 1, c0, _
                           modJsonLite.GetStr(CStr(items(i)), "purpose"), _
                           U3_SHEET & "/purpose"
        modUISheet.PutText ws, r0 + i - 1, c0 + 1, _
                           modJsonLite.GetStr(CStr(items(i)), "prompt_text"), _
                           U3_SHEET & "/prompt_text"
        modUISheet.EnsureButton ws, "btncopy_" & CStr(r0 + i - 1), "コピー", _
                                r0 + i - 1, c0 + 2, 52#, "modUICase4.CopyResearchRow"
    Next i
End Sub

' ============================================================================
' 案件入力の保存(画面 -> 案件)。OnActionハンドラ。
' ============================================================================
Public Sub CaseSave()
    If Not modUIProgress.TryEnterUiLock("案件入力の保存") Then Exit Sub
    On Error GoTo Done

    modUIProgress.ParkFocus

    ' 裁定書12 V1(13章§2.11): ci_case_id の**3値判定**。採番の根拠は揮発変数で
    ' はなくブックに残るこのセル1つだけ。
    '   「(新規)」  = [＋新規案件]が書いた新規モード -> ここで採番する
    '   有効な案件ID = その案件へ保存(SavePasteFields が再度突き合わせる)
    '   空・その他   = 描き切れなかった画面 -> 1欄も書かずに止める(採番へ倒すと
    '                  切り詰まった画面が別案件として確定する)
    Dim caseId As String
    Dim isNew As Boolean
    caseId = Trim$(modUISheet.ReadNamed("ci_case_id"))
    If StrComp(caseId, U3_NEW_MARK, vbBinaryCompare) = 0 Then
        isNew = True
        caseId = CreateCaseFromSheet()
        If LenB(caseId) = 0 Then
            Notice "案件を作成できませんでした。企業名と業種コードをご確認ください。"
            GoTo Done
        End If
    ElseIf Not modCaseStore.IsValidCaseId(caseId) Then
        ' 裁定書12 V3: 成功時と同じHOME遷移で戻し、hm_warning の警告文で
        ' 「保存できなかった」ことが判るようにする(MsgBoxは使わない)。
        modLog.LogError "E0302", U3_SRC & ".CaseSave", "case_id_blank"
        modUISheet.ShowSheet "ナビ"
        modUIHome.RefreshHome
        modUISheet.WriteNamed "hm_warning", U3_MSG_MISMATCH
        GoTo Done
    End If

    If SaveCaseInput(caseId) Then
        If isNew Then modUISheet.WriteNamed "hm_case_id", caseId
        modUISheet.ShowSheet "ナビ"
        modUIHome.RefreshHome
    End If

Done:
    modUIProgress.ExitUiLock
    modUIProgress.ParkFocus
End Sub

' 保存本体。True=保存できた。
Public Function SaveCaseInput(ByVal caseId As String) As Boolean
    On Error GoTo Failed

    Dim caseTypeJa As String
    caseTypeJa = modUISheet.ReadNamed("ci_case_type")

    Dim caseType As String
    caseType = modUICase.EnumEn("case_type", caseTypeJa)
    If LenB(caseType) = 0 Then caseType = "new"

    ' (0) 裁定書9 B15: 直前の再描画に失敗した画面は「この案件の内容」ではない。
    '     そのまま保存すると空欄が case_data を上書きして貼付元の原文が消える。
    If gDrawState = U3_ST_FAILED Then
        If StrComp(gDrawCaseId, caseId, vbBinaryCompare) = 0 Then
            modLog.LogError "E0101", U3_SRC & ".SaveCaseInput", "draw_failed_block"
            Notice "案件入力を画面へ読み込めていないため保存しませんでした。" & _
                   "HOMEの[案件入力を開く]で開き直してください。"
            Exit Function
        End If
    End If

    ' (1) 16章 E-01 必須入力の検査。欠けていれば1件も保存しない。
    Dim missing1 As String
    missing1 = MissingRequired(caseType)
    If LenB(missing1) > 0 Then
        modLog.LogError "E0101", U3_SRC & ".SaveCaseInput", "required_missing"
        Notice "次の欄が空です。入力してから保存してください: " & missing1
        Exit Function
    End If

    ' (2) 16章 E-05(1) PII走査。検知したら保存をブロックする。
    Dim piiText As String
    piiText = ScanPasteFields()
    If LenB(piiText) > 0 Then
        modLog.LogError "E0103", U3_SRC & ".SaveCaseInput", piiText
        Notice "個人情報らしき記述を検知したため保存しませんでした。該当箇所を直してください: " & _
               modUtil.SafeLeft(piiText, 300)
        Exit Function
    End If

    ' (3) 属性欄を案件一覧へ(日本語ラベル -> 機械値)。
    SaveAttributes caseId

    ' (4) 貼付欄を case_data へ(E-04 の浄化と E-31 の匿名化を通す)。
    '     裁定書11 Q1: 一致検査で止まったときは成功案内を出さずに抜ける。
    If Not SavePasteFields(caseId) Then Exit Function

    ' (5) 16章 E-02 入力が薄い場合の警告(続行可)。
    Dim hpLen As Long
    hpLen = Len(JoinField("ci_paste_hp_1", "ci_paste_hp_2;ci_paste_hp_3"))

    ' (6) 裁定書10 M4: 成功時の案内は hm_warning を上書きし MsgBox も出すため、
    '     B22 で保存を落とした欄名を**必ず案内に含める**。含めないと、続き欄に
    '     収まらず1欄まるごと保存しなかったことが利用者に伝わらないまま
    '     「保存できた」という案内だけが残る(サイレント部分保存)。
    Dim overText As String
    If LenB(gOverBuf) > 0 Then
        overText = "続き欄に収まらないため保存しなかった欄があります（" & gOverBuf & "）。"
    End If

    If hpLen < U3_THIN_HP Then
        Notice "HPテキストが" & CStr(hpLen) & "字と少なめです。" & _
               "このまま実行できますが、一般論に近い出力になりやすくなります。" & overText
    ElseIf LenB(overText) > 0 Then
        Notice "案件入力を保存しました。" & overText
    End If

    CountChars
    modLog.LogUsage "case_input_saved", caseId, "hp_len=" & CStr(hpLen)
    SaveCaseInput = True
    Exit Function

Failed:
    modLog.LogError "E0101", U3_SRC & ".SaveCaseInput", "save_failed", Err.Number
    SaveCaseInput = False
End Function

' 16章 E-01: 空の必須欄の一覧(画面ラベル)。空なら ""。
Private Function MissingRequired(ByVal caseType As String) As String
    Dim acc As String

    If LenB(modUISheet.ReadNamed("ci_company")) = 0 Then acc = AddName(acc, "企業名")
    If LenB(modUISheet.ReadNamed("ci_industry_code")) = 0 Then acc = AddName(acc, "業種コード")
    If LenB(modUISheet.ReadNamed("ci_industry_name")) = 0 Then acc = AddName(acc, "業種名")
    If LenB(modUISheet.ReadNamed("ci_channel")) = 0 Then acc = AddName(acc, "取引")
    If LenB(modUISheet.ReadNamed("ci_kanji")) = 0 Then acc = AddName(acc, "幹事")
    If LenB(modUISheet.ReadNamed("ci_bid")) = 0 Then acc = AddName(acc, "BID")
    If LenB(modUISheet.ReadNamed("ci_reins")) = 0 Then acc = AddName(acc, "再保・キャプティブ")

    Dim specs() As String
    specs = Split(PasteTable(), vbLf)
    Dim i As Long
    Dim dataKey As String
    Dim baseRange As String
    Dim contRanges As String
    Dim countRange As String
    Dim labelText As String
    Dim requiredKind As Long
    For i = LBound(specs) To UBound(specs)
        SplitPaste specs(i), dataKey, baseRange, contRanges, countRange, labelText, requiredKind
        If requiredKind = 1 Or (requiredKind = 2 And caseType = "renewal") Then
            If LenB(Trim$(JoinField(baseRange, contRanges))) = 0 Then
                acc = AddName(acc, labelText)
            End If
        End If
    Next i

    MissingRequired = acc
End Function

Private Function AddName(ByVal acc As String, ByVal nameText As String) As String
    If LenB(acc) > 0 Then
        AddName = acc & "／" & nameText
    Else
        AddName = nameText
    End If
End Function

' 16章 E-05(1): 貼付欄9欄(＋続き欄)の走査結果。検知なしなら ""。
Private Function ScanPasteFields() As String
    Dim buf() As String
    Dim cnt As Long
    modUtil.BufInit buf, cnt

    Dim specs() As String
    specs = Split(PasteTable(), vbLf)
    Dim i As Long
    Dim dataKey As String
    Dim baseRange As String
    Dim contRanges As String
    Dim countRange As String
    Dim labelText As String
    Dim requiredKind As Long
    Dim body As String
    Dim report As String
    For i = LBound(specs) To UBound(specs)
        SplitPaste specs(i), dataKey, baseRange, contRanges, countRange, labelText, requiredKind
        body = JoinField(baseRange, contRanges)
        If LenB(body) > 0 Then
            If modPii.HasPii(body) Then
                report = modPii.ScanReport(body, labelText)
                If LenB(report) > 0 Then modUtil.BufAdd buf, cnt, report
            End If
        End If
    Next i

    ScanPasteFields = modUtil.BufText(buf, cnt)
End Function

' 属性欄 -> 案件一覧(日本語ラベル -> 機械値)。
Private Sub SaveAttributes(ByVal caseId As String)
    Dim rows1() As String
    rows1 = Split(AttrTable(), vbLf)

    Dim pairs As String
    Dim i As Long
    Dim rangeName As String
    Dim colName As String
    Dim groupName As String
    Dim shown As String
    Dim stored As String
    For i = LBound(rows1) To UBound(rows1)
        SplitAttr rows1(i), rangeName, colName, groupName
        If LenB(rangeName) > 0 Then
            shown = modUISheet.ReadNamed(rangeName)
            If LenB(groupName) > 0 Then
                stored = modUICase.EnumEn(groupName, shown)
            Else
                stored = shown
            End If
            If LenB(pairs) > 0 Then pairs = pairs & vbLf
            pairs = pairs & colName & vbTab & stored
        End If
    Next i

    SetCaseCells caseId, pairs
End Sub

' 貼付欄 -> case_data。E-04(浄化)と E-31(匿名化)を通してから保存する。
' True=保存した。False=表示case_idと保存先が一致せず1欄も書いていない。
Private Function SavePasteFields(ByVal caseId As String) As Boolean
    ' 裁定書11 Q1: 貼付欄を上書きする前に、**ブックに残る**表示case_idと保存先を
    ' 突き合わせる(CaseSave の3値判定に続く第2の壁)。overflow・描画失敗のあとは
    ' ci_case_id が空なので、揮発変数が消えていてもここで必ず止まる。
    If StrComp(Trim$(modUISheet.ReadNamed("ci_case_id")), caseId, vbBinaryCompare) <> 0 Then
        ' 文言は画面へ直書きする(13章§2.12 B1ガードと同作法。モーダルを出さない
        ' ので、無人実行の層(b)テストからこのガードを検査できる)。
        modLog.LogError "E0302", U3_SRC & ".SavePasteFields", "case_id_mismatch:ci"
        modUISheet.WriteNamed "hm_warning", U3_MSG_MISMATCH
        Exit Function
    End If

    Dim company As String
    company = modUISheet.ReadNamed("ci_company")

    Dim anonymize As Boolean
    anonymize = AnonymizeOn()

    Dim specs() As String
    specs = Split(PasteTable(), vbLf)

    Dim i As Long
    Dim dataKey As String
    Dim baseRange As String
    Dim contRanges As String
    Dim countRange As String
    Dim labelText As String
    Dim requiredKind As Long
    Dim body As String
    Dim hits As Long
    Dim totalHits As Long
    For i = LBound(specs) To UBound(specs)
        SplitPaste specs(i), dataKey, baseRange, contRanges, countRange, labelText, requiredKind
        If Not IsOverflowed(dataKey) Then
            body = modUtilText.SanitizeInput(JoinField(baseRange, contRanges))
            If anonymize Then
                body = modUICase.AnonymizeText(body, company, hits)
                totalHits = totalHits + hits
            End If
            modCaseStore.SaveData caseId, dataKey, body
        End If
    Next i

    If totalHits > 0 Then
        modLog.LogError "E0104", U3_SRC & ".SavePasteFields", "replaced=" & CStr(totalHits)
    End If
    SavePasteFields = True
End Function

' 16章 E-22 / 裁定書10 M3: 本欄・続き欄を末尾番号の昇順に **区切り文字なしで**
' 連結する(空欄はスキップ)。分割側 modUtil.SplitForCells は原文を字数で機械的に
' 切っているだけなので、区切りを入れずに戻せば原文と1字も違わない
' (**完全に可逆**)。旧実装は境界へ vbLf を1つ挿入しており、描画->保存の往復の
' たびに原文が変質し、最終的に枠から溢れて末尾が削れていた。
Private Function JoinField(ByVal baseRange As String, ByVal contRanges As String) As String
    Dim targets() As String
    targets = Split(baseRange & ";" & contRanges, ";")

    Dim acc As String
    Dim i As Long
    Dim v As String
    For i = LBound(targets) To UBound(targets)
        If LenB(targets(i)) > 0 Then
            v = modUISheet.ReadNamed(targets(i))
            If LenB(v) > 0 Then acc = acc & v
        End If
    Next i
    JoinField = acc
End Function

' 16章 E-31: 匿名化の状態(ci_anonymize。既定は config anonymize_default)。
Private Function AnonymizeOn() As Boolean
    Dim v As String
    v = modUISheet.ReadNamed("ci_anonymize")
    If LenB(v) = 0 Then
        AnonymizeOn = modConfig.GetBool("anonymize_default", True)
        Exit Function
    End If
    AnonymizeOn = modConfig.ParseBoolText(v, modConfig.GetBool("anonymize_default", True))
End Function

' [匿名化]ボタン(11章 ④現契約サマリの横)。状態を切り替えるだけで、置換は保存時。
Public Sub CaseToggleAnonymize()
    If Not modUIProgress.TryEnterUiLock("匿名化の切替") Then Exit Sub
    On Error GoTo Done

    Dim nowOn As Boolean
    nowOn = AnonymizeOn()
    If nowOn Then
        modUISheet.WriteNamed "ci_anonymize", "FALSE"
        Notice "匿名化をオフにしました（実名のままLLMへ送ります）。"
    Else
        modUISheet.WriteNamed "ci_anonymize", "TRUE"
        Notice "匿名化をオンにしました（企業名と証券番号らしき英数列を伏せて送ります）。"
    End If

Done:
    modUIProgress.ExitUiLock
End Sub

' ============================================================================
' 文字数カウンタ(13章§2.11・16章 E-03/E-46。数式ではなくVBAが値で書く)
' ============================================================================
Public Sub CaseCountChars()
    If Not modUIProgress.TryEnterUiLock("字数を数える") Then Exit Sub
    On Error GoTo Done
    CountChars
Done:
    modUIProgress.ExitUiLock
End Sub

Private Sub CountChars()
    On Error Resume Next

    Dim specs() As String
    specs = Split(PasteTable(), vbLf)

    Dim i As Long
    Dim dataKey As String
    Dim baseRange As String
    Dim contRanges As String
    Dim countRange As String
    Dim labelText As String
    Dim requiredKind As Long
    Dim n As Long
    Dim total As Long
    For i = LBound(specs) To UBound(specs)
        SplitPaste specs(i), dataKey, baseRange, contRanges, countRange, labelText, requiredKind
        n = Len(JoinField(baseRange, contRanges))
        total = total + n
        If LenB(countRange) > 0 Then modUISheet.WriteNamed countRange, CStr(n)
    Next i

    modUISheet.WriteNamed "ci_count_total", CStr(total)
    modUISheet.MarkNamed "ci_count_total", (total > LimitChars())
End Sub

' 16章 E-03(1): 上限は収集ティアで決まる。
Private Function LimitChars() As Long
    Dim tierJa As String
    tierJa = modUISheet.ReadNamed("ci_dossier_tier")

    Dim tier As String
    tier = modUICase.EnumEn("dossier_tier", tierJa)
    If tier = "t1_quick" Or LenB(tier) = 0 Then
        LimitChars = modConfig.GetLong("max_context_chars", 40000)
    Else
        LimitChars = modConfig.GetLong("t2_max_context_chars", 100000)
    End If
End Function

' ============================================================================
' 共通の小物
' ============================================================================

' 利用者への案内。HOMEの警告欄へ書き、当該シートを見ている場合に備えて
' ダイアログでも知らせる(実行中ではないので E-50(c) の対象外)。
Private Sub Notice(ByVal messageText As String)
    On Error Resume Next
    modUISheet.WriteNamed "hm_warning", messageText
    MsgBox messageText, vbInformation, "リスク提案ナビ"
End Sub

Private Sub SplitAttr(ByVal rowText As String, ByRef rangeName As String, _
                      ByRef colName As String, ByRef groupName As String)
    rangeName = vbNullString
    colName = vbNullString
    groupName = vbNullString

    Dim flds() As String
    flds = Split(rowText, "|")
    If UBound(flds) - LBound(flds) < 2 Then Exit Sub
    rangeName = Trim$(flds(LBound(flds)))
    colName = Trim$(flds(LBound(flds) + 1))
    groupName = Trim$(flds(LBound(flds) + 2))
End Sub

Private Sub SplitPaste(ByVal rowText As String, ByRef dataKey As String, _
                       ByRef baseRange As String, ByRef contRanges As String, _
                       ByRef countRange As String, ByRef labelText As String, _
                       ByRef requiredKind As Long)
    dataKey = vbNullString
    baseRange = vbNullString
    contRanges = vbNullString
    countRange = vbNullString
    labelText = vbNullString
    requiredKind = 0

    Dim flds() As String
    flds = Split(rowText, "|")
    If UBound(flds) - LBound(flds) < 5 Then Exit Sub
    dataKey = Trim$(flds(LBound(flds)))
    baseRange = Trim$(flds(LBound(flds) + 1))
    contRanges = Trim$(flds(LBound(flds) + 2))
    countRange = Trim$(flds(LBound(flds) + 3))
    labelText = Trim$(flds(LBound(flds) + 4))
    requiredKind = CLng(Val(Trim$(flds(LBound(flds) + 5))))
End Sub
