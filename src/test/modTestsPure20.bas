Attribute VB_Name = "modTestsPure20"
Option Explicit

' ============================================================================
' modTestsPure20 - W10(データをブックの外へ・裁定書28)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は**仕様の文だけ**から手で書き出した(17章§1。実装の出力を
'   見てから期待値を合わせない)。
'
' 対象と根拠(全8本):
'   W10M 見出し -> 案件一覧行の写像 3本  13章§2.1 v2.7・裁定書28。
'     01 全列そろい   6列(case_id;company;industry_code;case_type;status;
'                     updated_at)が「列名<TAB>値」で vbLf 連結される
'     02 欠け列は空   企業ファイルに無いキーは**空値の行として出す**
'                     (古い企業ファイルでも落ちない=前方互換)
'     03 新旧優先     PickNewerHeader は updated_at が新しいほうを返し、
'                     **同時刻なら第1引数(本体側)を残す**
'   W10F ファイル名採番(禁止文字) 2本  13章§2.8 手順1・4。
'     01 禁止文字9種と制御文字が名前に1つも残らず、末尾は命名テンプレートの
'        `_企業カルテ.xlsx`(1社1ファイルなので `_2` の連番は付かない)
'     02 同じ会社名は同じ名前・違う会社名は違う名前(8桁が company 由来である
'        ことの帰結。case_id 由来だと同じ会社の2件目が別ファイルになる)
'   W10S schema_version の判定 2本  13章§2.8 v2.7。
'     01 版が空のファイル(W10 より前)は 1.0.0 とみなし、**読める**
'     02 major が現行(3)より新しい 4.0.0 は読まない(fail-closed)。同じ major の
'        2.9.9 / 3.9.9 は読める(列が欠けているだけなら「欠けは空」で読める)
'
'   W10P pii_flag の写像 1本  16章 E-05(7)・13章§2.8 v2.8(統合W10の裁定)。
'     01 dossier_case の列の並び(CaseSheetCols)が案件一覧の全列のあとに
'        schema_version と **pii_flag** をこの順で持ち、PiiFlagOf が
'        「検知あり -> TRUE / 検知なし(空・空白) -> FALSE」を返す。
'        自分の data_dir への自動保存は PII があっても行い、印だけ残すため。
'
' グループ単位の失敗隔離: modTestsPure.bas と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

' 13章§2.8 手順1の禁止文字(仕様の表から手で写した9種)。
Private Const P20_BAD As String = "\/:*?""<>|"
' 13章§2.8 の命名テンプレートの末尾(拡張子込み)。
Private Const P20_TAIL As String = "_企業カルテ.xlsx"

Public Sub RunAll()
    On Error GoTo FA
    T_W10M_HeaderToCaseRow
WB:
    On Error GoTo FB
    T_W10F_CompanyFileName
WC:
    On Error GoTo FC
    T_W10S_SchemaVersion
WD:
    On Error GoTo FD
    T_W10P_PiiFlag
WDone:
    Exit Sub
FA:
    GroupFail "W10M 見出し->案件一覧行の写像(裁定書28)"
    Resume WB
FB:
    GroupFail "W10F ファイル名採番(13章§2.8)"
    Resume WC
FC:
    GroupFail "W10S schema_version の判定(13章§2.8 v2.7)"
    Resume WD
FD:
    GroupFail "W10P pii_flag の写像(統合W10・16章E-05(7))"
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

Private Sub ChkT(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check nm, cond, detail
End Sub

' ============================================================================
' W10M 見出し -> 案件一覧行の写像(13章§2.1 v2.7)
' ============================================================================
Private Sub T_W10M_HeaderToCaseRow()
    ' 企業ファイルの見出し(ReadFileHeader が返す形。「key=value」の行並び)。
    Dim fullText As String
    fullText = "company=あい商事" & vbLf & _
               "industry_code=09" & vbLf & _
               "updated_at=2026-09-03 10:00:00" & vbLf & _
               "schema_version=3.0.0" & vbLf & _
               "case_id=C-20260901-001" & vbLf & _
               "case_type=new" & vbLf & _
               "status=s2_done" & vbLf

    ' 期待値は 13章§2.1 の列名と CF3_REBUILD_COLS の並びから手で書き出した。
    ChkS "Test_W10M_01_全列そろいの写像_13章§2.1", _
        modCompanyFile3.HeaderToCaseRow(fullText), _
        "case_id" & vbTab & "C-20260901-001" & vbLf & _
        "company" & vbTab & "あい商事" & vbLf & _
        "industry_code" & vbTab & "09" & vbLf & _
        "case_type" & vbTab & "new" & vbLf & _
        "status" & vbTab & "s2_done" & vbLf & _
        "updated_at" & vbTab & "2026-09-03 10:00:00"

    ' 古い企業ファイル(業種とステータスの見出しを持たない)。欠けは空で出す。
    Dim thinText As String
    thinText = "company=あい商事" & vbLf & _
               "updated_at=2026-09-03 10:00:00" & vbLf & _
               "case_id=C-20260901-001" & vbLf & _
               "case_type=new" & vbLf

    ChkS "Test_W10M_02_欠け列は空値の行として出す_13章§2.8前方互換", _
        modCompanyFile3.HeaderToCaseRow(thinText), _
        "case_id" & vbTab & "C-20260901-001" & vbLf & _
        "company" & vbTab & "あい商事" & vbLf & _
        "industry_code" & vbTab & vbLf & _
        "case_type" & vbTab & "new" & vbLf & _
        "status" & vbTab & vbLf & _
        "updated_at" & vbTab & "2026-09-03 10:00:00"

    ' 新旧優先: 新しいほうを返す。同時刻は第1引数(本体側)を残す。
    Dim oldText As String, newText As String
    oldText = "updated_at=2026-09-03 10:00:00"
    newText = "updated_at=2026-09-03 10:00:01"
    ChkT "Test_W10M_03_新旧優先は新しいほうを返し同時刻は本体を残す_裁定書28", _
        (modCompanyFile3.PickNewerHeader(oldText, newText) = newText) And _
        (modCompanyFile3.PickNewerHeader(newText, oldText) = newText) And _
        (modCompanyFile3.PickNewerHeader(oldText, oldText) = oldText), _
        "新しい方・順序を入れ替えても同じ・同時刻は第1引数、の3条件"
End Sub

' ============================================================================
' W10F ファイル名採番(13章§2.8 手順1・4)
' ============================================================================
Private Sub T_W10F_CompanyFileName()
    ' 手順1の禁止文字と制御文字を全部入れた会社名。
    Dim rawName As String
    rawName = "A" & P20_BAD & "社" & Chr$(9) & "B"

    Dim nameText As String
    nameText = modCompanyFile3.CompanyFileNameOf(rawName)

    Dim leftOver As Boolean
    Dim i As Long
    For i = 1 To Len(P20_BAD)
        If InStr(1, nameText, Mid$(P20_BAD, i, 1), vbBinaryCompare) > 0 Then leftOver = True
    Next i

    ChkT "Test_W10F_01_禁止文字が名前に残らず末尾は企業カルテ_13章§2.8手順1", _
        (Not leftOver) And (Len(nameText) > Len(P20_TAIL)) And _
        (Right$(nameText, Len(P20_TAIL)) = P20_TAIL), _
        "実際=[" & nameText & "]"

    ' 手順4: 8桁は company 由来。同じ会社名は同じ名前(=1社1ファイルで追記できる)、
    ' 違う会社名は違う名前(=同名社2件の衝突を避ける)。
    ChkT "Test_W10F_02_同じ会社名は同じ名前で違う会社名は違う名前_13章§2.8手順4", _
        (modCompanyFile3.CompanyFileNameOf("あい商事") = _
         modCompanyFile3.CompanyFileNameOf("あい商事")) And _
        (modCompanyFile3.CompanyFileNameOf("あい商事") <> _
         modCompanyFile3.CompanyFileNameOf("うえ商事")) And _
        (LenB(modCompanyFile3.CompanyFileNameOf(vbNullString)) = 0), _
        "同一・相異・空の会社名は空文字、の3条件"
End Sub

' ============================================================================
' W10S schema_version の判定(13章§2.8 v2.7)
' ============================================================================
Private Sub T_W10S_SchemaVersion()
    ChkT "Test_W10S_01_版が空のファイルは1_0_0とみなして読める_13章§2.8v2.7", _
        (modCompanyFile3.SchemaVersionOf(vbNullString) = "1.0.0") And _
        (modCompanyFile3.SchemaVersionOf("  ") = "1.0.0") And _
        modCompanyFile3.IsSchemaReadable(vbNullString), _
        "空・空白は 1.0.0 で、かつ読める"

    ChkT "Test_W10S_02_majorが現行より新しい版は読まない_13章§2.8v2.7", _
        (Not modCompanyFile3.IsSchemaReadable("4.0.0")) And _
        modCompanyFile3.IsSchemaReadable("3.9.9") And _
        modCompanyFile3.IsSchemaReadable("2.9.9"), _
        "4.0.0は不可・3.9.9と2.9.9は可(現行は " & _
        modCompanyFile3.SchemaVersionCurrent() & ")"
End Sub

' ============================================================================
' W10P pii_flag の写像(統合W10の裁定・16章 E-05(7)・13章§2.8 v2.8)
'   自分の data_dir への自動保存は「共有・送信・配布」ではないので PII 検知が
'   あっても行う。代わりに dossier_case の pii_flag に印を残す。よって
'   (1) 列の並びに pii_flag があること (2) 値の写像が TRUE/FALSE であること
'   の2点が仕様の全部で、1本で見る。
' ============================================================================
Private Sub T_W10P_PiiFlag()
    Dim colsText As String
    colsText = modCompanyFile3.CaseSheetCols()

    ChkT "Test_W10P_01_dossier_caseはpii_flag列を持ち検知有無がTRUE_FALSEへ写る_16章E-05(7)", _
        (Right$(colsText, Len(";schema_version;pii_flag")) = ";schema_version;pii_flag") And _
        (InStr(1, colsText, "case_id;", vbBinaryCompare) = 1) And _
        (modCompanyFile3.PiiFlagOf("個人名 x 1 (S1)") = "TRUE") And _
        (modCompanyFile3.PiiFlagOf(vbNullString) = "FALSE") And _
        (modCompanyFile3.PiiFlagOf("  ") = "FALSE"), _
        "列の並び=[" & colsText & "]"
End Sub
