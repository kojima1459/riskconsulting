Attribute VB_Name = "modBootData"
Option Explicit

' ============================================================================
' modBootData - 起動時の案件一覧の再構成(ui層・T-59・裁定書28 W10)
' ----------------------------------------------------------------------------
' なぜ modBoot の外に置くのか(12章§2):
'   modBoot は 36,172 字で30,000字契約を既に超えており、起動手順を1本足す
'   ぶんの実体を入れられない。そこで**実体は本モジュール**が持ち、modBoot は
'   手順(3)から1行呼ぶだけにした(起動処理の順序の正は従来どおり modBoot)。
'
' 何をするか(裁定書28 の「起動時再構成」):
'   本体xlsm は D: に置かれ**毎朝消えてよい**エンジンである。蓄積の正は
'   OneDrive の `データ\企業\<会社名>_<8桁>_企業カルテ.xlsx`(1社1ファイル)で
'   あり、本体の案件一覧は**作業用キャッシュ**でしかない(13章§2.1)。
'   起動のたびに企業フォルダを走査し、各ファイルの**見出しだけ**
'   (会社名・業種・案件種別・最終更新・ステータス・case_id)から案件一覧を
'   組み直す。case_data 等の本体は読まない(遅延読込=案件を選んで
'   [企業ファイルを開く]を押したときに modCompanyFile.ImportCompanyFile が読む)。
'
' 同じ case_id が本体にも企業ファイルにもあるとき:
'   **企業ファイルを正**とし、更新時刻の新しいほうで上書きする。ただし本体側の
'   ほうが新しい(=今日の作業がまだ書き出されていない)ときは触らない。
'
' 走査は Dir$ ループで行う(FSO は使わない=12章§4 の禁止)。**ファイル名を先に
'   全部集めてから**1件ずつ開く: Dir$ は列挙の状態をプロセスで1つしか持たず、
'   ループの中で別の Dir$(modCompanyFile2 の存在確認など)を呼ぶと列挙が
'   その場で壊れて残りのファイルを取りこぼすため。
'
' 失敗しても起動を止めない(12章§2.1 の起動手順の作法)。1件のファイルが壊れて
'   いても残りは読む。
' ============================================================================

Private Const BD_SRC As String = "modBootData"
Private Const BD_PATTERN As String = "*.xlsx"

' 1回の起動で読む企業ファイルの上限。担当15社(FR-45)に対して十分に広く、
' 事故(共有フォルダを丸ごと指した等)で起動が終わらなくなるのを止める幅。
Private Const BD_MAX_FILES As Long = 500

' ============================================================================
' RebuildCaseCache - 企業フォルダを走査して案件一覧を組み直す。
'   戻り値 = 案件一覧へ書いた(上書き・追加した)件数。
' ============================================================================
Public Function RebuildCaseCache() As Long
    On Error GoTo Failed

    Dim dirText As String
    dirText = modCompanyFile3.CompanyDir()
    If LenB(dirText) = 0 Then Exit Function

    Dim listText As String
    listText = FileNamesIn(dirText)
    If LenB(listText) = 0 Then Exit Function

    Dim names() As String
    names = Split(listText, vbLf)

    Dim n As Long
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If LenB(names(i)) > 0 Then
            If ApplyOneFile(dirText & "\" & names(i)) Then n = n + 1
        End If
    Next i

    modLog.LogUsage "case_cache_rebuilt", vbNullString, _
                    "files=" & CStr(UBound(names) - LBound(names) + 1) & _
                    " applied=" & CStr(n)
    RebuildCaseCache = n
    Exit Function

Failed:
    modLog.LogError "E0603", BD_SRC & ".RebuildCaseCache", "rebuild_failed", Err.Number
    RebuildCaseCache = 0
End Function

' 企業フォルダの .xlsx を vbLf 区切りで列挙する(Dir$ ループ。FSO禁止)。
'   **この関数の中では他の Dir$ を1つも呼ばない**(列挙状態が壊れるため)。
Private Function FileNamesIn(ByVal dirText As String) As String
    On Error GoTo Done0

    Dim acc As String
    Dim n As Long
    Dim nameText As String
    nameText = Dir$(dirText & "\" & BD_PATTERN)
    Do While LenB(nameText) > 0
        ' Excel が開いているファイルの一時名(~$...)は企業ファイルではない。
        If Left$(nameText, 2) <> "~$" Then
            If LenB(acc) > 0 Then acc = acc & vbLf
            acc = acc & nameText
            n = n + 1
            If n >= BD_MAX_FILES Then Exit Do
        End If
        nameText = Dir$
    Loop

Done0:
    FileNamesIn = acc
End Function

' 1ファイルぶん。見出しを読み、案件一覧より新しければ書く。書いたら True。
Private Function ApplyOneFile(ByVal pathText As String) As Boolean
    On Error GoTo Failed

    Dim headerText As String
    headerText = modCompanyFile3.ReadFileHeader(pathText)
    If LenB(headerText) = 0 Then Exit Function

    Dim caseId As String
    caseId = modCompanyFile3.HeaderValueOf(headerText, "case_id")
    If LenB(caseId) = 0 Then Exit Function

    ' 本体の案件一覧が既に新しいなら触らない(今日の作業を古い写しで潰さない)。
    If Not modCompanyFile3.IsFileNewer(headerText, caseId) Then Exit Function

    Dim pairsText As String
    pairsText = modCompanyFile3.HeaderToCaseRow(headerText)
    If LenB(pairsText) = 0 Then Exit Function

    ApplyOneFile = modCaseStore3.UpsertCaseRow(caseId, pairsText)
    Exit Function

Failed:
    modLog.LogError "E0603", BD_SRC & ".ApplyOneFile", "apply_failed", Err.Number
    ApplyOneFile = False
End Function
