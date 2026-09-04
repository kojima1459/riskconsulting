Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================================
' modConfig - config シート(name / value / 説明)の読み書き
' ----------------------------------------------------------------------------
' 役割:
'   設定値は全て config シートに集約し、非エンジニアがVBAを開かずセルを
'   書き換えるだけでチューニングできるようにする。キー台帳(既定値表)の正は
'   13章§2.3であり、【製品固有の既定値をcoreに焼かない】(12章§4)。
'   起動手順②(12章§2.1)で modBoot が RegisterDefault で13章§2.3の既定値を
'   登録し、そのあと LoadFromSheet がシートの値で上書きする。
'
' 移植元: PoC「マイ本棚AI」 src/core/modConfig.bas。
'   GetString/GetLong/GetDouble/GetBool の「どんな入力でも既定値へ落ちる」
'   契約と ParseBoolText の3分岐(真トークン/偽トークン/読めない)、
'   NarrowAscii(全角英数の半角化)はそのまま維持した。改めた点は3つ:
'     ・GetString -> GetStr(14章§2のコード例の関数名に合わせる)
'     ・シート直読みから【メモリ上のキャッシュ】方式へ。既定値表の登録と
'       シート値での上書きという起動手順②の形をそのまま表現でき、
'       config シートが欠損しても既定値のまま起動できる(16章 E-52)。
'     ・シート不在時の MsgBox を廃止。core層はUIを持たない(R4)ため、
'       E0608 を err_log へ警告記録して続行し、案内はui層に任せる。
'
' R4: config シートの読み書きが責務そのもののため、Excelトークンの使用を
'   明示的に許可されたモジュール(12章§4の9本の1つ)。
' 循環参照について: modLog は log_max_rows を読むため modConfig を呼び返す。
'   GetXxx は【一切ログを出さない】ので入れ子は1段で止まる。ログを出すのは
'   LoadFromSheet / SetValue だけ、という切り分けを崩さないこと。
' ============================================================================

' 裁定書28: data_dir\設定.txt から上書きしてよいキー(表はここ1箇所)。
Public Const SETTINGS_ALLOWED_KEYS As String = _
    "portal_url;dr_url_menu;dr_url_quick;dr_url_full;dr_open_after_copy;" & _
    "ui_fullscreen;kb_path;data_dir"

Private Const CFG_SHEET As String = "config"
Private Const CFG_COL_NAME As Long = 1
Private Const CFG_COL_VALUE As Long = 2
Private Const CFG_FIRST_ROW As Long = 2       ' 1行目はヘッダ
' xlUp の数値。Excelの組み込み定数名を書かずに済ませ、LibreOffice側の
' 構文チェック(tools/run_lo_tests.py モード2)で未定義名にならないようにする。
Private Const CFG_DIR_UP As Long = -4162

' キー(小文字)-> 値(文字列)。既定値表の登録とシート値の上書きを1本で持つ。
Private gCfgCache As Collection
' LoadFromSheet が最後に読めたキー数(-1=まだ読んでいない)。
Private gCfgLoadedCount As Long

' ============================================================================
' RegisterDefault - 既定値を1件登録する(12章§2.1 手順②の前半)。
'   13章§2.3の既定値表は app/ui 側(modBoot)が持ち、core へは焼かない。
' ============================================================================
Public Sub RegisterDefault(ByVal cfgKey As String, ByVal defaultText As String)
    PutCache cfgKey, defaultText
End Sub

' ============================================================================
' LoadFromSheet - config シートの値でキャッシュを上書きする(手順②の後半)。
' ----------------------------------------------------------------------------
'   戻り値: 読めたキー数。シートが無い・読めないときは -1 を返し、E0608 を
'           err_log へ【警告として】記録して続行する(16章 E-52。config の
'           欠損で製品を止めない。既定値だけで起動できる)。
'   detail には失敗したキー名/事象だけを書き、値は書かない(NFR-S3)。
' ============================================================================
Public Function LoadFromSheet() As Long
    On Error GoTo Failed

    Dim ws As Object
    Set ws = ConfigSheet()
    If ws Is Nothing Then
        gCfgLoadedCount = -1
        LoadFromSheet = -1
        modLog.LogError "E0608", "modConfig.LoadFromSheet", "config_sheet_missing"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, CFG_COL_NAME).End(CFG_DIR_UP).row
    Dim readCount As Long
    readCount = 0
    Dim failedKeys As String
    failedKeys = ""

    Dim r As Long
    For r = CFG_FIRST_ROW To lastRow
        Dim k As String
        k = Trim$(CStr(ws.Cells(r, CFG_COL_NAME).Value))
        If LenB(k) > 0 Then
            Dim v As Variant
            v = ws.Cells(r, CFG_COL_VALUE).Value
            If IsError(v) Then
                failedKeys = modUtil.AppendIdList(failedKeys, k)
            Else
                PutCache k, VariantToText(v)
                readCount = readCount + 1
            End If
        End If
    Next r

    gCfgLoadedCount = readCount
    LoadFromSheet = readCount
    If LenB(failedKeys) > 0 Then
        modLog.LogError "E0608", "modConfig.LoadFromSheet", _
            "cell_error_keys:" & modUtil.SafeLeft(failedKeys, 300)
    End If
    Exit Function

Failed:
    gCfgLoadedCount = -1
    LoadFromSheet = -1
    modLog.LogError "E0608", "modConfig.LoadFromSheet", "read_failed", Err.Number
End Function

' 直近の LoadFromSheet で読めたキー数(-1=未読込・読込失敗)。
Public Function LoadedKeyCount() As Long
    LoadedKeyCount = gCfgLoadedCount
End Function

Public Function HasKey(ByVal cfgKey As String) As Boolean
    Dim v As String
    HasKey = TryGetCached(cfgKey, v)
End Function

' ============================================================================
' GetStr / GetLong / GetDouble / GetBool
' ----------------------------------------------------------------------------
'   「読めなければ既定値へ落ちる」という契約を、どんな入力でも守り切る。
'   非エンジニアが直接編集する設計なので、セルにエラー値・桁あふれ・全角の
'   TRUE が混ざることは現実に起きる。1キーの事故が以降の全キーの取得を
'   例外化しないよう、各関数が単独で握って既定値を返す。
'   ここでログは出さない(modLog が log_max_rows を読むため無限再帰になる)。
' ============================================================================
Public Function GetStr(ByVal cfgKey As String, ByVal defaultText As String) As String
    On Error GoTo Fallback
    Dim v As String
    If TryGetCached(cfgKey, v) Then
        GetStr = v
    Else
        GetStr = defaultText
    End If
    Exit Function
Fallback:
    GetStr = defaultText
End Function

Public Function GetLong(ByVal cfgKey As String, ByVal defaultValue As Long) As Long
    On Error GoTo Fallback
    Dim v As String
    If Not TryGetCached(cfgKey, v) Then
        GetLong = defaultValue
        Exit Function
    End If
    v = Trim$(NarrowAscii(v))
    If LenB(v) = 0 Or Not IsNumeric(v) Then
        GetLong = defaultValue
        Exit Function
    End If
    ' "99999999999" のような桁あふれは CLng がオーバーフローする。
    ' Long の範囲外は設定として意味を成さないので既定値へ落とす。
    Dim d As Double
    d = CDbl(v)
    If d > 2147483647# Or d < -2147483648# Then
        GetLong = defaultValue
    Else
        GetLong = CLng(d)
    End If
    Exit Function
Fallback:
    GetLong = defaultValue
End Function

Public Function GetDouble(ByVal cfgKey As String, ByVal defaultValue As Double) As Double
    On Error GoTo Fallback
    Dim v As String
    If Not TryGetCached(cfgKey, v) Then
        GetDouble = defaultValue
        Exit Function
    End If
    v = Trim$(NarrowAscii(v))
    If LenB(v) = 0 Or Not IsNumeric(v) Then
        GetDouble = defaultValue
    Else
        GetDouble = CDbl(v)
    End If
    Exit Function
Fallback:
    GetDouble = defaultValue
End Function

Public Function GetBool(ByVal cfgKey As String, ByVal defaultValue As Boolean) As Boolean
    On Error GoTo Fallback
    Dim v As String
    If Not TryGetCached(cfgKey, v) Then
        GetBool = defaultValue
    Else
        GetBool = ParseBoolText(v, defaultValue)
    End If
    Exit Function
Fallback:
    GetBool = defaultValue
End Function

' ============================================================================
' ParseBoolText - config の値(文字列)を真偽へ解釈する純関数。
' ----------------------------------------------------------------------------
'   真トークン("true"/"1"/"yes"/"on")   -> True
'   偽トークン("false"/"0"/"no"/"off")  -> False   (明示的なFALSEは尊重する)
'   それ以外(空文字・空白のみ・誤記)   -> defaultValue(読めなかった)
'   偽トークンを独立させているのは、「FALSEと書いた」意図を既定TRUEのキーで
'   握り潰さないため。2分岐にすると keep_window_alive=FALSE の手入力が
'   効かなくなる(PoCで実際に起きた事故の再発防止)。
' ============================================================================
Public Function ParseBoolText(ByVal raw As String, ByVal defaultValue As Boolean) As Boolean
    Dim s As String
    s = LCase$(Trim$(NarrowAscii(raw)))
    Select Case s
        Case "true", "1", "yes", "on"
            ParseBoolText = True
        Case "false", "0", "no", "off"
            ParseBoolText = False
        Case Else
            ParseBoolText = defaultValue
    End Select
End Function

' ============================================================================
' SetValue - キャッシュと config シートの両方を更新する。
'   戻り値: シートまで書けたら True(キャッシュだけ更新できた場合は False)。
'   セルへの書込は modUtilText.SetCellSafe を通す(16章 NFR-S7①)。
' ============================================================================
Public Function SetValue(ByVal cfgKey As String, ByVal valueText As String) As Boolean
    On Error GoTo Failed
    PutCache cfgKey, valueText

    Dim ws As Object
    Set ws = ConfigSheet()
    If ws Is Nothing Then Exit Function

    Dim r As Long
    r = FindKeyRow(ws, cfgKey)
    If r = 0 Then
        r = ws.Cells(ws.Rows.count, CFG_COL_NAME).End(CFG_DIR_UP).row + 1
        modUtilText.SetCellSafe ws.Cells(r, CFG_COL_NAME), cfgKey, "config/name"
    End If
    modUtilText.SetCellSafe ws.Cells(r, CFG_COL_VALUE), valueText, "config/value"
    SetValue = True
    Exit Function
Failed:
    SetValue = False
End Function

' ============================================================================
' ParseSettingsText - data_dir\設定.txt の解析(純関数。裁定書28)
' ----------------------------------------------------------------------------
' 書式(利用者が手で書く前提の最小形):
'   ・1行1件の `key=value`(UTF-8)。key と value の前後の空白は落とす
'   ・`#` で始まる行はコメント。空行は無視
'   ・**許可キー以外は黙って捨てる**(設定ファイルから任意の config キーを
'     注入させない。log_max_rows や llm_transport を書き換えられないこと)
'   ・同じキーが2度出たら**後の行が勝つ**(手で足した行が効く)
' 返り値は `key=value` を vbLf でつないだもの(呼出側は Split して SetValue)。
' 値の中の改行は書式上あり得ない(1行1件)。値の中の "=" は残す(URLのため)。
' ============================================================================
Public Function ParseSettingsText(ByVal bodyText As String) As String
    If LenB(bodyText) = 0 Then Exit Function

    Dim t As String
    t = Replace(Replace(bodyText, vbCrLf, vbLf), vbCr, vbLf)

    Dim rows() As String
    rows = Split(t, vbLf)

    Dim outText As String
    Dim lineText As String
    Dim keyText As String
    Dim valueText As String
    Dim pos As Long
    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        lineText = Trim$(rows(i))
        If LenB(lineText) > 0 Then
            If Left$(lineText, 1) <> "#" Then
                pos = InStr(1, lineText, "=", vbBinaryCompare)
                If pos > 1 Then
                    keyText = LCase$(Trim$(Left$(lineText, pos - 1)))
                    valueText = Trim$(Mid$(lineText, pos + 1))
                    If IsSettingsKeyAllowed(keyText) Then
                        If LenB(outText) > 0 Then outText = outText & vbLf
                        outText = outText & keyText & "=" & valueText
                    End If
                End If
            End If
        End If
    Next i
    ParseSettingsText = outText
End Function

' 許可キーか(純関数)。表は SETTINGS_ALLOWED_KEYS 1箇所だけが持つ。
Public Function IsSettingsKeyAllowed(ByVal cfgKey As String) As Boolean
    If LenB(cfgKey) = 0 Then Exit Function
    IsSettingsKeyAllowed = (InStr(1, ";" & SETTINGS_ALLOWED_KEYS & ";", _
                                  ";" & LCase$(Trim$(cfgKey)) & ";", vbBinaryCompare) > 0)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------
Private Function ConfigSheet() As Object
    On Error GoTo NoSheet
    Set ConfigSheet = ThisWorkbook.Worksheets(CFG_SHEET)
    Exit Function
NoSheet:
    Set ConfigSheet = Nothing
End Function

Private Function FindKeyRow(ByVal ws As Object, ByVal cfgKey As String) As Long
    On Error GoTo NotFound
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, CFG_COL_NAME).End(CFG_DIR_UP).row
    Dim i As Long
    For i = CFG_FIRST_ROW To lastRow
        If StrComp(CStr(ws.Cells(i, CFG_COL_NAME).Value), cfgKey, vbTextCompare) = 0 Then
            FindKeyRow = i
            Exit Function
        End If
    Next i
    Exit Function
NotFound:
    FindKeyRow = 0
End Function

Private Sub PutCache(ByVal cfgKey As String, ByVal valueText As String)
    Dim k As String
    k = LCase$(Trim$(cfgKey))
    If LenB(k) = 0 Then Exit Sub
    If gCfgCache Is Nothing Then Set gCfgCache = New Collection

    On Error Resume Next
    gCfgCache.Remove k
    On Error GoTo 0
    gCfgCache.Add valueText, k
End Sub

Private Function TryGetCached(ByVal cfgKey As String, ByRef outValue As String) As Boolean
    outValue = ""
    If gCfgCache Is Nothing Then Exit Function
    On Error GoTo NotFound
    outValue = CStr(gCfgCache.Item(LCase$(Trim$(cfgKey))))
    TryGetCached = True
    Exit Function
NotFound:
    outValue = ""
    TryGetCached = False
End Function

' セルの値(数値・真偽・日付・文字列)を config の文字列表現へ均す。
' 日付は modUtilText の Iso 系を通す(和暦端末で年が元号年になるのを防ぐ)。
Private Function VariantToText(ByVal v As Variant) As String
    On Error GoTo AsText
    Select Case VarType(v)
        Case vbEmpty, vbNull
            VariantToText = ""
        Case vbBoolean
            VariantToText = IIf(CBool(v), "TRUE", "FALSE")
        Case vbDate
            VariantToText = modUtilText.IsoDateTime(CDate(v))
        Case Else
            VariantToText = CStr(v)
    End Select
    Exit Function
AsText:
    VariantToText = ""
End Function

' ----------------------------------------------------------------------------
' NarrowAscii - 全角の英数字と全角スペースだけを半角へ均す(純関数)。
'   StrConv(s, vbNarrow) は日本語ロケールでしか意図どおり動かない(他ロケール
'   では変換されない/半角カナまで巻き込む)ので使わない。判定に要るのは
'   トークンを構成する英数字だけなので Replace の小さな表で足りる。
'   全角スペース(U+3000)を半角へ均すのは、Trim$ が U+3000 を落とさず
'   「全角スペースだけのセル」が解釈不能扱いから漏れるのを防ぐため。
'   非ASCIIは必ず ChrW で組む(CP932へ潰れる文字をソースに直接置かない)。
' ----------------------------------------------------------------------------
Private Function NarrowAscii(ByVal s As String) As String
    Dim t As String
    t = Replace(s, ChrW(&H3000&), " ")

    Dim i As Long
    For i = 0 To 9
        t = Replace(t, ChrW(&HFF10& + i), Chr$(48 + i))
    Next i
    For i = 0 To 25
        t = Replace(t, ChrW(&HFF21& + i), Chr$(65 + i))
        t = Replace(t, ChrW(&HFF41& + i), Chr$(97 + i))
    Next i

    NarrowAscii = t
End Function
