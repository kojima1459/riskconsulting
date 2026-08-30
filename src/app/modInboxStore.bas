Attribute VB_Name = "modInboxStore"
Option Explicit

' ============================================================================
' modInboxStore - 受信箱シートの唯一の入出力口(app層・T-25)
' ----------------------------------------------------------------------------
' 正: 13章§1(採番 I-YYYYMM-NNN)/13章§2.6(受信箱の列)/11章§4(受信箱の状態
'   遷移)/16章 E-40・E-41(統制語彙の必須化)/10章 FR-17(関心度=同一テーマの
'   件数集計)/14章§6(NewInboxItem・SetInboxJudgement の宣言)/19章§3(enum)。
' 責務: 投函の起票・診断結果の格納・判定(統制語彙)の記録・未診断IDの列挙・
'   関心度の集計。**受信箱1枚しか触らない**。
' R4(12章§4): Excelトークン許可モジュールの1つ。許可の幅は「本体ブックの
'   受信箱1枚」。セル書込は全て modUtilText.SetCellSafe(16章 NFR-S7①)を通す
'   (body の32,000字打切りと先頭式記号の無害化もそこが唯一の実装であり、
'   13章§2.6「1セル上限で打切り+警告」はこの1本で満たす)。
' 純ロジックの分離: 採番・ID書式・状態遷移・E-41の必須検査・関心度キー・
'   関心度の集計(FR-17)はシートI/Oを含まない純関数として公開する(14章§6)。
'   規約をシートI/Oの中へ閉じ込めると層(a)から誰も検査できない(W2aの実害と
'   同じ轍を踏まない)。
' 投函下書き行(13章§2.6・裁定書10 C1): 先頭データ行は[＋投函]の入力欄であって
'   受信箱のデータではない。inbox_id 列が IB_DRAFT_MARK の行は走査・判定・集計の
'   すべてから除外する(IsDraftId が唯一の判定。行番号には依存しない)。
' 列アクセスは列名ベース(13章冒頭。列番号のハードコード禁止)。例外は投げず、
'   読めない・書けないは False / "" で返す。
' ============================================================================

' --- シート名(13章§2.6) ---
Private Const IB_SHEET As String = "受信箱"

' xlUp の数値(組込定数名を書かず LO の構文チェックで未定義名にしない)。
Private Const IB_DIR_UP As Long = -4162

' 見出し行を読む幅(受信箱16列に余裕を見た探索範囲。列番号ではない)。
Private Const IB_SCAN_COLS As Long = 24

' 同月連番の上限(案件の E-23 と同じ3桁採番)。
Private Const IB_SERIAL_MAX As Long = 999

' enum の区切り(modUtil.AppendIdList と同じ ";")。
Private Const IB_SEP As String = ";"

' 13章§2.6 / 19章§3 の enum。
Private Const IB_SOURCE_KINDS As String = "member_post;field_voice;watch"
Private Const IB_STATUSES As String = "undiagnosed;diagnosed;adopted;conditional_hold;rejected;merged"
Private Const IB_JUDGED As String = "adopted;conditional_hold;rejected;merged"
Private Const IB_DROP_TYPES As String = "T0;T1;T2;T3;T4;T5;T6;T7;T8;T9;T10"
Private Const IB_REVIVE_TAGS As String = "tech;regulation;partner;data;market"

' 投函下書き行の固定マーカー(13章§2.6・裁定書10 C1)。受信箱テーブルの先頭
' データ行(ヘッダ直下)は[＋投函]の入力欄であって受信箱のデータではない。
' inbox_id 列がこの値の行は**走査・判定・集計のどれからも必ず除外する**
' (未診断の抽出・関心度集計・ID一致検索。判定は ID 一致検索を通るので同時に閉じる)。
' 判定基準を1箇所に閉じるため、この定数と IsDraftId 以外に下書き行の知識を置かない。
Public Const IB_DRAFT_MARK As String = "(下書き)"

' 起票直後の状態(11章§4)。
Private Const IB_STATUS_NEW As String = "undiagnosed"
Private Const IB_STATUS_DIAGNOSED As String = "diagnosed"

' 見直し期日が未入力であることを表す値(VBAのDateは数値。0=1899-12-30)。
Private Const IB_NO_DATE As Double = 0

' 関心度の既定表示件数(11章 受信箱ワイヤーの「関心度: 熊対策12件 雹災5件」)。
Private Const IB_INTEREST_TOP As Long = 3
Private Const IB_INTEREST_SEP As String = " "
Private Const IB_INTEREST_UNIT As String = "件"

' --- 純ロジック(Excel非依存。シート操作側は必ずここを通す。14章§6) ---

' BuildInboxId - 受信箱IDの組立(13章§1: I-YYYYMM-NNN)。monthText は yyyymm の
'   6桁ちょうど(数字のみ)、seq は 1..999。範囲外は ""(呼び出し側が枯渇・不正
'   として扱う。例外は投げない)。第1引数名が monthText なのは Month が VBA の
'   組込関数名で、vba_lint の予約語検査が識別子として拒否するため。
Public Function BuildInboxId(ByVal monthText As String, ByVal seq As Long) As String
    If Len(monthText) <> 6 Then Exit Function
    If Not IsAllDigits(monthText) Then Exit Function
    If seq < 1 Or seq > IB_SERIAL_MAX Then Exit Function
    BuildInboxId = "I-" & monthText & "-" & Right$("00" & CStr(seq), 3)
End Function

' IsValidInboxId - "I-"+数字6桁+"-"+数字3桁(001..999)ちょうどか。前後空白は許さない。
Public Function IsValidInboxId(ByVal id As String) As Boolean
    If Len(id) <> 12 Then Exit Function
    If Left$(id, 2) <> "I-" Then Exit Function
    If Mid$(id, 9, 1) <> "-" Then Exit Function
    If Not IsAllDigits(Mid$(id, 3, 6)) Then Exit Function
    Dim tailText As String
    tailText = Right$(id, 3)
    If Not IsAllDigits(tailText) Then Exit Function
    IsValidInboxId = (CLng(tailText) >= 1)
End Function

' CanInboxTransition - 受信箱statusの遷移可否(11章§4)。許すのは次の2種だけ:
'   (1) undiagnosed -> diagnosed(プリフライト診断の完了)
'   (2) diagnosed -> adopted / conditional_hold / rejected / merged(判定)
'   自己遷移・判定済みからの再判定・診断を飛ばした判定は False。
'   enum(13章§2.6)に無い値はどちらの側でも False。
Public Function CanInboxTransition(ByVal fromStatus As String, ByVal toStatus As String) As Boolean
    Dim f As String, tgt As String
    f = Trim$(fromStatus)
    tgt = Trim$(toStatus)
    If Not IsListedValue(IB_STATUSES, f) Then Exit Function
    If Not IsListedValue(IB_STATUSES, tgt) Then Exit Function
    If StrComp(f, tgt, vbBinaryCompare) = 0 Then Exit Function
    If StrComp(f, IB_STATUS_NEW, vbBinaryCompare) = 0 Then
        CanInboxTransition = (StrComp(tgt, IB_STATUS_DIAGNOSED, vbBinaryCompare) = 0)
        Exit Function
    End If
    If StrComp(f, IB_STATUS_DIAGNOSED, vbBinaryCompare) = 0 Then
        CanInboxTransition = IsListedValue(IB_JUDGED, tgt)
    End If
End Function

' JudgementError - 16章 E-41(統制語彙の必須化)の唯一の判定。判定を保存して
'   よければ ""、保存を止めるなら理由の1行を返す(呼び出し側は ""以外なら
'   1列も書かない=保存ブロック)。
'   rejected         : drop_type(T0からT10)必須
'   conditional_hold : revive_tag(5値)と見直し期日の両方が必須
'   hasDue           : 見直し期日が入力されているか(Date型を純関数へ持ち込まない
'                      ため呼び出し側が真偽へ落として渡す)
'   status 自体も統制語彙(13章§2.6の6値)である。表に無い値は条件付き列の
'   検査へ進む前に止める(E-41 は「統制語彙の必須化」であって、判定列の値そのもの
'   が語彙外なら保存してよい理由が無い)。SetInboxJudgement も別途 IB_JUDGED で
'   絞るが、E-41 の唯一の判定点である本関数が語彙外を素通しすると、シートI/Oを
'   経ない呼び出し(層(a)のテスト・将来の別呼び出し口)で fail-open になる。
Public Function JudgementError(ByVal statusText As String, ByVal dropType As String, _
                               ByVal reviveTag As String, ByVal hasDue As Boolean) As String
    Dim tgt As String
    tgt = Trim$(statusText)

    If Not IsListedValue(IB_STATUSES, tgt) Then
        JudgementError = "受信箱の状態が統制語彙にありません: " & tgt
        Exit Function
    End If

    If StrComp(tgt, "rejected", vbBinaryCompare) = 0 Then
        If Not IsListedValue(IB_DROP_TYPES, dropType) Then
            JudgementError = "却下にはドロップ類型(T0からT10)の選択が必要です"
        End If
        Exit Function
    End If

    If StrComp(tgt, "conditional_hold", vbBinaryCompare) = 0 Then
        If Not IsListedValue(IB_REVIVE_TAGS, reviveTag) Then
            JudgementError = "条件付き保留には復活条件タグの選択が必要です"
            Exit Function
        End If
        If Not hasDue Then JudgementError = "条件付き保留には見直し期日の入力が必要です"
    End If
End Function

' InterestKeyOf - 関心度集計のテーマキー(10章 FR-17)。同一テーマの重複投函を
'   棄却せず件数へ寄せるための正規化で、改行・空白の揺れと大小文字だけを吸収
'   する(意味の同一視まではしない=人が読める粒度で数える)。空テーマは ""。
Public Function InterestKeyOf(ByVal themeText As String) As String
    Dim t As String
    ' 14章§6: 改行・空白の揺れと大小文字だけを吸収する。改行は空白へ畳む
    t = Replace(Replace(themeText, vbCrLf, " "), vbLf, " ")
    t = Replace(t, vbCr, " ")
    t = Trim$(modUtilText.NormalizeForHash(t))
    If LenB(t) = 0 Then Exit Function
    InterestKeyOf = LCase$(t)
End Function

' FmtInterestLine - 関心度1件の表示(11章 受信箱ワイヤー「熊対策12件」)。
'   件数0以下・テーマ空は ""(表示しない)。
Public Function FmtInterestLine(ByVal themeText As String, ByVal itemCount As Long) As String
    If itemCount <= 0 Then Exit Function
    If LenB(Trim$(themeText)) = 0 Then Exit Function
    FmtInterestLine = Trim$(themeText) & CStr(itemCount) & IB_INTEREST_UNIT
End Function

' InterestSummaryOf - 関心度の集計そのもの(10章 FR-17・11章 受信箱ワイヤー
'   「関心度: 熊対策12件 雹災5件」)。themeLines は**投函1件につき1行**
'   (vbLf区切り)のテーマ文字列。
'   規約(この1本にしか無い): (1)同一テーマの重複投函を棄却せず件数へ寄せる
'   (2)件数の多い順に並べる (3)**2件以上集まったテーマだけ**を載せる(1件は
'   「重複投函」ではない) (4)上限は IB_INTEREST_TOP 件 (5)区切りは半角空白。
'   同一視の粒度は InterestKeyOf、1件ぶんの表示は FmtInterestLine が唯一の
'   値源(答えを2箇所に書かない)。空・全行が空テーマなら ""(空の集計を
'   でっち上げない)。シートI/Oの InterestText は行を集めてここへ渡すだけで、
'   上の5つの規約を1つも持たない(裁定書9-1: 規約を層(a)から検査可能にする)。
Public Function InterestSummaryOf(ByVal themeLines As String) As String
    If LenB(Trim$(themeLines)) = 0 Then Exit Function

    Dim lineTexts() As String
    lineTexts = Split(themeLines, vbLf)

    Dim maxKinds As Long
    maxKinds = UBound(lineTexts) - LBound(lineTexts) + 1
    If maxKinds < 1 Then Exit Function

    Dim keys() As String
    Dim labels() As String
    Dim counts() As Long
    ReDim keys(1 To maxKinds)
    ReDim labels(1 To maxKinds)
    ReDim counts(1 To maxKinds)

    Dim kindCount As Long
    Dim i As Long
    For i = LBound(lineTexts) To UBound(lineTexts)
        Dim oneTheme As String
        oneTheme = Trim$(Replace(lineTexts(i), vbCr, vbNullString))
        Dim k As String
        k = InterestKeyOf(oneTheme)
        If LenB(k) > 0 Then
            Dim hit As Long
            hit = IndexOfKey(keys, kindCount, k)
            If hit = 0 Then
                kindCount = kindCount + 1
                keys(kindCount) = k
                labels(kindCount) = oneTheme
                counts(kindCount) = 1
            Else
                counts(hit) = counts(hit) + 1
            End If
        End If
    Next i

    InterestSummaryOf = TopInterest(labels, counts, kindCount, IB_INTEREST_TOP)
End Function

' 半角数字だけで構成されているか(空は False)。
Private Function IsAllDigits(ByVal s As String) As Boolean
    If LenB(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        If InStr(1, "0123456789", Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsAllDigits = True
End Function

' ";" 区切りの enum リストに値が含まれるか(前後空白は無視・大小文字は区別)。
Private Function IsListedValue(ByVal listText As String, ByVal valueText As String) As Boolean
    Dim v As String
    v = Trim$(valueText)
    If LenB(v) = 0 Then Exit Function
    IsListedValue = (InStr(1, IB_SEP & listText & IB_SEP, IB_SEP & v & IB_SEP, vbBinaryCompare) > 0)
End Function

' --- 受信箱シートI/O ---

' NewInboxItem - 投函を1件起票して inbox_id を返す(13章§1・§2.6・14章§6)。
'   採番は I-YYYYMM-NNN。当月の使用済み最大連番の次から始め、衝突したら E0605
'   を記録して次の番号へ(999で枯渇)。失敗時は ""(例外は投げない)。
'   posted_by_group は§6のシグネチャが受け取らないため空のまま起こす(13章
'   §2.6の必須は実行前検証の時点で満たされていればよい。modCaseStore.NewCase
'   の channel 等と同じ扱い)。body の32,000字打切りは SetCellSafe が行う。
Public Function NewInboxItem(ByVal sourceKind As String, ByVal theme As String, _
                             ByVal body As String) As String
    On Error GoTo Failed

    If Not IsListedValue(IB_SOURCE_KINDS, sourceKind) Then
        modLog.LogError "E0101", "modInboxStore.NewInboxItem", "invalid_source_kind"
        Exit Function
    End If
    If LenB(Trim$(theme)) = 0 Then
        modLog.LogError "E0101", "modInboxStore.NewInboxItem", "empty_required:theme"
        Exit Function
    End If

    Dim ws As Object
    Set ws = IbSheet()
    If ws Is Nothing Then
        modLog.LogError "E0603", "modInboxStore.NewInboxItem", "sheet_missing:inbox"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = IbLastRow(ws)
    Dim blk As Variant
    blk = IbBlock(ws, lastRow)
    Dim cId As Long
    cId = modUtil.FindHeaderCol(blk, "inbox_id")
    If cId <= 0 Then
        modLog.LogError "E0603", "modInboxStore.NewInboxItem", "header_missing:inbox_id"
        Exit Function
    End If

    Dim monthText As String
    monthText = Left$(modUtilText.IsoDateCompact(Date), 6)

    ' 当月ぶんの使用済み最大連番(裁定書10 C1: 下書き行のマーカーは ID の形を
    ' 満たさないので SerialOfInboxId が 0 を返し、採番に影響しない)。
    Dim usedMax As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim sn As Long
        sn = SerialOfInboxId(Trim$(CStr(blk(r, cId))), monthText)
        If sn > usedMax Then usedMax = sn
    Next r

    Dim serialNo As Long
    Dim newId As String
    serialNo = usedMax
    Do
        serialNo = serialNo + 1
        If serialNo > IB_SERIAL_MAX Then
            modLog.LogError "E0605", "modInboxStore.NewInboxItem", "serial_exhausted:" & monthText
            Exit Function
        End If
        newId = BuildInboxId(monthText, serialNo)
        If LenB(newId) = 0 Then
            modLog.LogError "E0605", "modInboxStore.NewInboxItem", "id_build_failed:" & monthText
            Exit Function
        End If
        If IbRowOf(blk, lastRow, cId, newId) <= 0 Then Exit Do
        modLog.LogError "E0605", "modInboxStore.NewInboxItem", "id_collision:" & CStr(serialNo)
    Loop

    Dim wr As Long
    wr = lastRow + 1
    IbPut ws, blk, wr, "inbox_id", newId
    IbPut ws, blk, wr, "posted_at", modUtil.NowStamp()
    IbPut ws, blk, wr, "source_kind", Trim$(sourceKind)
    IbPut ws, blk, wr, "theme", theme
    IbPut ws, blk, wr, "body", body
    IbPut ws, blk, wr, "status", IB_STATUS_NEW

    NewInboxItem = newId
    Exit Function

Failed:
    modLog.LogError "E0603", "modInboxStore.NewInboxItem", "unexpected", Err.Number
    NewInboxItem = vbNullString
End Function

' ReadInboxItem - 1件の投函を読む(modPlayOps がプリフライト診断の入力に使う
'   唯一の口。modPlayOps は R4 でシートに触れない)。
'   戻り値 False=シート・見出し・当該行が無い(呼び出し側は fail-closed で中止)。
Public Function ReadInboxItem(ByVal inboxId As String, ByRef theme As String, _
                              ByRef body As String, ByRef statusText As String) As Boolean
    On Error GoTo Failed

    theme = vbNullString
    body = vbNullString
    statusText = vbNullString
    If Not IsValidInboxId(inboxId) Then Exit Function

    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not IbLocate(inboxId, ws, blk, rowNo) Then Exit Function

    theme = IbValue(blk, rowNo, "theme")
    body = IbValue(blk, rowNo, "body")
    statusText = IbValue(blk, rowNo, "status")
    ReadInboxItem = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modInboxStore.ReadInboxItem", "read_failed", Err.Number
    ReadInboxItem = False
End Function

' UndiagnosedIds - 未診断(undiagnosed)の inbox_id を投函順に ";" 区切りで返す。
'   一括診断(16章 E-40)の対象一覧。1件も無ければ ""。
Public Function UndiagnosedIds() As String
    On Error GoTo Failed

    Dim ws As Object
    Set ws = IbSheet()
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = IbLastRow(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = IbBlock(ws, lastRow)

    Dim acc As String
    Dim r As Long
    For r = 2 To lastRow
        ' 裁定書10 C1: 投函下書き行は受信箱のデータではないので診断対象にしない。
        If Not IsDraftId(IbValue(blk, r, "inbox_id")) Then
            If StrComp(IbValue(blk, r, "status"), IB_STATUS_NEW, vbBinaryCompare) = 0 Then
                acc = modUtil.AppendIdList(acc, IbValue(blk, r, "inbox_id"))
            End If
        End If
    Next r
    UndiagnosedIds = acc
    Exit Function

Failed:
    modLog.LogError "E0603", "modInboxStore.UndiagnosedIds", "read_failed", Err.Number
    UndiagnosedIds = vbNullString
End Function

' SavePfResult - プリフライト診断の結果を格納する(13章§2.6 pf_json /
'   pf_survival / pf_pred_types)。あわせて status を undiagnosed -> diagnosed
'   へ進める(遷移の可否は CanInboxTransition が唯一の判定)。
'   既に diagnosed 以降の行は診断結果だけを上書きし status は動かさない
'   (判定済みの行を undiagnosed へ巻き戻さない)。
Public Function SavePfResult(ByVal inboxId As String, ByVal pfJson As String, _
                             ByVal survival As String, ByVal predTypes As String) As Boolean
    On Error GoTo Failed

    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not IbLocate(inboxId, ws, blk, rowNo) Then Exit Function

    IbPut ws, blk, rowNo, "pf_json", pfJson
    IbPut ws, blk, rowNo, "pf_survival", survival
    IbPut ws, blk, rowNo, "pf_pred_types", predTypes

    If CanInboxTransition(IbValue(blk, rowNo, "status"), IB_STATUS_DIAGNOSED) Then
        IbPut ws, blk, rowNo, "status", IB_STATUS_DIAGNOSED
    End If
    SavePfResult = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modInboxStore.SavePfResult", "write_failed", Err.Number
    SavePfResult = False
End Function

' SetInboxJudgement - 判定(統制語彙)を記録する(14章§6・16章 E-41)。
'   status は adopted / conditional_hold / rejected のみ受け付ける。
'   **merged は受け付けない**: 13章§2.6 は merged に merged_into(統合先の
'   inbox_id)を必須としているが、14章§6 のシグネチャは統合先を受け取る引数を
'   持たない。空の merged_into を書くと13章の必須を満たさない行ができるため
'   fail-closed で拒否し、E0101 を記録する(欠けている引数の追加は§6の裁定事項)。
'   E-41 の必須検査(JudgementError)に掛かったら**1列も書かない**(保存ブロック)。
'   遷移そのものの可否は CanInboxTransition が唯一の判定点。
'   判定の入力口(裁定書9 B2・N5): ui層は status 列ではなく入力列 judge_to
'   (13章§2.6)を読み、その値を statusText へ渡す。from(現在の status)と to が
'   同一セル由来になって自己遷移で必ず False になる経路は構造として無い。
'   成功時は status を書き換えたうえで **judge_to を空へ戻す**(判定済みの行に
'   入力値を残さない。judge_to 列が無いブックでは IbPut が何もしない)。
Public Function SetInboxJudgement(ByVal inboxId As String, ByVal statusText As String, _
                                  ByVal dropType As String, ByVal reviveTag As String, _
                                  ByVal reviveDue As Date) As Boolean
    On Error GoTo Failed

    Dim tgt As String
    tgt = Trim$(statusText)
    If StrComp(tgt, "merged", vbBinaryCompare) = 0 Then
        modLog.LogError "E0101", "modInboxStore.SetInboxJudgement", "merged_into_unavailable"
        Exit Function
    End If
    If Not IsListedValue(IB_JUDGED, tgt) Then
        modLog.LogError "E0101", "modInboxStore.SetInboxJudgement", "invalid_status:" & tgt
        Exit Function
    End If

    Dim blockReason As String
    blockReason = JudgementError(tgt, dropType, reviveTag, (CDbl(reviveDue) > IB_NO_DATE))
    If LenB(blockReason) > 0 Then
        modLog.LogError "E0101", "modInboxStore.SetInboxJudgement", _
                        "e41_blocked:" & tgt & " " & blockReason
        Exit Function
    End If

    Dim ws As Object
    Dim blk As Variant
    Dim rowNo As Long
    If Not IbLocate(inboxId, ws, blk, rowNo) Then Exit Function

    If Not CanInboxTransition(IbValue(blk, rowNo, "status"), tgt) Then
        modLog.LogError "E0101", "modInboxStore.SetInboxJudgement", "invalid_transition:" & tgt
        Exit Function
    End If

    IbPut ws, blk, rowNo, "status", tgt
    IbPut ws, blk, rowNo, "drop_type", DropTypeFor(tgt, dropType)
    IbPut ws, blk, rowNo, "revive_tag", ReviveTagFor(tgt, reviveTag)
    IbPut ws, blk, rowNo, "revive_due", ReviveDueFor(tgt, reviveDue)
    IbPut ws, blk, rowNo, "judged_at", modUtil.NowStamp()
    ' 裁定書10 m7(13章§2.6): 実施者を judged_by_group へ残す。値源は config の
    ' operator(run_log の operator 列と同じ値源)。個人名ではなくグループ名まで
    ' を入れる運用であり、未設定なら空のまま(でっち上げない)。
    IbPut ws, blk, rowNo, "judged_by_group", Trim$(modConfig.GetStr("operator", vbNullString))
    ' 裁定書9 B2: 判定の確定後に入力列 judge_to をクリアする(store側が唯一の実施点)。
    IbPut ws, blk, rowNo, "judge_to", vbNullString
    SetInboxJudgement = True
    Exit Function

Failed:
    modLog.LogError "E0603", "modInboxStore.SetInboxJudgement", "write_failed", Err.Number
    SetInboxJudgement = False
End Function

' InterestText - 関心度の集計表示(10章 FR-17・11章 受信箱ワイヤー)。受信箱の
'   theme 列を投函順に**1件1行**で集め、集計は純核 InterestSummaryOf へ渡す。
'   件数集計・降順整列・2件以上・上限件数の規約は本関数には1つも無い
'   (シートI/Oは行を集めるだけ)。上限件数は純核が持つため引数で受けない。
'   本体シートの読取だけで完結し、ナレッジブックへは書かない(12章§4)。
Public Function InterestText() As String
    On Error GoTo Failed

    Dim ws As Object
    Set ws = IbSheet()
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = IbLastRow(ws)
    If lastRow < 2 Then Exit Function

    Dim blk As Variant
    blk = IbBlock(ws, lastRow)

    Dim acc As String
    Dim shown As Long
    Dim r As Long
    For r = 2 To lastRow
        ' 裁定書10 C1: 投函下書き行のテーマは「投函1件」ではないので数えない。
        If Not IsDraftId(IbValue(blk, r, "inbox_id")) Then
            ' 1セル内の改行で1件が2件に化けないよう空白へ倒してから積む。
            If shown > 0 Then acc = acc & vbLf
            acc = acc & Replace(Replace(IbValue(blk, r, "theme"), vbCr, " "), vbLf, " ")
            shown = shown + 1
        End If
    Next r

    InterestText = InterestSummaryOf(acc)
    Exit Function

Failed:
    modLog.LogError "E0603", "modInboxStore.InterestText", "read_failed", Err.Number
    InterestText = vbNullString
End Function

' --- 内部ヘルパー ---

' 投函下書き行か(13章§2.6・裁定書10 C1)。判定基準は inbox_id 列の固定マーカー
' だけであり、行番号には依存しない(利用者が行を挿入しても崩れない)。
Private Function IsDraftId(ByVal idText As String) As Boolean
    IsDraftId = (StrComp(Trim$(idText), IB_DRAFT_MARK, vbBinaryCompare) = 0)
End Function

' 受信箱IDから当月ぶんの連番を取り出す(月不一致・形違いは0)。IDの形の判定は
' IsValidInboxId が唯一の値源(2箇所で書かない)。
Private Function SerialOfInboxId(ByVal inboxId As String, ByVal monthText As String) As Long
    If Not IsValidInboxId(inboxId) Then Exit Function
    If Mid$(inboxId, 3, 6) <> monthText Then Exit Function
    SerialOfInboxId = CLng(Right$(inboxId, 3))
End Function

' 判定に対応しない列は空へ戻す(却下を条件付き保留へ付け替えたとき、前の判定の
' 語彙が行に残らないようにする)。
Private Function DropTypeFor(ByVal tgt As String, ByVal dropType As String) As String
    If StrComp(tgt, "rejected", vbBinaryCompare) = 0 Then DropTypeFor = Trim$(dropType)
End Function

Private Function ReviveTagFor(ByVal tgt As String, ByVal reviveTag As String) As String
    If StrComp(tgt, "conditional_hold", vbBinaryCompare) = 0 Then ReviveTagFor = Trim$(reviveTag)
End Function

Private Function ReviveDueFor(ByVal tgt As String, ByVal reviveDue As Date) As String
    If StrComp(tgt, "conditional_hold", vbBinaryCompare) <> 0 Then Exit Function
    If CDbl(reviveDue) <= IB_NO_DATE Then Exit Function
    ReviveDueFor = modUtilText.IsoDate(reviveDue)
End Function

' 集計済みキーの位置(1始まり。無ければ0)。InterestSummaryOf の内部ヘルパ。
' 件数は投函の種類数ぶんしか回らない。
Private Function IndexOfKey(ByRef keys() As String, ByVal kindCount As Long, _
                            ByVal k As String) As Long
    Dim i As Long
    For i = 1 To kindCount
        If StrComp(keys(i), k, vbBinaryCompare) = 0 Then
            IndexOfKey = i
            Exit Function
        End If
    Next i
End Function

' 件数の多い順に topN 件を1行へ。2件以上のテーマだけを載せる(FR-17)。
' InterestSummaryOf の内部ヘルパ(呼ぶのは純核1本だけ。シートには触れない)。
Private Function TopInterest(ByRef labels() As String, ByRef counts() As Long, _
                             ByVal kindCount As Long, ByVal topN As Long) As String
    Dim used() As Boolean
    ReDim used(0 To kindCount)

    Dim acc As String
    Dim shown As Long
    Dim n As Long
    For n = 1 To topN
        Dim best As Long
        Dim bestCount As Long
        best = 0
        bestCount = 1
        Dim i As Long
        For i = 1 To kindCount
            If Not used(i) Then
                If counts(i) > bestCount Then
                    best = i
                    bestCount = counts(i)
                End If
            End If
        Next i
        If best = 0 Then Exit For
        used(best) = True
        Dim oneText As String
        oneText = FmtInterestLine(labels(best), counts(best))
        If LenB(oneText) > 0 Then
            If shown > 0 Then acc = acc & IB_INTEREST_SEP
            acc = acc & oneText
            shown = shown + 1
        End If
    Next n
    TopInterest = acc
End Function

' 名前でシートを取る。無ければ Nothing(実行時に生やすと列定義の欠けた表になる)。
Private Function IbSheet() As Object
    On Error GoTo NoSheet
    Set IbSheet = ThisWorkbook.Worksheets(IB_SHEET)
    Exit Function
NoSheet:
    Set IbSheet = Nothing
End Function

' A列基準の最終行。データが無ければ1(見出し行)。
Private Function IbLastRow(ByVal ws As Object) As Long
    Dim n As Long
    On Error GoTo One1
    n = ws.Cells(ws.Rows.count, 1).End(IB_DIR_UP).row
    If n < 1 Then n = 1
    IbLastRow = n
    Exit Function
One1:
    IbLastRow = 1
End Function

' 見出し行を含む矩形を一度だけ読む(1セルだけの Range は2次元配列にならない)。
Private Function IbBlock(ByVal ws As Object, ByVal lastRow As Long) As Variant
    On Error GoTo Empty1
    Dim hi As Long
    hi = lastRow
    If hi < 2 Then hi = 2
    IbBlock = ws.Range(ws.Cells(1, 1), ws.Cells(hi, IB_SCAN_COLS)).Value
    Exit Function
Empty1:
    IbBlock = Empty
End Function

' 読み込み済みブロックから inbox_id 一致行を探す(見つからなければ0)。
Private Function IbRowOf(ByVal blk As Variant, ByVal lastRow As Long, _
                         ByVal idCol As Long, ByVal inboxId As String) As Long
    On Error GoTo NotFound
    Dim n As Long
    For n = 2 To lastRow
        ' 裁定書10 C1: 投函下書き行は ID 一致検索の対象外(判定・診断結果の
        ' 書込・読取はすべてこの検索を通るので、ここ1点で下書き行を閉じる)。
        If Not IsDraftId(CStr(blk(n, idCol))) Then
            If Trim$(CStr(blk(n, idCol))) = inboxId Then
                IbRowOf = n
                Exit Function
            End If
        End If
    Next n
    Exit Function
NotFound:
    IbRowOf = 0
End Function

' 「受信箱の1行を触る」ための場所決め(シート -> 最終行 -> 矩形 -> 該当行)。
'   どこかで欠けたら False(呼び出し側は黙って中止する。ここでログは書かない)。
Private Function IbLocate(ByVal inboxId As String, ByRef ws As Object, _
                          ByRef blk As Variant, ByRef rowNo As Long) As Boolean
    rowNo = 0
    Set ws = IbSheet()
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = IbLastRow(ws)
    blk = IbBlock(ws, lastRow)

    Dim cId As Long
    cId = modUtil.FindHeaderCol(blk, "inbox_id")
    If cId <= 0 Then Exit Function

    rowNo = IbRowOf(blk, lastRow, cId, Trim$(inboxId))
    IbLocate = (rowNo > 0)
End Function

' 列名で1セルを引く(13章冒頭。列が無ければ "")。
Private Function IbValue(ByVal blk As Variant, ByVal r As Long, _
                         ByVal headerName As String) As String
    On Error GoTo Empty0
    Dim c As Long
    c = modUtil.FindHeaderCol(blk, headerName)
    If c <= 0 Then Exit Function
    IbValue = Trim$(CStr(blk(r, c)))
    Exit Function
Empty0:
    IbValue = vbNullString
End Function

' 列名で位置を引いて1セル書く。列が無ければ何もしない。書込は必ず SetCellSafe
' (16章 NFR-S7①。body の32,000字打切りと先頭式記号の無害化もここが唯一の実装)。
Private Sub IbPut(ByVal ws As Object, ByVal hdr As Variant, ByVal r As Long, _
                  ByVal colName As String, ByVal textValue As String)
    Dim c As Long
    c = modUtil.FindHeaderCol(hdr, colName)
    If c <= 0 Then Exit Sub
    modUtilText.SetCellSafe ws.Cells(r, c), textValue, IB_SHEET & "/" & colName
End Sub
