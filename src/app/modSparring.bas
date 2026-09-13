Attribute VB_Name = "modSparring"
Option Explicit

' ============================================================================
' modSparring - PL-04 商談の予行演習(自由対話)の実行制御と履歴管理(app層・T-27)
' ----------------------------------------------------------------------------
' 正: 15章§6.5(system本文・CallChat・履歴保存・「受信箱へ」)/14章§6(CallChat の
'   帯域外成否・prevU/prevA は ";;;" 連結の**新しい順**)/14章§1(Stepレジストリ:
'   step=sp・play=PL-04・呼び出し元は modSparring)/13章§2.17(商談の予行演習シートの
'   sparring_log ブロック)・§2.2(sparring_u / sparring_a への保存)/
'   16章 E-44(履歴上限)・E-05(3)(送信前の modPii 必須)・E-04(境界記号の偽装除去)。
' 責務: 発話1本の送信(system組立 -> 走査 -> CallChat -> 履歴保存)、再開時の
'   履歴復元、選択発話の受信箱送信。プロンプト本文は持たない(正は15章§6.5・
'   実体は modPromptsOps.BuildSparringSystem / AsmSparringSystem)。
' R4(12章§4): 本モジュールは Excelトークン許可の13本に**入っていない**。案件一覧の
'   読取は modCaseRead、case_data の読み書きは modCaseStore、受信箱は
'   modInboxStore が唯一の口である(シートに1セルも直接触れない)。
' スキーマ無し: 自由対話なので JSON防衛線(14章§5)は通さない。成否は CallChat の
'   ByRef ok / errCode だけで判定し、応答本文の "#ERR:" では判定しない(14章§6)。
' run_log: 自由対話には検証段が無いため CallChat が自分で1行書く(14章§6の注記)。
'   本モジュールは run_log を書かない(1呼び出し=2行にしない)。
'
' 【保存形式】13章§2.2 の sparring_u / sparring_a の中身。本モジュールが唯一の
'   読み書き点である:
'     1発話 = 1行 = `seq <TAB> spoke_at <TAB> 本文`   行区切りは vbLf
'     本文は modJsonLite.EscapeJsonStr で \n \t \\ を畳む(復元は UnescapeJsonStr)。
'   seq は13章§2.17 sparring_log と同じ**発話単位の通し連番**で、user と ai が
'   1本の番号列を共有する(発話=n / その応答=n+1)。
'   case_data の `seq` 列は §2.2 の定義どおり 32,000字の**分割連番**のままである。
'   発話単位の seq と spoke_at をこの行形式が持つのは、`SaveData` が data_key
'   単位で全行を置換し seq を分割連番に使う契約(14章§6)であり、1発話=1物理行を
'   case_data 側に持たせる口が無いためである(§2.2 の列定義と §2.2 data_key 注記
'   「発話単位seqで保存」の食い違いはここで吸収する)。
' ============================================================================

Private Const SP_SRC As String = "modSparring"

' 13章§2.2 の data_key(商談の予行演習履歴の保存先)。
Private Const SP_KEY_U As String = "sparring_u"
Private Const SP_KEY_A As String = "sparring_a"

' 13章§2.17 sparring_log の role(日本語表示 自分 / AI は ui 側)。
Private Const SP_ROLE_U As String = "user"
Private Const SP_ROLE_A As String = "ai"

' 13章§2.6 source_kind の enum のうち商談の予行演習発話が使う値(15章§6.5)。
Private Const SP_SOURCE_KIND As String = "field_voice"

' config 既定(13章§2.3 sparring_max_turns。NFR-M3)。
Private Const SP_MAX_TURNS_DFLT As Long = 12

' 受信箱テーマ(案件ID＋要約)と資料要約の切詰め字数。
Private Const SP_THEME_CHARS As Long = 40
Private Const SP_DIGEST_CHARS As Long = 400

' 貼付欄のうち商談の予行演習の資料要約へ載せる2欄(13章§2.2)。全文はS1が既に吸っており、
' ここへ再掲すると E-03 の予算を二重に食う。
Private Const SP_NOTE_KEYS As String = "input_field_notes|input_coverage_note"
Private Const SP_NOTE_LABELS As String = "現場メモ|付保の見立て"

' ============================================================================
' 純関数(シート・ログ・LLMに触れない判定核。14章§6・PURE_ALLOWLIST)
' ----------------------------------------------------------------------------
' 規約そのもの(送信可否・履歴上限・線上形式への変換)を実行制御の中に閉じ込めると
' 層(a)から誰も検査できない。modTestsPure から直接叩く前提で公開する。
' **Private へ戻すことは契約違反**(vba_lint の CONTRACT required が検出する)。
' ============================================================================

' CanContinueSparring - 発話を1本送ってよいかの【唯一の判定点】(fail-closed)。
'   (1) caseIdText が 13章§1 の案件ID書式(判定は modCaseStore.IsValidCaseId が
'       唯一の実装。書式を2箇所に書かない)
'   (2) 発話が空白・改行だけでない
'   (3) hasPii=False。16章 E-05(3)は商談の予行演習の発話送信前の検知で**送信をブロック**
'       する。走査そのものは modPii が唯一の実装なので、ここは結果の真偽だけを
'       受け取る(検知規則を2箇所に書かない)。
Public Function CanContinueSparring(ByVal caseIdText As String, _
                                    ByVal utterance As String, _
                                    ByVal hasPii As Boolean) As Boolean
    Dim t As String

    If hasPii Then Exit Function
    If Not modCaseStore.IsValidCaseId(caseIdText) Then Exit Function

    t = Replace(utterance, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    If LenB(Trim$(t)) = 0 Then Exit Function

    CanContinueSparring = True
End Function

' TrimHistoryOf - 保存形式の履歴を【直近 maxTurns 発話】へ切り詰める(古い順の
'   まま返す)。16章 E-44「渡す履歴を直近 sparring_max_turns 往復に制限」を
'   保存形式の側で行う唯一の点であり、全履歴は case_data に残る(切り捨てるのは
'   渡す分だけ)。maxTurns <= 0 は全件(切らない)。
'   線上形式(";;;"連結)側の最終防衛は modGatewayRPN.TrimHistoryPairs が別に持つ。
'   形式が違うので同じ実装は使えないが、**件数の値はどちらも config
'   sparring_max_turns の1箇所**から来る(答えの源は1つ)。
Public Function TrimHistoryOf(ByVal storedText As String, ByVal maxTurns As Long) As String
    Dim rows() As String
    Dim lo As Long, hi As Long
    Dim i As Long
    Dim buf() As String
    Dim cnt As Long

    If LenB(storedText) = 0 Then Exit Function
    If maxTurns <= 0 Then
        TrimHistoryOf = storedText
        Exit Function
    End If

    rows = RowsOf(storedText)
    lo = LBound(rows)
    hi = UBound(rows)
    If hi < lo Then Exit Function
    If (hi - lo + 1) <= maxTurns Then
        TrimHistoryOf = storedText
        Exit Function
    End If

    ' BufText の区切りは vbLf(14章§6)なので保存形式の行区切りと一致する。
    modUtil.BufInit buf, cnt
    For i = hi - maxTurns + 1 To hi
        modUtil.BufAdd buf, cnt, rows(i)
    Next i
    TrimHistoryOf = modUtil.BufText(buf, cnt)
End Function

' HistoryJoinOf - 保存形式の履歴から CallChat の histU / histA を組む【唯一の点】。
'   14章§6・15章§6.5: 直近 maxTurns 発話を**新しい順**に modGatewayRPN.GW_HIST_SEP
'   (";;;")で連結する(PoC 裁定D11の実証方式)。切詰めは TrimHistoryOf に委ねる。
'   本文は保存形式のエスケープを解いて原文へ戻す。本文中に区切りが現れたら ";" へ
'   潰す(線上形式だけの非可逆処理。case_data 側の原文は書き換えない)。
Public Function HistoryJoinOf(ByVal storedText As String, ByVal maxTurns As Long) As String
    Dim kept As String
    Dim rows() As String
    Dim lo As Long, hi As Long
    Dim i As Long
    Dim cnt As Long
    Dim outArr() As String
    Dim sepText As String

    kept = TrimHistoryOf(storedText, maxTurns)
    If LenB(kept) = 0 Then Exit Function

    rows = RowsOf(kept)
    lo = LBound(rows)
    hi = UBound(rows)
    If hi < lo Then Exit Function

    sepText = modGatewayRPN.GW_HIST_SEP
    ReDim outArr(0 To hi - lo)
    cnt = 0
    For i = hi To lo Step -1
        outArr(cnt) = Replace(BodyOfRow(rows(i)), sepText, ";")
        cnt = cnt + 1
    Next i
    HistoryJoinOf = Join(outArr, sepText)
End Function

' ============================================================================
' 実行制御
' ============================================================================

' ResumeSparring - 「商談の予行演習を開始/再開」(11章 商談の予行演習ワイヤー・15章§6.5)。
'   戻り値=保存済みの発話数(0=履歴なし＝新規開始)。**-1=案件一覧を読めない**
'   (fail-closed。ui は開始させない)。contextNote は 13章§2.17 の
'   `sp_context_note` へ出す表示文字列(例「ドシエ+S1-S3+型/機構 注入済」)。
'   dossier_tier の t3_sparring への自動昇格(13章§2.1・§2.17)は、裁定書9 A-1
'   により modCaseStore.PromoteTier(14章§6・N2)を唯一の書込口として**実行する**。
'   開始/再開のたびに呼ぶ(既に t3_sparring でも同値の書込で害はない)。昇格の
'   失敗(案件行が無い等)は usage_log へ事実を残して続行し、商談の予行演習の開始その
'   ものは止めない。
Public Function ResumeSparring(ByVal caseId As String, ByRef contextNote As String) As Long
    On Error GoTo Failed

    Dim ctx As TCaseCtx
    Dim rNo As Long
    Dim qMode As String
    Dim s4v As String
    Dim tierText As String

    contextNote = vbNullString
    ResumeSparring = -1

    If Not modCaseStore.IsValidCaseId(caseId) Then
        modLog.LogError "E0101", SP_SRC & ".ResumeSparring", "invalid_case_id"
        Exit Function
    End If
    If Not modCaseRead.ReadCaseCtx(caseId, ctx, rNo, qMode, s4v, tierText) Then
        modLog.LogUsage "case_ctx_unavailable", caseId, "案件一覧を読めないため商談の予行演習中止"
        Exit Function
    End If

    contextNote = ContextNoteOf(caseId)

    ' 裁定書9 A-1: t3_sparring への自動昇格を実行する(書込口は N2 の PromoteTier
    ' のみ)。失敗しても商談の予行演習の開始は止めない(fail-closed にしない設計どおり)。
    If modCaseStore.PromoteTier(caseId, "t3_sparring") Then
        modLog.LogUsage "sparring_tier_promoted", caseId, _
                        "dossier_tier=" & tierText & " -> t3_sparring"
    Else
        modLog.LogUsage "sparring_tier_promote_failed", caseId, _
                        "dossier_tier=" & tierText & " 昇格できず続行"
    End If

    ResumeSparring = RowCountOf(modCaseStore.LoadData(caseId, SP_KEY_U)) + _
                     RowCountOf(modCaseStore.LoadData(caseId, SP_KEY_A))
    Exit Function

Failed:
    modLog.LogError "E0603", SP_SRC & ".ResumeSparring", "unexpected", Err.Number
    ResumeSparring = -1
End Function

' HistoryOf - 保存済み履歴を保存形式のまま返す(11章の履歴表示・再開で ui が
'   seq / spoke_at / 本文へ分解する)。roleKind は 13章§2.17 の enum(user / ai)。
'   不正な role・不正な案件IDは ""。
Public Function HistoryOf(ByVal caseId As String, ByVal roleKind As String) As String
    Dim dataKey As String

    dataKey = DataKeyOf(roleKind)
    If LenB(dataKey) = 0 Then Exit Function
    If Not modCaseStore.IsValidCaseId(caseId) Then Exit Function
    HistoryOf = modCaseStore.LoadData(caseId, dataKey)
End Function

' SendSparring - 発話1本の送信(15章§6.5・PL-04)。
'   True=応答を受け取り、発話と応答を case_data へ保存できた。
'   replyText: 応答本文(True のときだけ入る)。
'   errCode  : 遮断・失敗の内訳。E0103=PII検知で送信前に遮断(16章E-05(3))/
'     E0101=前提不足/E0604=保存失敗/E02xx=経路失敗(CallChat が帯域外で返す値を
'     そのまま透す)。16章 E-44 の「往復数を減らして再開」の案内は E0204 で行う。
'   成否は CallChat の ByRef ok だけで判定する(応答本文の "#ERR:" では判定
'   しない。14章§6のエラー規約)。
Public Function SendSparring(ByVal caseId As String, ByVal utterance As String, _
                             ByRef replyText As String, ByRef errCode As String) As Boolean
    On Error GoTo Failed

    Dim detailAcc As String
    Dim safeText As String
    Dim piiNote As String
    Dim sysText As String
    Dim sysNote As String
    Dim histU As String
    Dim histA As String
    Dim maxTurns As Long
    Dim latencyMs As Long
    Dim rawText As String
    Dim seqNo As Long
    Dim callOk As Boolean

    replyText = vbNullString
    errCode = vbNullString

    ' 16章 E-04: 外部由来テキストは組立の前に無害化する(件数だけを detail へ)。
    safeText = Sanitized(utterance, detailAcc)

    ' 16章 E-05(3): 送信前の走査は modPii が唯一の実装。検知したら送信しない。
    piiNote = modPii.ScanReport(safeText, "商談の予行演習/sp_input")
    If Not CanContinueSparring(caseId, safeText, (LenB(piiNote) > 0)) Then
        If LenB(piiNote) > 0 Then
            errCode = "E0103"
            modLog.LogError errCode, SP_SRC & ".SendSparring", piiNote
        Else
            errCode = "E0101"
            modLog.LogError errCode, SP_SRC & ".SendSparring", "invalid_case_or_empty"
        End If
        Exit Function
    End If

    sysText = BuildSystemFor(caseId)
    If LenB(sysText) = 0 Then
        errCode = "E0101"
        Exit Function
    End If

    ' 資料側(案件JSON・ナレッジ)の走査は【記録だけ】行う。E-05 が遮断を命じる
    ' 実施点は (1)貼付欄 と (3)商談の予行演習の発話であり、貼付欄は入力時に遮断済みで
    ' ある。ここで資料側まで遮断すると遮断点がE-05の一覧から増え、かつ検証済み
    ' の sN_json で商談の予行演習が恒久的に開けなくなる。黙殺はしない。
    sysNote = modPii.ScanReport(sysText, "商談の予行演習/system")
    If LenB(sysNote) > 0 Then
        modLog.LogError "E0103", SP_SRC & ".SendSparring", sysNote
    End If

    maxTurns = MaxTurnsCfg()
    histU = HistoryJoinOf(modCaseStore.LoadData(caseId, SP_KEY_U), maxTurns)
    histA = HistoryJoinOf(modCaseStore.LoadData(caseId, SP_KEY_A), maxTurns)

    rawText = modGatewayRPN.CallChat(caseId, sysText, safeText, histU, histA, _
                                     callOk, errCode, latencyMs)
    If Not callOk Then
        If errCode = "E0204" Then
            ' 16章 E-44: 上限系は「往復数を減らして再開」を案内する。
            modLog.LogUsage "sparring_turn_limit", caseId, _
                            "履歴を減らして再開してください(max_turns=" & CStr(maxTurns) & ")"
        End If
        Exit Function
    End If

    If Not AppendPair(caseId, safeText, rawText, seqNo) Then
        errCode = "E0604"
        modLog.LogError errCode, SP_SRC & ".SendSparring", "history_save_failed"
        Exit Function
    End If

    AddNote detailAcc, "seq=" & CStr(seqNo)
    AddNote detailAcc, "sent_turns=" & CStr(maxTurns)
    ' 15章§6.5 は {{schemes}} を「全status」と書くが、modKnowledge の型ライブラリ
    ' 注入は S3用の proven/adopted 絞込しか口が無い(14章§6 SchemesFor)。黙って
    ' 狭めない(concerns へ挙げた既知の差分)。
    AddNote detailAcc, "schemes=proven_adopted_only"
    modLog.LogUsage "sparring_turn", caseId, detailAcc

    replyText = rawText
    SendSparring = True
    Exit Function

Failed:
    modLog.LogError "E0603", SP_SRC & ".SendSparring", "unexpected", Err.Number
    SendSparring = False
End Function

' SendToInbox - 選択した発話を受信箱へ登録する(15章§6.5「受信箱へ」)。
'   source_kind=field_voice / theme=案件ID＋発話の要約。戻り値=inbox_id
'   (失敗は "")。同一発話の二重送信の抑止は 13章§2.17 の `inbox_id` 列(商談の予行演習
'   シート側)が鍵なので、ここでは持たない(答えを2箇所に置かない)。
'   本文にPIIを検知したら登録しない(16章 E-05(2) 受信箱bodyは保存をブロック)。
Public Function SendToInbox(ByVal caseId As String, ByVal roleKind As String, _
                            ByVal seqNo As Long) As String
    On Error GoTo Failed

    Dim storedText As String
    Dim bodyText As String
    Dim piiNote As String

    storedText = HistoryOf(caseId, roleKind)
    If LenB(storedText) = 0 Then Exit Function

    bodyText = UtteranceOf(storedText, seqNo)
    If LenB(Trim$(bodyText)) = 0 Then
        modLog.LogError "E0101", SP_SRC & ".SendToInbox", "utterance_not_found"
        Exit Function
    End If

    piiNote = modPii.ScanReport(bodyText, "商談の予行演習/受信箱送信")
    If LenB(piiNote) > 0 Then
        modLog.LogError "E0103", SP_SRC & ".SendToInbox", piiNote
        Exit Function
    End If

    SendToInbox = modInboxStore.NewInboxItem(SP_SOURCE_KIND, _
                                             ThemeOf(caseId, bodyText), bodyText)
    Exit Function

Failed:
    modLog.LogError "E0603", SP_SRC & ".SendToInbox", "unexpected", Err.Number
    SendToInbox = vbNullString
End Function

' ============================================================================
' 内部ヘルパー(組立・保存・保存形式の読み書き)
' ============================================================================

' 15章§6.5 の system を1本組む。案件一覧を読めなければ "" を返す(fail-closed)。
Private Function BuildSystemFor(ByVal caseId As String) As String
    Dim ctx As TCaseCtx
    Dim rNo As Long
    Dim qMode As String
    Dim s4v As String
    Dim tierText As String
    Dim schemesText As String
    Dim patternsText As String
    Dim mechsText As String
    Dim rulesText As String

    If Not modCaseRead.ReadCaseCtx(caseId, ctx, rNo, qMode, s4v, tierText) Then
        modLog.LogUsage "case_ctx_unavailable", caseId, "案件一覧を読めないため商談の予行演習中止"
        Exit Function
    End If

    ' 注入IDの累積は呼び出しの都度初期化する(14章§6・13章§2.4)。
    modKnowledge.ResetInjectedIds
    schemesText = modKnowledge.SchemesFor(vbNullString)
    patternsText = modKnowledge.PatternsText()
    mechsText = modKnowledge.MechsText()
    rulesText = modKnowledge.RulesText()

    BuildSystemFor = modPromptsOps.AsmSparringSystem(DossierSummaryOf(caseId, ctx), _
                                                     S1S2S3JsonOf(caseId), _
                                                     schemesText, patternsText, _
                                                     mechsText, rulesText)
End Function

' {{dossierSummary}} = S1の business_summary ＋ 入力の要約(15章§6.5)。
Private Function DossierSummaryOf(ByVal caseId As String, ByRef ctx As TCaseCtx) As String
    Dim buf() As String
    Dim cnt As Long
    Dim s1Json As String
    Dim i As Long
    Dim keyName As String
    Dim labelText As String
    Dim noteText As String
    Dim dropped As Long
    Dim marked As Long

    modUtil.BufInit buf, cnt
    modUtil.BufAdd buf, cnt, "企業: " & ctx.company & "(" & ctx.industry_name & ")"
    modUtil.BufAdd buf, cnt, "案件区分: " & ctx.case_type & " / 調査の深さ: " & ctx.dossier_tier

    s1Json = modCaseStore.ResolveStepJson(caseId, 1)
    If LenB(s1Json) > 0 Then
        modUtil.BufAdd buf, cnt, "事業概要: " & modJsonLite.GetStr(s1Json, "business_summary")
    End If

    For i = 0 To 1
        keyName = PickAt(SP_NOTE_KEYS, i)
        labelText = PickAt(SP_NOTE_LABELS, i)
        ' 貼付欄は外部由来なので注入の前に無害化する(16章 E-04・15章§0 原則9)。
        noteText = modUtilText.SanitizeInput(modCaseStore.LoadData(caseId, keyName), _
                                             dropped, marked)
        noteText = Digest(noteText)
        If LenB(noteText) > 0 Then
            modUtil.BufAdd buf, cnt, labelText & ": " & noteText
        End If
    Next i

    DossierSummaryOf = modUtil.BufText(buf, cnt)
End Function

' {{s1s2s3Json}} = 現時点の分析結果(15章§6.5)。参照優先の解決は
'   modCaseStore.ResolveStepJson が唯一の口(13章§2.2)。未実行のStepは null。
Private Function S1S2S3JsonOf(ByVal caseId As String) As String
    Dim acc As String

    acc = "{""s1"": " & OrNullJson(modCaseStore.ResolveStepJson(caseId, 1))
    acc = acc & ", ""s2"": " & OrNullJson(modCaseStore.ResolveStepJson(caseId, 2))
    acc = acc & ", ""s3"": " & OrNullJson(modCaseStore.ResolveStepJson(caseId, 3))
    S1S2S3JsonOf = acc & "}"
End Function

Private Function OrNullJson(ByVal jsonText As String) As String
    If LenB(Trim$(jsonText)) = 0 Then
        OrNullJson = "null"
    Else
        OrNullJson = jsonText
    End If
End Function

' 13章§2.17 `sp_context_note` の表示文字列(例「ドシエ+S1-S3+型/機構 注入済」)。
Private Function ContextNoteOf(ByVal caseId As String) As String
    Dim acc As String
    Dim stepsText As String

    If LenB(modCaseStore.LoadData(caseId, "input_dossier")) > 0 Then acc = "ドシエ"

    stepsText = StepsPresentOf(caseId)
    If LenB(stepsText) > 0 Then
        If LenB(acc) > 0 Then acc = acc & "+"
        acc = acc & stepsText
    End If

    If LenB(acc) > 0 Then acc = acc & "+"
    ContextNoteOf = acc & "型/機構 注入済"
End Function

' 揃っているStepの範囲表示("S1-S3" / "S1-S2" / "S1" / "")。
Private Function StepsPresentOf(ByVal caseId As String) As String
    Dim topStep As Long
    Dim i As Long

    For i = 1 To 3
        If LenB(modCaseStore.ResolveStepJson(caseId, i)) = 0 Then Exit For
        topStep = i
    Next i

    If topStep <= 0 Then Exit Function
    If topStep = 1 Then
        StepsPresentOf = "S1"
    Else
        StepsPresentOf = "S1-S" & CStr(topStep)
    End If
End Function

' 発話と応答を1組ずつ保存形式へ追記する(13章§2.2)。seqNo には発話側に振った
'   通し連番を返す(応答は seqNo + 1)。どちらかの保存に失敗したら False。
Private Function AppendPair(ByVal caseId As String, ByVal userText As String, _
                            ByVal aiText As String, ByRef seqNo As Long) As Boolean
    Dim storedU As String
    Dim storedA As String
    Dim stampText As String
    Dim maxU As Long
    Dim maxA As Long

    storedU = modCaseStore.LoadData(caseId, SP_KEY_U)
    storedA = modCaseStore.LoadData(caseId, SP_KEY_A)
    maxU = MaxSeqOf(storedU)
    maxA = MaxSeqOf(storedA)
    seqNo = maxU
    If maxA > seqNo Then seqNo = maxA
    seqNo = seqNo + 1

    stampText = modUtil.NowStamp()
    If Not modCaseStore.SaveData(caseId, SP_KEY_U, _
                                 AppendRow(storedU, seqNo, stampText, userText)) Then
        Exit Function
    End If
    If Not modCaseStore.SaveData(caseId, SP_KEY_A, _
                                 AppendRow(storedA, seqNo + 1, stampText, aiText)) Then
        Exit Function
    End If
    AppendPair = True
End Function

' 保存形式へ1行足す(モジュール冒頭の【保存形式】)。
Private Function AppendRow(ByVal storedText As String, ByVal seqNo As Long, _
                           ByVal stampText As String, ByVal bodyText As String) As String
    Dim rowText As String

    rowText = CStr(seqNo) & vbTab & stampText & vbTab & _
              modJsonLite.EscapeJsonStr(bodyText)
    If LenB(storedText) = 0 Then
        AppendRow = rowText
    Else
        AppendRow = storedText & vbLf & rowText
    End If
End Function

' 保存形式を行へ分ける(空行は落とす。前後空白は本文の一部なので削らない)。
Private Function RowsOf(ByVal storedText As String) As String()
    Dim raw() As String
    Dim outArr() As String
    Dim i As Long
    Dim cnt As Long

    If LenB(storedText) = 0 Then
        RowsOf = Split(vbNullString)
        Exit Function
    End If

    raw = Split(storedText, vbLf)
    ReDim outArr(0 To UBound(raw) - LBound(raw))
    cnt = 0
    For i = LBound(raw) To UBound(raw)
        If LenB(raw(i)) > 0 Then
            outArr(cnt) = raw(i)
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        RowsOf = Split(vbNullString)
    Else
        ReDim Preserve outArr(0 To cnt - 1)
        RowsOf = outArr
    End If
End Function

Private Function RowCountOf(ByVal storedText As String) As Long
    Dim rows() As String

    rows = RowsOf(storedText)
    If UBound(rows) < LBound(rows) Then Exit Function
    RowCountOf = UBound(rows) - LBound(rows) + 1
End Function

' 保存形式1行の本文(エスケープを解いた原文)。区切りが無い行は行全体を本文と
'   みなす(手で編集されたブックでも履歴を落とさない)。
Private Function BodyOfRow(ByVal rowText As String) As String
    Dim parts() As String

    parts = Split(rowText, vbTab)
    If UBound(parts) - LBound(parts) < 2 Then
        BodyOfRow = rowText
        Exit Function
    End If
    BodyOfRow = modJsonLite.UnescapeJsonStr(parts(LBound(parts) + 2))
End Function

' 保存形式1行の seq(発話単位の通し連番)。読めない行は 0。
Private Function SeqOfRow(ByVal rowText As String) As Long
    Dim parts() As String

    parts = Split(rowText, vbTab)
    If UBound(parts) - LBound(parts) < 2 Then Exit Function
    SeqOfRow = CLng(Val(Trim$(parts(LBound(parts)))))
End Function

Private Function MaxSeqOf(ByVal storedText As String) As Long
    Dim rows() As String
    Dim i As Long
    Dim v As Long
    Dim topSeq As Long

    rows = RowsOf(storedText)
    For i = LBound(rows) To UBound(rows)
        v = SeqOfRow(rows(i))
        If v > topSeq Then topSeq = v
    Next i
    MaxSeqOf = topSeq
End Function

' seq を指定して1発話の本文を取り出す(受信箱送信・表示用)。無ければ ""。
Private Function UtteranceOf(ByVal storedText As String, ByVal seqNo As Long) As String
    Dim rows() As String
    Dim i As Long

    If seqNo <= 0 Then Exit Function
    rows = RowsOf(storedText)
    For i = LBound(rows) To UBound(rows)
        If SeqOfRow(rows(i)) = seqNo Then
            UtteranceOf = BodyOfRow(rows(i))
            Exit Function
        End If
    Next i
End Function

' 13章§2.17 の role(user / ai)を 13章§2.2 の data_key へ。表に無い値は ""。
Private Function DataKeyOf(ByVal roleKind As String) As String
    Select Case Trim$(roleKind)
        Case SP_ROLE_U
            DataKeyOf = SP_KEY_U
        Case SP_ROLE_A
            DataKeyOf = SP_KEY_A
    End Select
End Function

' 受信箱の theme(案件ID＋要約。15章§6.5)。
Private Function ThemeOf(ByVal caseId As String, ByVal bodyText As String) As String
    Dim t As String

    t = Replace(bodyText, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    ThemeOf = Trim$(caseId) & " " & Trim$(modUtil.SafeLeft(Trim$(t), SP_THEME_CHARS))
End Function

' 資料要約用の切詰め(改行は " / " へ畳む)。
Private Function Digest(ByVal rawText As String) As String
    Dim t As String

    If LenB(Trim$(rawText)) = 0 Then Exit Function
    t = Replace(rawText, vbCr, vbLf)
    t = Replace(t, vbLf, " / ")
    t = Replace(t, vbTab, " ")
    Digest = Trim$(modUtil.SafeLeft(Trim$(t), SP_DIGEST_CHARS))
End Function

Private Function MaxTurnsCfg() As Long
    MaxTurnsCfg = modConfig.GetLong("sparring_max_turns", SP_MAX_TURNS_DFLT)
End Function

' 外部由来テキストの無害化(16章 E-04)。件数だけ detail へ(本文は書かない)。
Private Function Sanitized(ByVal rawText As String, ByRef detailAcc As String) As String
    Dim removedN As Long
    Dim markerN As Long

    Sanitized = modUtilText.SanitizeInput(rawText, removedN, markerN)
    If removedN > 0 Then AddNote detailAcc, "e04_removed=" & CStr(removedN)
    If markerN > 0 Then AddNote detailAcc, "e04_marker=" & CStr(markerN)
End Function

' usage_log detail の積み方(400字上限は modLog.TruncDetail)。
Private Sub AddNote(ByRef acc As String, ByVal noteText As String)
    If LenB(noteText) = 0 Then Exit Sub
    If LenB(acc) > 0 Then acc = acc & ";"
    acc = acc & noteText
End Sub

' "a|b" の idx 番目(0始まり。範囲外は "")。
Private Function PickAt(ByVal listText As String, ByVal idx As Long) As String
    Dim parts() As String

    parts = Split(listText, "|")
    If idx < 0 Or idx > UBound(parts) Then Exit Function
    PickAt = parts(idx)
End Function
