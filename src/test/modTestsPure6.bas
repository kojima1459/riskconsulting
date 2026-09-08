Attribute VB_Name = "modTestsPure6"
Option Explicit

' ============================
' modTestsPure6 - mock受入条件・障害注入・防衛線(2.5)・ファイル名規則(層(a))
' ----------------------------
' 役割: modTestsPure5(15章§11の全64ケース)の姉妹モジュール。30,000字契約の
'   ため2分割した。本モジュールは次の4群を持つ。入口は Public Sub RunAll()。
'     G40  15章§8.1受入条件1: mock11応答が**自バリアント文脈**で modValidate に
'          合格すること(**警告判定のケースも発火させない**=戻り値が空文字)。
'          17章 T-22 DoD「T-14の時点では検証できない」分の実体がここ。
'     G41  15章§8.2の障害注入が modValidate でどう観測されるか。mock応答の
'          **形**(破損・M-9999混入・stories2件・空・上限文字列・#ERR行)と
'          DecideOk の判定は modTestsPure2 の G7/G10 が既に固定しているので、
'          本群はそこに無い「検証器の反応」だけを見る(答えを2箇所に書かない)。
'     G42  14章§5 防衛線(2.5) NormalizeLlmJson(16章E-49の重複排除・E0303)。
'     G43  13章§2.8 SanitizeFileName のファイル名生成規則。
'
' 統合者へ: modTestsPure5.RunAll の末尾から modTestsPure6.RunAll を呼ぶこと。
'   build/modules.json への登録と tests_expected の更新は modTestsPure5 の
'   ヘッダに記した3点セットのとおり(本ファイル29本)。
'
' テスト本数: 34本 = G40 10 / G40K 2 / G41 7 / G41K 1 / G42 4 / G43 5 / G32G 5
'   (K付きはID実在検査を通る群。裁定書7 A-1/A-2 で一覧テキストは Check系の
'    引数になったので、本モジュールが15章の1行書式で自給する Wl* を渡す。
'    modKnowledge には触れない)
'
' 設計判断(R4準拠): Excelトークン不使用。改行は vbLf 基準。乱数・時刻不使用。
' ============================

' 15章§8.1の11 mock ID(ResponseById のキーの正)。
Private Const MK_S1N As String = "MK-S1-NEW"
Private Const MK_S1R As String = "MK-S1-RNW"
Private Const MK_S2N As String = "MK-S2-NEW"
Private Const MK_S2R As String = "MK-S2-RNW"
Private Const MK_S3 As String = "MK-S3"
Private Const MK_S4 As String = "MK-S4"
Private Const MK_PF As String = "MK-PF"
Private Const MK_C2H As String = "MK-S2C-HIT"
Private Const MK_C2C As String = "MK-S2C-CLEAN"
Private Const MK_C3H As String = "MK-S3C-HIT"
Private Const MK_C3C As String = "MK-S3C-CLEAN"

' 13章§2.1 case_type の2値。
Private Const CT_NEW As String = "new"
Private Const CT_RNW As String = "renewal"

' ----------------------------
' RunAll: グループ単位で隔離実行する。1グループが実行時エラーで落ちても
'   残りのグループは走る(未実装/未注入の事実は GroupFail で1件の失敗として
'   可視化し、無かったことにしない)。
' ----------------------------
Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 7
        grpName = "G?" & i
        On Error Resume Next
        Err.Clear
        RunGroup i, grpName
        If Err.Number <> 0 Then
            GroupFail grpName
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    ' 姉妹モジュール(30,000字契約による分割)を同じ隔離作法で続けて回す。
    ' 数珠つなぎ: modTestsPure -> 2 -> 3 -> 4 -> 5 -> 6 -> 7(裁定書7 D-13/D-14)。
    On Error Resume Next
    Err.Clear
    modTestsPure7.RunAll
    If Err.Number <> 0 Then
        GroupFail "modTestsPure7.RunAll"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub RunGroup(ByVal grpNo As Long, ByRef grpName As String)
    Select Case grpNo
    Case 1
        grpName = "G40 mock自文脈"
        T_MockOk
    Case 2
        grpName = "G40K mock自文脈 ID実在"
        T_MockOkIds
    Case 3
        grpName = "G41 障害注入"
        T_Fault
    Case 4
        grpName = "G41K 障害注入 ID実在"
        T_FaultIds
    Case 5
        grpName = "G42 NormalizeLlmJson"
        T_Normalize
    Case 6
        grpName = "G43 SanitizeFileName"
        T_FileName
    Case 7
        grpName = "G32G CheckS3 growth_ideas"
        T_S3Growth
    End Select
End Sub

' ============================
' 共通ヘルパ(modTestsPure3/4 と同じ作法。各 modTestsPure* が自前で持つ)
' ============================
Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

Private Function HeadOf(ByVal s As String) As String
    If Len(s) <= 120 Then
        HeadOf = s
    Else
        HeadOf = Left$(s, 120) & "...(全" & Len(s) & "字)"
    End If
End Function

' ch を n 個ならべる(倍々に伸ばして切る)。
Private Function RepChar(ByVal ch As String, ByVal n As Long) As String
    Dim s As String
    If n <= 0 Or Len(ch) <= 0 Then Exit Function
    s = ch
    Do While Len(s) < n
        s = s & s
    Loop
    RepChar = Left$(s, n)
End Function

' 検証結果の中に「行頭が [ケースID]」の行があるか(15章§11・17章 T-22)。
Private Function HasCase(ByVal outText As String, ByVal caseId As String) As Boolean
    Dim lines() As String
    Dim i As Long
    Dim mark As String

    mark = "[" & caseId & "]"
    lines = Split(outText, vbLf)
    For i = 0 To UBound(lines)
        If Left$(LTrim$(lines(i)), Len(mark)) = mark Then
            HasCase = True
            Exit Function
        End If
    Next i
End Function

Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), _
        "期待=[" & HeadOf(want) & "] 実際=[" & HeadOf(act) & "]"
End Sub

Private Sub ChkN(ByVal nm As String, ByVal act As Long, ByVal want As Long)
    modTestRunner.Check nm, (act = want), "期待=" & want & " 実際=" & act
End Sub

Private Sub ChkB(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check nm, cond, detail
End Sub

' 15章§8.1受入条件1: 自バリアント文脈では警告を含め0件。
Private Sub ChkZero(ByVal nm As String, ByVal outText As String)
    modTestRunner.Check nm, (outText = ""), "0件(空文字)を期待。実際=[" & HeadOf(outText) & "]"
End Sub

' 障害注入mockが§8.2の表どおりのケースを踏むこと。ケースIDは頭と尾に割って
' 渡す(照合スクリプトの「1ケース1本」計数を乱さないため。本体の1本はG31/G32が持つ)。
Private Sub ChkFault(ByVal nm As String, ByVal outText As String, _
                     ByVal idHead As String, ByVal idTail As String)
    modTestRunner.Check nm, HasCase(outText, idHead & idTail), _
        "期待ケース=" & idHead & idTail & " 実際=[" & HeadOf(outText) & "]"
End Sub

' 配列件数(modTestsPure/2 と同じ作法。空Collection契約に合わせ Nothing は -1)。
Private Function ArrCount(ByVal js As String, ByVal ky As String) As Long
    Dim c As Collection
    Set c = modJsonLite.GetArrayItems(js, ky)
    If c Is Nothing Then
        ArrCount = -1
    Else
        ArrCount = c.Count
    End If
End Function

' ============================
' JSON最小改変ヘルパ(純文字列。mock正常応答から狙った1ケースだけを壊す)
'   キーは**最初に現れた1件**を対象にする。15章のスキーマ順に読むと、
'   risks/stories/slides 等の先頭要素の項目が最初に現れる。
' ============================
Private Function MockJson(ByVal mockId As String) As String
    MockJson = modJsonLite.ExtractJsonBlock(modMockLlm.ResponseById(mockId))
End Function

' ============================
' 実在ID一覧のホワイトリスト(裁定書7 A-1/A-3。層(a)で自給する)
' ----------------------------
'   ID実在検査は「呼出側が modKnowledge から取得して渡す一覧テキスト」を見る
'   (14章§6)。テストは modKnowledge に触れず、**15章の1行書式リテラル**で
'   小さな一覧を組み立てて引数に渡す。載せるIDは15章§8.1が mock 応答に許した
'   実在ID(M-0012 / L-04 / S-0004 / K-0003 / P9)だけで、幽霊素材に使う
'   M-9999 / L-99 / S-9999 / K-9999 は**1行も含めない**。
'   ・共通規約(15章§6.1): 1行1件・行頭は `[ID] `・項目区切りは " | "・
'     項目内の複数値は ";" 区切り・空項目は項目ごと省略・改行は vbLf。
'   ・MC-0107(機構ライブラリ)は壁打ちのsystemにしか注入されず Check系が読む
'     一覧には現れないため、本ホワイトリストには載せない(15章§6.1)。
'   Public なのは姉妹モジュール modTestsPure5 の G31K/G32K/G34K が同じ一覧を
'   使うため(30,000字契約により実体は本モジュール側に置く)。
' ============================

' S2の {{menusText}} = MenusSummaryFor(15章§3。ID・名称・対応カテゴリのみ)。
Public Function WlMenusSummary() As String
    WlMenusSummary = _
        "[M-0012] 食品工場リスク診断サービス | 対応カテゴリ:manufacturing_quality;supply_chain" & vbLf & _
        "[M-0031] 物流網BCP点検サービス | 対応カテゴリ:supply_chain"
End Function

' S3の {{menusText}} = MenusFor(15章§4。要約版に概要を加えた版)。
Public Function WlMenus() As String
    WlMenus = _
        "[M-0012] 食品工場リスク診断サービス | 概要:製造ラインの異物混入リスクを現地診断する" & _
        " | 対応カテゴリ:manufacturing_quality;supply_chain" & vbLf & _
        "[M-0031] 物流網BCP点検サービス | 概要:調達と配送の途絶点を洗い出す | 対応カテゴリ:supply_chain"
End Function

' S3の {{linesText}} = LinesText(15章§4。market_note が空の種目は項目ごと省略)。
Public Function WlLines() As String
    WlLines = _
        "[L-04] 賠償責任(一般) | 市場環境:再保険料率の上昇で限度額に慎重な傾向があり、案件による個別判断が求められる。" & vbLf & _
        "[L-07] 企業総合賠償責任保険"
End Function

' S3の {{schemesText}} = SchemesFor(15章§4。status が proven/adopted のみ)。
Public Function WlSchemes() As String
    WlSchemes = _
        "[S-0004] 見守りヤモリ型(P2) | 構造:検知パートナーと有事補償のバンドル" & _
        " | 成立条件:検知パートナーの実在;引受条件への接続 | 適用シグナル:設備の老朽化"
End Function

' S3の {{casesText}} = CasesFor(15章§4)。
Public Function WlCases() As String
    WlCases = _
        "[K-0003] 業種:09 顧客像:菓子製造の中堅 提示リスク:異物混入による自主回収" & _
        " 提案:生産物賠償の限度額増額 決め手:回収費用の実額提示"
End Function

' CheckPF の refIdsText(14章§6: menusSummary + patternsText + rulesText +
'   researchingText の連結)。V-PF-03 は M- / S- / K- で始まるIDを見るので、
'   mock が使う S-0004 / K-0003 の行も同じ連結へ載せる。
Public Function WlPfRefIds() As String
    WlPfRefIds = WlMenusSummary() & vbLf & _
        "[P9] 予兆検知と補償のバンドル | 構造:検知サービスとセットで残余リスクを保険がカバー" & _
        " | 成立条件:検知パートナーの実在;検知から引受条件化への接続" & _
        " | 代表例:漏水センサーと水濡れ補償 | 社内実績:見守りヤモリ型" & vbLf & _
        "[J-03] class:adverse_selection 基準:加入者が予兆を知っている設計は引き受けない" & _
        " | 破り方:entry_path" & vbLf & _
        "[RT-07] 高齢者見守り連携 status:researching 判定日:2026-05-20" & _
        " | メモ:自治体予算の裏取り待ち | 関連:P9" & vbLf & _
        WlSchemes() & vbLf & WlCases()
End Function

' ============================
' G40 mock正常応答の自バリアント文脈検証(15章§8.1受入条件1・17章T-22 DoD)
'   「警告判定のケースも発火させない」ため、期待値は空文字ちょうど。
'   受入条件1の「文脈」には**実在ID一覧の注入も含む**(裁定書7 A-2で15章§8.1へ
'   明記)。一覧を渡さない呼び出しは fail-closed で不合格になるので合格判定に
'   使えない。よってID実在検査を通る応答には Wl* の一覧を渡して判定する。
' ============================
Private Sub T_MockOk()
    ChkZero "G40_MK-S1-NEWがnew文脈で警告含め0件_15章§8.1", _
        modValidate.CheckS1(MockJson(MK_S1N), CT_NEW)

    ChkZero "G40_MK-S1-RNWがrenewal文脈で警告含め0件_15章§8.1", _
        modValidate.CheckS1(MockJson(MK_S1R), CT_RNW)

    ChkZero "G40_MK-S2-NEWがnew文脈で警告含め0件_15章§8.1", _
        modValidate.CheckS2(MockJson(MK_S2N), CT_NEW, WlMenusSummary())

    ChkZero "G40_MK-S2-RNWがrenewal文脈で警告含め0件_15章§8.1", _
        modValidate.CheckS2(MockJson(MK_S2R), CT_RNW, WlMenusSummary())

    ChkZero "G40_MK-S4が0件_15章§8.1", modValidate.CheckS4(MockJson(MK_S4))

    ChkZero "G40_MK-PFが0件_15章§8.1", _
        modValidate.CheckPF(MockJson(MK_PF), WlPfRefIds())

    ChkZero "G40_MK-S2C-HITが0件_15章§8.1", _
        modValidate.CheckS2C(MockJson(MK_C2H), MockJson(MK_S2R))

    ChkZero "G40_MK-S2C-CLEANが0件で改訂スキップ_15章§8.1", _
        modValidate.CheckS2C(MockJson(MK_C2C), MockJson(MK_S2R))

    ChkZero "G40_MK-S3C-HITが0件_15章§8.1", modValidate.CheckS3C(MockJson(MK_C3H))

    ChkZero "G40_MK-S3C-CLEANが0件で改訂スキップ_15章§8.1", _
        modValidate.CheckS3C(MockJson(MK_C3C))
End Sub

' ============================
' G40K 共通バリアントのS3はnew・renewalの両文脈で合格すること(受入条件1)
'   ID実在検査を通るためグループを分けてある。15章§8.1が言う「文脈」には
'   実在ID一覧の注入が含まれるので、4つの一覧(menus/lines/schemes/cases)を
'   15章§4の書式で自給して渡す。mockが使う M-0012 / L-04 / S-0004 / K-0003 が
'   一覧に載っているので0件になる=幽霊IDが1件も無いことの裏返しである。
' ============================
Private Sub T_MockOkIds()
    Dim s3 As String
    Dim wm As String
    Dim wn As String
    Dim ws As String
    Dim wc As String

    s3 = MockJson(MK_S3)
    wm = WlMenus()
    wn = WlLines()
    ws = WlSchemes()
    wc = WlCases()

    ChkZero "G40_MK-S3がnew文脈で0件_15章§8.1", _
        modValidate.CheckS3(s3, MockJson(MK_S2N), wm, wn, ws, wc)

    ChkZero "G40_MK-S3がrenewal文脈で0件_15章§8.1", _
        modValidate.CheckS3(s3, MockJson(MK_S2R), wm, wn, ws, wc)
End Sub

' ============================
' G41 障害注入mock(15章§8.2の表)を modValidate へ通したときの観測事実。
'   mock応答の**形**(破損・M-9999混入・stories2件・空・上限文字列・#ERR行)と
'   DecideOk の判定は modTestsPure2 G7/G10 が既に固定しているので、本群は
'   そこに無い「検証器がどう反応するか」だけを見る(同じ答えを2箇所に書かない)。
'   fault指定時に自バリアントの本文へ障害を注入するのはゲートウェイ入口
'   MockResponse(stepName, variantName, fault) の責務(14章§6)なので、
'   自文脈での合格を見る4本はこちらを使う。
' ============================
Private Sub T_Fault()
    Dim r1 As String
    Dim r2 As String

    ' enum_violation: s2 の category に未定義値 quality が入る。
    ChkFault "G41_enum_violationはS2のenum検査を落とす_15章§8.2", _
        modValidate.CheckS2(modJsonLite.ExtractJsonBlock( _
            modMockLlm.FaultResponse("enum_violation", "s2")), CT_NEW), "V-S2-", "03"

    ' count_violation: s3 の stories が2件。
    ChkFault "G41_count_violationはS3の件数検査を落とす_15章§8.2", _
        modValidate.CheckS3(modJsonLite.ExtractJsonBlock( _
            modMockLlm.FaultResponse("count_violation", "s3")), MockJson(MK_S2R)), "V-S3-", "01"

    ' fake_err: 先頭行が偽装エラーでも本文JSONは正常。検証器はそれを合格させる
    '   (エラーUIへ昇格させない根拠。14章§6の帯域外成否規約・16章E-46)。
    ChkZero "G41_fake_errの本文JSONは自文脈で合格する_15章§8.2", _
        modValidate.CheckS2(modJsonLite.ExtractJsonBlock( _
            modMockLlm.MockResponse("s2", CT_NEW, "fake_err")), CT_NEW, WlMenusSummary())

    ' broken_json: 修復呼出にも同じ破損が返る(毎回破損=状態レス)。だから
    '   修復後も不合格でStepが失敗する、という15章§8.2の期待挙動が成立する。
    r1 = modMockLlm.MockResponse("s2", CT_NEW, "broken_json")
    r2 = modMockLlm.MockResponse("s2", CT_NEW, "broken_json")
    ChkB "G41_broken_jsonは修復呼出でも毎回破損する_15章§8.2", _
        (modJsonLite.ExtractJsonBlock(r1) = "" And modJsonLite.ExtractJsonBlock(r2) = ""), _
        "2回目=[" & HeadOf(r2) & "]"

    ' broken_json_once: 2回目(修復呼出)の正常応答が自文脈で合格すること
    '   =修復リトライの成功系が成立する素材であること。
    modMockLlm.ResetFaultOnce
    r1 = modMockLlm.MockResponse("s2", CT_NEW, "broken_json_once")
    r2 = modMockLlm.MockResponse("s2", CT_NEW, "broken_json_once")
    ChkZero "G41_broken_json_onceの2回目は自文脈で合格する_15章§8.2", _
        modValidate.CheckS2(modJsonLite.ExtractJsonBlock(r2), CT_NEW, WlMenusSummary())

    ' 未知の値は正常応答へフォールバックする(config入力ミスでE2Eを暴走させない)。
    ChkZero "G41_未知のfault値は正常応答へフォールバックする_15章§8.2", _
        modValidate.CheckS2(modJsonLite.ExtractJsonBlock( _
            modMockLlm.MockResponse("s2", CT_NEW, "no_such_fault_kind")), CT_NEW, WlMenusSummary())

    ' fault が空のときは8.1の正常応答のみ(既定で異常系が混ざらない)。
    ChkZero "G41_fault空は正常応答のみを返す_15章§8.2", _
        modValidate.CheckS2(modJsonLite.ExtractJsonBlock( _
            modMockLlm.MockResponse("s2", CT_NEW, "")), CT_NEW, WlMenusSummary())
End Sub

' ============================
' G41K ghost_id(ID実在検査を通るためグループを分けてある)
'   **一覧を渡した上で**判定する。渡さないと fail-closed の申告行(裁定書7 A-2)
'   が同じケースIDで出てしまい、「幽霊IDを捕まえた」ことの証明にならない。
'   発火した行が幽霊ID M-9999 を挙げていることまで見て両者を分ける。
' ============================
Private Sub T_FaultIds()
    Dim outText As String

    outText = modValidate.CheckS3(modJsonLite.ExtractJsonBlock( _
        modMockLlm.FaultResponse("ghost_id", "s3")), MockJson(MK_S2R), _
        WlMenus(), WlLines(), WlSchemes(), WlCases())

    ChkB "G41_ghost_idはS3のID実在検査を落とす_15章§8.2", _
        (HasCase(outText, "V-S3-" & "03") And InStr(outText, "M-9999") > 0), _
        "期待=V-S3-" & "03 が M-9999 を挙げて発火。実際=[" & HeadOf(outText) & "]"
End Sub

' ============================
' G42 NormalizeLlmJson(14章§5 防衛線(2.5)・16章E-49。尾部劣化=同名項目の重複)
' ============================
Private Sub T_Normalize()
    Dim dupJson As String
    Dim uniqJson As String
    Dim outText As String
    Dim removed As Long

    dupJson = "{""risks"":[{""risk_no"":1,""risk_name"":""火災""}," & _
              "{""risk_no"":1,""risk_name"":""火災""}]}"
    uniqJson = "{""risks"":[{""risk_no"":1,""risk_name"":""火災""}," & _
               "{""risk_no"":2,""risk_name"":""水災""}]}"

    removed = -1
    outText = modValidate.NormalizeLlmJson("s2", dupJson, removed)
    ChkN "G42_NormalizeLlmJsonが完全重複を1件へ畳む_14章§5", ArrCount(outText, "risks"), 1
    ChkN "G42_NormalizeLlmJsonが除去件数を帯域外で返す_16章E-49", removed, 1

    removed = -1
    outText = modValidate.NormalizeLlmJson("s2", uniqJson, removed)
    ChkN "G42_重複が無ければ除去0件_14章§5", removed, 0
    ChkN "G42_重複が無ければ件数を保つ_14章§5", ArrCount(outText, "risks"), 2
End Sub

' ============================
' G43 SanitizeFileName(13章§2.8のファイル名生成規則1-3)
'   規則4(case_id由来8桁の付与)と規則5(パス240字超の置換)は、現行の1引数
'   宣言では case_id もパスも渡らないため層(a)から一意に固定できない。
'   よって本群は**先頭一致**で規則1-3だけを固定する(付与の有無に依存しない)。
' ============================
Private Sub T_FileName()
    Dim res As String

    res = modUtilText.SanitizeFileName("A/B:C*D?E""F<G>H|I")
    ChkS "G43_SanitizeFileNameが禁止文字を下線へ_13章§2.8", Left$(res, 17), "A_B_C_D_E_F_G_H_I"

    res = modUtilText.SanitizeFileName("A" & Chr(9) & "B")
    ChkS "G43_SanitizeFileNameが制御文字を下線へ_13章§2.8", Left$(res, 3), "A_B"

    res = modUtilText.SanitizeFileName("  浜松スイーツ.  ")
    ChkS "G43_SanitizeFileNameが前後空白と末尾ピリオドを除去_13章§2.8", _
        Left$(res, 6), "浜松スイーツ"

    res = modUtilText.SanitizeFileName(RepChar("あ", 40))
    ChkB "G43_SanitizeFileNameが32字で切詰め_13章§2.8", _
        (Left$(res, 32) = RepChar("あ", 32) And Mid$(res, 33, 1) <> "あ"), _
        "実際=[" & HeadOf(res) & "]"

    res = modUtilText.SanitizeFileName("浜松スイーツファクトリー")
    ChkS "G43_SanitizeFileNameが通常の社名を壊さない_13章§2.8", _
        Left$(res, 12), "浜松スイーツファクトリー"
End Sub

' ============================
' 判定の共通形(modTestsPure5 と同じ作法。テスト名の先頭が15章§11のケースID)
' ============================
Private Function CaseIdOf(ByVal nm As String) As String
    Dim p As Long
    p = InStr(nm, "_")
    If p > 1 Then
        CaseIdOf = Left$(nm, p - 1)
    Else
        CaseIdOf = nm
    End If
End Function

' 1ケース1本の共通形: 壊した素材で発火し、素材のままでは発火しないこと。
Private Sub ChkFire(ByVal nm As String, ByVal ngOut As String, ByVal okOut As String)
    modTestRunner.Check nm, _
        (HasCase(ngOut, CaseIdOf(nm)) And Not HasCase(okOut, CaseIdOf(nm))), _
        "壊した素材の結果=[" & HeadOf(ngOut) & "] 素材のままの結果=[" & HeadOf(okOut) & "]"
End Sub

' ============================
' G32G CheckS3 の growth_ideas(攻めの保険活用。15章§4 V-S3-14 ～ V-S3-18)
' ----------------------------
'   素材は MK-S3(growth_ideas 4件)。5ケースそれぞれに対し「壊した素材で発火し、
'   素材のままでは発火しない」を張る(modTestsPure5 の G32 と同じ形)。
'   growth_ideas の配列は本群だけで自給し(GIdea/WithGrowth)、他の3キーには
'   触れない = 既存 V-S3-01 ～ V-S3-13 の素材を汚さない。
' ============================
Private Sub T_S3Growth()
    Dim s3 As String
    Dim s2r As String
    Dim okOut As String
    Dim baseArr As String

    s3 = MockJson(MK_S3)
    s2r = MockJson(MK_S2R)
    okOut = modValidate.CheckS3(s3, s2r)
    baseArr = "[" & GIdea("成長案A", "5", "low") & "," & GIdea("成長案B", "4", "mid") & _
              "," & GIdea("成長案C", "3", "high") & "," & GIdea("成長案D", "2", "mid") & "]"

    ' 件数が4未満(3件)。
    ChkFire "V-S3-14_growth_ideasが4件未満_15章§11", _
        modValidate.CheckS3(WithGrowth(s3, "[" & GIdea("成長案A", "5", "low") & "," & _
            GIdea("成長案B", "4", "mid") & "," & GIdea("成長案C", "3", "high") & "]"), s2r), okOut

    ChkFire "V-S3-15_growth_ideasのeffectが1から5の外_15章§11", _
        modValidate.CheckS3(WithGrowth(s3, Replace(baseArr, """effect"":5", """effect"":6")), s2r), okOut

    ChkFire "V-S3-16_growth_ideasのdifficultyがenum外_15章§11", _
        modValidate.CheckS3(WithGrowth(s3, Replace(baseArr, """difficulty"":""low""", _
            """difficulty"":""medium""")), s2r), okOut

    ChkFire "V-S3-17_growth_ideasのtitleが30字超か重複_15章§11", _
        modValidate.CheckS3(WithGrowth(s3, Replace(baseArr, "成長案B", "成長案A")), s2r), okOut

    ' stories 先頭の headline と同じ文言を title に置く(同じ案を2箇所に出さない)。
    ChkFire "V-S3-18_growth_ideasのtitleがstoriesのheadlineと重複_15章§11", _
        modValidate.CheckS3(WithGrowth(s3, Replace(baseArr, "成長案A", _
            "PL保険のリコール費用特約拡充")), s2r), okOut
End Sub

' growth_ideas の1件(15章§4 Schema-S3 の6キー)。
Private Function GIdea(ByVal titleText As String, ByVal effectText As String, _
                       ByVal diffText As String) As String
    GIdea = "{""title"":""" & titleText & """,""what"":""何をするのか""," & _
            """why"":""なぜこの会社に効くのか"",""insurance_fit"":""保険との接点""," & _
            """effect"":" & effectText & ",""difficulty"":""" & diffText & """}"
End Function

' S3 JSON の growth_ideas 配列だけを差し替える(他の3キーは1字も触らない)。
Private Function WithGrowth(ByVal js As String, ByVal arrText As String) As String
    Dim p As Long
    p = InStr(js, ",""growth_ideas"":")
    If p = 0 Then
        WithGrowth = js
        Exit Function
    End If
    WithGrowth = Left$(js, p - 1) & ",""growth_ideas"":" & arrText & "}"
End Function
