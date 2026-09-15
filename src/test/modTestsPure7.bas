Attribute VB_Name = "modTestsPure7"
Option Explicit

' ============================
' modTestsPure7 - modPii(16章E-05)と modPipeline 判定核(14章§6)の回帰網
' ----------------------------
' テスト本数: 47本 = G50 modPii 22 / G60 modPipeline判定核 25
' ----------------------------
' 役割: 裁定書7 D-13 / D-14。**実装を1行も読まずに**、14章§6の公開契約と
'   16章E-03 / E-05 / E-07・15章§0.7 だけを根拠に入出力を固定する。
'   W2bの時点でこの2モジュールは回帰網が0本だった(命名権が未確定でPrivate
'   縛りだったため)。裁定書7 B-5 / B-6 で公開名が確定したので独立に張る。
'     G50 modPii 22本 - 検知規則(@付き / 電話番号 / 敬称つき人名)と
'         「本文を返さない」(16章 NFR-S3)の境界。
'     G60 modPipeline 判定核 25本 - 打切り計画・修復要否・結果分類・
'         失敗コードの切分け・予算配分・deep分岐。
'
' 結線(統合済み。本ファイル単体では1本も実行されないため3点セットで結線する):
'   (1) modTestsPure6.RunAll の末尾から modTestsPure7.RunAll を呼ぶ【結線済み】
'       (modTestsPure -> 2 -> 3 -> 4 -> 5 -> 6 -> 7 の数珠つなぎ)。
'   (2) build/modules.json へ1件追加(modTestsPure7。src/test/ 配下・
'       role=test・type=std・wave=T-24)【登録済み】。
'   (3) wintest/tests_expected.txt へ本ファイルの47本を足す【363 -> 410 済み。
'       363 は裁定書7 A-3(隔離グループ書き直し)後の実本数で、各 modTestsPure*
'       のヘッダ宣言 45/70/69/80/70/29/47 の総和と一致する】。
'   run_lo_tests.py の PURE_ALLOWLIST には modTestsPure7 / modPii /
'   modPipeline が既にある(裁定書7 B-6でmodPipelineを登録済み)。
'
' 判定の形: シート・config・ログ・LLMに触れない純関数だけを叩く(R4)。
'   期待値が仕様から一意に決まらない点は**テストを甘くせず**、決まる側面
'   (単調性・部分文字列の有無・境界)で固定し、その旨をコメントに残した。
'
' 設計判断(R4準拠): Excelトークン不使用。改行は vbLf 基準。乱数・時刻不使用。
'   PII素材は実在しない値で組む(example.co.jp はJPRSの文書用予約ドメイン、
'   電話番号は加入者番号を 0 で埋めた形)。
' ============================

' ----------------------------
' G50 の素材(16章E-05の3規則ぶん。1素材1規則に絞って境界を濁らせない)
' ----------------------------
Private Const PII_MAIL As String = "test.user@example.co.jp"
Private Const PII_TEL As String = "053-000-0000"
Private Const PII_TEL_FLAT As String = "0530000000"
Private Const PII_PERSON As String = "田中"

' 検知ゼロの一般業務文(@なし・電話書式なし・敬称なし)。金額と年月は
' 「数字があるだけ」で電話番号と見なしてはならない境界の素材でもある。
Private Const TXT_CLEAN As String = _
    "浜松第二工場のライン増設を2026年10月に予定しています。" & _
    "増設費は3,200万円の見込みで、稟議は総務部で回覧中です。"

' 走査箇所の申告(ScanReport の whereNote。13章§2.2の欄名を想定)。
Private Const WHERE_NOTE As String = "ナビ:現場メモ"

' ----------------------------
' G60 の素材(15章§0 原則10の書式で書いた検証エラー行)
'   ID幻覚のケースID(16章E-07)とそれ以外を1本ずつ持つ。
' ----------------------------
Private Const ERR_GHOST As String = "[V-S3-03] story_no 1 の menu_id M-9999 は実在しません"
Private Const ERR_PLAIN As String = "[V-S2-04] risk_no 1 の evidence.quote が空です"

' ----------------------------
' RunAll: グループ単位で隔離実行する。1グループが実行時エラーで落ちても
'   残りのグループは走る(未実装/未注入の事実は GroupFail で1件の失敗として
'   可視化し、無かったことにしない)。
' ----------------------------
Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 2
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
    ' 数珠つなぎ: modTestsPure -> 2 -> 3 -> 4 -> 5 -> 6 -> 7 -> 8(裁定書8)。
    On Error Resume Next
    Err.Clear
    modTestsPure8.RunAll
    If Err.Number <> 0 Then
        GroupFail "modTestsPure8.RunAll"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub RunGroup(ByVal grpNo As Long, ByRef grpName As String)
    Select Case grpNo
    Case 1
        grpName = "G50 modPii"
        T_Pii
    Case 2
        grpName = "G60 modPipeline判定核"
        T_Core
    End Select
End Sub

' ============================
' 共通ヘルパ(modTestsPure3/4/5/6 と同じ作法。各 modTestsPure* が自前で持つ)
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

Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), _
        "期待=[" & HeadOf(want) & "] 実際=[" & HeadOf(act) & "]"
End Sub

Private Sub ChkN(ByVal nm As String, ByVal act As Long, ByVal want As Long)
    modTestRunner.Check nm, (act = want), "期待=" & want & " 実際=" & act
End Sub

Private Sub ChkT(ByVal nm As String, ByVal act As Boolean)
    modTestRunner.Check nm, act, "True を期待。実際=False"
End Sub

Private Sub ChkF(ByVal nm As String, ByVal act As Boolean)
    modTestRunner.Check nm, Not act, "False を期待。実際=True"
End Sub

Private Sub ChkB(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check nm, cond, detail
End Sub

' ============================
' G50 modPii(16章E-05・14章§6・NFR-S3)
'   検知規則は「@付き / 電話番号 / 敬称つき人名」の3本。1素材1規則で当て、
'   一般業務文で誤検知しないことを対にして張る(検知器は「全部PII」と答えても
'   通ってしまうため、非検知側が無いと空虚なテストになる)。
' ============================
Private Sub T_Pii()
    Dim mailText As String
    Dim telText As String
    Dim telFlatText As String
    Dim bothText As String
    Dim samaText As String
    Dim shiText As String
    Dim rep As String
    Dim masked As String

    mailText = "ご連絡は " & PII_MAIL & " までお願いします。"
    telText = "工場の代表番号は " & PII_TEL & " です。"
    telFlatText = "工場の代表番号は " & PII_TEL_FLAT & " です。"
    bothText = "ご連絡は " & PII_MAIL & " または " & PII_TEL & " まで。"
    samaText = "担当の" & PII_PERSON & "様に工程表をお送りください。"
    shiText = "現場責任者は鈴木氏です。"

    ' ---- HasPii: 3規則それぞれで検知し、一般業務文では検知しない ----
    ChkT "G50_メールアドレスをPIIとして検知する_16章E-05", modPii.HasPii(mailText)

    ChkT "G50_市外局番形式のハイフンつき電話番号を検知する_16章E-05", modPii.HasPii(telText)

    ChkT "G50_ハイフン無しの電話番号を検知する_16章E-05", modPii.HasPii(telFlatText)

    ChkT "G50_敬称_様_つきの人名を検知する_16章E-05", modPii.HasPii(samaText)

    ' 裁定書7 D-13が例示する敬称「氏」。16章E-05の本文は「様」だけを挙げて
    ' いるので、ここが落ちたら仕様(E-05)と裁定書のどちらが正かを司令塔へ返す。
    ' (W2b時点の実測: modPii はカタカナ姓+様しか拾わず、漢字姓+様/氏/さんを
    '  取りこぼす。FR上いちばん多い「田中様」が素通りするので上の1本と併せて
    '  modPii側の修正を要する。テストは16章E-05の文言どおりに置く。)
    ChkT "G50_敬称_氏_つきの人名を検知する_16章E-05", modPii.HasPii(shiText)

    ' 金額(3,200万円)や年月(2026年10月)を電話番号と誤検知しないこと。
    ChkF "G50_数字を含む一般業務文では検知しない_16章E-05", modPii.HasPii(TXT_CLEAN)

    ' ---- DetectionCount: 件数 ----
    ChkN "G50_検知ゼロの本文は件数0_14章§6", modPii.DetectionCount(TXT_CLEAN), 0

    ChkN "G50_メール1件の本文は件数1_14章§6", modPii.DetectionCount(mailText), 1

    ' 種別の異なる2件が別々に数えられること(同じ本文を1件に丸めない)。
    ChkN "G50_メールと電話が同居する本文は件数2_14章§6", modPii.DetectionCount(bothText), 2

    ' ---- KindsOf: 種別("mail;phone" 形式。14章§6の例が語彙の正) ----
    ChkB "G50_KindsOfはメールをmailとして返す_14章§6", _
        (InStr(modPii.KindsOf(mailText), "mail") > 0 And _
         InStr(modPii.KindsOf(mailText), "phone") = 0), _
        "実際=[" & modPii.KindsOf(mailText) & "]"

    ChkB "G50_KindsOfは電話をphoneとして返す_14章§6", _
        (InStr(modPii.KindsOf(telText), "phone") > 0 And _
         InStr(modPii.KindsOf(telText), "mail") = 0), _
        "実際=[" & modPii.KindsOf(telText) & "]"

    ChkB "G50_KindsOfは複数種別をセミコロン区切りで返す_14章§6", _
        (InStr(modPii.KindsOf(bothText), "mail") > 0 And _
         InStr(modPii.KindsOf(bothText), "phone") > 0 And _
         InStr(modPii.KindsOf(bothText), ";") > 0), _
        "実際=[" & modPii.KindsOf(bothText) & "]"

    ChkS "G50_検知ゼロならKindsOfは空_14章§6", modPii.KindsOf(TXT_CLEAN), ""

    ' ---- ScanReport: 箇所名と種別だけを返し、本文は返さない(NFR-S3) ----
    rep = modPii.ScanReport(bothText, WHERE_NOTE)

    ChkB "G50_ScanReportは走査箇所の申告を含む_14章§6", _
        (InStr(rep, WHERE_NOTE) > 0), "実際=[" & HeadOf(rep) & "]"

    ChkB "G50_ScanReportは検知種別を含む_14章§6", _
        (InStr(rep, "mail") > 0 And InStr(rep, "phone") > 0), _
        "実際=[" & HeadOf(rep) & "]"

    ' err_log / dossier_meta へそのまま書ける=本文の断片が1つも残らないこと。
    ChkB "G50_ScanReportはメール本文を残さない_16章NFR-S3", _
        (InStr(rep, PII_MAIL) = 0 And InStr(rep, "test.user") = 0 And _
         InStr(rep, "example.co.jp") = 0), "実際=[" & HeadOf(rep) & "]"

    ChkB "G50_ScanReportは電話番号本文を残さない_16章NFR-S3", _
        (InStr(rep, PII_TEL) = 0 And InStr(rep, PII_TEL_FLAT) = 0), _
        "実際=[" & HeadOf(rep) & "]"

    ChkB "G50_検知ゼロなら種別をひとつも申告しない_14章§6", _
        (InStr(modPii.ScanReport(TXT_CLEAN, WHERE_NOTE), "mail") = 0 And _
         InStr(modPii.ScanReport(TXT_CLEAN, WHERE_NOTE), "phone") = 0), _
        "実際=[" & HeadOf(modPii.ScanReport(TXT_CLEAN, WHERE_NOTE)) & "]"

    ' ---- MaskText: 伏字案(16章E-05(5) の customer_quote 差し替え案) ----
    masked = modPii.MaskText(samaText)
    ChkB "G50_MaskTextは人名を伏字へ置換する_16章E-05", _
        (InStr(masked, PII_PERSON) = 0 And InStr(masked, "{{") > 0), _
        "実際=[" & HeadOf(masked) & "]"

    masked = modPii.MaskText(mailText)
    ChkB "G50_MaskTextはメール本文を残さない_16章E-05", _
        (InStr(masked, PII_MAIL) = 0 And InStr(masked, "{{") > 0), _
        "実際=[" & HeadOf(masked) & "]"

    ' 置換は検知箇所だけ。周りの業務文まで消してしまうと差し替え案にならない。
    ChkB "G50_MaskTextは非検知部分を保つ_16章E-05", _
        (InStr(masked, "ご連絡は") > 0 And InStr(masked, "までお願いします。") > 0), _
        "実際=[" & HeadOf(masked) & "]"

    ChkS "G50_検知ゼロの本文はMaskTextで変わらない_16章E-05", _
        modPii.MaskText(TXT_CLEAN), TXT_CLEAN
End Sub

' ============================
' G60 modPipeline の判定核(14章§6の16本のうち、規約そのものを持つ9本)
'   16章E-03(打切り)・E-07(ID幻覚)・15章§0.7(予算配分)・15章§4.5(deep)。
' ============================
Private Sub T_Core()
    T_CoreRepair
    T_CoreTrim
    T_CoreBudget
End Sub

' ---- 修復要否・結果分類・失敗コード(§5防衛線(4)・13章§2.4・16章E-07) ----
Private Sub T_CoreRepair()
    ' NeedsRepair: 検証エラーがあり、かつ再試行枠が残るときだけ True。
    ChkT "G60_検証エラーがあり再試行枠が残れば修復する_14章§5", _
        modPipeline.NeedsRepair(ERR_PLAIN, 1)

    ChkF "G60_検証エラーが無ければ修復しない_14章§5", _
        modPipeline.NeedsRepair("", 1)

    ChkF "G60_再試行枠が尽きていれば修復しない_14章§5", _
        modPipeline.NeedsRepair(ERR_PLAIN, 0)

    ' ClassifyResult: run_log の validate_result の3値(13章§2.4)。
    ChkS "G60_初回合格はokに分類する_13章§2.4", _
        modPipeline.ClassifyResult(True, False, False), "ok"

    ChkS "G60_修復後の合格はrepairedに分類する_13章§2.4", _
        modPipeline.ClassifyResult(False, True, True), "repaired"

    ChkS "G60_修復しても不合格はfailedに分類する_13章§2.4", _
        modPipeline.ClassifyResult(False, True, False), "failed"

    ' FailCodeOf: ID幻覚を含めば E0301、それ以外は E0302(16章E-07)。
    ChkS "G60_ID幻覚の不合格はE0301へ切り分ける_16章E-07", _
        modPipeline.FailCodeOf(ERR_GHOST), "E0301"

    ChkS "G60_ID幻覚でない不合格はE0302へ切り分ける_16章E-06", _
        modPipeline.FailCodeOf(ERR_PLAIN), "E0302"

    ' 1行でもID幻覚が混ざれば停止側(E0301)へ倒す。黙殺除去の禁止と対になる規約。
    ChkS "G60_ID幻覚が1行でも混ざればE0301へ倒す_16章E-07", _
        modPipeline.FailCodeOf(ERR_PLAIN & vbLf & ERR_GHOST), "E0301"

    ' CaseIdOfLine: 行頭の [ケースID] だけを取り出し、不一致は "" を返す
    ' (15章§0 原則10)。2面を1本で見る(FailCodeOf の判定はこの値に依るため)。
    ChkB "G60_行頭のケースIDだけを取り出す_15章§0", _
        (modPipeline.CaseIdOfLine(ERR_GHOST) = "V-S3-" & "03" And _
         modPipeline.CaseIdOfLine("story_no 1 の menu_id は実在しません") = ""), _
        "実際=[" & modPipeline.CaseIdOfLine(ERR_GHOST) & "]"
End Sub

' ---- 入力打切りの計画(16章E-03(2)(3)(4)) ----
Private Sub T_CoreTrim()
    Dim lens(0 To 5) As Long
    Dim plan As Variant
    Dim lo As Long

    ' lens(0..4)=切る順(HP/有報/追加ドシエ/前回更新メモ/営業メモ)の現在字数、
    ' lens(5)=打切らない4欄の合計(16章E-03(3))。裁定書47 G-1で順を反転。
    lens(0) = 1000
    lens(1) = 1000
    lens(2) = 1000
    lens(3) = 1000
    lens(4) = 1000
    lens(5) = 1000

    ' 合計6,000字が予算10,000字に収まる=1欄も削らない。
    plan = modPipeline.TrimInputPlan(lens, 10000)
    lo = LBound(plan)
    ChkB "G60_上限内なら1欄も削らない_16章E-03", _
        (plan(lo) = 1000 And plan(lo + 1) = 1000 And plan(lo + 2) = 1000 And _
         plan(lo + 3) = 1000 And plan(lo + 4) = 1000), _
        "実際=[" & PlanText(plan) & "]"

    ' budgetChars<=0 は上限なし(config未設定で全部削る事故を起こさない)。
    plan = modPipeline.TrimInputPlan(lens, 0)
    lo = LBound(plan)
    ChkB "G60_予算0は上限なしとして扱う_16章E-03", _
        (plan(lo) = 1000 And plan(lo + 4) = 1000), "実際=[" & PlanText(plan) & "]"

    ' 予算5,500字。先頭(HP)を削れば足りるので、そこだけが減り、
    ' 後ろの営業メモ(lens(4))は満額で残る=打切り順が守られていること。
    plan = modPipeline.TrimInputPlan(lens, 5500)
    lo = LBound(plan)
    ChkB "G60_1欄で足りるときはHPだけを削る_16章E-03", _
        (plan(lo) < 1000 And plan(lo + 1) = 1000 And plan(lo + 2) = 1000 And _
         plan(lo + 3) = 1000 And plan(lo + 4) = 1000), _
        "実際=[" & PlanText(plan) & "]"

    ' 予算4,500字。HPを全部落としても足りないので有報まで
    ' 及ぶ。追加ドシエより後ろ(lens(2..4))は満額で残る。
    plan = modPipeline.TrimInputPlan(lens, 4500)
    lo = LBound(plan)
    ChkB "G60_足りなければ次の欄へ順に及ぶ_16章E-03", _
        (plan(lo) = 0 And plan(lo + 1) < 1000 And plan(lo + 2) = 1000 And _
         plan(lo + 3) = 1000 And plan(lo + 4) = 1000), _
        "実際=[" & PlanText(plan) & "]"

    ' 打切らない4欄(lens(5))だけで予算を超える場合。ここで**仕様が一意に
    ' 決めている**のは次の3点だけなので、それだけを固定する。
    '   (i)  戻り値は5要素(打切らない4欄は計画に載らない=不可侵。14章§6)
    '   (ii) 計画が元の字数を増やすことはない
    '   (iii)E0102の条件(ProtectedOverBudget)が立つ(16章E-03(4))
    ' 5欄を0まで使い切るのか1字も触らないのかは、16章E-03(4)「打切らず、
    ' 実行前にE0102を警告」と14章§6「4欄だけで超過する場合は自動では削らない」
    ' の読みが割れる。どちらかに倒すと仕様に無い期待値を作るため固定しない。
    lens(5) = 20000
    plan = modPipeline.TrimInputPlan(lens, 10000)
    lo = LBound(plan)
    ChkB "G60_保護4欄だけで超過しても4欄は計画に載らずE0102が立つ_16章E-03", _
        (UBound(plan) - lo + 1 = 5 And _
         plan(lo) <= 1000 And plan(lo + 1) <= 1000 And plan(lo + 2) <= 1000 And _
         plan(lo + 3) <= 1000 And plan(lo + 4) <= 1000 And _
         modPipeline.ProtectedOverBudget(lens(5), 10000)), _
        "実際=[" & PlanText(plan) & "]"

    ' ProtectedOverBudget: E0102警告の唯一の条件(16章E-03(4))。
    ChkT "G60_保護4欄が上限を超えていればE0102の条件を満たす_16章E-03", _
        modPipeline.ProtectedOverBudget(10001, 10000)

    ' 上限ちょうどは「超過」ではない(境界で警告を出さない)。
    ChkF "G60_保護4欄が上限ちょうどなら超過ではない_16章E-03", _
        modPipeline.ProtectedOverBudget(10000, 10000)

    ' CutOrderLabels/CutOrderKeys: 切る順を機械で固定する窓(裁定書47 G-1)。
    ' 順を1つでも入れ替えたら赤になる(位置とラベル・data_keyの対応の逐語一致)。
    ChkS "G60_切る順のラベルはHP有報ドシエ前回更新メモ営業メモ_裁定書47G-1", _
        modPipeline.CutOrderLabels(), "hp|yuho|dossier|prev_renewal|memo"

    ChkS "G60_切る順のdata_keyはHP有報ドシエ前回更新メモ営業メモ_裁定書47G-1", _
        modPipeline.CutOrderKeys(), _
        "input_hp|input_yuho|input_dossier|input_prev_renewal|input_memo"
End Sub

' ---- 予算配分・切詰め注記・入念モード(15章§0.7・16章E-03(6)・15章§4.5) ----
Private Sub T_CoreBudget()
    Dim src As String
    Dim res As String

    src = RepChar("あ", 100)

    ' TruncField: allowedChars<=0 は注記だけを返す(本文を1字も残さない)。
    res = modPipeline.TruncField(src, 0)
    ChkB "G60_許容0字なら注記だけを返す_16章E-03", _
        (InStr(res, "一部省略") > 0 And InStr(res, "あ") = 0), _
        "実際=[" & HeadOf(res) & "]"

    ' 切詰めた欄には必ず注記が付く。先頭は元の本文のまま。
    res = modPipeline.TruncField(src, 10)
    ChkB "G60_切詰めた欄には一部省略の注記が付く_16章E-03", _
        (Left$(res, 10) = Left$(src, 10) And InStr(res, "一部省略") > 0 And _
         Len(res) < Len(src)), "実際=[" & HeadOf(res) & "]"

    ' BudgetOf: 15章§0.7 の配分(貼付=上限の7割 / ナレッジ=3割)。
    ChkN "G60_貼付の予算は上限の7割_15章§0.7", modPipeline.BudgetOf(40000, 7), 28000

    ChkN "G60_ナレッジの予算は上限の3割_15章§0.7", modPipeline.BudgetOf(40000, 3), 12000

    ' 7割と3割を足すと上限ちょうど(足して1.3倍になる読み方を封じる。16章E-03(1))。
    ChkN "G60_7割と3割の和は上限ちょうど_16章E-03", _
        modPipeline.BudgetOf(100000, 7) + modPipeline.BudgetOf(100000, 3), 100000

    ' DeepEnabled: 入念モードの分岐はS2/S3だけ(15章§4.5-4.7)。
    ChkB "G60_入念モードはS2とS3で有効_15章§4.5", _
        (modPipeline.DeepEnabled("deep", 2) And modPipeline.DeepEnabled("deep", 3)), _
        "S2=" & modPipeline.DeepEnabled("deep", 2) & " S3=" & modPipeline.DeepEnabled("deep", 3)

    ChkB "G60_入念モードでもS1とS4は対象外_15章§4.5", _
        (Not modPipeline.DeepEnabled("deep", 1) And Not modPipeline.DeepEnabled("deep", 4)), _
        "S1=" & modPipeline.DeepEnabled("deep", 1) & " S4=" & modPipeline.DeepEnabled("deep", 4)

    ChkF "G60_standardではS2でも入念にしない_15章§4.5", _
        modPipeline.DeepEnabled("standard", 2)
End Sub

' 打切り計画の失敗時に読める形へ(5要素を "/" で連ねるだけ)。
Private Function PlanText(ByVal plan As Variant) As String
    Dim i As Long
    Dim s As String
    On Error Resume Next
    For i = LBound(plan) To UBound(plan)
        If Len(s) > 0 Then s = s & "/"
        s = s & plan(i)
    Next i
    On Error GoTo 0
    PlanText = s
End Function
