Attribute VB_Name = "modUICase"
Option Explicit

' ============================================================================
' modUICase - enum変換表・入力規則・案件シート描画の入口(ui層・T-31)
' ----------------------------------------------------------------------------
' 11章§5「日本語ラベル⇔enumの変換は modUICase の共通変換表(19章と一致必須)
' のみで行う」の実体。**変換表はこのモジュールの EnumPairsCsv() 1本だけ**で、
' 画面の入力規則リスト(隠しレンジ)も、シート->JSON の逆変換も、JSON->シートの
' 表示も、すべてここから引く(同じ表を2箇所に書かない)。
'
' 19章§3との一致は目視ではなく `tools/enum_check.py` が機械照合する
' (17章§4-2の一致検査2本目)。同ツールは EnumPairsCsv() の本体を**静的に評価**
' するので、表をループや条件分岐で組み立て直すと照合が落ちる=骨抜きにできない。
'
' 30,000字契約(12章§2)による分割:
'   modUICase   = 変換表・入力規則・匿名化(E-31)の復元/置換・S1～S4の入口
'   modUICase2  = S1～S4シートの描画と逆シリアライズの本体(13章§2.12-§2.15)
'   modUICase3  = 案件入力・フィードバック・判断台帳(13章§2.5/§2.7/§2.11)
'   依存は modUICase2 / modUICase3 -> modUICase の一方向(変換表を引くだけ)。
'
' R4(12章§4): ui層なのでExcelトークンを使ってよい。セル書込は modUISheet 経由で
'   modUtilText.SetCellSafe に集約する。
' ============================================================================

' 入力規則の隠しレンジ(13章§2.9の実行時生成シート)。**列1は modBoot が
' data_key 用に使う**ため、本モジュールは列2以降を使う(書き合いを避ける)。
Private Const UC_ENUM_SHEET As String = "enum_hidden"
Private Const UC_ENUM_FIRST_COL As Long = 2
Private Const UC_ENUM_PREFIX As String = "enum_"
Private Const UC_DV_ROWS As Long = 120
Private Const UC_DV_TITLE As String = "入力できない値です"
Private Const UC_DV_MSG As String = "一覧から選んでください。"
Private Const UC_SRC As String = "modUICase"

' 16章E-31 のプレースホルダ(modExportHtml の復元と同じ2種)。
Public Const UC_PH_COMPANY As String = "{{COMPANY}}"
Public Const UC_PH_POLICY As String = "{{POLICY_NO}}"

' 証券番号らしさの判定(E-31)。英数字とハイフンだけで構成され、8文字以上、
' 数字を4つ以上含み、かつ数字以外を1つ以上含むトークンを候補とする
' (純粋な数字列は金額・年月と紛れるので対象にしない=過剰置換を避ける)。
Private Const UC_POLICY_MIN_LEN As Long = 8
Private Const UC_POLICY_MIN_DIGITS As Long = 4

' 変換表のキャッシュ(先頭に vbLf を足した形。行頭一致で引くため)。
Private gPairs As String

' ============================================================================
' EnumPairsCsv - 19章§3の変換表そのもの(`グループ,機械値,日本語` を vbLf 区切り)
' ----------------------------------------------------------------------------
'   **この関数だけが値の出どころ**。EnumJa / EnumEn / EnumLabels / 入力規則の
'   隠しレンジは、すべてこの戻り値を走査して答える。
'   並び順も19章§3の記載順であり、tools/enum_check.py が順序まで照合する。
'   行を足す・直すときは 19章§3 を直してから `python3 tools/enum_check.py
'   --dump-bas` の出力で本関数を差し替えること(手で写さない)。
' ============================================================================
Public Function EnumPairsCsv() As String
    Dim s As String
    s = s & "case_type,new,新規開拓" & vbLf
    s = s & "case_type,renewal,更新" & vbLf
    s = s & "dossier_tier,t1_quick,かんたん調査" & vbLf
    s = s & "dossier_tier,t2_full,しっかり調査" & vbLf
    s = s & "dossier_tier,t3_sparring,壁打ち" & vbLf
    s = s & "channel,wholesale,ホール" & vbLf
    s = s & "channel,retail,リテール" & vbLf
    s = s & "kanji,lead,幹事" & vbLf
    s = s & "kanji,non_lead,非幹事" & vbLf
    s = s & "kanji,coins,共保" & vbLf
    s = s & "bid,yes,あり" & vbLf
    s = s & "bid,no,なし" & vbLf
    s = s & "reins,none,なし" & vbLf
    s = s & "reins,reins,再保険" & vbLf
    s = s & "reins,captive,キャプティブ" & vbLf
    s = s & "case_status,draft,下書き" & vbLf
    s = s & "case_status,s1_done,S1完了" & vbLf
    s = s & "case_status,s2_done,S2完了" & vbLf
    s = s & "case_status,s3_done,S3完了" & vbLf
    s = s & "case_status,s4_done,S4完了" & vbLf
    s = s & "case_status,exported,出力済" & vbLf
    s = s & "case_status,feedback_done,記録済" & vbLf
    s = s & "case_status,error,エラー" & vbLf
    s = s & "s4_variant,proposal,保険提案書" & vbLf
    s = s & "s4_variant,alliance,協業提案書" & vbLf
    s = s & "quality_mode,standard,標準" & vbLf
    s = s & "quality_mode,deep,入念（批判→改訂つき）" & vbLf
    s = s & "risk_category,strategy_market,戦略・市場" & vbLf
    s = s & "risk_category,supply_chain,調達・供給網" & vbLf
    s = s & "risk_category,manufacturing_quality,製造・品質" & vbLf
    s = s & "risk_category,sales_customer,販売・顧客" & vbLf
    s = s & "risk_category,facility_bcp,施設・自然災害・BCP" & vbLf
    s = s & "risk_category,hr_labor,人材・労務" & vbLf
    s = s & "risk_category,digital_info,デジタル・情報" & vbLf
    s = s & "risk_category,legal_regulatory,法務・規制" & vbLf
    s = s & "risk_category,finance_counterparty,財務・取引先" & vbLf
    s = s & "risk_category,brand_social,ブランド・社会" & vbLf
    s = s & "transferability,cover,比較的移転しやすい" & vbLf
    s = s & "transferability,partial,条件付き・部分的" & vbLf
    s = s & "transferability,hard,保険化困難" & vbLf
    s = s & "risk_status,proposed,仮説" & vbLf
    s = s & "risk_status,confirmed,確認済み" & vbLf
    s = s & "risk_status,rejected,棄却（記録保持）" & vbLf
    s = s & "risk_status,new,新規発見" & vbLf
    s = s & "horizon,already,既に顕在化" & vbLf
    s = s & "horizon,near,1～3年" & vbLf
    s = s & "horizon,mid_long,3年超" & vbLf
    s = s & "frequency,high,高" & vbLf
    s = s & "frequency,mid,中" & vbLf
    s = s & "frequency,low,低" & vbLf
    s = s & "impact,large,大" & vbLf
    s = s & "impact,mid,中" & vbLf
    s = s & "impact,small,小" & vbLf
    s = s & "evidence_source,hp,HP" & vbLf
    s = s & "evidence_source,yuho,有報" & vbLf
    s = s & "evidence_source,memo,営業メモ" & vbLf
    s = s & "evidence_source,contract,現契約" & vbLf
    s = s & "evidence_source,prev_renewal,前回更新メモ" & vbLf
    s = s & "evidence_source,knowledge,社内ナレッジ" & vbLf
    s = s & "evidence_source,inference,推定" & vbLf
    s = s & "gap_type,uninsured,無保険" & vbLf
    s = s & "gap_type,underinsured,過小" & vbLf
    s = s & "gap_type,overlap,重複" & vbLf
    s = s & "proposal_kind,upsell,補償拡大" & vbLf
    s = s & "proposal_kind,cross_sell,新種目提案" & vbLf
    s = s & "proposal_kind,scheme,座組提案" & vbLf
    s = s & "growth_difficulty,low,低" & vbLf
    s = s & "growth_difficulty,mid,中" & vbLf
    s = s & "growth_difficulty,high,高" & vbLf
    s = s & "location_type,工場,工場" & vbLf
    s = s & "location_type,本社,本社" & vbLf
    s = s & "location_type,店舗,店舗" & vbLf
    s = s & "location_type,倉庫,倉庫" & vbLf
    s = s & "location_type,その他,その他" & vbLf
    s = s & "input_quality_aspect,profile,会社概要" & vbLf
    s = s & "input_quality_aspect,business,事業・製品" & vbLf
    s = s & "input_quality_aspect,sites,拠点・設備" & vbLf
    s = s & "input_quality_aspect,history,沿革" & vbLf
    s = s & "input_quality_aspect,news,直近の動き" & vbLf
    s = s & "input_quality_aspect,hr,採用・人員" & vbLf
    s = s & "input_quality_aspect,finance_risk,有報・財務リスク" & vbLf
    s = s & "input_quality_aspect,sales_memo,営業情報" & vbLf
    s = s & "input_quality_aspect,sns,SNS評判" & vbLf
    s = s & "input_quality_aspect,competitors,競合・業界事故" & vbLf
    s = s & "input_quality_aspect,market,市況・マクロ" & vbLf
    s = s & "input_quality_aspect,finance,財務状態" & vbLf
    s = s & "input_quality_aspect,insurance_ctx,付保・提案の経緯" & vbLf
    s = s & "input_quality_aspect,hazard,拠点ハザード" & vbLf
    s = s & "input_quality_status,ok,十分" & vbLf
    s = s & "input_quality_status,partial,断片的" & vbLf
    s = s & "input_quality_status,missing,無い" & vbLf
    s = s & "input_quality_overall,high,高" & vbLf
    s = s & "input_quality_overall,mid,中" & vbLf
    s = s & "input_quality_overall,low,低" & vbLf
    s = s & "field_insight_tag,risk_clue,リスクの手がかり" & vbLf
    s = s & "field_insight_tag,relationship,決裁・人間関係" & vbLf
    s = s & "field_insight_tag,competitor,競合・他社" & vbLf
    s = s & "field_insight_tag,constraint,制約・NG" & vbLf
    s = s & "field_insight_tag,opportunity,商機" & vbLf
    s = s & "field_insight_tag,other,その他" & vbLf
    s = s & "certainty,confirmed,確認済み" & vbLf
    s = s & "certainty,assumed,見立て" & vbLf
    s = s & "financials_source,yuho,有報" & vbLf
    s = s & "financials_source,kessan_kokoku,決算公告" & vbLf
    s = s & "financials_source,tdb,信用調査" & vbLf
    s = s & "financials_source,view,VIEW情報" & vbLf
    s = s & "financials_source,memo,営業メモ" & vbLf
    s = s & "financials_source,unknown,不明" & vbLf
    s = s & "sparring_role,user,自分" & vbLf
    s = s & "sparring_role,ai,AI" & vbLf
    s = s & "inbox_judge_to,adopted,採択" & vbLf
    s = s & "inbox_judge_to,conditional_hold,条件付き保留" & vbLf
    s = s & "inbox_judge_to,rejected,却下" & vbLf
    s = s & "inbox_source_kind,member_post,部内投稿" & vbLf
    s = s & "inbox_source_kind,field_voice,現場の声" & vbLf
    s = s & "inbox_source_kind,watch,ウォッチ" & vbLf
    s = s & "judge_decision,raise,増率" & vbLf
    s = s & "judge_decision,close,クローズ" & vbLf
    s = s & "judge_decision,restrict,縮小・限定" & vbLf
    s = s & "judge_decision,keep,条件維持" & vbLf
    s = s & "judge_decision,improve,ロス改善伴走" & vbLf
    s = s & "judge_result,won,成約" & vbLf
    s = s & "judge_result,lost,失注" & vbLf
    s = s & "judge_result,pending,未確定" & vbLf
    EnumPairsCsv = s
End Function

' 変換表のキャッシュ(毎回組み立て直さない。表そのものは上の1本が唯一の値源)。
Private Function PairsCache() As String
    If LenB(gPairs) = 0 Then gPairs = vbLf & EnumPairsCsv()
    PairsCache = gPairs
End Function

' ============================================================================
' EnumJa - 機械値 -> 日本語ラベル。表に無い組み合わせは ""。
' ============================================================================
Public Function EnumJa(ByVal groupName As String, ByVal enumValue As String) As String
    Dim t As String
    Dim needle As String
    Dim p As Long
    Dim q As Long

    If LenB(groupName) = 0 Then Exit Function
    t = PairsCache()
    needle = vbLf & groupName & "," & Trim$(enumValue) & ","
    p = InStr(1, t, needle, vbBinaryCompare)
    If p = 0 Then Exit Function

    p = p + Len(needle)
    q = InStr(p, t, vbLf, vbBinaryCompare)
    If q = 0 Then q = Len(t) + 1
    EnumJa = Mid$(t, p, q - p)
End Function

' ============================================================================
' EnumEn - 日本語ラベル -> 機械値。表に無いラベルは ""。
' ----------------------------------------------------------------------------
'   13章§2.2「表に無いラベルは検証不合格」を守るため、**推測で近いものを返さ
'   ない**。呼び出し側(逆シリアライズ)は "" を受け取ったら原文をそのまま
'   JSONへ載せ、modValidate に不合格として弾かせる(黙って直さない)。
'
'   唯一の例外(裁定書22 D13): **表示専用の欄に付ける丸括弧の補足**だけは落として
'   もう一度引く。`ci_dossier_tier` は「しっかり調査（貼った内容から自動で決まり
'   ます）」と表示する欄になり、ラベルそのものは変わっていないためである。
'   完全一致で引けたときは**この経路を1度も通らない**ので、既存の変換結果は
'   1つも変わらない(近いものを返す推測にはならない)。
' ============================================================================
Public Function EnumEn(ByVal groupName As String, ByVal labelText As String) As String
    EnumEn = EnumEnExact(groupName, labelText)
    If LenB(EnumEn) > 0 Then Exit Function

    Dim pos As Long
    pos = InStr(1, labelText, "（", vbBinaryCompare)
    If pos <= 1 Then Exit Function
    EnumEn = EnumEnExact(groupName, Left$(labelText, pos - 1))
End Function

' 完全一致だけの引き(上の唯一の値源)。
Private Function EnumEnExact(ByVal groupName As String, ByVal labelText As String) As String
    Dim t As String
    Dim pre As String
    Dim wanted As String
    Dim p As Long
    Dim q As Long
    Dim lineText As String
    Dim c1 As Long
    Dim c2 As Long

    If LenB(groupName) = 0 Then Exit Function
    wanted = Trim$(labelText)
    If LenB(wanted) = 0 Then Exit Function

    t = PairsCache()
    pre = vbLf & groupName & ","
    p = 1
    Do
        p = InStr(p, t, pre, vbBinaryCompare)
        If p = 0 Then Exit Do
        q = InStr(p + 1, t, vbLf, vbBinaryCompare)
        If q = 0 Then q = Len(t) + 1
        lineText = Mid$(t, p + 1, q - p - 1)
        c1 = InStr(1, lineText, ",", vbBinaryCompare)
        If c1 > 0 Then c2 = InStr(c1 + 1, lineText, ",", vbBinaryCompare)
        If c1 > 0 And c2 > c1 Then
            If Mid$(lineText, c2 + 1) = wanted Then
                EnumEnExact = Mid$(lineText, c1 + 1, c2 - c1 - 1)
                Exit Function
            End If
        End If
        p = q
    Loop
End Function

' ============================================================================
' EnumLabels / EnumGroups - 入力規則リストの源泉(11章§5)
' ----------------------------------------------------------------------------
'   EnumLabels は1グループの日本語ラベルを ";" 区切りで返す(ラベルに ";" は
'   含まれないことを tools/enum_check.py が別途検査している)。
' ============================================================================
Public Function EnumLabels(ByVal groupName As String) As String
    Dim t As String
    Dim pre As String
    Dim p As Long
    Dim q As Long
    Dim lineText As String
    Dim c2 As Long
    Dim acc As String

    If LenB(groupName) = 0 Then Exit Function
    t = PairsCache()
    pre = vbLf & groupName & ","
    p = 1
    Do
        p = InStr(p, t, pre, vbBinaryCompare)
        If p = 0 Then Exit Do
        q = InStr(p + 1, t, vbLf, vbBinaryCompare)
        If q = 0 Then q = Len(t) + 1
        lineText = Mid$(t, p + 1, q - p - 1)
        c2 = InStrRev(lineText, ",")
        If c2 > 0 Then
            If LenB(acc) > 0 Then acc = acc & ";"
            acc = acc & Mid$(lineText, c2 + 1)
        End If
        p = q
    Loop
    EnumLabels = acc
End Function

' 変換表に載っている全グループ名を出現順に ";" 区切りで返す。
Public Function EnumGroups() As String
    Dim t As String
    Dim p As Long
    Dim q As Long
    Dim lineText As String
    Dim c1 As Long
    Dim g As String
    Dim acc As String

    t = PairsCache()
    p = 1
    Do
        q = InStr(p + 1, t, vbLf, vbBinaryCompare)
        If q = 0 Then Exit Do
        lineText = Mid$(t, p + 1, q - p - 1)
        c1 = InStr(1, lineText, ",", vbBinaryCompare)
        If c1 > 1 Then
            g = Left$(lineText, c1 - 1)
            If InStr(1, ";" & acc & ";", ";" & g & ";", vbBinaryCompare) = 0 Then
                If LenB(acc) > 0 Then acc = acc & ";"
                acc = acc & g
            End If
        End If
        p = q
    Loop
    EnumGroups = acc
End Function

' ============================================================================
' 入力規則(データの入力規則リスト)の張り直し(11章§5・12章§2.1 手順(4))
' ----------------------------------------------------------------------------
' 11章§5: 「enum列はデータの入力規則(リスト)。源泉はナレッジブック/19章
' レジストリから起動時にコピーした隠しレンジ(ナレッジ未接続でも開ける)」。
' その隠しレンジを **変換表(EnumPairsCsv)から** 起動のたびに作り直す。
' 画面のドロップダウンと逆変換表の出どころが1本になるので、片方だけ古い、
' という状態が起こらない。
'
' 配置: veryHidden シート `enum_hidden` の列2以降に1グループ1列。
'   列1は modBoot が data_key(日本語ラベルを持たない内部キー)に使う。
'
' 束ねる先の表(グループ|種別|対象):
'   n = 帳票型の名前付きレンジ1点(13章§2.10/§2.11)
'   b = テーブル型のブロック列(アンカー名:列物理名。13章§2.12-§2.17)
'   f = フラット表(1行目が物理名ヘッダの1シート1テーブル)の列
'       (シート名:列物理名。13章§2.5-§2.7。ブロックのアンカー名を持たない)
' ============================================================================
Private Function BindingTable() As String
    Dim s As String
    s = s & "case_type|n|ci_case_type" & vbLf
    s = s & "dossier_tier|n|ci_dossier_tier" & vbLf
    s = s & "channel|n|ci_channel" & vbLf
    s = s & "kanji|n|ci_kanji" & vbLf
    s = s & "bid|n|ci_bid" & vbLf
    s = s & "reins|n|ci_reins" & vbLf
    s = s & "s4_variant|n|ci_s4_variant" & vbLf
    s = s & "quality_mode|n|ci_quality_mode" & vbLf
    s = s & "quality_mode|n|hm_quality_mode" & vbLf
    s = s & "input_quality_overall|b|s1_basic:input_quality_overall" & vbLf
    s = s & "location_type|b|s1_locations:type" & vbLf
    s = s & "field_insight_tag|b|s1_field_insights:tag" & vbLf
    s = s & "input_quality_aspect|b|s1_input_quality:aspect" & vbLf
    s = s & "input_quality_status|b|s1_input_quality:status" & vbLf
    s = s & "gap_type|b|s2_gaps:gap_type" & vbLf
    s = s & "risk_category|b|s2_risks:category" & vbLf
    s = s & "risk_status|b|s2_risks:status" & vbLf
    s = s & "frequency|b|s2_risks:frequency" & vbLf
    s = s & "impact|b|s2_risks:impact" & vbLf
    s = s & "evidence_source|b|s2_risks:evidence_source" & vbLf
    s = s & "transferability|b|s2_risks:transferability" & vbLf
    s = s & "risk_category|b|s2_emerging:category" & vbLf
    s = s & "horizon|b|s2_emerging:horizon" & vbLf
    s = s & "evidence_source|b|s2_emerging:evidence_source" & vbLf
    s = s & "proposal_kind|b|s3_stories:proposal_kind" & vbLf
    s = s & "sparring_role|b|sparring_log:role" & vbLf
    ' 裁定書10 M7: 受信箱の判定入力列 judge_to は日本語ラベルで選ばせる
    ' (13章§2.6)。受信箱は columns だけのフラット表なのでブロックのアンカーが
    ' 無く、束ね種別 f で「シート名:列物理名」を指す。この1行が無いと
    ' inbox_judge_to の隠しレンジはどのセルにも束ねられない死んだ登録になる。
    s = s & "inbox_judge_to|f|受信箱:judge_to" & vbLf
    ' 裁定書11 Q4: 投函下書き行の source_kind も利用者が選ぶ入力列なので
    ' 日本語ラベルで束ねる(19章§3 inbox.source_kind)。
    s = s & "inbox_source_kind|f|受信箱:source_kind" & vbLf
    ' 裁定書12 V7: 判断台帳の decision / result も利用者が選ぶ入力列なので
    ' 受信箱と同作法(f種別)で束ねる(19章§3 judgement.decision / judgement.result)。
    ' ビルド時の静的DVは51行目までしか無く、52行目以降は日本語ラベルの一覧が
    ' 消えていた。f種別は2行目以降の列全体へ張るのでこの欠落も同時に解消する。
    s = s & "judge_decision|f|判断台帳:decision" & vbLf
    s = s & "judge_result|f|判断台帳:result" & vbLf
    BindingTable = s
End Function

' ============================================================================
' ApplyEnumValidation - 隠しレンジの複製と入力規則の張り直し。
'   modBoot 手順(4)から呼ぶ(起動のたびに実行して構わない冪等な処理)。
'   戻り値=張れた入力規則の本数(0でも起動は止めない)。
' ============================================================================
Public Function ApplyEnumValidation() As Long
    On Error GoTo Failed

    Dim ws As Object
    Set ws = modUISheet.EnsureHiddenSheet(UC_ENUM_SHEET)
    If ws Is Nothing Then Exit Function

    ' (1) グループごとに1列ずつ複製し、名前付きレンジを張る。
    Dim groups() As String
    groups = Split(EnumGroups(), ";")

    Dim i As Long
    For i = LBound(groups) To UBound(groups)
        If LenB(groups(i)) > 0 Then
            modUISheet.PutEnumRange ws, UC_ENUM_FIRST_COL + i - LBound(groups), _
                                    UC_ENUM_PREFIX & groups(i), _
                                    EnumLabels(groups(i)), ";"
        End If
    Next i

    ' (2) 束ねる先へ張り直す。
    Dim binds() As String
    binds = Split(BindingTable(), vbLf)

    Dim done As Long
    For i = LBound(binds) To UBound(binds)
        If LenB(Trim$(binds(i))) > 0 Then
            If BindOne(binds(i)) Then done = done + 1
        End If
    Next i

    modLog.LogUsage "enum_validation", vbNullString, "bound=" & CStr(done)
    ApplyEnumValidation = done
    Exit Function

Failed:
    modLog.LogError "E0603", UC_SRC & ".ApplyEnumValidation", "bind_failed", Err.Number
    ApplyEnumValidation = 0
End Function

' ============================================================================
' RebindFlatValidation - 1シートぶんのフラット表(f種別)の束ねを張り直す
' ----------------------------------------------------------------------------
' 裁定書11 Q3(a): 受信箱の下書き行を Rows(2).Insert で挿すと、挿入行は直上の
' 見出し行から書式を継ぐため利用者が選ぶ列の入力規則が付かない。行を挿した
' 側からこれを呼び、当該シートの f種別の束ねだけを張り直す(ApplyEnumValidation
' 全体を起動時以外に走らせない)。戻り値=張れた本数。
' ============================================================================
Public Function RebindFlatValidation(ByVal sheetName As String) As Long
    On Error GoTo Zero0
    If LenB(sheetName) = 0 Then Exit Function

    Dim binds() As String
    binds = Split(BindingTable(), vbLf)

    Dim i As Long
    Dim done As Long
    For i = LBound(binds) To UBound(binds)
        If InStr(1, binds(i), "|f|" & sheetName & ":", vbBinaryCompare) > 0 Then
            If BindOne(binds(i)) Then done = done + 1
        End If
    Next i
    RebindFlatValidation = done
    Exit Function
Zero0:
    RebindFlatValidation = 0
End Function

' 束ね1件。書式は BindingTable のコメントのとおり。張れたら True。
Private Function BindOne(ByVal rowText As String) As Boolean
    On Error GoTo Failed

    Dim flds() As String
    flds = Split(rowText, "|")
    If UBound(flds) - LBound(flds) <> 2 Then Exit Function

    Dim groupName As String
    Dim kindText As String
    Dim target As String
    groupName = Trim$(flds(LBound(flds)))
    kindText = Trim$(flds(LBound(flds) + 1))
    target = Trim$(flds(LBound(flds) + 2))

    Dim rangeName As String
    rangeName = UC_ENUM_PREFIX & groupName

    If kindText = "n" Then
        If modUISheet.NamedCell(target) Is Nothing Then Exit Function
        modUISheet.BindNamedValidation target, rangeName, UC_DV_TITLE, UC_DV_MSG
        BindOne = True
        Exit Function
    End If

    Dim parts() As String
    parts = Split(target, ":")
    If UBound(parts) - LBound(parts) <> 1 Then Exit Function

    Dim anchorName As String
    Dim colName As String
    anchorName = parts(LBound(parts))
    colName = parts(LBound(parts) + 1)

    Dim ws As Object
    Dim headerRow As Long
    Dim firstCol As Long

    If kindText = "f" Then
        ' 裁定書10 M7: フラット表は1行目が物理名ヘッダで左端列が1(13章§2.9)。
        ' ブロックのアンカー名を持たないので、シート名から直に引く。
        Set ws = modUISheet.SheetOf(anchorName)
        headerRow = 1
        firstCol = 1
    ElseIf kindText = "b" Then
        Set ws = modUISheet.BlockSheet(anchorName)
        headerRow = modUISheet.BlockRow(anchorName)
        firstCol = modUISheet.BlockCol(anchorName)
    Else
        Exit Function
    End If

    If ws Is Nothing Then Exit Function
    If headerRow <= 0 Or firstCol <= 0 Then Exit Function

    Dim hdr As Variant
    hdr = modUISheet.HeaderOf(ws, headerRow, firstCol, 0)

    Dim colNo As Long
    colNo = modUISheet.ColOf(hdr, firstCol, colName)
    If colNo <= 0 Then Exit Function

    ' 裁定書11 Q5: フラット表(受信箱・判断台帳)は投函・壁打ち・watchが積み上がり
    ' 行数の上限が無い表なので、DVは行2から UC_DV_ROWS までではなく**2行目以降の列全体**
    ' へ張る(121行目以降で日本語ラベルの一覧が消え、手入力が EnumEn の完全一致で
    ' 弾かれる乖離を無くす)。ブロック型(b)は行数が枠で決まるので現状維持。
    Dim dvRows As Long
    dvRows = UC_DV_ROWS
    If kindText = "f" Then dvRows = ColumnRowsBelow(ws, headerRow)

    modUISheet.BindBlockValidation ws, headerRow, colNo, dvRows, _
                                   rangeName, UC_DV_TITLE, UC_DV_MSG
    BindOne = True
    Exit Function

Failed:
    BindOne = False
End Function

' 見出し行の下にある行数(=シート最終行まで)。取れなければ UC_DV_ROWS。
Private Function ColumnRowsBelow(ByVal ws As Object, ByVal headerRow As Long) As Long
    On Error GoTo Fallback0
    ColumnRowsBelow = ws.Rows.count - headerRow
    If ColumnRowsBelow <= 0 Then ColumnRowsBelow = UC_DV_ROWS
    Exit Function
Fallback0:
    ColumnRowsBelow = UC_DV_ROWS
End Function

' ============================================================================
' 匿名化(16章 E-31)のプレースホルダ置換と復元
' ----------------------------------------------------------------------------
' 置換は保存の直前(案件入力 -> case_data)に modUICase3 が、復元は画面表示と
' レポート生成の直前に行う。**LLMには実名と契約情報の組を渡さない**という
' E-31 の要点を、置換と復元を1組にしてここで持つ(規則を2箇所に書かない)。
' ============================================================================

' 実名 -> プレースホルダ。hits に置換件数(会社名+証券番号)を返す。
Public Function AnonymizeText(ByVal sText As String, ByVal company As String, _
                              ByRef hits As Long) As String
    hits = 0
    AnonymizeText = sText
    If LenB(sText) = 0 Then Exit Function

    Dim t As String
    t = sText

    ' (1) 企業名と、その表記ゆれ(法人格の有無)。長い順に消す。
    Dim forms() As String
    forms = Split(CompanyForms(company), vbLf)
    Dim i As Long
    For i = LBound(forms) To UBound(forms)
        If Len(forms(i)) >= 2 Then
            hits = hits + CountOccurrences(t, forms(i))
            t = Replace(t, forms(i), UC_PH_COMPANY)
        End If
    Next i

    ' (2) 証券番号らしき英数列。
    Dim masked As Long
    t = MaskPolicyNumbers(t, masked)
    hits = hits + masked

    AnonymizeText = t
End Function

' プレースホルダ -> 実名(画面表示・レポート生成の直前)。
'   `{{POLICY_NO}}` は復元表を持たないためプレースホルダのまま残す
'   (modExportHtml の復元と同じ扱い。黙って消さない)。
Public Function RestoreNames(ByVal sText As String, ByVal company As String) As String
    RestoreNames = sText
    If LenB(sText) = 0 Then Exit Function
    If LenB(Trim$(company)) = 0 Then Exit Function
    If InStr(1, sText, UC_PH_COMPANY, vbBinaryCompare) = 0 Then Exit Function
    RestoreNames = Replace(sText, UC_PH_COMPANY, company)
End Function

' 企業名の表記ゆれ候補(長い順に vbLf 区切り)。法人格を落とした形も候補にする。
Private Function CompanyForms(ByVal company As String) As String
    Dim base1 As String
    base1 = Trim$(company)
    If LenB(base1) = 0 Then Exit Function

    Dim short1 As String
    short1 = base1
    short1 = Replace(short1, "株式会社", vbNullString)
    short1 = Replace(short1, "有限会社", vbNullString)
    short1 = Replace(short1, "合同会社", vbNullString)
    short1 = Trim$(short1)

    If LenB(short1) > 0 And short1 <> base1 Then
        CompanyForms = base1 & vbLf & short1
    Else
        CompanyForms = base1
    End If
End Function

' 出現回数。needle が空なら0(無限ループにしない)。
Private Function CountOccurrences(ByVal haystack As String, ByVal needle As String) As Long
    Dim n As Long
    Dim p As Long
    If LenB(needle) = 0 Then Exit Function
    p = InStr(1, haystack, needle, vbBinaryCompare)
    Do While p > 0
        n = n + 1
        p = InStr(p + Len(needle), haystack, needle, vbBinaryCompare)
    Loop
    CountOccurrences = n
End Function

' 証券番号らしき英数列を {{POLICY_NO}} へ置換する(E-31)。
'   条件は本モジュール冒頭の定数のとおり。純粋な数字列は対象にしない
'   (金額・年月と紛れるため。過剰置換で本文を壊さないほうを採る)。
Private Function MaskPolicyNumbers(ByVal sText As String, ByRef maskedCount As Long) As String
    maskedCount = 0

    Dim outText As String
    Dim token As String
    Dim ch As String
    Dim i As Long

    For i = 1 To Len(sText) + 1
        If i <= Len(sText) Then
            ch = Mid$(sText, i, 1)
        Else
            ch = vbNullString
        End If

        If IsTokenChar(ch) Then
            token = token & ch
        Else
            If LooksLikePolicyNo(token) Then
                outText = outText & UC_PH_POLICY
                maskedCount = maskedCount + 1
            Else
                outText = outText & token
            End If
            token = vbNullString
            outText = outText & ch
        End If
    Next i

    MaskPolicyNumbers = outText
End Function

' 英数字(半角)とハイフンだけをトークン構成文字とみなす。
Private Function IsTokenChar(ByVal ch As String) As Boolean
    If LenB(ch) = 0 Then Exit Function
    Dim c As Long
    c = AscW(ch)
    If c >= 48 And c <= 57 Then
        IsTokenChar = True
    ElseIf c >= 65 And c <= 90 Then
        IsTokenChar = True
    ElseIf c >= 97 And c <= 122 Then
        IsTokenChar = True
    ElseIf ch = "-" Then
        IsTokenChar = True
    End If
End Function

Private Function LooksLikePolicyNo(ByVal token As String) As Boolean
    If Len(token) < UC_POLICY_MIN_LEN Then Exit Function

    Dim digits As Long
    Dim others As Long
    Dim i As Long
    Dim c As Long
    For i = 1 To Len(token)
        c = AscW(Mid$(token, i, 1))
        If c >= 48 And c <= 57 Then
            digits = digits + 1
        Else
            others = others + 1
        End If
    Next i
    If digits < UC_POLICY_MIN_DIGITS Then Exit Function
    If others < 1 Then Exit Function
    LooksLikePolicyNo = True
End Function

' ============================================================================
' S1～S4シートの入口(13章§2.2 が名指しする `modUICase.SerializeSheet`)
' ----------------------------------------------------------------------------
' 実体は modUICase2(30,000字契約による分割先)。13章が名指しした名前を
' この層に残し、呼び出し側がどちらを呼ぶか迷わないようにする。
' ============================================================================

' 13章§2.2 逆シリアライズ規約1。対象シートの全ブロックを読んでJSONを組む。
'   組めなければ ""(捏造しない)。
Public Function SerializeSheet(ByVal stepNo As Long) As String
    SerializeSheet = modUICase2.SerializeStep(stepNo)
End Function

' JSON -> シート。参照優先の解決は modCaseStore.ResolveStepJson が行う。
Public Function DrawSheet(ByVal caseId As String, ByVal stepNo As Long) As Boolean
    DrawSheet = modUICase2.DrawStep(caseId, stepNo)
End Function
