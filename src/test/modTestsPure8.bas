Attribute VB_Name = "modTestsPure8"
Option Explicit

' ============================
' modTestsPure8 - W2c(T-25/T-27/T-28)の判定核の契約テスト(17章§4-1 層(a))
' ----------------------------
' テスト本数: 57本 = G70 13 / G72 6 / G73 7 / G74 6 / G76 9 / G77 5 /
'                    G78 6 / G79 5
' ----------------------------
' 役割: 裁定書8-11。**実装を1行も読まずに**、13章§2.6/§2.7・15章§4.5-4.7/
'   §6.5/§8.1・16章 E-35/E-36/E-41/E-44 だけを根拠に入出力を固定した独立
'   テスト。原版は T-25/T-27/T-28 の実装前に書かれている。
'     G70 統制語彙の必須化        16章E-41(受信箱の判定を保存してよいか)
'     G72 プリフライト要約列      13章§2.6 pf_survival(15章§6 Schema-PF)
'     G73 履歴の切詰め            16章E-44 sparring_max_turns の境界
'     G74 履歴の線上形式           14章§6 CallChat の";;;"連結(新しい順)
'     G76 改訂要否                16章E-35・V-S2C-05/V-S3C-05
'     G77 改訂の採否              16章E-36
'     G78 批判ダイジェスト        15章§4.7 {{critiqueDigest}}
'     G79 E-35/E-36の分岐合成
'
' 【統合時の§6整合(統合者)】原版は 14章§6 がまだ1本も宣言していない時点で
'   書かれたため、テスト側から3本の**契約提案**(HasControlledVocab /
'   ShouldReviseOf / AdoptRevisionOf)と2本の配置提案(PreflightVerdictOf を
'   modInboxStore へ / CanContinueSparring を帯域外ok判定に)を置いていた。
'   命名権は14章§6であり(原版ヘッダも「§6が別の形を宣言したら§6が正」と
'   明記)、裁定書8 B-7/B-9/B-10 で §6 が別の形を宣言済みなので、
'   **期待値は1件も動かさず**呼び先だけを §6 の宣言へ寄せた:
'     HasControlledVocab(st,dt,rt,dueText)
'         -> modInboxStore.JudgementError(st,dt,rt,hasDue) が ""(=保存可)か
'     PreflightVerdictOf(pfJson)   -> modPlayOps.PfSurvivalOf(pfJson)
'     TrimHistoryOf(hist,max)      -> modSparring.TrimHistoryOf(同型・変更なし)
'     HistoryJoinOf(hist)          -> modSparring.HistoryJoinOf(hist, maxTurns)
'     ShouldReviseOf(n,json,ok)    -> modPipeline2.DeepOutcomeOf(E-35の分類)
'                                     + modPipeline2.NeedsRevision(json,n)
'     AdoptRevisionOf(base,rev,ok) -> modPipeline2.DeepOutcomeOf(E-36の分類)
'     CritiqueDigestOf(json)       -> modPipeline2.CritiqueDigest(json, stepNo)
'   合成が要るものは下の「§6整合アダプタ」に private で置いた(ChkS/Ctn と
'   同じテスト側の私的ヘルパで、公開名は1つも増やしていない)。
'   §6に対応が無く**落とした**ぶんは司令塔への懸念として返した(甘い期待値へ
'   書き換えて実装を追認するより、無い事実を表に出す):
'     - G71 関心度集計 6本: 件数集計と降順整列の純関数が§6に無い
'       (modInboxStore.InterestText はシートI/O。InterestKeyOf /
'        FmtInterestLine は1件分の正規化と整形しか持たない)
'     - G75 壁打ち継続可否 4本: §6の CanContinueSparring は
'       (caseIdText, utterance, hasPii) の**送信前** fail-closed 判定で、
'       原版が仮定した(ok, errCode)の**応答後**の継続可否とは別の契約
'     - G78 の executive_reactions 全件保持 1本: 15章§4.7 の注記は3配列を
'       名指しするが、14章§6は「S3は lands=false の反応 -> issues の順」と
'       宣言していて lands=true の反応は digest に載らない(仕様間の食い違い)
'
' 判定の形: シート・config・ログ・LLMに触れない純関数だけを叩く(R4)。
'   期待値が仕様から一意に決まらない項目は**テストにしない**(甘い期待値を
'   置いて実装を追認しない)。落とした分は末尾の一覧に残す。
'
' 結線(統合済み。本ファイル単体では1本も実行されない):
'   (1) modTestsPure7.RunAll の末尾から modTestsPure8.RunAll を呼ぶ【結線済み】
'       (modTestsPure -> 2 -> 3 -> 4 -> 5 -> 6 -> 7 -> 8 の数珠つなぎ)。
'   (2) build/modules.json へ1件追加(modTestsPure8 / src/test/modTestsPure8.bas
'       / role=test / type=std / wave=T-28)【登録済み】。12章§2のtest層一覧も
'       modTestsPure + 2..8 へ更新済み。
'   (3) wintest/tests_expected.txt を 414 -> 471 へ(本ファイル57本)【更新済み】。
'   run_lo_tests.py の PURE_ALLOWLIST には modTestsPure8 / modInboxStore /
'   modSparring / modPipeline2 / modPlayOps が既にある(裁定書8 B-7/B-9/B-10)。
'
' 設計判断(R4準拠): Excelトークン不使用。改行は vbLf 基準。乱数・時刻不使用。
'   素材は実在しない架空値で組む(15章§8.1の浜松スイーツファクトリーの文脈)。
' ============================

' ----------------------------
' 素材の定数
' ----------------------------
' 14章§6 CallChat が使う履歴の区切り(PoC裁定D11の実証方式)。実装の
' modGatewayRPN.GW_HIST_SEP を参照せず、仕様の値をテスト側に独立に置く。
Private Const SEP_HIST As String = ";;;"

' 13章§2.6 受信箱の統制語彙(status / drop_type / revive_tag)。
Private Const ST_REJECTED As String = "rejected"
Private Const ST_HOLD As String = "conditional_hold"
Private Const ST_ADOPTED As String = "adopted"
Private Const ST_UNDIAG As String = "undiagnosed"

' 13章§2.3 sparring_max_turns の既定値。
Private Const MAX_TURNS_DEFAULT As Long = 12

' 16章E-35/E-36 の結末(14章§6 DeepOutcomeOf の4値のうち本群が使う2つ)。
' 実装の Private Const を参照できないので仕様側の値を独立に置く。
Private Const OUT_CRITIQUE_SKIPPED As String = "critique_skipped"
Private Const OUT_REVISED As String = "revised"

' 15章§4.5 S2C(MK-S2C-HIT相当)の指摘本文。CritiqueDigest が1件も
' 落とさずに日本語整形することを確かめるための照合語。
Private Const C2_D1 As String = "高温化による原料調達難が拾われていない"
Private Const C2_S1 As String = "気候リスクを1件追加する"
Private Const C2_D2 As String = "業種名を変えても通る一般論が2件ある"
Private Const C2_S2 As String = "浜松2工場の固有事情に接地させる"
Private Const C2_D3 As String = "移転困難なリスクをcoverとしている"
Private Const C2_S3 As String = "hardへ改め条件を添える"
Private Const C2_AR As String = "原料調達の途絶"

' 15章§4.6 S3C(MK-S3C-HIT相当)の反応と指摘。
Private Const C3_R1 As String = "そこはもう入っている"
Private Const C3_R2 As String = "で、いくらかかるの"
Private Const C3_R3 As String = "うちには関係ない話だ"
Private Const C3_D1 As String = "3本目は経営者の関心から遠い"
Private Const C3_D2 As String = "係争中の先のD&Oは引受部門が難色を示す"

' E-36 の採否を見分けるための印(改訂前/改訂版)。
Private Const MK_BASE As String = "確定前の生成版"
Private Const MK_REV As String = "審査を踏まえた改訂版"

' ----------------------------
' RunAll: グループ単位で隔離実行する。1グループが実行時エラーで落ちても
'   残りのグループは走る(未実装/未注入の事実は GroupFail で1件の失敗として
'   可視化し、無かったことにしない)。本ファイルは数珠つなぎの末端。
' ----------------------------
Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 8
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
End Sub

Private Sub RunGroup(ByVal grpNo As Long, ByRef grpName As String)
    Select Case grpNo
    Case 1
        grpName = "G70 統制語彙"
        T_Vocab
    Case 2
        grpName = "G72 PF要約列"
        T_PfVerdict
    Case 3
        grpName = "G73 履歴切詰め"
        T_TrimHist
    Case 4
        grpName = "G74 履歴連結"
        T_JoinHist
    Case 5
        grpName = "G76 改訂要否"
        T_ShouldRevise
    Case 6
        grpName = "G77 改訂採否"
        T_Adopt
    Case 7
        grpName = "G78 批判ダイジェスト"
        T_Digest
    Case 8
        grpName = "G79 E-35/E-36分岐"
        T_DeepBranch
    End Select
End Sub

' ============================
' 共通ヘルパ(modTestsPure3/4/5/6/7 と同じ作法。各 modTestsPure* が自前で持つ)
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

Private Function Ctn(ByVal hay As String, ByVal needle As String) As Boolean
    Ctn = (InStr(hay, needle) > 0)
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

' vbLf区切りの行数(0行=空文字は0)。切詰めの件数照合に使う。
Private Function LineCount(ByVal s As String) As Long
    If LenB(s) = 0 Then Exit Function
    LineCount = UBound(Split(s, vbLf)) + 1
End Function

' ============================
' §6整合アダプタ(統合時に追加。公開名は増やさない private ヘルパ)
' ----------------------------
' 原版が仮定した3本を 14章§6 が宣言した形へ**期待値を変えずに**寄せるだけの
' 薄い層。規約そのもの(E-41の必須列・E-35の批判不合格・E-36の採否)は
' すべて呼び先の実装側にあり、ここには置かない。
' ============================

' VocabOk - 原版 HasControlledVocab(status, dropType, reviveTag, reviveDueText)。
'   §6は同じE-41判定を modInboxStore.JudgementError(...) As String で宣言して
'   いる(""=保存可・非空=保存ブロックの理由)。期日は Date を純関数へ持ち込ま
'   ないため真偽で受ける契約なので、原版の文字列の空/非空をそのまま落とす。
Private Function VocabOk(ByVal statusText As String, ByVal dropType As String, _
                         ByVal reviveTag As String, ByVal reviveDueText As String) As Boolean
    Dim reason As String
    reason = modInboxStore.JudgementError(statusText, dropType, reviveTag, _
                                          (LenB(reviveDueText) > 0))
    VocabOk = (LenB(reason) = 0)
End Function

' ReviseNeeded - 原版 ShouldReviseOf(stepNo, critiqueJson, critiqueOk)。
'   §6は2本に分けて宣言している: 批判が不合格のときの結末(E-35 =
'   critique_skipped。改訂へ進まない)は DeepOutcomeOf が唯一の分類点で、
'   指摘の有無による改訂要否は NeedsRevision が唯一の判定点。
Private Function ReviseNeeded(ByVal stepNo As Long, ByVal critiqueJson As String, _
                              ByVal critiqueOk As Boolean) As Boolean
    Dim outcome As String
    outcome = modPipeline2.DeepOutcomeOf(critiqueOk, False, False)
    If StrComp(outcome, OUT_CRITIQUE_SKIPPED, vbBinaryCompare) = 0 Then Exit Function
    ReviseNeeded = modPipeline2.NeedsRevision(critiqueJson, stepNo)
End Function

' AdoptedJson - 原版 AdoptRevisionOf(baseJson, revisedJson, revisedOk)。
'   §6に「確定JSONを返す」関数は無く、E-36の採否そのものは DeepOutcomeOf が
'   revised(改訂版を採用) / revision_discarded(改訂を破棄し改訂前を採用)へ
'   分類する。確定JSONの選択は modPipeline2.RunPipe(層(b))にあるため、層(a)
'   では分類の側を叩いて同じ規約を固定する。
Private Function AdoptedJson(ByVal baseJson As String, ByVal revisedJson As String, _
                             ByVal revisedOk As Boolean) As String
    Dim outcome As String
    outcome = modPipeline2.DeepOutcomeOf(True, True, revisedOk)
    If StrComp(outcome, OUT_REVISED, vbBinaryCompare) = 0 Then
        AdoptedJson = revisedJson
    Else
        AdoptedJson = baseJson
    End If
End Function

' ============================
' 素材(JSON)。長い文字列は論理行1023字の制約があるので継ぎ足しで組む。
' ============================

' 15章§4.5 Schema-S2C。issues 3件(missing / generic / insurability_error)と
'   additional_risks 1件。MK-S2C-HIT の形(15章§8.1 #8)。
Private Function JsS2cHit() As String
    Dim s As String
    s = "{""verdict_summary"":""検討の幅は足りているが固有性が弱い。"","
    s = s & """issues"":["
    s = s & "{""target"":""risk_no:3"",""issue_type"":""missing"","
    s = s & """detail"":""" & C2_D1 & "。"",""suggestion"":""" & C2_S1 & "。""},"
    s = s & "{""target"":""overall"",""issue_type"":""generic"","
    s = s & """detail"":""" & C2_D2 & "。"",""suggestion"":""" & C2_S2 & "。""},"
    s = s & "{""target"":""gap_no:1"",""issue_type"":""insurability_error"","
    s = s & """detail"":""" & C2_D3 & "。"",""suggestion"":""" & C2_S3 & "。""}],"
    s = s & """additional_risks"":[{""risk_name"":""" & C2_AR & ""","
    s = s & """why"":""field_insights に果実の産地集中の記述がある。""}]}"
    JsS2cHit = s
End Function

' MK-S2C-CLEAN(15章§8.1 #9・V-S2C-05)。issues も additional_risks も0件。
Private Function JsS2cClean() As String
    JsS2cClean = "{""verdict_summary"":""重大な指摘はない。""," & _
                 """issues"":[],""additional_risks"":[]}"
End Function

' 15章§4.6 Schema-S3C。MK-S3C-HIT の形(lands=false 1件・issues 2件)。
Private Function JsS3cHit() As String
    Dim s As String
    s = "{""executive_reactions"":["
    s = s & "{""story_no"":1,""reaction"":""" & C3_R1 & "。"",""lands"":false},"
    s = s & "{""story_no"":2,""reaction"":""" & C3_R2 & "。"",""lands"":true},"
    s = s & "{""story_no"":3,""reaction"":""" & C3_R3 & "。"",""lands"":true}],"
    s = s & """issues"":["
    s = s & "{""target"":""story_no:3"",""issue_type"":""wont_land"","
    s = s & """detail"":""" & C3_D1 & "。"",""suggestion"":""順序を入れ替える。""},"
    s = s & "{""target"":""overall"",""issue_type"":""uw_concern"","
    s = s & """detail"":""" & C3_D2 & "。"",""suggestion"":""do_not_propose へ回す。""}]}"
    JsS3cHit = s
End Function

' MK-S3C-CLEAN(15章§8.1 #11・V-S3C-05)。lands が3件とも true かつ issues 0件。
Private Function JsS3cClean() As String
    Dim s As String
    s = "{""executive_reactions"":["
    s = s & "{""story_no"":1,""reaction"":""筋は通っている。"",""lands"":true},"
    s = s & "{""story_no"":2,""reaction"":""社内で検討する。"",""lands"":true},"
    s = s & "{""story_no"":3,""reaction"":""資料を置いていって。"",""lands"":true}],"
    s = s & """issues"":[]}"
    JsS3cClean = s
End Function

' V-S3C-05 の**両条件**(lands全true かつ issues0件)のうち片方だけを満たす形。
'   issues は0件だが lands=false が1件ある。スキップ条件は成立しない。
Private Function JsS3cLandsFalse() As String
    Dim s As String
    s = "{""executive_reactions"":["
    s = s & "{""story_no"":1,""reaction"":""" & C3_R3 & "。"",""lands"":false},"
    s = s & "{""story_no"":2,""reaction"":""社内で検討する。"",""lands"":true},"
    s = s & "{""story_no"":3,""reaction"":""資料を置いていって。"",""lands"":true}],"
    s = s & """issues"":[]}"
    JsS3cLandsFalse = s
End Function

' 15章§6 Schema-PF の要点だけを持つ応答。survival 以外の場所に
'   "high" の語を置き、素朴な InStr 実装が引っかかることを確かめる罠。
Private Function JsPf(ByVal survival As String) As String
    Dim s As String
    s = "{""summary"":""high の水準には届かないが芽はある。"","
    s = s & """predicted_drop_types"":[""T4"",""T9""],"
    If LenB(survival) > 0 Then
        s = s & """survival"":""" & survival & ""","
    End If
    s = s & """advice_to_poster"":""加入経路を組み替えて再投函のこと。""}"
    JsPf = s
End Function

' E-36 の採否を見分ける改訂前/改訂版。中身の差は marker だけにして、
'   「どちらを返したか」以外の差でテストが通らないようにする。
Private Function JsRisk(ByVal marker As String) As String
    Dim s As String
    s = "{""risks"":[{""risk_no"":1,""risk_name"":""浜松2工場の高温化"","
    s = s & """category"":""facility_bcp"",""note"":""" & marker & """}],"
    s = s & """gaps"":[],""emerging_risks"":[]}"
    JsRisk = s
End Function

' 発話 u1..uN を seq昇順(古い順)の vbLf 区切りで組む。
Private Function HistOf(ByVal n As Long) As String
    Dim i As Long
    Dim s As String
    For i = 1 To n
        If i > 1 Then s = s & vbLf
        s = s & "u" & i
    Next i
    HistOf = s
End Function

' ============================
' G70 統制語彙(16章E-41・13章§2.6。modInboxStore.JudgementError)
'   E-41: rejected は drop_type 必須 / conditional_hold は revive_tag と期日が
'   必須。どちらも欠けたら保存ブロック。enum の端(T0 / T10 / market)を必ず
'   当てる(範囲を1つ狭く実装しても落ちるように)。
' ============================
Private Sub T_Vocab()
    ' ---- rejected: drop_type(T0-T10)が必須 ----
    ChkT "G70_rejectedはdrop_type指定で保存できる_16章E-41", _
        VocabOk(ST_REJECTED, "T4", "", "")

    ChkF "G70_rejectedでdrop_type未選択は保存できない_16章E-41", _
        VocabOk(ST_REJECTED, "", "", "")

    ' 13章§2.6 の drop_type は T0 から T10。両端を含むことを固定する。
    ChkT "G70_drop_typeの下端T0を受け付ける_13章§2.6", _
        VocabOk(ST_REJECTED, "T0", "", "")

    ChkT "G70_drop_typeの上端T10を受け付ける_13章§2.6", _
        VocabOk(ST_REJECTED, "T10", "", "")

    ChkF "G70_drop_typeのT11は範囲外_13章§2.6", _
        VocabOk(ST_REJECTED, "T11", "", "")

    ' ---- conditional_hold: revive_tag と revive_due の**両方**が必須 ----
    ChkT "G70_conditional_holdはタグと期日の両方があれば保存できる_16章E-41", _
        VocabOk(ST_HOLD, "", "tech", "2026/12/31")

    ChkF "G70_conditional_holdで期日未入力は保存できない_16章E-41", _
        VocabOk(ST_HOLD, "", "tech", "")

    ChkF "G70_conditional_holdでタグ未入力は保存できない_16章E-41", _
        VocabOk(ST_HOLD, "", "", "2026/12/31")

    ' revive_tag の enum は tech / regulation / partner / data / market の5値。
    ChkT "G70_revive_tagの末尾marketを受け付ける_13章§2.6", _
        VocabOk(ST_HOLD, "", "market", "2026/12/31")

    ChkF "G70_revive_tagのenum外は保存できない_13章§2.6", _
        VocabOk(ST_HOLD, "", "weather", "2026/12/31")

    ' ---- 条件付き列を要求しない status ----
    ChkT "G70_adoptedは条件付き列が空でも保存できる_13章§2.6", _
        VocabOk(ST_ADOPTED, "", "", "")

    ChkT "G70_undiagnosedは条件付き列が空でも保存できる_13章§2.6", _
        VocabOk(ST_UNDIAG, "", "", "")

    ' status 自体も統制語彙(6値)。表に無い値は通さない。
    ChkF "G70_status自体がenum外なら保存できない_13章§2.6", _
        VocabOk("dropped", "", "", "")
End Sub

' ============================
' G72 プリフライト要約列(13章§2.6・15章§6 Schema-PF。modPlayOps.PfSurvivalOf)
'   pf_survival は「診断結果の要約列」で enum high / mid / low。日本語ラベル
'   ではなく enum 値を入れる列である(13章§2.6)。
' ============================
Private Sub T_PfVerdict()
    ChkS "G72_survivalのhighをそのまま返す_13章§2.6", _
        modPlayOps.PfSurvivalOf(JsPf("high")), "high"

    ChkS "G72_survivalのmidをそのまま返す_13章§2.6", _
        modPlayOps.PfSurvivalOf(JsPf("mid")), "mid"

    ' summary に "high" の語が入っていても survival の値を返すこと
    ' (キーを見ずに本文を拾う実装だとここで落ちる)。
    ChkS "G72_summaryのhighに引きずられずsurvivalを返す_15章§6", _
        modPlayOps.PfSurvivalOf(JsPf("low")), "low"

    ' survival キーが無い応答(検証不合格の生JSON等)は要約列を作らない。
    ChkS "G72_survivalキー不在は空を返す_13章§2.6", _
        modPlayOps.PfSurvivalOf(JsPf("")), ""

    ' 壊れたJSON(途中で切れた応答)でも例外を投げず空を返す(14章§6の
    ' 「純内部関数は例外を投げない」規約)。
    ChkS "G72_壊れたJSONは空を返す_14章§6", _
        modPlayOps.PfSurvivalOf("{""survival"":""hi"), ""

    ChkS "G72_空文字は空を返す_14章§6", _
        modPlayOps.PfSurvivalOf(""), ""
End Sub

' ============================
' G73 modSparring.TrimHistoryOf(16章E-44・13章§2.3)
'   E-44: リボンへ渡す履歴を直近 sparring_max_turns 往復に制限する。
'   13章§2.3: 超過は**古い順に切捨て**。全履歴は case_data に残るので、
'   ここで捨てるのは「渡す分」だけである。
'   数え方: 14章§6 CallChat は histU / histA を**別々に**受けるので、本群は
'   片側の発話列(1行=1往復ぶんの片側。case_data の sparring_u または
'   sparring_a)を渡す前提で置いている。
' ============================
Private Sub T_TrimHist()
    Dim src As String
    Dim res As String

    src = HistOf(5)

    res = modSparring.TrimHistoryOf(src, 3)
    ChkN "G73_上限3なら3件だけ渡す_16章E-44", LineCount(res), 3

    ' 残るのは新しい3本。並びは元のまま(seq昇順)で反転しない
    ' (新しい順への並べ替えは HistoryJoinOf の責務。G74)。
    ChkS "G73_残るのは新しい3件で並びは元のまま_13章§2.3", res, _
        "u3" & vbLf & "u4" & vbLf & "u5"

    ChkB "G73_切り捨てるのは古い側_13章§2.3", _
        (Not Ctn(res, "u1")), "実際=[" & HeadOf(res) & "]"

    ' 件数ちょうどの境界(1件も落とさない)。
    ChkS "G73_件数ちょうどの上限では1件も落とさない_16章E-44", _
        modSparring.TrimHistoryOf(src, 5), src

    ' 上限が件数を上回るとき(まだ肥大していない通常運転)。
    ChkS "G73_上限が件数を上回るなら全件渡す_16章E-44", _
        modSparring.TrimHistoryOf(src, 9), src

    ' 履歴が無いとき(初回発話)は空のまま。
    ChkS "G73_履歴が無ければ空を返す_16章E-44", _
        modSparring.TrimHistoryOf("", 3), ""

    ' 既定値 sparring_max_turns=12(13章§2.3)の境界。13件目で1件落ちる。
    ChkS "G73_既定12往復の境界で最古の1件が落ちる_13章§2.3", _
        Left$(modSparring.TrimHistoryOf(HistOf(13), MAX_TURNS_DEFAULT), 2), "u2"
End Sub

' ============================
' G74 modSparring.HistoryJoinOf(14章§6 CallChat・15章§6.5)
'   14章§6: 履歴は「新しい順」に ";;;" 区切りで連結して渡す(PoC裁定D11)。
'   第2引数の maxTurns は 0(=全件。切詰めは G73 が別に固定する)を渡す。
' ============================
Private Sub T_JoinHist()
    Dim res As String

    res = modSparring.HistoryJoinOf(HistOf(3), 0)

    ChkS "G74_新しい順に区切りで連結する_14章§6", res, _
        "u3" & SEP_HIST & "u2" & SEP_HIST & "u1"

    ChkS "G74_2件でも新しい順を保つ_14章§6", _
        modSparring.HistoryJoinOf(HistOf(2), 0), "u2" & SEP_HIST & "u1"

    ' 末尾に区切りを残さない(空の往復を1本増やしてリボンへ渡さない)。
    ChkB "G74_末尾に区切りを付けない_14章§6", _
        (Right$(res, Len(SEP_HIST)) <> SEP_HIST), "実際=[" & HeadOf(res) & "]"

    ChkS "G74_1件なら区切りを付けない_14章§6", _
        modSparring.HistoryJoinOf(HistOf(1), 0), "u1"

    ChkS "G74_履歴が無ければ空を返す_14章§6", _
        modSparring.HistoryJoinOf("", 0), ""

    ' 発話本文を改変しない(日本語・記号を含む実際の発話で確認する)。
    ChkS "G74_発話本文を改変しない_15章§6.5", _
        modSparring.HistoryJoinOf("器は誰が持つ?" & vbLf & "ヤモリ型の応用でいけるか", 0), _
        "ヤモリ型の応用でいけるか" & SEP_HIST & "器は誰が持つ?"
End Sub

' ============================
' G76 改訂要否(16章E-35・15章§4.7・V-S2C-05/V-S3C-05)
'   改訂パスへ進むかどうかの判定。批判が不合格(修復1回でも直らない)なら
'   E-35 で批判をスキップして生成版を確定するので、改訂も走らない。
'   §6の宣言では DeepOutcomeOf(E-35の分類)と NeedsRevision(指摘の有無)の
'   2本に分かれる(ReviseNeeded アダプタ参照)。
' ============================
Private Sub T_ShouldRevise()
    ' MK-S2C-HIT(issues 3件) -> 改訂へ。
    ChkT "G76_S2C指摘ありは改訂へ進む_15章§8.1", _
        ReviseNeeded(2, JsS2cHit(), True)

    ' MK-S2C-CLEAN(issues 0件) -> 改訂スキップ(V-S2C-05。呼び出しを節約する)。
    ChkF "G76_S2C指摘0件は改訂をスキップする_15章§4.5スキップ条件", _
        ReviseNeeded(2, JsS2cClean(), True)

    ' MK-S3C-HIT(issues 2件) -> 改訂へ。
    ChkT "G76_S3C指摘ありは改訂へ進む_15章§8.1", _
        ReviseNeeded(3, JsS3cHit(), True)

    ' MK-S3C-CLEAN(lands全true かつ issues 0件) -> 改訂スキップ(V-S3C-05)。
    ChkF "G76_S3Cはlands全trueかつ指摘0件でスキップする_15章§4.6スキップ条件", _
        ReviseNeeded(3, JsS3cClean(), True)

    ' V-S3C-05 は**両条件**が要る。issues 0件でも lands=false があれば改訂へ。
    ChkT "G76_S3Cはissues0件でもlands偽があれば改訂へ進む_15章§4.6両条件", _
        ReviseNeeded(3, JsS3cLandsFalse(), True)

    ' E-35: 批判が不合格なら批判をスキップして生成版を確定する(改訂しない)。
    ChkF "G76_S2C批判が不合格なら改訂しない_16章E-35", _
        ReviseNeeded(2, JsS2cHit(), False)

    ChkF "G76_S3C批判が不合格なら改訂しない_16章E-35", _
        ReviseNeeded(3, JsS3cHit(), False)

    ' 入念モードの対象は S2 / S3 だけ(15章§4.5-4.7・DeepEnabled と同じ範囲)。
    ChkF "G76_S1は入念モードの対象外_15章§4.5", _
        ReviseNeeded(1, JsS2cHit(), True)

    ChkF "G76_S4は入念モードの対象外_15章§4.5", _
        ReviseNeeded(4, JsS2cHit(), True)
End Sub

' ============================
' G77 改訂の採否(16章E-36・15章§4.7)
'   改訂結果は改訂前と同じ Check 関数(CheckS2/CheckS3)を通す。不合格なら
'   改訂を破棄し、検証合格済みの改訂前JSONを確定として採用する(Step失敗に
'   しない)。s2r_json / s3r_json に入るのは合格した改訂版だけ。
' ============================
Private Sub T_Adopt()
    Dim baseJson As String
    Dim revJson As String

    baseJson = JsRisk(MK_BASE)
    revJson = JsRisk(MK_REV)

    ChkS "G77_改訂版が合格なら改訂版を確定する_16章E-36", _
        AdoptedJson(baseJson, revJson, True), revJson

    ChkS "G77_改訂版が不合格なら改訂前を確定する_16章E-36", _
        AdoptedJson(baseJson, revJson, False), baseJson

    ' 不合格の改訂版の断片を確定JSONへ混ぜない(2本を継ぎ合わせない)。
    ChkB "G77_不合格の改訂版は確定JSONに混ざらない_16章E-36", _
        (Not Ctn(AdoptedJson(baseJson, revJson, False), MK_REV)), _
        "実際=[" & HeadOf(AdoptedJson(baseJson, revJson, False)) & "]"

    ChkB "G77_合格の改訂版に改訂前が混ざらない_16章E-36", _
        (Not Ctn(AdoptedJson(baseJson, revJson, True), MK_BASE)), _
        "実際=[" & HeadOf(AdoptedJson(baseJson, revJson, True)) & "]"

    ' 改訂の応答自体が返らなかった場合も改訂前で続行する(Step失敗にしない)。
    ChkS "G77_改訂が空応答でも改訂前で続行する_16章E-36", _
        AdoptedJson(baseJson, "", False), baseJson
End Sub

' ============================
' G78 批判ダイジェスト(15章§4.7。modPipeline2.CritiqueDigest)
'   {{critiqueDigest}} = critiqueJson の issues / additional_risks /
'   executive_reactions を**日本語整形**した文字列。改訂プロンプトに載る
'   唯一の指摘の運び手なので、1件も落とさないことが要件。
'   §6の第2引数 stepNo は S2C=2 / S3C=3。
' ============================
Private Sub T_Digest()
    Dim d2 As String
    Dim d3 As String

    d2 = modPipeline2.CritiqueDigest(JsS2cHit(), 2)
    d3 = modPipeline2.CritiqueDigest(JsS3cHit(), 3)

    ChkB "G78_S2Cの指摘detailを1件も落とさない_15章§4.7", _
        (Ctn(d2, C2_D1) And Ctn(d2, C2_D2) And Ctn(d2, C2_D3)), _
        "実際=[" & HeadOf(d2) & "]"

    ChkB "G78_S2Cの改善方向suggestionを1件も落とさない_15章§4.7", _
        (Ctn(d2, C2_S1) And Ctn(d2, C2_S2) And Ctn(d2, C2_S3)), _
        "実際=[" & HeadOf(d2) & "]"

    ChkB "G78_S2Cのadditional_risksを落とさない_15章§4.7", _
        Ctn(d2, C2_AR), "実際=[" & HeadOf(d2) & "]"

    ChkB "G78_S3Cの指摘detailを落とさない_15章§4.7", _
        (Ctn(d3, C3_D1) And Ctn(d3, C3_D2)), "実際=[" & HeadOf(d3) & "]"

    ' 「日本語整形した文字列」であってJSONの素通しではない(生のキー名が
    ' 残っていたら整形していない)。
    ChkB "G78_生JSONのキー名を素通しにしない_15章§4.7", _
        (Not Ctn(d2, """issue_type""") And Not Ctn(d2, """suggestion""")), _
        "実際=[" & HeadOf(d2) & "]"

    ' 複数の指摘は行で分ける(1行に詰め込むと改訂プロンプトで読めない)。
    ChkB "G78_複数の指摘は改行で分ける_15章§4.7", _
        Ctn(d2, vbLf), "実際=[" & HeadOf(d2) & "]"
End Sub

' ============================
' G79 E-35/E-36 の分岐合成(15章§8.1 受入条件2)
'   quality_mode=deep で S2 -> MK-S2C-HIT -> 改訂 と
'   S2 -> MK-S2C-CLEAN -> 改訂スキップ の両経路が流れること。S3C も同様。
'   2つの純核を合成した結果が16章E-35/E-36の「確定する成果物」に一致する
'   ことを見る(判定を1本ずつ見るG76/G77とは別の層の検問)。
' ============================
Private Sub T_DeepBranch()
    Dim baseJson As String
    Dim revJson As String
    Dim doRevise As Boolean

    baseJson = JsRisk(MK_BASE)
    revJson = JsRisk(MK_REV)

    ' (1) MK-S2C-HIT -> 改訂 -> 改訂版が合格 -> 改訂版を確定(s2r_json)。
    doRevise = ReviseNeeded(2, JsS2cHit(), True)
    ChkB "G79_S2C_HITは改訂へ進み合格した改訂版を確定する_15章§8.1", _
        (doRevise And AdoptedJson(baseJson, revJson, True) = revJson), _
        "改訂要否=" & doRevise

    ' (2) MK-S2C-HIT -> 改訂 -> 改訂版が不合格 -> 改訂前を確定(E-36)。
    ChkB "G79_S2C_HITでも改訂不合格なら改訂前を確定する_16章E-36", _
        (ReviseNeeded(2, JsS2cHit(), True) And _
         AdoptedJson(baseJson, revJson, False) = baseJson), _
        "実際=[" & HeadOf(AdoptedJson(baseJson, revJson, False)) & "]"

    ' (3) MK-S2C-CLEAN -> 改訂スキップ -> 生成版が確定(改訂を呼ばない)。
    ChkB "G79_S2C_CLEANは改訂を呼ばず生成版を確定する_15章§4.5スキップ条件", _
        ((Not ReviseNeeded(2, JsS2cClean(), True)) And _
         AdoptedJson(baseJson, "", False) = baseJson), _
        "実際=[" & HeadOf(AdoptedJson(baseJson, "", False)) & "]"

    ' (4) MK-S3C-HIT -> 改訂 -> 不合格 -> 改訂前を確定(S3も同じ規約)。
    ChkB "G79_S3C_HITでも改訂不合格なら改訂前を確定する_16章E-36", _
        (ReviseNeeded(3, JsS3cHit(), True) And _
         AdoptedJson(baseJson, revJson, False) = baseJson), _
        "実際=[" & HeadOf(AdoptedJson(baseJson, revJson, False)) & "]"

    ' (5) 批判が不合格(E-35) -> 批判をスキップして生成版を確定。
    '     本体Stepは成功のままなので、確定する成果物は生成版である。
    ChkB "G79_批判不合格は改訂せず生成版を確定する_16章E-35", _
        ((Not ReviseNeeded(2, JsS2cHit(), False)) And _
         AdoptedJson(baseJson, "", False) = baseJson), _
        "実際=[" & HeadOf(AdoptedJson(baseJson, "", False)) & "]"
End Sub

' ============================
' 意図的に未テスト(期待値が15章/16章/13章から一意に定まらないもの)
'   司令塔へ懸念として返した項目。**甘い期待値を置いて実装を追認しない**
'   ために、ここへ列挙して空白のまま残す。
'     (a) JudgementError の status=merged(13章§2.6は merged_into 必須と
'         するが E-41 は触れず、SetInboxJudgement の引数にも merged_into が
'         無い。§6は「SetInboxJudgement 側で fail-closed に拒否」と宣言)。
'     (b) revive_due の**書式**検査の有無(13章§2.6は「期日」とだけ書く)。
'         本ファイルは空/非空だけを当てている。
'     (c) 関心度集計(FR-17)の 1件だけのテーマを載せるか・同数のときの並び・
'         0件のときの既定文言・上限件数。§6に純関数が無いのでG71ごと落とした。
'     (d) PfSurvivalOf と pf_pred_types(T1..T10 の ";" 連結)の関係
'         (13章§2.6は2列あるのに要約列の宣言は1本)。
'     (e) TrimHistoryOf の maxTurns<=0(§6は「全件」と宣言したのでG73では
'         当てず、G74が全件渡しとして使うにとどめた)。
'     (f) 壁打ちの応答後の継続可否(E0202/E0201/E0204/E0205/E0206の扱い)。
'         §6の CanContinueSparring は**送信前**判定なのでG75ごと落とした。
'     (g) ShouldReviseOf で issues 0件だが additional_risks が非空のとき
'         (§6は「S2Cは additional_risks が非空なら True」と宣言済み)。
'     (h) CritiqueDigest の指摘0件のときの戻り(空文字か既定文言か)と、
'         target(risk_no:3 等)を digest に載せるか。
'     (i) digest に lands=true の反応を載せるか(15章§4.7の注記と14章§6の
'         宣言が食い違う。仕様裁定待ちのためG78から1本落とした)。
' ============================
