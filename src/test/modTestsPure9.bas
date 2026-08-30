Attribute VB_Name = "modTestsPure9"
Option Explicit

' ============================
' modTestsPure9 - 14章§6が宣言済みで回帰網の無かった純核の契約テスト
'                 (17章§4-1 層(a)。30,000字契約による modTestsPure8 の分割先)
' ----------------------------
' テスト本数: 32本 = G71 8 / G80 6 / G81 5 / G82 5 / G83 8
' ----------------------------
' 役割: 14章§6 v2.4.5/v2.4.6 が「Private へ戻すことは契約違反」として公開を
'   宣言しながら、層(a)から1本も叩かれていなかった純核13本に回帰網を張る。
'   あわせて modTestsPure8 から **G71 関心度の集計**(FR-17)を引き取り、
'   集計の唯一の値源 InterestSummaryOf とその下請け2本を1モジュールへ束ねた
'   (期待値は modTestsPure8 の原版から1文字も動かしていない)。
'   modTestsPure8 と同じく**実装を1行も読まず**、14章§6の宣言・13章§1/§2.6/
'   §2.7・11章§4・15章§6.1/§11・16章E-06/E-07・19章§3だけを根拠に入出力を
'   固定する(仕様が値源であり、実装は追認の対象ではない)。
'     G71 関心度の集計(10章FR-17。件数集計・降順・2件以上・上限3件)
'         modInboxStore.InterestSummaryOf
'     G80 受信箱の採番・ID書式・状態遷移
'         modInboxStore.BuildInboxId / IsValidInboxId / CanInboxTransition
'     G81 関心度の同一視の粒度と1件ぶんの表示(FR-17の下請け2本)
'         modInboxStore.InterestKeyOf / FmtInterestLine
'     G82 判断台帳の採番・ID書式・enum検証
'         modJudgeStore.BuildJudgeId / IsValidJudgeId / IsValidDecision /
'                       IsValidJudgeResult
'     G83 プリフライトの判定核
'         modPlayOps.PfPredTypesOf / PfRefIds / PfFailCodeOf / CaseIdOfPfLine
'
' 判定の形: シート・config・ログ・LLMに触れない純関数だけを叩く(R4)。期待値が
'   仕様から一意に決まらない項目はテストにせず、末尾の一覧に残して司令塔へ
'   懸念として返す(甘い期待値を置いて実装を追認しない)。
'
' 結線: modTestsPure8.RunAll の末尾から本モジュールの RunAll を呼び、本モジュール
'   の末尾から **modTestsPure10.RunAll**(T-35 HTMLレポートの純部)を呼ぶ
'   (数珠つなぎ modTestsPure -> 2 .. -> 8 -> 9 -> 10。本ファイル単体では1本も
'   実行しない)。build/modules.json(role=test / wave=T-28)と run_lo_tests.py の
'   PURE_ALLOWLIST へ登録する。叩く製品モジュール(modInboxStore /
'   modJudgeStore / modPlayOps)はいずれも登録済み。
'
' 設計判断(R4準拠): Excelトークン・乱数・時刻を使わず改行は vbLf 基準。素材は
'   実在しない架空値(15章§8.1の浜松スイーツファクトリーの文脈)。
' ============================

' ----------------------------
' 素材の定数
' ----------------------------
' 10章FR-17・11章 受信箱ワイヤーの関心度(`関心度: 熊対策12件 雹災5件`)。
Private Const TH_KUMA As String = "熊対策"
Private Const TH_HYOU As String = "雹災"
Private Const TH_KAZE As String = "強風"
Private Const TH_MIZU As String = "浸水"

' 14章§6 InterestSummaryOf の既定表示件数(「多い順に既定3件まで」)。
Private Const INTEREST_TOP_DEFAULT As Long = 3

' 15章§6.1 の5種の1行書式(行頭は必ず `[ID] `)。PfRefIds はこの書式を
' **1行のまま** vbLf で連ねるだけの関数なので、素材も1行1件で置く。
Private Const RL_RULE As String = _
    "[J-03] class:adverse_selection 基準:加入者が予兆を知っている設計は引き受けない"
Private Const RL_MENU As String = _
    "[M-0012] 食品工場リスク診断サービス | 対応カテゴリ:manufacturing_quality"
Private Const RL_SCHEME As String = _
    "[S-0011] 見守りヤモリ型 | 器:自治体が契約者となり住民が加入する"
Private Const RL_PATTERN As String = _
    "[P2] 検知×補償バンドル | 構造:検知サービスとセットで残余リスクを保険がカバー"
Private Const RL_RESEARCH As String = _
    "[RT-07] 高齢者見守り連携 status:researching 判定日:2026-05-20"

' 15章§11 の検証エラー1行(行頭 `[ケースID] `)。E-07のID幻覚は V-PF-03 だけ。
Private Const ERR_GHOST As String = _
    "[V-PF-03] duplicates[0].ref_id M-9999 は実在しません"
Private Const ERR_OTHER As String = _
    "[V-PF-05] predicted_drop_types に不正な値があります: T99"

' ----------------------------
' RunAll: グループ単位で隔離実行(1グループが実行時エラーで落ちても残りは走る。
'   未実装/未注入は GroupFail で1件の失敗として可視化し、隠さない)。
' ----------------------------
Public Sub RunAll()
    Dim i As Long
    Dim grpName As String

    For i = 1 To 5
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

    ' 姉妹モジュール(T-35 HTMLレポートの純部)を同じ隔離作法で続けて回す。
    ' 数珠つなぎ: modTestsPure -> 2 -> .. -> 8 -> 9 -> 10。
    On Error Resume Next
    Err.Clear
    modTestsPure10.RunAll
    If Err.Number <> 0 Then
        GroupFail "modTestsPure10.RunAll"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub RunGroup(ByVal grpNo As Long, ByRef grpName As String)
    Select Case grpNo
    Case 1
        grpName = "G80 受信箱ID・遷移"
        T_InboxId
    Case 2
        grpName = "G81 関心度の粒度と表示"
        T_InterestParts
    Case 3
        grpName = "G82 判断台帳ID・enum"
        T_JudgeId
    Case 4
        grpName = "G83 PF判定核"
        T_PfCore
    Case 5
        grpName = "G71 関心度集計"
        T_Interest
    End Select
End Sub

' ============================
' 共通ヘルパ(modTestsPure3..8 と同じ作法。各 modTestsPure* が自前で持つ)
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

Private Sub ChkB(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    modTestRunner.Check nm, cond, detail
End Sub

' vbLf区切りの行数(0行=空文字は0)。連結結果の件数照合に使う。
Private Function LineCount(ByVal s As String) As Long
    If LenB(s) = 0 Then Exit Function
    LineCount = UBound(Split(s, vbLf)) + 1
End Function

' 15章§6 Schema-PF のうち predicted_drop_types だけを差し替えられる素材。
'   survival / summary は別の列の値源なので、本群の期待値には関与させない。
Private Function JsPfTypes(ByVal listText As String) As String
    Dim s As String
    s = "{""summary"":""芽はあるが加入経路が弱い。"","
    s = s & """predicted_drop_types"":[" & listText & "],"
    s = s & """survival"":""mid"",""advice_to_poster"":""組み替えて再投函のこと。""}"
    JsPfTypes = s
End Function

' ============================
' G80 受信箱の採番・ID書式・状態遷移(13章§1・13章§2.6・11章§4)
'   BuildInboxId は `I-YYYYMM-NNN`(monthText=数字6桁ちょうど・seq=1..999)、
'   IsValidInboxId は同じ形ちょうど(前後空白を許さない)、CanInboxTransition
'   は11章§4の遷移表が唯一の判定点。
' ============================
Private Sub T_InboxId()
    Dim toSt As Variant
    Dim i As Long
    Dim okAll As Boolean
    Dim ng As String

    ' 13章§1 の実例 I-202609-045。連番は3桁ゼロ詰め。
    ChkS "G80_受信箱IDを13章§1の形で採番する_13章§1", _
        modInboxStore.BuildInboxId("202609", 45), "I-202609-045"

    ' 上端999まで採れる(枯渇はその先)。
    ChkS "G80_受信箱IDの連番は999まで採れる_14章§6", _
        modInboxStore.BuildInboxId("202609", 999), "I-202609-999"

    ' 範囲外・桁違い・非数字はすべて ""(例外を投げない)。
    ChkB "G80_受信箱IDの範囲外と桁違いと非数字は空文字_14章§6", _
        ((LenB(modInboxStore.BuildInboxId("202609", 0)) = 0) And _
         (LenB(modInboxStore.BuildInboxId("202609", 1000)) = 0) And _
         (LenB(modInboxStore.BuildInboxId("20260", 1)) = 0) And _
         (LenB(modInboxStore.BuildInboxId("2026a9", 1)) = 0)), _
        "seq0=[" & modInboxStore.BuildInboxId("202609", 0) & "] " & _
        "5桁=[" & modInboxStore.BuildInboxId("20260", 1) & "]"

    ' 形ちょうど。000は不正・前後空白は許さない・判断台帳IDは別物。
    ChkB "G80_受信箱IDは13章§1の形ちょうどだけを通す_14章§6", _
        (modInboxStore.IsValidInboxId("I-202609-045") And _
         (Not modInboxStore.IsValidInboxId("I-202609-000")) And _
         (Not modInboxStore.IsValidInboxId(" I-202609-045")) And _
         (Not modInboxStore.IsValidInboxId("I-202609-045 ")) And _
         (Not modInboxStore.IsValidInboxId("J-202609-045"))), _
        "I-202609-045 の判定=" & modInboxStore.IsValidInboxId("I-202609-045")

    ' 11章§4が許すのは undiagnosed->diagnosed と diagnosed->判定4値 の5組だけ。
    okAll = True
    If Not modInboxStore.CanInboxTransition("undiagnosed", "diagnosed") Then
        okAll = False
        ng = ng & "undiagnosed->diagnosed "
    End If
    toSt = Array("adopted", "conditional_hold", "rejected", "merged")
    For i = 0 To UBound(toSt)
        If Not modInboxStore.CanInboxTransition("diagnosed", CStr(toSt(i))) Then
            okAll = False
            ng = ng & "diagnosed->" & CStr(toSt(i)) & " "
        End If
    Next i
    ChkB "G80_受信箱の遷移表が許す5組をすべて通す_11章§4", okAll, "不許可: " & ng

    ' 診断を飛ばした判定・自己遷移・判定済みからの再判定・enum外はすべてFalse。
    ChkB "G80_診断飛ばしと自己遷移と再判定とenum外は許さない_11章§4", _
        ((Not modInboxStore.CanInboxTransition("undiagnosed", "adopted")) And _
         (Not modInboxStore.CanInboxTransition("diagnosed", "diagnosed")) And _
         (Not modInboxStore.CanInboxTransition("adopted", "rejected")) And _
         (Not modInboxStore.CanInboxTransition("diagnosed", "dropped")) And _
         (Not modInboxStore.CanInboxTransition("unknown", "diagnosed"))), _
        "11章§4の遷移表に無い組を通した"
End Sub

' ============================
' G81 関心度の同一視の粒度と1件ぶんの表示(10章FR-17・11章 受信箱ワイヤー)
'   InterestSummaryOf(G71)が集計の唯一の値源であるのに対し、同一視の粒度は
'   InterestKeyOf、1件ぶんの表示は FmtInterestLine が唯一の値源(14章§6)。
'   集計を通さず下請け2本を直接当てる(集計側の実装で答えが出せないように)。
' ============================
Private Sub T_InterestParts()
    ' 吸収するのは前後空白の揺れと大小文字。
    '   【要裁定・下の(i)】§6は「**改行**・空白の揺れと大小文字だけを吸収する」
    '   と書くが、実測では末尾 vbLf が残る(本テストを改行込みで当てると赤)。
    '   実装を追認して甘くしないため、改行の1条件だけを本テストから外して
    '   司令塔へ差し戻す(§6を直すか modInboxStore を直すかの二択)。
    ChkB "G81_テーマキーは前後空白と大小文字の揺れを吸収する_14章§6", _
        ((modInboxStore.InterestKeyOf(" " & TH_KUMA & " ") = _
          modInboxStore.InterestKeyOf(TH_KUMA)) And _
         (modInboxStore.InterestKeyOf(TH_KUMA & vbLf) = _
          modInboxStore.InterestKeyOf(TH_KUMA)) And _
         (modInboxStore.InterestKeyOf("Bear Risk") = _
          modInboxStore.InterestKeyOf("bear risk"))), _
        "素=[" & modInboxStore.InterestKeyOf(TH_KUMA) & _
        "] 前後空白=[" & modInboxStore.InterestKeyOf(" " & TH_KUMA & " ") & _
        "] 大=[" & modInboxStore.InterestKeyOf("Bear Risk") & _
        "] 小=[" & modInboxStore.InterestKeyOf("bear risk") & "]"

    ' 意味の同一視はしない(人が読める粒度で数える。表記違いは別テーマ)。
    ChkB "G81_テーマキーは意味の同一視をしない_FR-17", _
        (modInboxStore.InterestKeyOf(TH_KUMA) <> _
         modInboxStore.InterestKeyOf("クマ対策")), _
        "熊対策=[" & modInboxStore.InterestKeyOf(TH_KUMA) & _
        "] クマ対策=[" & modInboxStore.InterestKeyOf("クマ対策") & "]"

    ChkS "G81_空テーマのキーは空を返す_14章§6", _
        modInboxStore.InterestKeyOf(""), ""

    ' 11章の表示は `熊対策12件`(テーマ+件数+件。区切りを挟まない)。
    ChkS "G81_関心度1件はテーマと件数を連ねて表示する_11章受信箱", _
        modInboxStore.FmtInterestLine(TH_KUMA, 12), TH_KUMA & "12件"

    ' 件数0以下・テーマ空は表示そのものを作らない。
    ChkB "G81_関心度1件は件数0以下とテーマ空で表示を作らない_14章§6", _
        ((LenB(modInboxStore.FmtInterestLine(TH_KUMA, 0)) = 0) And _
         (LenB(modInboxStore.FmtInterestLine(TH_KUMA, -1)) = 0) And _
         (LenB(modInboxStore.FmtInterestLine("", 3)) = 0)), _
        "0件=[" & modInboxStore.FmtInterestLine(TH_KUMA, 0) & _
        "] 空テーマ=[" & modInboxStore.FmtInterestLine("", 3) & "]"
End Sub

' ============================
' G82 判断台帳の採番・ID書式・enum検証(13章§1・13章§2.7・19章§3)
'   BuildJudgeId は `J-YYYYMM-NNN` で BuildInboxId と同型。IsValidJudgeId は
'   判断基準ID(`J-NN`)と桁数で区別できることが要件(14章§6)。
' ============================
Private Sub T_JudgeId()
    Dim en As Variant
    Dim i As Long
    Dim okAll As Boolean
    Dim ng As String

    ' 13章§1 の実例 J-202609-012。
    ChkS "G82_判断IDを13章§1の形で採番する_13章§1", _
        modJudgeStore.BuildJudgeId("202609", 12), "J-202609-012"

    ChkB "G82_判断IDの範囲外と桁違いと非数字は空文字_14章§6", _
        ((LenB(modJudgeStore.BuildJudgeId("202609", 0)) = 0) And _
         (LenB(modJudgeStore.BuildJudgeId("202609", 1000)) = 0) And _
         (LenB(modJudgeStore.BuildJudgeId("2026091", 1)) = 0) And _
         (LenB(modJudgeStore.BuildJudgeId("2026a9", 1)) = 0)), _
        "seq0=[" & modJudgeStore.BuildJudgeId("202609", 0) & "] " & _
        "7桁=[" & modJudgeStore.BuildJudgeId("2026091", 1) & "]"

    ' 形ちょうど。判断基準ID `J-03` は桁数が違うので通してはならない。
    ChkB "G82_判断IDは13章§1の形ちょうどだけを通す_14章§6", _
        (modJudgeStore.IsValidJudgeId("J-202609-012") And _
         (Not modJudgeStore.IsValidJudgeId("J-202609-000")) And _
         (Not modJudgeStore.IsValidJudgeId(" J-202609-012")) And _
         (Not modJudgeStore.IsValidJudgeId("J-03")) And _
         (Not modJudgeStore.IsValidJudgeId("I-202609-012"))), _
        "J-202609-012 の判定=" & modJudgeStore.IsValidJudgeId("J-202609-012") & _
        " J-03 の判定=" & modJudgeStore.IsValidJudgeId("J-03")

    ' 19章§3 の decision enum 5値ちょうど。空とenum外はFalse。
    okAll = True
    en = Array("raise", "close", "restrict", "keep", "improve")
    For i = 0 To UBound(en)
        If Not modJudgeStore.IsValidDecision(CStr(en(i))) Then
            okAll = False
            ng = ng & CStr(en(i)) & " "
        End If
    Next i
    ChkB "G82_decisionは19章§3の5値だけを通す_19章§3", _
        (okAll And (Not modJudgeStore.IsValidDecision("")) And _
         (Not modJudgeStore.IsValidDecision("reject"))), _
        "通らなかったenum: " & ng

    ' 13章§2.7 の result enum 3値ちょうど。空はFalse(任意列の扱いは呼出側)。
    okAll = True
    ng = ""
    en = Array("won", "lost", "pending")
    For i = 0 To UBound(en)
        If Not modJudgeStore.IsValidJudgeResult(CStr(en(i))) Then
            okAll = False
            ng = ng & CStr(en(i)) & " "
        End If
    Next i
    ChkB "G82_resultは13章§2.7の3値だけを通す_13章§2.7", _
        (okAll And (Not modJudgeStore.IsValidJudgeResult("")) And _
         (Not modJudgeStore.IsValidJudgeResult("win"))), _
        "通らなかったenum: " & ng
End Sub

' ============================
' G83 プリフライトの判定核(13章§2.6・15章§6.1・15章§11・16章E-06/E-07)
'   PfPredTypesOf は受信箱の `pf_pred_types` 列(";"区切り T1..T10)の値源、
'   PfRefIds は V-PF-03 の実在検査へ渡す一覧テキストの値源、PfFailCodeOf は
'   不合格の内訳をE0301/E0302へ切り分ける唯一の点、CaseIdOfPfLine は検証
'   エラー1行からケースIDを取り出す。
' ============================
Private Sub T_PfCore()
    Dim got As String

    ' 列の書式は ";" 区切り(13章§2.6)。
    ChkS "G83_予測棄却類型を区切り付きの1列へまとめる_13章§2.6", _
        modPlayOps.PfPredTypesOf(JsPfTypes("""T4"",""T9""")), "T4;T9"

    ' 重複は1件へ寄せ、enum外(T11)と T0(PFには現れない。19章§3)は落とす。
    ChkS "G83_予測棄却類型の重複とenum外とT0を落とす_19章§3", _
        modPlayOps.PfPredTypesOf(JsPfTypes("""T4"",""T11"",""T4"",""T0"",""T9""")), _
        "T4;T9"

    ' 予測0件の応答から列をでっち上げない。
    ChkS "G83_予測棄却類型が0件なら空を返す_13章§2.6", _
        modPlayOps.PfPredTypesOf(JsPfTypes("")), ""

    ' 15章§6.1 の5種を1行書式のまま vbLf で連ねる(行を作り替えない)。
    got = modPlayOps.PfRefIds(RL_RULE, RL_MENU, RL_SCHEME, RL_PATTERN, RL_RESEARCH)
    ChkB "G83_一覧テキストは5種を1行書式のまま行で連ねる_15章§6.1", _
        (Ctn(got, RL_RULE) And Ctn(got, RL_MENU) And Ctn(got, RL_SCHEME) And _
         Ctn(got, RL_PATTERN) And Ctn(got, RL_RESEARCH) And _
         (LineCount(got) = 5)), _
        "行数=" & LineCount(got) & " 実際=[" & HeadOf(got) & "]"

    ' 空の注入は**行ごと**落とす(空行を残すと1行1件の書式が崩れる)。
    got = modPlayOps.PfRefIds(RL_RULE, "", RL_SCHEME, "", "")
    ChkB "G83_一覧テキストは空の注入を行ごと落とす_14章§6", _
        ((LineCount(got) = 2) And Ctn(got, RL_RULE) And Ctn(got, RL_SCHEME)), _
        "行数=" & LineCount(got) & " 実際=[" & HeadOf(got) & "]"

    ' 16章E-07: ID幻覚を1件でも**含めば** E0301(他のケースと混在していても)。
    ChkS "G83_ID幻覚を含む不合格はE0301へ切り分ける_16章E-07", _
        modPlayOps.PfFailCodeOf(ERR_OTHER & vbLf & ERR_GHOST), "E0301"

    ' 16章E-06: ID幻覚を含まない不合格は E0302。
    ChkS "G83_ID幻覚を含まない不合格はE0302へ切り分ける_16章E-06", _
        modPlayOps.PfFailCodeOf(ERR_OTHER), "E0302"

    ' 検証エラー1行の行頭 `[ケースID] ` からIDだけを取り出す。不一致は ""。
    ChkB "G83_検証エラー行の先頭からケースIDを取り出す_15章§0原則10", _
        ((modPlayOps.CaseIdOfPfLine(ERR_GHOST) = "V-PF-03") And _
         (LenB(modPlayOps.CaseIdOfPfLine("ref_id が実在しません")) = 0)), _
        "実際=[" & modPlayOps.CaseIdOfPfLine(ERR_GHOST) & "] " & _
        "行頭なし=[" & modPlayOps.CaseIdOfPfLine("ref_id が実在しません") & "]"
End Sub

' ============================
' G71 関心度の集計(10章 FR-17・11章 受信箱ワイヤー。裁定書9-1で復帰)
'   FR-17: 同一テーマの重複投函は棄却せず関心度スコア(件数)として集計する。
'   11章の表示は `関心度: 熊対策12件 雹災5件`(件数降順・半角空白区切り)。
'   素材は熊対策3件・雹災2件を**交互**に並べ、隣接行の数え上げでは正解に
'   ならないようにする。原版6本は期待値を1文字も変えず呼び先だけを与えた。
' ============================
Private Sub T_Interest()
    Dim src As String
    Dim res As String

    src = TH_KUMA & vbLf & TH_HYOU & vbLf & TH_KUMA & vbLf & _
          TH_HYOU & vbLf & TH_KUMA
    res = modInboxStore.InterestSummaryOf(src)

    ChkB "G71_同一テーマ3件が3件として集計される_FR-17", _
        Ctn(res, TH_KUMA & "3件"), "実際=[" & HeadOf(res) & "]"

    ChkB "G71_別テーマは別に集計される_FR-17", _
        Ctn(res, TH_HYOU & "2件"), "実際=[" & HeadOf(res) & "]"

    ' 重複投函を1件に丸めない(棄却しないのがFR-17の要求)。
    ChkB "G71_重複投函を1件に丸めない_FR-17", _
        (Not Ctn(res, TH_KUMA & "1件")), "実際=[" & HeadOf(res) & "]"

    ' 別テーマを1つのテーマに混ぜない(3+2=5件にしない)。
    ChkB "G71_別テーマを合算しない_FR-17", _
        (Not Ctn(res, TH_KUMA & "5件")), "実際=[" & HeadOf(res) & "]"

    ' 件数降順(11章の表示順。多いテーマが先)。
    ChkB "G71_件数の多いテーマが先に並ぶ_11章受信箱", _
        (InStr(res, TH_KUMA) > 0 And InStr(res, TH_KUMA) < InStr(res, TH_HYOU)), _
        "実際=[" & HeadOf(res) & "]"

    ' 0件のときに件数表記を作らない(空の集計をでっち上げない)。
    ChkB "G71_投函0件では件数表記を出さない_FR-17", _
        (Not Ctn(modInboxStore.InterestSummaryOf(""), "件")), _
        "実際=[" & HeadOf(modInboxStore.InterestSummaryOf("")) & "]"

    ' 【裁定書9-1で追加】原版が「§6に規約が無い」として当てなかった2点は、
    '   §6が InterestSummaryOf の宣言で明記したので固定する。
    ' (1) 1件だけのテーマは載せない(FR-17の「重複投函」ではない)。
    Dim soloText As String
    soloText = modInboxStore.InterestSummaryOf(TH_KUMA & vbLf & TH_HYOU & vbLf & TH_KUMA)
    ChkB "G71_1件だけのテーマは載せない_FR-17", _
        (Ctn(soloText, TH_KUMA & "2件") And (Not Ctn(soloText, TH_HYOU))), _
        "実際=[" & HeadOf(soloText) & "]"

    ' (2) 上限件数(既定3件)を超えたテーマは載せない。落ちるのは件数の少ない側。
    Dim manyText As String
    manyText = modInboxStore.InterestSummaryOf(RepLines(TH_KUMA, 5) & _
        RepLines(TH_HYOU, 4) & RepLines(TH_KAZE, 3) & RepLines(TH_MIZU, 2))
    ChkB "G71_上限件数を超えたテーマは載せない_FR-17", _
        (Ctn(manyText, TH_KUMA & "5件") And Ctn(manyText, TH_HYOU & "4件") And _
         Ctn(manyText, TH_KAZE & "3件") And (Not Ctn(manyText, TH_MIZU))), _
        "上限=" & INTEREST_TOP_DEFAULT & " 実際=[" & HeadOf(manyText) & "]"
End Sub

' 同じテーマを n 行ぶん積む(1行=1投函。末尾は改行で閉じる)。
Private Function RepLines(ByVal themeText As String, ByVal n As Long) As String
    Dim i As Long
    For i = 1 To n
        RepLines = RepLines & themeText & vbLf
    Next i
End Function

' ============================
' 意図的に未テスト(期待値が14章§6・13章・15章・19章から一意に定まらないもの)
'   司令塔へ懸念として返した項目。**甘い期待値を置いて実装を追認しない**
'   ために、ここへ列挙して空白のまま残す。
'     (a) InterestKeyOf の**空白だけのテーマ**(" " や vbLf のみ)。§6は
'         「空白の揺れを吸収する」「空テーマは ""」と別々に書くだけで、
'         吸収後に空になった場合を空テーマと見なすとは書いていない。
'     (b) InterestKeyOf が語中の空白を畳むか(前後の除去だけか)。
'         `Bear Risk` と `BearRisk` を同一視するかの規約が無い。
'     (c) IsValidDecision / IsValidJudgeResult の**大小文字**。§6は
'         「enumに一致するか」とだけ書き、`Raise` の扱いを決めていない。
'     (d) PfPredTypesOf の**並び順**(応答の出現順か昇順か)。本群は順序が
'         どちらでも同じ答えになる素材だけを使って当てている。
'     (e) PfPredTypesOf の壊れたJSON・キー不在の戻り(PfSurvivalOf は §6が
'         「enum以外は ""」と書くが、こちらは配列の規約しか書いていない)。
'     (f) PfRefIds の**連結順**(§6の引数順のままか)と、15章§6.1の
'         `(登録なし)` を空の注入と同じく落とすか。
'     (g) PfFailCodeOf の errText が**空**のとき(不合格でない呼び出し)。
'     (h) CaseIdOfPfLine が角括弧内にケースID以外の語を持つ行をどう扱うか。
'     (i) 【未テストではなく**食い違い**。要裁定】InterestKeyOf の**改行**。
'         14章§6 は「改行・空白の揺れと大小文字だけを吸収する」と宣言するが、
'         実測では `InterestKeyOf("熊対策" & vbLf)` が末尾の vbLf を残し
'         `InterestKeyOf("熊対策")` と一致しない(前後空白と大小文字は吸収
'         されている)。§6の「改行」を落とすか modInboxStore を直すかは本
'         ファイルの権限外なので、G81 から当該1条件だけを外して差し戻す。
' ============================
