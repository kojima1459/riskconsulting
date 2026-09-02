Attribute VB_Name = "modTestsPure12"
Option Explicit

' ============================================================================
' modTestsPure12 - W4.1(裁定書9)の回帰テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針(裁定書9 §7): 実装コード(src/app・src/ui)を一切読まずに、
'   裁定書9 と spec班改訂後の docs/spec(13/14/16/19章)だけから期待値を導いた。
'   期待値をあとから実装に合わせて緩めることは禁止(17章§1)。
'   呼び先の関数名は14章§6の逐語のみを使う。
'
' 対象と根拠:
'   W41A  B5   配列セルの「;」全角カンマ置換     13章§2.2(v2.5適用点の注記)
'   W41C  A-6  充足度highラベル「高」            19章§3 input_quality.overall / 14章§6 EnumPairsCsv
'   W41D  B2   受信箱の判定遷移(judge_to方式)    13章§2.6 / 14章§6 CanInboxTransition / 16章E-41
'   W41E  B4   ファイル名規則(純層で組める範囲)  13章§2.8 / 14章§6 SanitizeFileName
'   W41F  B9/N1 E-35/E-36警告文言と初期状態      16章E-35/E-36 / 14章§6 DeepWarningOf・LastDeepOutcome
'   (B20・N2 は純層で書けないため末尾コメントに期待値だけ残した。concerns参照)
'
' グループ単位の失敗隔離: modTestsPure.bas と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**
'   (本モジュールの Check は42本。modTestRunner への結線=RunAll呼び出しの追加と
'    tests_expected の更新は統合担当が行う)。
' ============================================================================

Public Sub RunAll()
    On Error GoTo FA
    T_W41A_SemicolonReplace
WB:
    On Error GoTo FC
    T_W41C_QualityLabel
WD:
    On Error GoTo FD
    T_W41D_InboxJudgement
WE:
    On Error GoTo FE
    T_W41E_FileNameRules
WF:
    On Error GoTo FF
    T_W41F_DeepWarnings
WDone:
    On Error GoTo F13
    modTestsPure13.RunAll
    On Error GoTo 0
    Exit Sub
F13:
    GroupFail "modTestsPure13.RunAll"
    Resume Next
FA:
    GroupFail "W41A B5 セミコロン置換"
    Resume WB
FC:
    GroupFail "W41C A-6 充足度ラベル"
    Resume WD
FD:
    GroupFail "W41D B2 受信箱判定"
    Resume WE
FE:
    GroupFail "W41E B4 ファイル名規則"
    Resume WF
FF:
    GroupFail "W41F B9/N1 入念モード警告"
    Resume WDone
End Sub

Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

' 文字列一致の1本。期待値と実際値の両方をレポートへ残す。
Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), "期待=[" & want & "] 実際=[" & act & "]"
End Sub

' 真偽一致の1本。
Private Sub ChkB(ByVal nm As String, ByVal act As Boolean, ByVal want As Boolean)
    modTestRunner.Check nm, (act = want), "期待=" & want & " 実際=" & act
End Sub

' 数値一致の1本。
Private Sub ChkN(ByVal nm As String, ByVal act As Long, ByVal want As Long)
    modTestRunner.Check nm, (act = want), "期待=" & want & " 実際=" & act
End Sub

' String() の要素数(未割当は0)。
Private Function ArrLenOf(ByRef a() As String) As Long
    On Error GoTo Bad
    ArrLenOf = UBound(a) - LBound(a) + 1
    Exit Function
Bad:
    ArrLenOf = 0
End Function

' 16進小文字/大文字のみで構成されているか。
Private Function IsHexText(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    If LenB(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsHexText = True
End Function

' ----------------------------------------------------------------------------
' W41A: B5 - 配列セルの「;」置換(裁定書9 B5。13章§2.2 v2.5)。
' 仕様: 配列セルを組み立てる関数(PathArrCell / PrevCell / IntArrCell)が要素1本を
'   書き出す直前に Replace(one, ";", ChrW(&HFF0C&)) を通す(非可逆・復元しない)。
'   要素単位で置換するため区切りの「; 」は温存され、読み戻し(「; 」で分割し
'   前後空白を除去=14章§6 modUtil.SplitKeepNonEmpty)で要素数が保たれる。
' 注: PathArrCell 等は ui層(modUICaseFmt)の関数で14章§6の公開契約に無いため、
'   層(a)からは直接呼べない。本グループは13章§2.2の置換規則そのもの
'   (期待値=置換後の格納文字列と読み戻し要素数)を、§6にある modUtil.
'   SplitKeepNonEmpty との組で固定する。実装関数の直接検証は層(b)に委ねる。
' ----------------------------------------------------------------------------
Private Sub T_W41A_SemicolonReplace()
    Dim raw1 As String
    Dim stored1 As String
    Dim want1 As String
    Dim cellText As String
    Dim parts() As String

    ' LLMが返した1要素に「; 」が含まれるケース(裁定書9 B5の発火条件そのもの)
    raw1 = "調達先A; 調達先B依存"
    ' 13章§2.2 v2.5 の式を要素単位で適用した期待格納値
    stored1 = Replace(raw1, ";", ChrW(&HFF0C&))
    want1 = "調達先A" & ChrW(&HFF0C&) & " 調達先B依存"

    ' Test_W41_02: 置換後の格納文字列の期待値(「;」だけが全角カンマU+FF0Cへ。
    '   後続の半角スペースは温存される)
    ChkS "Test_W41_02_B5_置換後の要素は;のみFF0C化_13章2.2", stored1, want1

    ' Test_W41_03: 非可逆置換の結果、要素本文に「;」が残らない
    ChkN "Test_W41_03_B5_置換後の要素に;が残らない_13章2.2", _
        InStr(1, stored1, ";", vbBinaryCompare), 0

    ' Test_W41_01: 置換済み要素を「; 」で連結した1セルは、読み戻し
    '   (「; 」で分割し前後空白を除去=SplitKeepNonEmpty)で2要素のまま。
    '   置換が無いと3要素へ分裂する(B5の欠陥再現条件)
    cellText = stored1 & "; " & "自然災害"
    parts = modUtil.SplitKeepNonEmpty(cellText, ";")
    ChkN "Test_W41_01_B5_置換済みセルの読み戻しは2要素_13章2.2", ArrLenOf(parts), 2
End Sub

' ----------------------------------------------------------------------------
' W41C: A-6 - 充足度(input_quality.overall)の high ラベル(裁定書9 A-6。19章§3 v2.5)。
' 仕様: 19章§3 input_quality.overall = high/mid/low -> 高/中/低。
'   v2.5で high の「充足度 高」を「高」へ改めた(表示側が「充足度: 」を前置)。
'   変換表の唯一の値源は modUICase.EnumPairsCsv(14章§6。書式は
'   「グループ,機械値,日本語」の vbLf 区切り・19章§3の記載順)。
' ----------------------------------------------------------------------------
Private Sub T_W41C_QualityLabel()
    Dim csv As String
    csv = modUICase.EnumPairsCsv()

    ' Test_W41_04: 機械値->日本語(A-6の本丸)
    ChkS "Test_W41_04_A6_EnumJa_overall_highは高_19章3", _
        modUICase.EnumJa("input_quality_overall", "high"), "高"

    ' Test_W41_05: 変換表そのものに「高」の行がある
    ChkB "Test_W41_05_A6_EnumPairsCsvにoverall_high_高の行_19章3", _
        (InStr(1, csv, "input_quality_overall,high,高", vbBinaryCompare) > 0), True

    ' Test_W41_06: 旧ラベル「充足度 高」は変換表から消えている
    ChkN "Test_W41_06_A6_旧ラベル充足度高が表に無い_19章3v2.5", _
        InStr(1, csv, "充足度 高", vbBinaryCompare), 0

    ' Test_W41_07: 日本語->機械値の逆引き
    ChkS "Test_W41_07_A6_EnumEn_高はhigh_19章3", _
        modUICase.EnumEn("input_quality_overall", "高"), "high"

    ' Test_W41_08: 旧ラベルの逆引きは""(表に無いラベルは推測で返さない=14章§6)
    ChkS "Test_W41_08_A6_EnumEn_旧ラベルは空_14章6", _
        modUICase.EnumEn("input_quality_overall", "充足度 高"), ""

    ' Test_W41_09/10: mid/low は従来どおり(A-6はhighだけの改訂)
    ChkS "Test_W41_09_A6_EnumJa_overall_midは中_19章3", _
        modUICase.EnumJa("input_quality_overall", "mid"), "中"
    ChkS "Test_W41_10_A6_EnumJa_overall_lowは低_19章3", _
        modUICase.EnumJa("input_quality_overall", "low"), "低"
End Sub

' ----------------------------------------------------------------------------
' W41D: B2 - 受信箱の判定(judge_to入力列方式)の遷移規約(裁定書9 B2/N5。
'   13章§2.6 v2.5・14章§6 CanInboxTransition・16章E-41・19章§3)。
' 仕様: ui層は judge_to 列を読んで SetInboxJudgement の status 引数へ渡し、
'   CanInboxTransition は「現在の status(diagnosed)->judge_toの値」を検査する。
'   許すのは (1) undiagnosed->diagnosed (2) diagnosed->adopted/conditional_hold/
'   rejected/merged の2種だけ。自己遷移(judged->judgedを含む)・判定済みからの
'   再判定・診断を飛ばした判定・enum外はすべて False。
' ----------------------------------------------------------------------------
Private Sub T_W41D_InboxJudgement()
    ' --- diagnosed -> 各判定先(可) ---
    ChkB "Test_W41_11_B2_diagnosedからadopted可_11章4", _
        modInboxStore.CanInboxTransition("diagnosed", "adopted"), True
    ChkB "Test_W41_12_B2_diagnosedからconditional_hold可_11章4", _
        modInboxStore.CanInboxTransition("diagnosed", "conditional_hold"), True
    ChkB "Test_W41_13_B2_diagnosedからrejected可_11章4", _
        modInboxStore.CanInboxTransition("diagnosed", "rejected"), True
    ChkB "Test_W41_14_B2_diagnosedからmerged可_11章4", _
        modInboxStore.CanInboxTransition("diagnosed", "merged"), True
    ChkB "Test_W41_15_B2_undiagnosedからdiagnosed可_11章4", _
        modInboxStore.CanInboxTransition("undiagnosed", "diagnosed"), True

    ' --- 自己遷移は不可(旧方式=status列を書き換えて渡すと必ずここで死ぬ。B2の核心) ---
    ChkB "Test_W41_16_B2_diagnosed自己遷移不可_11章4", _
        modInboxStore.CanInboxTransition("diagnosed", "diagnosed"), False
    ChkB "Test_W41_17_B2_adopted自己遷移不可_11章4", _
        modInboxStore.CanInboxTransition("adopted", "adopted"), False
    ChkB "Test_W41_18_B2_conditional_hold自己遷移不可_11章4", _
        modInboxStore.CanInboxTransition("conditional_hold", "conditional_hold"), False
    ChkB "Test_W41_19_B2_rejected自己遷移不可_11章4", _
        modInboxStore.CanInboxTransition("rejected", "rejected"), False
    ChkB "Test_W41_20_B2_merged自己遷移不可_11章4", _
        modInboxStore.CanInboxTransition("merged", "merged"), False

    ' --- 判定済みからの再判定・診断飛ばし・enum外は不可 ---
    ChkB "Test_W41_21_B2_判定済みからの再判定不可_11章4", _
        modInboxStore.CanInboxTransition("adopted", "rejected"), False
    ChkB "Test_W41_22_B2_診断を飛ばした判定不可_11章4", _
        modInboxStore.CanInboxTransition("undiagnosed", "adopted"), False
    ChkB "Test_W41_23_B2_enum外は不可_13章2.6", _
        modInboxStore.CanInboxTransition("diagnosed", "approved"), False

    ' --- judge_to 入力列の選択肢(19章§3 v2.5): adopted/conditional_hold/rejected。
    '     merged は選択肢に置かない(SetInboxJudgementがmergedをfail-closedで拒否するため) ---
    ChkS "Test_W41_24_N5_judge_to_adoptedは採択_19章3", _
        modUICase.EnumJa("inbox_judge_to", "adopted"), "採択"
    ChkS "Test_W41_25_N5_judge_to_conditional_holdは条件付き保留_19章3", _
        modUICase.EnumJa("inbox_judge_to", "conditional_hold"), "条件付き保留"
    ChkS "Test_W41_26_N5_judge_to_rejectedは却下_19章3", _
        modUICase.EnumJa("inbox_judge_to", "rejected"), "却下"
    ChkS "Test_W41_27_N5_judge_toにmergedは無い_19章3", _
        modUICase.EnumJa("inbox_judge_to", "merged"), ""

    ' --- E-41(統制語彙の必須化)の唯一の判定 JudgementError(14章§6) ---
    ChkB "Test_W41_28_B2_却下はdrop_type必須_16章E41", _
        (LenB(modInboxStore.JudgementError("rejected", "", "", False)) > 0), True
    ChkS "Test_W41_29_B2_却下T4は保存可_16章E41", _
        modInboxStore.JudgementError("rejected", "T4", "", False), ""
    ChkB "Test_W41_30_B2_条件付き保留は期日必須_16章E41", _
        (LenB(modInboxStore.JudgementError("conditional_hold", "", "tech", False)) > 0), True
    ChkS "Test_W41_31_B2_条件付き保留tagと期日で保存可_16章E41", _
        modInboxStore.JudgementError("conditional_hold", "", "tech", True), ""
End Sub

' ----------------------------------------------------------------------------
' W41E: B4/B7 - ファイル名規則の純層で組める範囲(裁定書9 B4・B7。13章§2.8)。
' 仕様(13章§2.8 SanitizeFileName 手順1から3):
'   1. 禁止文字 \ / : * ? " < > | と制御文字(Chr(0)-Chr(31))を "_" へ
'   2. 前後の空白を除去し、末尾のピリオドを除去
'   3. 先頭から32字で切詰め
'   (手順4の8桁付与・手順5の240字超は case_id/ディレクトリが要るため層(b))
' 日付サフィックス: <yyyymmdd> は modUtilText.IsoDateCompact(Date) の値で
'   テンプレートの一部として必ず付与(v2.5・B4。例 _20260901)。
' ※ 同名存在時の _2 _3 連番探索(B4後段)はファイルシステム依存のため層(a)では
'    書けない。期待値: 「存在検査と連番探索を必ず通す。adSaveCreateOverWrite
'    一択にしない。企業ドシエだけは連番を作らない(1社1ファイル)」(13章§2.8)。
' ----------------------------------------------------------------------------
Private Sub T_W41E_FileNameRules()
    Dim h1 As String
    Dim h2 As String

    ' Test_W41_32: 禁止文字9種はすべて "_"
    ChkS "Test_W41_32_B4_禁止文字は下線へ_13章2.8手順1", _
        modUtilText.SanitizeFileName("a\b/c:d*e?f" & Chr(34) & "g<h>i|j"), _
        "a_b_c_d_e_f_g_h_i_j"

    ' Test_W41_33: 制御文字(Chr(0)-Chr(31))も "_"
    ChkS "Test_W41_33_B4_制御文字は下線へ_13章2.8手順1", _
        modUtilText.SanitizeFileName("a" & Chr(9) & "b"), "a_b"

    ' Test_W41_34: 前後空白の除去と末尾ピリオドの除去
    ChkS "Test_W41_34_B4_前後空白と末尾ピリオド除去_13章2.8手順2", _
        modUtilText.SanitizeFileName(" 株式会社テスト. "), "株式会社テスト"

    ' Test_W41_35: 32字で切詰め
    ChkS "Test_W41_35_B4_32字切詰め_13章2.8手順3", _
        modUtilText.SanitizeFileName(String$(40, "a")), String$(32, "a")

    ' Test_W41_36: 日付サフィックスの実体 IsoDateCompact は yyyymmdd の8桁
    '   (13章§2.8 の例示 _20260901 と同じ日付で固定検証)
    ChkS "Test_W41_36_B4_IsoDateCompactはyyyymmdd_13章2.8", _
        modUtilText.IsoDateCompact(DateSerial(2026, 9, 1)), "20260901"

    ' Test_W41_37: 企業ドシエの8桁は company 由来
    '   Left$(Fnv1a64Hex(NormalizeForHash(company)), 8)(B7・13章§2.8手順4)。
    '   純層で言える性質: 8桁の16進で、同じ会社名なら常に同じ値(決定的)。
    '   これが案件をまたいで同一=1社1ファイルの根拠
    h1 = Left$(modUtilText.Fnv1a64Hex(modUtilText.NormalizeForHash("テスト製菓株式会社")), 8)
    h2 = Left$(modUtilText.Fnv1a64Hex(modUtilText.NormalizeForHash("テスト製菓株式会社")), 8)
    ChkB "Test_W41_37_B7_企業由来8桁は決定的な16進8桁_13章2.8手順4", _
        ((Len(h1) = 8) And IsHexText(h1) And (h1 = h2)), True
End Sub

' ----------------------------------------------------------------------------
' W41F: B9/N1 - 入念モードのE-35/E-36警告(裁定書9 B9・N1。16章E-35/E-36・14章§6)。
' 仕様: DeepWarningOf(outcome) は hm_warning へ出す文言(16章の逐語)。
'   警告の要らない結末(revised / revision_skipped)は ""。
'   LastDeepOutcome() は直近RunStepの結末で、値は critique_skipped /
'   revision_discarded / "" の3通りのみ。RunStep開始時に必ず""へリセットされる
'   ので、RunStep未実行の本テスト時点では "" が正。
' ----------------------------------------------------------------------------
Private Sub T_W41F_DeepWarnings()
    ' Test_W41_38: E-35の逐語(16章E-35)
    ChkS "Test_W41_38_B9_E35文言_審査省略_16章E35", _
        modPipeline2.DeepWarningOf("critique_skipped"), _
        "入念モードの審査を省略しました（生成版で続行）"

    ' Test_W41_39: E-36の逐語(16章E-36)
    ChkS "Test_W41_39_B9_E36文言_改訂破棄_16章E36", _
        modPipeline2.DeepWarningOf("revision_discarded"), _
        "入念モードの改訂を破棄しました（改訂前で続行）"

    ' Test_W41_40/41: 警告の要らない結末は ""(14章§6 DeepWarningOf)
    ChkS "Test_W41_40_B9_revisedは警告なし_14章6", _
        modPipeline2.DeepWarningOf("revised"), ""
    ChkS "Test_W41_41_B9_revision_skippedは警告なし_14章6", _
        modPipeline2.DeepWarningOf("revision_skipped"), ""

    ' Test_W41_42: RunStep未実行のLastDeepOutcomeは""(N1。14章§6)
    ChkS "Test_W41_42_N1_RunStep未実行のLastDeepOutcomeは空_14章6", _
        modPipeline2.LastDeepOutcome(), ""
End Sub

' ----------------------------------------------------------------------------
' 意図的に未テスト(純層で書けないもの。期待値だけを残す):
'
' [N2] modCaseStore.PromoteTier のenum検査(裁定書9 N2・A-1。14章§6):
'   期待値: tierText が t1_quick / t2_full / t3_sparring 以外なら1列も書かず
'   False(E0101を記録)。案件行が無いときも False。status は動かさない。
'   本関数は案件一覧シートへの書込とE0101のログ記録(err_logシート)を内包し、
'   仕様はenum検査がシートアクセスに先行することを保証していないため、
'   シート不在の層(a)から呼ぶと enum検査の合否とシート不在の False を
'   区別できない(=検査として無意味)。層(b)のExcelテストで
'   「PromoteTier(caseId, "t9_bogus") = False かつ dossier_tier 列が不変」
'   「PromoteTier(caseId, "t3_sparring") = True かつ dossier_tier = t3_sparring
'    かつ status 不変」を実証すること。
'
' [B20] InvalidateDownstream の last_ok_step min取り(裁定書9 B20。14章§6):
'   期待値: 書き込む値は Min(keepStep, 現在のlast_ok_step)。
'   ApplyRepairedState の冒頭で keepStep > 現在値 なら keepStep = 現在値 へ丸め、
'   無効化の呼び出しで案件が昇格しうる経路を残さない
'   (例: last_ok_step=1 の案件へ InvalidateDownstream(caseId, 3) を呼んでも
'    last_ok_step は 1 のまま。status も引き上げない)。
'   ApplyRepairedState は14章§6の公開契約面に無く(モジュール内部)、
'   InvalidateDownstream は案件一覧シートを書くため層(a)から検証できない。
'   層(b)のExcelテストで上の例を実証すること。
'
' [B4後段] レポートファイル名の連番探索(裁定書9 B4。13章§2.8):
'   期待値: 同名ファイルが存在する場合は上書きせず _2 _3 と連番で空きを探し
'   新規ファイルとして作る。企業ドシエファイルは1社1ファイルのため連番対象外。
'   ファイルシステム依存のため層(a)から検証できない。
' ----------------------------------------------------------------------------
