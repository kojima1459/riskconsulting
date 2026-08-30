Attribute VB_Name = "modAppTypes"
Option Explicit

' ============================================================================
' modAppTypes - app層のドメイン型(Type)だけを置くモジュール
' ----------------------------------------------------------------------------
' 役割:
'   案件・提案の語彙を持つ構造体を集約する。Sub/Function は置かない(型のみ)。
'
' なぜ core ではなく app なのか(12章§2・§4・14章§6):
'   「core層のコードに製品固有の語彙を書かない」が依存規則であり、TCaseCtx は
'   案件区分・ドシエ階層・チャネル・幹事区分といった本製品固有の語彙そのもの
'   なので core層 modTypes から app層へ移設した。core は app を参照しないため、
'   TCaseCtx を引数に取れるのは app層以上(modPrompts* / modPipeline 等)に限る。
'
' R4準拠: Worksheets/Range/Application/ThisWorkbook/MsgBox/ActiveSheet には
'   一切触れない純粋な型定義。
'
' 注意(tools/run_lo_tests.py モード2): 隔離ライブラリへ自動同梱されるのは
'   core層の modTypes だけなので、本モジュールの型を使うモジュールを
'   コンパイル検査に掛けるときは run_lo_tests 側の同梱対象に modAppTypes を
'   足す必要がある(W2で TCaseCtx を引数に取る modPrompts* が入った時点の課題)。
' ============================================================================

' ----------------------------------------------------------------------------
' TCaseCtx - 1案件の文脈(14章§6の定義の10項目)。
' ----------------------------------------------------------------------------
'   15章の各プロンプトへ「この案件は何者か」を1本で渡すための入れ物。
'   値はすべて機械値(enum)または原文の文字列で持ち、日本語ラベルへの変換は
'   ui層(modUICase)の変換表が行う(19章§3)。
'   enumの正は19章§3:
'     case_type    : new / renewal
'     dossier_tier : t1_quick / t2_full / t3_sparring
'     channel      : wholesale / retail
'     kanji        : lead / non_lead / coins
'     bid          : yes / no
'     reins        : none / reins / captive
'   other_insurers / company / industry_code / industry_name は自由記述および
'   業種マスタ由来の値。
' ----------------------------------------------------------------------------
Public Type TCaseCtx
    case_type As String
    dossier_tier As String
    channel As String
    kanji As String
    bid As String
    reins As String
    other_insurers As String
    company As String
    industry_code As String
    industry_name As String
End Type

' ----------------------------------------------------------------------------
' TJudgement - 判断台帳1件の入出力レコード(13章§2.7・裁定書8 B-8)。
' ----------------------------------------------------------------------------
'   modJudgeStore.NewJudgement / ReadJudgement が使う唯一の入れ物(14章§6)。
'   13章§2.7の11列のうち、システムが払い出す judge_id / judged_at の2列を
'   除いた残り9列をそのまま持つ(TCaseCtxと同じく物理名=フィールド名)。
'   enumの正は19章§3(decision)・13章§2.7(result)。空でよい列(case_ref /
'   factor_note / result / post_loss)は空文字列のまま渡す。
' ----------------------------------------------------------------------------
Public Type TJudgement
    line_id As String
    case_ref As String
    situation As String
    decision As String
    factor_note As String
    key_reason As String
    result As String
    post_loss As String
    recorded_by As String
End Type
