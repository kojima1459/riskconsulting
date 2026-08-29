Attribute VB_Name = "modTypes"
Option Explicit

' ============================================================================
' modTypes - core層の汎用ユーザー定義型(Type)だけを置くモジュール
' ----------------------------------------------------------------------------
' 役割:
'   モジュールをまたいで受け渡す構造体のうち、【製品固有の語彙を持たない
'   汎用のもの】をここに集約する。Sub/Function は一切置かない(型のみ)。
'
' 12章§4の依存規則:
'   ドメイン型(TCaseCtx 等の案件・提案の語彙を持つ型)は core に置かず
'   app層 modAppTypes が持つ。ここに置いてよいのは「ログ1行」のような
'   基盤の責務そのものに属する型に限る。
'
' 移植元: PoC「マイ本棚AI」 src/core/modTypes.bas の骨格(型のみ・R4準拠)。
'   PoC固有の型(ExtractedPage / ShelfChunk / Hit)はRPNに対応物が無いため
'   非移植。
'
' R4準拠: Worksheets/Range/Application/ThisWorkbook/MsgBox/ActiveSheet には
'   一切触れない。LibreOffice の構文チェック(tools/run_lo_tests.py モード2)は
'   全モジュールの隔離ライブラリへ本モジュールを同梱するため、ここに書いた型は
'   どのモジュールからでも参照できる。
' ============================================================================

' ----------------------------------------------------------------------------
' TRunLogRec - run_log の1行(13章§2.4)。
' ----------------------------------------------------------------------------
'   run_log は「1行=1 LLM呼び出し」で14列ある。列ごとの引数にすると
'   modLog.LogRun が14引数になり、呼び出し側で順序を1つ間違えても
'   コンパイルが通ってしまう(model と transport の取り違えは実行時にも
'   気付けない)。名前で埋める構造体にして取り違えを構造的に防ぐ。
'
'   フィールド名は run_log の列名と同じにしてある。ただし step 列だけは
'   VBAの予約語 Step と衝突するため stepName とした(値は列 step へ書く)。
'   round_no は「受信箱起点の pf/wt/fg では空」を表現する必要があるため
'   Long ではなく String で持つ(空文字=記録しない)。
'
'   detail は 400字上限・本文非記録(NFR-S3)。上限の切詰めは modLog が行う
'   ので、呼び出し側は素の文字列を入れてよい。
' ----------------------------------------------------------------------------
Public Type TRunLogRec
    run_at As String              ' 空なら modLog が記録時刻で埋める
    case_id As String             ' 案件ID または 受信箱ID
    round_no As String            ' 案件一覧の現在値の複写。受信箱起点は空
    stepName As String            ' 列名は step。全12値(19章§4)
    play As String                ' PL-01..08
    transport As String           ' ribbon / direct / mock
    model As String
    latency_ms As Long
    input_chars As Long
    output_chars As Long
    injected_kb_ids As String     ' ";"区切り(modKnowledge.LastInjectedIds)
    validate_result As String     ' ok / repaired / failed
    detail As String              ' 最大400字・本文非記録
    operator As String
End Type
