Attribute VB_Name = "modUIProgress"
Option Explicit

' ============================================================================
' modUIProgress - 進捗表示・ui_lock・フォーカス退避(ui層・T-30)
' ----------------------------------------------------------------------------
' 14章§6が宣言する公開4本(SetStage / TryEnterUiLock / ExitUiLock / ParkFocus)
' の実装。11章§5・16章 E-11 / E-50 / E-51 が本モジュールの規約の正。
'
' ui_lock の置き場(16章E-11・11章§5):
'   **モジュール変数**(保持者Step名＋取得時の Timer 値)で保持し、シート・config
'   へは永続化しない。Excelプロセスの終了で自然消滅する=強制終了しても次回起動が
'   ロックされたままにならない。保持時間が `llm_wait_sec + 120` 秒を超えたロックは
'   失効とみなして自動解除し、E0602 を記録してから新しい取得を許す(想定外の実行時
'   エラーで ExitUiLock に到達しなかった場合の回復路)。
'
' 進捗の書き切り(16章E-50(a)):
'   リボン呼出は Application.Run の同期ブロックであり、呼出中はVBAが制御を返さ
'   ないため画面を更新できない。SetStage は**呼出の前に**4行(Step名・開始時刻・
'   最大待ち時間・白画面の案内)を書き切り、DoEvents で描画を確定させる。
'
' ScreenUpdating の復帰(16章E-50(d)):
'   True への復帰は ExitUiLock に一元化する。個々のハンドラは「必ず ExitUiLock を
'   呼ぶ」だけを守ればよく、復帰漏れが構造的に起きない。
' ============================================================================

Private Const UP_SRC As String = "modUIProgress"
Private Const UP_HOME As String = "HOME"

' 13章§2.10 の名前付きレンジ(進捗ブロックの4行)。
Private Const UP_NAME_STEP As String = "hm_progress_step"
Private Const UP_NAME_STARTED As String = "hm_progress_started_at"
Private Const UP_NAME_MAXWAIT As String = "hm_progress_max_wait"
Private Const UP_NAME_NOTE As String = "hm_progress_note"

' 16章E-50(a) の固定案内文(11章のワイヤーの文言)。
Private Const UP_NOTE_TEXT As String = "画面が白くなっても処理は続いています"

' 16章E-11: ロックの失効判定に足す秒数(llm_wait_sec + 120)。
Private Const UP_LOCK_MARGIN_SEC As Double = 120#
' 1日の秒数(Timer が日付をまたいで巻き戻ったときの補正に使う)。
Private Const UP_DAY_SEC As Double = 86400#
Private Const UP_WAIT_DEFAULT As Long = 1200

' ui_lock の実体(シートにもconfigにも書かない。11章§5)。
Private gLockStep As String
Private gLockAt As Double

' ============================================================================
' SetStage - LLM呼出の**前**に進捗4行を確定表示する(16章E-50(a)・14章§6)。
'   maxWaitSec は config `llm_wait_sec`(呼び出し側が渡す)。0以下は既定1200秒。
' ============================================================================
Public Sub SetStage(ByVal stepName As String, ByVal maxWaitSec As Long)
    On Error Resume Next

    Dim waitSec As Long
    waitSec = maxWaitSec
    If waitSec <= 0 Then waitSec = UP_WAIT_DEFAULT

    modUISheet.WriteNamed UP_NAME_STEP, stepName
    modUISheet.WriteNamed UP_NAME_STARTED, Format$(Now, "hh:nn:ss")
    modUISheet.WriteNamed UP_NAME_MAXWAIT, MaxWaitText(waitSec)
    modUISheet.WriteNamed UP_NAME_NOTE, UP_NOTE_TEXT

    ' 呼出中は描画できないので、ここで描画を確定させる(E-50(a)(c))。
    DoEvents
End Sub

' 「最大20分（1Stepあたり llm_wait_sec 秒）」相当の文言(11章のワイヤー)。
'   秒を分へ換算し、端数は切り上げる(短く見せない)。
Private Function MaxWaitText(ByVal waitSec As Long) As String
    Dim mins As Long
    mins = waitSec \ 60
    If mins * 60 < waitSec Then mins = mins + 1
    If mins < 1 Then mins = 1
    MaxWaitText = "最大" & CStr(mins) & "分（1Stepあたり " & CStr(waitSec) & " 秒）"
End Function

' ============================================================================
' TryEnterUiLock - 多重実行ガード(16章E-11・14章§6)。True=取得できた。
'   保持中でも `llm_wait_sec + 120` 秒を超えていれば失効として奪い、E0602 を
'   記録する。**取得できなかった場合はモーダルダイアログを出さない**
'   (E-50(c): ui_lock中のクリックは案内表示のみ)。
' ============================================================================
Public Function TryEnterUiLock(ByVal stepName As String) As Boolean
    On Error GoTo Failed

    If LenB(gLockStep) > 0 Then
        If HeldSeconds() <= LockLimitSec() Then
            modLog.LogError "E0602", UP_SRC & ".TryEnterUiLock", _
                            "busy:" & gLockStep & ">" & stepName
            NoticeBusy stepName
            Exit Function
        End If
        ' 失効(ExitUiLock へ到達しなかった経路の回復路)。奪ってから記録する。
        modLog.LogError "E0602", UP_SRC & ".TryEnterUiLock", _
                        "lock_expired:" & gLockStep
    End If

    gLockStep = stepName
    gLockAt = Timer
    TryEnterUiLock = True
    Exit Function

Failed:
    ' ロック機構自体が壊れた場合は**取得させない**(fail-closed)。
    modLog.LogError "E0602", UP_SRC & ".TryEnterUiLock", "lock_failed", Err.Number
    TryEnterUiLock = False
End Function

' 保持時間(秒)。Timer は0時で巻き戻るため、負になったら1日ぶんを足す。
Private Function HeldSeconds() As Double
    Dim d As Double
    d = Timer - gLockAt
    If d < 0 Then d = d + UP_DAY_SEC
    HeldSeconds = d
End Function

' 失効までの秒数 = config `llm_wait_sec` + 120(16章E-11)。
Private Function LockLimitSec() As Double
    Dim waitSec As Long
    waitSec = modConfig.GetLong("llm_wait_sec", UP_WAIT_DEFAULT)
    If waitSec <= 0 Then waitSec = UP_WAIT_DEFAULT
    LockLimitSec = CDbl(waitSec) + UP_LOCK_MARGIN_SEC
End Function

' 実行中クリックの案内。ダイアログを出さず警告欄へ1行書くだけにする(E-50(c))。
Private Sub NoticeBusy(ByVal stepName As String)
    On Error Resume Next
    modUISheet.WriteNamed "hm_warning", _
        "実行中のため「" & stepName & "」は受け付けませんでした（" & _
        gLockStep & " の完了をお待ちください）。"
End Sub

' ============================================================================
' ExitUiLock - 正常終了・異常終了のどちらでも必ず呼ぶ(14章§6・16章E-50(d))。
'   ScreenUpdating の True 復帰をここに一元化する。
' ============================================================================
Public Sub ExitUiLock()
    On Error Resume Next
    gLockStep = vbNullString
    gLockAt = 0

    ' 実行が終わったので「実行中のStep」は空へ戻す(古い表示を残さない)。
    ' 開始時刻・最大待ち時間・案内文は直前の実行の記録として残す。
    modUISheet.WriteNamed UP_NAME_STEP, vbNullString

    ' 16章E-50(d): ScreenUpdating の True 復帰はここに一元化する
    ' (個々のハンドラでの復帰漏れを構造的に無効化する)。
    Application.ScreenUpdating = True
End Sub

' ============================================================================
' ParkFocus - 全アクション完了時のフォーカス退避(11章§5・16章E-51(c))。
' ----------------------------------------------------------------------------
'   セル編集モードのままVBAが止まるのを防ぐため、フォーカスを編集対象外の
'   待避セル(表示中シートの左上)へ戻す。**実行開始時にも通してから処理へ入る**。
'   起動直後はガードシートの非表示化で活性シートが不定になりうるので、
'   活性シートが取れない・非表示のときは HOME を活性化してから退避する
'   (modBoot が暫定で持っていた private ParkFocusAtBoot の役割をここへ集約。
'    17章 T-30 DoD「フォーカス退避の実装を2箇所に残さない」)。
' ============================================================================
Public Sub ParkFocus()
    On Error Resume Next

    Dim ws As Object
    Set ws = ActiveParkSheet()
    If ws Is Nothing Then Exit Sub

    ws.Cells(1, 1).Select
End Sub

' 退避先シート。活性シートが使えなければ HOME へ落とす。
'   エラーハンドラ(On Error GoTo)を使わないのは、ハンドラ稼働中に
'   On Error Resume Next を重ねられない(同一プロシージャで捕捉できなくなる)
'   ためで、区間ごとに Resume Next / GoTo 0 を挟む形にしてある。
Private Function ActiveParkSheet() As Object
    Dim ws As Object
    Dim vis As Long

    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0

    If Not ws Is Nothing Then
        vis = -2
        On Error Resume Next
        vis = CLng(ws.Visible)
        On Error GoTo 0
        If vis = -1 Then
            Set ActiveParkSheet = ws
            Exit Function
        End If
    End If

    Set ws = modUISheet.SheetOf(UP_HOME)
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    ws.Activate
    On Error GoTo 0
    Set ActiveParkSheet = ws
End Function
