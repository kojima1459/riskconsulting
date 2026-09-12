Attribute VB_Name = "modReportMail"
Option Explicit

' modReportMail - 報告メールの下書き作成(W12-d・Z-45。裁定書38 §1 班E)。
'
' 髙橋回答 Q-8(docs/29 §10.1)・v1.1 §6.1「メールの作成/下書き保存は可能。
' 自動送信は禁止(送信は人が行う)」を実装する。方式は「報告シート＋
' [報告文をコピー]」から**Outlookの下書きを開いて提示する**へ変更する
' (自動送信はしない=.Display のみ。.Send は絶対に呼ばない)。
'
' 生成は Outlook.Application を **late-bound**(参照設定を足さず CreateObject
' のみ)で行う。cmd・WMI・スクリプトシェル(AV が重く見る起動系)はこの経路に一切登場しない
' (16章 E-70・docs/24・29 の社内AV事情)。Outlook が無い・COMが拒否された
' 場合は、既存の「報告文をコピー」相当のフォールバック(呼び出し側が
' report_text をコピー導線へ渡す)に自動で切り替える。
'
' 宛先/CC は config `report_mail_to` / `report_mail_cc`(既定は空。19章§4・
' 13章§2.3・build/sheets_main.json の config シートが正)。

Private Const RM_SRC As String = "modReportMail"
Private Const RM_TAIL_LINES As Long = 20

' SendReportMail - HTML画面の[報告メールを作成]から呼ぶ唯一の入口。
' 成功時: Outlookの下書きを開いて ok:true を返す(送信はしない)。
' 失敗時: ok:false + report_text/report_subject を返し、呼び出し側(画面)が
'   既存の[コピー]導線へフォールバックする(16章 E-70)。
Public Function SendReportMail(ByVal caseId As String) As String
    Dim subjectText As String, bodyText As String, toAddr As String, ccAddr As String
    subjectText = ReportSubject(caseId)
    bodyText = BuildReportBody(caseId)
    toAddr = modConfig.GetStr("report_mail_to", vbNullString)
    ccAddr = modConfig.GetStr("report_mail_cc", vbNullString)
    If TryOutlookDisplay(subjectText, bodyText, toAddr, ccAddr) Then
        SendReportMail = "{""ok"":true,""message"":" & _
            modNaviJson.Q("Outlookの下書きを開きました。内容をご確認のうえ、送信はご自身で行ってください（自動送信はしません）。") & "}"
    Else
        SendReportMail = "{""ok"":false,""kind"":""fallback"",""message"":" & _
            modNaviJson.Q("Outlookの下書きを開けなかったため、報告文をコピーしました。ご利用のメールソフトに貼り付けてください。") & _
            ",""report_text"":" & modNaviJson.Q(bodyText) & _
            ",""report_subject"":" & modNaviJson.Q(subjectText) & "}"
    End If
End Function

Public Function ReportSubject(ByVal caseId As String) As String
    ReportSubject = "【" & modBootNavi.AppDisplayName() & "】報告 " & caseId & " " & modUtil.NowStamp()
End Function

' BuildReportBody - 現行の報告文を流用する(err_log 末尾N行・data_dir パス
' 入り。裁定書38 §1 班E)。案件IDが空でも(現象が起きた案件が無くても)
' err_log 末尾と保存先だけは常に持てる。
Public Function BuildReportBody(ByVal caseId As String) As String
    Dim buf() As String, n As Long
    Dim ws As Object, blk As Variant, lastRow As Long, r As Long, startRow As Long
    modUtil.BufInit buf, n
    modUtil.BufAdd buf, n, modBootNavi.AppDisplayName() & " 報告"
    modUtil.BufAdd buf, n, "案件ID: " & IIf(LenB(caseId) > 0, caseId, "(未選択)")
    modUtil.BufAdd buf, n, "保存先: " & modUtil.ResolveDataDir(modConfig.GetStr("data_dir", vbNullString), ThisWorkbook.Path)
    modUtil.BufAdd buf, n, "app_version: " & modConfig.GetStr("app_version", vbNullString)
    modUtil.BufAdd buf, n, "----- err_log 末尾" & CStr(RM_TAIL_LINES) & "行 -----"
    Set ws = modCaseStore2.SheetOf("err_log")
    If Not ws Is Nothing Then
        lastRow = modCaseStore2.LastRowOf(ws)
        blk = modCaseStore2.ReadBlock(ws, lastRow)
        If Not IsEmpty(blk) Then
            startRow = lastRow - RM_TAIL_LINES + 1
            If startRow < 2 Then startRow = 2
            For r = startRow To lastRow
                modUtil.BufAdd buf, n, modNaviStore.CellValue(blk, r, "logged_at") & " " & _
                    modNaviStore.CellValue(blk, r, "err_code") & " " & _
                    modNaviStore.CellValue(blk, r, "source") & " " & _
                    modNaviStore.CellValue(blk, r, "detail")
            Next r
            If lastRow < 2 Then modUtil.BufAdd buf, n, "(記録なし)"
        Else
            modUtil.BufAdd buf, n, "(記録なし)"
        End If
    Else
        modUtil.BufAdd buf, n, "(err_logシートを読めませんでした)"
    End If
    BuildReportBody = modUtil.BufText(buf, n)
End Function

' TryOutlookDisplay - Outlookの下書きを開く(late-bound。参照設定を足さない)。
' .Display のみを呼び、.Send は絶対に呼ばない(v1.1 §6.1「自動送信は禁止」)。
' Outlook 未導入・COM生成拒否(AV等)はすべてここで False にして呼び出し側の
' フォールバックへ渡す(16章 E-70)。
Private Function TryOutlookDisplay(ByVal subjectText As String, ByVal bodyText As String, _
                                    ByVal toAddr As String, ByVal ccAddr As String) As Boolean
    On Error GoTo Failed
    Dim outlookApp As Object, mailItem As Object
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0) ' olMailItem。定数は参照設定なしのため直値
    mailItem.Subject = subjectText
    mailItem.Body = bodyText
    If LenB(toAddr) > 0 Then mailItem.To = toAddr
    If LenB(ccAddr) > 0 Then mailItem.CC = ccAddr
    mailItem.Display ' 下書きの提示のみ。.Send は呼ばない
    TryOutlookDisplay = True
    Exit Function
Failed:
    modLog.LogError "E0609", RM_SRC & ".TryOutlookDisplay", "Outlook.Application", Err.Number
    TryOutlookDisplay = False
End Function
