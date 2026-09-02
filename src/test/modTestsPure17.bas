Attribute VB_Name = "modTestsPure17"
Option Explicit

' ============================================================================
' modTestsPure17 - W7 統合(17章 T-57)の純層テスト(17章§4-1 層(a))
' ----------------------------------------------------------------------------
' 執筆方針: 期待値は 15章§0.7 の切詰め表**だけ**から手計算した。実装の出力を
'   見てから期待値を合わせることは禁止(17章§1)。
'
' 対象と根拠(15章§0.7 の6段。裁定書25 S6 が「順2=事故事例」を足した):
'   W7G TrimPlan 12本  6段の半減順に**正負1本ずつ**。
'     負= その段の直前で予算に収まるので**その段は発火しない**
'     正= 1字だけ足りないので**その段まで発火する**
'
' 期待値の手計算(15章§0.7「1段ずつ適用し、そのつど総量を再計算する」):
'   素材 = 行数 5/5/10/60/20/20(成功事例/事故事例/型/メニュー/種目/リスクライブラリ)
'          文字数は各1000字 -> 総量6000字。文字数は行数に比例と見積もる。
'   半減は端数切上げ・下限 0/0/0/5/5/5。
'     段1 成功事例 5->3   1000->600  総量 6000 -> 5600
'     段2 事故事例 5->3   1000->600  総量 5600 -> 5200
'     段3 型       10->5  1000->500  総量 5200 -> 4700
'     段4 メニュー 60->30 1000->500  総量 4700 -> 4200
'     段5 種目     20->10 1000->500  総量 4200 -> 3700
'     段6 リスクL  20->10 1000->500  総量 3700 -> 3200
'   よって「段kの直前の総量」を予算に与えると段kは発火せず(負)、そこから
'   1字減らすと段kまで発火する(正)。
'
' グループ単位の失敗隔離: modTestsPure.bas と同じ On Error GoTo 方式。
' **テストを増減したら wintest/tests_expected.txt を必ず同時に更新すること**。
' ============================================================================

Public Sub RunAll()
    On Error GoTo FA
    T_W7G_TrimPlan6
WDone:
    Exit Sub
FA:
    GroupFail "W7G TrimPlan(15章§0.7 の6段)"
    Resume WDone
End Sub

Private Sub GroupFail(ByVal grpName As String)
    modTestRunner.Check grpName & "(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & _
        ") ※未実装/未注入の可能性"
End Sub

Private Sub ChkS(ByVal nm As String, ByVal act As String, ByVal want As String)
    modTestRunner.Check nm, (act = want), "期待=[" & want & "] 実際=[" & act & "]"
End Sub

' ============================================================================
' W7G TrimPlan(15章§0.7)
' ============================================================================
Private Sub T_W7G_TrimPlan6()
    Dim c(0 To 11) As Long
    Dim i As Long

    c(0) = 5
    c(1) = 5
    c(2) = 10
    c(3) = 60
    c(4) = 20
    c(5) = 20
    For i = 6 To 11
        c(i) = 1000
    Next i

    ' 段1 成功事例(表の順1)
    ChkS "Test_W7G_01_段1は総量6000ちょうどなら発火しない_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 6000)), "5,5,10,60,20,20"
    ChkS "Test_W7G_02_段1は5999で成功事例だけ半減_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 5999)), "3,5,10,60,20,20"

    ' 段2 事故事例(表の順2。裁定書25 S6 で新設)
    ChkS "Test_W7G_03_段2は総量5600ちょうどなら発火しない_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 5600)), "3,5,10,60,20,20"
    ChkS "Test_W7G_04_段2は5599で事故事例まで半減_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 5599)), "3,3,10,60,20,20"

    ' 段3 型ライブラリ(表の順3)
    ChkS "Test_W7G_05_段3は総量5200ちょうどなら発火しない_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 5200)), "3,3,10,60,20,20"
    ChkS "Test_W7G_06_段3は5199で型ライブラリまで半減_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 5199)), "3,3,5,60,20,20"

    ' 段4 メニュー(表の順4)
    ChkS "Test_W7G_07_段4は総量4700ちょうどなら発火しない_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 4700)), "3,3,5,60,20,20"
    ChkS "Test_W7G_08_段4は4699でメニューまで半減_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 4699)), "3,3,5,30,20,20"

    ' 段5 種目(表の順5)
    ChkS "Test_W7G_09_段5は総量4200ちょうどなら発火しない_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 4200)), "3,3,5,30,20,20"
    ChkS "Test_W7G_10_段5は4199で種目まで半減_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 4199)), "3,3,5,30,10,20"

    ' 段6 リスクライブラリ(表の順6)
    ChkS "Test_W7G_11_段6は総量3700ちょうどなら発火しない_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 3700)), "3,3,5,30,10,20"
    ChkS "Test_W7G_12_段6は3699でリスクライブラリまで半減_15章§0.7", _
        PlanText(modPipeline4.TrimPlan(c, 3699)), "3,3,5,30,10,10"
End Sub

' TrimPlan の戻り(6要素・0始まり)を "a,b,c,d,e,f" の1行にする。6要素でない場合は
' その事実が期待値との差分として現れるように件数を出す。
Private Function PlanText(ByVal p As Variant) As String
    Dim i As Long
    Dim n As Long
    Dim s As String

    n = 0
    On Error Resume Next
    n = UBound(p) - LBound(p) + 1
    On Error GoTo 0
    If n <> 6 Then
        PlanText = "(要素数=" & n & ")"
        Exit Function
    End If
    s = ""
    For i = 0 To 5
        If i > 0 Then s = s & ","
        s = s & p(LBound(p) + i)
    Next i
    PlanText = s
End Function
