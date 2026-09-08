VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNaviHtml
   Caption         =   "Riscon Navi"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11100
   OleObjectBlob   =   "frmNaviHtml.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "frmNaviHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Specification 4 / 8.1. Early binding needs Microsoft Internet Controls.
Private WithEvents mBrowser As SHDocVw.WebBrowser
Attribute mBrowser.VB_VarHelpID = -1
Private mBrowserHost As MSForms.Control
Private mInjected As Boolean
Private mReady As Boolean
Private mSize As Long
Private mTextScale As String
Private mOpticalZoom As Long
Public AllowClose As Boolean
Private Sub UserForm_Initialize()
    ' 裁定書34 §0.4/§1.2: 題名は config app_display_name の1箇所から。
    Me.Caption = modBootNavi.AppDisplayName()
    CycleSize
End Sub
Public Sub CycleSize()
    mSize = (mSize + 1) Mod 3
    Select Case mSize
    Case 0: Me.Width = 740: Me.Height = 520
    Case 1: Me.Width = Application.Width * 0.92: Me.Height = Application.Height * 0.88
    Case 2: Me.Width = Application.Width * 0.98: Me.Height = Application.Height * 0.95
    End Select
    If Me.Width < 640 Then Me.Width = 640
    If Me.Height < 400 Then Me.Height = 400
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    LayoutBrowser
End Sub
Public Sub EnsureBrowser()
    On Error GoTo Failed
    If Not mBrowser Is Nothing Then Exit Sub
    Set mBrowserHost = Me.Controls.Add("Shell.Explorer.2", "naviBrowser", True)
    Set mBrowser = mBrowserHost
    mBrowser.Silent = True
    LayoutBrowser
    mBrowser.Navigate "about:blank"
    Exit Sub
Failed:
    modNaviHost.HostFailure "browser_create", Err.Number
    ClearBrowser
    Err.Raise 5, "frmNaviHtml", "browser_unavailable"
End Sub
Private Sub ClearBrowser()
    On Error GoTo Done
    Set mBrowser = Nothing
    If Not mBrowserHost Is Nothing Then Me.Controls.Remove "naviBrowser"
    Set mBrowserHost = Nothing
    mReady = False: mInjected = False
Done:
End Sub
Private Sub LayoutBrowser()
    If mBrowserHost Is Nothing Then Exit Sub
    mBrowserHost.Left = 0: mBrowserHost.Top = 0
    mBrowserHost.Width = Me.InsideWidth: mBrowserHost.Height = Me.InsideHeight
End Sub
Private Sub UserForm_Resize()
    On Error GoTo Failed
    LayoutBrowser
    Exit Sub
Failed:
    modNaviHost.HostFailure "resize", Err.Number
End Sub
Private Sub mBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    On Error GoTo Failed
    Dim target As String
    target = LCase$(CStr(URL))
    If target = "vba://dispatch" Or target = "vba://dispatch/" Then
        Cancel = True
        modNaviHost.HostDispatchPending
    ElseIf Left$(target, 5) = "blob:" Then
        ' Report preview: an isolated child frame may load a host-supplied HTML blob.
        ' Top-level blob navigation stays blocked. HTTP/file navigation is never admitted.
        Cancel = (pDisp Is mBrowser)
    ElseIf target <> "about:blank" Then
        Cancel = True
    End If
    Exit Sub
Failed:
    Cancel = True
    modNaviHost.HostFailure "navigate", Err.Number
End Sub
Private Sub mBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Cancel = True
End Sub
Private Sub mBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim doc As Object, html As String
    On Error GoTo Failed
    If mInjected Then Exit Sub
    If LCase$(CStr(URL)) <> "about:blank" Then Exit Sub
    If Not pDisp Is mBrowser Then Exit Sub
    mInjected = True
    Set doc = mBrowser.Document
    html = modNaviHost.HostReadPage()
    doc.Open
    doc.Write html
    doc.Close
    mReady = True
    ApplyTextScale modConfig.GetStr("ui_font_scale", "medium")
    Exit Sub
Failed:
    modNaviHost.HostFailure "html_load", Err.Number
    mReady = False
    Application.StatusBar = "HTMLを読み込めませんでした。uiフォルダーと参照設定を確認してください。"
End Sub
Public Function TakePendingText() As String
    Dim field As Object
    If mBrowser Is Nothing Then Exit Function
    Set field = mBrowser.Document.getElementById("vbaPayload")
    If field Is Nothing Then Exit Function
    TakePendingText = CStr(field.Value)
    field.Value = ""
End Function
Public Sub DeliverResponse(ByVal json As String)
    Dim field As Object
    If mBrowser Is Nothing Then Exit Sub
    Set field = mBrowser.Document.getElementById("vbaResponse")
    If field Is Nothing Then Exit Sub
    ' DOM .value assignment preserves text; never interpolate customer data in JS or HTML.
    field.Value = json
    mBrowser.Document.parentWindow.execScript "window.naviApp.receiveFromHost();", "JavaScript"
End Sub
Public Sub ShowBusy(ByVal message As String, ByVal progressJson As String)
    If Len(progressJson) = 0 Then progressJson = "{}"
    DeliverResponse "{""ok"":true,""busy"":true,""message"":" & modNaviJson.Q(message) & ",""progress"":" & progressJson & "}"
End Sub
Public Sub ApplyTextScale(ByVal fontScale As String)
    Dim zoom As Variant
    On Error GoTo Failed
    ' HTML rem units own small/medium/large. Optical zoom only corrects WebOC DPI.
    Select Case fontScale
    Case "small": zoom = CLng(100)
    Case "large": zoom = CLng(100)
    Case Else: fontScale = "medium": zoom = CLng(100)
    End Select
    mTextScale = fontScale
    If mBrowser Is Nothing Then Exit Sub
    zoom = CLng(BaseOpticalZoom())
    mBrowser.ExecWB 63, 2, zoom
    mOpticalZoom = CLng(zoom)
    Exit Sub
Failed:
    ' 裁定書34 §1.2 / 16章 E-66: 拡大率命令(ExecWB 63)の非対応
    '   (-2147221248 = OLECMDERR_E_NOTSUPPORTED)は表示に影響しない。
    '   err_log には書かず usage_log に1行だけ残す(err_log を「見なくていい行」で
    '   埋めない)。文字の大きさは HTML 側の CSS で付いているので表示は続く。
    modLog.LogUsage "zoom_unsupported", vbNullString, fontScale & ":" & CStr(Err.Number)
End Sub
Private Function BaseOpticalZoom() As Long
    Dim scr As Object, systemDpi As Double, layoutDpi As Double, value As Double
    BaseOpticalZoom = 100
    On Error GoTo Done
    Set scr = mBrowser.Document.parentWindow.screen
    systemDpi = CDbl(scr.systemXDPI)
    layoutDpi = CDbl(scr.logicalXDPI)
    If systemDpi < 96 Or layoutDpi < 96 Then Exit Function
    ' WebOC without a DPI-aware site maps layout pixels using logical/device DPI.
    ' Normalize CSS pixels to the Windows logical size, independently of text size.
    value = 100# * (systemDpi / 96#) * (layoutDpi / 96#)
    If value > 1000# Then value = 1000#
    BaseOpticalZoom = CLng(value)
Done:
End Function
Public Sub RequestClose()
    On Error GoTo Failed
    If modNaviHost.HostIsBusy() Then Exit Sub
    If mReady Then
        mBrowser.Document.parentWindow.execScript "window.naviApp.requestClose();", "JavaScript"
    Else
        modNaviHost.CloseNaviTool
    End If
    Exit Sub
Failed:
    modNaviHost.HostFailure "request_close", Err.Number
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If AllowClose Then Exit Sub
    Cancel = True
    RequestClose
End Sub
Private Sub UserForm_Terminate()
    Set mBrowser = Nothing
    Set mBrowserHost = Nothing
End Sub
Public Property Get IsReady() As Boolean
    IsReady = mReady
End Property
Public Property Get DocumentMode() As Long
    If mReady Then DocumentMode = mBrowser.Document.DocumentMode
End Property
Public Property Get HtmlLength() As Long
    If mReady Then HtmlLength = Len(mBrowser.Document.DocumentElement.outerHTML)
End Property
Public Property Get TextScale() As String
    TextScale = mTextScale
End Property
Public Property Get OpticalZoom() As Long
    OpticalZoom = mOpticalZoom
End Property
