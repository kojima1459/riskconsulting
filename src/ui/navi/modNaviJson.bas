Attribute VB_Name = "modNaviJson"
Option Explicit
' Specification 7.1: transport envelope only. Step documents stay opaque strings.
Public Function Q(ByVal srcValue As String) As String
    Q = Chr$(34) & modJsonLite.EscapeJsonStr(srcValue) & Chr$(34)
End Function
Public Function Flag(ByVal srcValue As Boolean) As String
    If srcValue Then Flag = "true" Else Flag = "false"
End Function
Public Function IsValidJson(ByVal srcText As String) As Boolean
    Dim pos As Long
    pos = 1
    If Len(srcText) > 4000000 Then Exit Function
    If Not Consume(srcText, pos, 0) Then Exit Function
    SkipSpace srcText, pos
    IsValidJson = (pos > Len(srcText))
End Function
Public Function StringField(ByVal srcText As String, ByVal key As String) As String
    Dim raw As String
    raw = RawField(srcText, key)
    If Left$(raw, 1) <> Chr$(34) Then Exit Function
    StringField = modJsonLite.UnescapeJsonStr(Mid$(raw, 2, Len(raw) - 2))
End Function
Public Function ObjectField(ByVal srcText As String, ByVal key As String) As String
    Dim raw As String
    raw = RawField(srcText, key)
    If Left$(raw, 1) = "{" Then ObjectField = raw Else ObjectField = "{}"
End Function
Public Function RawField(ByVal srcText As String, ByVal key As String) As String
    Dim first As Long, last As Long
    If FindField(srcText, key, first, last) Then RawField = Mid$(srcText, first, last - first)
End Function
Public Function ReplaceTextField(ByVal srcText As String, ByVal key As String, ByVal srcValue As String) As String
    Dim first As Long, last As Long
    ReplaceTextField = srcText
    If FindField(srcText, key, first, last) Then
        If Mid$(srcText, first, 1) = Chr$(34) Then ReplaceTextField = Left$(srcText, first - 1) & Q(srcValue) & Mid$(srcText, last)
    End If
End Function
Private Function FindField(ByVal srcText As String, ByVal key As String, ByRef first As Long, ByRef last As Long) As Boolean
    Dim pos As Long, startKey As Long, foundKey As String
    pos = 1: SkipSpace srcText, pos
    If Mid$(srcText, pos, 1) <> "{" Then Exit Function
    pos = pos + 1: SkipSpace srcText, pos
    Do While pos <= Len(srcText)
        startKey = pos
        If Not ReadString(srcText, pos) Then Exit Function
        foundKey = modJsonLite.UnescapeJsonStr(Mid$(srcText, startKey + 1, pos - startKey - 2))
        SkipSpace srcText, pos
        If Mid$(srcText, pos, 1) <> ":" Then Exit Function
        pos = pos + 1: SkipSpace srcText, pos
        first = pos
        If Not Consume(srcText, pos, 0) Then Exit Function
        last = pos
        If foundKey = key Then FindField = True: Exit Function
        SkipSpace srcText, pos
        If Mid$(srcText, pos, 1) <> "," Then Exit Function
        pos = pos + 1: SkipSpace srcText, pos
    Loop
End Function
Private Sub SkipSpace(ByVal srcText As String, ByRef pos As Long)
    Dim ch As String
    Do While pos <= Len(srcText)
        ch = Mid$(srcText, pos, 1)
        If ch <> " " And ch <> vbCr And ch <> vbLf And ch <> vbTab Then Exit Do
        pos = pos + 1
    Loop
End Sub
Private Function ReadString(ByVal srcText As String, ByRef pos As Long) As Boolean
    Dim ch As String, i As Long
    If Mid$(srcText, pos, 1) <> Chr$(34) Then Exit Function
    pos = pos + 1
    Do While pos <= Len(srcText)
        ch = Mid$(srcText, pos, 1): pos = pos + 1
        If ch = Chr$(34) Then ReadString = True: Exit Function
        If AscW(ch) >= 0 And AscW(ch) < 32 Then Exit Function
        If ch = Chr$(92) Then
            ch = Mid$(srcText, pos, 1): pos = pos + 1
            If Len(ch) = 0 Then Exit Function
            If ch = "u" Then
                For i = 1 To 4
                    ch = LCase$(Mid$(srcText, pos, 1))
                    If Len(ch) = 0 Then Exit Function
                    If InStr(1, "0123456789abcdef", ch, 0) = 0 Then Exit Function
                    pos = pos + 1
                Next i
            ElseIf InStr(1, Chr$(34) & Chr$(92) & "/bfnrt", ch, 0) = 0 Then
                Exit Function
            End If
        End If
    Loop
End Function
Private Function Consume(ByVal srcText As String, ByRef pos As Long, ByVal depth As Long) As Boolean
    Dim ch As String, closing As String, isObject As Boolean
    Dim keyStart As Long, keyName As String, keys As Collection
    If depth > 40 Then Exit Function
    SkipSpace srcText, pos
    ch = Mid$(srcText, pos, 1)
    If ch = Chr$(34) Then Consume = ReadString(srcText, pos): Exit Function
    If ch = "{" Or ch = "[" Then
        isObject = (ch = "{")
        If isObject Then closing = "}" Else closing = "]"
        Set keys = New Collection
        pos = pos + 1: SkipSpace srcText, pos
        If Mid$(srcText, pos, 1) = closing Then pos = pos + 1: Consume = True: Exit Function
        Do
            If isObject Then
                keyStart = pos
                If Not ReadString(srcText, pos) Then Exit Function
                keyName = modJsonLite.UnescapeJsonStr(Mid$(srcText, keyStart + 1, pos - keyStart - 2))
                If Not AddKey(keys, keyName) Then Exit Function
                SkipSpace srcText, pos
                If Mid$(srcText, pos, 1) <> ":" Then Exit Function
                pos = pos + 1
            End If
            If Not Consume(srcText, pos, depth + 1) Then Exit Function
            SkipSpace srcText, pos
            ch = Mid$(srcText, pos, 1): pos = pos + 1
            If ch = closing Then Consume = True: Exit Function
            If ch <> "," Then Exit Function
            SkipSpace srcText, pos
        Loop
    End If
    If Mid$(srcText, pos, 4) = "true" Or Mid$(srcText, pos, 4) = "null" Then
        pos = pos + 4: Consume = True: Exit Function
    End If
    If Mid$(srcText, pos, 5) = "false" Then pos = pos + 5: Consume = True: Exit Function
    Consume = ReadNumber(srcText, pos)
End Function
Private Function AddKey(ByVal keys As Collection, ByVal key As String) As Boolean
    On Error GoTo Duplicate
    ' Prefix permits an empty JSON property name. VBA Collection keys compare case-insensitively.
    keys.Add True, "k:" & key
    AddKey = True
    Exit Function
Duplicate:
    AddKey = False
End Function
Private Function IsDigit(ByVal ch As String) As Boolean
    If Len(ch) = 1 Then IsDigit = (ch >= "0" And ch <= "9")
End Function
Private Function ReadNumber(ByVal srcText As String, ByRef pos As Long) As Boolean
    Dim start As Long, ch As String
    start = pos
    If Mid$(srcText, pos, 1) = "-" Then pos = pos + 1
    ch = Mid$(srcText, pos, 1)
    If ch = "0" Then
        pos = pos + 1
    Else
        If Not IsDigit(ch) Then Exit Function
        Do While IsDigit(Mid$(srcText, pos, 1)): pos = pos + 1: Loop
    End If
    If Mid$(srcText, pos, 1) = "." Then
        pos = pos + 1
        If Not IsDigit(Mid$(srcText, pos, 1)) Then Exit Function
        Do While IsDigit(Mid$(srcText, pos, 1)): pos = pos + 1: Loop
    End If
    ch = LCase$(Mid$(srcText, pos, 1))
    If ch = "e" Then
        pos = pos + 1: ch = Mid$(srcText, pos, 1)
        If ch = "+" Or ch = "-" Then pos = pos + 1
        If Not IsDigit(Mid$(srcText, pos, 1)) Then Exit Function
        Do While IsDigit(Mid$(srcText, pos, 1)): pos = pos + 1: Loop
    End If
    ReadNumber = (pos > start)
End Function
