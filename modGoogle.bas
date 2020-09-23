Attribute VB_Name = "modGoogle"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public sURL As String
Public strPacket As String
Public strNext As String
Public strPrev As String

Public Function GetStringBetween(ByVal Str As String, ByVal str1 As String, ByVal str2 As String, Optional ByVal st As Long = 0) As String
    On Error Resume Next
    Dim s1, s2, s, l As Long
    Dim foundstr As String
    
    s1 = InStr(st + 1, Str, str1, vbTextCompare)
    s2 = InStr(s1 + 1, Str, str2, vbTextCompare)
    
    If s1 = 0 Or s2 = 0 Or IsNull(s1) Or IsNull(s2) Then
        foundstr = Str
    Else
        s = s1 + Len(str1)
        l = s2 - s
        foundstr = Mid(Str, s, l)
    End If
    
    GetStringBetween = foundstr
End Function

Sub OpenURL(URL As String)
'Open URL in Default Browser
    ShellExecute hwnd, "open", URL, vbNullString, vbNullString, conSwNormal
End Sub

Sub CenterForm(frm As Form)
'Centers a Form
    If frm.WindowState = 0 Then
        frm.Top = Screen.Height / 2 - frm.Height / 2
        frm.Left = Screen.Width / 2 - frm.Width / 2
    End If
End Sub

Public Function CleanUp(sData As String)

    If InStr(sData, LCase("&amp;")) Then _
        sData = Replace(sData, LCase("&amp;"), "&")
        
    If InStr(sData, LCase("&quot;")) Then _
        sData = Replace(sData, LCase("&quot;"), Chr(34))
        
    If InStr(sData, LCase("&nbsp;")) Then _
        sData = Replace(sData, LCase("&nbsp;"), " ")
        
    If InStr(sData, LCase("&copy;")) Then _
        sData = Replace(sData, LCase("&copy;"), "©")

    If InStr(sData, LCase("&trade;")) Then _
        sData = Replace(sData, LCase("&trade;"), "™")

    '<b>BOLD</b>
    If InStr(sData, "<b>") Then
        sData = Replace(sData, "<b>", "")
        sData = Replace(sData, "</b>", "")
    End If
    
    '<a href=>LINK</a>
    If InStr(sData, "</a>") Then _
        sData = Replace(sData, "</a>", "")
    
    If InStr(sData, "<a href=") Then
        temp$ = GetStringBetween(sData, "<a href=", ">")
        sData = Replace(sData, "<a href=", "")
        sData = Replace(sData, ">", "")
        sData = Replace(sData, temp$, "")
        sData = temp$ & " - " & sData
    End If
    
CleanUp = sData
End Function
