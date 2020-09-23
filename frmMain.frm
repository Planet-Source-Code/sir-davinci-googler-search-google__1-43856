VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Googler"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrev 
      Caption         =   "< Prev"
      Enabled         =   0   'False
      Height          =   325
      Left            =   6600
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   7680
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   2
      Top             =   1800
      Width           =   1035
   End
   Begin VB.TextBox txtQuery 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "davinci"
      Top             =   1800
      Width           =   7335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   120
      Width           =   5835
      Begin MSWinsockLib.Winsock Socket 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemoteHost      =   "www.google.com"
         RemotePort      =   80
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   480
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":29B9
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title:"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description:"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "URL:"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "double click on result to view site"
      Height          =   195
      Left            =   6240
      TabIndex        =   7
      Top             =   120
      Width           =   2340
   End
   Begin VB.Label lbResults 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Idle"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4480
      Width           =   330
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
    strPacket = strNext
    With Socket
        .Close
        .Connect
    End With
    cmdNext.Enabled = False
End Sub

Private Sub cmdPrev_Click()
    strPacket = strPrev
    With Socket
        .Close
        .Connect
    End With
    cmdPrev.Enabled = False
End Sub

Private Sub cmdSearch_Click()
    strPacket = "/palm?q=" & txtQuery
    With Socket
        .Close
        .Connect
    End With
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
    OpenURL ListView1.SelectedItem.SubItems(2)
End Sub

Private Sub Socket_Connect()
Dim strGoogle As String
    
    '*** Space %20 Support ***
    If InStr(strPacket, " ") Then
        strPacket = Replace(strPacket, " ", "%20")
    End If
    '***
    
    strGoogle = "GET " & strPacket & " HTTP/1.0" & vbNewLine
    strGoogle = strGoogle & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbNewLine
    strGoogle = strGoogle & "Accept-Language: en-us" & vbNewLine
    strGoogle = strGoogle & "Accept-Encoding: gzip, deflate" & vbNewLine
    strGoogle = strGoogle & "Host: www.google.com" & vbNewLine
    strGoogle = strGoogle & "Connection: Keep-Alive" & vbNewLine & vbNewLine
    
    ListView1.ListItems.Clear
    
Socket.SendData strGoogle
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim strInfo As String
Socket.GetData strInfo, vbString
               
        Dim strData As Variant
        strData = Split(strInfo, "<p>")
        
        Dim ListX As ListItem
        
        '*** No Matches ***
        For g = 0 To UBound(strData)
            If InStr(strData(g), "Your search - <b>" & txtQuery & "</b> - did not match any documents.") Then
                Set ListX = ListView1.ListItems.Add(, , "No Matches Found...")
            End If
        Next g
        '***
        
        '*** Results ***
        For x = 0 To 0
            If InStr(strData(i), "Results") Then
                results = "Results " & Trim(GetStringBetween(strData(x), "Results", "<hr>"))
                results = Replace(results, "<b>", "")
                results = Replace(results, "</b>", "")
                lbResults.Caption = results
            End If
        Next x
        '***
        
        '*** Query Matches ***
        For i = 1 To UBound(strData) '- 2
        If InStr(strData(i), "<a href=http://") Then
            Set ListX = ListView1.ListItems.Add(, , , , 1)
            a = GetStringBetween(strData(i), "<a href=http://", ">")

            Dim strDec As String
            strDec = GetStringBetween(strData(i), " - ", " - ")
            strDec = CleanUp(strDec)
            
            Dim strTitle As String
            strTitle = GetStringBetween(strData(i), a & ">", "</a>")
            If InStr(strTitle, "<b>") Then
                strTitle = Replace(strTitle, "<b>", "")
                strTitle = Replace(strTitle, "</b>", "")
            End If
            
            ListX.Text = strTitle
            ListX.SubItems(1) = strDec
            ListX.SubItems(2) = "http://" & a
        End If
        Next i
        '***
        
        '*** Next & Prev ***
        For q = 0 To UBound(strData)
        If InStr(strData(q), "sa=N>Next</") Then
            strNext = GetStringBetween(strData(q), "<A HREF=", "Next</A>")
            strNext = Replace(strNext, ">", "")
            cmdNext.Enabled = True
        End If
        If InStr(strData(q), "sa=N>Prev</") Then
            strPrev = GetStringBetween(strData(q), "<A HREF=", "Prev</A>")
            strPrev = Replace(strPrev, ">", "")
            cmdPrev.Enabled = True
        End If
        Next q
        '***
End Sub
