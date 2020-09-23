VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dCoolXDCC"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7995
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox theuserlist 
      Height          =   450
      Left            =   2760
      TabIndex        =   24
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Timer reconnecttimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   360
   End
   Begin VB.TextBox txtPM 
      BackColor       =   &H80000004&
      Height          =   735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Text            =   "form1.frx":628A
      Top             =   4920
      Width           =   5895
   End
   Begin VB.TextBox thisversioninfo 
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Text            =   "form1.frx":629D
      Top             =   1680
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Timer updatetimer 
      Interval        =   1000
      Left            =   4800
      Top             =   120
   End
   Begin VB.TextBox thisversion 
      Height          =   285
      Left            =   240
      TabIndex        =   20
      Text            =   "2.0"
      Top             =   1320
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Timer pausexdcc 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   65000
      Left            =   3120
      Top             =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "XD&CC Settings"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "XDCC"
      Height          =   3735
      Left            =   6000
      TabIndex        =   8
      Top             =   2760
      Width           =   1935
      Begin VB.CommandButton Command6 
         Caption         =   "Add Channel"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   18
         Top             =   3120
         Width           =   1695
      End
      Begin VB.ListBox List5 
         Height          =   840
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ListBox List4 
         Height          =   840
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ListBox List3 
         Height          =   840
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ad Channels:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Packs:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Users:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Channels"
      Height          =   2655
      Left            =   6000
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add Channel"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         Height          =   1035
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         Height          =   645
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Timer xdcctimer 
      Interval        =   60000
      Left            =   1680
      Top             =   0
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Basic &Settings"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox status2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "form1.frx":630D
      Top             =   0
      Width           =   5895
   End
   Begin VB.TextBox Status 
      BackColor       =   &H80000004&
      Height          =   4245
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "form1.frx":6325
      Top             =   600
      Width           =   5895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuNull1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPM 
         Caption         =   "Send Private Message"
      End
      Begin VB.Menu mnuNull2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuXDCC 
      Caption         =   "X&DCC"
      Begin VB.Menu mnuXDCCEn 
         Caption         =   "&Enable"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuXDCCDis 
         Caption         =   "D&isable"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause for..."
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuBasicSettings 
         Caption         =   "&Basic Settings"
      End
      Begin VB.Menu mnuXDCCSettings 
         Caption         =   "&XDCC Settings"
      End
   End
   Begin VB.Menu mnuUpdate 
      Caption         =   "&Update"
      Begin VB.Menu mnuCheckUpdate 
         Caption         =   "&Check for Updates"
      End
      Begin VB.Menu mnuCheckAddons 
         Caption         =   "Check for Add&ons"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelpmenu 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "dCoolXDCC Help"
      End
      Begin VB.Menu mnutrouble 
         Caption         =   "dCoolXDCC Troubleshooting"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuWebsite 
         Caption         =   "&Visit Website - dCool101d.tk"
      End
      Begin VB.Menu mnuAboutdCoolXDCC 
         Caption         =   "About d&CoolXDCC"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, fLength, ret                                '// Declare Variables
Dim Buffer As String                              '// Declare Buffer
Dim bSize As Long
Dim wanttodisconnect As Boolean
Public adtime As Integer

Private Sub Command1_Click()
wanttodisconnect = False
On Error Resume Next
Winsock1.RemoteHost = Form2.Text1.Text
Winsock1.RemotePort = Form2.Text2.Text
'Tell winsock to connect
Winsock1.LocalPort = 1560
Winsock1.Connect
mnuPM.Enabled = True
intmaxsends = 5
Text3.Enabled = True
PingDone = 0
Command4.Enabled = False
Frame1.Enabled = True
'Disable button 1
Command1.Enabled = False
List1.Enabled = True
List2.Enabled = True
Form2.Text1.Enabled = False
Form2.Text2.Enabled = False
Form2.txtnick.Enabled = False
Form2.txtname.Enabled = False
Form2.txtpassword.Enabled = False
Form2.txtemail.Enabled = False
Form2.txtemail2.Enabled = False
Form2.Text3.Enabled = False
Form2.chkreg.Enabled = False
Form2.Command1.Enabled = False
'Enable button 2
Command2.Enabled = True
For X = 1 To intmaxsends
    Load SendFileForm(X)
Next
End Sub

Private Sub Command2_Click()
mnuPM.Enabled = False
Text3.Enabled = False
PingDone = 0
Command4.Enabled = True
Frame1.Enabled = False
wanttodisconnect = True
Winsock1.Close
List1.Enabled = False
List2.Enabled = False
Dim X As Long
For X = 1 To List2.ListCount
List2.Selected(X - 1) = True
If List2.Text <> "-Server-" Then
List1.AddItem List2.Text
End If
Next
List2.Clear
Form2.Text1.Enabled = True
Form2.Text2.Enabled = True
Form2.txtnick.Enabled = True
Form2.txtname.Enabled = True
Form2.txtpassword.Enabled = True
Form2.txtemail.Enabled = True
Form2.txtemail2.Enabled = True
Form2.Text3.Enabled = True
Form2.chkreg.Enabled = True
Form2.Command1.Enabled = True
'Enable button  1
Command1.Enabled = True
'Disable button 2
Command2.Enabled = False
List3.Clear
End Sub

Private Sub Command3_Click()
Dim X As Long
If Left(Text4.Text, 1) <> "#" Then
    Text4.Text = "#" & Text4.Text
End If
List1.AddItem Text4.Text
Text4.Text = ""
Open "chnllst.dat" For Output As #9
For X = 1 To List1.ListCount
List1.Selected(X - 1) = True
Write #9, List1.Text
Next
Close #9
End Sub

Private Sub Command4_Click()
    Me.Enabled = False
    Form2.Show
End Sub

Private Sub Command5_Click()
    Me.Enabled = False
    Form3.Show
End Sub

Private Sub Command6_Click()
    List5.AddItem Text10.Text
    Text10.Text = ""
End Sub

Private Sub Form_Load()
Load frmPM
List2.AddItem "-Server-"
Load frmtray
Load frmMessage
XDCCEnabled = True
Dim X As Long
Load Form3

Dim email2 As String
Dim theserver As String
Dim theport As String
Dim xdccdir As String
strChannel = ""
'Set the text in Text1 to your IP Address
On Error GoTo nextpart
Open "chnllst.dat" For Input As #9
Dim channeladdtolist As String
Do While Not EOF(9)
Input #9, channeladdtolist
List1.AddItem channeladdtolist
Loop
Close #9
nextpart:
    Me.Enabled = False
    Form2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
If blnexit = False Then
    Cancel = 1
    Me.Hide
Else
    Unload frmtray
    End
End If
End Sub

Private Sub List2_Click()
    If List1.Text = "-Server-" Then
    strChannel = ""
    Else
    strChannel = List2.Text & " :"
    Text10.Text = List2.Text
    End If
    
End Sub

Private Sub List1_DblClick()
If Winsock1.State <> sckClosed Then
Winsock1.SendData "JOIN " & List1.Text & vbCrLf
Status.Text = Status.Text & "***Joined " & List1.Text & vbNewLine
List2.AddItem List1.Text
List1.RemoveItem (List1.ListIndex)
List2.Selected(List2.ListCount - 1) = True
End If
End Sub

Private Sub List2_DblClick()
If Winsock1.State <> sckClosed Then
If List1.Text <> "-Server-" Then
Winsock1.SendData "PART " & List2.Text & vbCrLf
Status.Text = Status.Text & "***Parted " & List2.Text & vbNewLine
List1.AddItem List2.Text
List2.RemoveItem (List2.ListIndex)
End If
End If
End Sub

Private Sub List5_DblClick()
    List5.RemoveItem (List5.ListIndex)
End Sub

Private Sub mnuAboutdCoolXDCC_Click()
    frmAbout.Show
End Sub

Private Sub mnuBasicSettings_Click()
    Call Command4_Click
End Sub

Private Sub mnuCheckUpdate_Click()
    Me.Enabled = False
    frmUpdate.Show
End Sub

Private Sub mnuConnect_Click()
    Call Command1_Click
End Sub

Private Sub mnuDisconnect_Click()
    Call Command2_Click
End Sub

Public Sub mnuExit_Click()
    Call Command2_Click
    blnexit = True
    Unload Me
End Sub

Private Sub mnuHelp_Click()
    MsgBox "Entering help mode, click the exit button in the upper-right corner to return to normal mode", , "Entering Help Mode"
    Form1.Hide
    frmHelp.Show
End Sub

Private Sub mnuPause_Click()
    Dim intNull As Integer
    intNull = Val(InputBox("Pause for how many seconds?", "XDCC Pause")) * 1000
    If intNull > 0 Then
        pausexdcc.Interval = intNull
        pausexdcc.Enabled = True
    End If
End Sub

Private Sub mnuPM_Click()
    frmPM.Show
End Sub

Private Sub mnuWebsite_Click()
    Const conSwNormal = 1
    ShellExecute Me.hwnd, "open", "http://mywebpages.comcast.net/khameneh/", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub mnuXDCCDis_Click()
    mnuXDCCDis.Checked = True
    mnuXDCCEn.Checked = False
    mnuPause.Enabled = False
End Sub

Private Sub mnuXDCCEn_Click()
    mnuXDCCDis.Checked = False
    mnuXDCCEn.Checked = True
    mnuPause.Enabled = True
End Sub

Private Sub mnuXDCCSettings_Click()
    Me.Enabled = False
    Form3.Show
End Sub

Private Sub pausexdcc_Timer()
    Call mnuXDCCEn_Click
    Me.Enabled = False
End Sub

Private Sub status2_Change()
status2.SelStart = Len(status2.Text) - 1
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Winsock1.State <> sckClosed Then
    If KeyAscii = 13 Then
        Winsock1.SendData "PRIVMSG " & strChannel & Text3.Text & vbCrLf
        Status.Text = Status.Text & "-" & "<" & strNick & ">" & strChannel & "  " & Text3.Text & vbNewLine
        Text3.Text = ""
        Status.SelStart = Len(Status)
    End If
End If
End Sub

Private Sub theuserlist_Click()
Clipboard.SetText theuserlist.Text
End Sub

Private Sub Timer1_Timer()
    Status.Text = lastmessage & "-----Cleared-----" & vbNewLine
    status2.Text = "-----Cleared-----"
End Sub

Private Sub reconnecttimer_Timer()
    Call Command1_Click
    reconnecttimer.Enabled = False
End Sub

Private Sub updatetimer_Timer()
Dim Filename As String
Dim tmpStrg As String
Dim op As SHFILEOPSTRUCT
Dim id As Long

    tmpStrg = Dir$(App.Path & "\*.exe")
    If tmpStrg <> "" Then
            Filename = tmpStrg
        If Left(Filename, 17) = "dCoolXDCC_Update_" Then
                If Mid(Filename, 17, 3) > Val(thisversion.Text) Then
                    On Error Resume Next
                    DestroyFile "dCoolXDCC.exe"
                    With op
                        .wFunc = FO_COPY ' Set function
                        .pTo = "dCoolXDCC.exe" ' Set new path
                        .pFrom = Filename ' Set current path
                        .fFlags = FOF_SIMPLEPROGRESS + FOF_SILENT + FOF_NOCONFIRMATION
                    End With
                    SHFileOperation op
                    id = Shell("dCoolXDCC.exe", 1)
                    DestroyFile Filename
                    blnexit = True
                    End
                End If
            End If
        tmpStrg = Dir$
        While Len(tmpStrg) > 0
                Filename = tmpStrg
            If Left(Filename, 17) = "dCoolXDCC_Update_" Then
                If Mid(Filename, 17, 3) > Val(thisversion.Text) Then
                    On Error Resume Next
                    DestroyFile "dCoolXDCC.exe"
                    With op
                        .wFunc = FO_COPY ' Set function
                        .pTo = "dCoolXDCC.exe" ' Set new path
                        .pFrom = Filename ' Set current path
                        .fFlags = FOF_SIMPLEPROGRESS + FOF_SILENT + FOF_NOCONFIRMATION
                    End With
                    SHFileOperation op
                    id = Shell("dCoolXDCC.exe", 1)
                    DestroyFile Filename
                    blnexit = True
                    End
                End If
            End If
            tmpStrg = Dir$
        Wend
    Else
    End If
updatetimer.Enabled = False
End Sub

Private Sub updatetimer2_Timer()

End Sub

Private Sub Winsock1_Close()
If wanttodisconnect = False Then
reconnecttimer.Enabled = True
End If

Call Command2_Click

Dim X As Long
For X = 1 To intmaxsends
SendFileForm(1).tcpSend(SendFileForm(1).Tag).Close
SendFileForm(1).inuse = False
Next
End Sub

Private Sub Winsock1_Connect()
    strNick = Form2.txtnick.Text
    strName = Form2.txtname.Text
    strEmail = Form2.txtemail.Text
    strPassword = Form2.txtpassword.Text
    strQuit = "...XDCC signing off!"
  'When we are connected tell the user
  If Winsock1.State <> sckClosed Then
  Winsock1.SendData "NICK " & strNick & vbCrLf
  Winsock1.SendData "USER " & strEmail & " " & Chr(54) & strName & Chr(54) & " " & Chr(54) & Winsock1.LocalIP & Chr(54) & " :" & Name & vbCrLf
  End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strLine As String
    Winsock1.GetData strData
    strLines = Split(strData, vbCrLf)
    Dim X As Long
    For X = 0 To UBound(strLines)
        strLine = strLines(X)
        Call TextArrival(bytesTotal, strLine)
    Next
End Sub

Private Sub TextArrival(bytesTotal As Long, strData As String)
    Dim Retval As Boolean
    Dim X As Long
    Dim Y As Integer
    Dim z As Long
    Dim strString As String
    Static PingDone As Integer
    Dim user As String
    Dim message As String
    Dim gotmessage As Boolean
    Dim strData2 As String
    gotmessage = False
    Dim gotnick As Boolean
    gotnick = False
    status2.Text = status2.Text & strData & vbNewLine
    strData2 = strData
    strData = IRCStripColor(strData)
    strData = Replace(strData, "", "")
    If UCase(Right(strData, 20)) = "NO SUCH NICK/CHANNEL" Then
        nickorchannel = Split(strData, " ")
        txtPM.Text = txtPM.Text & "- No such Nick/Channel: " & nickorchannel(UBound(nickorchannel) - 3) & " -" & vbNewLine
        frmPM.txtPM.Text = txtPM.Text
        frmPM.txtPM.SelStart = Len(frmPM.txtPM.Text) - 1
    End If
    On Error Resume Next
    For z = 1 To Len(strData)
        If Mid(strData, z, 1) = "!" And z < 20 And gotnick = False Then
            user = Left(strData, z - 1)
            user = "<" & Right(user, Len(user) - 1) & ">"
            gotnick = True
        End If
        If UCase(Mid(strData, z, 1)) = "#" And gotmessage = False Then
            message = Right(strData, Len(strData) - z + 1)
            gotmessage = True
        End If
    Next
    
    If UCase(Right(strData, Len("please choose a different nick."))) = UCase("please choose a different nick.") Or UCase(Right(strData, Len("you have not registered."))) = UCase("you have not registered.") Then
        If Winsock1.State <> sckClosed Then
        If Form2.chkreg.Value = 1 Then
        Winsock1.SendData "PRIVMSG NickServ :REGISTER " & strPassword & " " & strEmail & "@" & txtemail2.Text & vbCrLf
        End If
        Winsock1.SendData "PRIVMSG NickServ :IDENTIFY " & strPassword & vbCrLf
        
        For X = 1 To List2.ListCount
        List2.Selected(X - 1) = True
        If List2.Text <> "-Server-" Then
        Winsock1.SendData "JOIN " & List2.Text & vbCrLf
        Status.Text = Status.Text & "***Joined " & List2.Text & vbNewLine
        End If
        Next
        
        List1.Enabled = True
        List2.Enabled = True
        End If
    End If
    
    For Y = 1 To Len(message)
        If Mid(message, Y, 8 + Len(strNick)) = " 353 " & strNick & " = " Then
            z = (Y + 8 + Len(strNick))
            Dim gotfirst As Integer
            gotfirst = 0
            For X = (Y + 8 + Len(strNick)) To Len(message)
                If Mid(message, X, 1) = " " Then
                    If gotfirst = 2 Then
                        Dim thenickadding As String
                        thenickadding = Replace((Mid(message, z + 1, X - z - 1)), "@", "")
                        thenickadding = Replace(thenickadding, "%", "")
                        thenickadding = Replace(thenickadding, "+", "")
                        theuserlist.AddItem thenickadding
                    Else
                        gotfirst = gotfirst + 1
                    End If
                    z = X
                End If
            Next
            theuserlist.RemoveItem (theuserlist.ListCount - 1)
            theuserlist.RemoveItem (theuserlist.ListCount - 1)
            theuserlist.RemoveItem (theuserlist.ListCount - 1)
            theuserlist.RemoveItem (theuserlist.ListCount - 1)
            theuserlist.RemoveItem (theuserlist.ListCount - 1)
            theuserlist.RemoveItem (theuserlist.ListCount - 1)
            theuserlist.RemoveItem (theuserlist.ListCount - 1)
            theuserlist.Refresh
            GoTo gotnicklist
        End If
    Next
    
    If Left(message, 1) = "#" Then
        For X = 1 To List2.ListCount
            If Right(message, Len(List2.Text)) = List2.Text And Len(message) = Len(List2.Text) + 1 Then
                Dim donewith As Boolean
                For Y = 1 To theuserlist.ListCount
                    theuserlist.Selected(Y - 1) = True
                    If theuserlist.Text = Mid(user, 2, Len(user) - 2) Then
                        donewith = True
                        theuserlist.RemoveItem (Y - 1)
                    End If
                Next
                If donewith <> True Then
                    theuserlist.AddItem Mid(user, 2, Len(user) - 2)
                End If
                GoTo gotnicklist
            End If
        Next
    End If
    
    lastmessage = user & message & vbNewLine
    
gotnicklist:
    If user = "" Then
        lastmessage = ""
    End If
    If message = "" Then
        lastmessage = ""
    End If
    
    If lastmessage <> "" Then
        Status.Text = Status.Text & lastmessage
    End If
    
    strLines = Split(strData, vbCrLf)
    strLineParts = Split(strLines(i), " ")
    
    For X = 0 To UBound(strLines)
        strLineParts = Split(strLines(X), " ")
        If UBound(strLineParts) <> -1 Then
            Select Case strLineParts(0)
            Case "PING"
            Winsock1.SendData ("PONG " & Right(strLines(X), Len(strLines(X)) - Len("PING "))) & vbCrLf
            status2.Text = status2.Text & ("PONG " & Right(strLines(X), Len(strLines(X)) - Len("PING "))) & vbCrLf
            End Select
        End If
    Next X
    
    Status.SelStart = Len(Status.Text)
    
    Dim intNull As Integer
    Dim strNull1 As String
    Dim strNull2 As String
    Dim theplace As String
    For X = 1 To Len(strData2)
        If UCase(Mid(strData2, X, 24)) = "DCC RESUME FILE.EXT 200" Then
        For z = 1 To Len(strData)
            If Mid(strData2, z, 1) = "@" Then
                For Y = z To Len(strData)
                    If Mid(strData, Y, 1) = " " Then
                    theplace = Mid(strData2, z + 1, Y - z - 1)
                    MsgBox theplace
                    GoTo donewiththeplace
                    End If
                Next
            End If
        Next
donewiththeplace:
            If Val(Mid(strData2, X + 24, 1)) <= intmaxsends Then
                For intnull2 = X + 26 To Len(strData2)
                    If Mid(strData2, intnull2, 1) = "" Then
                        SendFileForm(Val(Mid(strData2, X + 24, 1)) + 1).ByteSent = Val(Mid(strData2, X + 26, intnull2 - X - 26)) + 1
                    End If
                Next
                Winsock1.SendData ("PRIVMSG " & SendFileForm(Val(Mid(strData2, X + 24, 1)) + 1).Nicktosend & " 200" & Val(Mid(strData2, X + 24, 1)) & Val(Mid(strData2, X + 26, intnull2 - X - 26)) & "")
                Status.SelStart = Len(Status)
            End If
            GoTo thenextpart
        End If
    Next
    
thenextpart:
On Error Resume Next
    If mnuXDCCEn.Checked = True Then
    For X = 1 To Len(strData)
        strString = UCase(Mid(strData, Len(strData) - X + 1, 13))
        If UCase(strString) = " :XDCC SEND #" Then
            Open "tmpfilecnfg.dat" For Input As #3
            For Y = 1 To intmaxsends
                If SendFileForm(Y).inuse = False Then
                    If Mid(strData, Len(strData) - X + 1 + 14, 1) <> 1 And Mid(strData, Len(strData) - X + 1 + 14, 1) <> 2 And Mid(strData, Len(strData) - X + 1 + 14, 1) <> 3 And Mid(strData, Len(strData) - X + 1 + 14, 1) <> 4 And Mid(strData, Len(strData) - X + 1 + 14, 1) <> 5 And Mid(strData, Len(strData) - X + 1 + 14, 1) <> 6 And Mid(strData, Len(strData) - X + 1 + 14, 1) <> 7 And Mid(strData, Len(strData) - X + 1 + 14, 1) <> 8 And Mid(strData, Len(strData) - X + 1 + 14, 1) <> 9 And Mid(strData, Len(strData) - X + 1 + 14, 1) <> 0 Then
                        SendFileForm(Y).Packnumber = Mid(strData, Len(strData) - X + 1 + 13, 1)
                    Else
                        If Mid(strData, Len(strData) - X + 1 + 15, 1) <> 1 And Mid(strData, Len(strData) - X + 1 + 15, 1) <> 2 And Mid(strData, Len(strData) - X + 1 + 15, 1) <> 3 And Mid(strData, Len(strData) - X + 1 + 15, 1) <> 4 And Mid(strData, Len(strData) - X + 1 + 15, 1) <> 5 And Mid(strData, Len(strData) - X + 1 + 15, 1) <> 6 And Mid(strData, Len(strData) - X + 1 + 15, 1) <> 7 And Mid(strData, Len(strData) - X + 1 + 15, 1) <> 8 And Mid(strData, Len(strData) - X + 1 + 15, 1) <> 9 And Mid(strData, Len(strData) - X + 1 + 15, 1) <> 0 Then
                            SendFileForm(Y).Packnumber = Mid(strData, Len(strData) - X + 1 + 13, 2)
                        Else
                            SendFileForm(Y).Packnumber = Mid(strData, Len(strData) - X + 1 + 13, 3)
                        End If
                    End If
                        Do While Not EOF(1)
                            Input #3, intNull, strNull1, strNull2
                            If intNull = SendFileForm(Y).Packnumber And SendFileForm(Y).inuse = False Then
                                SendFileForm(Y).Filetosend = strNull2
                                SendFileForm(Y).Filename = strNull1
                                'put in file size for lblfilesize(x)
                                SendFileForm(Y).Nicktosend = getnicktosend(strData)
                                List3.AddItem (SendFileForm(Y).Nicktosend & "-" & Val(SendFileForm(Y).Packnumber))
                                Call SendFileForm(Y).sendfiletonick
                                SendFileForm(Y).inuse = True
                                Close #3
                                GoTo gotdone
                            End If
                        Loop
                    
                End If
                If Y = intmaxsends And SendFileForm(intmaxsends).inuse = True Then
                    Winsock1.SendData "PRIVMSG " & Nicktosend(Y) & " : All slots taken, please come back later!" & vbCrLf
                    Close #3
                    GoTo gotdone
                End If
            Next
gotdone:
            
            GoTo gotdone2
        End If
    Next

    
    Dim thenick As String
    For X = 1 To Len(strData)
        strString = UCase(Mid(strData, Len(strData) - X + 1, 11))
        If strString = UCase(" :XDCC LIST") Then
            txtPM.Text = strNick & " :XDCC list" & vbNewLine
            txtPM.SelStart = Len(txtPM.Text) - 1
            thenick = getnicktosend(strData)
            Winsock1.SendData "PRIVMSG " & thenick & " : dCoolXDCC Listing " & strNick & vbCrLf
            For Y = 1 To List4.ListCount
                List4.Selected(Y - 1) = True
                Winsock1.SendData "PRIVMSG " & thenick & " :" & "#" & List4.Text & vbCrLf
            Next
            Winsock1.SendData "PRIVMSG " & thenick & " :/msg " & strNick & " XDCC Send #PACKNUMBER" & vbCrLf
        End If
    Next
    Else
        Winsock1.SendData "PRIVMSG " & thenick & " : XDCC is currently Disabled"
    End If
    
    For X = 1 To Len(strData)
        If UCase(Mid(strData, X, 8 + Len(strNick))) = "PRIVMSG " & UCase(strNick) Then
            For Y = 1 To Len(strData)
                If Mid(strData, X + Y + 8, 1) = " " Then
                        Load frmMessage
                        For z = 1 To Len(strData)
                            If Mid(strData, z, 1) = "!" And z < 20 Then
                                frmMessage.lblAction = Right(Left(strData, z - 1), Len(Left(strData, z - 1)) - 1) & " :"
                                GoTo gotname
                            End If
                        Next
gotname:
                        frmMessage.lblMessage = Right(strData, Len(strData) - X - Y - 9)
                        If allowalert = True Then
                        frmMessage.Show
                        frmMessage.Timer1.Enabled = True
                        End If
                        txtPM.Text = txtPM.Text & frmMessage.lblAction & " " & frmMessage.lblMessage & vbNewLine
                        txtPM.SelStart = Len(txtPM.Text) - 1
                        On Error Resume Next
                        Dim Foundnickinlist As Boolean
                        If frmPM.pmuserlist.ListCount <> 0 Then
                            For z = 1 To frmPM.pmuserlist.ListCount
                                frmPM.pmuserlist.Selected(z - 1) = True
                                If Left(frmMessage.lblAction, Len(frmMessage.lblAction) - 1) = frmPM.pmuserlist.Text Then
                                    GoTo Nickwasinlist
                                End If
                            Next
                            If Foundnickinlist = False Then
                                frmPM.pmuserlist.AddItem Left(frmMessage.lblAction, Len(frmMessage.lblAction) - 1)
                            End If
                        Else
                            frmPM.pmuserlist.AddItem Left(frmMessage.lblAction, Len(frmMessage.lblAction) - 1)
                        End If
Nickwasinlist:
                        frmPM.txtPM.Text = txtPM.Text
                        frmPM.txtPM.SelStart = Len(frmPM.txtPM.Text) - 1
                    GoTo done
                End If
            Next
        End If
    Next
done:
gotdone2:
    Status.SelStart = Len(Status) - 1
    status2.SelStart = Len(status2) - 1
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  'If an error occurs tell the user
  MsgBox "Error has occurred", vbCritical, "Connect Error"
  'Add the error in the Sattus Box
  Status.Text = Status.Text & Description & " - Error number: " & Number & vbNewLine
  'Enable button  1
  Command1.Enabled = True
  'Disable button 2
  Command2.Enabled = False
  wanttodisconnect = True
  Call Command2_Click
End Sub

Private Function getnicktosend(strData As String)
Dim X As Long
For X = 1 To Len(strData)
If Mid(strData, X, 1) = "!" Then
getnicktosend = Mid(strData, 2, X - 2)
End If
Next
End Function

Private Sub xdcctimer_Timer()
On Error Resume Next
    Static siTick As Integer
    If siTick >= adtime Then
    siTick = 0
    If Winsock1.State <> sckClosed Then
Dim z As Integer
Dim X As Integer
Dim adverchannel As String
Dim inchannel As Boolean
inchannel = False
For z = 1 To List5.ListCount
    List5.Selected(z - 1) = True
    For X = 1 To List2.ListCount
        List2.Selected(X - 1) = True
        If UCase(List2.Text) = UCase(List5.Text) Then
            inchannel = True
        End If
    Next
    If inchannel = True And Winsock1.State <> sckClosed Then
        Winsock1.SendData "PRIVMSG " & List5.Text & " :  - dCoolXDCC Online - " & strNick & vbCrLf
        Winsock1.SendData "PRIVMSG " & List5.Text & " : /msg " & strNick & " XDCC list" & vbCrLf
        If strNick = "dCool101d-XDCC" Then
            Winsock1.SendData "PRIVMSG " & List5.Text & " : Now Serving dCoolXDCC v" & thisversion.Text & vbCrLf
        End If
        
        If Form3.Text3.Text <> "" Then
            Winsock1.SendData "PRIVMSG " & List5.Text & " : " & Form3.Text3.Text & vbCrLf
        End If
        Status.Text = Status.Text & "*****Ad Placed in " & List5.Text & " *****" & vbNewLine
        Status.SelStart = Len(Status.Text)
    End If
    inchannel = False
Next
End If
    Else
        siTick = siTick + 1
    End If
End Sub

Private Function IRCStripColor(strText As String) As String
    Dim nc As Integer
    Dim i As Integer
    Dim col As Integer
    Dim slen As Integer
    Dim new_str As String
    Dim X As Integer
    nc = 0
    i = 0
    col = 0
    X = 1
    slen = Len(strText)
    Do While (slen > 0)
    If (((col And isDigit(Mid(strText, X, 1)) And (nc < 2)) Or _
    ((col And Mid(strText, X, 1) = ",") And _
    (isDigit(Mid(strText, (X + 1), 1))) And (nc < 3)))) Then
    nc = nc + 1
    If (Mid(strText, X, 1) = ",") Then
    nc = 0
    End If
    Else
    col = 0
    Select Case (Asc(Mid(strText, X, 1)))
    Case (3): ' color
    col = 1
    nc = 0
    GoTo Skip_Byte
    Case (7): ' Beep
    GoTo Skip_Byte
    Case (15): ' Reset
    GoTo Skip_Byte
    Case (22): ' Reverse
    GoTo Skip_Byte
    Case (2): ' Bold
    GoTo Skip_Byte
    Case (31): ' Underline
    GoTo Skip_Byte
    Case Else:
    new_str = new_str & Mid(strText, X, 1)
    i = i + 1
    End Select
    End If
Skip_Byte:
    X = X + 1
    slen = slen - 1
    Loop
    IRCStripColor = new_str
End Function

Private Function isDigit(digit As String) As Boolean
    ' based on the C/C++ ctype.h function.
    Dim X As Integer
    Dim c As String
    c = Left(digit, 1)
    If (c = "") Then
    isDigit = False
    Exit Function
    End If
    X = Asc(c)
    If ((X >= 48) And (X <= 57)) Then
    isDigit = True
    Else
    isDigit = False
    End If
End Function

