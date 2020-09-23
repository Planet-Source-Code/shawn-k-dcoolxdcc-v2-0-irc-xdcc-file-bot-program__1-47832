VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmSendFile 
   Caption         =   "DCC SEND"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Tag             =   "0"
   Begin VB.Timer filesendtimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   960
   End
   Begin VB.TextBox Filename 
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Packnumber 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Filetosend 
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox lblFileSize 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox TempFileName 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Nicktosend 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox ByteSent 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock tcpSend 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "irc.winebta.org"
      RemotePort      =   1560
      LocalPort       =   2000
   End
End
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFullPath As String
Dim i, fLength, ret                               '// Declare Variables
Dim Buffer As String                              '// Declare Buffer
Dim bSize As Long
Public inuse As Boolean
Dim bytessent As Long
Dim continue As Boolean
Public theremoteplace As String

Public Function sendfile() As Boolean
    continue = True
    'i = FreeFile                                      '// Set I As FreeFile
    i = Me.Tag + 5
    'Open strFullPath For Binary Access Read As #i
    Close #i
    Open Filetosend For Binary Access Read As i
    filesendtimer.Enabled = True
End Function
Public Sub sendfiletonick()
    i = Me.Tag + 5
    On Error Resume Next
    bytessent = 1
    Open Filetosend For Binary Access Read As (Me.Tag + 5)
    lblFileSize = LOF(Me.Tag + 5)
    Close Me.Tag + 10
    tcpSend(Me.Tag).Close
    tcpSend(Me.Tag).Listen
    Me.Tag = lasttag
    Dim LIP As String
    Dim TempFileName As String
    TempFileName = Replace(Filetosend, " ", "_")
    LIP = IrcGetLongIP(Form2.Text3.Text)
    Form1.Winsock1.SendData "NOTICE " & Nicktosend & " :DCC SEND " & Filetosend & "(" & Form2.Text3.Text & ")" & vbCrLf
    Form1.Winsock1.SendData "PRIVMSG " & Nicktosend & " :DCC SEND " & TempFileName & " " & LIP & " " & tcpSend(Me.Tag).LocalPort & " " & " " & Val(lblFileSize) & "" & vbCrLf
End Sub

Private Sub filesendtimer_Timer()
On Error Resume Next
If continue = True And tcpSend(Me.Tag).State <> sckConnecting And tcpSend(Me.Tag).State <> sckResolvingHost Then
Dim X As Integer
    bSize = 100 * maxbandwithperuser
    fLength = LOF(i)
        If EOF(i) = False Then                          '// Begin A Loop Until EOF
                                                       
            If fLength - Loc(i) <= bSize Then     '// If The Buffer Is Larger Than
                bSize = fLength - Loc(i)          '// The Rest Of the File. Make The
            End If                                '// New Buffer Size The Rest Of The
                                                  '// File
            If bSize = 0 Then
            filesendtimer.Enabled = False
            Close i                                       '// Close File
                inuse = False
                For X = 1 To Form1.List3.ListCount
                    Form1.List3.Selected(X - 1) = True
                    If Form1.List3.Text = Nicktosend & "-" & Packnumber Then
                        Form1.List3.RemoveItem (X - 1)
                    End If
                Next
                filesendtimer.Enabled = False
                Exit Sub
            End If '// If Buffer Size Is 0 Send Done
        
            ByteSent = Val(ByteSent) + bSize           '// Adds The Buffer To Bytes Sent
            Buffer = Space$(bSize)                '// Get The Buffer From The BlockSize
            Get i, bytessent, Buffer                       '// Take Block From File
            bytessent = bytessent + bSize
            tcpSend(Me.Tag).SendData Buffer                   '// Send Block
        End If                                      '// Loop
    If EOF(i) = True Then
        Close i                                       '// Close File
                inuse = False
                For X = 1 To Form1.List3.ListCount
                    Form1.List3.Selected(X - 1) = True
                    If Form1.List3.Text = Nicktosend & "-" & Packnumber Then
                        Form1.List3.RemoveItem (X - 1)
                    End If
                Next
                filesendtimer.Enabled = False
            Exit Sub
    End If
    continue = False
End If
End Sub

Private Sub Form_Load()
    tcpSend(Me.Tag).LocalPort = 2000 + Me.Tag
End Sub

Private Sub tcpSend_Close(Index As Integer)
    Close i                                       '// Close File
                inuse = False
                For X = 1 To Form1.List3.ListCount
                    Form1.List3.Selected(X - 1) = True
                    If Form1.List3.Text = Nicktosend & "-" & Packnumber Then
                        Form1.List3.RemoveItem (X - 1)
                    End If
                Next
                Exit Sub
End Sub

Private Sub tcpSend_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If tcpSend(Me.Tag).State <> sckClosed Then tcpSend(Me.Tag).Close
    tcpSend(Me.Tag).Accept requestID
    sendfile
End Sub

Private Sub tcpSend_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    tcpSend(Me.Tag).GetData strData
    continue = True
End Sub

Private Sub tcpSend_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description
End Sub
