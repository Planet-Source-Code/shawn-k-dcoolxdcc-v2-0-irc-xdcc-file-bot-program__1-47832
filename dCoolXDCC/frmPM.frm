VERSION 5.00
Begin VB.Form frmPM 
   Caption         =   "Send Private Message"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   Icon            =   "frmPM.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3795
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox pmuserlist 
      Height          =   1815
      Left            =   5040
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtPM 
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   4815
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "MESSAGE:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TO:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   270
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
txtPM.Height = frmPM.Height - 1890 - 180 + 55
pmuserlist.Height = frmPM.Height - 980 + 40
txtPM.Width = frmPM.Width - 200 - 2200
pmuserlist.Left = frmPM.Width - 200 - 50 - pmuserlist.Width
txtUser.Width = frmPM.Width - 200 - 2200
txtMessage.Width = frmPM.Width - 200 - 2200
End Sub

Private Sub Form_Resize()
On Error Resume Next
txtPM.Height = frmPM.Height - 1890 - 180 + 55
pmuserlist.Height = frmPM.Height - 980 + 40
txtPM.Width = frmPM.Width - 200 - 2200
pmuserlist.Left = frmPM.Width - 200 - 50 - pmuserlist.Width
txtUser.Width = frmPM.Width - 200 - 2200
txtMessage.Width = frmPM.Width - 200 - 2200
End Sub

Private Sub Form_Unload(Cancel As Integer)
If blnexit = False Then
Cancel = 1
Me.Hide
End If
End Sub

Private Sub pmuserlist_Click()
txtUser.Text = pmuserlist.Text
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
If Form1.Winsock1.State <> sckClosed And txtMessage.Text <> "" Then
    If KeyAscii = 13 Then
        Form1.Winsock1.SendData "PRIVMSG " & txtUser.Text & " :" & txtMessage.Text & vbCrLf
        Form1.txtPM.Text = Form1.txtPM.Text & "-TO- " & txtUser.Text & " :" & txtMessage.Text & vbNewLine
        txtMessage.Text = ""
        txtPM.Text = Form1.txtPM.Text
        txtPM.SelStart = Len(txtPM.Text) - 1
    End If
End If
End Sub
