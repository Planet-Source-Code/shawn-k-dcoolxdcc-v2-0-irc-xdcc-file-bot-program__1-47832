VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connection Settings"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7455
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.CheckBox Check1 
         Caption         =   "Enable alert system (PM alerts, XDCC alerts)"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   3120
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   6600
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton Command6 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2760
         TabIndex        =   12
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox txtemail2 
         Height          =   285
         Left            =   5160
         MaxLength       =   100
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtemail 
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   10
         Top             =   2160
         Width           =   3255
      End
      Begin VB.CheckBox chkreg 
         Caption         =   "Register on Connect (First time users ONLY)"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtpassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   100
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1800
         Width           =   5535
      End
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1440
         Width           =   5535
      End
      Begin VB.TextBox txtnick 
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   5
         Text            =   "irc.winbeta.org"
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "6667"
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   3
         Top             =   2520
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get IP Address"
         Height          =   255
         Left            =   5760
         TabIndex        =   2
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtspace 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   $"Form2.frx":0000
         Top             =   3480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4920
         TabIndex        =   20
         Top             =   2160
         Width           =   210
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Adress:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Server address:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Port to connect to:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "IP address:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1

Private Sub Command1_Click()
    Text3.Text = "Please wait..."
    Text3.Text = Replace(Inet1.OpenURL("http://www.pchelplive.com/ip.php"), txtspace.Text, "")
End Sub

Private Sub Form_Resize()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Load()
    On Error GoTo err:
    allowalert = True
    Open "cnfg.dat" For Input As #2
    Input #2, strNick, strName, strEmail, email2, strPassword, theserver, theport, theIP, allowalert
    Close #2
    If allowalert = True Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    txtnick.Text = strNick
    txtname.Text = strName
    txtemail.Text = strEmail
    txtpassword.Text = strPassword
    txtemail2.Text = email2
    Text1.Text = theserver
    Text2.Text = theport
    If theIP = "" Then
    Text3.Text = Form1.Winsock1.LocalIP
    Else
    Text3.Text = theIP
    End If
err:
    Close #2
End Sub

Private Sub Command6_Click()
    strNick = txtnick.Text
    strName = txtname.Text
    strEmail = txtemail.Text
    strPassword = txtpassword.Text
    email2 = txtemail2.Text
    theserver = Text1.Text
    theport = Text2.Text
    theIP = Text3.Text
    If Check1.Value = 1 Then
        allowalert = True
    Else
        allowalert = False
    End If
    Open "cnfg.dat" For Output As #2
    Write #2, strNick, strName, strEmail, email2, strPassword, theserver, theport, theIP, allowalert
    Close #2
    Form1.Enabled = True
    Me.Hide
    Form1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Form1.Enabled = True
    Form1.Show
If blnexit = False Then
    Cancel = 1
End If
End Sub

