VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dCoolXDCC - Help"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7995
   Icon            =   "frmhelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPM 
      BackColor       =   &H80000004&
      Height          =   735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Text            =   "frmhelp.frx":628A
      Top             =   4920
      Width           =   5895
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
         Height          =   1035
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.ListBox List2 
         Height          =   645
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
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
      Text            =   "frmhelp.frx":62DC
      Top             =   0
      Width           =   5895
   End
   Begin VB.TextBox Status 
      BackColor       =   &H80000004&
      Height          =   4245
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "frmhelp.frx":62F2
      Top             =   600
      Width           =   5895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Enabled         =   0   'False
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
      Enabled         =   0   'False
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
      Enabled         =   0   'False
      Begin VB.Menu mnuBasicSettings 
         Caption         =   "&Basic Settings"
      End
      Begin VB.Menu mnuXDCCSettings 
         Caption         =   "&XDCC Settings"
      End
   End
   Begin VB.Menu mnuUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
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
      Enabled         =   0   'False
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
      Enabled         =   0   'False
      Begin VB.Menu mnuWebsite 
         Caption         =   "&Visit Website - dCool101d.tk"
      End
      Begin VB.Menu mnuAboutdCoolXDCC 
         Caption         =   "About d&CoolXDCC"
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Click to connect XDCC to IRC server"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Click to disconnect XDCC from IRC server"
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Click to add the channel entered above"
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Click to re-configure your basic connection settings"
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Click to configure your XDCC settings"
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Click to add the text entered above to your Ad Channel list"
End Sub

Private Sub Form_Load()
    frmHelp2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmHelp2
    Form1.Show
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Double-Click to join a selected channel"
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Double-Click to part from the seleted channel"
End Sub

Private Sub List3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "This is a list of user who are recieving files from you XDCC"
End Sub

Private Sub List4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "This is a list of files that your XDCC has listed"
End Sub

Private Sub List5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "This is a list of channels that you XDCC will display an add in a specified amount of time"
End Sub

Private Sub Status_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "This display all the incoming messages after they have been filtered"
End Sub

Private Sub status2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "This displays all the raw incoming messages from the IRC server"
End Sub

Private Sub Text10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Enter a channel to add to your Ad Channel list"
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Enter text you would like to send to the selected channel and hit enter (from the Joined channel list)"
End Sub

Private Sub Text4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "Enter a channel name that you would like to add to you channel list"
End Sub

Private Sub txtPM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmHelp2.lblStatus.Caption = "This display all your incoming Private messages. To respond, Go to: File>Send Private Message"
End Sub
