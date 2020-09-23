VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "dCoolXDCC Alert"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3030
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3500
      Left            =   2640
      Top             =   480
   End
   Begin VB.Label lblMessage 
      Caption         =   "message"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblAction 
      Alignment       =   2  'Center
      Caption         =   "Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1

Private Sub Form_Click()
    Load frmPM
    frmPM.txtUser = Left(lblAction.Caption, Len(lblAction.Caption) - 1)
    frmPM.Show
End Sub

Private Sub Form_Resize()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Load()
Me.Top = Screen.Height - Me.Height - 400
Me.Left = Screen.Width - Me.Width
End Sub

Private Sub lblAction_Click()
    Load frmPM
    frmPM.txtUser = Left(lblAction.Caption, Len(lblAction.Caption) - 1)
    frmPM.Show
End Sub

Private Sub lblMessage_Click()
    Load frmPM
    frmPM.txtUser = Left(lblAction.Caption, Len(lblAction.Caption) - 1)
    frmPM.Show
End Sub

Private Sub Timer1_Timer()
Me.Hide
Timer1.Enabled = False
End Sub
