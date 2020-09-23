VERSION 5.00
Begin VB.Form frmtray 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TrayCode"
   ClientHeight    =   1155
   ClientLeft      =   1335
   ClientTop       =   2160
   ClientWidth     =   1230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1155
   ScaleWidth      =   1230
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   120
      Picture         =   "tray.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Menu mPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore dCoolXDCC"
      End
      Begin VB.Menu mnuNull 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close dCoolXDCC"
      End
   End
End
Attribute VB_Name = "frmtray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IconIndex As Integer, IntDown As Integer
Sub MyPopupmenu(X As Integer, Y As Integer)
    hMenu = GetMenu(Me.hwnd)
    hSubMenu = GetSubMenu(hMenu, 0)
    hTray = FindWindow("Shell_TrayWnd", 0&)
    If hTray > 0 Then
        Dim TrayRect As RECT
        Result = GetWindowRect(hTray, TrayRect)
        lResult = TrackPopupMenu(hSubMenu, TPM_RIGHTALIGN, TrayRect.Right, TrayRect.Bottom, 0&, Form2.hwnd, ByVal 0&)
    End If
End Sub

Private Sub Form_Load()
    Dim iData As NOTIFYICONDATA
    
    iData.cbSize = Len(iData)
    iData.hwnd = Picture1.hwnd
    iData.uID = 9999
    iData.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
    iData.uCallbackMessage = WM_LBUTTONDOWN
    iData.hIcon = Picture1.Picture
    iData.szTip = "dCoolXDCC" & Chr$(0)
    
    Result = Shell_NotifyIcon(NIM_ADD, iData)
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim iData As NOTIFYICONDATA
    
    iData.cbSize = Len(iData)
    iData.hwnd = Picture1.hwnd
    iData.uID = 9999
    
    Result = Shell_NotifyIcon(NIM_DELETE, iData)

    End
    
End Sub

Private Sub Image2_Click()

End Sub

Private Sub mnuClose_Click()
    Form1.SetFocus
    blnexit = True
    Call Form1.mnuExit_Click
End Sub

Private Sub mnuRestore_Click()
    Form1.Show
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case X
        Case TRAY_MSG_MOUSEMOVE
        
        Case TRAY_MSG_LEFTBTN_DOWN
        
        Case TRAY_MSG_LEFTBTN_UP
            Form1.Show
        Case TRAY_MSG_LEFTBTN_DBLCLICK
            
        Case TRAY_MSG_RIGHTBTN_DOWN
        
        Case TRAY_MSG_RIGHTBTN_UP
            Form1.Show
        Case TRAY_MSG_RIGHTBTN_DBLCLICK
            MyPopupmenu (X), (Y)
    End Select

End Sub
