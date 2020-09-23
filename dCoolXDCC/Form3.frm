VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "XDCC Settings"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMaxSends 
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "20"
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox bandwith 
      Height          =   285
      Left            =   1920
      MaxLength       =   100
      TabIndex        =   11
      Text            =   "200"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "Brought to you by Winbeta, dCoolXDCC, and...ME!"
      Top             =   2760
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "5"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Place ad in specified channels every"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Max Sends:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "kb/s"
      Height          =   195
      Left            =   2400
      TabIndex        =   12
      Top             =   1920
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Max Bandwith Per User:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1710
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Credit Line:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "min."
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XDCC Folders/ Individual Files:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2190
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1

Dim intNum As Long

Private Sub bandwith_Change()
If Val(bandwith.Text) > 200 Then
    bandwith.Text = "200"
End If
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Text2.Enabled = True
        Text3.Enabled = True
        Label3.Enabled = True
    Else
        Text2.Enabled = False
        Text3.Enabled = False
        Label3.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    Open "tmpfilecnfg.dat" For Output As #3
    Dim foldertoadd As String
    foldertoadd = Text1.Text
    If Mid(foldertoadd, Len(foldertoadd) - 3, 1) = "." Or Mid(foldertoadd, Len(foldertoadd) - 2, 1) = "." Or Mid(foldertoadd, Len(foldertoadd) - 4, 1) = "." Then
        List1.AddItem foldertoadd
        intNum = intNum + 1
        filenameparts = Split(foldertoadd, "/")
        Write #3, intNum, Left(filenameparts(UBound(filenameparts)), Len(filenameparts(UBound(filenameparts))) - 4), foldertoadd
        Form1.List4.AddItem intNum & " - " & FileName2
    Else
        If Right(foldertoadd, 1) <> "\" Then
            foldertoadd = foldertoadd & "\"
        End If
        List1.AddItem foldertoadd
        Call addXDCCfolder(foldertoadd)
    End If
    Text1.Text = ""
    Close #3
End Sub

Private Sub Command2_Click()
    Form1.adtime = 5
    Dim displayad As Integer
    Dim adtime As Integer
    Dim creditline As String
    Dim X As Long
    Dim xdcclist As String
    Open "xdccfolders.dat" For Output As #4
    displayad = Check1.Value
    adtime = Val(Text2.Text)
    Form1.adtime = adtime
    creditline = Text3.Text
    If Val(bandwith.Text) < 1 Then bandwith.Text = "1"
    If Val(txtMaxSends.Text) < 1 Then txtMaxSends.Text = "1"
    If Val(txtMaxSends.Text) > 20 Then txtMaxSends.Text = "20"
    maxbandwithperuser = Val(bandwith.Text)
    intmaxsends = Val(txtMaxSends.Text)
    Write #4, displayad, adtime, creditline, maxbandwithperuser, intmaxsends
    For X = 1 To List1.ListCount
    List1.Selected(X - 1) = True
    xdcclist = List1.Text
    Write #4, xdcclist
    Next
    Close #4
err:
    maxbandwithperuser = Val(bandwith.Text)
    Me.Hide
    Form1.Enabled = True
    Form1.Show
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim displayad As Integer
    Dim adtime As Integer
    Dim creditline As String
    On Error GoTo err:
    Dim xdcclist As String
    Open "xdccfolders.dat" For Input As #1
    Input #1, displayad, adtime, creditline, maxbandwithperuser, intmaxsends
    Form1.adtime = adtime
    txtMaxSends.Text = intmaxsends
    For X = 1 To 20
        SendFileForm(X).Tag = X - 1
    Next
    Check1.Value = displayad
    Text2.Text = adtime
    Open "tmpfilecnfg.dat" For Output As #3
    Do While Not EOF(1)
        Input #1, xdcclist
        List1.AddItem xdcclist
        If Mid(xdcclist, Len(xdcclist) - 4, 1) = "." Then
            intNum = intNum + 1
            filenameparts = Split(xdcclist, "/")
            Write #3, intNum, Left(filenameparts(UBound(filenameparts)), Len(filenameparts(UBound(filenameparts))) - 4), xdcclist
            Form1.List4.AddItem intNum & " - " & FileName2
        Else
            addXDCCfolder (xdcclist)
        End If
    Loop
    Close #3
    Close #1
    bandwith.Text = maxbandwithperuser
    Text3.Text = creditline
    If Check1.Value = 1 Then
        Form1.xdcctimer.Enabled = True
    Else
        Form1.xdcctimer.Enabled = False
    End If
err:
    On Error GoTo err2
    Close #1
err2:
End Sub

Private Sub Form_Resize()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function addXDCCfolder(directory As String)
    Dim Filename As String
    Dim FileName2 As String
    Dim tmpStrg As String

    tmpStrg = Dir$(directory & "*.*")
    
    If tmpStrg <> "" Then
    
        FileName2 = Left$(tmpStrg, Len(tmpStrg) - 4)
        FileName2 = Replace(FileName2, "_", " ")
        Filename = tmpStrg
        tmpStrg = Dir$
        
        intNum = intNum + 1
        Write #3, intNum, FileName2, directory & Filename
        Form1.List4.AddItem intNum & " - " & FileName2
        
        
        While Len(tmpStrg) > 0
            FileName2 = Left$(tmpStrg, Len(tmpStrg) - 4)
            Filename = tmpStrg
            tmpStrg = Dir$
            intNum = intNum + 1
            Write #3, intNum, FileName2, directory & Filename
            Form1.List4.AddItem intNum & " - " & FileName2
            
        Wend
    Else
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Form1.Enabled = True
    Form1.Show
If blnexit = False Then
    Cancel = 1
End If
End Sub

Private Sub List1_DblClick()
    List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub Text2_Change()
    Text2.Text = Val(Text2.Text)
End Sub

