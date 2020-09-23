VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "dCoolXDCC Update"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlg 
      Left            =   0
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Latest Version"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5415
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Version"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Version "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check for Update"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   3360
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock wscHttp 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReadURL 
      Caption         =   "Update dCoolXDCC"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   3360
      Width           =   2055
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_strRemoteHost As String    'the web server to connect to
Private m_strFilePath As String      'relative path to the file to retrieve
Private m_strHttpResponse As String  'the server response
Private m_bResponseReceived As Boolean
Dim FileURL As String
Dim Filename As String
Dim gotversion As Boolean
Dim latestversion As Currency
Dim versionname As String
Dim fileinfo As String

Public Sub GetFile()
    '
    Dim strURL As String    'temporary buffer
    '
    '
    'check the textbox
    If Len(FileURL) = 0 Then
        MsgBox "Please, enter the URL to retrieve.", vbInformation
        Exit Sub
    End If
    '
    'if the user has entered "http://", remove this substring
    '
    If Left(FileURL, 7) = "http://" Then
        strURL = Mid(FileURL, 8)
    Else
        strURL = FileURL
    End If
    '
    'get remote host name
    '
    m_strRemoteHost = Left$(strURL, InStr(1, strURL, "/") - 1)
    '
    'get relative path to the file to retrieve
    '
    m_strFilePath = Mid$(strURL, InStr(1, strURL, "/"))
    '
    'clear the RichTextBox
    '
    '
    'clear the buffer
    '
    m_strHttpResponse = ""
    '
    'turn off the m_bResponseReceived flag
    '
    m_bResponseReceived = False
    '
    'establish the connection
    '
    With wscHttp
        .Close
        .LocalPort = 0
        .Connect m_strRemoteHost, 80
    End With
    '
EXIT_LABEL:
    Exit Sub
End Sub

Private Sub cmdReadURL_Click()
    On Error GoTo err:
    Dim id As Long
    FileURL = "http://home.comcast.net/~khameneh/dCoolXDCC.exe"
    Label1.Caption = "Retrieving Update..."
    Call GetFile
    Exit Sub
err:
MsgBox "Error Updating"
Form1.Enabled = True
Form1.Show
Unload Me
End Sub

Private Sub Command1_Click()
    FileURL = "http://mywebpages.comcast.net/khameneh/dCoolXDCC.txt"
    Filename = "C:\dCoolXDCC.txt"
    Call GetFile
    Command1.Enabled = False
    Label1.Caption = "Checking for updates..."
End Sub

Private Sub Form_Load()
Label2.Caption = Label2.Caption & Form1.thisversion.Text
Text1.Text = Form1.thisversioninfo.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
Form1.Show
End Sub

Private Sub thisversion_Change()

End Sub

Private Sub wscHttp_Close()
    '
    Dim strHttpResponseHeader As String
    '
    'to cut of the header info, we must find 
    'a blank line (vbCrLf & vbCrLf)
    'that separates the message body from the header
    '
    If Not m_bResponseReceived Then
        strHttpResponseHeader = Left$(m_strHttpResponse, _
                                InStr(1, m_strHttpResponse, _
                                vbCrLf & vbCrLf) - 1)
        Debug.Print strHttpResponseHeader
        m_strHttpResponse = Mid(m_strHttpResponse, _
                            InStr(1, m_strHttpResponse, _
                            vbCrLf & vbCrLf) + 4)
        '
        'pass the document data to the RichTextBox control
        '
        Open Filename For Binary As #1
        Put #1, , m_strHttpResponse
        Close #1
        '
        'turn on the m_bResponseReceived flag
        '
        m_bResponseReceived = True
        '
    End If
    
    If gotversion = False Then
    gotversion = True
    Open Filename For Input As #1
    Input #1, latestversion
    Input #1, versionname
    Input #1, fileinfo
    Close #1
    If latestversion > Val(Form1.thisversion.Text) Then
        Label1.Caption = versionname
        Text2.Text = fileinfo
        cmdReadURL.Enabled = True
    Else
        Label1.Caption = "No new updates found!"
    End If
    
    
    End If
    If Filename = "dCoolXDCC_Update_" & latestversion & ".exe" Then
        id = Shell(Filename, 1)
        blnexit = True
        End
        Exit Sub
    End If
    Filename = "dCoolXDCC_Update_" & latestversion & ".exe"
End Sub

Private Sub wscHttp_Connect()
    '
    Dim strHttpRequest As String
    '
    'create the HTTP Request
    '
    'build request line that contains the HTTP method, 
    'path to the file to retrieve,
    'and HTTP version info. Each line of the request 
    'must be completed by the vbCrLf
    strHttpRequest = "GET " & m_strFilePath & " HTTP/1.1" & vbCrLf
    '
    'add HTTP headers to the request
    '
    'add required header - "Host", that contains the remote host name
    '
    strHttpRequest = strHttpRequest & "Host: " & m_strRemoteHost & vbCrLf
    '
    'add the "Connection" header to force the server to close the connection
    '
    strHttpRequest = strHttpRequest & "Connection: close" & vbCrLf
    '
    'add optional header "Accept"
    '
    strHttpRequest = strHttpRequest & "Accept: */*" & vbCrLf
    '
    'add other optional headers
    '
    'strHttpRequest = strHttpRequest & <Header Name> & _
                      <Header Value> & vbCrLf
    '. . .
    '
    'add a blank line that indicates the end of the request
    strHttpRequest = strHttpRequest & vbCrLf
    '
    'send the request
    wscHttp.SendData strHttpRequest
    '
    Debug.Print strHttpRequest
    '
End Sub

Private Sub wscHttp_DataArrival(ByVal bytesTotal As Long)
    '
    On Error Resume Next
    '
    Dim strData As String
    '
    'get arrived data from winsock buffer
    '
    wscHttp.GetData strData
    '
    'store the data in the m_strHttpResponse variable
    m_strHttpResponse = m_strHttpResponse & strData
    '
End Sub
