Attribute VB_Name = "Module1"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long

Public lastmessage As String
Public strNick As String
Public strEmail As String
Public strName As String
Public strPassword As String
Public strQuit As String
Public strChannel As String
Public Filetosend(1 To 5) As String
Public Filename(1 To 5) As String
Public Packnumber(1 To 5) As Long
Public Nicktosend(1 To 5) As String
Public lasttag As Long
Public XDCCEnabled As Boolean
Public strFullPath As String
Public theIP As String
Public lblFileSize(1 To 5) As Long
Public intmaxsends As Long
Public SendFileForm(1 To 20) As New frmSendFile
Public blnexit As Boolean
Public maxbandwithperuser As Long
Public allowalert As Boolean



Public Sub DestroyFile(sFileName As String)
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, offset As Long
    'Create two buffers with a specified 'wi
    '     pe-out' characters
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    'Overwrite the file contents with the wi
    '     pe-out characters
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
    Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1


    For iLoop = 1 To Blocks
        offset = Seek(hFileHandle)
        Put hFileHandle, , Block1
        Put hFileHandle, offset, Block2
    Next iLoop
    Close hFileHandle
    'Now you can delete the file, which cont
    '     ains no sensitive data
    Kill sFileName
End Sub


Function IrcGetLongIP(ByVal AscIp$) As String
    'this function converts an ascii ip string into a long ip in network byte order
    'and stick it in a string suitable for use in a DCC command.
    On Error GoTo IrcGetLongIpError:
    Dim inn&
    inn = inet_addr(AscIp)
    inn = htonl(inn)
    If inn < 0 Then
        IrcGetLongIP = CVar(inn + 4294967296#)
        Exit Function
    Else
        IrcGetLongIP = CVar(inn)
        Exit Function
    End If
    Exit Function
IrcGetLongIpError:
    IrcGetLongIP = "0"
    Exit Function
    Resume
End Function
