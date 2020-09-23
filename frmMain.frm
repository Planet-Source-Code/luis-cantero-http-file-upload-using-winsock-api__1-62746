VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTTPUpload"
   ClientHeight    =   6750
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   6495
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox DataArrival 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtServerResponse 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3000
      Width           =   6255
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   1725
      TabIndex        =   8
      Text            =   "http://localhost/FastASPUpload/upload.asp"
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtFile2 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtFile1 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Upload Script URL:"
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   210
      Width           =   1380
   End
   Begin VB.Label lblAuctionImage2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "File 2:"
      Height          =   195
      Left            =   825
      TabIndex        =   5
      Top             =   1290
      Width           =   420
   End
   Begin VB.Label lblAuctionImage 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "File 1:"
      Height          =   195
      Left            =   825
      TabIndex        =   3
      Top             =   810
      Width           =   420
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Ready"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   2640
      Width           =   465
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Winsock
Dim intSock As Integer
Dim strReceivedData As String

Private Sub cmdBrowse_Click(Index As Integer)

    Select Case Index
      Case 1
        txtFile1 = OpenDialog("JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|CompuServe GIF (*.gif)|*.gif", "Select a file...", "*.jpg", App.Path, Me.hwnd)

      Case 2
        txtFile2 = OpenDialog("AVI (*.avi)|*.avi|MPEG (*.mpg;*.mpeg)|*.mpg;*.mpeg|WMV (*.wmv)|*.wmv|QuickTime (*.mov)|*.mov|Flash (*.swf)|*.swf", "Select a file...", "*.avi", App.Path, Me.hwnd)

    End Select

End Sub

Private Sub cmdUpload_Click()

  Dim objURL As CURL

    DataArrival = ""

    Set objURL = New CURL

    With objURL
        Call .ParseURL(txtURL, GetProxy())
    End With

    strReceivedData = ""

    Call UploadFiles(objURL)

End Sub

Private Sub DataArrival_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim lngBytes As Long
  Dim strServerResponse As String
  Dim MsgBuffer As String * 8192

    On Error Resume Next

        'A Socket is open
        If intSock > 0 Then
            'Receive up to 8192 chars
            lngBytes = recv(intSock, MsgBuffer, 8192, 0)

            If lngBytes > 0 Then
                strServerResponse = Mid$(MsgBuffer, 1, lngBytes)

                'Beginning of HTTP Response
                If Left$(strServerResponse, 7) = "HTTP/1." Then
                    'Skip server information
                    strServerResponse = Mid$(strServerResponse, InStr(1, strServerResponse, vbCrLf & vbCrLf) + 4)
                End If

                txtServerResponse = txtServerResponse & strServerResponse
                txtServerResponse.SelStart = Len(txtServerResponse)

                strReceivedData = strReceivedData & strServerResponse

                '0 Bytes received, close sock to indicate end of receive
              ElseIf WSAGetLastError() <> WSAEWOULDBLOCK Then
                closesocket (intSock)
                Call EndWinsock 'Very important!
                intSock = 0
            End If
        End If

        Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If intSock > 0 Then closesocket (intSock)
    Call EndWinsock
    End

End Sub

Private Function GetFileExtension(strFileName As String) As String

  'Error check

    If Len(strFileName) < 3 Or InStr(1, strFileName, ".") = 0 Then
        GetFileExtension = ""
        Exit Function
    End If

    'Return
    GetFileExtension = Mid$(strFileName, InStrRev(strFileName, ".") + 1)

End Function

Private Function GetMimeType(strFileName As String) As String

  Dim strExtension As String

    strExtension = LCase$(GetFileExtension(strFileName))

    'Error check
    If strExtension = "" Then
        GetMimeType = "text/plain"
        Exit Function
    End If

    Select Case strExtension
      Case "bmp"
        GetMimeType = "image/bmp"

      Case "gif"
        GetMimeType = "image/gif"

      Case "jpg", "jpeg"
        GetMimeType = "image/jpeg"

      Case "swf"
        GetMimeType = "application/x-shockwave-flash"

      Case "mpg", "mpeg"
        GetMimeType = "video/mpeg"

      Case "wmv"
        GetMimeType = "video/x-ms-wmv"

      Case "avi"
        GetMimeType = "video/avi"

      Case Else
        GetMimeType = "text/plain"

    End Select

End Function

Function GetProxy() As String

  Dim strProxy As String
  Dim blnProxyEnabled As Boolean
  Dim intBeg As Integer
  Dim intEnd As Integer

    strProxy = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer", REG_SZ)
    blnProxyEnabled = CBool(QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable", REG_DWORD))

    intBeg = InStr(1, strProxy, "http=", vbTextCompare)

    If intBeg > 0 And blnProxyEnabled Then
        intEnd = InStr(intBeg + 1, strProxy, ";")
        If intEnd = 0 Then intEnd = Len(strProxy)
        strProxy = Mid$(strProxy, intBeg + 5, intEnd - intBeg - 5)
    End If

    'Return
    GetProxy = strProxy

End Function

Private Sub UploadFiles(objURL As CURL)

  Dim lngStart As Long
  Dim strQuery As String
  Dim strFileContents1 As String
  Dim strFileContents2 As String
  Dim objHTTPRequest As CHTTPRequest

  Const intSecondsToWait = 10 'Seconds to wait = 10

    lblStatus.Caption = "Connecting to " & objURL.Host
    DoEvents

    intSock = ConnectSock(objURL.Host, objURL.Port, DataArrival.hwnd)

    If intSock = SOCKET_ERROR Then
        lblStatus.Caption = "Could not connect to " & objURL.Host
        Exit Sub
    End If

    Set objHTTPRequest = New CHTTPRequest

    'File 1
    If txtFile1 <> "" Then
        If Dir(txtFile1) <> "" Then  'File exists
            strFileContents1 = GetFileQuick(txtFile1)
        End If
    End If

    'File 2
    If txtFile2 <> "" Then
        If Dir(txtFile2) <> "" Then  'File exists
            strFileContents2 = GetFileQuick(txtFile2)
        End If
    End If

    With objHTTPRequest
        .Host = objURL.Host
        .Proxy = objURL.Proxy
        .Path = objURL.Path
        .UserAgent = App.Title

        .MimeBoundary = "LcEnTeRpRiSeS"

        'Form fields
        Call .AddFormData("Field1", "test field")
        Call .AddFormData("Field2", "test field 2")

        'Files
        Call .AddFile("File1", txtFile1, strFileContents1, GetMimeType(txtFile1))
        Call .AddFile("File2", txtFile2, strFileContents2, GetMimeType(txtFile2))

        'MsgBox .GetGETQuery
        strQuery = .GetPOSTQuery
    End With

    lblStatus.Caption = "Sending request..."
    DoEvents

    'Send request
    Call SendData(intSock, strQuery)

    'Wait for page to be downloaded
    lngStart = timeGetTime
    While intSecondsToWait - Int((timeGetTime - lngStart) / 1000) > 0 And intSock > 0
        lblStatus.Caption = "Waiting for response from " & objURL.Host & "... " & intSecondsToWait - Int((timeGetTime - lngStart) / 1000)
        DoEvents

        'You can put a routine that will check if a boolean variable is True here
        'This could indicate that the request has been canceled
        'If CancelFlag = True Then
        '   lblStatus.Caption = "Cancelled request"
        '   Exit Sub
        'End If
    Wend

    lblStatus.Caption = "Ready"

End Sub

':) Ulli's VB Code Formatter V2.13.6 (02.10.2005 19:34:05) 5 + 238 = 243 Lines
