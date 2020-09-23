VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Personal Proxy Server"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabProxy 
      Height          =   4995
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8811
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Log"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraLog"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Header"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraRequest"
      Tab(1).Control(1)=   "fraResponse"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Stat"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraStat"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraStat 
         Caption         =   "Connection Request Statistic"
         Height          =   4485
         Left            =   90
         TabIndex        =   10
         Top             =   390
         Width           =   6435
         Begin MSFlexGridLib.MSFlexGrid flxStatistic 
            Height          =   4095
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   7223
            _Version        =   393216
            AllowBigSelection=   0   'False
            SelectionMode   =   1
         End
      End
      Begin VB.Frame fraResponse 
         Caption         =   "Response Header"
         Height          =   1905
         Left            =   -74910
         TabIndex        =   7
         Top             =   2340
         Width           =   6465
         Begin VB.TextBox txtResponse 
            Height          =   1515
            Left            =   150
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   270
            Width           =   6195
         End
      End
      Begin VB.Frame fraRequest 
         Caption         =   "Request Header"
         Height          =   1905
         Left            =   -74910
         TabIndex        =   6
         Top             =   390
         Width           =   6465
         Begin VB.TextBox txtRequest 
            Height          =   1515
            Left            =   150
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   6195
         End
      End
      Begin VB.Frame fraLog 
         Caption         =   "Proxy Logs"
         Height          =   4485
         Left            =   -74910
         TabIndex        =   4
         Top             =   390
         Width           =   6435
         Begin VB.ListBox lstLog 
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   4155
            ItemData        =   "frmMain.frx":035E
            Left            =   120
            List            =   "frmMain.frx":0360
            TabIndex        =   5
            Top             =   210
            Width           =   6195
         End
      End
   End
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "Clear"
      Height          =   405
      Left            =   2100
      TabIndex        =   2
      Top             =   0
      Width           =   1035
   End
   Begin VB.CommandButton cmdConfiguration 
      Caption         =   "Config"
      Height          =   405
      Left            =   1050
      TabIndex        =   1
      Top             =   0
      Width           =   1035
   End
   Begin VB.Timer tmrClient 
      Index           =   0
      Interval        =   10
      Left            =   4470
      Top             =   0
   End
   Begin VB.Timer tmrServer 
      Index           =   0
      Interval        =   10
      Left            =   4020
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   0
      Left            =   3150
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   3570
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Start"
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1035
   End
   Begin VB.Menu mnuHeader 
      Caption         =   "Header"
      Visible         =   0   'False
      Begin VB.Menu mnuClearHeader 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "Log"
      Visible         =   0   'False
      Begin VB.Menu mnuClearLog 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEBUG_MODE = False

'collection to hold data buffer for each connections
'serverconnection keep the buffer for the sckserver socket (which sent the data to browser client)
'clientconnection keep the buffer for the sckclient socket (which sent the data to the host server or the proxy server)
Dim ServerConnection As Collection
Dim ClientConnection As Collection

Private Sub cmdClearLog_Click()
    lstLog.Clear
    txtRequest.Text = ""
    txtResponse.Text = ""
End Sub

Private Sub cmdConfiguration_Click()
    LoadConfigurationScreen Me, netProxy
End Sub

Private Sub cmdSwitch_Click()
    If cmdSwitch.Caption = "Start" Then
        StartProxy ListeningPort
        cmdSwitch.Caption = "Stop"
    Else
        StopProxy
        cmdSwitch.Caption = "Start"
    End If
End Sub

Private Sub StartProxy(LocalPort As Long)
    'starting the server by binding the local port property to the listening port we use for proxy
    'and issue the listen method to start listening
    SendToLog "Initializing Proxy server"
    InitializeSocket sckServer(0)
    sckServer(0).LocalPort = LocalPort
    sckServer(0).Listen
    SendToLog "Listening on port " & LocalPort
End Sub

Private Sub StopProxy()
Dim Socket As Winsock

    'looping each connections and close 'em
    SendToLog "Disconnecting all connection"
    For Each Socket In sckServer
        tmrServer(Socket.Index).Enabled = False
        tmrClient(Socket.Index).Enabled = False
        DoEvents
        CloseSocket Socket.Index
    Next
    SendToLog "Proxy server stopped"
End Sub

Private Sub InitializeSocket(Socket As Winsock)
On Error Resume Next

    'initialize socket before use
    SendToLog "Initialize Socket " & Socket.LocalPort
    Socket.Close
    Socket.LocalPort = 0
End Sub

Private Sub SendToLog(Message As String)
    lstLog.AddItem "[" & Now & "] " & Message
    If lstLog.ListCount > 10000 Then lstLog.Clear
End Sub

Private Sub flxStatistic_Click()
Dim i As Long

    With flxStatistic
        For i = 1 To ConnectionRequest.Count
            DoEvents
            If i = .Rows Then
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = ConnectionRequest(i).IPAddress
                .TextMatrix(i, 2) = ConnectionRequest(i).HostName
            End If
            .TextMatrix(i, 3) = ConnectionRequest(i).Stat_Connect_Count
        Next i
    End With
End Sub

Private Sub Form_Load()

    LocalIP = sckServer(0).LocalIP
    LoadUser UserList, "UserList.txt"
    InitializeGrid
    
    Set netProxy = New CProxy
    
    If Len(Dir(App.Path & "\" & ConfigFileName)) = 0 Then LocalComputerName = sckServer(0).LocalHostName
    LoadProxyConfiguration
    
    Set ServerConnection = New Collection
    Set ClientConnection = New Collection
End Sub

Private Sub Form_Resize()
    tabProxy.Width = Me.ScaleWidth
    If Me.ScaleHeight > (cmdSwitch.Height - 30) Then
        tabProxy.Height = Me.ScaleHeight - (cmdSwitch.Height - 30)
    End If
    If tabProxy.Width > 200 Then
        fraLog.Width = tabProxy.Width - 200
        fraStat.Width = tabProxy.Width - 200
    End If
    If tabProxy.Height > 500 Then
        fraLog.Height = tabProxy.Height - 500
        fraStat.Height = tabProxy.Height - 500
    End If
    If fraLog.Width > 200 Then
        lstLog.Width = fraLog.Width - 200
        flxStatistic.Width = fraStat.Width - 200
    End If
    If fraLog.Height > 240 Then
        lstLog.Height = fraLog.Height - 240
    End If
    If fraStat.Height > 320 Then
        flxStatistic.Height = fraStat.Height - 320
    End If
    fraRequest.Width = fraLog.Width
    fraRequest.Height = fraLog.Height \ 2
    If fraRequest.Height > 400 Then
        txtRequest.Height = fraRequest.Height - 400
    End If
    If fraRequest.Width > 300 Then
        txtRequest.Width = fraRequest.Width - 300
    End If
    fraResponse.Top = fraRequest.Top + fraRequest.Height
    fraResponse.Width = fraLog.Width
    fraResponse.Height = fraLog.Height \ 2
    If fraResponse.Height > 400 Then
        txtResponse.Height = fraResponse.Height - 400
    End If
    If fraResponse.Width > 300 Then
        txtResponse.Width = fraResponse.Width - 300
    End If
    
    With flxStatistic
        .ColWidth(0) = Abs(500 / 4500 * (.Width - 100))
        .ColWidth(1) = Abs(1000 / 4500 * (.Width - 100))
        .ColWidth(2) = Abs(2000 / 4500 * (.Width - 100))
        .ColWidth(3) = Abs(1000 / 4500 * (.Width - 100))
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Socket As Winsock

    'unload socket control array
    For Each Socket In sckClient
        CloseSocket Socket.Index
        If Socket.Index <> 0 Then
            Unload Socket
        End If
    Next
    
    For Each Socket In sckServer
        CloseSocket Socket.Index
        If Socket.Index <> 0 Then
            Unload Socket
        End If
    Next
        
    Set netProxy = Nothing
    Set ServerConnection = Nothing
    Set ClientConnection = Nothing
End Sub

Private Sub lstLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLog, vbPopupMenuRightButton, X + 240, Y + 1060
    End If
End Sub

Private Sub mnuClearHeader_Click()
    txtRequest.Text = ""
    txtResponse.Text = ""
End Sub

Private Sub mnuClearLog_Click()
    lstLog.Clear
End Sub

Private Sub sckClient_Close(Index As Integer)
    InitializeSocket sckClient(Index)
    ClientConnection(Index).ClearBuffer
End Sub

Private Sub sckClient_Connect(Index As Integer)
Dim vData As String

    'send the data request when connected
    If sckClient(Index).State = sckConnected Then
        vData = ClientConnection(Index).SendBuffer
        If Len(vData) <> 0 Then
            vData = ClientConnection(Index).SendBuffer.GetString
            SendDataTo sckClient(Index), vData
            SendToLog "Connected to Server " & sckClient(Index).RemoteHostIP & ":" & sckClient(Index).RemotePort
            If DEBUG_MODE Then Debug.Print "send to server " & vbCrLf & vData
        End If
    End If
End Sub

Private Sub sckClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim vData As String
Dim lpos As Long
Dim Header As String, Data As String

    If Index <> 0 And sckClient(Index).State = sckConnected Then
        SendToLog "Receive data from server " & sckClient(Index).RemoteHostIP & ":" & sckClient(Index).RemotePort & " size: " & bytesTotal & " bytes"
        
        'ok, this is rather tricky but here goes...
        'first we append the data to a buffer in serverconnection which it'll keep in seperate buffer
        'until the header is received.
        'when the header has been received then we send the header buffer to the sent buffer
        'which will be sent by the timer,
        'and we set the connected flag to true to tell the
        'append function to directly put the data in the sent buffer.
        
        sckClient(Index).GetData vData
        ServerConnection(Index).Append vData
        
        If ServerConnection(Index).HeaderReceived And Not ServerConnection(Index).Connected Then
            If DEBUG_MODE Then Debug.Print "received from server " & vbCrLf & ServerConnection(Index).Header
            Header = FilterResponseHeader(ServerConnection(Index).Header)
            ServerConnection(Index).SendBuffer = Header & vbCrLf & ServerConnection(Index).Data
            ServerConnection(Index).DataSent = ServerConnection(Index).DataSent + Len(ServerConnection(Index).Data)
            ServerConnection(Index).Connected = True
            SendResponseHeader "Socket " & Index & " :" & vbCrLf & Header
            If DEBUG_MODE Then Debug.Print "send to client buffer " & vbCrLf & Header
        End If
    End If
End Sub

Private Sub sckClient_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    InitializeSocket sckServer(Index)
    If Index <> 0 Then
        ClientConnection(Index).ClearBuffer
    End If
    
    InitializeSocket sckClient(Index)
    If Index <> 0 Then
        ServerConnection(Index).ClearBuffer
    End If
End Sub

Private Sub sckServer_Close(Index As Integer)
    CloseSocket Index
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Long, ActiveConnection As Long, ReceivingSocket As Winsock
    If Index = 0 Then
        'count the active connection
        ActiveConnection = 0
        For i = 0 To sckServer.Count - 1
            If i <> 0 Then If sckServer(i).State <> sckClosed Then ActiveConnection = ActiveConnection + 1
        Next i
    
        'receiving connection
        Set ReceivingSocket = AvailableSocket
        ReceivingSocket.Accept requestID
        
        'checking the maximum connection
        If ActiveConnection < MaximumConnection Then
            SendToLog "Accept connection request from client " & AvailableSocket.RemoteHostIP & ":" & ReceivingSocket.RemotePort
        Else
            ServerConnection(ReceivingSocket.Index).Rejected = True
            SendToLog "Maximum connection reached, Connection request from client " & ReceivingSocket.RemoteHostIP & ":" & ReceivingSocket.RemotePort & " rejected"
        End If
    End If
    
    Set ReceivingSocket = Nothing
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim i As Long, lpos As Long
Dim vData As String
Static Blocking As Boolean
Dim Header As String

    If Index <> 0 And sckServer(Index).State = sckConnected Then
        SendToLog "Receive data from client " & sckServer(Index).RemoteHostIP & ":" & sckServer(Index).RemotePort & " size: " & bytesTotal & " bytes"
            
        sckServer(Index).GetData vData
        ClientConnection(Index).Append vData
        
        If ClientConnection(Index).HeaderReceived And Not ClientConnection(Index).Connected Then
            'if header received then check the authorization status,
            'if rejected because maximum connection reached then we send the rejected form
            'if authorization is needed we send the request for password form.
        
            If DEBUG_MODE Then Debug.Print "received from client " & vbCrLf & ClientConnection(Index).Header
            If ServerConnection(Index).Rejected Then
                'maximum connection reached
                Header = GenerateHTMLForm(ftRejected)
                ServerConnection(Index).SendBuffer = Header
                SendResponseHeader "Socket " & Index & " :" & vbCrLf & Header
                DoEvents
                CloseSocket Index
                If DEBUG_MODE Then Debug.Print "send to client buffer " & vbCrLf & Header
            ElseIf Not ServerConnection(Index).AuthorizeUser Then
                ServerConnection(Index).AuthorizeUser = CheckCredential(sckServer(Index), ClientConnection(Index).Header)
                If Not ServerConnection(Index).AuthorizeUser Then
                    'not yet authorized
                    Header = GenerateHTMLForm(ftAuthenticate)
                    ServerConnection(Index).SendBuffer = Header
                    ClientConnection(Index).ClearBuffer
                    SendResponseHeader "Socket " & Index & " :" & vbCrLf & Header
                    If DEBUG_MODE Then Debug.Print "send to client buffer " & vbCrLf & Header
                ElseIf Left$(ClientConnection(Index).Header, 7) = "OPTIONS" Then
                    'this form used when the browser sent OPTION Method (Opera)
                    'i don't know what it's used for... but we just send them not found form
                    Header = GenerateHTMLForm(ftNotFound)
                    ServerConnection(Index).SendBuffer = Header
                    SendResponseHeader "Socket " & Index & " :" & vbCrLf & Header
                    DoEvents
                    CloseSocket Index
                Else
                    'authorized
                    InitializeSocket sckClient(Index)
                    Header = FilterRequestHeader(ClientConnection(Index).Header)
                    ClientConnection(Index).SendBuffer = Header & vbCrLf & ClientConnection(Index).Data
                    ClientConnection(Index).DataSent = ClientConnection(Index).DataSent + Len(ClientConnection(Index).Data)
                    ClientConnection(Index).Connected = True
                    SendRequestHeader "Socket " & Index & " :" & vbCrLf & Header
                    AddConnectionStatistic sckServer(Index)
                    If DEBUG_MODE Then Debug.Print "send to server buffer " & vbCrLf & Header
                End If
            ElseIf Left$(ClientConnection(Index).Header, 7) = "OPTIONS" Then
                'this form used when the browser sent OPTION Method (Opera)
                'i don't know what it's used for... but we just send them not found form
                Header = GenerateHTMLForm(ftNotFound)
                ServerConnection(Index).SendBuffer = Header
                SendResponseHeader "Socket " & Index & " :" & vbCrLf & Header
                DoEvents
                CloseSocket Index
            Else
                'ok
                InitializeSocket sckClient(Index)
                Header = FilterRequestHeader(ClientConnection(Index).Header)
                ClientConnection(Index).SendBuffer = Header & vbCrLf & ClientConnection(Index).Data
                ClientConnection(Index).DataSent = ClientConnection(Index).DataSent + Len(ClientConnection(Index).Data)
                ClientConnection(Index).Connected = True
                SendRequestHeader "Socket " & Index & " :" & vbCrLf & Header
                AddConnectionStatistic sckServer(Index)
                If DEBUG_MODE Then Debug.Print "send to server buffer " & vbCrLf & Header
            End If
        End If
    End If
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CloseSocket Index
End Sub

Private Function AvailableSocket() As Winsock
Dim Socket As Winsock

    'seach socket array for closed connection
    For Each Socket In sckServer
        DoEvents
        If Socket.State = sckClosed Then
            ServerConnection(Socket.Index).ClearBuffer
            ClientConnection(Socket.Index).ClearBuffer
            Set AvailableSocket = Socket
            Exit Function
        End If
    Next
    
    'if there aren't any closed connection we load a new socket to received the connection request
    Set AvailableSocket = AddNewConnection
End Function

Private Function AddNewConnection() As Winsock
Dim ServerData As New CBuffer
Dim ClientData As New CBuffer
Dim NewSocket As Long
    
    NewSocket = sckServer.Count
    
    Load sckServer(NewSocket)
    Load tmrServer(NewSocket)
    ServerData.HeaderType = htResponse
    ServerData.ClearBuffer
    ServerConnection.Add ServerData, Chr(NewSocket)
    
    Load sckClient(NewSocket)
    Load tmrClient(NewSocket)
    ClientData.HeaderType = htRequest
    ClientData.ClearBuffer
    ClientConnection.Add ClientData, Chr(NewSocket)

    Set AddNewConnection = sckServer(NewSocket)
End Function

Private Sub sckServer_SendComplete(Index As Integer)
    'we close both connection when the requested data has been sent to the browser.
    If ServerConnection(Index).DataSent >= ServerConnection(Index).ResponseHeader.Length Then
        If ServerConnection(Index).ResponseHeader.GetHeader("Content-Length") <> "" Then
            If sckClient(Index).State <> sckConnected And Len(ServerConnection(Index).SendBuffer) = 0 Then
                CloseSocket Index
            End If
        End If
    End If
End Sub

Private Sub tmrClient_Timer(Index As Integer)
Dim i As Long
Dim vData As String

    'timer to send the data to host server/proxy server if there are any data in the sent buffer
    If Index <> 0 Then
        i = Index
        vData = ClientConnection(i).SendBuffer
        If Len(vData) <> 0 Then
            If sckClient(i).State <> sckConnected And sckClient(i).State <> sckConnecting Then
                ConnectSocket sckClient(i), ClientConnection(i)
            ElseIf sckClient(i).State = sckConnected And Len(vData) <> 0 Then
                vData = ClientConnection(i).SendBuffer.GetString
                SendDataTo sckClient(i), vData
                If DEBUG_MODE Then Debug.Print "send to server " & vbCrLf & vData
            End If
        End If
    End If
End Sub

Private Sub SendDataTo(Socket As Winsock, vData As String)
    'send data to specified socket
    Socket.SendData vData
    SendToLog "Sending data to " & Socket.RemoteHostIP & ":" & Socket.RemotePort & " Size:" & Len(vData)
End Sub

Private Sub ConnectSocket(Socket As Winsock, BufferConnection As CBuffer)
Dim vProxyServer As String, vProxyPort As Long

    On Error GoTo errHandler
    
    'connecting to host server or proxy server
    If UseProxy Then
        vProxyServer = netProxy.Server
        vProxyPort = netProxy.Port
    Else
        vProxyServer = BufferConnection.Server
        vProxyPort = BufferConnection.Port
    End If
    Socket.Connect vProxyServer, vProxyPort
    DoEvents
    SendToLog "Connecting to server " & vProxyServer & ":" & vProxyPort
    Exit Sub
errHandler:
End Sub

Private Sub tmrServer_Timer(Index As Integer)
Dim i As Long
Dim vData As String

    If Index <> 0 Then
        i = Index
        If sckServer(i).State = sckConnected Then
            vData = ServerConnection(i).SendBuffer
            If Len(vData) <> 0 Then
                vData = ServerConnection(i).SendBuffer.GetString
                SendDataTo sckServer(i), vData
                If DEBUG_MODE Then Debug.Print "send to client " & vbCrLf & vData
            End If
        End If
    End If
End Sub

Private Sub CloseSocket(Index As Integer)
    'close both socket connection (client and server)
    
    InitializeSocket sckClient(Index)
    If Index <> 0 Then
        ServerConnection(Index).ClearBuffer
    End If
    
    InitializeSocket sckServer(Index)
    If Index <> 0 Then
        ClientConnection(Index).ClearBuffer
    End If
End Sub

Private Sub SendRequestHeader(Message As String)
    'showing the request header to textbox
    If Len(txtRequest.Text) > 16384 Then
        txtRequest.Text = ""
    End If
    txtRequest.Text = txtRequest.Text & Message & vbCrLf
End Sub

Private Sub SendResponseHeader(Message As String)
    'showing the response header to textbox
    If Len(txtResponse.Text) > 16384 Then
        txtResponse.Text = ""
    End If
    txtResponse.Text = txtResponse.Text & Message & vbCrLf
End Sub

Private Sub InitializeGrid()
    'init statistic grid
    With flxStatistic
        .Clear
        .Rows = 1
        .Cols = 4
        
        .ColAlignment(0) = flexAlignLeftCenter
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "IP Address"
        .TextMatrix(0, 2) = "Host Name"
        .TextMatrix(0, 3) = "Connection"
    End With
End Sub

