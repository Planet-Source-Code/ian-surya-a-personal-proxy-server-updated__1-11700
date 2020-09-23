VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Passthrough Server"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrClient 
      Index           =   0
      Interval        =   10
      Left            =   4080
      Top             =   0
   End
   Begin VB.Timer tmrServer 
      Index           =   0
      Interval        =   10
      Left            =   3390
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   0
      Left            =   2130
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   2700
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   405
      Left            =   1020
      TabIndex        =   2
      Top             =   0
      Width           =   1035
   End
   Begin VB.TextBox txtLog 
      Height          =   3675
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   420
      Width           =   6735
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Listen"
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private UseNTLM As Boolean
Private netProxy As New InetProxy

Private ProxyServer As String
Private ProxyPort As Long

Dim ServerConnection As Collection
Dim ClientConnection As Collection

Private Sub cmdSwitch_Click()
Dim Socket As Winsock

    If cmdSwitch.Caption = "Listen" Then
        InitializeSocket sckServer(0)
        sckServer(0).LocalPort = 8080
        sckServer(0).Listen
        SendMessage "Socket(0) listening on port " & sckServer(0).LocalPort
        cmdSwitch.Caption = "Stop"
    Else
        InitializeSocket sckServer(0)
        SendMessage "Socket(0) stop listening"
        cmdSwitch.Caption = "Listen"
    
        For Each Socket In sckServer
            CloseSocket Socket.Index
        Next
    End If
End Sub

Private Sub Command1_Click()
    txtLog = ""
End Sub

Private Sub InitializeSocket(Socket As Winsock)
On Error Resume Next
    Socket.Close
    Socket.LocalPort = 0
End Sub

Private Sub SendMessage(Message As String)
    txtLog = txtLog & Message & vbCrLf & vbCrLf
    If Len(txtLog) > 20000 Then txtLog = ""
    txtLog.SelStart = Len(txtLog)
End Sub

Private Sub Form_Load()
    UseNTLM = True
    
    SetProxy Me, netProxy
    netProxy.Access = inetNamedProxy
    netProxy.Server = "HO_PROXY"
    netProxy.Port = 80
    
    SetLocalProxy netProxy.Server, netProxy.Port
    
    Set ServerConnection = New Collection
    Set ClientConnection = New Collection
End Sub

Private Sub Form_Resize()
    
    txtLog.Width = Me.ScaleWidth
    If Me.WindowState <> vbMinimized Then txtLog.Height = Abs(Me.ScaleHeight - txtLog.Top)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Socket As Winsock

    For Each Socket In sckServer
        CloseSocket Socket.Index
        If Socket.Index <> 0 Then Unload Socket
    Next
    
    For Each Socket In sckClient
        CloseSocket Socket.Index
        If Socket.Index <> 0 Then Unload Socket
    Next
    
    Set ServerConnection = Nothing
    Set ClientConnection = Nothing
End Sub

Private Sub SetLocalProxy(HostName As String, HostPort As Long)
    ProxyServer = HostName
    ProxyPort = HostPort
End Sub

Private Sub sckClient_Close(Index As Integer)
    InitializeSocket sckClient(Index)
    ClientConnection(Index).ClearBuffer
End Sub

Private Sub sckClient_Connect(Index As Integer)
Dim Data As String, vHeader As String, vData As String

    vHeader = ClientConnection(Index).BufferHeader
    vData = ClientConnection(Index).BufferData
    If sckClient(Index).State = sckConnected Then
        sckClient(Index).SendData vHeader & vData
        ClientConnection(Index).LastBuffer = vHeader & vData
        ClientConnection(Index).BufferHeader = Mid(ClientConnection(Index).BufferHeader, Len(vHeader) + 1)
        ClientConnection(Index).BufferData = Mid(ClientConnection(Index).BufferData, Len(vData) + 1)
    End If
End Sub

Private Sub sckClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
Dim i As Long
Dim lpos As Long
Dim tmpHeader As New HttpHeader
Static Header As String
Dim AuthorizationString As String, tmpString As String

    If Index <> 0 And sckClient(Index).State = sckConnected Then
        sckClient(Index).GetData Data
        SendMessage "Socket " & Index & " to server receive data from server " & sckClient(Index).RemoteHostIP & ":" & sckClient(Index).RemotePort & " size: " & bytesTotal & " bytes"
        
        'If UseNTLM Then Debug.Print Data
        
        If Not ServerConnection(Index).HeaderReceived Then
            If IsHTTPHeader(Data) Then
                lpos = InStr(1, Data, vbCrLf & vbCrLf, vbTextCompare)
                If lpos = 0 Then
                    Header = Data
                Else
                    Header = Header & Data
                    lpos = InStr(1, Header, vbCrLf & vbCrLf, vbTextCompare)
                    tmpHeader.ParseHeader Left$(Header, lpos + 1)
                    If tmpHeader.Status = inetProxyUnauthorized And UseNTLM Then
                        If tmpHeader.GetHeader("Proxy-Authenticate") = "NTLM" Then
                            ClientConnection(Index).NTLM.CloseSecurity
                        End If
                        If ClientConnection(Index).NTLM.NTLMAuthenticate(tmpHeader) Then
                            tmpString = ClientConnection(Index).LastBuffer
                            tmpString = Replace(tmpString, "Proxy-Connection: Keep-Alive" & vbCrLf, "", 1, 1)
                            tmpString = ReplaceAuthorization(tmpString)
                            AuthorizationString = "Proxy-Connection: Keep-Alive" & vbCrLf & ClientConnection(Index).NTLM.GetNTLMToken
                            tmpString = Replace(tmpString, vbCrLf, vbCrLf & AuthorizationString, 1, 1)
                            tmpString = Replace(tmpString, "HTTP/1.0", "HTTP/1.1", 1, 1)
                        End If
                        If LCase(tmpHeader.GetHeader("Proxy-Connection")) = "keep-alive" Then
                            ClientConnection(Index).ClearBuffer
                            ClientConnection(Index).AppendBuffer tmpString
                        Else
                            InitializeSocket sckClient(Index)
                            ClientConnection(Index).ClearBuffer
                            ClientConnection(Index).AppendBuffer tmpString
                            sckClient(Index).Connect ProxyServer, ProxyPort
                        End If
                    Else
                        'ClientConnection(Index).NTLM.CloseSecurity
                        ServerConnection(Index).AppendBuffer Data
                    End If
                    Header = ""
                End If
            Else
                'ClientConnection(Index).NTLM.CloseSecurity
                ServerConnection(Index).AppendBuffer Data
            End If
        Else
            'ClientConnection(Index).NTLM.CloseSecurity
            ServerConnection(Index).AppendBuffer Data
        End If
    End If
End Sub

Private Function ReplaceAuthorization(Data As String) As String
Dim lpos As Long
Dim endpos As Long
Dim AuthString As String

    lpos = InStr(1, Data, "Proxy-Authorization:", vbTextCompare)
    If lpos <> 0 Then
        endpos = InStr(lpos + 1, Data, vbCrLf, vbTextCompare)
        AuthString = Mid$(Data, lpos, endpos - lpos)
        ReplaceAuthorization = Replace(Data, AuthString & vbCrLf, "", 1, 1, vbTextCompare)
    Else
        ReplaceAuthorization = Data
    End If
    
End Function

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
    If Index = 0 Then
        AvailableSocket.Accept requestID
        SendMessage "Socket 0 to client receive request from client " & AvailableSocket.RemoteHostIP & ":" & AvailableSocket.RemotePort
    End If
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim i As Long
Dim Data As String
Dim tmpData As String

    If Index <> 0 And sckServer(Index).State = sckConnected Then
        sckServer(Index).GetData Data
        SendMessage "Socket " & Index & " to client receive data from client " & sckServer(Index).RemoteHostIP & ":" & sckServer(Index).RemotePort & " size: " & bytesTotal & " bytes"
        ClientConnection(Index).AppendBuffer Data
    End If
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CloseSocket Index
End Sub

Private Function AvailableSocket() As Winsock
Dim ServerData As New CBuffer
Dim ClientData As New CBuffer
Dim Socket As Winsock
Dim NewSocket As Long

    For Each Socket In sckServer
        If Socket.State = sckClosed Then
            ServerConnection(Socket.Index).ClearBuffer
            ClientConnection(Socket.Index).ClearBuffer
            Set AvailableSocket = Socket
            Exit Function
        End If
    Next
    
    NewSocket = sckServer.Count
    Load sckServer(NewSocket)
    Load sckClient(NewSocket)
    Load tmrServer(NewSocket)
    Load tmrClient(NewSocket)
    Set ServerData.NTLM.Proxy = netProxy
    Set ClientData.NTLM.Proxy = netProxy
    ServerConnection.Add ServerData, Chr(NewSocket)
    ClientConnection.Add ClientData, Chr(NewSocket)
    Set AvailableSocket = sckServer(NewSocket)
End Function

Private Sub sckServer_SendComplete(Index As Integer)
    'CloseSocket Index
End Sub

Private Sub tmrClient_Timer(Index As Integer)
Dim i As Long
Dim Data As String, vHeader As String, vData As String

    On Error GoTo errHandler

    i = Index
    vHeader = ClientConnection(i).BufferHeader
    vData = ClientConnection(i).BufferData
    If Len(vHeader) <> 0 Or Len(vData) <> 0 Then
        If sckClient(i).State <> sckConnected Then
            sckClient(i).Connect ProxyServer, ProxyPort
        End If
        If sckClient(i).State = sckConnected Then
            sckClient(i).SendData vHeader & vData
            ClientConnection(Index).LastBuffer = vHeader & vData
            ClientConnection(i).BufferHeader = Mid(ClientConnection(i).BufferHeader, Len(vHeader) + 1)
            ClientConnection(i).BufferData = Mid(ClientConnection(i).BufferData, Len(vData) + 1)
            SendMessage "Socket " & i & " to server sending data"
        End If
    End If
    Exit Sub
errHandler:
End Sub

Private Sub tmrServer_Timer(Index As Integer)
Dim i As Long
Dim Data As String, vHeader As String, vData As String

    On Error GoTo errHandler

    i = Index
    vHeader = ServerConnection(i).BufferHeader
    vData = ServerConnection(i).BufferData
    If Len(vHeader) <> 0 Or Len(vData) <> 0 Then
        vHeader = ServerConnection(i).BufferHeader
        vData = ServerConnection(i).BufferData
        If sckServer(i).State = sckConnected Then
            sckServer(i).SendData vHeader & vData
            ServerConnection(Index).LastBuffer = vHeader & vData
            ServerConnection(i).BufferHeader = Mid(ServerConnection(i).BufferHeader, Len(vHeader) + 1)
            ServerConnection(i).BufferData = Mid(ServerConnection(i).BufferData, Len(vData) + 1)
            SendMessage "Socket " & i & " to client sending data"
        End If
    End If
    Exit Sub
errHandler:
End Sub

Private Sub CloseSocket(Index As Integer)
    InitializeSocket sckServer(Index)
    If Index <> 0 Then
        ClientConnection(Index).ClearBuffer
    End If
    
    InitializeSocket sckClient(Index)
    If Index <> 0 Then
        ServerConnection(Index).ClearBuffer
    End If
End Sub
