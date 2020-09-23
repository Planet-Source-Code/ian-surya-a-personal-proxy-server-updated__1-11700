Attribute VB_Name = "modStats"
Option Explicit

Public ConnectionRequest As New Collection

'add the connection request to statistic collection
Public Sub AddConnectionStatistic(Socket As Winsock)
Dim CStat As Cconnection

    If IsInCollection(ConnectionRequest, Socket.RemoteHostIP) Then
        Set CStat = ConnectionRequest(Socket.RemoteHostIP)
        CStat.Stat_Connect_Count = CStat.Stat_Connect_Count + 1
    Else
        Set CStat = New Cconnection
        CStat.IPAddress = Socket.RemoteHostIP
        CStat.HostName = Socket.RemoteHost
        CStat.Key = CStat.IPAddress
        CStat.Stat_Connect_Count = CStat.Stat_Connect_Count + 1
        ConnectionRequest.Add CStat, CStat.Key
    End If
    
    Set CStat = Nothing
End Sub

