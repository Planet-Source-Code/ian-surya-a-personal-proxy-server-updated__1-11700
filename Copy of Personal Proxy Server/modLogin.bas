Attribute VB_Name = "modLogin"
Option Explicit

Public UserList As Collection
Public InvalidList As Collection

Public Sub InitializeList(Col As Collection)
    Set Col = New Collection
End Sub

Public Sub AddUser(Col As Collection, Authorization As String)
Dim User As New Cconnection

    User.Key = Authorization
    If Not IsInCollection(Col, User.Key) Then
        Col.Add User, User.Key
    End If
    
    Set User = Nothing
End Sub

'save the user in user list to text file
Public Sub SaveUser(Col As Collection, vFileName As String)
Dim ff As Integer
Dim i As Long
Dim tmpString As String
    
    If Col.Count <> 0 Then
        tmpString = ""
        For i = 1 To Col.Count
            tmpString = tmpString & DoXOR(Col(i).Key, "PersonalProxyServer") & vbCrLf
        Next i
        
        ff = FreeFile
        Open App.Path & "\" & vFileName For Output As #ff
        Print #ff, tmpString
        Close #ff
    End If
End Sub

'load the user list from file
Public Sub LoadUser(Col As Collection, vFileName As String)
Dim ff As Integer
Dim i As Long
Dim tmpString As String

    InitializeList Col
    If Dir(App.Path & "\" & vFileName) = vFileName Then
        
        ff = FreeFile
        Open App.Path & "\" & vFileName For Input As #ff
            Do While Not EOF(ff)
                Line Input #ff, tmpString
                If tmpString <> "" Then AddUser Col, DoXOR(tmpString, "PersonalProxyServer")
            Loop
        Close #ff
    End If
End Sub

'checking if the authorization string the client sent match the user list
Public Function CheckCredential(Socket As Winsock, Header As String) As Boolean
Dim AuthorizationString As String

    'skip if the request comes from local computer
    If Socket.RemoteHostIP = LocalIP Or Socket.RemoteHostIP = "127.0.0.1" Then
        CheckCredential = True
    Else
        AuthorizationString = GetHttpHeader(Header, "Proxy-Authorization")
        If UCase(Left$(AuthorizationString, 5)) = "BASIC" Then
            AuthorizationString = Mid$(AuthorizationString, 7)
            If IsInUserCollection(UserList, AuthorizationString) Then
                If IsInCollection(ConnectionRequest, Socket.RemoteHostIP) Then
                    ConnectionRequest(Socket.RemoteHostIP).Authorized = True
                End If
                CheckCredential = True
            End If
        End If
    End If
End Function

Public Function GetUser(AuthorizationString As String) As String
Dim tmpString As String
Dim lpos As Long

    tmpString = Base64Decode(AuthorizationString)
    lpos = InStr(1, tmpString, ":", vbTextCompare)
    If lpos <> 0 Then
        GetUser = Left$(tmpString, lpos - 1)
    Else
        GetUser = tmpString
    End If
End Function

Public Function GetPassword(AuthorizationString As String) As String
Dim tmpString As String
Dim lpos As Long

    tmpString = Base64Decode(AuthorizationString)
    lpos = InStr(1, tmpString, ":", vbTextCompare)
    If lpos <> 0 Then
        GetPassword = Mid$(tmpString, lpos + 1)
    Else
        GetPassword = ""
    End If
End Function

Public Function IsUserExist(UserName As String) As Long
Dim i As Long

    For i = 1 To UserList.Count
        If UCase(GetUser(UserList(i).Key)) = UCase(UserName) Then
            IsUserExist = i
            Exit For
        End If
    Next i
End Function
