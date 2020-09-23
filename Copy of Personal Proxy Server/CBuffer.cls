VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public AuthenticationCounter As Long
Public LastAuthorization As String

Public LastHeader As String

Public Enum ENUM_HEADER_TYPE
    htRequest
    htResponse
End Enum

Public HeaderType As ENUM_HEADER_TYPE
Public HeaderReceived As Boolean
Public ResponseHeader As CHttpHeader
Public Header As CStringBuffer
Public Data As CStringBuffer

Private HeaderBuffer As CStringBuffer

Public Connected As Boolean
Public SendBuffer As CStringBuffer
Public DataSent As Long

Public AuthorizeUser As Boolean
Public Rejected As Boolean

Public Sub ClearBuffer()
    AuthorizeUser = Not UseAuthentication
    Rejected = False
    Connected = False
    DataSent = 0
    SendBuffer.Clear
    HeaderBuffer.Clear
    Data.Clear
    Header.Clear
    HeaderReceived = False
    Set ResponseHeader = Nothing
    Set ResponseHeader = New CHttpHeader
End Sub

Public Function Server() As String
Dim lpos As Long
    
    Server = GetRequestHeader("Host")
    lpos = InStr(1, Server, ":", vbTextCompare)
    If lpos <> 0 Then
        Server = Left$(Server, lpos - 1)
    End If
End Function

Public Function Port() As Long
Dim lpos As Long, vServer As String
    
    Port = 80
    vServer = GetRequestHeader("Host")
    lpos = InStr(1, vServer, ":", vbTextCompare)
    If lpos <> 0 Then
        Port = Val(Mid$(vServer, lpos + 1))
    End If
End Function

Private Function GetRequestHeader(Optional HeaderCaption As String = "") As String
Dim lpos As Long, endpos As Long

    GetRequestHeader = Header
    If HeaderCaption <> "" Then
        lpos = InStr(1, Header, HeaderCaption & ": ", vbTextCompare)
        If lpos <> 0 Then
            endpos = InStr(lpos, Header, vbCrLf, vbTextCompare)
            GetRequestHeader = Mid$(Header, lpos + 6, endpos - (lpos + 6))
        End If
    End If
End Function

Public Sub Append(vData As String)
Dim lpos As Long

    If Connected Then
        SendBuffer.Append vData
    Else
        If Not HeaderReceived Then
            If IsHTTPHeader(vData) Then
                ClearBuffer
                HeaderBuffer = vData
            Else
                HeaderBuffer.Append vData
            End If
            lpos = InStr(1, HeaderBuffer, vbCrLf & vbCrLf, vbTextCompare)
            If lpos <> 0 Then
                Data = Mid$(HeaderBuffer, lpos + 4)
                Header = Left$(HeaderBuffer, lpos + 1)
                If HeaderType = htResponse Then
                    ResponseHeader.ParseHeader Left$(HeaderBuffer, lpos + 1)
                End If
                HeaderReceived = True
                HeaderBuffer = ""
                LastHeader = Header
            End If
        Else
            Data.Append vData
        End If
    End If
End Sub

Private Sub Class_Initialize()
    Set ResponseHeader = New CHttpHeader
    Set SendBuffer = New CStringBuffer
    Set HeaderBuffer = New CStringBuffer
    Set Data = New CStringBuffer
    Set Header = New CStringBuffer
End Sub

Private Sub Class_Terminate()
    Set ResponseHeader = Nothing
    Set SendBuffer = Nothing
    Set HeaderBuffer = Nothing
    Set Data = Nothing
    Set Header = Nothing
End Sub

Private Function RedirectHeader(Header As String) As String
Dim httpMethod As String
Dim Resource As String
Dim Host As String
Dim lpos As Long
Dim endpos As Long
Dim newResource As String
Dim newHost As String
Dim newHeader As String

    httpMethod = UCase(Trim(Left$(Header, 4)))
    lpos = Len(httpMethod) + 2
    endpos = InStr(lpos, Header, " ", vbTextCompare)
    Resource = Mid$(Header, lpos, endpos - lpos)
    
    Host = GetHttpHeader(Header, "Host")
    
    lpos = InStr(1, Resource, "://", vbTextCompare)
    If lpos <> 0 And Left$(Resource, 14) = "/~redirected~/" Then
        newResource = Mid$(Resource, lpos + 3)
        lpos = InStr(1, newResource, "/", vbTextCompare)
        If lpos <> 0 Then
            newHost = Left$(newResource, lpos - 1)
            newResource = Mid$(newResource, lpos)
        Else
            newHost = newResource
            newResource = "/"
        End If
        newHeader = Replace(Header, Resource, newResource, 1, 1)
        newHeader = AddHttpHeader(newHeader, "Host", newHost)
        RedirectHeader = newHeader
    Else
        RedirectHeader = Header
    End If
End Function
