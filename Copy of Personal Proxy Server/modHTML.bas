Attribute VB_Name = "modHTML"
Option Explicit

Public Enum ENUM_FORM_TYPE
    ftRejected
    ftAuthenticate
    ftNotFound
End Enum

'get the description for http error message
Private Function GetErrorMessage(ErrorNo As Long) As String
Dim ErrorDesc As String

    Select Case ErrorNo
    Case 200
        ErrorDesc = "OK"
    Case 201
        ErrorDesc = "Created"
    Case 202
        ErrorDesc = "Accepted"
    Case 204
        ErrorDesc = "No Content"
    Case 301
        ErrorDesc = "Moved Permanently"
    Case 302
        ErrorDesc = "Moved Temporarily"
    Case 304
        ErrorDesc = "Not Modified"
    Case 400
        ErrorDesc = "Bad Request"
    Case 401
        ErrorDesc = "Unauthorized"
    Case 403
        ErrorDesc = "Forbidden"
    Case 404
        ErrorDesc = "Not Found"
    Case 407
        ErrorDesc = "Proxy authentication required"
    Case 500
        ErrorDesc = "Internal Server Error"
    Case 501
        ErrorDesc = "Not Implemented"
    Case 502
        ErrorDesc = "Bad Gateway"
    Case 503
        ErrorDesc = "Service Unavailable"
    Case Else
        ErrorDesc = "Extended Code"
    End Select
    
    GetErrorMessage = ErrorNo & " " & ErrorDesc
End Function

'generate html form
Public Function GenerateHTMLForm(FormType As ENUM_FORM_TYPE) As String
Dim Header As String
Dim Data As String

    Select Case FormType
    Case ftRejected
        Data = "Error Access denied, Connection limit reached."
        Header = "HTTP/1.0 " & GetErrorMessage(403) & vbCrLf
        Header = Header & "Server" & ": " & PersonalProxyName & vbCrLf
        Header = Header & "Content-type" & ": " & "text/html" & vbCrLf
        Header = Header & "Date" & ": " & Format(Now, "ddd, dd mmm yyyy hh:mm:ss") & " GMT" & vbCrLf
        Header = Header & "Content-Length" & ": " & Len(Data) & vbCrLf
        Header = Header & "Connection" & ": " & "close" & vbCrLf
    Case ftNotFound
        Data = "Object not found."
        Header = "HTTP/1.0 " & GetErrorMessage(404) & vbCrLf
        Header = Header & "Server" & ": " & PersonalProxyName & vbCrLf
        Header = Header & "Content-type" & ": " & "text/html" & vbCrLf
        Header = Header & "Date" & ": " & Format(Now, "ddd, dd mmm yyyy hh:mm:ss") & " GMT" & vbCrLf
        Header = Header & "Content-Length" & ": " & Len(Data) & vbCrLf
        Header = Header & "Connection" & ": " & "close" & vbCrLf
    Case ftAuthenticate
        Data = "Error Access denied, authentication required."
        Header = "HTTP/1.0 " & GetErrorMessage(407) & vbCrLf
        Header = Header & "Proxy-Authenticate" & ": " & "Basic" & " " & "realm=Personal Proxy Server" & vbCrLf
        Header = Header & "Server" & ": " & PersonalProxyName & vbCrLf
        Header = Header & "Content-type" & ": " & "text/html" & vbCrLf
        Header = Header & "Date" & ": " & Format(Now, "ddd, dd mmm yyyy hh:mm:ss") & " GMT" & vbCrLf
        Header = Header & "Content-Length" & ": " & Len(Data) & vbCrLf
        Header = Header & "Proxy-Connection" & ": " & "Keep-Alive" & vbCrLf
    End Select
    
    GenerateHTMLForm = Header & vbCrLf & Data
    
End Function
