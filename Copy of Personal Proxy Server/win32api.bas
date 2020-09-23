Attribute VB_Name = "modWin32API"
Public Const errNotFound = ""

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function WriteInitString(ByVal appname As String, ByVal KeyName As String, DescString As String, FileName As String) As Long
    errnumber& = WritePrivateProfileString(appname, KeyName, DescString, FileName)
    WriteInitString = errnumber
End Function

Function GetInitString(ByVal appname As String, ByVal KeyName As String, ByVal FileName As String) As String
    Dim Defaultstr As String
    Defaultstr = errNotFound
    ReturnedString$ = String(255, Chr$(0))
    StringSize& = GetPrivateProfileString(appname, KeyName, Defaultstr, ReturnedString, Len(ReturnedString), FileName)
    GetInitString = Left(ReturnedString, StringSize)
End Function
