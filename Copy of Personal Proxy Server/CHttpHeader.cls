VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHttpHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Header As String

Public Sub ParseHeader(HeaderData As String)
    Header = HeaderData
End Sub

Public Function Length() As Long
    Length = Val(GetHeader("Content-Length"))
End Function

Public Function GetHeader(Optional HeaderCaption As String = "") As String
Dim lpos As Long
Dim endpos As Long
Dim HeaderData As String

    If HeaderCaption <> "" Then
        lpos = InStr(1, Header, HeaderCaption & ":", vbTextCompare)
        If lpos <> 0 Then
            endpos = InStr(lpos + 1, Header, vbCrLf, vbTextCompare)
            HeaderData = Mid$(Header, lpos + Len(HeaderCaption) + 2, endpos - (lpos + Len(HeaderCaption) + 2))
        End If
        GetHeader = HeaderData
    Else
        GetHeader = Header
    End If
End Function

