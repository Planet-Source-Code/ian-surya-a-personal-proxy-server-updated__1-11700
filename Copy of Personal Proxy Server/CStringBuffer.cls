VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const BUFFER_LENGTH = 2 ^ 16

Private m_Length As Long
Private m_Data As String

Public Property Get Data() As String
Attribute Data.VB_UserMemId = 0
    Data = Left$(m_Data, m_Length)
End Property

Public Property Let Data(vData As String)
    m_Data = vData
    m_Length = Len(vData)
End Property

Public Sub Append(vData As String)
Dim lenData As Long

    lenData = Len(vData)
    If Len(m_Data) - m_Length <= lenData Then
        m_Data = m_Data & String$(BUFFER_LENGTH, 0)
    End If
    Mid$(m_Data, m_Length + 1, lenData) = vData
    m_Length = m_Length + lenData
End Sub

Public Sub Clear()
    m_Data = ""
    m_Length = 0
End Sub

Public Function GetString() As String
    GetString = Left$(m_Data, m_Length)
    m_Data = Mid$(m_Data, m_Length + 1)
    m_Length = 0
End Function
