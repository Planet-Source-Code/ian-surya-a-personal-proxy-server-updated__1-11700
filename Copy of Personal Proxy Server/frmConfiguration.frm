VERSION 5.00
Begin VB.Form frmConfiguration 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proxy Configuration"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5910
   Icon            =   "frmConfiguration.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConfiguration 
      Caption         =   "Proxy Configuration"
      Height          =   2955
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   5895
      Begin VB.CommandButton cmdUserList 
         Caption         =   "User List"
         Height          =   315
         Left            =   2130
         TabIndex        =   22
         Top             =   2040
         Width           =   1000
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Authenticate User"
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   19
         Top             =   2070
         Width           =   1845
      End
      Begin VB.TextBox txtProxySetting 
         Height          =   345
         Index           =   6
         Left            =   2130
         TabIndex        =   18
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtProxySetting 
         Height          =   345
         Index           =   5
         Left            =   2130
         TabIndex        =   16
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtProxySetting 
         Height          =   345
         Index           =   4
         Left            =   2130
         TabIndex        =   14
         Top             =   960
         Width           =   3615
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Replace User Agent"
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   17
         Top             =   1710
         Width           =   2055
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Hide Actual Server"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   990
         Width           =   2055
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Hide Actual Proxy"
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   15
         Top             =   1350
         Width           =   2055
      End
      Begin VB.CheckBox chkConfig 
         Alignment       =   1  'Right Justify
         Caption         =   "Disable Cookie"
         Height          =   285
         Index           =   1
         Left            =   4290
         TabIndex        =   12
         Top             =   630
         Width           =   1455
      End
      Begin VB.CheckBox chkConfig 
         Alignment       =   1  'Right Justify
         Caption         =   "Force Reload"
         Height          =   285
         Index           =   0
         Left            =   4290
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtProxySetting 
         Height          =   345
         Index           =   3
         Left            =   2130
         TabIndex        =   11
         Top             =   600
         Width           =   1755
      End
      Begin VB.TextBox txtProxySetting 
         Height          =   345
         Index           =   2
         Left            =   2130
         TabIndex        =   8
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label9 
         Caption         =   "Internet Explorer 3 or later. (Does not apply to local machine)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   690
         TabIndex        =   24
         Top             =   2670
         Width           =   4665
      End
      Begin VB.Label Label8 
         Caption         =   "Notes : Authentication required client browser that support basic authentication such as"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   2490
         Width           =   5625
      End
      Begin VB.Label Label7 
         Caption         =   "Local Port"
         Height          =   285
         Left            =   150
         TabIndex        =   10
         Top             =   660
         Width           =   1905
      End
      Begin VB.Label Label6 
         Caption         =   "Maximum Connection"
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   1905
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4830
      TabIndex        =   21
      Top             =   4110
      Width           =   1005
   End
   Begin VB.CheckBox chkUseProxy 
      Caption         =   "Use Proxy Server"
      Height          =   165
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   1575
   End
   Begin VB.Frame fraProxy 
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   5895
      Begin VB.TextBox txtProxySetting 
         Height          =   345
         Index           =   1
         Left            =   1140
         TabIndex        =   5
         Top             =   570
         Width           =   4665
      End
      Begin VB.TextBox txtProxySetting 
         Height          =   345
         Index           =   0
         Left            =   1140
         TabIndex        =   3
         Top             =   210
         Width           =   4665
      End
      Begin VB.Label Label2 
         Caption         =   "Proxy Port"
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Proxy Server"
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3780
      TabIndex        =   20
      Top             =   4110
      Width           =   1005
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub OpenCommunication(ParentForm As Form, commtype As Byte, ParamArray X())
    Select Case commtype
    Case 0
        chkUseProxy.Value = X(0)
        txtProxySetting(0).Text = X(1)
        txtProxySetting(1).Text = X(2)
        txtProxySetting(2).Text = X(3)
        txtProxySetting(3).Text = X(4)
        txtProxySetting(3).Tag = X(4)
        chkConfig(0).Value = X(5)
        chkConfig(1).Value = X(6)
        chkConfig(2).Value = X(7)
        chkConfig(3).Value = X(8)
        chkConfig(4).Value = X(9)
        txtProxySetting(4).Text = X(10)
        txtProxySetting(5).Text = X(11)
        txtProxySetting(6).Text = X(12)
        chkConfig(5).Value = X(13)
        chkConfig_Click 2
        chkConfig_Click 3
        chkConfig_Click 4
        Me.Show vbModal, ParentForm
    End Select
End Sub

Private Sub SetFrame(objForm As Form, objFrame As Frame, Flag As Boolean)
Dim objControl As Control

    objFrame.Enabled = Flag
    For Each objControl In objForm
        If objControl.Container.Name = objFrame.Name Then
            objControl.Enabled = Flag
            Select Case LCase(TypeName(objControl))
            Case "label"
            Case "textbox"
                If Flag Then
                    objControl.BackColor = vbWhite
                Else
                    objControl.BackColor = vbTransparant
                End If
            Case "frame"
                If Not Flag Then SetFrame Me, objControl, Flag
            Case "commandbutton"
            Case "combobox"
            Case "listbox"
            Case "checkbox"
            End Select
        End If
    Next objControl
End Sub

Private Sub chkConfig_Click(Index As Integer)
    If Index > 1 And Index < 5 Then
        If chkConfig(Index) = vbChecked Then
            txtProxySetting(Index + 2).Enabled = True
            txtProxySetting(Index + 2).BackColor = vbWhite
        Else
            txtProxySetting(Index + 2).Enabled = False
            txtProxySetting(Index + 2).BackColor = vbTransparant
        End If
    ElseIf Index = 5 Then
        If chkConfig(Index) = vbChecked Then
            cmdUserList.Enabled = True
        Else
            cmdUserList.Enabled = False
        End If
    End If
End Sub

Private Sub chkUseProxy_Click()
    If chkUseProxy.Value = vbChecked Then
        SetFrame Me, fraProxy, True
    Else
        SetFrame Me, fraProxy, False
    End If
End Sub

Private Sub cmdCancel_Click()
    ProxySet = False
    ChildOpen = False
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    If Val(txtProxySetting(6).Text) <> Val(txtProxySetting(6).Tag) Then MsgBox "Changing Listening Port required proxy server to restart", vbExclamation
    If txtProxySetting(3).Text = "<hidden>" Then txtProxySetting(3).Text = ""
    ProxySet = True
    ChildOpen = False
    Me.Hide
End Sub

Private Sub cmdUserList_Click()
    Me.Enabled = False
    UserOpen = True
    frmUser.Show vbModal
    Do While UserOpen
        DoEvents
    Loop
    Unload frmUser
    Me.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    ElseIf KeyAscii = vbKeyEscape Then
        ProxySet = False
        ChildOpen = False
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    SetFrame Me, fraProxy, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProxySet = False
    ChildOpen = False
End Sub

Private Sub txtProxySetting_LostFocus(Index As Integer)
    If Index = 1 Or Index = 5 Or Index = 6 Then
        txtProxySetting(Index).Text = Abs(Val(txtProxySetting(Index).Text))
    End If
End Sub
