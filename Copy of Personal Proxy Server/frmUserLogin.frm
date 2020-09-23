VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Domain Login"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmUserLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1350
      Width           =   1005
   End
   Begin VB.TextBox txtDomain 
      Height          =   345
      Left            =   1890
      TabIndex        =   5
      Top             =   810
      Width           =   2715
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   420
      Width           =   2715
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1890
      TabIndex        =   1
      Top             =   30
      Width           =   2715
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2610
      TabIndex        =   6
      Top             =   1350
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Domain"
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   450
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1785
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub OpenCommunication(ParentForm As Form, commtype As Byte, ParamArray X())
    Select Case commtype
    Case 0
        txtUserName.Text = X(0)
        txtPassword.Text = X(1)
        txtDomain.Text = X(2)
        txtDomain.Enabled = False
        txtDomain.BackColor = vbTransparant
        Me.Show vbModal
    Case 1
        txtUserName.Text = X(0)
        txtPassword.Text = X(1)
        txtDomain.Text = X(2)
        txtUserName.Enabled = False
        txtUserName.BackColor = vbTransparant
        txtDomain.Enabled = False
        txtDomain.BackColor = vbTransparant
        Me.Show vbModal
    End Select
End Sub

Private Sub Command1_Click()
    LoginSet = True
    LoginOpen = False
    Me.Hide
End Sub

Private Sub Command2_Click()
    LoginSet = False
    LoginOpen = False
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    ElseIf KeyAscii = vbKeyEscape Then
        LoginSet = False
        LoginOpen = False
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LoginSet = False
    LoginOpen = False
End Sub

Private Sub txtUserName_LostFocus()
    txtUserName = LCase(txtUserName)
End Sub
