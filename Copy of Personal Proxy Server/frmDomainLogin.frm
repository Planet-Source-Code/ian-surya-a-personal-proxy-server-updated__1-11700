VERSION 5.00
Begin VB.Form frmDomainLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trusted Domain Login"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
Attribute VB_Name = "frmDomainLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ProxySet = True
    ChildOpen = False
    Me.Hide
End Sub

Private Sub Command2_Click()
    ProxySet = False
    ChildOpen = False
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    ElseIf KeyAscii = vbKeyEscape Then
        Command2_Click
    End If
End Sub

