VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUser 
   Caption         =   "User List"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxUser 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   7223
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
   End
   Begin VB.Menu mnuUser 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub flxUser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuUser, , X, Y
    End If
End Sub

Private Sub Form_Load()
    InitializeGrid
    FillGrid
End Sub

Private Sub Form_Resize()
    flxUser.Width = Me.Width
    flxUser.Height = Me.Height
    ResizeGrid
End Sub

Private Sub InitializeGrid()
    With flxUser
        .Clear
        .Rows = 1
        .Cols = 2
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "User"
        .ColAlignment(0) = flexAlignLeftCenter
    End With
End Sub

Private Sub ResizeGrid()
    With flxUser
        .ColWidth(0) = 300 / 2300 * (flxUser.Width - 180)
        .ColWidth(1) = 2000 / 2300 * (flxUser.Width - 180)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UserOpen = False
    SaveUser UserList, "UserList.txt"
End Sub

Private Sub FillGrid()
Dim i As Long
    With flxUser
        For i = 1 To UserList.Count
            AddGrid i, GetUser(UserList(i).Key)
        Next i
    End With
End Sub

Private Sub AddGrid(Row As Long, UserName As String)
    With flxUser
        If .Rows = Row Then .Rows = .Rows + 1
        .TextMatrix(Row, 0) = Row
        .TextMatrix(Row, 1) = UserName
    End With
End Sub

Private Sub mnuAdd_Click()
Dim UserName As String, Password As String
    Me.Enabled = False
    LoginOpen = True
    LoginSet = False
    UserName = ""
    Password = ""
    frmUserLogin.OpenCommunication Me, 0, "", "", "Personal Proxy Server"
    Do While LoginOpen
        DoEvents
    Loop
    UserName = Trim(frmUserLogin.txtUserName)
    If UserName <> "" And LoginSet Then
        If frmUserLogin.txtPassword <> "<hidden>" Then Password = Trim(frmUserLogin.txtPassword)
        If CBool(IsUserExist(UserName)) Then
            If MsgBox("Do you want to update this user ?", vbYesNo) = vbYes Then
                UserList.Remove IsUserExist(UserName)
                AddUser UserList, Base64Encode(UserName & ":" & Password)
                InitializeGrid
                FillGrid
            End If
        Else
            AddUser UserList, Base64Encode(UserName & ":" & Password)
            InitializeGrid
            FillGrid
        End If
    End If
    Unload frmUserLogin
    Me.Enabled = True
End Sub

Private Sub mnuDelete_Click()
    If flxUser.RowSel <> 0 Then
        If MsgBox("Delete this user ?", vbYesNo) = vbYes Then
            UserList.Remove flxUser.RowSel
            InitializeGrid
            FillGrid
        End If
    End If
End Sub

Private Sub mnuEdit_Click()
Dim UserName As String, Password As String
    If flxUser.RowSel <> 0 Then
        Me.Enabled = False
        LoginOpen = True
        LoginSet = False
        UserName = GetUser(UserList(flxUser.RowSel).Key)
        Password = GetPassword(UserList(flxUser.RowSel).Key)
        frmUserLogin.OpenCommunication Me, 1, UserName, "<hidden>", "Personal Proxy Server"
        Do While LoginOpen
            DoEvents
        Loop
        If LoginSet Then
            If frmUserLogin.txtPassword <> "<hidden>" Then Password = Trim(frmUserLogin.txtPassword)
            UserList.Remove flxUser.RowSel
            AddUser UserList, Base64Encode(UserName & ":" & Password)
            InitializeGrid
            FillGrid
        End If
        Unload frmUserLogin
        Me.Enabled = True
    End If
End Sub
