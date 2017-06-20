VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登陆"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox Text2 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "用户名:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "密码:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tcount As Integer

Private Sub Command1_Click()
    If Text1.Text = "" And Text2.Text = "" Then
        MsgBox "请输入用户名和密码！", , "错误"
    ElseIf Text1.Text = "" Then
        MsgBox "你还没有输入用户名！", , "错误"
        Text1.SetFocus
    ElseIf Text2.Text = "" Then
        MsgBox "你还没有输入密码！", , "错误"
        Text2.SetFocus
    Else
        Dim user As String
        user = "用户名=" & "'" & Text1.Text & "' and " & "密码=" & "'" & Text2.Text & "'"
        
        Dim db As ADODB.Connection
        Dim ADOset As ADODB.Recordset
        
        Set db = New ADODB.Connection
        Set ADOset = New ADODB.Recordset
        
        db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=shujuku.mdb;Persist Security Info=False"
        ADOset.Open "select * from yonghuming where " & user, db, adOpenStatic, adLockOptimistic
        
        If ADOset.EOF Then
            tcount = tcount + 1
            If tcount > 2 Then
                MsgBox "非法用户,不能使用本系统", , "三次尝试已过"
                End
            End If
            MsgBox "没有该用户，或密码错误！", , "错误"
            Text2.Text = ""
            Exit Sub
        Else
            MDIForm1.Show
            MDIForm1.user = Text1.Text
            Unload Me
        End If
    End If
End Sub

Private Sub Command2_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub Form_Load()
    tcount = 0
End Sub

Private Sub Text1_Change()
    Text1.Text = Trim(Text1.Text)
End Sub
