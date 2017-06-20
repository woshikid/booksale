VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "密码设置"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3390
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   3390
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改密码"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "确认新密码:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "新密码:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "原密码:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "用户名:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text2.Text = "" And Text3.Text = "" And Text4.Text = "" Then
        MsgBox "你没有输入信息！", , "错误"
    ElseIf Text2.Text = "" Then
        MsgBox "请输入原密码！", , "错误"
    ElseIf Text3.Text = "" Then
        MsgBox "你没有输入新密码！", , "错误"
    ElseIf Text4.Text = "" Then
        MsgBox "请确认你的新密码！", , "错误"
    ElseIf Text3.Text <> Text4.Text Then
        MsgBox "你两次输入的密码不同，请重新输入新密码！", , "错误"
        Text3.Text = ""
        Text4.Text = ""
        Text3.SetFocus
    Else
        Dim db As ADODB.Connection
        Dim ADOset As ADODB.Recordset
    
        Set db = New ADODB.Connection
        Set ADOset = New ADODB.Recordset
    
        db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=shujuku.mdb;Persist Security Info=False"
        ADOset.Open "select * from yonghuming where 用户名='" & Label5.Caption & "' and 密码='" & Text2.Text & "'", db, adOpenStatic, adLockOptimistic
    
        If ADOset.EOF Then
            MsgBox "原密码错误,请重新输入", , "错误"
            Exit Sub
        Else
            db.Execute "update yonghuming set 密码='" & Text3.Text & "' where 用户名='" & Label5.Caption & "'"
            MsgBox "你已经成功的更换了新密码，请使用新密码！", , "成功"
            Unload Me
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label5.Caption = MDIForm1.user
End Sub
