VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "添加新用户"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3975
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   3975
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "确认密码:"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "密码:"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "用户名:"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
        MsgBox "请输入完整的注册信息", , "错误"
    ElseIf Text1.Text = "" Then
        MsgBox "必须输入用户名", , "错误"
    ElseIf Text2.Text = "" Then
        MsgBox "请输入密码", , "错误"
    ElseIf Text2.Text <> Text3.Text Then
        MsgBox "你两次输入的密码不同", , "错误"
        Text2.Text = ""
        Text3.Text = ""
        Text2.SetFocus
    Else
        On Error Resume Next
        Dim db As ADODB.Connection
        Set db = New ADODB.Connection
    
        db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=shujuku.mdb;Persist Security Info=False"
        db.Execute "insert into yonghuming values('" & Tex1.Text & "','" & Text2.Text & "')"
        If Err Then
            MsgBox Err.Description
            Exit Sub
        End If
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Text1_Change()
    Text1.Text = Trim(Text1.Text)
End Sub
