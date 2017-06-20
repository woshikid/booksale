VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "注册"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   Icon            =   "reg.frx":0000
   LinkTopic       =   "Form14"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "试用"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "注册"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "序列号："
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "使用者："
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "您尚未注册，部分功能有所限制，注册请填写如下信息："
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 260
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Sub Command1_Click()
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "" Then
        MsgBox "请输入姓名", , "注册"
        Exit Sub
    End If
    
    If Text2.Text = "0KSB7NP56RQU" Then
        Open SystemDir() & "\bookreg.ini" For Output As #1
            Write #1, Text1.Text
        Close #1
        MsgBox "注册成功", , "注册"
        Form1.Show
        Unload Me
    Else
        MsgBox "序列号错误", , "注册"
    End If
End Sub

Private Sub Command2_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Dim name As String
    If CBool(PathFileExists(SystemDir() & "\bookreg.ini")) = False Then
        Me.Visible = True
    Else
        Form1.Show
        Unload Me
    End If
End Sub
Public Function SystemDir() As String
    Dim tmp As String
    tmp = Space$(MAX_PATH)
    SystemDir = Left$(tmp, GetSystemDirectory(tmp, MAX_PATH))
End Function
