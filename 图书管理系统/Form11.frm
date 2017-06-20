VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5265
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "5/28/2006"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "作者:Cocaine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "图书管理系统1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 260
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Sub Form_Load()
    If CBool(PathFileExists(SystemDir() & "\bookreg.ini")) = False Then
        Label4.Caption = "未注册"
    Else
        Dim name As String
        Open SystemDir() & "\bookreg.ini" For Input As #1
            Input #1, name
        Close #1
        Label4.Caption = "注册给：" & name
    End If
End Sub

Public Function SystemDir() As String
    Dim tmp As String
    tmp = Space$(MAX_PATH)
    SystemDir = Left$(tmp, GetSystemDirectory(tmp, MAX_PATH))
End Function
