VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
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
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�޸�����"
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
      Caption         =   "ȷ��������:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "������:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "ԭ����:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�û���:"
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
        MsgBox "��û��������Ϣ��", , "����"
    ElseIf Text2.Text = "" Then
        MsgBox "������ԭ���룡", , "����"
    ElseIf Text3.Text = "" Then
        MsgBox "��û�����������룡", , "����"
    ElseIf Text4.Text = "" Then
        MsgBox "��ȷ����������룡", , "����"
    ElseIf Text3.Text <> Text4.Text Then
        MsgBox "��������������벻ͬ�����������������룡", , "����"
        Text3.Text = ""
        Text4.Text = ""
        Text3.SetFocus
    Else
        Dim db As ADODB.Connection
        Dim ADOset As ADODB.Recordset
    
        Set db = New ADODB.Connection
        Set ADOset = New ADODB.Recordset
    
        db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=shujuku.mdb;Persist Security Info=False"
        ADOset.Open "select * from yonghuming where �û���='" & Label5.Caption & "' and ����='" & Text2.Text & "'", db, adOpenStatic, adLockOptimistic
    
        If ADOset.EOF Then
            MsgBox "ԭ�������,����������", , "����"
            Exit Sub
        Else
            db.Execute "update yonghuming set ����='" & Text3.Text & "' where �û���='" & Label5.Caption & "'"
            MsgBox "���Ѿ��ɹ��ĸ����������룬��ʹ�������룡", , "�ɹ�"
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
