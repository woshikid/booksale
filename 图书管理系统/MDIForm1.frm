VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "����Ա"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   9885
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu richangyewu 
      Caption         =   "�ճ�ҵ��"
      Begin VB.Menu richangxiaoshou 
         Caption         =   "�ճ�����"
      End
   End
   Begin VB.Menu jibenziliao 
      Caption         =   "��������"
      Begin VB.Menu tushujibenziliao 
         Caption         =   "ͼ���������"
      End
      Begin VB.Menu kehujibenziliao 
         Caption         =   "�ͻ���������"
         Begin VB.Menu huiyuan 
            Caption         =   "��Ա"
         End
         Begin VB.Menu pifashang 
            Caption         =   "������"
         End
      End
   End
   Begin VB.Menu xitonggongneng 
      Caption         =   "ϵͳ����"
      Begin VB.Menu mimashezhi 
         Caption         =   "��������"
      End
      Begin VB.Menu yonghuguanli 
         Caption         =   "�û�����"
      End
      Begin VB.Menu tushuzhekoushezhi 
         Caption         =   "ͼ���ۿ�����"
      End
      Begin VB.Menu pandian 
         Caption         =   "�̵�"
      End
   End
   Begin VB.Menu tongjifenxi 
      Caption         =   "ͳ�Ʒ���"
      Begin VB.Menu xiaoshoufenxi 
         Caption         =   "���۷���"
      End
      Begin VB.Menu xiaoshoutongji 
         Caption         =   "����ͳ��"
         Begin VB.Menu ribaobiao 
            Caption         =   "�ձ���"
         End
         Begin VB.Menu yuebaobiao 
            Caption         =   "�±���"
         End
         Begin VB.Menu nianbaobiao 
            Caption         =   "�걨��"
         End
      End
      Begin VB.Menu kucunfenxi 
         Caption         =   "������"
      End
   End
   Begin VB.Menu bangzhu 
      Caption         =   "����"
      Begin VB.Menu guanyu 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 260
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public user As String

Private Sub guanyu_Click()
    Form11.Show
End Sub

Private Sub huiyuan_Click()
    Form5.Show
End Sub

Private Sub kucunfenxi_Click()
    Form12.Show
End Sub

Private Sub MDIForm_Load()
    If CBool(PathFileExists(SystemDir() & "\bookreg.ini")) = False Then
        xiaoshoufenxi.Enabled = False
        xiaoshoutongji.Enabled = False
        kucunfenxi.Enabled = False
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub mimashezhi_Click()
    Form7.Show
End Sub

Private Sub nianbaobiao_Click()
    DataEnvironment1.Commands(3).CommandText = "select ���,sum(�۳��۸�) as �����۶�,sum(����) as �������� from booksaled where �������� >#" & (Date - 356) & "# group by ��� order by sum(����) desc"
    DataReport3.Show
End Sub

Private Sub pandian_Click()
    Form10.Show
End Sub

Private Sub pifashang_Click()
    Form6.Show
End Sub

Private Sub ribaobiao_Click()
    DataEnvironment1.Commands(1).CommandText = "select ���,sum(�۳��۸�) as �����۶�,sum(����) as �������� from booksaled where �������� =#" & Date & "# group by ��� order by sum(����) desc"
    DataReport1.Show
End Sub

Private Sub richangxiaoshou_Click()
    Form3.Show
End Sub

Private Sub tushujibenziliao_Click()
    Form4.Show
End Sub

Private Sub tushuzhekoushezhi_Click()
    Form9.Show
End Sub

Private Sub xiaoshoufenxi_Click()
    Form13.Show
End Sub

Private Sub yonghuguanli_Click()
    Form8.Show
End Sub

Private Sub yuebaobiao_Click()
    DataEnvironment1.Commands(2).CommandText = "select ���,sum(�۳��۸�) as �����۶�,sum(����) as �������� from booksaled where �������� >#" & (Date - 30) & "# group by ��� order by sum(����) desc"
    DataReport2.Show
End Sub

Public Function SystemDir() As String
    Dim tmp As String
    tmp = Space$(MAX_PATH)
    SystemDir = Left$(tmp, GetSystemDirectory(tmp, MAX_PATH))
End Function
