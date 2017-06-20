VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "管理员"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   9885
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu richangyewu 
      Caption         =   "日常业务"
      Begin VB.Menu richangxiaoshou 
         Caption         =   "日常销售"
      End
   End
   Begin VB.Menu jibenziliao 
      Caption         =   "基本资料"
      Begin VB.Menu tushujibenziliao 
         Caption         =   "图书基本资料"
      End
      Begin VB.Menu kehujibenziliao 
         Caption         =   "客户基本资料"
         Begin VB.Menu huiyuan 
            Caption         =   "会员"
         End
         Begin VB.Menu pifashang 
            Caption         =   "批发商"
         End
      End
   End
   Begin VB.Menu xitonggongneng 
      Caption         =   "系统功能"
      Begin VB.Menu mimashezhi 
         Caption         =   "密码设置"
      End
      Begin VB.Menu yonghuguanli 
         Caption         =   "用户管理"
      End
      Begin VB.Menu tushuzhekoushezhi 
         Caption         =   "图书折扣设置"
      End
      Begin VB.Menu pandian 
         Caption         =   "盘点"
      End
   End
   Begin VB.Menu tongjifenxi 
      Caption         =   "统计分析"
      Begin VB.Menu xiaoshoufenxi 
         Caption         =   "销售分析"
      End
      Begin VB.Menu xiaoshoutongji 
         Caption         =   "销售统计"
         Begin VB.Menu ribaobiao 
            Caption         =   "日报表"
         End
         Begin VB.Menu yuebaobiao 
            Caption         =   "月报表"
         End
         Begin VB.Menu nianbaobiao 
            Caption         =   "年报表"
         End
      End
      Begin VB.Menu kucunfenxi 
         Caption         =   "库存分析"
      End
   End
   Begin VB.Menu bangzhu 
      Caption         =   "帮助"
      Begin VB.Menu guanyu 
         Caption         =   "关于"
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
    DataEnvironment1.Commands(3).CommandText = "select 书号,sum(售出价格) as 总销售额,sum(数量) as 总销售量 from booksaled where 出售日期 >#" & (Date - 356) & "# group by 书号 order by sum(数量) desc"
    DataReport3.Show
End Sub

Private Sub pandian_Click()
    Form10.Show
End Sub

Private Sub pifashang_Click()
    Form6.Show
End Sub

Private Sub ribaobiao_Click()
    DataEnvironment1.Commands(1).CommandText = "select 书号,sum(售出价格) as 总销售额,sum(数量) as 总销售量 from booksaled where 出售日期 =#" & Date & "# group by 书号 order by sum(数量) desc"
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
    DataEnvironment1.Commands(2).CommandText = "select 书号,sum(售出价格) as 总销售额,sum(数量) as 总销售量 from booksaled where 出售日期 >#" & (Date - 30) & "# group by 书号 order by sum(数量) desc"
    DataReport2.Show
End Sub

Public Function SystemDir() As String
    Dim tmp As String
    tmp = Space$(MAX_PATH)
    SystemDir = Left$(tmp, GetSystemDirectory(tmp, MAX_PATH))
End Function
