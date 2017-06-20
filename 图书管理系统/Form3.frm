VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "日常销售"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6855
   FillColor       =   &H8000000F&
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   6855
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2040
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=shujuku.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=shujuku.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from booksaled order by 出售日期 desc"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":7D42
      Height          =   3975
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "书号"
         Caption         =   "书号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "出售日期"
         Caption         =   "出售日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "售出价格"
         Caption         =   "售出价格"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "数量"
         Caption         =   "数量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1244.976
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "销售"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6375
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         Default         =   -1  'True
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "批发"
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "会员"
         Height          =   255
         Left            =   4680
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "普通"
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "会员号/客户号:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "数量："
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "书号："
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "销售记录："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = Trim(Text1.Text)
    Text2.Text = Trim(Text2.Text)
    Text3.Text = Trim(Text3.Text)
    
    If Text1.Text = "" Then
        MsgBox "请输入书号", , "提示"
        Exit Sub
    End If
    
    If Text2.Text = 0 Then
        MsgBox "购买数量为0，请检查", , "提示"
        Exit Sub
    End If
    
    If Option1.Value = False And (Text3.Text = "" Or Text3.Text = "0") Then
        MsgBox "请输入会员号或客户号", , "提示"
        Exit Sub
    End If
    
    Dim db As ADODB.Connection
    Dim ADOset As ADODB.Recordset
    Dim price As ADODB.Recordset
    
    Set db = New ADODB.Connection
    Set ADOset = New ADODB.Recordset
    Set price = New ADODB.Recordset
    
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=shujuku.mdb;Persist Security Info=False"
    ADOset.Open "select * from book where 书号='" & Text1.Text & "'", db, adOpenStatic, adLockOptimistic
    
    If ADOset.RecordCount <> 1 Then
        MsgBox "没有该种类的书，请检查书号", , "错误"
        Exit Sub
    End If
    
    price.Open "select * from leibie where 类别号=" & ADOset("所属类别"), db, adOpenStatic, adLockOptimistic
    
    If Option1.Value = True Then
        If MsgBox("请确认购买信息:" & ADOset(0) & " " & ADOset(1) & " " & ADOset(2) & " " & ADOset(3) & " 数量:" & Text2.Text & " 普通顾客 价格:" & ADOset("零售价") * Text2.Text, vbOKCancel + vbQuestion, "确认") = vbCancel Then
            Exit Sub
        End If
        
        If Val(ADOset("库存量")) < Val(Text2.Text) Then
            MsgBox "没有足够的库存,剩余数量:" & ADOset("库存量"), , "库存不足"
            Exit Sub
        End If
        
        sql = "insert into booksaled values('" & ADOset(0) & "','" & Date & "'," & ADOset("零售价") * Text2.Text & "," & Text2.Text & ")"
        db.Execute sql
        sql = "update book set 库存量=库存量-" & Text2.Text & " where 书号='" & Text1.Text & "'"
        db.Execute sql
    ElseIf Option2.Value = True Then
        If MsgBox("请确认购买信息:" & ADOset(0) & " " & ADOset(1) & " " & ADOset(2) & " " & ADOset(3) & " 数量:" & Text2.Text & " 会员 价格:" & (Int(ADOset("零售价") * price("会员折扣") * 100)) / 100 * Text2.Text, vbOKCancel + vbQuestion, "确认") = vbCancel Then
            Exit Sub
        End If
        
        If Val(ADOset("库存量")) < Val(Text2.Text) Then
            MsgBox "没有足够的库存,剩余数量:" & ADOset("库存量"), , "库存不足"
            Exit Sub
        End If
        
        sql = "insert into booksaled values('" & ADOset(0) & "','" & Date & "'," & (Int(ADOset("零售价") * price("会员折扣") * 100)) / 100 * Text2.Text & "," & Text2.Text & ")"
        db.Execute sql
        sql = "insert into huiyuansale values(" & Text3.Text & ",'" & ADOset(0) & "','" & Date & "'," & (Int(ADOset("零售价") * price("会员折扣") * 100)) / 100 * Text2.Text & "," & Text2.Text & ")"
        db.Execute sql
        sql = "update book set 库存量=库存量-" & Text2.Text & " where 书号='" & Text1.Text & "'"
        db.Execute sql
    Else
        If MsgBox("请确认购买信息:" & ADOset(0) & " " & ADOset(1) & " " & ADOset(2) & " " & ADOset(3) & " 数量:" & Text2.Text & " 批发 价格:" & (Int(ADOset("零售价") * price("批发折扣") * 100)) / 100 * Text2.Text, vbOKCancel + vbQuestion, "确认") = vbCancel Then
            Exit Sub
        End If
        
        If Val(ADOset("库存量")) < Val(Text2.Text) Then
            MsgBox "没有足够的库存,剩余数量:" & ADOset("库存量"), , "库存不足"
            Exit Sub
        End If
        
        sql = "insert into booksaled values('" & ADOset(0) & "','" & Date & "'," & (Int(ADOset("零售价") * price("会员折扣") * 100)) / 100 * Text2.Text & "," & Text2.Text & ")"
        db.Execute sql
        sql = "insert into customersale values(" & Text3.Text & ",'" & ADOset(0) & "','" & Date & "'," & (Int(ADOset("零售价") * price("会员折扣") * 100)) / 100 * Text2.Text & "," & Text2.Text & ")"
        db.Execute sql
        sql = "update book set 库存量=库存量-" & Text2.Text & " where 书号='" & Text1.Text & "'"
        db.Execute sql
    End If
        
    Text1.Text = ""
    Text2.Text = 0
    Text3.Text = ""
    Option1.Value = True
    Adodc1.Refresh
End Sub

Private Sub Option1_Click()
    Text3.Enabled = False
End Sub

Private Sub Option2_Click()
    Text3.Enabled = True
End Sub

Private Sub Option3_Click()
    Text3.Enabled = True
End Sub

Private Sub Text2_Change()
    Text2.Text = Val(Text2.Text)
End Sub

Private Sub Text3_Change()
    Text3.Text = Val(Text3.Text)
End Sub
