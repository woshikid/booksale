VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "顾客图书查询"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9840
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "查询结果"
      Height          =   4815
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   9375
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form2.frx":7D42
         Height          =   4455
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
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
         ColumnCount     =   10
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
            DataField       =   "书名"
            Caption         =   "书名"
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
            DataField       =   "作者"
            Caption         =   "作者"
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
            DataField       =   "出版社"
            Caption         =   "出版社"
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
         BeginProperty Column04 
            DataField       =   "出版日期"
            Caption         =   "出版日期"
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
         BeginProperty Column05 
            DataField       =   "所在架位"
            Caption         =   "所在架位"
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
         BeginProperty Column06 
            DataField       =   "库存量"
            Caption         =   "库存量"
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
         BeginProperty Column07 
            DataField       =   "类别名"
            Caption         =   "类别名"
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
         BeginProperty Column08 
            DataField       =   "零售价"
            Caption         =   "零售价"
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
         BeginProperty Column09 
            DataField       =   "会员价"
            Caption         =   "会员价"
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
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2129.953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1560.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   764.787
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3960
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      RecordSource    =   $"Form2.frx":7D57
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
   Begin VB.Frame Frame1 
      Caption         =   "智能查询"
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   9375
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Default         =   -1  'True
         Height          =   375
         Left            =   7680
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "模糊查找"
         Height          =   255
         Left            =   6000
         TabIndex        =   9
         Top             =   960
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "精确查找"
         Height          =   255
         Left            =   4560
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7440
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "书号："
         Height          =   255
         Left            =   6840
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "出版社："
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "作者："
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "书名："
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "请输入您的查询条件："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = Trim(Text1.Text) '去掉条件的前后空格
    Text2.Text = Trim(Text2.Text)
    Text3.Text = Trim(Text3.Text)
    Text4.Text = Trim(Text4.Text)
    
    If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" Then
        MsgBox "请输入查询条件", , "查询条件"
        Exit Sub
    End If
    
    Dim sql As String '查询的sql语句
    sql = "select 书号,书名,作者,出版社,出版日期,所在架位,库存量,类别名,零售价,零售价 * 会员折扣 as 会员价 from book,leibie where 类别号=所属类别"
    
    If Option1.Value = True Then '精确查找
        If Text1.Text <> "" Then
            sql = sql & " and 书名='" & Text1.Text & "'"
        End If
        If Text2.Text <> "" Then
            sql = sql & " and 作者='" & Text2.Text & "'"
        End If
        If Text3.Text <> "" Then
            sql = sql & " and 书号='" & Text3.Text & "'"
        End If
        If Text4.Text <> "" Then
            sql = sql & " and 出版社='" & Text4.Text & "'"
        End If
    Else
        If Text1.Text <> "" Then
            sql = sql & " and 书名 like '%" & Text1.Text & "%'"
        End If
        If Text2.Text <> "" Then
            sql = sql & " and 作者 like '%" & Text2.Text & "%'"
        End If
        If Text3.Text <> "" Then
            sql = sql & " and 书号 like '%" & Text3.Text & "&'"
        End If
        If Text4.Text <> "" Then
            sql = sql & " and 出版社 like '%" & Text4.Text & "%'"
        End If
    End If
    
    Adodc1.RecordSource = sql
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "没有找到相关的记录", , "查询结果"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

