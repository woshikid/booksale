VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "图书基本资料"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   7815
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":7D42
      Height          =   3255
      Left            =   120
      TabIndex        =   26
      Top             =   2640
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5741
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
      ColumnCount     =   11
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
         DataField       =   "所属类别"
         Caption         =   "所属类别"
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
      BeginProperty Column08 
         DataField       =   "进货价"
         Caption         =   "进货价"
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
      BeginProperty Column10 
         DataField       =   "截至日期"
         Caption         =   "截至日期"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3600
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "book"
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
      Caption         =   "图书基本资料"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton Command3 
         Caption         =   "删除"
         Height          =   375
         Left            =   5760
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改"
         Height          =   375
         Left            =   4320
         TabIndex        =   24
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "增加"
         Height          =   375
         Left            =   2880
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         DataField       =   "截至日期"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         DataField       =   "零售价"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text9 
         DataField       =   "进货价"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text8 
         DataField       =   "库存量"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text7 
         DataField       =   "所属类别"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text6 
         DataField       =   "所在架位"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         DataField       =   "出版日期"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         DataField       =   "出版社"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         DataField       =   "作者"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         DataField       =   "书名"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "书号"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "截至日期:"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "零售价:"
         Height          =   255
         Left            =   5400
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "进货价:"
         Height          =   375
         Left            =   3600
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "库存量:"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "所属类别:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "所在架位:"
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "出版日期:"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "出版社:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "作者:"
         Height          =   255
         Left            =   5280
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "书名:"
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "书号:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    If Command1.Caption = "增加" Then
        Adodc1.Recordset.AddNew
        Text1.Locked = False
        Text2.Locked = False
        Text3.Locked = False
        Text4.Locked = False
        Text5.Locked = False
        Text6.Locked = False
        Text7.Locked = False
        Text8.Locked = False
        Text9.Locked = False
        Text10.Locked = False
        Text11.Locked = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command1.Caption = "保存"
    Else
        Text1.Text = Trim(Text1.Text)
        Text2.Text = Trim(Text2.Text)
        Text3.Text = Trim(Text3.Text)
        Text4.Text = Trim(Text4.Text)
        Text5.Text = Trim(Text5.Text)
        Text6.Text = Trim(Text6.Text)
        Text7.Text = Trim(Text7.Text)
        Text8.Text = Trim(Text8.Text)
        Text9.Text = Trim(Text9.Text)
        Text10.Text = Trim(Text10.Text)
        Text11.Text = Trim(Text11.Text)
        
        Adodc1.Recordset.Update
        If Err Then
            MsgBox Err.Description, , "请仔细检查数据"
        End If
        Text1.Locked = True
        Text2.Locked = True
        Text3.Locked = True
        Text4.Locked = True
        Text5.Locked = True
        Text6.Locked = True
        Text7.Locked = True
        Text8.Locked = True
        Text9.Locked = True
        Text10.Locked = True
        Text11.Locked = True
        Command2.Enabled = True
        Command3.Enabled = True
        Command1.Caption = "增加"
    End If
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    If Command2.Caption = "修改" Then
        Text1.Locked = False
        Text2.Locked = False
        Text3.Locked = False
        Text4.Locked = False
        Text5.Locked = False
        Text6.Locked = False
        Text7.Locked = False
        Text8.Locked = False
        Text9.Locked = False
        Text10.Locked = False
        Text11.Locked = False
        Command1.Enabled = False
        Command3.Enabled = False
        Command2.Caption = "保存"
    Else
        Text1.Text = Trim(Text1.Text)
        Text2.Text = Trim(Text2.Text)
        Text3.Text = Trim(Text3.Text)
        Text4.Text = Trim(Text4.Text)
        Text5.Text = Trim(Text5.Text)
        Text6.Text = Trim(Text6.Text)
        Text7.Text = Trim(Text7.Text)
        Text8.Text = Trim(Text8.Text)
        Text9.Text = Trim(Text9.Text)
        Text10.Text = Trim(Text10.Text)
        Text11.Text = Trim(Text11.Text)
        Adodc1.Recordset.Update
        If Err Then
            MsgBox Err.Description, , "请仔细检查数据"
        End If
        Text1.Locked = True
        Text2.Locked = True
        Text3.Locked = True
        Text4.Locked = True
        Text5.Locked = True
        Text6.Locked = True
        Text7.Locked = True
        Text8.Locked = True
        Text9.Locked = True
        Text10.Locked = True
        Text11.Locked = True
        Command1.Enabled = True
        Command3.Enabled = True
        Command2.Caption = "修改"
    End If
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.RecordCount = 0 Then
        Command2.Enabled = False
        Command3.Enabled = False
    End If
    
    If Err Then
        MsgBox Err.Description, , "错误"
    End If
End Sub

Private Sub Form_Load()
    If Adodc1.Recordset.RecordCount = 0 Then
        Command2.Enabled = False
        Command3.Enabled = False
    End If
End Sub

Private Sub Text7_Change()
    Text7.Text = Val(Text7.Text)
End Sub

Private Sub Text8_Change()
    Text7.Text = Val(Text7.Text)
End Sub

