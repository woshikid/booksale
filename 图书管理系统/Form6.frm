VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "批发商"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9705
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   9705
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3600
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      RecordSource    =   "customer"
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
      Bindings        =   "Form6.frx":7D42
      Height          =   4095
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7223
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "客户号"
         Caption         =   "客户号"
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
         DataField       =   "客户名称"
         Caption         =   "客户名称"
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
         DataField       =   "所在地"
         Caption         =   "所在地"
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
         DataField       =   "邮政编码"
         Caption         =   "邮政编码"
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
         DataField       =   "联系人"
         Caption         =   "联系人"
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
         DataField       =   "联系人手机"
         Caption         =   "联系人手机"
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
         DataField       =   "开户银行"
         Caption         =   "开户银行"
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
         DataField       =   "银行账号"
         Caption         =   "银行账号"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "批发商资料"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton Command3 
         Caption         =   "删除"
         Height          =   375
         Left            =   6480
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改"
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "增加"
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         DataField       =   "银行账号"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         DataField       =   "开户银行"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         DataField       =   "联系人手机"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         DataField       =   "联系人"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         DataField       =   "邮政编码"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text3 
         DataField       =   "所在地"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         DataField       =   "客户名称"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         DataField       =   "客户号"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "银行帐号:"
         Height          =   255
         Left            =   6960
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "开户银行:"
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "联系人手机:"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "联系人:"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "邮政编码:"
         Height          =   255
         Left            =   6960
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "所在地:"
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "客户名称:"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "客户号:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form6"
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

Private Sub Text1_Change()
    Text1.Text = Val(Text1.Text)
End Sub

Private Sub Text4_Change()
    Text4.Text = Val(Text4.Text)
End Sub

Private Sub Text6_Change()
    Text6.Text = Val(Text6.Text)
End Sub

Private Sub Text8_Change()
    Text8.Text = Val(Text8.Text)
End Sub
