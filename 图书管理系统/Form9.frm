VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "图书折扣设置"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7170
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   7170
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1800
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
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
      RecordSource    =   "leibie"
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
      Bindings        =   "Form9.frx":7D42
      Height          =   1815
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3201
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
         DataField       =   "类别号"
         Caption         =   "类别号"
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
      BeginProperty Column02 
         DataField       =   "会员折扣"
         Caption         =   "会员折扣"
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
         DataField       =   "批发折扣"
         Caption         =   "批发折扣"
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
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1395.213
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "图书折扣"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "修改折扣"
         Default         =   -1  'True
         Height          =   735
         Left            =   5280
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         DataField       =   "批发折扣"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text3 
         DataField       =   "会员折扣"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         DataField       =   "类别名"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         DataField       =   "类别号"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "批发折扣:"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "会员折扣:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "类别名:"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "类别号:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Command1.Caption = "修改折扣" Then
        Text3.Locked = False
        Text4.Locked = False
        Command1.Caption = "保存"
    Else
        If Val(Text3.Text) > 1 Or Val(Text3.Text) <= 0 Or Val(Text4.Text) > 1 Or Val(Text4.Text) <= 0 Then
            MsgBox "会员折扣和批发折扣数值应为:0<折扣百分比<=1例:0.8", , "错误"
            Exit Sub
        End If
        Adodc1.Recordset.Update
        
        If Err Then
            MsgBox Err.Description, , "请仔细检查数据"
        End If
        Text3.Locked = True
        Text4.Locked = True
        Command1.Caption = "修改折扣"
    End If
End Sub

Private Sub Text3_Change()
    Text3.Text = Trim(Text3.Text)
End Sub

Private Sub Text4_Change()
    Text4.Text = Trim(Text4.Text)
End Sub
