VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Forml603 
   BackColor       =   &H00C0E0FF&
   Caption         =   "外加工出库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data8 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data7 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data6 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   720
      TabIndex        =   43
      Top             =   1560
      Width           =   2775
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "进行"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   13080
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text1111 
      Height          =   270
      Left            =   5520
      TabIndex        =   37
      Text            =   "Text4"
      Top             =   6720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Forml603.frx":0000
      Height          =   3615
      Left            =   4440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5760
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   10200
      TabIndex        =   34
      Text            =   "Text2"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5760
      TabIndex        =   32
      Text            =   "Text2"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   12240
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3720
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   9720
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   8040
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   6360
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   10800
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   9240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9240
      TabIndex        =   15
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80674817
      CurrentDate     =   39557
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Forml603.frx":0014
      Height          =   360
      Left            =   6360
      TabIndex        =   29
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Forml603.frx":0028
      Height          =   2415
      Left            =   4440
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   11
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Forml603.frx":003C
      Left            =   8040
      List            =   "Forml603.frx":004F
      TabIndex        =   39
      Text            =   "Combo1"
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6375
      Left            =   720
      TabIndex        =   46
      Top             =   3000
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11245
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1920
      TabIndex        =   47
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80674817
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1920
      TabIndex        =   48
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80674817
      CurrentDate     =   39557
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   720
      TabIndex        =   50
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   720
      TabIndex        =   49
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单价"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   13080
      TabIndex        =   41
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "颜色刷新"
      Height          =   375
      Left            =   12240
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   8880
      TabIndex        =   35
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "单号刷新"
      Height          =   375
      Left            =   7800
      TabIndex        =   33
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   4440
      TabIndex        =   31
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "出库数量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   10800
      TabIndex        =   27
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单号"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   26
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5640
      TabIndex        =   25
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7080
      TabIndex        =   24
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   23
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "计划"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   9720
      TabIndex        =   22
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "位置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4440
      TabIndex        =   21
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "外协单位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   6360
      TabIndex        =   20
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   9240
      TabIndex        =   19
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   12240
      TabIndex        =   18
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "类别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8040
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "Forml603"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public l, S1, S2, s3, s4, c, r As Integer


Private Sub Combo1_Click()
Text1(6).Text = Combo1.Text
End Sub


Private Sub Command1_Click()
If Text1(7).Text <> "" Then
Data1.RecordSource = "select * from wxjl where 单位='" & Text1(7).Text & "' and 日期=cdate('" & Text1(9).Text & "') order by 款号,序号"
Data1.Refresh
Else
Data1.RecordSource = "select * from wxjl where 日期=cdate('" & Text1(9).Text & "') order by 款号,序号"
Data1.Refresh
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next


If Text1(5).Text = "" Then
MsgBox ("选择位置")
Exit Sub
End If

If Text1(7).Text = "" Then
MsgBox ("选择单位")
Exit Sub
End If



If Text1(2).Text = "" Or Text1(4).Text = "" Or Text1(6).Text = "" Then
MsgBox ("输入不完整!")
Exit Sub
End If

Data1.Recordset.AddNew
For i = 0 To 11
Data1.Recordset.Fields(i) = Text1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

Data2.Refresh
Text1(10).Text = 1
If Data2.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data2.Recordset.Fields(0) + 1
End If
End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
On Error Resume Next
For i = 1 To 4
Text1(i).Text = ""
Next
For i = 8 To 10
Text1(i).Text = ""
Next

Text1(9).Text = Date
DTPicker1.Value = Date

If Text3.Text = "" Then
Data1.RecordSource = "select * from wxjl where 单号='" & Text2.Text & "'  order by 日期 desc,款号,位置 desc,颜色,规格,序号"
Data1.Refresh
Else
Data1.RecordSource = "select * from wxjl where 单号='" & Text2.Text & "' and 颜色='" & Text3.Text & "' order by 日期 desc,款号,位置 desc,颜色,规格,序号"
Data1.Refresh
End If
Data2.RecordSource = "SELECT MAX(序号) FROM wxjl where 单号='" & Text2.Text & "' and 颜色='" & Text3.Text & "'"
Data2.Refresh

Text1(10).Text = 1
If Data2.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
If MsgBox("确定删除吗？，删除不能回复！", vbYesNo) = vbNo Then Exit Sub

If s3 = 0 Or s4 = 0 Then
MsgBox ("请选择记录！")
Exit Sub
End If

If s3 < 1 Or s4 < 1 Then
MsgBox ("选择记录")
Exit Sub
End If

If s3 > s4 Then
MsgBox ("注意选择顺序！")
Exit Sub
End If

k = s4 - s3

If k = 0 Then
Data4.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data1.Recordset.Move s3 - 1
Data1.Recordset.Delete
Else
Data1.Recordset.MoveFirst
Data1.Recordset.Move s3 - 1
For l = 1 To k + 1
Data1.Recordset.Delete
Data1.Recordset.MoveNext
Next
End If

Data1.Refresh
Data2.Refresh
Text1(10).Text = 1
If Data2.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data2.Recordset.Fields(0) + 1
End If
End Sub


Private Sub Command7_Click()
Call cwbb(MSFlexGrid2, "报表日期" + Text1(9).Text)
End Sub

Private Sub Command8_Click()
Call tree
Call zk
End Sub

Private Sub DBCombo1_Click(Area As Integer)
Text1(7).Text = DBCombo1.Text
End Sub

Private Sub DTPicker1_Change()
On Error Resume Next
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from wxjl where 日期=cdate('" & Text1(9).Text & "') order by 款号,位置 desc,颜色,规格,序号"
Data1.Refresh
Text1(9).Text = DTPicker1.Value

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM wxjl where 日期=cdate('" & Text1(9).Text & "')"
Data2.Refresh

Text1(10).Text = 1
If Data2.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data2.Recordset.Fields(0) + 1
End If


End Sub

Private Sub DTPicker1_CloseUp()
On Error Resume Next
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from wxjl where 日期=cdate('" & Text1(9).Text & "') order by 款号,位置 desc,颜色,规格,序号"
Data1.Refresh
Text1(9).Text = DTPicker1.Value

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM wxjl where 日期=cdate('" & Text1(9).Text & "')"
Data2.Refresh

Text1(10).Text = 1
If Data2.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data2.Recordset.Fields(0) + 1
End If



End Sub

Private Sub Form_Load()
On Error Resume Next
DBCombo2.Text = ""
For i = 0 To 11
Text1(i).Text = ""
Next
Text1(9).Text = Date
Text2.Text = ""
Text3.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date - 30
DTPicker3.Value = Date
DBCombo1.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from wxjl where 单号='" & Text2.Text & "' and 颜色='" & Text3.Text & "' order by 序号 desc"
Data1.Refresh
Combo1.Text = ""
Option4.Value = True
Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM wxjl where 单号='" & Text2.Text & "' and 颜色='" & Text3.Text & "'"
Data2.Refresh

Text1(10).Text = 1
If Data2.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data2.Recordset.Fields(0) + 1
End If

Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data5.RecordSource = "select 简称 from khzl where instr(代码,'外')>0 group by 简称"
Data5.Refresh
Data7.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data8.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data6.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(1) = 0
For i = 2 To 5
MSFlexGrid2.ColWidth(i) = 1000
Next

MSFlexGrid1.ColWidth(0) = 300
For i = 2 To 5
MSFlexGrid1.ColWidth(i) = 1100
Next
MSFlexGrid1.ColWidth(10) = 0

End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 11
khbl = 13
Forml202.Text1.Text = Text2.Text
Forml202.Show
End Select
End Sub

Private Sub Label2_DBLClick()
On Error Resume Next
lo = "d:\数据库\\htgl\2011\scjd.MDB"
Data4.Database.Execute "delete * from wxxx"
Data5.Database.Execute "insert into wxxx(单号,款号,颜色,计划) in'" & lo & "' select 单号,款号,颜色,sum(数量) from SCZY_XDH where instr(单号,'" & Text2.Text & "')>0 group by 单号,款号,颜色"
Data4.Database.Execute "insert into wxxx(单号,款号,颜色,裁剪) select 单号,款号,颜色,sum(val(裁剪)) from cjrb where instr(单号,'" & Text2.Text & "')>0 group by 单号,款号,颜色"
Data4.Database.Execute "insert into wxxx(单号,款号,颜色,外协) select 单号,款号,颜色,sum(val(数量)) from wxjl where instr(单号,'" & Text2.Text & "')>0 group by 单号,款号,颜色"
Data4.Database.Execute "insert into wxxx(单号,款号,颜色,返回) select 单号,款号,颜色,sum(val(数量)) from wxrk where instr(单号,'" & Text2.Text & "')>0 group by 单号,款号,颜色"
Data4.Database.Execute "update wxxx set 计划='0' where 计划=null"
Data4.Database.Execute "update wxxx set 裁剪='0' where 裁剪=null"
Data4.Database.Execute "update wxxx set 外协='0' where 外协=null"
Data4.Database.Execute "update wxxx set 返回='0' where 返回=null"
Data4.Database.Execute "update wxxx set 共量='0'"
Data4.Database.Execute "insert into wxxx(单号,款号,颜色,计划,裁剪,外协,返回) select 单号,款号,颜色,sum(val(计划)),sum(val(裁剪)),sum(val(外协)),sum(val(返回)) from wxxx group by 单号,款号,颜色"
Data4.Database.Execute "update wxxx set 未协=val(裁剪)-val(外协),未回=val(外协)-val(返回)"
Data4.Database.Execute "delete * from wxxx where 共量='0'"
Data4.RecordSource = "select * from wxxx order by 款号,颜色"
Data4.Refresh
End Sub

Private Sub Label3_Click()
Data4.RecordSource = "select * from cjrb where 单号='" & Text2.Text & "' and 颜色='" & Text3.Text & "' and 发至<>'缝制' and 发至<>'裁剪' order by 日期,序号"
Data4.Refresh
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data4.Recordset.Move rs - 1
Text3.Text = Data4.Recordset.Fields(2)

For i = 0 To 2
Text1(i).Text = Data4.Recordset.Fields(i)
Next
Text1(3).Text = ""
Text1(4).Text = Data4.Recordset.Fields(3)
Text1(8).Text = Data4.Recordset.Fields(5)

End Sub


Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 2
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from wxjl where 单号='" & Text2.Text & "' and 颜色='" & Text3.Text & "' order by 序号 desc"
Data1.Refresh

       Case 6
       Case 9
End Select
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid1.RowSel
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid1.RowSel
End Sub

Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
s3 = MSFlexGrid2.RowSel
End Sub

Private Sub MSFlexGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
s4 = MSFlexGrid2.RowSel
End Sub


Private Sub Text1_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case 2
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from wxjl where 单号='" & Text2.Text & "' and 颜色='" & Text3.Text & "' order by 序号 desc"
Data1.Refresh
End Select

End Sub


Private Sub MSFlex()
On Error Resume Next
With MSFlexGrid2
    c = .Col: r = .Row    '''''C列，，R行
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        ms = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid2.Text = ms
    MSFlexGrid2.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data1.Recordset.MoveFirst
Data1.Recordset.Move r - 1
Data1.Recordset.Edit
Data1.Recordset.Fields(c - 1) = Text1111.Text
Data1.Recordset.Update
Text1111.Visible = False
MSFlexGrid2.Text = Text1111.Text
MSFlexGrid2.SetFocus
End If
End Sub


Private Sub Text2_Change()
Call Label2_DBLClick
End Sub

'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Data1.DatabaseName = "e:\excel\sjzz.MDB"
If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") = 0 Then
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
Text2.Text = l1
End If


'DBCombo2.Text = Node.Index
'DBCombo3.Text = TreeView1.Nodes(Node.Index).FullPath
End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub


Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex
   TreeView1.Nodes.Clear
 

If Option4.Value = True Then
    Data7.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
    Data7.Refresh
    m = 1
    If Not Data7.Recordset.EOF Then  'make sure there are records in the table
        Data7.Recordset.MoveFirst
        Do While Not Data7.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data7.Recordset.Fields(0)
        intIndex = mNode.Index
        Data8.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data7.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
        Data8.Refresh
        
        If Not Data8.Recordset.EOF Then
        Data8.Recordset.MoveFirst
        Do While Not Data8.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data8.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data6.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data8.Recordset.Fields(0) & "' and 进度='进行'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        Data6.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        
        Data8.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        Data7.Recordset.MoveNext
        m = m + 1
        Loop
    End If
End If


If Option5.Value = True Then
    Data7.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
    Data7.Refresh
    m = 1
    If Not Data7.Recordset.EOF Then  'make sure there are records in the table
        Data7.Recordset.MoveFirst
        Do While Not Data7.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(, , Data7.Recordset.Fields(0), Data7.Recordset.Fields(0))
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data7.Recordset.Fields(0)
        intIndex = mNode.Index
        Data8.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data7.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
        Data8.Refresh
        
        If Not Data8.Recordset.EOF Then
        Data8.Recordset.MoveFirst
        Do While Not Data8.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data8.Recordset.Fields(0))
        intIndex = mNode.Index
        Data6.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data8.Recordset.Fields(0) & "' and 进度='结束'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        Data6.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data8.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data7.Recordset.MoveNext
        Loop
    End If
End If

End Sub

