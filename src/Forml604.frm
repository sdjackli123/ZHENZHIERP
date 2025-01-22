VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Forml604 
   BackColor       =   &H00C0E0FF&
   Caption         =   "外加工入库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data9 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data8 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data7 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2775
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   480
      TabIndex        =   40
      Top             =   1560
      Width           =   2775
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1560
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "进行"
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   12
      Left            =   12480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Text            =   "Forml604.frx":0000
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   11280
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   33
      Text            =   "Text2"
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   12480
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3495
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   855
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Width           =   735
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   735
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   735
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   735
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   10320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   8640
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   11280
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   10080
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   8880
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   7560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   14040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12480
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80674817
      CurrentDate     =   39557
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Forml604.frx":0006
      Height          =   4455
      Left            =   3480
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4920
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Forml604.frx":001A
      Height          =   360
      Left            =   8640
      TabIndex        =   19
      Top             =   4320
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
      Bindings        =   "Forml604.frx":002E
      Height          =   2055
      Left            =   3480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   10
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
      ItemData        =   "Forml604.frx":0042
      Left            =   11280
      List            =   "Forml604.frx":0055
      TabIndex        =   39
      Text            =   "Combo1"
      Top             =   3480
      Width           =   1095
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6375
      Left            =   480
      TabIndex        =   44
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
      Left            =   1680
      TabIndex        =   45
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
      Left            =   1680
      TabIndex        =   46
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
      Left            =   480
      TabIndex        =   48
      Top             =   480
      Width           =   1215
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
      Left            =   480
      TabIndex        =   47
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "废片"
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
      Left            =   11280
      TabIndex        =   38
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
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
      Left            =   12480
      TabIndex        =   37
      Top             =   3960
      Width           =   2295
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
      Left            =   3480
      TabIndex        =   32
      Top             =   480
      Width           =   1335
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
      Left            =   11280
      TabIndex        =   31
      Top             =   3120
      Width           =   1095
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
      Left            =   14040
      TabIndex        =   30
      Top             =   3120
      Width           =   735
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
      Left            =   12480
      TabIndex        =   29
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单位"
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
      Left            =   8640
      TabIndex        =   28
      Top             =   3960
      Width           =   1575
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
      Left            =   10080
      TabIndex        =   27
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "计划"
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
      Left            =   8880
      TabIndex        =   26
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格"
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
      Left            =   7560
      TabIndex        =   25
      Top             =   3120
      Width           =   1215
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
      Index           =   4
      Left            =   6120
      TabIndex        =   24
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
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
      Left            =   4680
      TabIndex        =   23
      Top             =   3120
      Width           =   1335
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
      Index           =   2
      Left            =   3480
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "入库数量"
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
      Left            =   10320
      TabIndex        =   21
      Top             =   3960
      Width           =   855
   End
End
Attribute VB_Name = "Forml604"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public l As Integer


Private Sub Combo1_Click()
Text1(6).Text = Combo1.Text
End Sub

Private Sub Command1_Click()
If Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(6).Text = "" Then
MsgBox ("输入不完整!")
Exit Sub
End If

Data1.Recordset.Edit
For i = 0 To 12
Data1.Recordset.Fields(i) = Text1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

Text1(7).Text = ""
Text1(8).Text = ""
Text1(11).Text = 0
Data2.Refresh
Text1(10).Text = 1
If Data2.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data2.Recordset.Fields(0) + 1
End If
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False

End Sub

Private Sub Command2_Click()
'On Error Resume Next
Command2.Enabled = False
Data6.Database.Execute "delete * from wxkc"
Data6.Database.Execute "insert into wxkc(单号,款号,颜色,规格,计划,位置,类别,单位,数量) select 单号,款号,颜色,规格,计划,位置,类别,单位,数量 from wxjl where 单号='" & Text2.Text & "'"
Data6.Database.Execute "insert into wxkc(单号,款号,颜色,规格,计划,位置,类别,单位,数量) select 单号,款号,颜色,规格,计划,位置,类别,单位,-val(数量) from wxrk where 单号='" & Text2.Text & "'"
Data6.Database.Execute "update wxkc set 共量='1'"
Data6.Database.Execute "insert into wxkc(单号,款号,颜色,规格,计划,位置,类别,单位,数量) select 单号,款号,颜色,规格,计划,位置,类别,单位,sum(val(数量)) as 余量 from wxkc group by 单号,款号,颜色,规格,计划,位置,类别,单位 order by 单位,位置"
Data6.Database.Execute "delete * from  wxkc where 共量='1'"
Data4.RecordSource = "select 单号,款号,颜色,规格,计划,位置,类别,单位,数量 from wxkc where val(数量)<>0"
Data4.Refresh
Command2.Enabled = True
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text1(2).Text = "" Or Text1(4).Text = "" Or Text1(6).Text = "" Then
MsgBox ("输入不完整!")
Exit Sub
End If

Data1.Recordset.AddNew
For i = 0 To 12
Data1.Recordset.Fields(i) = Text1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

Text1(7).Text = ""
Text1(8).Text = ""
Text1(11).Text = 0
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
For i = 3 To 12
Text1(i).Text = ""
Next
Text1(9).Text = Date
Text1(11).Text = 0
DTPicker1.Value = Date
Data1.RecordSource = "select * from wxrk where 日期=cdate('" & Text1(9).Text & "') order by 序号 desc"
Data1.Refresh
Data2.RecordSource = "SELECT MAX(序号) FROM wxrk where 日期=cdate('" & Text1(9).Text & "')"
Data2.Refresh

Text1(10).Text = 1
If Data2.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data2.Recordset.Fields(0) + 1
End If
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False
End Sub

Private Sub Command6_Click()
On Error Resume Next
If MsgBox("确定删除吗？，删除不能回复！", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
Text1(7).Text = ""
Text1(8).Text = ""
Text1(11).Text = 0
Data2.Refresh
Text1(10).Text = 1
If Data2.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data2.Recordset.Fields(0) + 1
End If
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False

End Sub

Private Sub Command7_Click()
Call cwrk(MSFlexGrid2, "入外加工报表日期" + Text1(9).Text)
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
Data1.RecordSource = "select * from wxrk where 日期=cdate('" & Text1(9).Text & "') order by 序号 desc"
Data1.Refresh
Text1(9).Text = DTPicker1.Value

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM wxrk where 日期=cdate('" & Text1(9).Text & "')"
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
Data1.RecordSource = "select * from wxrk where 日期=cdate('" & Text1(9).Text & "') order by 序号 desc"
Data1.Refresh
Text1(9).Text = DTPicker1.Value

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM wxrk where 日期=cdate('" & Text1(9).Text & "')"
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
For i = 0 To 12
Text1(i).Text = ""
Next
Text1(9).Text = Date
DTPicker1.Value = Date
DTPicker2.Value = Date - 30
DTPicker3.Value = Date
Text1(11).Text = 0
Text2.Text = ""
DBCombo1.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from wxrk where 日期=cdate('" & Text1(9).Text & "') order by 序号 desc"
Data1.Refresh
Combo1.Text = ""
Option4.Value = True
Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM wxrk where 日期=cdate('" & Text1(9).Text & "')"
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
Data6.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data7.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data8.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data9.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

MSFlexGrid2.ColWidth(0) = 300
For i = 1 To 5
MSFlexGrid2.ColWidth(i) = 1000
Next

MSFlexGrid1.ColWidth(0) = 300
For i = 1 To 5
MSFlexGrid1.ColWidth(i) = 1200
Next

Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False
End Sub


Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data4.Recordset.Move rs - 1
For i = 0 To 8
Text1(i).Text = Data4.Recordset.Fields(i)
Next

End Sub

Private Sub MSFlexGrid2_dblClick()
If Data1.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To 12
Text1(i).Text = Data1.Recordset.Fields(i)
Next
Command3.Enabled = False
Command1.Enabled = True
Command6.Enabled = True
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 6
If Text1(6).Text = "外协" Then
Data5.RecordSource = "select 简称 from wxzl group by 简称"
Data5.Refresh
End If
If Text1(6).Text = "绣印" Then
Data5.RecordSource = "select 简称 from yhzl group by 简称"
Data5.Refresh
End If
End Select
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
        Data9.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data8.Recordset.Fields(0) & "' and 进度='进行'"
        Data9.Refresh
        
        If Not Data9.Recordset.EOF Then
        Data9.Recordset.MoveFirst
        Do While Not Data9.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data9.Recordset.Fields(0))
        Data9.Recordset.MoveNext
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
        Data9.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data8.Recordset.Fields(0) & "' and 进度='结束'"
        Data9.Refresh
        
        If Not Data9.Recordset.EOF Then
        Data9.Recordset.MoveFirst
        Do While Not Data9.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data9.Recordset.Fields(0))
        Data9.Recordset.MoveNext
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

Private Sub Text2_Change()
On Error Resume Next
Command2.Enabled = False
Data6.Database.Execute "delete * from wxkc"
Data6.Database.Execute "insert into wxkc(单号,款号,颜色,规格,计划,位置,类别,单位,数量) select 单号,款号,颜色,规格,计划,位置,类别,单位,数量 from wxjl where 单号='" & Text2.Text & "'"
Data6.Database.Execute "insert into wxkc(单号,款号,颜色,规格,计划,位置,类别,单位,数量) select 单号,款号,颜色,规格,计划,位置,类别,单位,-val(数量) from wxrk where 单号='" & Text2.Text & "'"
Data6.Database.Execute "update wxkc set 共量='1'"
Data6.Database.Execute "insert into wxkc(单号,款号,颜色,规格,计划,位置,类别,单位,数量) select 单号,款号,颜色,规格,计划,位置,类别,单位,sum(val(数量)) as 余量 from wxkc group by 单号,款号,颜色,规格,计划,位置,类别,单位 order by 单位,位置"
Data6.Database.Execute "delete * from  wxkc where 共量='1'"
Data4.RecordSource = "select 单号,款号,颜色,规格,计划,位置,类别,单位,数量 from wxkc where val(数量)<>0"
Data4.Refresh
Command2.Enabled = True
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

