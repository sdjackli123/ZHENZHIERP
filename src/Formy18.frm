VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy18 
   BackColor       =   &H00C0E0FF&
   Caption         =   "主料分析 "
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form18"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5760
      TabIndex        =   157
      Text            =   "Text7"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   4320
      TabIndex        =   156
      Text            =   "Text7"
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "批量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   6480
      Width           =   855
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   5175
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":0000
      Height          =   330
      Index           =   0
      Left            =   4800
      TabIndex        =   138
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   15
      Left            =   12120
      TabIndex        =   136
      Text            =   "Text6"
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   14
      Left            =   12120
      TabIndex        =   135
      Text            =   "Text6"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   13
      Left            =   12120
      TabIndex        =   134
      Text            =   "Text6"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   12
      Left            =   12120
      TabIndex        =   133
      Text            =   "Text6"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   12120
      TabIndex        =   132
      Text            =   "Text6"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   12120
      TabIndex        =   131
      Text            =   "Text6"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   12120
      TabIndex        =   130
      Text            =   "Text6"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   12120
      TabIndex        =   129
      Text            =   "Text6"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   5640
      TabIndex        =   128
      Text            =   "Text6"
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   127
      Text            =   "Text6"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   5640
      TabIndex        =   126
      Text            =   "Text6"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   5640
      TabIndex        =   125
      Text            =   "Text6"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   5640
      TabIndex        =   124
      Text            =   "Text6"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   123
      Text            =   "Text6"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   122
      Text            =   "Text6"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   121
      Text            =   "Text6"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "辅料"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   6480
      Width           =   855
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":0015
      Height          =   330
      Index           =   0
      Left            =   3960
      TabIndex        =   65
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   63
      Text            =   "Text3"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查看"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   6480
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   0
      Left            =   6360
      TabIndex        =   38
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":0029
      Height          =   330
      Index           =   0
      Left            =   2760
      TabIndex        =   30
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":003E
      Height          =   330
      Index           =   0
      Left            =   1320
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   0
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy18.frx":0052
      Height          =   2775
      Left            =   480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7200
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   7
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy18.frx":0066
      Height          =   330
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "款号"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy18.frx":007B
      Height          =   330
      Index           =   2
      Left            =   1320
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "颜色"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":008F
      Height          =   330
      Index           =   1
      Left            =   1320
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":00A3
      Height          =   330
      Index           =   2
      Left            =   1320
      TabIndex        =   18
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":00B7
      Height          =   330
      Index           =   3
      Left            =   1320
      TabIndex        =   19
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy18.frx":00CB
      Height          =   330
      Index           =   1
      Left            =   2760
      TabIndex        =   20
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "刀模名称"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy18.frx":00E0
      Height          =   330
      Index           =   2
      Left            =   2760
      TabIndex        =   21
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "刀模名称"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":00F5
      Height          =   330
      Index           =   4
      Left            =   1320
      TabIndex        =   22
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy18.frx":0109
      Height          =   330
      Index           =   4
      Left            =   2760
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "刀模名称"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":011E
      Height          =   330
      Index           =   5
      Left            =   1320
      TabIndex        =   24
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy18.frx":0132
      Height          =   330
      Index           =   5
      Left            =   2760
      TabIndex        =   25
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "刀模名称"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy18.frx":0147
      Height          =   330
      Index           =   6
      Left            =   2760
      TabIndex        =   26
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "刀模名称"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":015C
      Height          =   330
      Index           =   6
      Left            =   1320
      TabIndex        =   27
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":0170
      Height          =   330
      Index           =   7
      Left            =   1320
      TabIndex        =   28
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":0184
      Height          =   330
      Index           =   1
      Left            =   2760
      TabIndex        =   31
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":0199
      Height          =   330
      Index           =   2
      Left            =   2760
      TabIndex        =   32
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":01AE
      Height          =   330
      Index           =   3
      Left            =   2760
      TabIndex        =   33
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":01C3
      Height          =   330
      Index           =   4
      Left            =   2760
      TabIndex        =   34
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":01D8
      Height          =   330
      Index           =   5
      Left            =   2760
      TabIndex        =   35
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":01ED
      Height          =   330
      Index           =   6
      Left            =   2760
      TabIndex        =   36
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":0202
      Height          =   330
      Index           =   7
      Left            =   2760
      TabIndex        =   37
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   1
      Left            =   6360
      TabIndex        =   39
      Top             =   3000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   2
      Left            =   6360
      TabIndex        =   40
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   3
      Left            =   6360
      TabIndex        =   41
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   4
      Left            =   6360
      TabIndex        =   42
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   5
      Left            =   6360
      TabIndex        =   43
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   6
      Left            =   6360
      TabIndex        =   44
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   7
      Left            =   6360
      TabIndex        =   45
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":0217
      Height          =   330
      Index           =   1
      Left            =   3960
      TabIndex        =   66
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":022B
      Height          =   330
      Index           =   2
      Left            =   3960
      TabIndex        =   67
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":023F
      Height          =   330
      Index           =   3
      Left            =   3960
      TabIndex        =   68
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":0253
      Height          =   330
      Index           =   4
      Left            =   3960
      TabIndex        =   69
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":0267
      Height          =   330
      Index           =   5
      Left            =   3960
      TabIndex        =   70
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":027B
      Height          =   330
      Index           =   6
      Left            =   3960
      TabIndex        =   71
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":028F
      Height          =   330
      Index           =   7
      Left            =   3960
      TabIndex        =   72
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":02A3
      Height          =   330
      Index           =   8
      Left            =   12840
      TabIndex        =   74
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   8
      Left            =   13680
      TabIndex        =   75
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":02B7
      Height          =   330
      Index           =   8
      Left            =   10200
      TabIndex        =   76
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy18.frx":02CC
      Height          =   330
      Index           =   8
      Left            =   11160
      TabIndex        =   77
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "刀模名称"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":02E1
      Height          =   330
      Index           =   8
      Left            =   8760
      TabIndex        =   78
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":02F5
      Height          =   330
      Index           =   9
      Left            =   8760
      TabIndex        =   79
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":0309
      Height          =   330
      Index           =   10
      Left            =   8760
      TabIndex        =   80
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":031D
      Height          =   330
      Index           =   11
      Left            =   8760
      TabIndex        =   81
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":0331
      Height          =   330
      Index           =   12
      Left            =   8760
      TabIndex        =   82
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy18.frx":0345
      Height          =   330
      Index           =   12
      Left            =   11160
      TabIndex        =   83
      Top             =   7800
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "刀模名称"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":035A
      Height          =   330
      Index           =   13
      Left            =   8760
      TabIndex        =   84
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":036E
      Height          =   330
      Index           =   14
      Left            =   8760
      TabIndex        =   85
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy18.frx":0382
      Height          =   330
      Index           =   15
      Left            =   8760
      TabIndex        =   86
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":0396
      Height          =   330
      Index           =   9
      Left            =   10200
      TabIndex        =   87
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":03AB
      Height          =   330
      Index           =   10
      Left            =   10200
      TabIndex        =   88
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":03C0
      Height          =   330
      Index           =   11
      Left            =   10200
      TabIndex        =   89
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":03D5
      Height          =   330
      Index           =   12
      Left            =   10200
      TabIndex        =   90
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":03EA
      Height          =   330
      Index           =   13
      Left            =   10200
      TabIndex        =   91
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":03FF
      Height          =   330
      Index           =   14
      Left            =   10200
      TabIndex        =   92
      Top             =   5400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy18.frx":0414
      Height          =   330
      Index           =   15
      Left            =   10200
      TabIndex        =   93
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   9
      Left            =   13680
      TabIndex        =   94
      Top             =   3000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   10
      Left            =   13680
      TabIndex        =   95
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   11
      Left            =   13680
      TabIndex        =   96
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   12
      Left            =   13680
      TabIndex        =   97
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   13
      Left            =   13680
      TabIndex        =   98
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   14
      Left            =   13680
      TabIndex        =   99
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   15
      Left            =   13680
      TabIndex        =   100
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":0429
      Height          =   330
      Index           =   9
      Left            =   12840
      TabIndex        =   101
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":043D
      Height          =   330
      Index           =   10
      Left            =   12840
      TabIndex        =   102
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":0451
      Height          =   330
      Index           =   11
      Left            =   12840
      TabIndex        =   103
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":0465
      Height          =   330
      Index           =   12
      Left            =   12840
      TabIndex        =   104
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":0479
      Height          =   330
      Index           =   13
      Left            =   12840
      TabIndex        =   105
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":048D
      Height          =   330
      Index           =   14
      Left            =   12840
      TabIndex        =   106
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy18.frx":04A1
      Height          =   330
      Index           =   15
      Left            =   12840
      TabIndex        =   107
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":04B5
      Height          =   330
      Index           =   1
      Left            =   4800
      TabIndex        =   139
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":04CA
      Height          =   330
      Index           =   2
      Left            =   4800
      TabIndex        =   140
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":04DF
      Height          =   330
      Index           =   3
      Left            =   4800
      TabIndex        =   141
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":04F4
      Height          =   330
      Index           =   4
      Left            =   4800
      TabIndex        =   142
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":0509
      Height          =   330
      Index           =   5
      Left            =   4800
      TabIndex        =   143
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":051E
      Height          =   330
      Index           =   6
      Left            =   4800
      TabIndex        =   144
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":0533
      Height          =   330
      Index           =   7
      Left            =   4800
      TabIndex        =   145
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":0548
      Height          =   330
      Index           =   8
      Left            =   11280
      TabIndex        =   146
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":055D
      Height          =   330
      Index           =   9
      Left            =   11280
      TabIndex        =   147
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":0572
      Height          =   330
      Index           =   10
      Left            =   11280
      TabIndex        =   148
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":0587
      Height          =   330
      Index           =   11
      Left            =   11280
      TabIndex        =   149
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":059C
      Height          =   330
      Index           =   12
      Left            =   11280
      TabIndex        =   150
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":05B1
      Height          =   330
      Index           =   13
      Left            =   11280
      TabIndex        =   151
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":05C6
      Height          =   330
      Index           =   14
      Left            =   11280
      TabIndex        =   152
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy18.frx":05DB
      Height          =   330
      Index           =   15
      Left            =   11280
      TabIndex        =   153
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo6"
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "部位"
      Height          =   375
      Index           =   8
      Left            =   5160
      TabIndex        =   155
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "批号"
      Height          =   375
      Index           =   7
      Left            =   12120
      TabIndex        =   137
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "批号"
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   120
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   5
      Left            =   11220
      TabIndex        =   119
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "刀模"
      Height          =   375
      Index           =   4
      Left            =   11160
      TabIndex        =   118
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   117
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前页"
      Height          =   375
      Index           =   15
      Left            =   8040
      TabIndex        =   115
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "中页"
      Height          =   375
      Index           =   14
      Left            =   8040
      TabIndex        =   114
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "后页"
      Height          =   375
      Index           =   13
      Left            =   8040
      TabIndex        =   113
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "上眉"
      Height          =   375
      Index           =   12
      Left            =   8040
      TabIndex        =   112
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "下眉"
      Height          =   375
      Index           =   11
      Left            =   8040
      TabIndex        =   111
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前衬"
      Height          =   375
      Index           =   10
      Left            =   8040
      TabIndex        =   110
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "主料"
      Height          =   375
      Index           =   9
      Left            =   8040
      TabIndex        =   109
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "主料"
      Height          =   375
      Index           =   8
      Left            =   8040
      TabIndex        =   108
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料单位"
      Height          =   375
      Index           =   1
      Left            =   12840
      TabIndex        =   73
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料单位"
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   64
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "主料"
      Height          =   375
      Index           =   7
      Left            =   480
      TabIndex        =   61
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "主料"
      Height          =   375
      Index           =   6
      Left            =   480
      TabIndex        =   60
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前衬"
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   59
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "下眉"
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   58
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "上眉"
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   57
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "后页"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   56
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "中页"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   55
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前页"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   54
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料数量"
      Height          =   375
      Index           =   34
      Left            =   13680
      TabIndex        =   53
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料规格"
      Height          =   375
      Index           =   33
      Left            =   10200
      TabIndex        =   52
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料名称"
      Height          =   375
      Index           =   32
      Left            =   8760
      TabIndex        =   51
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "主料"
      Height          =   375
      Index           =   31
      Left            =   8040
      TabIndex        =   50
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料名称"
      Height          =   375
      Index           =   30
      Left            =   1320
      TabIndex        =   49
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "刀模"
      Height          =   375
      Index           =   29
      Left            =   2760
      TabIndex        =   48
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "主料"
      Height          =   375
      Index           =   44
      Left            =   480
      TabIndex        =   47
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料数量"
      Height          =   375
      Index           =   37
      Left            =   6360
      TabIndex        =   46
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料规格"
      Height          =   375
      Index           =   28
      Left            =   2760
      TabIndex        =   29
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  成衣面料单耗分析表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   4920
      TabIndex        =   15
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   375
      Index           =   13
      Left            =   480
      TabIndex        =   13
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   15
      Left            =   11160
      TabIndex        =   12
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "交货期"
      Height          =   375
      Index           =   16
      Left            =   11160
      TabIndex        =   11
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   19
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "Formy18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X, c, r As Integer: Public ms As String
Dim ba As Database: Dim rr As Integer
Dim rs As Single: Dim rd1 As Recordset: Dim ba1 As Database: Public ll As Integer: Public RQ As Date
Dim rd As Recordset: Public mm As Date: Public ml As Date


Private Sub Command12_Click()
Formy181.Text1.Text = DBCombo1(0).Text
Formy181.Text2.Text = DBCombo1(1).Text
Formy181.DBCombo1.Text = DBCombo1(2).Text
Formy181.Show
End Sub


Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
If DBCombo1(2).Text = "" Then
MsgBox ("请确认颜色")
Exit Sub
End If

If DBCombo1(1).Text = "" Then
MsgBox ("请确认款号")
Exit Sub
End If

If Text7.Text = "" Then
If MsgBox("请输入部位,不输入可以继续，继续吗？", vbYesNo) = vbNo Then Exit Sub
End If

rd.AddNew

For i = 0 To 15
If DBCombo5(i).Text <> "" Then
rd.AddNew
rd.Fields(0) = Trim(DBCombo1(0).Text)
rd.Fields(1) = Trim(DBCombo1(1).Text)
rd.Fields(2) = Trim(DBCombo1(2).Text)
rd.Fields(3) = Trim(Label3(i).Caption)
rd.Fields(4) = Trim(DBCombo2(i).Text)
rd.Fields(5) = Trim(DBCombo4(i).Text)
rd.Fields(6) = Trim(DBCombo9(i).Text)
rd.Fields(7) = Trim(DBCombo6(i).Text)
rd.Fields(8) = Trim(Text6(i).Text)
rd.Fields(9) = Trim(DBCombo5(i).Text)
rd.Fields(10) = "1主料库"
rd.Fields(11) = Trim(Text7.Text)
rd.Update
End If
Next

MsgBox ("保存成功！")

Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.RecordSource = "select * from dlclb WHERE 款号='" & DBCombo1(1).Text & "' AND 款号='" & DBCombo1(1).Text & "' AND 订单颜色='" & DBCombo1(2).Text & "' AND 材料库类='1主料库' order by 部位 desc"
Data4.Refresh

End Sub

Private Sub Command2_Click()
On Error Resume Next

If DBCombo1(2).Text = "" Then
MsgBox ("请确认颜色")
Exit Sub
End If

If DBCombo1(1).Text = "" Then
MsgBox ("请确认款号")
Exit Sub
End If

If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If

If Text7.Text = "" Then
If MsgBox("请输入部位,不输入可以继续，继续吗？", vbYesNo) = vbNo Then Exit Sub
End If

Data4.Recordset.Edit
For i = 0 To 15
If DBCombo5(i).Text <> "" Then
Data4.Recordset.Edit
Data4.Recordset.Fields(0) = Trim(DBCombo1(0).Text)
Data4.Recordset.Fields(1) = Trim(DBCombo1(1).Text)
Data4.Recordset.Fields(2) = Trim(DBCombo1(2).Text)
Data4.Recordset.Fields(3) = Trim(Label3(i).Caption)
Data4.Recordset.Fields(4) = Trim(DBCombo2(i).Text)
Data4.Recordset.Fields(5) = Trim(DBCombo4(i).Text)
Data4.Recordset.Fields(6) = Trim(DBCombo9(i).Text)
Data4.Recordset.Fields(7) = Trim(DBCombo6(i).Text)
Data4.Recordset.Fields(8) = Trim(Text6(i).Text)
Data4.Recordset.Fields(9) = Trim(DBCombo5(i).Text)
Data4.Recordset.Fields(11) = Trim(Text7.Text)
Data4.Recordset.Update
End If
Next
Data4.Recordset.Update
MsgBox ("修改保存成功！")
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.RecordSource = "select * from dlclb WHERE 款号='" & DBCombo1(1).Text & "' AND 款号='" & DBCombo1(1).Text & "' AND 订单颜色='" & DBCombo1(2).Text & "' AND 材料库类='1主料库' order by 部位 desc"
Data4.Refresh



End Sub

Private Sub Command4_Click()

On Error Resume Next
If MsgBox("确定删除吗?,删除将不能恢复!", vbYesNo) = vbNo Then Exit Sub
Data4.Recordset.Delete
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.RecordSource = "select * from dlclb WHERE 款号='" & DBCombo1(1).Text & "' AND 款号='" & DBCombo1(1).Text & "' AND 订单颜色='" & DBCombo1(2).Text & "' AND 材料库类='1主料库' order by 部位 desc"
Data4.Refresh

End Sub

Private Sub Command5_Click()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.EOF
For i = 0 To 15
If Label3(i).Caption = Data4.Recordset.Fields(3) Then
DBCombo2(i).Text = Data4.Recordset.Fields(4)
DBCombo4(i).Text = Data4.Recordset.Fields(5)
DBCombo5(i).Text = Format(Data4.Recordset.Fields(9), "#0.00")
DBCombo9(i).Text = Data4.Recordset.Fields(6)
DBCombo6(i).Text = Data4.Recordset.Fields(7)
Text6(i).Text = Data4.Recordset.Fields(8)
End If
Next
Data4.Recordset.MoveNext
Loop

End Sub

Private Sub Command6_Click()
'DataEnvironment4.cldfd DBCombo1(0).Text
'DataReport5.Show 1
'DataEnvironment4.rscldfd.Close
Call MXOutDataToExcel(MSFlexGrid1, "主料耗料表")
End Sub


Private Sub Command8_Click()
On Error Resume Next
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh
Data6.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False


End Sub





Private Sub DBCombo11_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo12_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo13_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo14_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo15_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo16_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub


Private Sub DTPicker1_Change()

End Sub

Private Sub DTPicker2_Change()
Text5.Text = DTPicker2.Value
End Sub

Private Sub DTPicker3_Change()
Text8.Text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
Text8.Text = DTPicker3.Value
Text8.SetFocus
End Sub

Private Sub Command9_Click()
On Error Resume Next
For i = 0 To 4
Formy23.DBCombo1(i).Text = DBCombo1(i).Text
Next
Formy23.Show
End Sub

Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data11.RecordSource = "select 款号 from KSNR WHERE 款号='" & DBCombo1(1).Text & "'  GROUP BY 款号 "
Data11.Refresh
       Case 1


Data5.RecordSource = "select 颜色 from KSNR WHERE 款号='" & DBCombo1(1).Text & "' GROUP BY 颜色 "
Data5.Refresh


DBCombo1(3).Text = 1
     Case 2
End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 1


Data5.RecordSource = "select 颜色 from KSNR WHERE 款号='" & DBCombo1(1).Text & "' AND 款号='" & DBCombo1(1).Text & "' GROUP BY 颜色 "
Data5.Refresh


     
     Case 2
Data1.RecordSource = "select * from KSNR WHERE 款号='" & DBCombo1(1).Text & "' AND 颜色='" & DBCombo1(2).Text & "' "
Data1.Refresh
If Data1.Recordset.EOF Then
For i = 0 To 15
Label3(i).Caption = ""
Next
Else
l = 0
For i = 0 To 7
If Data1.Recordset.Fields(4 + l) <> "" Then
Label3(i).Caption = Data1.Recordset.Fields(4 + l)
Else
Label3(i).Caption = ""
End If
l = l + 1
Next
End If

Text1.Text = Data1.Recordset.Fields(16)
Text2.Text = Data1.Recordset.Fields(17)
Text3.Text = DBCombo1(2).Text

'For i = 0 To 15
'DBCombo6(i).Text = Text3.Text
'Next


For i = 3 To Data1.Recordset.Fields.Count - 1
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next

Data4.RecordSource = "select * from dlclb WHERE 款号='" & DBCombo1(1).Text & "' AND 订单颜色='" & DBCombo1(2).Text & "' AND 材料库类='1主料库' order by 部位 desc"
Data4.Refresh
DBCombo1(3).Text = 12
Text5.Text = DBCombo1(4).Text
End Select

End Sub

Private Sub DBCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo2_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label3(i).Caption <> "" Then
       DBCombo2(i).Text = DBCombo2(Index).Text
       End If
       Next
End Select
End Sub

Private Sub DBCombo2_Click(Index As Integer, Area As Integer)
Select Case Index
       Case Index
       Data13.RecordSource = "SELECT 材料规格 FROM CLMC WHERE 库类='1主料库' AND 材料名称='" & DBCombo2(Index).Text & "' GROUP BY 材料规格 "
       Data13.Refresh
       Data14.RecordSource = "SELECT 库类,材料名称,材料规格,材料单位,颜色 FROM CLMC WHERE 库类='1主料库' AND 材料名称='" & DBCombo2(Index).Text & "' GROUP BY 库类,材料名称,材料规格,材料单位,颜色"
       Data14.Refresh
End Select
 
End Sub



Private Sub DBCombo3_Click(Index As Integer, Area As Integer)
Select Case Index
       Case Index
  '     Data11.RecordSource = "SELECT * FROM DMSZ WHERE DMSZ.刀模名称='" & DBCombo3(Index).Text & "' AND INSTR('" & Label3(Index).Caption & "',DMSZ.刀模位置)>0"
  '     Data11.Refresh
  '     If Data11.Recordset.EOF Then
       
  '     DBCombo5(Index).Text = ""
  '     DBCombo9(Index).Text = ""
 '      Else
  '     DBCombo5(Index).Text = Data11.Recordset.Fields(3)
  '     DBCombo9(Index).Text = Data11.Recordset.Fields(2)
  '     End If
End Select
End Sub



Private Sub DBCombo4_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label3(i).Caption <> "" Then
       DBCombo4(i).Text = DBCombo4(Index).Text
       End If
       Next
End Select

End Sub
Private Sub DBCombo5_Click(Index As Integer, Area As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label3(i).Caption <> "" Then
       DBCombo5(i).Text = DBCombo5(Index).Text
       End If
       Next
End Select
End Sub

Private Sub DBCombo6_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label3(i).Caption <> "" Then
       DBCombo6(i).Text = DBCombo6(Index).Text
       End If
       Next
End Select

End Sub

Private Sub DBCombo6_Click(Index As Integer, Area As Integer)
Select Case Index
       Case Index
       Data14.RecordSource = "SELECT 库类,材料名称,材料规格,材料单位,颜色 FROM CLMC WHERE 库类='1主料库' AND 材料名称='" & DBCombo2(Index).Text & "' AND 颜色='" & DBCombo6(Index).Text & "' GROUP BY 库类,材料名称,材料规格,材料单位,颜色 "
       Data14.Refresh
End Select

End Sub

Private Sub DBCombo9_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label3(i).Caption <> "" Then
       DBCombo9(i).Text = DBCombo9(Index).Text
       End If
       Next
End Select

End Sub


Private Sub Form_Load()
On Error Resume Next

Set ba = OpenDatabase("d:\数据库\\htgl\2011\SCZYJHD.MDB")
Set rd = ba.OpenRecordset("DLCLB", dbOpenDynaset)

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text7.Text = ""
For i = 0 To 19
DBCombo1(i).Text = ""
Next

For i = 0 To 15
DBCombo2(i).Text = ""
Next

For i = 0 To 15
DBCombo3(i).Text = ""
Next

For i = 0 To 15
DBCombo4(i).Text = ""
Next

For i = 0 To 15
DBCombo5(i).Text = ""
Next

For i = 0 To 15
DBCombo9(i).Text = "公斤"
Next

For i = 0 To 15
DBCombo6(i).Text = ""
Text6(i).Text = ""
Next

Text5.Text = ""


Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from KSNR WHERE 款号='" & DBCombo1(1).Text & "' ORDER BY VAL(序号) DESC"
Data1.Refresh

For i = 0 To Data1.Recordset.Fields.Count - 1
DBCombo1(i).Text = ""
Next



Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "select MAX(VAL(序号)) from KSNR  WHERE 款号='" & DBCombo1(1).Text & "'"
Data2.Refresh



Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select * from SCZY_Z WHERE SCZY_Z.款号='" & DBCombo1(1).Text & "' "
Data3.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data4.RecordSource = "select * from dlclb WHERE 款号='" & DBCombo1(1).Text & "' AND 款号='" & DBCombo1(1).Text & "' AND 订单颜色='" & DBCombo1(2).Text & "' AND 材料库类='1主料库' order by 部位 desc"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"

Data6.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data6.RecordSource = "SELECT 材料名称 FROM CLMC WHERE 库类='1主料库' GROUP BY 材料名称 "
Data6.Refresh

Data7.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data7.RecordSource = "SELECT cldw.mc FROM cldw  GROUP BY cldw.mc"
Data7.Refresh


Data9.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data9.RecordSource = "SELECT * FROM ZHLMX"
Data9.Refresh



Data11.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data11.Refresh


Data12.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data12.RecordSource = "SELECT YS FROM YS  GROUP BY YS "
Data12.Refresh




For i = 0 To 15
Label3(i).Caption = ""
Next

i = 0
Data9.Recordset.MoveFirst
Do While Not Data9.Recordset.EOF
Label3(i).Caption = Data9.Recordset.Fields(0)
i = i + 1
Data9.Recordset.MoveNext
Loop

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 0
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 1200
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.ColWidth(6) = 1200
MSFlexGrid1.ColWidth(7) = 1200
MSFlexGrid1.ColWidth(8) = 1200

DBCombo1(1).TabIndex = 0
End Sub



Private Sub Label2_DBLClick(Index As Integer)
On Error Resume Next
Select Case Index
       Case 44
For i = 0 To 15
       
If DBCombo3(i).Text <> "" Then
       Data11.RecordSource = "SELECT * FROM DMSZ WHERE DMSZ.刀模名称='" & DBCombo3(Index).Text & "' AND INSTR('" & Label3(Index).Caption & "',DMSZ.刀模位置)>0"
       Data11.Refresh
       If Data11.Recordset.EOF Then
       
       DBCombo5(Index).Text = ""
       DBCombo9(Index).Text = ""
       Else
       DBCombo5(Index).Text = Data11.Recordset.Fields(3)
       DBCombo9(Index).Text = Data11.Recordset.Fields(2)
       End If
End If

Next
End Select
End Sub




Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
Data4.Recordset.MoveFirst
Data4.Recordset.Move rs - 1

For i = 0 To 15
If Data4.Recordset.Fields(3) = Label3(i).Caption Then
DBCombo2(i).Text = Data4.Recordset.Fields(4)
DBCombo4(i).Text = Data4.Recordset.Fields(5)
DBCombo5(i).Text = Format(Data4.Recordset.Fields(9), "#0.0000")
DBCombo9(i).Text = Data4.Recordset.Fields(6)
DBCombo6(i).Text = Data4.Recordset.Fields(7)
Text6(i).Text = Trim(Data4.Recordset.Fields(8))
Else
DBCombo2(i).Text = ""
DBCombo4(i).Text = ""
DBCombo5(i).Text = ""
DBCombo9(i).Text = ""
DBCombo6(i).Text = ""
Text6(i).Text = ""
End If

Next

Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text6_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label3(i).Caption <> "" Then
       Text6(i).Text = Text6(Index).Text
       End If
       Next
End Select
End Sub

Private Sub MSFlex()
On Error Resume Next
With MSFlexGrid1
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

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid1.Text = ms
    MSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data4.Recordset.MoveFirst
Data4.Recordset.Move r - 1
Data4.Recordset.Edit
Data4.Recordset.Fields(c - 1) = Text1111.Text
Data4.Recordset.Update
Text1111.Visible = False
MSFlexGrid1.Text = Text1111.Text
MSFlexGrid1.SetFocus
End If
End Sub


