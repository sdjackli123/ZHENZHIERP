VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formc29 
   BackColor       =   &H00C0E0FF&
   Caption         =   "库存材料备料"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form29"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data16 
      Caption         =   "Data16"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   960
      Width           =   1095
   End
   Begin VB.Data Data15 
      Caption         =   "Data15"
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
      Top             =   10560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   67
      Text            =   "Text2"
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按库类查看"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按材料查看"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按颜色查看"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7680
      Width           =   1455
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formc29.frx":0000
      Height          =   330
      Left            =   960
      TabIndex        =   60
      Top             =   7680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo2"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库存信息"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
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
      Top             =   8160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   -120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   -120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command33 
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
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
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
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Width           =   3135
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   13440
      Top             =   4560
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc29.frx":0015
      Height          =   2535
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5040
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   14
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
      Height          =   330
      Index           =   2
      Left            =   840
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc29.frx":0029
      Height          =   330
      Index           =   0
      Left            =   840
      TabIndex        =   12
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      DataSource      =   "Data2"
      Height          =   330
      Index           =   1
      Left            =   840
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc29.frx":003D
      Height          =   330
      Index           =   3
      Left            =   840
      TabIndex        =   14
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   4
      Left            =   840
      TabIndex        =   15
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   5
      Left            =   3360
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      DataSource      =   "Data3"
      Height          =   330
      Index           =   9
      Left            =   3360
      TabIndex        =   17
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   -120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81199105
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5520
      TabIndex        =   19
      Top             =   -120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81199105
      CurrentDate     =   39177
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   240
      Index           =   23
      Left            =   5160
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   423
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   5.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formc29.frx":0051
      Height          =   1815
      Left            =   120
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8040
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   18
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "客户信息"
      Height          =   2775
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   7575
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   6
         Left            =   3240
         TabIndex        =   23
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   8
         Left            =   3240
         TabIndex        =   24
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc29.frx":0066
         Height          =   330
         Index           =   10
         Left            =   4920
         TabIndex        =   25
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "XM"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   3240
         TabIndex        =   26
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc29.frx":007A
         Height          =   330
         Index           =   11
         Left            =   4920
         TabIndex        =   46
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "XM"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc29.frx":008E
         Height          =   330
         Index           =   12
         Left            =   4920
         TabIndex        =   47
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "MC"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc29.frx":00A2
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   4920
         TabIndex        =   48
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "MC"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   4920
         TabIndex        =   50
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   15
         Left            =   6360
         TabIndex        =   55
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   6360
         TabIndex        =   56
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "库类"
         Height          =   375
         Index           =   4
         Left            =   4320
         TabIndex        =   58
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "库别"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   57
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "仓务员"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   49
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "序号"
         Height          =   375
         Index           =   15
         Left            =   6360
         TabIndex        =   40
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "日期"
         Height          =   375
         Index           =   14
         Left            =   6360
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "备注"
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   38
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "数量"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   37
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单价"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   36
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "批号"
         Height          =   375
         Index           =   10
         Left            =   2640
         TabIndex        =   35
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单号"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "合计金额"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "颜色"
         Height          =   375
         Index           =   9
         Left            =   2640
         TabIndex        =   32
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "材料规格"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "合约号"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "材料单位"
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "材料名称"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "客户"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   120
      TabIndex        =   52
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81199105
      CurrentDate     =   39883
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formc29.frx":00B6
      Height          =   4335
      Left            =   7800
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "记录号 "
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Bindings        =   "Formc29.frx":00CB
      Height          =   2655
      Left            =   7440
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   5040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "记录号 "
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "当前的操作单号"
      Height          =   375
      Index           =   6
      Left            =   2160
      TabIndex        =   66
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF00FF&
      Caption         =   "颜色"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   61
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   54
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作月份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   53
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单据号"
      Height          =   375
      Index           =   17
      Left            =   7080
      TabIndex        =   45
      Top             =   -240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   10440
      TabIndex        =   44
      Top             =   -120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   10440
      TabIndex        =   43
      Top             =   -75
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "制 衣 材 料 出 库 "
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
      Left            =   2160
      TabIndex        =   42
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF00FF&
      Caption         =   "库存数量"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   41
      Top             =   4320
      Width           =   1335
   End
End
Attribute VB_Name = "Formc29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X, BAR, SHX, CZBZH, CZSX, SJBL As Integer ''''''''CZBZH操作标志
Dim ba As Database: Dim rr As Integer
Dim rs As Single: Dim rd As Recordset: Dim ba1 As Database: Public ll As Integer: Public K1, K2 As String
Dim rd1 As Recordset
Dim A As String  '中间变量
Dim B As Double
Dim c As Integer
Dim kg As Integer
Dim bb As Long
Dim cc As String
Dim kkf As Integer
Dim N As Integer
Dim DH  As Integer
Dim fh As String



Private Sub Command12_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command1_Click()
Data2.Database.Execute "DELETE * FROM CLRCZZ"
Data2.Database.Execute "DELETE * FROM CLRCZZHZ"
Data2.Database.Execute "INSERT INTO CLRCZZ(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类) select CKGL.材料名称,CKGL.材料规格,CKGL.材料单位,CKGL.颜色,CKGL.批次,CKGL.数量,CKGL.单价,CKGL.库类 from ckgl WHERE CKGL.库别='清库库存' AND CKGL.库类='" & DBCombo1(12).Text & "' AND CKGL.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data2.Database.Execute "UPDATE CLRCZZ SET 库别='入库' where 库别=NULL"
Data2.Database.Execute "INSERT INTO CLRCZZ(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类) select CKBL.材料名称,CKBL.材料规格,CKBL.材料单位,CKBL.颜色,CKBL.批次,CKBL.数量,CKBL.单价,CKBL.库类 from ckBL WHERE CKBL.库别='清库库存' AND CKBL.库类='" & DBCombo1(12).Text & "' AND CKBL.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data2.Database.Execute "UPDATE CLRCZZ SET 库别='出库',数量=-数量 WHERE 库别=NULL"
Data2.Database.Execute "INSERT INTO CLRCZZHZ(库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价) SELECT CLRCZZ.库类,CLRCZZ.材料名称,CLRCZZ.材料规格,CLRCZZ.材料单位,CLRCZZ.颜色,CLRCZZ.批次,SUM(CLRCZZ.数量) AS L,AVG(CLRCZZ.单价) AS D FROM CLRCZZ GROUP BY CLRCZZ.库类,CLRCZZ.材料名称,CLRCZZ.材料规格,CLRCZZ.材料单位,CLRCZZ.颜色,CLRCZZ.批次"
Data13.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.数量>0"
Data13.Refresh
End Sub

Private Sub Command10_Click()
Call SX2(Data15, MSFlexGrid3, 7)
End Sub

Private Sub Command3_Click()
Data13.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.材料名称='" & DBCombo1(3).Text & "' AND CLRCZZHZ.颜色='" & DBCombo2.Text & "' AND CLRCZZHZ.数量>0"
Data13.Refresh
End Sub

Private Sub Command33_Click()
Unload Me
End Sub


Private Sub Command5_Click()
Data13.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.材料名称='" & DBCombo1(3).Text & "' AND CLRCZZHZ.数量>0"
Data13.Refresh
End Sub

Private Sub Command7_Click()

If Text4.Text = "" Then
MsgBox ("请输入日期!")
Exit Sub
End If
If Text5.Text = "" Then
MsgBox ("请输入日期!")
Exit Sub
End If

Data5.RecordSource = "select   * from ckBL WHERE CKBL.单号='" & DBCombo1(1).Text & "' order by Val(ckBL.序号)"
Data5.Refresh

End Sub


Private Sub Command11_Click()
On Error Resume Next
If CZBZH = 0 Then
MsgBox ("请选择仓库库存信息")
Exit Sub
End If

If DBCombo1(11).Text = "" Then
MsgBox ("请选择负责人！")
Exit Sub
End If

If DBCombo1(8).Text = "" Or DBCombo1(9).Text = "" Then
MsgBox ("数量、单价不能为空！！")
Exit Sub
End If

If Val(DBCombo1(8).Text) > Val(Text1.Text) Then
MsgBox ("输入数量有误，请重新输入")
Exit Sub
End If


rd.AddNew
For i = 0 To rd.Fields.Count - 1
rd.Fields(i) = DBCombo1(i).Text
Next
rd.Update

'For i = 3 To RD.Fields.Count - 3
'If i = 18 Then
'CWY = DBCombo1(i).Text
'End If
'DBCombo1(i).Text = ""
'Next

'Data13.RecordSource = "select  * from ckgl WHERE CKgl.材料名称='" & DBCombo1(3).Text & "' and CKGL.数量>CKGL.实领量 and CKGL.库类='" & DBCombo1(15).Text & "'  order by Val(ckgl.序号)"
Data13.Refresh
Data5.RecordSource = "select   * from ckBL WHERE CKBL.单号='" & DBCombo1(1).Text & "' order by Val(ckBL.序号)"
Data5.Refresh
Data7.RecordSource = "select MAX(VAL(CKbl.序号)) from ckbl "
Data7.Refresh
If Data5.Recordset.EOF Then
MSFlexGrid1.TextMatrix(0, 0) = "记录号"
Exit Sub
End If
Data5.Recordset.MoveLast
MSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To Data5.Recordset.RecordCount
MSFlexGrid1.TextMatrix(i, 0) = i
Next

Data7.RecordSource = "select   MAX(VAL(CKGL.序号)) from ckgl "
Data7.Refresh
DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(15).Text = Date

CZBZH = 0
End Sub

Private Sub Command2_Click()
On Error Resume Next

If DBCombo1(11).Text = "" Then
MsgBox ("请选择负责人！")
Exit Sub
End If

If DBCombo1(8).Text = "" Or DBCombo1(9).Text = "" Then
MsgBox ("数量、单价不能为空！！")
Exit Sub
End If

Data5.Recordset.Edit
For i = 0 To Data5.Recordset.Fields.Count - 1
Data5.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data5.Recordset.Update


'For i = 3 To RD.Fields.Count - 3
'If i = 18 Then
'CWY = DBCombo1(i).Text
'End If
'DBCombo1(i).Text = ""
'Next

Data5.RecordSource = "select   * from ckBL WHERE CKBL.单号='" & DBCombo1(1).Text & "' order by Val(ckBL.序号)"
Data5.Refresh

Data7.RecordSource = "select   MAX(VAL(CKGL.序号)) from ckgl "
Data7.Refresh
DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(15).Text = Date
DBCombo1(0).SetFocus
End Sub

Private Sub Command4_Click()
On Error Resume Next

Data5.Recordset.Delete

'For i = 3 To RD.Fields.Count - 2
'If i = 18 Then
'CWY = DBCombo1(i).Text
'End If
'DBCombo1(i).Text = ""
'Next


Data5.RecordSource = "select   * from ckBL WHERE CKBL.单号='" & DBCombo1(1).Text & "' order by Val(ckBL.序号)"
Data5.Refresh
Data7.RecordSource = "select   MAX(VAL(CKGL.序号)) from ckgl "
Data7.Refresh
DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(15).Text = Date
DBCombo1(0).SetFocus
End Sub


Private Sub Command8_Click()
On Error Resume Next
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
DBCombo1(15).Text = Date
DBCombo1(8).Text = 0
SHX = 0
Data1.Refresh
Data3.Refresh
Data4.Refresh
Data6.Refresh
Data8.Refresh
Data9.Refresh
Data7.Database.Execute "UPDATe CKGL SET 序号='0'  WHERE 序号=null"
Data7.RecordSource = "select   MAX(VAL(CKGL.序号)) from ckgl "
Data7.Refresh
DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(15).Text = Date
Data5.RecordSource = "select   * from ckBL WHERE CKBL.单号='" & DBCombo1(1).Text & "' order by Val(ckBL.序号)"
Data5.Refresh
End Sub

Private Sub Command6_Click()
If Data5.Recordset.EOF Then
MsgBox ("此单据号中无记录，不能打印！")
Exit Sub
End If

BAR = 1
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub


Private Sub Command9_Click()
Data13.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.库类='" & DBCombo1(12).Text & "' AND CLRCZZHZ.数量>0"
Data13.Refresh
End Sub

Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 1
       Data5.RecordSource = "select   * from ckBL WHERE CKBL.单号='" & DBCombo1(1).Text & "' order by Val(ckBL.序号)"
       Data5.Refresh
       Case 8
       DBCombo1(10).Text = Format(Val(DBCombo1(8).Text) * Val(DBCombo1(9).Text), "#0.00")
End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 1
       Data5.RecordSource = "select   * from ckBL WHERE CKBL.单号='" & DBCombo1(1).Text & "' order by Val(ckBL.序号)"
       Data5.Refresh
       Case 8
       DBCombo1(10).Text = Format(Val(DBCombo1(8).Text) * Val(DBCombo1(9).Text), "#0.00")
End Select

End Sub

Private Sub DBCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text4.Text = DTPicker1.Value
Text4.SetFocus
End Sub

Private Sub DTPicker2_Change()
Text5.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text5.Text = DTPicker2.Value
Text5.SetFocus
End Sub

Private Sub DTPicker3_Change()
Text3.Text = Month(DTPicker3.Value)
End Sub

Private Sub DTPicker3_CloseUp()
Text3.Text = Month(DTPicker3.Value)
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
Text2.Text = ""
Text3.Text = Month(Date)
Select Case Text3.Text
       Case 1
K1 = Format(Date, "YYYY") + "-" + "01" + "-01"
K2 = Format(Date, "YYYY") + "-" + "01" + "-31"
       Case 2
If Val(Format(Date, "YYYY")) / 4 = Int(Val(Format(Date, "YYYY")) / 4) Then
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-29"
Else
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-28"
End If
       Case 3
K1 = Format(Date, "YYYY") + "-" + "03" + "-01"
K2 = Format(Date, "YYYY") + "-" + "03" + "-31"
       Case 4
K1 = Format(Date, "YYYY") + "-" + "04" + "-01"
K2 = Format(Date, "YYYY") + "-" + "04" + "-30"
       Case 5
K1 = Format(Date, "YYYY") + "-" + "05" + "-01"
K2 = Format(Date, "YYYY") + "-" + "05" + "-31"
       Case 6
K1 = Format(Date, "YYYY") + "-" + "06" + "-01"
K2 = Format(Date, "YYYY") + "-" + "06" + "-30"
       Case 7
K1 = Format(Date, "YYYY") + "-" + "07" + "-01"
K2 = Format(Date, "YYYY") + "-" + "07" + "-31"
       Case 8
K1 = Format(Date, "YYYY") + "-" + "08" + "-01"
K2 = Format(Date, "YYYY") + "-" + "08" + "-30"
       Case 9
K1 = Format(Date, "YYYY") + "-" + "09" + "-01"
K2 = Format(Date, "YYYY") + "-" + "09" + "-31"
       Case 10
K1 = Format(Date, "YYYY") + "-" + "10" + "-01"
K2 = Format(Date, "YYYY") + "-" + "10" + "-30"
       Case 11
K1 = Format(Date, "YYYY") + "-" + "11" + "-01"
K2 = Format(Date, "YYYY") + "-" + "11" + "-31"
       Case 12
K1 = Format(Date, "YYYY") + "-" + "12" + "-01"
K2 = Format(Date, "YYYY") + "-" + "12" + "-30"
End Select

Set ba = OpenDatabase("d:\数据库\\htgl\2011\ckgl.MDB")
Set rd = ba.OpenRecordset("ckBL", dbOpenDynaset)

For i = 0 To rd.Fields.Count - 1
DBCombo1(i).Text = ""
Next
DBCombo2.Text = ""
DBCombo1(15).Text = Date
DBCombo1(16).Enabled = False
DBCombo1(8).Text = 0
DBCombo1(13).Text = "清库库存"


Data1.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data1.RecordSource = "select 简称 from KHZL group by 简称"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\CKGL.mdb"
Data2.RecordSource = "select CKGL.材料名称 from ckgl WHERE CKGL.库别<>'采购入库' group by ckgl.材料名称"
Data2.Refresh


Data3.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data3.RecordSource = "select CW.MC from CW group by CW.MC"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data4.RecordSource = "select fzr.xm  from fzr group by fzr.xm"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data5.RecordSource = "select   * from ckBL WHERE CKBL.单号='" & DBCombo1(1).Text & "' order by Val(ckBL.序号)"
Data5.Refresh

Data6.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data6.RecordSource = "select KL.MC from KL group by KL.MC"
Data6.Refresh

Data7.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data7.RecordSource = "select   MAX(VAL(CKBL.序号)) from ckBL "
Data7.Refresh

DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1

Data8.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data8.RecordSource = "select KB.MC from KB group by KB.MC"
Data8.Refresh


Data13.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"

Data9.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data9.RecordSource = "select CLDW.MC from CLDW group by CLDW.MC"
Data9.Refresh

Data10.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"

Data11.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"

Data14.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data14.RecordSource = "select YS.YS  from YS group by YS.YS"
Data14.Refresh

Data15.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data15.Refresh

Data16.DatabaseName = "d:\数据库\\htgl\2011\CKGL.mdb"
Data16.Refresh

MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1000
MSFlexGrid1.ColWidth(8) = 1200
MSFlexGrid1.ColWidth(9) = 1200
MSFlexGrid1.ColWidth(14) = 1200
MSFlexGrid1.ColWidth(15) = 1200

MSFlexGrid2.ColWidth(1) = 2500
MSFlexGrid2.ColWidth(2) = 1500
MSFlexGrid3.ColWidth(0) = 300



DBCombo1(0).TabIndex = 0

Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

End Sub




Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
Data5.Recordset.MoveFirst
Data5.Recordset.Move rs - 1
For i = 0 To Data5.Recordset.Fields.Count - 1
DBCombo1(i).Text = Data5.Recordset.Fields(i)
Next
Command11.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub

Private Sub MSFlexGrid2_dblClick()
On Error Resume Next
If Data13.Recordset.EOF Then
Exit Sub
End If
rs = MSFlexGrid2.Row
Data13.Recordset.MoveFirst
Data13.Recordset.Move rs - 1
Text1.Text = Data13.Recordset.Fields(5)
For i = 0 To Data13.Recordset.Fields.Count - 2
DBCombo1(3 + i).Text = Data13.Recordset.Fields(i)
Next
CZBZH = 1
End Sub


Private Sub MSFlexGrid3_Click()
On Error Resume Next
rs = MSFlexGrid3.Row
'If Data2.Recordset.EOF Then Exit Sub
Data15.Recordset.MoveFirst
Data15.Recordset.Move rs - 1
DBCombo1(12).Text = Data15.Recordset.Fields(0)
Data16.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.库类='" & Data15.Recordset.Fields(0) & "' AND CLRCZZHZ.材料名称='" & Data15.Recordset.Fields(1) & "' AND 颜色='" & Data15.Recordset.Fields(4) & "' AND CLRCZZHZ.数量>0"
Data16.Refresh
End Sub

Private Sub MSFlexGrid4_Click()
On Error Resume Next
rs = MSFlexGrid4.Row
Data16.Recordset.MoveFirst
Data16.Recordset.Move rs - 1
DBCombo1(12).Text = Data16.Recordset.Fields(7)
DBCombo1(3).Text = Data16.Recordset.Fields(0)
DBCombo2.Text = Data16.Recordset.Fields(3)
DBCombo1(1).Text = Text2.Text
End Sub

Private Sub Text2_Change()
Data15.RecordSource = "SELECT DHCLB.材料库类,DHCLB.材料名称,DHCLB.材料规格,DHCLB.材料单位,DHCLB.材料颜色,DHCLB.材料批号,SUM(DHCLB.材料数量) AS 备料量 FROM DHCLB WHERE DHCLB.单号='" & Text2.Text & "' GROUP BY DHCLB.材料库类,DHCLB.材料名称,DHCLB.材料规格,DHCLB.材料单位,DHCLB.材料颜色,DHCLB.材料批号"
Data15.Refresh
End Sub

Private Sub Text3_Change()
Select Case Text3.Text
       Case 1
K1 = Format(Date, "YYYY") + "-" + "01" + "-01"
K2 = Format(Date, "YYYY") + "-" + "01" + "-31"
       Case 2
If Val(Format(Date, "YYYY")) / 4 = Int(Val(Format(Date, "YYYY")) / 4) Then
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-29"
Else
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-28"
End If
       Case 3
K1 = Format(Date, "YYYY") + "-" + "03" + "-01"
K2 = Format(Date, "YYYY") + "-" + "03" + "-31"
       Case 4
K1 = Format(Date, "YYYY") + "-" + "04" + "-01"
K2 = Format(Date, "YYYY") + "-" + "04" + "-30"
       Case 5
K1 = Format(Date, "YYYY") + "-" + "05" + "-01"
K2 = Format(Date, "YYYY") + "-" + "05" + "-31"
       Case 6
K1 = Format(Date, "YYYY") + "-" + "06" + "-01"
K2 = Format(Date, "YYYY") + "-" + "06" + "-30"
       Case 7
K1 = Format(Date, "YYYY") + "-" + "07" + "-01"
K2 = Format(Date, "YYYY") + "-" + "07" + "-31"
       Case 8
K1 = Format(Date, "YYYY") + "-" + "08" + "-01"
K2 = Format(Date, "YYYY") + "-" + "08" + "-30"
       Case 9
K1 = Format(Date, "YYYY") + "-" + "09" + "-01"
K2 = Format(Date, "YYYY") + "-" + "09" + "-31"
       Case 10
K1 = Format(Date, "YYYY") + "-" + "10" + "-01"
K2 = Format(Date, "YYYY") + "-" + "10" + "-30"
       Case 11
K1 = Format(Date, "YYYY") + "-" + "11" + "-01"
K2 = Format(Date, "YYYY") + "-" + "11" + "-31"
       Case 12
K1 = Format(Date, "YYYY") + "-" + "12" + "-01"
K2 = Format(Date, "YYYY") + "-" + "12" + "-30"
End Select
End Sub



