VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Formd111111 
   BackColor       =   &H00C0E0FF&
   Caption         =   "大货配方单"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form23"
   MDIChild        =   -1  'True
   ScaleHeight     =   10590
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "取消退出"
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   3720
      Width           =   975
   End
   Begin VB.Data Data16 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data15 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data14 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   103
      Text            =   "Text4"
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   102
      Text            =   "Text4"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   101
      Text            =   "Text4"
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   100
      Text            =   "Text4"
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   99
      Text            =   "Text4"
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   98
      Text            =   "Text4"
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   97
      Text            =   "Text4"
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "日期确认"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "浴比确认"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   6360
      TabIndex        =   94
      Text            =   "Text3"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   -120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   11760
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3720
      TabIndex        =   83
      Top             =   4440
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "曲线打印"
      Height          =   375
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Left            =   11040
      TabIndex        =   80
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
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
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选择打印"
      Height          =   375
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   -120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "责任修改"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
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
      Top             =   -120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "配方单"
      Height          =   3495
      Left            =   3600
      TabIndex        =   25
      Top             =   240
      Width           =   11535
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   6
         Left            =   5800
         TabIndex        =   92
         Text            =   "Text3"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   5
         Left            =   5800
         TabIndex        =   91
         Text            =   "Text3"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   4
         Left            =   5800
         TabIndex        =   90
         Text            =   "Text3"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   3
         Left            =   5800
         TabIndex        =   89
         Text            =   "Text3"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   2
         Left            =   5800
         TabIndex        =   88
         Text            =   "Text3"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   1
         Left            =   5800
         TabIndex        =   87
         Text            =   "Text3"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   0
         Left            =   5800
         TabIndex        =   85
         Text            =   "Text3"
         Top             =   720
         Width           =   615
      End
      Begin MSDBCtls.DBCombo DBCombo4 
         Bindings        =   "Formd111111.frx":0000
         Height          =   330
         Left            =   1200
         TabIndex        =   79
         Top             =   2880
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         BackColor       =   12648447
         ListField       =   "工艺编号"
         Text            =   "DBCombo4"
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   6
         Left            =   8640
         TabIndex        =   75
         Text            =   "Text2"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   5
         Left            =   8640
         TabIndex        =   74
         Text            =   "Text2"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   4
         Left            =   8640
         TabIndex        =   73
         Text            =   "Text2"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   3
         Left            =   8640
         TabIndex        =   72
         Text            =   "Text2"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   2
         Left            =   8640
         TabIndex        =   71
         Text            =   "Text2"
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   1
         Left            =   8640
         TabIndex        =   70
         Text            =   "Text2"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   0
         Left            =   8640
         TabIndex        =   69
         Text            =   "Text2"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   6
         Left            =   7200
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   5
         Left            =   7200
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   4
         Left            =   7200
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   3
         Left            =   7200
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   1880
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   2
         Left            =   7200
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   1
         Left            =   7200
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   0
         Left            =   7200
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "Formd111111.frx":0015
         Height          =   330
         Index           =   0
         Left            =   6480
         TabIndex        =   55
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "dw"
         Text            =   "DBCombo3"
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Formd111111.frx":0029
         Height          =   330
         Index           =   0
         Left            =   3360
         TabIndex        =   48
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "染料名称"
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formd111111.frx":003D
         Height          =   360
         Index           =   4
         Left            =   1200
         TabIndex        =   8
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         BackColor       =   12648447
         ListField       =   "工艺工序"
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   5
         Left            =   1200
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formd111111.frx":0051
         Height          =   360
         Index           =   6
         Left            =   1200
         TabIndex        =   10
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         BackColor       =   12648447
         ListField       =   "染化助库名"
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formd111111.frx":0065
         Height          =   360
         Index           =   7
         Left            =   9600
         TabIndex        =   12
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         BackColor       =   12648447
         ListField       =   "染料名称"
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formd111111.frx":0079
         Height          =   360
         Index           =   8
         Left            =   9360
         TabIndex        =   13
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "dw"
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   9
         Left            =   10080
         TabIndex        =   14
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   10
         Left            =   10800
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   11
         Left            =   9480
         TabIndex        =   16
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formd111111.frx":008D
         Height          =   360
         Index           =   12
         Left            =   1200
         TabIndex        =   7
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "工艺编号"
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formd111111.frx":00A1
         Height          =   360
         Index           =   13
         Left            =   1200
         TabIndex        =   11
         Top             =   2400
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "标志"
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   16
         Left            =   9480
         TabIndex        =   41
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   17
         Left            =   9480
         TabIndex        =   44
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   18
         Left            =   9480
         TabIndex        =   45
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Formd111111.frx":00B6
         Height          =   330
         Index           =   1
         Left            =   3360
         TabIndex        =   49
         Top             =   1110
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "染料名称"
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Formd111111.frx":00CA
         Height          =   330
         Index           =   2
         Left            =   3360
         TabIndex        =   50
         Top             =   1500
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "染料名称"
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Formd111111.frx":00DE
         Height          =   330
         Index           =   3
         Left            =   3360
         TabIndex        =   51
         Top             =   1880
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "染料名称"
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Formd111111.frx":00F2
         Height          =   330
         Index           =   4
         Left            =   3360
         TabIndex        =   52
         Top             =   2250
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "染料名称"
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Formd111111.frx":0106
         Height          =   330
         Index           =   5
         Left            =   3360
         TabIndex        =   53
         Top             =   2620
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "染料名称"
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Formd111111.frx":011A
         Height          =   330
         Index           =   6
         Left            =   3360
         TabIndex        =   54
         Top             =   3000
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "染料名称"
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "Formd111111.frx":012E
         Height          =   330
         Index           =   1
         Left            =   6480
         TabIndex        =   56
         Top             =   1080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "dw"
         Text            =   "DBCombo3"
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "Formd111111.frx":0142
         Height          =   330
         Index           =   2
         Left            =   6480
         TabIndex        =   57
         Top             =   1500
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "dw"
         Text            =   "DBCombo3"
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "Formd111111.frx":0156
         Height          =   330
         Index           =   3
         Left            =   6480
         TabIndex        =   58
         Top             =   1920
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "dw"
         Text            =   "DBCombo3"
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "Formd111111.frx":016A
         Height          =   330
         Index           =   4
         Left            =   6480
         TabIndex        =   59
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "dw"
         Text            =   "DBCombo3"
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "Formd111111.frx":017E
         Height          =   330
         Index           =   5
         Left            =   6480
         TabIndex        =   60
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "dw"
         Text            =   "DBCombo3"
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "Formd111111.frx":0192
         Height          =   330
         Index           =   6
         Left            =   6480
         TabIndex        =   61
         Top             =   3000
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "dw"
         Text            =   "DBCombo3"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   19
         Left            =   1200
         TabIndex        =   93
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "批次"
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
         Index           =   20
         Left            =   5805
         TabIndex        =   86
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "常规工艺号"
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
         Index           =   19
         Left            =   120
         TabIndex        =   78
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "压力"
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
         Index           =   18
         Left            =   9480
         TabIndex        =   47
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "车速"
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
         Index           =   17
         Left            =   9480
         TabIndex        =   46
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "次序号"
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
         Index           =   16
         Left            =   9480
         TabIndex        =   42
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "工序名称"
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
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "浴比"
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
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "染化助名称"
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
         Left            =   3360
         TabIndex        =   36
         Top             =   240
         Width           =   2415
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
         Index           =   6
         Left            =   6480
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "配方"
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
         Left            =   7200
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "工艺日期"
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
         Index           =   8
         Left            =   9480
         TabIndex        =   33
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "配方编号"
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
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "校正值"
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
         Left            =   8640
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "染化助代码"
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
         Left            =   120
         TabIndex        =   29
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "染化助库"
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
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "客户信息"
      Height          =   3495
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   3255
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   3
         Left            =   1200
         TabIndex        =   3
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   14
         Left            =   1200
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   360
         Index           =   15
         Left            =   1200
         TabIndex        =   5
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
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
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "生产类别"
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
         Left            =   120
         TabIndex        =   40
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "负责人"
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
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "客户名称"
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "品名"
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
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "色号"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "颜色 "
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
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "调整确认"
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "下一编号"
      Height          =   375
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   -120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formd111111.frx":01A6
      Height          =   5535
      Left            =   240
      TabIndex        =   31
      Top             =   4320
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   9763
      _Version        =   393216
      BackColorFixed  =   8421631
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查看"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "工艺曲线"
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
      Left            =   10200
      TabIndex        =   81
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Formd111111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S1, S2 As Integer: Public c, r As Integer
Dim BA As Database: Dim RD As Recordset: Dim sz(56) As String

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Command10_Click()
'Call TPOutAdodcToExcel(BJ, "配方编号：" + Trim(dataCombo1(12).Text), "E:\" + Trim(dataCombo5.Text) + ".BMP")
End Sub

Private Sub Command12_Click()
If DataCombo1(12).Text = "" Then
MsgBox ("没有配方编号")
Exit Sub
End If
If DataCombo1(5).Text = "" Then
MsgBox ("请输入浴比")
Exit Sub
End If
Adodc7.Database.Execute "UPDATE PFD1 SET 浴比='" & DataCombo1(5).Text & "' WHERE 配方编号='" & DataCombo1(12).Text & "'"
Adodc7.RecordSource = "SELECT * FROM PFD1 WHERE PFD1.配方编号='" & DataCombo1(12).Text & "'ORDER BY val(PFD1.工序名称),次序号"
Adodc7.Refresh
       If Adodc7.Recordset.EOF Then
       DataCombo1(16).Text = 1
       Else
       DataCombo1(16).Text = Adodc7.Recordset.RecordCount + 1
       End If
End Sub

Private Sub Command13_Click()
If DataCombo1(12).Text = "" Then
MsgBox ("没有配方编号")
Exit Sub
End If
If DataCombo1(11).Text = "" Then
MsgBox ("请输入配方日期")
Exit Sub
End If
Adodc7.Database.Execute "UPDATE PFD1 SET 配方日期='" & DataCombo1(11).Text & "' WHERE 配方编号='" & DataCombo1(12).Text & "'"
Adodc7.RecordSource = "SELECT * FROM PFD1 WHERE PFD1.配方编号='" & DataCombo1(12).Text & "'ORDER BY val(PFD1.工序名称),次序号"
Adodc7.Refresh
       If Adodc7.Recordset.EOF Then
       DataCombo1(16).Text = 1
       Else
       DataCombo1(16).Text = Adodc7.Recordset.RecordCount + 1
       End If
End Sub

Private Sub Command14_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If DataCombo1(0).Text = "" Or DataCombo1(1).Text = "" Or DataCombo1(2).Text = "" Or DataCombo1(3).Text = "" Or DataCombo1(12).Text = "" Then
MsgBox ("客户、品名、色号、颜色、负责人、生产类别、配方编号须填完整！")
Exit Sub
End If

For i = 0 To 3
DataCombo1(i).Enabled = False
Next
DataCombo1(10).Enabled = False
DataCombo1(11).Enabled = False
DataCombo1(12).Enabled = False

DataCombo1(13).Text = ""        '''''''''''''代码清离

For i = 0 To 6     '''''''''''''''''''''''''
If Text1(i).Text <> "" Then
DataCombo1(7).Text = DataCombo2(i).Text
DataCombo1(8).Text = DataCombo3(i).Text
DataCombo1(9).Text = Text1(i).Text
DataCombo1(10).Text = Text2(i).Text
DataCombo1(17).Text = Text4(i).Text
DataCombo1(19).Text = Text3(i).Text
Adodc6.Recordset.AddNew
For P = 0 To Adodc6.Recordset.Fields.Count - 1
Adodc6.Recordset.Fields(P) = DataCombo1(P).Text
Next
Adodc6.Recordset.Fields(16) = Adodc7.Recordset.RecordCount + 1
Adodc6.Recordset.Update
Adodc7.Refresh
End If
Next
                '''''''''''''''''''''''
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = 1
Text3(i).Text = ""
Text4(i).Text = ""
Next
DataCombo1(16).Enabled = False
DataCombo1(16).Text = Adodc7.Recordset.RecordCount + 1
DataCombo1(4).SetFocus
End Sub

Private Sub Command3_Click()
If DataCombo1(0).Text = "" Or DataCombo1(1).Text = "" Or DataCombo1(2).Text = "" Or DataCombo1(3).Text = "" Or DataCombo1(12).Text = "" Then
MsgBox ("客户、品名、色号、颜色、负责人、生产类别、配方编号须填完整！")
Exit Sub
End If

Adodc7.Recordset.Edit
DataCombo1(7).Text = DataCombo2(0).Text
DataCombo1(8).Text = DataCombo3(0).Text
DataCombo1(9).Text = Text1(0).Text
DataCombo1(10).Text = Text2(0).Text
DataCombo1(17).Text = Text4(0).Text
DataCombo1(19).Text = Text3(0).Text
For i = 0 To RD.Fields.Count - 1
Adodc7.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc7.Recordset.Update
Adodc7.Refresh
For i = 0 To 3
DataCombo1(i).Enabled = False
Next
DataCombo1(16).Enabled = False
DataCombo1(16).Text = Adodc7.Recordset.RecordCount + 1
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = 1
Text3(i).Text = ""
Text4(i).Text = ""
Next
End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc7.Recordset.Delete
Adodc7.Refresh
For i = 0 To 3
DataCombo1(i).Enabled = False
Next

DataCombo1(16).Text = Adodc7.Recordset.RecordCount + 1
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = 1
Text3(i).Text = ""
Text4(i).Text = ""
Next

DataCombo1(0).SetFocus
End Sub

Private Sub Command6_Click()
Adodc14.RecordSource = "select distinct 工序名称 from pfd1 where 配方编号='" & DataCombo1(12).Text & "' order by 工序名称"
Adodc14.Refresh
If Adodc14.Recordset.EOF Then Exit Sub
If MsgBox("确定生成吗？", vbYesNo) = vbNo Then Exit Sub
Adodc14.Recordset.MoveFirst
i = 0
sz(i) = DataCombo1(0).Text
i = i + 1
sz(i) = DataCombo1(1).Text
i = i + 1
sz(i) = DataCombo1(2).Text
i = i + 1
sz(i) = DataCombo1(3).Text
i = i + 1
sz(i) = DataCombo1(12).Text
i = i + 1
sz(i) = DataCombo1(14).Text
i = i + 1
sz(i) = DataCombo1(11).Text
i = i + 1

Do While Not Adodc14.Recordset.EOF
Adodc15.RecordSource = "select 工序名称,浴比,染化助库,染化助名称,单位,配方,车速 from pfd1 where 配方编号='" & DataCombo1(12).Text & "' and 工序名称='" & Adodc14.Recordset.Fields(0) & "' order by 次序号"
Adodc15.Refresh

If Not Adodc15.Recordset.EOF Then
Adodc15.Recordset.MoveFirst
Do While Not Adodc15.Recordset.EOF
sz(i) = Adodc15.Recordset.Fields(0) + "(" + Adodc15.Recordset.Fields(1) + ")" + Adodc15.Recordset.Fields(2) + "-" + Adodc15.Recordset.Fields(3) + "\" + Adodc15.Recordset.Fields(4) + "#" + Adodc15.Recordset.Fields(5) + "^" + Adodc15.Recordset.Fields(6)
i = i + 1
Adodc15.Recordset.MoveNext
Loop
End If

Adodc14.Recordset.MoveNext
Loop

If i < 57 Then
For L = i To 56
sz(L) = ""
Next
End If

Adodc16.RecordSource = "select 编号 from pfd where 编号='" & DataCombo1(12).Text & "'"
Adodc16.Refresh
If Adodc16.Recordset.EOF Then
Adodc16.Database.Execute "INSERT INTO pfd(客户,品名,色号,颜色,编号,技术,日期,N1,N2,N3,N4,N5,N6,N7,N8,N9,N10,N11,N12,N13,N14,N15,N16,N17,N18,N19,N20,N21,N22,N23,N24,N25,N26,N27,N28,N29,N30,N31,N32,N33,N34,N35,N36,N37,N38,N39,N40,N41,N42,N43,N44,N45,N46,N47,N48,N49,N50) VALUES('" & sz(0) & "','" & sz(1) & "','" & sz(2) & "','" & sz(3) & "','" & sz(4) & "','" & sz(5) & "',cdate('" & sz(6) & "'),'" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & sz(10) & "',  " & _
                                                                        "'" & sz(11) & "','" & sz(12) & "','" & sz(13) & "','" & sz(14) & "','" & sz(15) & "','" & sz(16) & "','" & sz(17) & "','" & sz(18) & "','" & sz(19) & "','" & sz(20) & "','" & sz(21) & "', " & _
                                                                        "'" & sz(22) & "','" & sz(23) & "','" & sz(24) & "','" & sz(25) & "','" & sz(26) & "','" & sz(27) & "','" & sz(28) & "','" & sz(29) & "','" & sz(30) & "','" & sz(31) & "','" & sz(32) & "', " & _
                                                                        "'" & sz(33) & "','" & sz(34) & "','" & sz(35) & "','" & sz(36) & "','" & sz(37) & "','" & sz(38) & "','" & sz(39) & "','" & sz(40) & "','" & sz(41) & "','" & sz(42) & "','" & sz(43) & "', " & _
                                                                        "'" & sz(44) & "','" & sz(45) & "','" & sz(46) & "','" & sz(47) & "','" & sz(48) & "','" & sz(49) & "','" & sz(50) & "','" & sz(51) & "','" & sz(52) & "','" & sz(53) & "','" & sz(54) & "', " & _
                                                                        "'" & sz(55) & "','" & sz(56) & "')"
Else
Adodc16.Database.Execute "delete * from pfd where 编号='" & DataCombo1(12).Text & "'"
Adodc16.Database.Execute "INSERT INTO pfd(客户,品名,色号,颜色,编号,技术,日期,N1,N2,N3,N4,N5,N6,N7,N8,N9,N10,N11,N12,N13,N14,N15,N16,N17,N18,N19,N20,N21,N22,N23,N24,N25,N26,N27,N28,N29,N30,N31,N32,N33,N34,N35,N36,N37,N38,N39,N40,N41,N42,N43,N44,N45,N46,N47,N48,N49,N50) VALUES('" & sz(0) & "','" & sz(1) & "','" & sz(2) & "','" & sz(3) & "','" & sz(4) & "','" & sz(5) & "',cdate('" & sz(6) & "'),'" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & sz(10) & "',  " & _
                                                                        "'" & sz(11) & "','" & sz(12) & "','" & sz(13) & "','" & sz(14) & "','" & sz(15) & "','" & sz(16) & "','" & sz(17) & "','" & sz(18) & "','" & sz(19) & "','" & sz(20) & "','" & sz(21) & "', " & _
                                                                        "'" & sz(22) & "','" & sz(23) & "','" & sz(24) & "','" & sz(25) & "','" & sz(26) & "','" & sz(27) & "','" & sz(28) & "','" & sz(29) & "','" & sz(30) & "','" & sz(31) & "','" & sz(32) & "', " & _
                                                                        "'" & sz(33) & "','" & sz(34) & "','" & sz(35) & "','" & sz(36) & "','" & sz(37) & "','" & sz(38) & "','" & sz(39) & "','" & sz(40) & "','" & sz(41) & "','" & sz(42) & "','" & sz(43) & "', " & _
                                                                        "'" & sz(44) & "','" & sz(45) & "','" & sz(46) & "','" & sz(47) & "','" & sz(48) & "','" & sz(49) & "','" & sz(50) & "','" & sz(51) & "','" & sz(52) & "','" & sz(53) & "','" & sz(54) & "', " & _
                                                                        "'" & sz(55) & "','" & sz(56) & "')"
End If



Unload Me

End Sub

Private Sub Command8_Click()
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 4
       Adodc12.RecordSource = "SELECT 工艺编号 FROM CGGY WHERE INSTR('" & DataCombo1(4).Text & "',工艺名称)>0 GROUP BY 工艺编号"
       Adodc12.Refresh
       Case 12
       Adodc7.RecordSource = "SELECT * FROM PFD1 WHERE PFD1.配方编号='" & DataCombo1(12).Text & "'ORDER BY val(PFD1.工序名称),次序号"
       Adodc7.Refresh

       If Adodc7.Recordset.EOF Then
       DataCombo1(16).Text = 1
       Else
       DataCombo1(16).Text = Adodc7.Recordset.RecordCount + 1
       End If
       
       For i = 0 To 3
       DataCombo1(i).Text = Adodc7.Recordset.Fields(i)
       Next
       DataCombo1(14).Text = Adodc7.Recordset.Fields(14)
       Case 6
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Adodc10.RecordSource = "SELECT 标志 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 标志 "
       Adodc10.Refresh
       Case 13
       If DataCombo1(13).Text = "" Then
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Else
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH  where 染化助库名='" & DataCombo1(6).Text & "' AND INSTR(标志,'" & DataCombo1(13).Text & "')>0 GROUP BY 染料名称"
       Adodc8.Refresh
       End If
End Select
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 4
       Adodc12.RecordSource = "SELECT 工艺编号 FROM CGGY WHERE INSTR('" & DataCombo1(4).Text & "',工艺名称)>0 GROUP BY 工艺编号"
       Adodc12.Refresh

       Case 12
       Adodc7.RecordSource = "SELECT * FROM PFD1 WHERE PFD1.配方编号='" & DataCombo1(12).Text & "'ORDER BY val(PFD1.工序名称),次序号"
       Adodc7.Refresh

       If Adodc7.Recordset.EOF Then
        DataCombo1(16).Text = 1
       Else
         DataCombo1(16).Text = Adodc7.Recordset.RecordCount + 1
       End If
       
       For i = 0 To 3
       DataCombo1(i).Text = Adodc7.Recordset.Fields(i)
       Next
       DataCombo1(14).Text = Adodc7.Recordset.Fields(14)
       
       Case 6
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH GROUP BY 染料名称 where 染化助库名='" & DataCombo1(6).Text & "'"
       Adodc8.Refresh
       Adodc10.RecordSource = "SELECT 标志 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 标志 "
       Adodc10.Refresh
       Case 13
       If DataCombo1(13).Text = "" Then
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Else
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH  where 染化助库名='" & DataCombo1(6).Text & "' AND INSTR(标志,'" & DataCombo1(13).Text & "')>0 GROUP BY 染料名称"
       Adodc8.Refresh
       End If
End Select
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode

End Sub


Private Sub dataCombo4_Click(Area As Integer)
On Error Resume Next
For i = 0 To 6
DataCombo2(i).Text = ""
Text1(i).Text = ""
Next
Adodc13.RecordSource = "SELECT * FROM CGGY WHERE 工艺编号='" & DataCombo4.Text & "' AND INSTR('" & DataCombo1(4).Text & "',工艺名称)>0 ORDER BY 序号"
Adodc13.Refresh
If Adodc13.Recordset.EOF Then
For i = 0 To 6
Text1(i).Text = ""
Next
Else
Adodc13.Recordset.MoveFirst
i = 0
Do While Not Adodc13.Recordset.EOF
DataCombo1(6).Text = Adodc13.Recordset.Fields(2)
DataCombo1(13).Text = Adodc13.Recordset.Fields(3)
DataCombo2(i).Text = Adodc13.Recordset.Fields(4)
DataCombo3(i).Text = Adodc13.Recordset.Fields(5)
Text1(i).Text = Adodc13.Recordset.Fields(6)
Text4(i).Text = Adodc13.Recordset.Fields(8)
i = i + 1
Adodc13.Recordset.MoveNext
Loop
End If
End Sub


Private Sub Form_Load()
Me.Caption = Me.Caption + "操作年度 " + ljb
'On Error Resume Next
Dim L As String
Set BA = OpenDatabase("d:\excel\枣庄\DB.mdb")
Set RD = BA.OpenRecordset("PFD1", dbOpenDynaset)
mb = 1
For i = 0 To RD.Fields.Count - 1
DataCombo1(i) = ""
Next
Timer1.Enabled = False
ProgressBar1.Visible = False
For i = 0 To 3
DataCombo1(i).Enabled = False
Next
DataCombo4.Text = ""
DataCombo5.Text = ""
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = 1
Text3(i).Text = ""
Text4(i).Text = ""
Next

Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False

DataCombo1(10).Text = 1
DataCombo1(11).Text = Date
'dataCombo1(11).Enabled = False
DataCombo1(12).Enabled = False
DataCombo1(14).Enabled = False
DataCombo1(15).Enabled = False

Adodc1.DatabaseName = "d:\excel\枣庄\DB.mdb"


Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"


Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc4.RecordSource = "select 编号,工艺工序 from gx group by 编号,工艺工序 ORDER BY VAL(工艺工序)"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc5.RecordSource = "select dw,IP from dw group by dw,IP ORDER BY IP"
Adodc5.Refresh


Adodc6.DatabaseName = "d:\excel\枣庄\DB.mdb"
Adodc6.RecordSource = "PFD1"
Adodc6.Refresh

Adodc7.DatabaseName = "d:\excel\枣庄\DB.mdb"
Adodc7.RecordSource = "SELECT * FROM PFD1 WHERE PFD1.配方编号='" & DataCombo1(12).Text & "' ORDER BY val(PFD1.工序名称),PFD1.次序号"
Adodc7.Refresh

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc8.RecordSource = "SELECT 染化助库名 FROM RHZH GROUP BY 染化助库名"
Adodc8.Refresh

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc9.RecordSource = "SELECT 染化助库名 FROM RHZH GROUP BY 染化助库名"
Adodc9.Refresh

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc10.RecordSource = "SELECT 标志 FROM RHZH GROUP BY 标志"
Adodc10.Refresh

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc11.RecordSource = "SELECT 负责人姓名 FROM GR GROUP BY 负责人姓名"
Adodc11.Refresh

Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc12.RecordSource = "SELECT 工艺编号 FROM CGGY GROUP BY 工艺编号"
Adodc12.Refresh

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc13.Refresh

If Adodc7.Recordset.EOF Then
DataCombo1(16).Text = 1
Else
DataCombo1(16).Text = Adodc7.Recordset.RecordCount + 1
End If

Adodc14.DatabaseName = "d:\excel\枣庄\DB.mdb"
Adodc15.DatabaseName = "d:\excel\枣庄\DB.mdb"
Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"

DataCombo1(0).TabIndex = 0

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(4) = 0
VSFlexGrid1.ColWidth(5) = 1000
VSFlexGrid1.ColWidth(6) = 600
VSFlexGrid1.ColWidth(7) = 1500
VSFlexGrid1.ColWidth(8) = 2000
VSFlexGrid1.ColWidth(9) = 1000
VSFlexGrid1.ColWidth(10) = 1500
VSFlexGrid1.ColWidth(11) = 600
VSFlexGrid1.ColWidth(12) = 1200
VSFlexGrid1.ColWidth(15) = 0
VSFlexGrid1.ColWidth(16) = 0
VSFlexGrid1.ColWidth(18) = 1600
VSFlexGrid1.ColWidth(19) = 0
VSFlexGrid1.ColWidth(20) = 600
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 16
       DataCombo1(16).Enabled = True
       Case 12
       DataCombo1(12).Enabled = True
       Case 11
       DataCombo1(10).Enabled = True
       Case 8
       DataCombo1(11).Enabled = True
       Case 9
       DataCombo1(12).Enabled = True
End Select
End Sub
Private Sub vSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
S1 = VSFlexGrid1.RowSel
End Sub

Private Sub vSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
S2 = VSFlexGrid1.RowSel
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc7.Recordset.MoveFirst
Adodc7.Recordset.Move rs - 1
For i = 0 To RD.Fields.Count - 1
If i <> 14 Then
DataCombo1(i).Text = Adodc7.Recordset.Fields(i)
End If
Next
DataCombo2(0).Text = DataCombo1(7).Text
DataCombo3(0).Text = DataCombo1(8).Text
Text1(0).Text = DataCombo1(9).Text
Text2(0).Text = DataCombo1(10).Text
Text3(0).Text = DataCombo1(19).Text
Text4(0).Text = DataCombo1(17).Text
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub vSFlexGrid1_Dbl()
With VSFlexGrid1
    c = .Col: r = .Row    '''''C列，，R行
If c = 9 Or c = 10 Or c = 11 Or c = 18 Or c = 20 Then
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End If
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call vSFlexGrid1_Dbl
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc7.Recordset.MoveFirst
Adodc7.Recordset.Move r - 1
Adodc7.Recordset.Edit
Adodc7.Recordset.Fields(c - 1) = Text1111.Text
Adodc7.Recordset.Update
VSFlexGrid1.Text = Text1111.Text
Text1111.Visible = False
VSFlexGrid1.SetFocus
End If
End Sub


