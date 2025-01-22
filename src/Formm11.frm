VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formm11 
   BackColor       =   &H00C0E0FF&
   Caption         =   "单机备活"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "Formm11.frx":0000
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   77
      Text            =   "Text7"
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "复制原锅号"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "毛坯库存"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   5280
      Width           =   1815
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Left            =   12960
      TabIndex        =   72
      Top             =   6000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo3"
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0000C0C0&
      Caption         =   "返修"
      Height          =   375
      Left            =   2880
      TabIndex        =   71
      Top             =   1920
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯"
      Height          =   375
      Left            =   1680
      TabIndex        =   70
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染色"
      Height          =   375
      Left            =   480
      TabIndex        =   69
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0FF&
      Caption         =   "计划取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H008080FF&
      Caption         =   "条码预览"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data Data25 
      Caption         =   "Data25"
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   4800
      TabIndex        =   65
      Text            =   "Text12"
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H008080FF&
      Caption         =   "条码打印"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H008080FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   61
      Text            =   "Text8"
      Top             =   960
      Width           =   2415
   End
   Begin VB.Data Data24 
      Caption         =   "Data24"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Data Data23 
      Caption         =   "Data23"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "订单刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "计划结束"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Data Data22 
      Caption         =   "Data22"
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
      Width           =   3135
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   13200
      TabIndex        =   55
      Text            =   "Text3"
      Top             =   5160
      Width           =   495
   End
   Begin VB.Data Data21 
      Caption         =   "Data21"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   4320
      TabIndex        =   54
      Text            =   "Text12"
      Top             =   8400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Formm11.frx":440A
      Left            =   13800
      List            =   "Formm11.frx":441A
      TabIndex        =   53
      Text            =   "Combo2"
      Top             =   5160
      Width           =   975
   End
   Begin VB.Data Data20 
      Caption         =   "Data20"
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
      Top             =   -120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data19 
      Caption         =   "Data19"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formm11.frx":4436
      Height          =   1695
      Left            =   4200
      TabIndex        =   51
      Top             =   3480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   8
      BackColorFixed  =   8421631
      BackColorBkg    =   40863
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   2880
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H008080FF&
      Caption         =   "新锅号"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4680
      Width           =   975
   End
   Begin VB.Data Data18 
      Caption         =   "Data18"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formm11.frx":444B
      Height          =   330
      Left            =   10800
      TabIndex        =   30
      Top             =   5160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "车台编号"
      Text            =   ""
   End
   Begin VB.Data Data17 
      Caption         =   "Data17"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data16 
      Caption         =   "Data16"
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
      Top             =   -120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data15 
      Caption         =   "Data15"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CKGL"
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "打印预览"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   -240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   -120
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
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   29
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11760
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   240
      Top             =   0
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   855
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "订单详情"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   "ok"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text9"
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Text            =   "Text10"
      Top             =   10560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
      Bindings        =   "Formm11.frx":4460
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Bindings        =   "Formm11.frx":4474
      Height          =   330
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "xm"
      Text            =   "DBCombo5"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1440
      TabIndex        =   33
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   83558401
      CurrentDate     =   36892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Bindings        =   "Formm11.frx":4489
      Height          =   2055
      Left            =   480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7440
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632319
      BackColorBkg    =   34952
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "记录号"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formm11.frx":449D
      Height          =   1695
      Left            =   4920
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5640
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   11
      BackColor       =   16777215
      BackColorFixed  =   12632319
      BackColorBkg    =   8421440
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   83558401
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   83558401
      CurrentDate     =   39177
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formm11.frx":44B1
      Height          =   330
      Left            =   1560
      TabIndex        =   20
      Top             =   2400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   1
      Left            =   7440
      TabIndex        =   21
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   2
      Left            =   10800
      TabIndex        =   22
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   3
      Left            =   11760
      TabIndex        =   23
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   4
      Left            =   12720
      TabIndex        =   24
      Top             =   4080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   5
      Left            =   13560
      TabIndex        =   25
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   6
      Left            =   9480
      TabIndex        =   26
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   7
      Left            =   9480
      TabIndex        =   27
      Top             =   5160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formm11.frx":44C5
      Height          =   330
      Index           =   8
      Left            =   7440
      TabIndex        =   28
      Top             =   5160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo4"
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1440
      TabIndex        =   31
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   83558401
      CurrentDate     =   36892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formm11.frx":44DA
      Height          =   1935
      Left            =   4200
      TabIndex        =   58
      Top             =   1440
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   8
      BackColorFixed  =   8421631
      BackColorBkg    =   40863
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   375
      Left            =   11760
      TabIndex        =   79
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   83558401
      CurrentDate     =   39177
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "色号-色别"
      Height          =   375
      Left            =   9480
      TabIndex        =   80
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "机台"
      Height          =   375
      Index           =   16
      Left            =   10800
      TabIndex        =   78
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请输入原锅号"
      Height          =   855
      Index           =   15
      Left            =   2400
      TabIndex        =   76
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "毛坯备注"
      Height          =   375
      Index           =   14
      Left            =   12960
      TabIndex        =   73
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作月份"
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   68
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请输入合约号"
      Height          =   375
      Index           =   13
      Left            =   480
      TabIndex        =   62
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP"
      Height          =   375
      Index           =   12
      Left            =   13200
      TabIndex        =   56
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "整理类别"
      Height          =   375
      Index           =   11
      Left            =   13800
      TabIndex        =   52
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入单号"
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   50
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "说明：为了保证作业的准确，建议输入计划作业单号！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   48
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   375
      Index           =   8
      Left            =   9480
      TabIndex        =   47
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "技术要求"
      Height          =   375
      Index           =   7
      Left            =   7440
      TabIndex        =   46
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "重量（公斤）"
      Height          =   375
      Index           =   4
      Left            =   13560
      TabIndex        =   45
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "匹数"
      Height          =   375
      Index           =   3
      Left            =   12720
      TabIndex        =   44
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光胚幅宽(cm)"
      Height          =   375
      Index           =   2
      Left            =   11760
      TabIndex        =   43
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "毛胚幅宽(寸)"
      Height          =   375
      Index           =   1
      Left            =   10800
      TabIndex        =   42
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "品名"
      Height          =   375
      Index           =   6
      Left            =   7440
      TabIndex        =   41
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   40
      Top             =   840
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   7080
      X2              =   8280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   39
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户名称"
      Height          =   375
      Index           =   8
      Left            =   480
      TabIndex        =   38
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   37
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   36
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "日期"
      Height          =   375
      Index           =   9
      Left            =   11760
      TabIndex        =   35
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请输入锅号"
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   34
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择负责人"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   32
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Formm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X As Integer: Public BI As Integer ''''BI PANDUAN CHURU KU BIANLIANG
Dim BA As Database: Dim rr As Integer: Public gh, K1, K2 As String: Public hg As Date: Dim BA3 As Database: Dim RD3 As Recordset
Public ZL As Single  ''''''重量变量
Rem ' 中间转换变量
Dim rs As Single: Dim RD1 As Recordset: Dim BA1 As Database: Public ll, c, r As Integer: Public lbj As Long
Dim RD As Recordset: Public mm As Date: Public ML As Date: Dim BA2 As Database: Dim RD2 As Recordset

Private Sub Command10_Click()
If DBCombo8.Text = "" Then
MsgBox ("请输入订单号")
Exit Sub
End If
If MsgBox("确定订单结束吗？" + DBCombo8.Text, vbYesNo) = vbNo Then Exit Sub
Data22.Database.Execute "UPDATE SCZY_Z SET 排产='Y' WHERE 单号='" & DBCombo8.Text & "'"
End Sub

Private Sub Command12_Click()
On Error Resume Next
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox ("请选择备活信息")
Exit Sub
End If

If Option1.Value = True Then
Data23.RecordSource = "select MAX(VAL(MID(kpd.锅号,8))) as h  from kpd where 日期=CDATE('" & Text6.Text & "') "
Data23.Refresh
Text7.Text = Format(CDate(Text6.Text), "YYMMDD") + "-1"
If Data23.Recordset.EOF Then
Text7.Text = Format(CDate(Text6.Text), "YYMMDD") + "-1"
Else
Text7.Text = Format(CDate(Text6.Text), "YYMMDD") + "-" + Trim(Data23.Recordset.Fields(0) + 1)
End If

  Text3.Text = 1
  Data9.RecordSource = "select count(kpd.Ip) as bj from kpd where kpd.锅号= '" & Text7.Text & "' "
  Data9.Refresh
  Text3.Text = Data9.Recordset.Fields(0) + 1
End If

If Option2.Value = True Then

Data23.RecordSource = "select MAX(VAL(MID(kpd.锅号,8))) as h  from kpd where 日期=CDATE('" & Text6.Text & "')  AND INSTR(锅号,'W')>0"
Data23.Refresh
Text7.Text = Format(CDate(Text6.Text), "YYMMDD") + "W1"
If Data23.Recordset.EOF Then
Text7.Text = Format(CDate(Text6.Text), "YYMMDD") + "W1"
Else
Text7.Text = Format(CDate(Text6.Text), "YYMMDD") + "W" + Trim(Data23.Recordset.Fields(0) + 1)
End If
  Text3.Text = 1
  Data9.RecordSource = "select count(kpd.Ip) as bj from kpd where kpd.锅号= '" & Text7.Text & "' "
  Data9.Refresh
  Text3.Text = Data9.Recordset.Fields(0) + 1
End If

If Option3.Value = True Then

Data23.RecordSource = "select MAX(VAL(MID(kpd.锅号,8))) as h  from kpd where 日期=CDATE('" & Text6.Text & "')  AND INSTR(锅号,'F')>0"
Data23.Refresh
Text7.Text = Format(CDate(Text6.Text), "YYMMDD") + "F1"
If Data23.Recordset.EOF Then
Text7.Text = Format(CDate(Text6.Text), "YYMMDD") + "F1"
Else
Text7.Text = Format(CDate(Text6.Text), "YYMMDD") + "F" + Trim(Data23.Recordset.Fields(0) + 1)
End If
  Text3.Text = 1
  Data9.RecordSource = "select count(kpd.Ip) as bj from kpd where kpd.锅号= '" & Text7.Text & "' "
  Data9.Refresh
  Text3.Text = Data9.Recordset.Fields(0) + 1
End If
End Sub


Private Sub Command13_Click()
On Error Resume Next
Data7.RecordSource = "SELECT 款号与和约号 FROM sczy_Z WHERE 单号='" & DBCombo8.Text & "'"
Data7.Refresh
Text9.Text = ""
If Not Data7.Recordset.EOF Then
Text9.Text = Data7.Recordset.Fields(0)
End If
If DBCombo8.Text = "" Then Exit Sub
Data20.Database.Execute "delete * from ZHXH4"     ''''''清空数据表
Data19.Recordset.FindFirst "单号='" & DBCombo8.Text & "' "
If Data19.Recordset.NoMatch Then
Data20.RecordSource = "SELECT ZHXH4.色别,format(SUM(ZHXH4.计划量),'#0.0') AS 计划量,format(SUM(ZHXH4.备活量),'#0.0') AS 备活量 FROM ZHXH4 WHERE ZHXH4.单号='" & DBCombo8.Text & "' GROUP BY ZHXH4.色别 "
Data20.Refresh
Exit Sub
Else
Data20.Recordset.FindFirst "单号='" & DBCombo8.Text & "'"
If Data20.Recordset.NoMatch Then
Data19.Database.Execute "INSERT INTO  ZHXH4(单号,品名,色别,毛坯幅宽,光坯幅宽,计划量) IN 'd:\数据库\bfrz\" + LJB + "\ZHXH.mdB'  SELECT sczy_X.单号,SCZY_X.名称,SCZY_X.色别,SCZY_X.毛胚幅宽,SCZY_X.光胚幅宽,SCZY_X.计划量 From sczy_x WHERE SCZY_X.单号='" & DBCombo8.Text & "'"
Data20.Database.Execute "UPDATE ZHXH4 SET ZHXH4.备活量=0 where ZHXH4.单号='" & DBCombo8.Text & "'"
Data20.RecordSource = "SELECT ZHXH4.色别,format(SUM(ZHXH4.计划量),'#0.0') AS 计划量,format(SUM(ZHXH4.备活量),'#0.0') AS 备活量 FROM ZHXH4 WHERE ZHXH4.单号='" & DBCombo8.Text & "' GROUP BY ZHXH4.色别 "
Data20.Refresh
Else
Data20.RecordSource = "SELECT ZHXH4.色别,format(SUM(ZHXH4.计划量),'#0.0') AS 计划量,format(SUM(ZHXH4.备活量),'#0.0') AS 备活量 FROM ZHXH4 WHERE ZHXH4.单号='" & DBCombo8.Text & "' GROUP BY ZHXH4.色别 "
Data20.Refresh
Exit Sub
End If
End If
Data20.Database.Execute "DELETE * FROM ZHXH4 WHERE ZHXH4.备活量>0 AND 单号='" & DBCombo8.Text & "'"
Data8.Database.Execute "INSERT INTO  ZHXH4(单号,品名,色别,毛坯幅宽,光坯幅宽,备活量) IN 'd:\数据库\bfrz\" + LJB + "\ZHXH.mdB'  SELECT KPD.单号,KPD.品名,KPD.色别,KPD.毛胚幅宽,KPD.光胚幅宽,KPD.重量 From KPD WHERE KPD.单号='" & DBCombo8.Text & "'"
Data20.Database.Execute "UPDATE ZHXH4 SET ZHXH4.计划量=0 where ZHXH4.备活量>0 AND 单号='" & DBCombo8.Text & "'"
Data20.RecordSource = "SELECT ZHXH4.色别,format(SUM(ZHXH4.计划量),'#0.0') AS 计划量,format(SUM(ZHXH4.备活量),'#0.0') AS 备活量 FROM ZHXH4 WHERE ZHXH4.单号='" & DBCombo8.Text & "' GROUP BY ZHXH4.色别 "
Data20.Refresh

End Sub


Private Sub Command14_Click()
Call lcd(Data13, Text7.Text)
End Sub

Private Sub Command15_Click()
Call ZJTMDY("*" + Trim(Text7.Text) + "J" + "*", Text7.Text)
End Sub

Private Sub Command16_Click()
Call TMDY("*" + Trim(Text7.Text) + "J" + "*", Text7.Text)
End Sub

Private Sub Command17_Click()
If DBCombo8.Text = "" Then
MsgBox ("请输入订单号")
Exit Sub
End If
If MsgBox("确定取消结束吗？" + DBCombo8.Text, vbYesNo) = vbNo Then Exit Sub
Data22.Database.Execute "UPDATE SCZY_Z SET 排产='N' WHERE 单号='" & DBCombo8.Text & "'"
End Sub


Private Sub Command2_Click()
Command2.Enabled = False
Call mpkc7
Command2.Enabled = True
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If Text11.Text = "" Then
MsgBox ("请输入原锅号")
Exit Sub
End If


If Text7.Text = "" Then
MsgBox ("请输入锅号")
Exit Sub
End If

If MsgBox("要复制原锅号" + Text11.Text + "新锅号为" + Text7.Text + "吗？", vbYesNo) = vbNo Then Exit Sub
Data10.Database.Execute "insert into kpd(客户名称,单号,锅号,色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,日期,备注,技术要求,IP,标签,kp,kp1,CKY,负责人,pb,rs,ts,xdx,ddx,fh) select 客户名称,单号,'" & Text7.Text & "',色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,'" & Date & "',备注,技术要求,IP,标签,'N','N',CKY,负责人,'Y','N','N','N','N','N' from kpd where 锅号='" & Text11.Text & "'"
Data8.RecordSource = "select kpd.客户名称,kpd.锅号,kpd.IP,kpd.品名,kpd.毛胚幅宽,kpd.光胚幅宽,匹数,kpd.重量,kpd.色别,kpd.标签 as 合约号,kpd.备注,kpd.技术要求,类别,CKY as 毛坯备注,车台  from kpd where kpd.锅号='" & Text7.Text & "' order by val(ip)"
Data8.Refresh

End Sub

Private Sub Command5_Click()
On Error Resume Next

Call lcd2(Data6, Text7.Text)

'Data6.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"
'Data6.RecordSource = "select max(kpd.重量) as zl from kpd where kpd.锅号='" & Text7.Text & "'"
'Data6.Refresh
'a = Data6.Recordset.Fields("zl")

'Data6.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"
'Data6.RecordSource = "select * from kpd where kpd.锅号='" & Text7.Text & "' And kpd.重量 = VAL('" & a & "')  "
'Data6.Refresh
'On Error Resume Next
'b = Data6.Recordset.Fields("ip")



'DataEnvironment1.kp Text7.Text, b
'DataReport1.Show 1
'DataEnvironment1.rskp.Close
End Sub

Private Sub Command7_Click()
Form12.Text1.Text = DBCombo8.Text
Form12.Show
End Sub



Private Sub Command1_Click()
On Error Resume Next
If DBCombo5.Text = "" Then
MsgBox ("请选择负责人！")
Exit Sub
End If

If DBCombo1.Text = "" Then
MsgBox ("请输入客户！")
Exit Sub
End If

RD2.FindFirst "ip='" & Text3.Text & "' AND 锅号='" & Text7.Text & "'"
   If RD2.NoMatch Then
If DBCombo1.Text = "" Then
MsgBox ("wrong")
Exit Sub
End If
If Text3.Text = "" Then Text3.Text = 1
RD2.AddNew
RD2.Fields(0) = DBCombo1.Text
RD2.Fields(1) = DBCombo8.Text
RD2.Fields(2) = Text7.Text
RD2.Fields(11) = Text3.Text
RD2.Fields(12) = Text6.Text
RD2.Fields(13) = Text9.Text
RD2.Fields(14) = DBCombo2.Text
RD2.Fields(15) = DBCombo5.Text
RD2.Fields(16) = DBCombo3.Text
RD2.Fields(17) = "N"
RD2.Fields(18) = "N"
RD2.Fields(19) = Combo2.Text
RD2.Fields(20) = "0"
RD2.Fields(21) = "Y"
RD2.Fields(22) = "N"
RD2.Fields(23) = "N"
RD2.Fields(24) = "N"
RD2.Fields(25) = "N"
RD2.Fields(26) = "N"


For i = 1 To 8
RD2.Fields(i + 2) = DBCombo4(i).Text
Next
RD2.Update



Rem 'shuaxin 开票单

Data8.RecordSource = "select kpd.客户名称,kpd.锅号,kpd.IP,kpd.品名,kpd.毛胚幅宽,kpd.光胚幅宽,匹数,kpd.重量,kpd.色别,kpd.标签 as 合约号,kpd.备注,kpd.技术要求,类别,CKY as 毛坯备注,车台  from kpd where kpd.锅号='" & Text7.Text & "' "
Data8.Refresh

Data8.RecordSource = "select kpd.客户名称,kpd.锅号,kpd.IP,kpd.品名,kpd.毛胚幅宽,kpd.光胚幅宽,匹数,kpd.重量,kpd.色别,kpd.标签 as 合约号,kpd.备注,kpd.技术要求  from kpd where kpd.锅号='" & Text7.Text & "' order by val(ip)"
Data8.Refresh

If Not Data8.Recordset.EOF Then
Data8.Recordset.MoveLast
MSFlexGrid4.TextMatrix(0, 0) = "记录号"
For i = 1 To Data8.Recordset.RecordCount
MSFlexGrid4.TextMatrix(i, 0) = i
Next
End If

Else
MsgBox ("此IP号已存在，禁止重复！")
DBCombo3.SetFocus
End If
  
  Text3.Text = 1
  Data9.RecordSource = "select count(kpd.Ip) as bj from kpd where kpd.锅号= '" & Text7.Text & "' "
  Data9.Refresh
  Text3.Text = Data9.Recordset.Fields(0) + 1

End Sub


Private Sub Command8_Click()

Rem 'shuaxin 开票单
If Text1.Text = "" Or Text2.Text = "" Then
Data8.RecordSource = "select * from kpd "
Data8.Refresh
Else
RQ = CDate(Text2.Text)

Data8.RecordSource = "select kpd.客户名称,kpd.锅号,kpd.IP,kpd.品名,kpd.毛胚幅宽,kpd.光胚幅宽,匹数,kpd.重量,kpd.色别,kpd.标签 as 合约号,kpd.备注,kpd.技术要求  from kpd where kpd.锅号='" & Text7.Text & "' and kpd.日期 between # " & CDate(Text1.Text) & "#  and   # " & CDate(RQ) & " # and kpd.客户名称='" & DBCombo1.Text & "'order by kpd.日期 desc"
Data8.Refresh
End If
If Data8.Recordset.EOF Then
MSFlexGrid4.TextMatrix(0, 0) = "记录号"
GoTo 200
End If
Data8.Recordset.MoveLast
MSFlexGrid4.TextMatrix(0, 0) = "记录号"
For i = 1 To Data8.Recordset.RecordCount
MSFlexGrid4.TextMatrix(i, 0) = i
Next


Rem ' shuaxin 仓库单
200:



300:

  Data9.RecordSource = "select count(kpd.Ip) as bj from kpd where kpd.锅号= '" & Text7.Text & "' "
  Data9.Refresh
  Text3.Text = Data9.Recordset.Fields(0) + 1



Text7.Enabled = True

End Sub

Private Sub Command6_Click()
On Error Resume Next
RQ = Text2.Text
If DBCombo1.Text <> "" Then
Data8.RecordSource = "select kpd.客户名称,kpd.锅号,kpd.IP,kpd.品名,kpd.毛胚幅宽,kpd.光胚幅宽,匹数,kpd.重量,kpd.色别,kpd.标签 as 合约号,kpd.备注,kpd.技术要求  from kpd where kpd.日期 between CDate(# " & Text1.Text & "#)  and   CDate(' " & RQ & " ') and kpd.客户名称='" & DBCombo1.Text & "'order by kpd.日期,val(mid(锅号,9))"
Data8.Refresh
Else
Data8.RecordSource = "select kpd.客户名称,kpd.锅号,kpd.IP,kpd.品名,kpd.毛胚幅宽,kpd.光胚幅宽,匹数,kpd.重量,kpd.色别,kpd.标签 as 合约号,kpd.备注,kpd.技术要求  from kpd where kpd.日期 between CDate(# " & Text1.Text & "#)  and   CDate(' " & RQ & " ') order by kpd.日期,val(mid(锅号,9))"
Data8.Refresh
End If
End Sub

Private Sub Command9_Click()
On Error Resume Next
If Text3.Text = "" Then Exit Sub
If MsgBox("确定删除" + Text3.Text + "吗？", vbYesNo) = vbNo Then Exit Sub
Data8.Recordset.Delete
Data8.Refresh
Text3.Text = 1
Data9.RecordSource = "select count(kpd.Ip) as bj from kpd where kpd.锅号= '" & Text7.Text & "' "
Data9.Refresh
Text3.Text = Data9.Recordset.Fields(0) + 1
End Sub

Private Sub DBCombo1_Change()
On Error Resume Next
 ww = 0
If Text4.Text = "" Or Text5.Text = "" Then
Exit Sub
End If
RQ = CDate(Text5.Text)
 Data9.RecordSource = "select count(kpd.Ip) as bj from kpd where kpd.锅号= '" & Text7.Text & "' "
  Data9.Refresh
  Text3.Text = Data9.Recordset.Fields(0) + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''订单计划信息
Data22.RecordSource = "select 客户名称,sczy_z.单号,sczy_z.胚布名称,sczy_z.生产规格,sczy_z.总备注,sczy_z.克重,sczy_z.款号与和约号 from sczy_z where sczy_z.日期 between CDate(' " & Text4.Text & "')  and  CDate( '" & Text5.Text & "')  and sczy_z.客户名称='" & DBCombo1.Text & "' and 排产<>'Y'  order by sczy_z.日期 "
Data22.Refresh


End Sub

Private Sub DBCombo1_Click(Area As Integer)
' On Error Resume Next

 ww = 0
If Text4.Text = "" Or Text5.Text = "" Then
End If
RQ = CDate(Text5.Text)
op = 0.5

  Data9.RecordSource = "select count(kpd.Ip) as bj from kpd where kpd.锅号= '" & Text7.Text & "' "
  Data9.Refresh
  Text3.Text = Data9.Recordset.Fields(0) + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''订单计划信息
Data22.RecordSource = "select 客户名称,sczy_z.单号,sczy_z.胚布名称,sczy_z.生产规格,sczy_z.总备注,sczy_z.克重,sczy_z.款号与和约号 from sczy_z where sczy_z.日期 between CDate(' " & Text4.Text & "')  and  CDate( '" & Text5.Text & "')  and sczy_z.客户名称='" & DBCombo1.Text & "' and 排产<>'Y' order by sczy_z.日期 "
Data22.Refresh

End Sub

Private Sub DBCOMBO1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub DBCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub DBCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub DBCombo4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub dbcombo5_Change()
On Error Resume Next
Data22.RecordSource = "select 客户名称,sczy_z.单号,sczy_z.胚布名称,sczy_z.生产规格,sczy_z.总备注,sczy_z.克重,sczy_z.款号与和约号 from sczy_z where sczy_z.日期 between CDate('" & Text4.Text & "')  and  CDate('" & Text5.Text & "')  and 排产<>'Y'  order by sczy_z.日期 "
Data22.Refresh
End Sub

Private Sub DBCombo5_Click(Area As Integer)
Data22.RecordSource = "select 客户名称,sczy_z.单号,sczy_z.胚布名称,sczy_z.生产规格,sczy_z.总备注,sczy_z.克重,sczy_z.款号与和约号 from sczy_z where sczy_z.日期 between CDate('" & Text4.Text & "')  and  CDate('" & Text5.Text & "')  and 排产<>'Y'  order by sczy_z.日期 "
Data22.Refresh
End Sub

Private Sub DBCombo5_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DBCombo6_Change()
On Error Resume Next
DBCombo4(5).Enabled = False
Label14.Caption = "锁定状态,双击解锁"
If BI = 1 Then
Data13.RecordSource = "select ckgl.IP,ckgl.布类,ckgl.存放位置,ckgl.毛胚幅宽,ckgl.毛胚重量,ckgl.实际投放量,ckgl.毛胚匹数,ckgl.备注 from ckgl  where ckgl.客户名称='" & DBCombo1.Text & " ' and ckgl.布类='" & DBCombo4(1).Text & " ' and ckgl.毛胚幅宽='" & DBCombo4(2).Text & " ' and ckgl.  ckgl.存放位置='" & DBCombo6.Text & "' and ckgl.毛胚重量>ckgl.实际投放量 "
Data13.Refresh
Else
       Data13.RecordSource = "select ckgl.IP,ckgl.布类,ckgl.存放位置,ckgl.毛胚幅宽,ckgl.毛胚重量,ckgl.实际投放量,ckgl.毛胚匹数,ckgl.备注 from ckgl  where ckgl.客户名称='" & DBCombo1.Text & " ' and ckgl.布类='" & DBCombo4(1).Text & " ' and ckgl.毛胚幅宽='" & DBCombo4(2).Text & " ' and ckgl.   ckgl.毛胚重量>ckgl.实际投放量 "
       Data13.Refresh
       End If
Text11.Text = Data13.Recordset.Fields("IP")
DBCombo7.Text = Data13.Recordset.Fields("毛胚重量") - Data13.Recordset.Fields("实际投放量")
End Sub


Private Sub DBCombo6_DblClick(Area As Integer)
On Error Resume Next
DBCombo4(5).Enabled = False
If BI = 1 Then
Data13.RecordSource = "select ckgl.IP,ckgl.布类,ckgl.存放位置,ckgl.毛胚幅宽,ckgl.毛胚重量,ckgl.实际投放量,ckgl.毛胚匹数,ckgl.备注 from ckgl  where ckgl.客户名称='" & DBCombo1.Text & " ' and ckgl.布类='" & DBCombo4(1).Text & " ' and ckgl.毛胚幅宽='" & DBCombo4(2).Text & " ' and ckgl.  ckgl.存放位置='" & DBCombo6.Text & "' and ckgl.毛胚重量>ckgl.实际投放量"
Data13.Refresh
Else
       Data13.RecordSource = "select ckgl.IP,ckgl.布类,ckgl.存放位置,ckgl.毛胚幅宽,ckgl.毛胚重量,ckgl.实际投放量,ckgl.毛胚匹数,ckgl.备注 from ckgl  where ckgl.客户名称='" & DBCombo1.Text & " ' and ckgl.布类='" & DBCombo4(1).Text & " ' and ckgl.毛胚幅宽='" & DBCombo4(2).Text & " ' and ckgl.   ckgl.毛胚重量>ckgl.实际投放量"
       Data13.Refresh
       End If
Text11.Text = Data13.Recordset.Fields("IP")
DBCombo7.Text = Data13.Recordset.Fields("毛胚重量") - Data13.Recordset.Fields("实际投放量")
End Sub

Private Sub DBCombo6_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub DBCombo7_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub DBCombo8_KeyDown(KeyCode As Integer, Shift As Integer)
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
Text1.Text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = DTPicker3.Value
Text1.SetFocus
End Sub

Private Sub DTPicker4_Change()
Text2.Text = DTPicker4.Value
End Sub

Private Sub DTPicker4_CloseUp()
Text2.Text = DTPicker4.Value
Text2.SetFocus
End Sub

Private Sub DTPicker5_Change()
Text6.Text = DTPicker5.Value
End Sub

Private Sub DTPicker5_CloseUp()
Text6.Text = DTPicker5.Value
End Sub
Private Sub Form_Load()
On Error Resume Next
Me.Caption = Me.Caption + "操作年度： " + LJB
Set BA1 = OpenDatabase("d:\数据库\bfrz\" + LJB + "\JH.MDB")
Set RD1 = BA1.OpenRecordset("KPD", dbOpenDynaset)
Set BA = OpenDatabase("d:\数据库\bfrz\" + LJB + "\ckgl.MDB")
Set RD = BA.OpenRecordset("ckgl", dbOpenDynaset)
Set BA3 = OpenDatabase("d:\数据库\bfrz\" + LJB + "\jh.MDB")
Set RD3 = BA3.OpenRecordset("jh", dbOpenDynaset)

Set BA2 = OpenDatabase("d:\数据库\bfrz\" + LJB + "\ckgl.MDB")
Set RD2 = BA2.OpenRecordset("kpd", dbOpenDynaset)
DBCombo8.Text = ""
Combo2.Text = "圆筒"
Data3.DatabaseName = "d:\数据库\bfrz\" + LJB + "\sczyjhd.mdb"
Data3.RecordSource = "select 简称  from khzl  group by 简称"
Data3.Refresh

Text3.Text = ""
Text7.Text = ""
Text9.Text = ""
DBCombo3.Text = ""
Text11.Text = ""
DTPicker5.Value = Date
Data12.DatabaseName = "d:\数据库\bfrz\" + LJB + "\sczyjhd.mdb"
Data12.RecordSource = "select xm  from ywf group by xm"
Data12.Refresh


Data1.DatabaseName = "d:\数据库\bfrz\" + LJB + "\jh.mdb"
Data10.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"
Data11.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"


Data19.DatabaseName = "d:\数据库\bfrz\" + LJB + "\SCZYJHD.mdb"
Data19.RecordSource = "SCZY_X"
Data19.Refresh

Data20.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ZHXH.mdb"
Data20.Refresh

Data13.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"

Data14.DatabaseName = "d:\数据库\bfrz\" + LJB + "\jh.mdb"

Data15.DatabaseName = "d:\数据库\bfrz\" + LJB + "\CKGL.mdb"

Data16.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"

Data17.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"

Data18.DatabaseName = "d:\数据库\bfrz\" + LJB + "\JH.mdb"
Data18.RecordSource = "SELECT CT.车台编号 FROM CT GROUP BY CT.车台编号"
Data18.Refresh

Data21.DatabaseName = "d:\数据库\bfrz\" + LJB + "\SCJD.mdb"
Data22.DatabaseName = "d:\数据库\bfrz\" + LJB + "\SCZYJHD.mdb"


Data2.DatabaseName = "d:\数据库\bfrz\" + LJB + "\jh.mdb"

Data23.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"
Data24.DatabaseName = "d:\数据库\bfrz\" + LJB + "\CJBB.mdb"



Data4.DatabaseName = "d:\数据库\bfrz\" + LJB + "\jh.mdb"

Data5.DatabaseName = "d:\数据库\bfrz\" + LJB + "\jh.mdb"

Data6.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"

Data7.DatabaseName = "d:\数据库\bfrz\" + LJB + "\sczyjhd.mdb"

Data8.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"

Data9.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"


Data25.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"
Data25.RecordSource = "select MC from JSYQ group by MC"
Data25.Refresh


Text6.Text = Date
Text8.Text = Month(Text6.Text)


DBCombo5.Text = ""


DBCombo4(4).Enabled = True
DBCombo4(5).Enabled = True
Text1.Text = Date
Text2.Text = Date

BI = 1 ''''''''''设置出入库状态为出库



DBCombo7.Text = ""

For i = 1 To 8
DBCombo4(i).Text = ""
Next




MSFlexGrid3.ColWidth(0) = 100
MSFlexGrid3.ColWidth(1) = 1000
MSFlexGrid3.ColWidth(2) = 1800
MSFlexGrid3.ColWidth(3) = 1200

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1000
MSFlexGrid1.ColWidth(3) = 1000

MSFlexGrid2.ColWidth(0) = 100
MSFlexGrid2.ColWidth(1) = 1500
MSFlexGrid2.ColWidth(2) = 1500
MSFlexGrid2.ColWidth(3) = 1500
MSFlexGrid2.ColWidth(4) = 1500
MSFlexGrid2.ColWidth(5) = 1500
MSFlexGrid2.ColWidth(6) = 1200
MSFlexGrid2.ColWidth(7) = 1500


MSFlexGrid4.ColWidth(0) = 100
MSFlexGrid4.ColWidth(2) = 1500
MSFlexGrid4.ColWidth(3) = 500
MSFlexGrid4.ColWidth(4) = 1600
MSFlexGrid4.ColWidth(8) = 1000
MSFlexGrid4.ColWidth(9) = 1800

ZL = 0

DTPicker1.Value = Date - 30
DTPicker2.Value = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
DBCombo1.Text = ""
Text11.Text = ""
Text4.Text = Date - 30
Text5.Text = Date
Text4.TabIndex = 0
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label10_Click()

End Sub



Private Sub Label1_Click()
Form38.Text1.Text = DBCombo4(6).Text
Form38.Show
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 6
Form113.Text1.Text = DBCombo4(1).Text
Form113.Show
       Case 8
beizhu = 11
Form112.Show
End Select
End Sub

Private Sub Label9_Click()
Form18.Text1.Text = DBCombo2.Text
Form18.Show
End Sub

Private Sub MSFlexGrid1_Click()
On Error Resume Next
If Data20.Recordset.EOF Then Exit Sub
rs = MSFlexGrid1.Row
Data20.Recordset.MoveFirst
Data20.Recordset.Move rs - 1
DBCombo4(6).Text = Data20.Recordset.Fields(0)

End Sub

Private Sub Label14_DblClick()
DBCombo4(5).Enabled = True
End Sub

Private Sub Label15_DblClick()
Label12.Caption = Format(DBCombo4(5).Text, "###0.00")
       Combo1.Text = "入库"
       
       BI = 0
       Data13.RecordSource = "select ckgl.客户名称,ckgl.布类,ckgl.存放位置,ckgl.毛胚幅宽,ckgl.毛胚重量,ckgl.实际投放量,ckgl.毛胚匹数,ckgl.备注,CKGL.IP from ckgl  where ckgl.客户名称='" & DBCombo1.Text & " ' and ckgl.布类='" & DBCombo4(1).Text & " ' and ckgl.毛胚幅宽='" & DBCombo4(2).Text & " ' and ckgl.   VAL(ckgl.毛胚重量)-VAL(CKGL.实际投放量)>=0 AND VAL(CKGL.实际投放量)>0 order by ckgl.存放位置"
       Data13.Refresh
End Sub

Private Sub MSFlexGrid2_Click()
On Error Resume Next
If Data22.Recordset.EOF Then
DBCombo2.Text = ""
Exit Sub
End If
rs = MSFlexGrid2.Row
Data22.Recordset.MoveFirst
Data22.Recordset.Move rs - 1
DBCombo8.Text = Data22.Recordset.Fields(1)
Text9.Text = Data22.Recordset.Fields(6)
End Sub

Private Sub MSFlexGrid3_Click()
On Error Resume Next
If Data6.Recordset.EOF Then Exit Sub
rs = MSFlexGrid3.Row
Data6.Recordset.MoveFirst
Data6.Recordset.Move rs - 1
DBCombo4(1).Text = Data6.Recordset.Fields(1)
DBCombo4(2).Text = Data6.Recordset.Fields(2)
DBCombo3.Text = Data6.Recordset.Fields(5)
DBCombo4(4).Text = Data6.Recordset.Fields(3)
DBCombo4(5).Text = Data6.Recordset.Fields(4)
End Sub

Private Sub MSFlexGrid4_DblClick()
If Data8.Recordset.EOF Then Exit Sub
rs = MSFlexGrid4.Row
Data8.Recordset.MoveFirst
Data8.Recordset.Move rs - 1
Text3.Text = Data8.Recordset.Fields(2)
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Text7_Change()
On Error Resume Next
Data8.RecordSource = "select kpd.客户名称,kpd.锅号,kpd.IP,kpd.品名,kpd.毛胚幅宽,kpd.光胚幅宽,匹数,kpd.重量,kpd.色别,kpd.标签 as 合约号,kpd.备注,kpd.技术要求,类别,CKY as 毛坯备注,车台  from kpd where kpd.锅号='" & Text7.Text & "' order by val(ip)"
Data8.Refresh
  Text3.Text = 1
  Data9.RecordSource = "select count(kpd.Ip) as bj from kpd where kpd.锅号= '" & Text7.Text & "' "
  Data9.Refresh
  Text3.Text = Data9.Recordset.Fields(0) + 1
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub SX()
If Data20.Recordset.EOF Then Exit Sub
Data20.Recordset.MoveFirst
i = 1
Do While Not Data20.Recordset.EOF
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = i
MSFlexGrid1.Text = Format(Data20.Recordset.Fields(2), "##0.0")
Data20.Recordset.MoveNext
i = i + 1
Loop

End Sub

Private Sub MSFlex()
With MSFlexGrid4
    c = .Col: r = .Row    '''''C列，，R行
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub


Private Sub MSFlexGrid4_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid4.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid4.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid4.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data8.Recordset.MoveFirst
Data8.Recordset.Move r - 1
Data8.Recordset.Edit
Data8.Recordset.Fields(c - 1) = Text1111.Text
Data8.Recordset.Update

Text1111.Visible = False
MSFlexGrid4.SetFocus
End Sub

Private Sub mpkc7()
If DBCombo1.Text = "" Then
MsgBox ("请输入客户")
Exit Sub
End If

Data6.DatabaseName = "d:\数据库\bfrz\" + LJB + "\ckgl.mdb"
Data6.Database.Execute "delete * from mpkc7"
Data6.Database.Execute "insert into mpkc7(客户名称,品名,匹数,重量,备注) select 客户名称,布类,毛胚匹数,毛胚重量,备注 from ckgl where 客户名称='" & DBCombo1.Text & "' "
'Data6.Database.Execute "insert into mpkc7(客户名称,品名,匹数,重量,备注) select 客户名称,品名,匹数,重量,备注 from kcjl where 客户名称='" & DBCombo1.Text & "'"
Data6.Database.Execute "update mpkc7 set 库类='1'"
Data6.Database.Execute "insert into mpkc7(客户名称,品名,匹数,重量,备注) select 客户名称,品名,匹数,重量,CKY from kpd where 客户名称='" & DBCombo1.Text & "' AND instr(锅号,'F')=0 and instr(锅号,'H')=0 "
Data6.Database.Execute "update mpkc7 set 匹数=-匹数,重量=-重量,库类='-1' where 库类=null"
Data6.Database.Execute "insert into mpkc7(客户名称,品名,匹数,重量,备注) select 客户名称,布类,毛胚匹数,毛胚重量,ny from chk where 客户名称='" & DBCombo1.Text & "' "
Data6.Database.Execute "update mpkc7 set 匹数=-匹数,重量=-重量,库类='-1' where 库类=null"
Data6.Database.Execute "insert into mpkc7(客户名称,品名,备注,匹数,重量) select 客户名称,品名,备注,sum(匹数),format(sum(重量),'#0.0') from mpkc7 group by 客户名称,品名,备注"
Data6.Database.Execute "update mpkc7 set 库类='0' where 库类=null"
Data6.Database.Execute "delete * from mpkc7 where 库类<>'0'"
Data6.RecordSource = "select * from mpkc7 WHERE 重量<>0 and  instr(备注,'c')=0 order by 客户名称,品名"
Data6.Refresh
End Sub
