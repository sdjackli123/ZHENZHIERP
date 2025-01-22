VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy22 
   BackColor       =   &H00C0E0FF&
   Caption         =   "定量分析---线类"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form22"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   4080
      TabIndex        =   192
      Text            =   "Text4"
      Top             =   960
      Width           =   1815
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
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
      Width           =   5535
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":0000
      Height          =   330
      Index           =   0
      Left            =   5040
      TabIndex        =   157
      Top             =   3480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":0015
      Height          =   330
      Index           =   0
      Left            =   4200
      TabIndex        =   156
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   13680
      TabIndex        =   153
      Text            =   "Text6"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   13680
      TabIndex        =   152
      Text            =   "Text6"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   13560
      TabIndex        =   151
      Text            =   "Text6"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Left            =   7560
      TabIndex        =   150
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo2"
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7560
      TabIndex        =   148
      Text            =   "Text5"
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0E0FF&
      Height          =   1935
      Left            =   6840
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   146
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4455
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
      Top             =   10800
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   840
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   6480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "Formy22.frx":002A
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   2295
      Left            =   9600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Formy22.frx":0030
      Top             =   720
      Width           =   5535
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
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
      Left            =   1080
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
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
      Top             =   10680
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
      Left            =   120
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
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
      Width           =   4935
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":0036
      Height          =   330
      Index           =   0
      Left            =   3360
      TabIndex        =   10
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   0
      Left            =   5760
      TabIndex        =   11
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":004A
      Height          =   330
      Index           =   0
      Left            =   2520
      TabIndex        =   12
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":005F
      Height          =   330
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   0
      Left            =   960
      TabIndex        =   14
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy22.frx":0073
      Height          =   2415
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7800
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   4260
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
      Bindings        =   "Formy22.frx":0087
      Height          =   330
      Index           =   1
      Left            =   960
      TabIndex        =   16
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "款号"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy22.frx":009C
      Height          =   330
      Index           =   2
      Left            =   960
      TabIndex        =   17
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      BackColor       =   12648447
      ListField       =   "颜色"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   3
      Left            =   11880
      TabIndex        =   18
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   4
      Left            =   2640
      TabIndex        =   19
      Top             =   -120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   5
      Left            =   7920
      TabIndex        =   20
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   6
      Left            =   8040
      TabIndex        =   21
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   7
      Left            =   8040
      TabIndex        =   22
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   8
      Left            =   7800
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   9
      Left            =   8040
      TabIndex        =   24
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   10
      Left            =   7800
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   11
      Left            =   7920
      TabIndex        =   26
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   12
      Left            =   7800
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   13
      Left            =   7920
      TabIndex        =   28
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   14
      Left            =   7920
      TabIndex        =   29
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   15
      Left            =   7920
      TabIndex        =   30
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   16
      Left            =   9840
      TabIndex        =   31
      Top             =   8400
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   17
      Left            =   9840
      TabIndex        =   32
      Top             =   8400
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   18
      Left            =   10320
      TabIndex        =   33
      Top             =   8400
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   19
      Left            =   10680
      TabIndex        =   34
      Top             =   8160
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":00B0
      Height          =   330
      Index           =   1
      Left            =   1080
      TabIndex        =   35
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":00C4
      Height          =   330
      Index           =   2
      Left            =   1080
      TabIndex        =   36
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":00D8
      Height          =   330
      Index           =   3
      Left            =   1080
      TabIndex        =   37
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":00EC
      Height          =   330
      Index           =   4
      Left            =   1080
      TabIndex        =   38
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":0100
      Height          =   330
      Index           =   5
      Left            =   1080
      TabIndex        =   39
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":0114
      Height          =   330
      Index           =   6
      Left            =   1080
      TabIndex        =   40
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":0128
      Height          =   330
      Index           =   7
      Left            =   1080
      TabIndex        =   41
      Top             =   6840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":013C
      Height          =   330
      Index           =   8
      Left            =   9360
      TabIndex        =   42
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":0150
      Height          =   330
      Index           =   9
      Left            =   9360
      TabIndex        =   43
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":0164
      Height          =   330
      Index           =   10
      Left            =   9360
      TabIndex        =   44
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":0178
      Height          =   330
      Index           =   1
      Left            =   2520
      TabIndex        =   45
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":018D
      Height          =   330
      Index           =   2
      Left            =   2520
      TabIndex        =   46
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":01A2
      Height          =   330
      Index           =   3
      Left            =   2520
      TabIndex        =   47
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":01B7
      Height          =   330
      Index           =   4
      Left            =   2520
      TabIndex        =   48
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":01CC
      Height          =   330
      Index           =   5
      Left            =   2520
      TabIndex        =   49
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":01E1
      Height          =   330
      Index           =   6
      Left            =   2520
      TabIndex        =   50
      Top             =   6360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":01F6
      Height          =   330
      Index           =   7
      Left            =   2520
      TabIndex        =   51
      Top             =   6840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":020B
      Height          =   330
      Index           =   8
      Left            =   10800
      TabIndex        =   52
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":0220
      Height          =   330
      Index           =   9
      Left            =   10800
      TabIndex        =   53
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":0235
      Height          =   330
      Index           =   10
      Left            =   10800
      TabIndex        =   54
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   1
      Left            =   5760
      TabIndex        =   55
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   2
      Left            =   5760
      TabIndex        =   56
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   3
      Left            =   5760
      TabIndex        =   57
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   4
      Left            =   5760
      TabIndex        =   58
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   5
      Left            =   5760
      TabIndex        =   59
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   6
      Left            =   5760
      TabIndex        =   60
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   7
      Left            =   5760
      TabIndex        =   61
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   8
      Left            =   14160
      TabIndex        =   62
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   9
      Left            =   14160
      TabIndex        =   63
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   10
      Left            =   14160
      TabIndex        =   64
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":024A
      Height          =   330
      Index           =   11
      Left            =   9360
      TabIndex        =   65
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":025E
      Height          =   330
      Index           =   11
      Left            =   10800
      TabIndex        =   66
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   11
      Left            =   14160
      TabIndex        =   67
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":0273
      Height          =   330
      Index           =   1
      Left            =   3360
      TabIndex        =   68
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":0287
      Height          =   330
      Index           =   2
      Left            =   3360
      TabIndex        =   69
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":029B
      Height          =   330
      Index           =   3
      Left            =   3360
      TabIndex        =   70
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":02AF
      Height          =   330
      Index           =   4
      Left            =   3360
      TabIndex        =   71
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":02C3
      Height          =   330
      Index           =   5
      Left            =   3360
      TabIndex        =   72
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":02D7
      Height          =   330
      Index           =   6
      Left            =   3360
      TabIndex        =   73
      Top             =   6360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":02EB
      Height          =   330
      Index           =   7
      Left            =   3360
      TabIndex        =   74
      Top             =   6840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":02FF
      Height          =   330
      Index           =   8
      Left            =   11640
      TabIndex        =   75
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":0313
      Height          =   330
      Index           =   9
      Left            =   11640
      TabIndex        =   76
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":0327
      Height          =   330
      Index           =   10
      Left            =   11640
      TabIndex        =   77
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":033B
      Height          =   330
      Index           =   11
      Left            =   11640
      TabIndex        =   78
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":034F
      Height          =   330
      Index           =   12
      Left            =   11640
      TabIndex        =   79
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   12
      Left            =   14160
      TabIndex        =   80
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":0363
      Height          =   330
      Index           =   12
      Left            =   10800
      TabIndex        =   81
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":0378
      Height          =   330
      Index           =   12
      Left            =   9360
      TabIndex        =   82
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   20
      Left            =   8760
      TabIndex        =   83
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   21
      Left            =   8640
      TabIndex        =   84
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   22
      Left            =   8640
      TabIndex        =   85
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   23
      Left            =   8640
      TabIndex        =   86
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":038C
      Height          =   330
      Index           =   13
      Left            =   9360
      TabIndex        =   87
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":03A0
      Height          =   330
      Index           =   14
      Left            =   9360
      TabIndex        =   88
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":03B4
      Height          =   330
      Index           =   15
      Left            =   9360
      TabIndex        =   89
      Top             =   6840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":03C8
      Height          =   330
      Index           =   13
      Left            =   10800
      TabIndex        =   90
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":03DD
      Height          =   330
      Index           =   14
      Left            =   10800
      TabIndex        =   91
      Top             =   6360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":03F2
      Height          =   330
      Index           =   15
      Left            =   10800
      TabIndex        =   92
      Top             =   6840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   13
      Left            =   14160
      TabIndex        =   93
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   14
      Left            =   14160
      TabIndex        =   94
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   15
      Left            =   14160
      TabIndex        =   95
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":0407
      Height          =   330
      Index           =   13
      Left            =   11640
      TabIndex        =   96
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":041B
      Height          =   330
      Index           =   14
      Left            =   11640
      TabIndex        =   97
      Top             =   6360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":042F
      Height          =   330
      Index           =   15
      Left            =   11640
      TabIndex        =   98
      Top             =   6840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   24
      Left            =   8640
      TabIndex        =   99
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy22.frx":0443
      Height          =   330
      Index           =   16
      Left            =   9120
      TabIndex        =   141
      Top             =   10200
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy22.frx":0457
      Height          =   330
      Index           =   16
      Left            =   10560
      TabIndex        =   142
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   16
      Left            =   13920
      TabIndex        =   143
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy22.frx":046C
      Height          =   330
      Index           =   16
      Left            =   11400
      TabIndex        =   144
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":0480
      Height          =   330
      Index           =   1
      Left            =   4200
      TabIndex        =   158
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":0495
      Height          =   330
      Index           =   1
      Left            =   5040
      TabIndex        =   159
      Top             =   3960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":04AA
      Height          =   330
      Index           =   2
      Left            =   4200
      TabIndex        =   160
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":04BF
      Height          =   330
      Index           =   2
      Left            =   5040
      TabIndex        =   161
      Top             =   4440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":04D4
      Height          =   330
      Index           =   3
      Left            =   4200
      TabIndex        =   162
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":04E9
      Height          =   330
      Index           =   3
      Left            =   5040
      TabIndex        =   163
      Top             =   4920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":04FE
      Height          =   330
      Index           =   4
      Left            =   4200
      TabIndex        =   164
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":0513
      Height          =   330
      Index           =   4
      Left            =   5040
      TabIndex        =   165
      Top             =   5400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":0528
      Height          =   330
      Index           =   5
      Left            =   4200
      TabIndex        =   166
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":053D
      Height          =   330
      Index           =   5
      Left            =   5040
      TabIndex        =   167
      Top             =   5880
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":0552
      Height          =   330
      Index           =   6
      Left            =   4200
      TabIndex        =   168
      Top             =   6360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":0567
      Height          =   330
      Index           =   6
      Left            =   5040
      TabIndex        =   169
      Top             =   6360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":057C
      Height          =   330
      Index           =   7
      Left            =   4200
      TabIndex        =   170
      Top             =   6840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":0591
      Height          =   330
      Index           =   7
      Left            =   5040
      TabIndex        =   171
      Top             =   6840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":05A6
      Height          =   330
      Index           =   8
      Left            =   12480
      TabIndex        =   172
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":05BB
      Height          =   330
      Index           =   8
      Left            =   13440
      TabIndex        =   173
      Top             =   3480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":05D0
      Height          =   330
      Index           =   9
      Left            =   12480
      TabIndex        =   174
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":05E5
      Height          =   330
      Index           =   9
      Left            =   13440
      TabIndex        =   175
      Top             =   3960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":05FA
      Height          =   330
      Index           =   10
      Left            =   12480
      TabIndex        =   176
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":060F
      Height          =   330
      Index           =   10
      Left            =   13440
      TabIndex        =   177
      Top             =   4440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":0624
      Height          =   330
      Index           =   11
      Left            =   12480
      TabIndex        =   178
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":0639
      Height          =   330
      Index           =   11
      Left            =   13440
      TabIndex        =   179
      Top             =   4920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":064E
      Height          =   330
      Index           =   12
      Left            =   12480
      TabIndex        =   180
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":0663
      Height          =   330
      Index           =   12
      Left            =   13440
      TabIndex        =   181
      Top             =   5400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":0678
      Height          =   330
      Index           =   13
      Left            =   12480
      TabIndex        =   182
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":068D
      Height          =   330
      Index           =   13
      Left            =   13440
      TabIndex        =   183
      Top             =   5880
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":06A2
      Height          =   330
      Index           =   14
      Left            =   12480
      TabIndex        =   184
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":06B7
      Height          =   330
      Index           =   14
      Left            =   13440
      TabIndex        =   185
      Top             =   6360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":06CC
      Height          =   330
      Index           =   15
      Left            =   12480
      TabIndex        =   186
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":06E1
      Height          =   330
      Index           =   15
      Left            =   13440
      TabIndex        =   187
      Top             =   6840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy22.frx":06F6
      Height          =   330
      Index           =   16
      Left            =   12240
      TabIndex        =   188
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy22.frx":070B
      Height          =   330
      Index           =   16
      Left            =   13200
      TabIndex        =   189
      Top             =   10200
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "批次"
      Text            =   "DBCombo4"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy22.frx":0720
      Height          =   1695
      Left            =   120
      TabIndex        =   190
      TabStop         =   0   'False
      Top             =   1320
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2990
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
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "合约号"
      Height          =   255
      Index           =   21
      Left            =   4080
      TabIndex        =   191
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "批号"
      Height          =   375
      Index           =   20
      Left            =   13440
      TabIndex        =   155
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "批号"
      Height          =   375
      Index           =   17
      Left            =   5040
      TabIndex        =   154
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣花编号"
      Height          =   375
      Index           =   11
      Left            =   6840
      TabIndex        =   149
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣花信息"
      Height          =   375
      Index           =   9
      Left            =   6840
      TabIndex        =   147
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "帽芯"
      Height          =   375
      Index           =   16
      Left            =   8280
      TabIndex        =   145
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "高头明线"
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   140
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "间线"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   139
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣眼"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   138
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   137
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料单位"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   136
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "商标"
      Height          =   375
      Index           =   10
      Left            =   8520
      TabIndex        =   135
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料顶扣"
      Height          =   375
      Index           =   9
      Left            =   8520
      TabIndex        =   134
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料后扣"
      Height          =   375
      Index           =   8
      Left            =   8520
      TabIndex        =   133
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "高头明线"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   132
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "胶条"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   131
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "间线"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   130
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "帽芯"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   129
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣眼"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   128
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "高头明线"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   127
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "间线"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   126
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣眼"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   125
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料数量"
      Height          =   375
      Index           =   34
      Left            =   5760
      TabIndex        =   124
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料规格"
      Height          =   375
      Index           =   33
      Left            =   2520
      TabIndex        =   123
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料名称"
      Height          =   375
      Index           =   32
      Left            =   1080
      TabIndex        =   122
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料明细"
      Height          =   375
      Index           =   31
      Left            =   120
      TabIndex        =   121
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  成衣定量备料分析表"
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
      TabIndex        =   120
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   119
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "具体要求"
      Height          =   495
      Index           =   10
      Left            =   6000
      TabIndex        =   118
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣花(印刷)说明"
      Height          =   495
      Index           =   12
      Left            =   8640
      TabIndex        =   117
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   375
      Index           =   13
      Left            =   120
      TabIndex        =   116
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "数量"
      Height          =   375
      Index           =   14
      Left            =   11040
      TabIndex        =   115
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   15
      Left            =   9840
      TabIndex        =   114
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "交货期"
      Height          =   375
      Index           =   16
      Left            =   9720
      TabIndex        =   113
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "刀模"
      Height          =   375
      Index           =   18
      Left            =   3960
      TabIndex        =   112
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   19
      Left            =   120
      TabIndex        =   111
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "双针"
      Height          =   375
      Index           =   11
      Left            =   8520
      TabIndex        =   110
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "汉带"
      Height          =   375
      Index           =   12
      Left            =   8520
      TabIndex        =   109
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前页"
      Height          =   375
      Index           =   13
      Left            =   8520
      TabIndex        =   108
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣眼"
      Height          =   375
      Index           =   14
      Left            =   8520
      TabIndex        =   107
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "帽芯"
      Height          =   375
      Index           =   15
      Left            =   8520
      TabIndex        =   106
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料明细"
      Height          =   375
      Index           =   0
      Left            =   8520
      TabIndex        =   105
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料名称"
      Height          =   375
      Index           =   4
      Left            =   9360
      TabIndex        =   104
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料规格"
      Height          =   375
      Index           =   5
      Left            =   10800
      TabIndex        =   103
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料数量"
      Height          =   375
      Index           =   6
      Left            =   14160
      TabIndex        =   102
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料单位"
      Height          =   375
      Index           =   7
      Left            =   11640
      TabIndex        =   101
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   8
      Left            =   12540
      TabIndex        =   100
      Top             =   3120
      Width           =   855
   End
End
Attribute VB_Name = "Formy22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X As Integer
Dim ba As Database: Dim rr As Integer
Dim rs As Single: Dim rd1 As Recordset: Dim ba1 As Database: Public ll As Integer: Public RQ As Date
Dim rd As Recordset: Public mm As Date: Public ml As Date

Private Sub Command12_Click()
Unload Me
Formy4.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
If DBCombo1(0).Text = "" Then
MsgBox ("请确认单号")
Exit Sub
End If

If DBCombo1(1).Text = "" Then
MsgBox ("请确认款号")
Exit Sub
End If


rd.AddNew


For i = 0 To 16
If DBCombo8(i).Text <> "" Then
rd.AddNew
rd.Fields(0) = Trim(DBCombo1(0).Text)
rd.Fields(1) = Trim(DBCombo1(1).Text)
rd.Fields(2) = Trim(DBCombo1(2).Text)
rd.Fields(3) = Trim(Label4(i).Caption)
rd.Fields(4) = Trim(DBCombo6(i).Text)
rd.Fields(5) = Trim(DBCombo7(i).Text)
rd.Fields(6) = Trim(DBCombo10(i).Text)
rd.Fields(7) = Trim(DBCombo3(i).Text)
rd.Fields(8) = Trim(DBCombo4(i).Text)
rd.Fields(9) = Trim(DBCombo8(i).Text)
rd.Fields(10) = "3线料库"
rd.Update
End If
Next


Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3线料库' "
Data4.Refresh



End Sub

Private Sub Command2_Click()
On Error Resume Next

If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If


Data4.Recordset.Edit
For i = 0 To 16
If DBCombo8(i).Text <> "" Then
Data4.Recordset.Edit
Data4.Recordset.Fields(0) = Trim(DBCombo1(0).Text)
Data4.Recordset.Fields(1) = Trim(DBCombo1(1).Text)
Data4.Recordset.Fields(2) = Trim(DBCombo1(2).Text)
Data4.Recordset.Fields(3) = Trim(Label4(i).Caption)
Data4.Recordset.Fields(4) = Trim(DBCombo6(i).Text)
Data4.Recordset.Fields(5) = Trim(DBCombo7(i).Text)
Data4.Recordset.Fields(6) = Trim(DBCombo10(i).Text)
Data4.Recordset.Fields(7) = Trim(DBCombo3(i).Text)
Data4.Recordset.Fields(8) = Trim(DBCombo4(i).Text)
Data4.Recordset.Fields(9) = Trim(DBCombo8(i).Text)
Data4.Recordset.Update
End If
Next
Data4.Recordset.Update

Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3线料库' "
Data4.Refresh



End Sub

Private Sub Command4_Click()

On Error Resume Next

If MsgBox("确定删除吗?,删除将不能恢复!", vbYesNo) = vbNo Then Exit Sub
Data4.Recordset.Delete
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3线料库' "
Data4.Refresh


End Sub

Private Sub Command5_Click()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.EOF

For i = 0 To 16
If Label4(i).Caption = Data4.Recordset.Fields(3) Then
DBCombo6(i).Text = Data4.Recordset.Fields(4)
DBCombo7(i).Text = Data4.Recordset.Fields(5)
DBCombo8(i).Text = Format(Data4.Recordset.Fields(9), "#0.00")
DBCombo10(i).Text = Data4.Recordset.Fields(6)
DBCombo3(i).Text = Data4.Recordset.Fields(7)
DBCombo4(i).Text = Data4.Recordset.Fields(8)
End If
Next
Data4.Recordset.MoveNext
Loop

End Sub

Private Sub Command6_Click()
DataEnvironment4.cldfd DBCombo1(0).Text
DataReport5.Show 1
DataEnvironment4.rscldfd.Close

End Sub



Private Sub Command8_Click()
On Error Resume Next
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh
Data8.Refresh
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

Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data10.RecordSource = "select SCZY_x.款号 from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "'  GROUP BY SCZY_X.款号 "
Data10.Refresh
       Case 1
Data1.RecordSource = "select * from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "' AND SCZY_X.款号='" & DBCombo1(1).Text & "' "
Data1.Refresh

Data5.RecordSource = "select SCZY_x.颜色 from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "' AND SCZY_X.款号='" & DBCombo1(1).Text & "' GROUP BY SCZY_X.颜色 "
Data5.Refresh

For i = 1 To 15     '''''''''''''15个字段付值
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next

Text1.Text = Data1.Recordset.Fields(16)
Text2.Text = Data1.Recordset.Fields(17)

Data3.RecordSource = "select * from SCZY_Z WHERE SCZY_Z.单号='" & DBCombo1(0).Text & "' "
Data3.Refresh


Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3线料库' AND 主辅名称<>'汉带' "
Data4.Refresh

Text3.Text = DBCombo1(2).Text
DBCombo1(3).Text = 12
     
     Case 2

End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
     Case 2
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "'  AND DLCLB.材料库类='3线料库'  AND 主辅名称<>'汉带'"
Data4.Refresh

Data1.RecordSource = "select * from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "' AND SCZY_X.款号='" & DBCombo1(1).Text & "'AND SCZY_X.颜色='" & DBCombo1(2).Text & "' "
Data1.Refresh
If Data1.Recordset.EOF Then
For i = 0 To 15
Label4(i).Caption = ""
Next
Else
l = 0
For i = 0 To 9
If Data1.Recordset.Fields(10 + 2 * l) <> "" Then
Label4(i).Caption = Data1.Recordset.Fields(10 + 2 * l)
Else
Label4(i).Caption = ""
End If
l = l + 1
Next
End If

Text1.Text = Data1.Recordset.Fields(16)
Text2.Text = Data1.Recordset.Fields(17)
Text3.Text = DBCombo1(2).Text
DBCombo1(3).Text = 12
DBCombo1(8).Text = Data1.Recordset.Fields(8)
DBCombo1(10).Text = Data1.Recordset.Fields(10)
DBCombo1(12).Text = Data1.Recordset.Fields(12)

For i = 0 To 16     '''''''''''''17个字段付值
DBCombo3(i).Text = DBCombo1(2).Text
Next
End Select

End Sub

Private Sub DBCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub



Private Sub DBCombo3_Change(Index As Integer)
Select Case Index
       Case Index
       If DBCombo6(Index).Text = "绣花线" Then
       Data13.RecordSource = "SELECT 批次 FROM CKGL WHERE CKGL.库类='3线料库' AND 材料名称='" & DBCombo6(Index).Text & "' AND 颜色='" & DBCombo3(Index).Text & "' GROUP BY 批次 "
       Data13.Refresh
       Else
       Data13.RecordSource = "SELECT * FROM CKGL WHERE 批次=null"
       Data13.Refresh
       End If
       Data14.RecordSource = "SELECT 库类,材料名称,材料规格,材料单位,颜色 FROM CLMC WHERE CLMC.库类='3线料库' AND CLMC.材料名称='" & DBCombo6(Index).Text & "' AND 颜色='" & DBCombo3(Index).Text & "' GROUP BY 库类,材料名称,材料规格,材料单位,颜色"
       Data14.Refresh

End Select
End Sub

Private Sub DBCombo3_Click(Index As Integer, Area As Integer)
Select Case Index
       Case Index
       If DBCombo6(Index).Text = "绣花线" Then
       Data13.RecordSource = "SELECT 批次 FROM CKGL WHERE CKGL.库类='3线料库' AND 材料名称='" & DBCombo6(Index).Text & "' AND 颜色='" & DBCombo3(Index).Text & "' GROUP BY 批次 "
       Data13.Refresh
       Else
       Data13.RecordSource = "SELECT * FROM CKGL WHERE 批次=null"
       Data13.Refresh
       End If
       Data14.RecordSource = "SELECT 库类,材料名称,材料规格,材料单位,颜色 FROM CLMC WHERE CLMC.库类='3线料库' AND CLMC.材料名称='" & DBCombo6(Index).Text & "' AND 颜色='" & DBCombo3(Index).Text & "' GROUP BY 库类,材料名称,材料规格,材料单位,颜色"
       Data14.Refresh
       If DBCombo3(Index).Text <> "" Then
       For i = Index + 1 To 15
       DBCombo3(i).Text = DBCombo3(Index).Text
       Next
       End If
End Select
End Sub

Private Sub DBCombo6_Change(Index As Integer)
Select Case Index
       Case Index
       Data11.RecordSource = "SELECT CLMC.材料规格 FROM CLMC WHERE CLMC.库类='3线料库' AND 材料名称='" & DBCombo6(Index).Text & "' GROUP BY CLMC.材料规格 "
       Data11.Refresh
       Data14.RecordSource = "SELECT 库类,材料名称,材料规格,材料单位,颜色 FROM CLMC WHERE CLMC.库类='3线料库' AND CLMC.材料名称='" & DBCombo6(Index).Text & "' GROUP BY 库类,材料名称,材料规格,材料单位,颜色 "
       Data14.Refresh
End Select
End Sub

Private Sub DBCombo6_Click(Index As Integer, Area As Integer)
Select Case Index
       Case Index
       Data11.RecordSource = "SELECT CLMC.材料规格 FROM CLMC WHERE CLMC.库类='3线料库' AND 材料名称='" & DBCombo6(Index).Text & "' GROUP BY CLMC.材料规格 "
       Data11.Refresh
       Data14.RecordSource = "SELECT 库类,材料名称,材料规格,材料单位,颜色 FROM CLMC WHERE CLMC.库类='3线料库' AND CLMC.材料名称='" & DBCombo6(Index).Text & "' GROUP BY 库类,材料名称,材料规格,材料单位,颜色 "
       Data14.Refresh
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
For i = 0 To 24
DBCombo1(i).Text = ""
Next

For i = 0 To 16
DBCombo6(i).Text = ""
Next

For i = 0 To 16
DBCombo7(i).Text = ""
Next

For i = 0 To 16
DBCombo8(i).Text = ""
Next

For i = 0 To 16
DBCombo10(i).Text = ""
Next

For i = 0 To 16
DBCombo3(i).Text = ""
DBCombo4(i).Text = ""
Next


For i = 0 To 2
Text6(i).Text = 0
Next

Text5.Text = ""

DBCombo2.Text = ""

Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "' ORDER BY VAL(SCZY_X.序号) DESC"
Data1.Refresh

For i = 0 To Data1.Recordset.Fields.Count - 1
DBCombo1(i).Text = ""
Next



Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "select MAX(VAL(SCZY_X.序号)) from SCZY_X  WHERE SCZY_X.单号='" & DBCombo1(0).Text & "'"
Data2.Refresh

DBCombo1(19).Text = 1
DBCombo1(19).Text = Data2.Recordset.Fields(0) + 1

Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select * from SCZY_Z WHERE SCZY_Z.单号='" & DBCombo1(0).Text & "' "
Data3.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3线料库' "
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"

Data6.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data6.RecordSource = "SELECT CKGL.材料名称 FROM CKGL WHERE CKGL.库类='1主料库' GROUP BY CKGL.材料名称 "
Data6.Refresh

Data7.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data7.RecordSource = "SELECT cldw.mc FROM cldw  GROUP BY cldw.mc"
Data7.Refresh

Data8.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data8.RecordSource = "SELECT CLMC.材料名称 FROM CLMC WHERE CLMC.库类='3线料库' GROUP BY CLMC.材料名称 "
Data8.Refresh

Data9.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data9.RecordSource = "SELECT * FROM XLMX"
Data9.Refresh

Data10.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data10.Refresh


Data11.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data11.Refresh

Data12.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data12.RecordSource = "SELECT YS FROM YS  GROUP BY YS "
Data12.Refresh


Data13.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data13.Refresh

Data14.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data14.Refresh
For i = 0 To 16
Label4(i).Caption = ""
Next

i = 0
Data9.Recordset.MoveFirst
Do While Not Data9.Recordset.EOF
Label4(i).Caption = Data9.Recordset.Fields(0)
i = i + 1
Data9.Recordset.MoveNext
Loop

MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 2000
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.ColWidth(6) = 1200
MSFlexGrid1.ColWidth(7) = 1200
MSFlexGrid1.ColWidth(8) = 1500

DBCombo1(1).TabIndex = 0
End Sub

Private Sub Label2_DBLClick(Index As Integer)
Select Case Index
   Case 9
   DBCombo17.Enabled = True
   End Select
End Sub


Private Sub Label4_dblClick(Index As Integer)
Select Case Index
       Case 0
       Formy30.Show
       Case 1
       Formy9.Show
       Case 2
       Formy33.Show
       Case 13
       Formy43.Show
       Case 14
       Formy44.Show
End Select
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
Data4.Recordset.MoveFirst
Data4.Recordset.Move rs - 1


For i = 0 To 16
If Trim(Data4.Recordset.Fields(3)) = Trim(Label4(i).Caption) Then
DBCombo3(i).Text = Data4.Recordset.Fields(7)
DBCombo4(i).Text = Data4.Recordset.Fields(8)
DBCombo6(i).Text = Data4.Recordset.Fields(4)
DBCombo7(i).Text = Data4.Recordset.Fields(5)
DBCombo8(i).Text = Format(Data4.Recordset.Fields(9), "#0.0000")
DBCombo10(i).Text = Data4.Recordset.Fields(6)
Else
DBCombo3(i).Text = ""
DBCombo4(i).Text = ""
DBCombo6(i).Text = ""
DBCombo7(i).Text = ""
DBCombo8(i).Text = ""
DBCombo10(i).Text = ""
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



