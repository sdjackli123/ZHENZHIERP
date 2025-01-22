VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy24 
   BackColor       =   &H00C0E0FF&
   Caption         =   "定量分析包装库类"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form24"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   5040
      TabIndex        =   199
      Text            =   "Text8"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Data Data20 
      Caption         =   "Data20"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "线类"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   13260
      TabIndex        =   35
      Text            =   "Text4"
      Top             =   10680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   14040
      TabIndex        =   34
      Text            =   "Text4"
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   12780
      TabIndex        =   33
      Text            =   "Text4"
      Top             =   10560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   12780
      TabIndex        =   32
      Text            =   "Text4"
      Top             =   10560
      Visible         =   0   'False
      Width           =   495
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
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
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
      TabIndex        =   31
      Text            =   "Text3"
      Top             =   1080
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6360
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
      Left            =   360
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
      Height          =   615
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Text            =   "Formy24.frx":0000
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   1695
      Left            =   10680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Text            =   "Formy24.frx":0006
      Top             =   840
      Width           =   4455
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
      TabIndex        =   27
      Top             =   6360
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6360
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6360
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6360
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6360
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
      Left            =   1200
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6360
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
      Left            =   120
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
      Left            =   360
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
      Top             =   11040
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
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
      Top             =   10920
      Width           =   4935
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
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
      Top             =   11040
      Width           =   4815
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
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
      Width           =   4695
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11160
      Width           =   4335
   End
   Begin VB.Data Data15 
      Caption         =   "Data15"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11040
      Width           =   4215
   End
   Begin VB.Data Data16 
      Caption         =   "Data16"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10920
      Width           =   4215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Text            =   "Formy24.frx":000C
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   19
      Text            =   "Text6"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   18
      Text            =   "Text6"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   17
      Text            =   "Text6"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   16
      Text            =   "Text6"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   4920
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   13680
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   13680
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   13680
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   13680
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   13680
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   12
      Left            =   13680
      TabIndex        =   8
      Text            =   "Text6"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   13
      Left            =   13680
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   14
      Left            =   13680
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   6360
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   15
      Left            =   13680
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   8280
      Width           =   495
   End
   Begin VB.Data Data17 
      Caption         =   "Data17"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   16
      Left            =   13560
      TabIndex        =   4
      Text            =   "Text6"
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   17
      Left            =   13560
      TabIndex        =   3
      Text            =   "Text6"
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Data Data18 
      Caption         =   "Data18"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data19 
      Caption         =   "Data19"
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
      Width           =   4335
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Text            =   "Text7"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Text            =   "Text7"
      Top             =   8040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy24.frx":0012
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6840
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   5106
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
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy24.frx":0026
      Height          =   330
      Index           =   0
      Left            =   3120
      TabIndex        =   37
      Top             =   3000
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
      Left            =   5400
      TabIndex        =   38
      Top             =   3000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":003A
      Height          =   330
      Index           =   0
      Left            =   2280
      TabIndex        =   39
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":004F
      Height          =   330
      Index           =   0
      Left            =   1080
      TabIndex        =   40
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
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
      TabIndex        =   41
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy24.frx":0063
      Height          =   330
      Index           =   1
      Left            =   960
      TabIndex        =   42
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "款号"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy24.frx":0078
      Height          =   330
      Index           =   2
      Left            =   960
      TabIndex        =   43
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "颜色"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   3
      Left            =   12600
      TabIndex        =   44
      Top             =   360
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
      Left            =   2760
      TabIndex        =   45
      Top             =   120
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
      Left            =   1320
      TabIndex        =   46
      Top             =   9120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   6
      Left            =   7680
      TabIndex        =   47
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
      Index           =   7
      Left            =   7680
      TabIndex        =   48
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   8
      Left            =   7680
      TabIndex        =   49
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
      Index           =   9
      Left            =   1320
      TabIndex        =   50
      Top             =   9240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   10
      Left            =   8040
      TabIndex        =   51
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
      Index           =   11
      Left            =   1320
      TabIndex        =   52
      Top             =   10440
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   12
      Left            =   8160
      TabIndex        =   53
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
      Index           =   13
      Left            =   1320
      TabIndex        =   54
      Top             =   10920
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   14
      Left            =   8280
      TabIndex        =   55
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   15
      Left            =   7680
      TabIndex        =   56
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   16
      Left            =   6120
      TabIndex        =   57
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   17
      Left            =   9360
      TabIndex        =   58
      Top             =   10800
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
      Left            =   8280
      TabIndex        =   59
      Top             =   8880
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
      Left            =   8400
      TabIndex        =   60
      Top             =   9240
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":008C
      Height          =   330
      Index           =   1
      Left            =   1080
      TabIndex        =   61
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":00A0
      Height          =   330
      Index           =   2
      Left            =   1080
      TabIndex        =   62
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":00B4
      Height          =   330
      Index           =   3
      Left            =   1080
      TabIndex        =   63
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":00C8
      Height          =   330
      Index           =   4
      Left            =   1080
      TabIndex        =   64
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":00DC
      Height          =   330
      Index           =   5
      Left            =   1080
      TabIndex        =   65
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":00F0
      Height          =   330
      Index           =   6
      Left            =   1080
      TabIndex        =   66
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":0104
      Height          =   330
      Index           =   7
      Left            =   9840
      TabIndex        =   67
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":0118
      Height          =   330
      Index           =   8
      Left            =   9840
      TabIndex        =   68
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":012C
      Height          =   330
      Index           =   9
      Left            =   9840
      TabIndex        =   69
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":0140
      Height          =   330
      Index           =   10
      Left            =   9840
      TabIndex        =   70
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   1
      Left            =   2280
      TabIndex        =   71
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":0154
      Height          =   330
      Index           =   2
      Left            =   2280
      TabIndex        =   72
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
      Height          =   330
      Index           =   3
      Left            =   2280
      TabIndex        =   73
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":0169
      Height          =   330
      Index           =   4
      Left            =   2280
      TabIndex        =   74
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
      Height          =   330
      Index           =   5
      Left            =   2280
      TabIndex        =   75
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":017E
      Height          =   330
      Index           =   6
      Left            =   2280
      TabIndex        =   76
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
      Height          =   330
      Index           =   7
      Left            =   11040
      TabIndex        =   77
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":0193
      Height          =   330
      Index           =   8
      Left            =   11040
      TabIndex        =   78
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
      Bindings        =   "Formy24.frx":01A8
      Height          =   330
      Index           =   9
      Left            =   11040
      TabIndex        =   79
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
      Bindings        =   "Formy24.frx":01BD
      Height          =   330
      Index           =   10
      Left            =   11040
      TabIndex        =   80
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
      Left            =   5400
      TabIndex        =   81
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
      Index           =   2
      Left            =   5400
      TabIndex        =   82
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
      Index           =   3
      Left            =   5400
      TabIndex        =   83
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
      Index           =   4
      Left            =   5400
      TabIndex        =   84
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
      Index           =   5
      Left            =   5400
      TabIndex        =   85
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
      Index           =   6
      Left            =   5400
      TabIndex        =   86
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
      Index           =   7
      Left            =   14160
      TabIndex        =   87
      Top             =   3000
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
      TabIndex        =   88
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
      TabIndex        =   89
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
      TabIndex        =   90
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":01D2
      Height          =   330
      Index           =   11
      Left            =   9840
      TabIndex        =   91
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":01E6
      Height          =   330
      Index           =   11
      Left            =   11040
      TabIndex        =   92
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
      TabIndex        =   93
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy24.frx":01FB
      Height          =   330
      Index           =   1
      Left            =   3120
      TabIndex        =   94
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
      Bindings        =   "Formy24.frx":020F
      Height          =   330
      Index           =   2
      Left            =   3120
      TabIndex        =   95
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
      Bindings        =   "Formy24.frx":0223
      Height          =   330
      Index           =   3
      Left            =   3120
      TabIndex        =   96
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
      Bindings        =   "Formy24.frx":0237
      Height          =   330
      Index           =   4
      Left            =   3120
      TabIndex        =   97
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
      Bindings        =   "Formy24.frx":024B
      Height          =   330
      Index           =   5
      Left            =   3120
      TabIndex        =   98
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
      Bindings        =   "Formy24.frx":025F
      Height          =   330
      Index           =   6
      Left            =   3120
      TabIndex        =   99
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
      Bindings        =   "Formy24.frx":0273
      Height          =   330
      Index           =   7
      Left            =   11880
      TabIndex        =   100
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy24.frx":0287
      Height          =   330
      Index           =   8
      Left            =   11880
      TabIndex        =   101
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
      Bindings        =   "Formy24.frx":029B
      Height          =   330
      Index           =   9
      Left            =   11880
      TabIndex        =   102
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
      Bindings        =   "Formy24.frx":02AF
      Height          =   330
      Index           =   10
      Left            =   11880
      TabIndex        =   103
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
      Bindings        =   "Formy24.frx":02C3
      Height          =   330
      Index           =   11
      Left            =   11880
      TabIndex        =   104
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
      Bindings        =   "Formy24.frx":02D7
      Height          =   330
      Index           =   12
      Left            =   11880
      TabIndex        =   105
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
      TabIndex        =   106
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":02EB
      Height          =   330
      Index           =   12
      Left            =   11040
      TabIndex        =   107
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
      Bindings        =   "Formy24.frx":0300
      Height          =   330
      Index           =   12
      Left            =   9840
      TabIndex        =   108
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   20
      Left            =   7680
      TabIndex        =   109
      Top             =   8040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   21
      Left            =   1320
      TabIndex        =   110
      Top             =   10320
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   22
      Left            =   7680
      TabIndex        =   111
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   23
      Left            =   7680
      TabIndex        =   112
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":0314
      Height          =   330
      Index           =   13
      Left            =   9840
      TabIndex        =   113
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":0328
      Height          =   330
      Index           =   14
      Left            =   9840
      TabIndex        =   114
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":033C
      Height          =   330
      Index           =   15
      Left            =   9840
      TabIndex        =   115
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":0350
      Height          =   330
      Index           =   13
      Left            =   11040
      TabIndex        =   116
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
      Bindings        =   "Formy24.frx":0365
      Height          =   330
      Index           =   14
      Left            =   11040
      TabIndex        =   117
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
      Bindings        =   "Formy24.frx":037A
      Height          =   330
      Index           =   15
      Left            =   11040
      TabIndex        =   118
      Top             =   8280
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
      TabIndex        =   119
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
      TabIndex        =   120
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
      TabIndex        =   121
      Top             =   9000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy24.frx":038F
      Height          =   330
      Index           =   13
      Left            =   11880
      TabIndex        =   122
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
      Bindings        =   "Formy24.frx":03A3
      Height          =   330
      Index           =   14
      Left            =   11880
      TabIndex        =   123
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
      Bindings        =   "Formy24.frx":03B7
      Height          =   330
      Index           =   15
      Left            =   11880
      TabIndex        =   124
      Top             =   8280
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
      Left            =   1320
      TabIndex        =   125
      Top             =   10800
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   25
      Left            =   1320
      TabIndex        =   126
      Top             =   10440
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":03CB
      Height          =   330
      Index           =   16
      Left            =   9720
      TabIndex        =   127
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy24.frx":03DF
      Height          =   330
      Index           =   17
      Left            =   9720
      TabIndex        =   128
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":03F3
      Height          =   330
      Index           =   16
      Left            =   10920
      TabIndex        =   129
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formy24.frx":0408
      Height          =   330
      Index           =   17
      Left            =   10920
      TabIndex        =   130
      Top             =   6840
      Visible         =   0   'False
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
      Index           =   16
      Left            =   14040
      TabIndex        =   131
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   17
      Left            =   14040
      TabIndex        =   132
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy24.frx":041D
      Height          =   330
      Index           =   16
      Left            =   11760
      TabIndex        =   133
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy24.frx":0431
      Height          =   330
      Index           =   17
      Left            =   11760
      TabIndex        =   134
      Top             =   6840
      Visible         =   0   'False
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
      Index           =   26
      Left            =   7680
      TabIndex        =   135
      Top             =   8640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":0445
      Height          =   330
      Index           =   0
      Left            =   4020
      TabIndex        =   136
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":045A
      Height          =   330
      Index           =   1
      Left            =   4020
      TabIndex        =   137
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":046F
      Height          =   330
      Index           =   2
      Left            =   4020
      TabIndex        =   138
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":0484
      Height          =   330
      Index           =   3
      Left            =   4020
      TabIndex        =   139
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":0499
      Height          =   330
      Index           =   4
      Left            =   4020
      TabIndex        =   140
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":04AE
      Height          =   330
      Index           =   5
      Left            =   4020
      TabIndex        =   141
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":04C3
      Height          =   330
      Index           =   6
      Left            =   4020
      TabIndex        =   142
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":04D8
      Height          =   330
      Index           =   7
      Left            =   12800
      TabIndex        =   143
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":04ED
      Height          =   330
      Index           =   8
      Left            =   12800
      TabIndex        =   144
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":0502
      Height          =   330
      Index           =   9
      Left            =   12800
      TabIndex        =   145
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":0517
      Height          =   330
      Index           =   10
      Left            =   12800
      TabIndex        =   146
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":052C
      Height          =   330
      Index           =   11
      Left            =   12800
      TabIndex        =   147
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":0541
      Height          =   330
      Index           =   12
      Left            =   12800
      TabIndex        =   148
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":0556
      Height          =   330
      Index           =   13
      Left            =   12800
      TabIndex        =   149
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":056B
      Height          =   330
      Index           =   14
      Left            =   12800
      TabIndex        =   150
      Top             =   6360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":0580
      Height          =   330
      Index           =   15
      Left            =   12720
      TabIndex        =   151
      Top             =   8280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":0595
      Height          =   330
      Index           =   16
      Left            =   12600
      TabIndex        =   152
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy24.frx":05AA
      Height          =   330
      Index           =   17
      Left            =   12600
      TabIndex        =   153
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo3"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy24.frx":05BF
      Height          =   975
      Left            =   120
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   1560
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1720
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
      Caption         =   "部位"
      Height          =   375
      Index           =   20
      Left            =   4200
      TabIndex        =   200
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   3
      Left            =   4020
      TabIndex        =   198
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料单位"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   197
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "商标"
      Height          =   375
      Index           =   10
      Left            =   7440
      TabIndex        =   196
      Top             =   10680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料顶扣"
      Height          =   375
      Index           =   9
      Left            =   7440
      TabIndex        =   195
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料后扣"
      Height          =   375
      Index           =   8
      Left            =   8880
      TabIndex        =   194
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "高头明线"
      Height          =   375
      Index           =   7
      Left            =   8880
      TabIndex        =   193
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "胶条"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   192
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "间线"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   191
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "帽芯"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   190
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣眼"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   189
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前页"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   188
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "汉带"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   187
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "双针"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   186
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料数量"
      Height          =   375
      Index           =   34
      Left            =   5400
      TabIndex        =   185
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料规格"
      Height          =   375
      Index           =   33
      Left            =   2280
      TabIndex        =   184
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料名称"
      Height          =   375
      Index           =   32
      Left            =   1080
      TabIndex        =   183
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料明细"
      Height          =   375
      Index           =   31
      Left            =   120
      TabIndex        =   182
      Top             =   2640
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
      TabIndex        =   181
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   180
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "具体要求"
      Height          =   615
      Index           =   10
      Left            =   7080
      TabIndex        =   179
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣花(印刷)说明"
      Height          =   495
      Index           =   12
      Left            =   9720
      TabIndex        =   178
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   375
      Index           =   13
      Left            =   120
      TabIndex        =   177
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "数量"
      Height          =   375
      Index           =   14
      Left            =   11760
      TabIndex        =   176
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   15
      Left            =   6120
      TabIndex        =   175
      Top             =   9360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "交货期"
      Height          =   375
      Index           =   16
      Left            =   5520
      TabIndex        =   174
      Top             =   10320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "刀模"
      Height          =   375
      Index           =   18
      Left            =   7440
      TabIndex        =   173
      Top             =   7440
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
      TabIndex        =   172
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "双针"
      Height          =   375
      Index           =   11
      Left            =   8880
      TabIndex        =   171
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "汉带"
      Height          =   375
      Index           =   12
      Left            =   8880
      TabIndex        =   170
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前页"
      Height          =   375
      Index           =   13
      Left            =   8880
      TabIndex        =   169
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣眼"
      Height          =   375
      Index           =   14
      Left            =   8880
      TabIndex        =   168
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "帽芯"
      Height          =   375
      Index           =   15
      Left            =   5520
      TabIndex        =   167
      Top             =   10560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料明细"
      Height          =   375
      Index           =   0
      Left            =   8880
      TabIndex        =   166
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料名称"
      Height          =   375
      Index           =   4
      Left            =   9840
      TabIndex        =   165
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料规格"
      Height          =   375
      Index           =   5
      Left            =   11040
      TabIndex        =   164
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料数量"
      Height          =   375
      Index           =   6
      Left            =   14160
      TabIndex        =   163
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料单位"
      Height          =   375
      Index           =   7
      Left            =   11880
      TabIndex        =   162
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   8
      Left            =   12780
      TabIndex        =   161
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "批号"
      Height          =   375
      Index           =   9
      Left            =   4920
      TabIndex        =   160
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "批号"
      Height          =   375
      Index           =   11
      Left            =   13680
      TabIndex        =   159
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "其它"
      Height          =   375
      Index           =   16
      Left            =   5520
      TabIndex        =   158
      Top             =   10680
      Visible         =   0   'False
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "其它"
      Height          =   375
      Index           =   17
      Left            =   8880
      TabIndex        =   157
      Top             =   5880
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "合约号"
      Height          =   255
      Index           =   17
      Left            =   6840
      TabIndex        =   156
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "其它"
      Height          =   375
      Index           =   18
      Left            =   8880
      TabIndex        =   155
      Top             =   6360
      Width           =   855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Formy24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X, c, r As Integer: Public ms As String
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


For i = 0 To 17
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
rd.Fields(8) = Trim(Text6(i).Text)
rd.Fields(9) = Trim(DBCombo8(i).Text)
rd.Fields(10) = "3包装库"
rd.Fields(11) = Text8.Text
rd.Update
End If
Next


Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "'  AND DLCLB.材料库类='3包装库' order by 材料名称,订单颜色"
Data4.Refresh
Data20.RecordSource = "select max(val(部位)) from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3包装库'"
Data20.Refresh
Text8.Text = "01"
If Data20.Recordset.EOF Then
Text8.Text = "01"
Else
If Len(Data20.Recordset.Fields(0) + 1) = 1 Then
Text8.Text = "0" + Trim(Data20.Recordset.Fields(0) + 1)
Else
Text8.Text = Trim(Data20.Recordset.Fields(0) + 1)
End If
End If


End Sub

Private Sub Command2_Click()
On Error Resume Next

If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If


Data4.Recordset.Edit
For i = 0 To 17
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
Data4.Recordset.Fields(8) = Trim(Text6(i).Text)
Data4.Recordset.Fields(9) = Trim(DBCombo8(i).Text)
Data4.Recordset.Fields(11) = Text8.Text
Data4.Recordset.Update
End If
Next
Data4.Recordset.Update

Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "'  AND DLCLB.材料库类='3包装库' order by 材料名称,订单颜色"
Data4.Refresh

Data20.RecordSource = "select max(val(部位)) from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3包装库'"
Data20.Refresh
Text8.Text = "01"
If Data20.Recordset.EOF Then
Text8.Text = "01"
Else
If Len(Data20.Recordset.Fields(0) + 1) = 1 Then
Text8.Text = "0" + Trim(Data20.Recordset.Fields(0) + 1)
Else
Text8.Text = Trim(Data20.Recordset.Fields(0) + 1)
End If
End If

End Sub

Private Sub Command4_Click()

On Error Resume Next

If MsgBox("确定删除吗?,删除将不能恢复!", vbYesNo) = vbNo Then Exit Sub
Data4.Recordset.Delete
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "'  AND DLCLB.材料库类='3包装库' order by 材料名称,订单颜色"
Data4.Refresh

Data20.RecordSource = "select max(val(部位)) from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3包装库'"
Data20.Refresh
Text8.Text = "01"
If Data20.Recordset.EOF Then
Text8.Text = "01"
Else
If Len(Data20.Recordset.Fields(0) + 1) = 1 Then
Text8.Text = "0" + Trim(Data20.Recordset.Fields(0) + 1)
Else
Text8.Text = Trim(Data20.Recordset.Fields(0) + 1)
End If
End If
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
Text6(i).Text = Data4.Recordset.Fields(8)
End If
Next
Data4.Recordset.MoveNext
Loop

End Sub

Private Sub Command6_Click()
'DataEnvironment4.CLDFL DBCombo1(0).Text
'DataReport8.Show 1
'DataEnvironment4.rsCLDFL.Close
Call MXOutDataToExcel(MSFlexGrid1, "辅料耗料表")
End Sub

Private Sub Command7_Click()
On Error Resume Next
For i = 0 To 4
Formy22.DBCombo1(i).Text = DBCombo1(i).Text
Next
Formy22.Text1 = Text1.Text
Formy22.Text2 = Text2.Text
Formy22.Text3 = Text3.Text
Formy22.Show
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
Data20.RecordSource = "select max(val(部位)) from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3包装库'"
Data20.Refresh
Text8.Text = "01"
If Data20.Recordset.EOF Then
Text8.Text = "01"
Else
If Len(Data20.Recordset.Fields(0) + 1) = 1 Then
Text8.Text = "0" + Trim(Data20.Recordset.Fields(0) + 1)
Else
Text8.Text = Trim(Data20.Recordset.Fields(0) + 1)
End If
End If
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

Private Sub Command9_Click()

End Sub

Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data11.RecordSource = "select SCZY_x.款号 from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "'  GROUP BY SCZY_X.款号 "
Data11.Refresh

       Case 1
Data1.RecordSource = "select * from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "' AND SCZY_X.款号='" & DBCombo1(1).Text & "' "
Data1.Refresh

Data5.RecordSource = "select SCZY_x.颜色 from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "' AND SCZY_X.款号='" & DBCombo1(1).Text & "' GROUP BY SCZY_X.颜色 "
Data5.Refresh

For i = 1 To 17     '''''''''''''17个字段付值
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next

Text1.Text = Data1.Recordset.Fields(16)
Text2.Text = Data1.Recordset.Fields(17)

Data3.RecordSource = "select * from SCZY_Z WHERE SCZY_Z.单号='" & DBCombo1(0).Text & "' "
Data3.Refresh


Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND 主辅名称<>'汉带' AND DLCLB.材料库类='3包装库' AND 主辅名称<>'汉带' order by 材料名称,订单颜色"
Data4.Refresh

Text3.Text = DBCombo1(2).Text
DBCombo1(3).Text = 12
     
Data20.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data20.RecordSource = "select max(val(部位)) from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3包装库'"
Data20.Refresh
Text8.Text = "01"
If Data20.Recordset.EOF Then
Text8.Text = "01"
Else
If Len(Data20.Recordset.Fields(0) + 1) = 1 Then
Text8.Text = "0" + Trim(Data20.Recordset.Fields(0) + 1)
Else
Text8.Text = Trim(Data20.Recordset.Fields(0) + 1)
End If
End If

End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
     Case 2
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND 主辅名称<>'汉带' AND DLCLB.材料库类='3包装库' AND 主辅名称<>'汉带' order by 材料名称,订单颜色 "
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

For i = 5 To 17     '''''''''''''17个字段付值
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next

'For i = 0 To 19    '''''''''''''17个字段付值
'DBCombo3(i).Text = DBCombo1(2).Text
'Next

Data20.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data20.RecordSource = "select max(val(部位)) from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3包装库'"
Data20.Refresh
Text8.Text = "01"
If Data20.Recordset.EOF Then
Text8.Text = "01"
Else
If Len(Data20.Recordset.Fields(0) + 1) = 1 Then
Text8.Text = "0" + Trim(Data20.Recordset.Fields(0) + 1)
Else
Text8.Text = Trim(Data20.Recordset.Fields(0) + 1)
End If
End If


Text1.Text = Data1.Recordset.Fields(16)
Text2.Text = Data1.Recordset.Fields(17)
Text3.Text = DBCombo1(2).Text
DBCombo1(3).Text = 12
End Select

End Sub

Private Sub DBCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub






Private Sub DBCombo10_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label4(i).Caption <> "" Then
       DBCombo10(i).Text = DBCombo10(Index).Text
       End If
       Next
End Select

End Sub

Private Sub DBCombo3_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label4(i).Caption <> "" Then
       DBCombo3(i).Text = DBCombo3(Index).Text
       End If
       Next
End Select

End Sub

Private Sub DBCombo3_Click(Index As Integer, Area As Integer)
Select Case Index
       Case Index
       Data19.RecordSource = "SELECT 库类,材料名称,材料规格,材料单位,颜色 FROM CLMC WHERE CLMC.库类='3包装库' AND CLMC.材料名称='" & DBCombo6(Index).Text & "' AND 颜色='" & DBCombo3(Index).Text & "' GROUP BY 库类,材料名称,材料规格,材料单位,颜色"
       Data19.Refresh
       If DBCombo3(Index).Text <> "" Then
       For i = Index + 1 To 17
       DBCombo3(i).Text = DBCombo3(Index).Text
       Next
       End If
End Select
End Sub

Private Sub DBCombo6_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label4(i).Caption <> "" Then
       DBCombo6(i).Text = DBCombo6(Index).Text
       End If
       Next
End Select
End Sub

Private Sub DBCombo6_Click(Index As Integer, Area As Integer)
Select Case Index
       Case Index
       Data17.RecordSource = "SELECT CLMC.材料规格 FROM CLMC WHERE CLMC.库类='3包装库' AND 材料名称='" & DBCombo6(Index).Text & "' GROUP BY CLMC.材料规格 "
       Data17.Refresh
       Data19.RecordSource = "SELECT 库类,材料名称,材料规格,材料单位,颜色 FROM CLMC WHERE CLMC.库类='3包装库' AND CLMC.材料名称='" & DBCombo6(Index).Text & "' GROUP BY 库类,材料名称,材料规格,材料单位,颜色 "
       Data19.Refresh
End Select
End Sub


Private Sub DBCombo7_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label4(i).Caption <> "" Then
       DBCombo7(i).Text = DBCombo7(Index).Text
       End If
       Next
End Select

End Sub

Private Sub DBCombo8_Change(Index As Integer)
Select Case Index
       Case Index
       For i = Index + 1 To 10
       If Label4(i).Caption <> "" Then
       DBCombo8(i).Text = DBCombo8(Index).Text
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
Text7.Text = ""
Text8.Text = ""
For i = 0 To 26
DBCombo1(i).Text = ""
Next

For i = 0 To 17
DBCombo6(i).Text = ""
Next

For i = 0 To 17
DBCombo7(i).Text = ""
Next

For i = 0 To 17
DBCombo8(i).Text = ""
Next

For i = 0 To 17
DBCombo10(i).Text = ""
Next

For i = 0 To 17
DBCombo3(i).Text = ""
Text6(i).Text = ""
Next

Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "' ORDER BY VAL(SCZY_X.序号) DESC"
Data1.Refresh

For i = 0 To Data1.Recordset.Fields.Count - 1
DBCombo1(i).Text = ""
Next

Text5.Text = ""

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
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' AND DLCLB.材料库类='3包装库' order by 材料名称,订单颜色"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"

Data6.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data6.RecordSource = "SELECT CLMC.材料名称 FROM CLMC WHERE CLMC.库类='1主料库' GROUP BY CLMC.材料名称 "
Data6.Refresh

Data7.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data7.RecordSource = "SELECT cldw.mc FROM cldw  GROUP BY cldw.mc"
Data7.Refresh

Data8.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data8.RecordSource = "SELECT CLMC.材料名称 FROM CLMC WHERE CLMC.库类='3包装库' GROUP BY CLMC.材料名称 "
Data8.Refresh

Data9.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data9.RecordSource = "SELECT * FROM FLMX"
Data9.Refresh

Data10.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data11.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data11.Refresh

Data12.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data12.Refresh

Data17.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data17.Refresh

Data18.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data18.RecordSource = "SELECT YS FROM YS  GROUP BY YS "
Data18.Refresh

Data19.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data19.Refresh

For i = 0 To 17
Label4(i).Caption = ""
Next

i = 0
Data9.Recordset.MoveFirst
Do While Not Data9.Recordset.EOF
Label4(i).Caption = Data9.Recordset.Fields(0)
i = i + 1
Data9.Recordset.MoveNext
Loop

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1200
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
Select Case Index
   Case 31
     

Data10.RecordSource = "SELECT * FROM SZSZ WHERE  INSTR('" & DBCombo1(5).Text & "',SZSZ.双针名称)>0"
Data10.Refresh
       If Data10.Recordset.EOF Then
       DBCombo6(0).Text = ""
       DBCombo7(0).Text = ""
       DBCombo10(0).Text = ""
       DBCombo8(0).Text = ""
       Else
       DBCombo6(0).Text = Data10.Recordset.Fields(1)
       DBCombo7(0).Text = Data10.Recordset.Fields(2)
       DBCombo10(0).Text = Data10.Recordset.Fields(3)
       DBCombo8(0).Text = Data10.Recordset.Fields(4)
       End If
       
Data10.RecordSource = "SELECT * FROM MXSZ WHERE  INSTR('" & DBCombo1(9).Text & "',MXSZ.帽芯名称)>0"
Data10.Refresh
       If Data10.Recordset.EOF Then
       DBCombo6(4).Text = ""
       DBCombo7(4).Text = ""
       DBCombo10(4).Text = ""
       DBCombo8(4).Text = ""
       Else
       DBCombo6(4).Text = Data10.Recordset.Fields(1)
       DBCombo7(4).Text = Data10.Recordset.Fields(2)
       DBCombo10(4).Text = Data10.Recordset.Fields(3)
       DBCombo8(4).Text = Data10.Recordset.Fields(4)
       End If

Data10.RecordSource = "SELECT * FROM JTSZ WHERE  INSTR('" & DBCombo1(11).Text & "',JTSZ.胶条名称)>0"
Data10.Refresh
       If Data10.Recordset.EOF Then
       DBCombo6(6).Text = ""
       DBCombo7(6).Text = ""
       DBCombo10(6).Text = ""
       DBCombo8(6).Text = ""
       Else
       DBCombo6(6).Text = Data10.Recordset.Fields(1)
       DBCombo7(6).Text = Data10.Recordset.Fields(2)
       DBCombo10(6).Text = Data10.Recordset.Fields(3)
       DBCombo8(6).Text = Data10.Recordset.Fields(4)
       End If

Data10.RecordSource = "SELECT * FROM HKSZ WHERE  INSTR('" & DBCombo1(11).Text & "',HKSZ.后扣名称)>0"
Data10.Refresh
       If Data10.Recordset.EOF Then
       DBCombo6(8).Text = ""
       DBCombo7(8).Text = ""
       DBCombo10(8).Text = ""
       DBCombo8(8).Text = ""
       Else
       DBCombo6(8).Text = Data10.Recordset.Fields(1)
       DBCombo7(8).Text = Data10.Recordset.Fields(2)
       DBCombo10(8).Text = Data10.Recordset.Fields(3)
       DBCombo8(8).Text = Data10.Recordset.Fields(4)
       End If
   
Data10.RecordSource = "SELECT * FROM DKSZ WHERE  INSTR('" & DBCombo1(14).Text & "',DKSZ.顶扣名称)>0"
Data10.Refresh
       If Data10.Recordset.EOF Then
       DBCombo6(9).Text = ""
       DBCombo7(9).Text = ""
       DBCombo10(9).Text = ""
       DBCombo8(9).Text = ""
       Else
       DBCombo6(9).Text = Data10.Recordset.Fields(1)
       DBCombo7(9).Text = Data10.Recordset.Fields(2)
       DBCombo10(9).Text = Data10.Recordset.Fields(3)
       DBCombo8(9).Text = Data10.Recordset.Fields(4)
       End If
   
Data10.RecordSource = "SELECT * FROM SBSZ WHERE  INSTR('" & DBCombo1(15).Text & "',SBSZ.商标名称)>0"
Data10.Refresh
       If Data10.Recordset.EOF Then
       DBCombo6(10).Text = ""
       DBCombo7(10).Text = ""
       DBCombo10(10).Text = ""
       DBCombo8(10).Text = ""
       Else
       DBCombo6(10).Text = Data10.Recordset.Fields(1)
       DBCombo7(10).Text = Data10.Recordset.Fields(2)
       DBCombo10(10).Text = Data10.Recordset.Fields(3)
       DBCombo8(10).Text = Data10.Recordset.Fields(4)
       End If
   
   End Select
End Sub




Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
Data4.Recordset.MoveFirst
Data4.Recordset.Move rs - 1


For i = 0 To 16
If Data4.Recordset.Fields(3) = Label4(i).Caption Then
DBCombo6(i).Text = Data4.Recordset.Fields(4)
DBCombo7(i).Text = Data4.Recordset.Fields(5)
DBCombo8(i).Text = Format(Data4.Recordset.Fields(9), "#0.0000")
DBCombo10(i).Text = Data4.Recordset.Fields(6)
DBCombo3(i).Text = Data4.Recordset.Fields(7)
Text6(i).Text = Data4.Recordset.Fields(8)
Text8.Text = Data4.Recordset.Fields(11)
Else
DBCombo6(i).Text = ""
DBCombo7(i).Text = ""
DBCombo8(i).Text = ""
DBCombo10(i).Text = ""
DBCombo3(i).Text = ""
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
       If Label4(i).Caption <> "" Then
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
On Error Resume Next
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




