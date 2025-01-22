VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formy303 
   BackColor       =   &H00C0E0FF&
   Caption         =   "织布漂染"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form42"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data11 
      Caption         =   "Data1"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Data Data10 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data9 
      Caption         =   "Data3"
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data8 
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data5 
      Caption         =   "Data4"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data6 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data7 
      Caption         =   "Data2"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   12
      Left            =   7920
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   4080
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
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   6360
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   10680
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   9600
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   11520
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   9000
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   3840
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   4680
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   8400
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   6360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4080
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   4920
      TabIndex        =   17
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy303.frx":0000
      Height          =   1935
      Left            =   3840
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy303.frx":0014
      Height          =   3255
      Left            =   3840
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6360
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   14
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6360
      TabIndex        =   20
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80543745
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7920
      TabIndex        =   35
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80543745
      CurrentDate     =   36892
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7695
      Left            =   360
      TabIndex        =   39
      Top             =   1920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   13573
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1440
      TabIndex        =   40
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   80543745
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1440
      TabIndex        =   41
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   80543745
      CurrentDate     =   39177
   End
   Begin VB.Label Label7 
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
      Index           =   0
      Left            =   360
      TabIndex        =   43
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Index           =   0
      Left            =   360
      TabIndex        =   42
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "漂染期限"
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
      Left            =   7920
      TabIndex        =   36
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "光坯幅宽"
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
      Left            =   11520
      TabIndex        =   33
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择单号"
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
      Left            =   3840
      TabIndex        =   32
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择颜色"
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
      Left            =   5400
      TabIndex        =   31
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择款号"
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
      Left            =   7560
      TabIndex        =   30
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择单号"
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
      Left            =   3840
      TabIndex        =   29
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯幅宽"
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
      Left            =   9600
      TabIndex        =   28
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "克重"
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
      Left            =   10680
      TabIndex        =   27
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料名称"
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
      TabIndex        =   26
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "织布量"
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
      Left            =   4680
      TabIndex        =   25
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染耗"
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
      Left            =   3840
      TabIndex        =   24
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料数量"
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
      Left            =   8400
      TabIndex        =   23
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "裁耗"
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
      Left            =   9000
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "织布期限"
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
      Left            =   6360
      TabIndex        =   21
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "Formy303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call tree
Call zk
End Sub

Private Sub Command3_Click()
Call zbmx(MSFlexGrid1, "织布明细")
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
If MsgBox("删除不能回复！确定删除吗？", vbYesNo) = vbNo Then Exit Sub
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.Delete
Data4.Refresh
Text1(4).Text = ""
Text1(5).Text = ""
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command7_Click()
If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(9).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data4.Recordset.AddNew
For i = 0 To 12
Data4.Recordset.Fields(i) = Text1(i).Text
Next
Data4.Recordset.Update
Data4.Refresh

Text1(4).Text = ""
Text1(5).Text = ""
Text1(4).SetFocus
End Sub

Private Sub Command8_Click()
If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(9).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data4.Recordset.Edit
For i = 0 To 12
Data4.Recordset.Fields(i) = Text1(i).Text
Next
Data4.Recordset.Update
Data4.Refresh

Text1(4).Text = ""
Text1(5).Text = ""
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command9_Click()
If DBCombo2.Text = "" Then
Data4.RecordSource = "select * from zbjh where 单号='" & DBCombo1.Text & "' order by 款号,材料名称,颜色"
Data4.Refresh
Else
Data4.RecordSource = "select * from zbjh where 单号='" & DBCombo1.Text & "' and instr(款号,'" & DBCombo2.Text & "')>0 order by 材料名称,颜色"
Data4.Refresh
End If
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub DBCombo1_Change()
Data2.RecordSource = "SELECT  单号,材料名称,材料颜色,材料规格,材料单位,材料批号,材料数量 as 采购量,订单颜色 as 采购类别 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "' and  材料库类='1主料库'   order by 材料名称,材料规格,材料单位,材料批号"
Data2.Refresh
Data4.RecordSource = "select * from zbfl where 单号='" & DBCombo1.Text & "' order by 颜色"
Data4.Refresh
End Sub

Private Sub DTPicker1_Change()
Text1(11).Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text1(11).Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Change()
Text1(12).Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text1(12).Text = DTPicker2.Value
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
For i = 0 To 12
Text1(i).Text = ""
Next
Text1(11).Text = Date
DTPicker1.Value = Date
Text1(12).Text = Date
DTPicker2.Value = Date

Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(4) = 2000
MSFlexGrid1.ColWidth(12) = 1300
MSFlexGrid1.ColWidth(13) = 1300

MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(4) = 1800
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 1
xqbl = 3
Formy41.Show
End Select
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data4.Recordset.Move rs - 1
Text1(10).Text = Data4.Recordset.Fields(10)
Text1(11).Text = Data4.Recordset.Fields(11)
DTPicker1.Value = Text1(11).Text
Text1(12).Text = Data4.Recordset.Fields(12)
DTPicker2.Value = Text1(12).Text
For i = 0 To 9
Text1(i).Text = Data4.Recordset.Fields(i)
Next
Command7.Enabled = False
Command8.Enabled = True
Command6.Enabled = True
End Sub

Private Sub MSFlexGrid2_dblClick()
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
Text1(0).Text = Data2.Recordset.Fields(0)
Text1(1).Text = Data2.Recordset.Fields(1)
Text1(2).Text = Data2.Recordset.Fields(6)
Text1(3).Text = Data2.Recordset.Fields(3)
Text1(10).Text = Data2.Recordset.Fields(8)
Text1(4).Text = Data2.Recordset.Fields(7)
Text1(6).Text = Data2.Recordset.Fields(4)
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
      Case 8
If Val(Text1(8).Text) > 0 Then
Text1(9).Text = Format(Val(Text1(10).Text) / (1 - Val(Text1(8).Text) / 100), "#0.00")
End If
      Case 10
If Val(Text1(8).Text) > 0 Then
Text1(9).Text = Format(Val(Text1(10).Text) / (1 - Val(Text1(8).Text) / 100), "#0.00")
End If
End Select
End Sub

Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex

   TreeView1.Nodes.Clear
 
    Data6.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
    Data6.Refresh
    m = 1
    If Not Data6.Recordset.EOF Then  'make sure there are records in the table
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data6.Recordset.Fields(0)
        intIndex = mNode.Index
        Data5.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data6.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
        Data5.Refresh
        
        If Not Data5.Recordset.EOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data5.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data7.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data5.Recordset.Fields(0) & "' and 进度='进行'"
        Data7.Refresh
        
        If Not Data7.Recordset.EOF Then
        Data7.Recordset.MoveFirst
        Do While Not Data7.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data7.Recordset.Fields(0))
        Data7.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data5.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data6.Recordset.MoveNext
        Loop
    End If

End Sub


'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next

If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") > 0 Then
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
DBCombo1.Text = l1
End If

End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub



