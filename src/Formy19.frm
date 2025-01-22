VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy19 
   BackColor       =   &H00C0E0FF&
   Caption         =   "大货生产工艺通知单"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号打印"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data8 
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
      Top             =   5160
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   5280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      Left            =   9840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   47
      Text            =   "Formy19.frx":0000
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   2055
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   46
      Text            =   "Formy19.frx":0006
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      ScrollBars      =   3  'Both
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   3495
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
      Left            =   1680
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1095
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   0
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   10200
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "尺码打印"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
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
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
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
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
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
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
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
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "尺码表"
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
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy19.frx":000C
      Height          =   330
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo18 
      Bindings        =   "Formy19.frx":0020
      Height          =   330
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "xm"
      Text            =   "DBCombo18"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy19.frx":0034
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5400
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   12
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   32896
      WordWrap        =   -1  'True
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81133569
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   10200
      TabIndex        =   14
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81133569
      CurrentDate     =   39177
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   1
      Left            =   960
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   2
      Left            =   3600
      TabIndex        =   16
      Top             =   6960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   3
      Left            =   3600
      TabIndex        =   17
      Top             =   7440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   4
      Left            =   3600
      TabIndex        =   18
      Top             =   7920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   5
      Left            =   960
      TabIndex        =   19
      Top             =   3960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   6
      Left            =   12360
      TabIndex        =   20
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   7
      Left            =   4560
      TabIndex        =   21
      Top             =   2520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   8
      Left            =   960
      TabIndex        =   22
      Top             =   3000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   9
      Left            =   12360
      TabIndex        =   23
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   10
      Left            =   4560
      TabIndex        =   24
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   11
      Left            =   4560
      TabIndex        =   25
      Top             =   3480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   12
      Left            =   4560
      TabIndex        =   26
      Top             =   3960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   13
      Left            =   1560
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   14
      Left            =   960
      TabIndex        =   28
      Top             =   3480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   15
      Left            =   10920
      TabIndex        =   50
      Top             =   7080
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   16
      Left            =   10920
      TabIndex        =   51
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   17
      Left            =   4560
      TabIndex        =   53
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   18
      Left            =   120
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy19.frx":0048
      Height          =   2295
      Left            =   10920
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4048
      _Version        =   393216
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      FormatString    =   "记录号 "
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "接单日期"
      Height          =   375
      Left            =   120
      TabIndex        =   61
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "生产单号"
      Height          =   375
      Left            =   3720
      TabIndex        =   59
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "样品单号"
      Height          =   375
      Left            =   3720
      TabIndex        =   58
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "唛头"
      Height          =   375
      Index           =   16
      Left            =   10080
      TabIndex        =   54
      Top             =   7200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "包装"
      Height          =   375
      Index           =   15
      Left            =   10080
      TabIndex        =   52
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "外箱唛头"
      Height          =   255
      Index           =   13
      Left            =   9840
      TabIndex        =   49
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "包装方式"
      Height          =   255
      Index           =   12
      Left            =   6840
      TabIndex        =   48
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  成衣生产工艺通知单"
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
      TabIndex        =   45
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "刀模"
      Height          =   375
      Left            =   11520
      TabIndex        =   44
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "日期与时间"
      Height          =   375
      Left            =   3720
      TabIndex        =   43
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Left            =   4680
      TabIndex        =   42
      Top             =   960
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   7560
      X2              =   8760
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Left            =   9120
      TabIndex        =   41
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "选择负责人"
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
      Index           =   13
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   375
      Left            =   3720
      TabIndex        =   39
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "状态"
      Height          =   375
      Index           =   7
      Left            =   11520
      TabIndex        =   38
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "件数"
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   37
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   36
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款式"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   35
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "系数"
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   34
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "工作编号"
      Height          =   375
      Index           =   8
      Left            =   2760
      TabIndex        =   32
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "交期"
      Height          =   375
      Index           =   9
      Left            =   3720
      TabIndex        =   31
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   10
      Left            =   3720
      TabIndex        =   30
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "面料"
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   29
      Top             =   3480
      Width           =   855
   End
End
Attribute VB_Name = "Formy19"
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
If DBCombo18.Text = "" Then
MsgBox ("请选择负责人！")
Exit Sub
End If
If DBCombo1(4).Text = "" Then DBCombo1(4).Text = 0

DBCombo1(13).Text = DBCombo18.Text
DBCombo1(6).Text = "开始"

rd.AddNew
For i = 0 To rd.Fields.Count - 1
rd.Fields(i) = DBCombo1(i).Text
Next
rd.Fields(15) = Text1.Text
rd.Fields(16) = Text3.Text
rd.Fields(18) = "进行"
rd.Update

Data1.Refresh
Data2.Refresh
Data3.Refresh

For i = 0 To Data1.Recordset.Fields.Count - 1
DBCombo1(i).Text = ""
Next

DBCombo1(12).Text = 1
DBCombo1(5).Text = Date
DBCombo1(11).Text = Date
DBCombo1(12).Text = Data2.Recordset.Fields(0) + 1

DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYYYMMDD")) + "1"
If Data3.Recordset.EOF Then
DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYYYMMDD")) + "1"
Else
DBCombo1(7).Text = "LDH" + Trim(Data3.Recordset.Fields(0) + 1)
End If

End Sub

Private Sub Command2_Click()
On Error Resume Next

If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If

If DBCombo1(4).Text = "" Then DBCombo1(4).Text = 0
DBCombo1(13).Text = DBCombo18.Text
Data1.Recordset.Edit
For i = 0 To Data1.Recordset.Fields.Count - 1
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Fields(15) = Text1.Text
Data1.Recordset.Fields(16) = Text3.Text
Data1.Recordset.Update

Data1.Refresh
Data2.Refresh
Data3.Refresh

DBCombo1(13).Text = DBCombo18.Text

For i = 0 To Data1.Recordset.Fields.Count - 1
DBCombo1(i).Text = ""
Next
DBCombo1(12).Text = 1
DBCombo1(5).Text = Date
DBCombo1(11).Text = Date
DBCombo1(12).Text = Data2.Recordset.Fields(0) + 1

DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYYYMMDD")) + "1"
If Data3.Recordset.EOF Then
DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYYYMMDD")) + "1"
Else
DBCombo1(7).Text = "LDH" + Trim(Data3.Recordset.Fields(0) + 1)
End If


End Sub

Private Sub Command4_Click()

On Error Resume Next

If MsgBox("确定删除吗?,删除将不能恢复!", vbYesNo) = vbNo Then Exit Sub


Data1.Recordset.Delete
Data1.Refresh
Data2.Refresh
Data3.Refresh

For i = 0 To Data1.Recordset.Fields.Count - 1
DBCombo1(i).Text = ""
Next
DBCombo1(12).Text = 1
DBCombo1(5).Text = Date
DBCombo1(11).Text = Date
DBCombo1(12).Text = Data2.Recordset.Fields(0) + 1

DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYYYMMDD")) + "1"
If Data3.Recordset.EOF Then
DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYYYMMDD")) + "1"
Else
DBCombo1(7).Text = "LDH" + Trim(Data3.Recordset.Fields(0) + 1)
End If

End Sub


Private Sub Command6_Click()
Call DGYDOutDataToExcel(Data1, Data2, DBCombo1(7).Text)
Data1.RecordSource = "select * from SCZY_ZDH WHERE INSTR(SCZY_ZDH.单号,'L')>0   AND SCZY_ZDH.日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "') ORDER BY VAL(SCZY_ZDH.序号) DESC"
Data1.Refresh
End Sub

Private Sub Command7_Click()
If DBCombo1(17).Text = "" Then
MsgBox ("请输入样品单号")
Exit Sub
End If
Data10.Database.Execute "DELETE * FROM SCZY_XDH WHERE 单号='" & DBCombo1(7).Text & "'"
If MsgBox("确定由此样品生成大货吗？", vbYesNo) = vbNo Then Exit Sub
Data10.Database.Execute "INSERT INTO SCZY_XDH(款号,颜色,数量,刀模,双针,汉带,前页,绣眼,帽芯,间线,胶条,高头明线,后扣,顶扣,商标,具体要求,绣印说明,交货期,序号,工作编号) SELECT 款号,颜色,数量,刀模,双针,汉带,前页,绣眼,帽芯,间线,胶条,高头明线,后扣,顶扣,商标,具体要求,绣印说明,交货期,序号,工作编号 FROM SCZY_X WHERE 单号='" & DBCombo1(17).Text & "' AND 数量>0"
Data10.Database.Execute "UPDATE SCZY_XDH SET 单号='" & DBCombo1(7).Text & "' WHERE 单号=NULL"
MsgBox ("生成成功！")
End Sub

Private Sub Command8_Click()
On Error Resume Next
Data1.RecordSource = "select * from SCZY_ZDH WHERE INSTR(SCZY_ZDH.单号,'L')>0 AND SCZY_ZDH.日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "') ORDER BY VAL(SCZY_ZDH.序号) DESC"
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data7.Refresh
Data8.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
DBCombo1(12).Text = 1
DBCombo1(5).Text = Date
DBCombo1(12).Text = Data2.Recordset.Fields(0) + 1

Data3.RecordSource = "select MAX(VAL(MID(SCZY_ZDH.单号,10))) from SCZY_ZDH WHERE INSTR(SCZY_ZDH.单号,'L')>0 AND SCZY_ZDH.日期=CDATE('" & Date & "')"
Data3.Refresh
DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYMMDD")) + "1"
If Data3.Recordset.EOF Then
DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYMMDD")) + "1"
Else
DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYMMDD")) + Trim(Data3.Recordset.Fields(0) + 1)
End If

Data9.Refresh
End Sub




Private Sub DBCombo10_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

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




Private Sub DBCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo4_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo5_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

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

Private Sub DBCombo9_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub


Private Sub Command9_Click()
Call BTDY(Data2, DBCombo1(7).Text)
End Sub

Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 17
        HYH = ""
        Data4.RecordSource = "select * from SCZY_Z WHERE SCZY_Z.单号='" & DBCombo1(17).Text & "'"
        Data4.Refresh
        For i = 0 To 1
        DBCombo1(i).Text = Data4.Recordset.Fields(i)
        Next
        DBCombo1(9).Text = Data4.Recordset.Fields(9)
        DBCombo1(8).Text = Data4.Recordset.Fields(8)
        HYH = Data4.Recordset.Fields(7)
End Select
End Sub

Private Sub DBCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo18_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.Value

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

Private Sub Form_Load()
On Error Resume Next

Set ba = OpenDatabase("d:\数据库\\htgl\2011\SCZYJHD.MDB")
Set rd = ba.OpenRecordset("SCZY_ZDH", dbOpenDynaset)

Text1.Text = ""
Text3.Text = ""
Text4.Text = Date - 15
Text5.Text = Date
DTPicker1.Value = Date - 15
DTPicker2.Value = Date

DBCombo18.Text = ""

Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from SCZY_ZDH WHERE INSTR(SCZY_ZDH.单号,'L')>0   AND SCZY_ZDH.日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "') ORDER BY SCZY_ZDH.序号 DESC"
Data1.Refresh

For i = 0 To 18
DBCombo1(i).Text = ""
Next

Data10.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"

Data11.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data12.DatabaseName = "d:\数据库\\htgl\2011\XHXX.MDB"
Data13.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data14.DatabaseName = "d:\数据库\\htgl\2011\XHXX.MDB"

Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "select MAX(VAL(SCZY_ZDH.序号)) from SCZY_ZDH"
Data2.Refresh


Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select MAX(VAL(MID(SCZY_ZDH.单号,10))) from SCZY_ZDH WHERE INSTR(SCZY_ZDH.单号,'L')>0 AND SCZY_ZDH.日期=CDATE('" & Date & "')"
Data3.Refresh
DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYMMDD")) + "1"
If Data3.Recordset.EOF Then
DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYMMDD")) + "1"
Else
DBCombo1(7).Text = "LDH" + Trim(Format(Date, "YYMMDD")) + Trim(Data3.Recordset.Fields(0) + 1)
End If


Data7.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data7.RecordSource = "SELECT YWF.xm from ywf GROUP BY YWF.XM"
Data7.Refresh


Data8.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data8.RecordSource = "SELECT 简称  from khzl GROUP BY 简称"
Data8.Refresh


DBCombo1(12).Text = 1
DBCombo1(5).Text = Date
DBCombo1(11).Text = Date
DBCombo1(12).Text = Data2.Recordset.Fields(0) + 1


Data4.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data4.RecordSource = "select * from SCZY_Z WHERE SCZY_Z.单号='" & DBCombo1(17).Text & "'"
Data4.Refresh


Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False


MSFlexGrid1.ColWidth(0) = 100
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 1200
MSFlexGrid1.ColWidth(4) = 1500
MSFlexGrid1.ColWidth(5) = 1000
MSFlexGrid1.ColWidth(6) = 1200
MSFlexGrid1.ColWidth(7) = 0
MSFlexGrid1.ColWidth(8) = 1200
MSFlexGrid1.ColWidth(9) = 1200
MSFlexGrid1.ColWidth(10) = 1200
MSFlexGrid1.ColWidth(11) = 1200
MSFlexGrid1.ColWidth(12) = 1200
MSFlexGrid1.ColWidth(19) = 0
MSFlexGrid1.ColWidth(20) = 0
MSFlexGrid1.ColWidth(21) = 0
MSFlexGrid1.ColWidth(22) = 0
MSFlexGrid1.ColWidth(23) = 0
MSFlexGrid1.ColWidth(24) = 0
MSFlexGrid1.ColWidth(25) = 0
MSFlexGrid1.ColWidth(26) = 0
MSFlexGrid1.ColWidth(27) = 0
MSFlexGrid1.ColWidth(28) = 0
MSFlexGrid1.ColWidth(29) = 0

MSFlexGrid2.ColWidth(1) = 2000

DBCombo18.TabIndex = 0
Data9.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data9.RecordSource = "SELECT 工作编号  from SCZY_ZDH ORDER BY 工作编号 DESC"
Data9.Refresh

End Sub


Private Sub Label10_Click()
DBCombo1(5).Enabled = True
End Sub

Private Sub Label10_DblClick()
DBCombo1(5).Enabled = False
End Sub

Private Sub Label2_DBLClick()
Formy8.Show
End Sub

Private Sub Label3_dblClick(Index As Integer)
Select Case Index
       Case 7
DBCombo6.Enabled = True
End Select
End Sub

Private Sub Label8_Click()
DBCombo1(7).Enabled = False
End Sub

Private Sub Label8_DblClick()
DBCombo1(7).Enabled = True
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row

If Data1.Recordset.EOF Then
DBCombo18.Text = ""
Exit Sub
End If

Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To Data1.Recordset.Fields.Count - 1
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next
DBCombo18.Text = Data1.Recordset.Fields(13)
Text1.Text = Data1.Recordset.Fields(15)
Text3.Text = Data1.Recordset.Fields(16)

Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Timer1_Timer()
Text2.Text = Now
End Sub


