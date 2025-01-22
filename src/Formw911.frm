VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw911 
   BackColor       =   &H00C0E0FF&
   Caption         =   "成品退库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新单据"
      Height          =   375
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Data Data7 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command11 
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw911.frx":0000
      Height          =   6015
      Left            =   480
      TabIndex        =   7
      Top             =   3480
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   10610
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   0
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   2160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   2
      Left            =   1920
      TabIndex        =   10
      Top             =   2760
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   3
      Left            =   6120
      TabIndex        =   11
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   4
      Left            =   6120
      TabIndex        =   12
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw911.frx":0014
      Height          =   390
      Index           =   5
      Left            =   6120
      TabIndex        =   13
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   6
      Left            =   6120
      TabIndex        =   14
      Top             =   2160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   7
      Left            =   10680
      TabIndex        =   15
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   8
      Left            =   10680
      TabIndex        =   16
      Top             =   1560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81199105
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9120
      TabIndex        =   18
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81199105
      CurrentDate     =   39557
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   9
      Left            =   1920
      TabIndex        =   19
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   10
      Left            =   10680
      TabIndex        =   20
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   11
      Left            =   10680
      TabIndex        =   35
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "条码扫描"
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
      Index           =   5
      Left            =   480
      TabIndex        =   37
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单据"
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
      Index           =   1
      Left            =   9240
      TabIndex        =   36
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单号"
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
      Index           =   7
      Left            =   480
      TabIndex        =   33
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "条码"
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
      Index           =   6
      Left            =   9240
      TabIndex        =   32
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
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
      Index           =   4
      Left            =   8160
      TabIndex        =   31
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
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
      Index           =   0
      Left            =   5640
      TabIndex        =   30
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "数量"
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
      Index           =   3
      Left            =   4680
      TabIndex        =   29
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
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
      Index           =   3
      Left            =   480
      TabIndex        =   28
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格"
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
      Index           =   2
      Left            =   4680
      TabIndex        =   27
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "单位"
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
      Index           =   2
      Left            =   4680
      TabIndex        =   26
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
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
      Index           =   1
      Left            =   480
      TabIndex        =   25
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
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
      Index           =   0
      Left            =   4680
      TabIndex        =   24
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
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
      Index           =   4
      Left            =   9240
      TabIndex        =   23
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
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
      Index           =   5
      Left            =   9240
      TabIndex        =   22
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "款号"
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
      TabIndex        =   21
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "Formw911"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer
Private Sub Command1_Click()
If DBCombo1(2).Text = "" Or DBCombo1(6).Text = "" Then
Exit Sub
End If
Data1.Recordset.AddNew
For i = 0 To 11
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh
DBCombo1(6).Text = 0
Data6.RecordSource = "SELECT max(序号) FROM LSTK WHERE 单据='" & DBCombo1(11).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If

End Sub


Private Sub Command11_Click()
On Error Resume Next
Data1.RecordSource = "SELECT * FROM LSTK WHERE 单据='" & DBCombo1(11).Text & "' ORDER BY 序号 DESC"
Data1.Refresh
Data6.RecordSource = "SELECT max(序号) FROM LSTK WHERE 单据='" & DBCombo1(11).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If
DBCombo1(6).Text = 0
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If DBCombo1(2).Text = "" Or DBCombo1(6).Text = "" Then
Exit Sub
End If
If MsgBox("确定修改吗", vbYesNo) = vbNo Then Exit Sub
If DBCombo1(0).Text = "" Then Exit Sub
Data1.Recordset.Edit
For i = 0 To 11
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh
DBCombo1(6).Text = 0
Data6.RecordSource = "SELECT max(序号) FROM LSTK WHERE 单据='" & DBCombo1(11).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If

End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除吗", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
Data6.RecordSource = "SELECT max(序号) FROM LSTK WHERE 单据='" & DBCombo1(11).Text & "'"
Data6.Refresh
DBCombo1(6).Text = 0
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command5_Click()
Call OutDataToExcel(MSFlexGrid1, 7, DBCombo1(0).Text)
End Sub



Private Sub Command7_Click()
If DBCombo1(2).Text <> "" Then
Data1.RecordSource = "SELECT * FROM LSTK WHERE 品名='" & DBCombo1(2).Text & "' order by 日期,单号,序号"
Data1.Refresh
Else
Data1.RecordSource = "SELECT * FROM LSTK WHERE  日期 between CDATE('" & DTPicker1.Value & "') and CDATE('" & DTPicker2.Value & "') order by 日期,单号,序号"
Data1.Refresh
End If
End Sub


Private Sub Command8_Click()
On Error Resume Next
Data7.DatabaseName = "d:\数据库\\htgl\2011\cpck.MDB"
Data7.RecordSource = "select MAX(VAL(MID(单据,10))) from LSTK WHERE  日期=CDATE('" & DBCombo1(0).Text & "')"
Data7.Refresh
DBCombo1(11).Text = "TDH" + Trim(Format(CDate(DBCombo1(0).Text), "YYMMDD")) + "1"
If Data7.Recordset.EOF Then
DBCombo1(11).Text = "TDH" + Trim(Format(CDate(DBCombo1(0).Text), "YYMMDD")) + "1"
Else
DBCombo1(11).Text = "TDH" + Trim(Format(CDate(DBCombo1(0).Text), "YYMMDD")) + Trim(Data7.Recordset.Fields(0) + 1)
End If
Data6.RecordSource = "SELECT max(序号) FROM LSTK WHERE 单据='" & DBCombo1(11).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If
DBCombo1(6).Text = 0
Data1.RecordSource = "SELECT * FROM LSTK WHERE 单据='" & DBCombo1(11).Text & "' ORDER BY 序号 DESC"
Data1.Refresh
End Sub

Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data1.RecordSource = "SELECT * FROM LSTK WHERE 日期=CDATE('" & DBCombo1(0).Text & "')"
Data1.Refresh
End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data1.RecordSource = "SELECT * FROM LSTK WHERE 日期=CDATE('" & DBCombo1(0).Text & "')"
Data1.Refresh
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
S1 = 0
S2 = 0
For i = 0 To 11
DBCombo1(i).Text = ""
Next
Text1.Text = ""
DBCombo1(0).Text = Date
DBCombo1(5).Text = "件"
DTPicker1.Value = Date
DTPicker2.Value = Date
Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select 简称 from khzl GROUP BY 简称"
Data3.Refresh
Data4.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data4.RecordSource = "select MC from CLDW GROUP BY MC"
Data4.Refresh
Data5.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data6.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data6.RecordSource = "SELECT max(序号) FROM LSTK WHERE 单据='" & DBCombo1(11).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If

Data7.DatabaseName = "d:\数据库\\htgl\2011\cpck.MDB"
Data7.RecordSource = "select MAX(VAL(MID(单据,10))) from LSTK WHERE  日期=CDATE('" & DBCombo1(0).Text & "')"
Data7.Refresh
DBCombo1(11).Text = "TDH" + Trim(Format(DBCombo1(0).Text, "YYMMDD")) + "1"
If Data7.Recordset.EOF Then
DBCombo1(11).Text = "TDH" + Trim(Format(DBCombo1(0).Text, "YYMMDD")) + "1"
Else
DBCombo1(11).Text = "TDH" + Trim(Format(DBCombo1(0).Text, "YYMMDD")) + Trim(Data7.Recordset.Fields(0) + 1)
End If

Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data1.RecordSource = "SELECT * FROM LSTK WHERE 单据='" & DBCombo1(11).Text & "' ORDER BY 序号 DESC"
Data1.Refresh


MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1200
For i = 1 To 5
MSFlexGrid1.ColWidth(i) = 1600
Next
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Label1_dblClick(Index As Integer)
Select Case Index
       Case 1
khbl = 7
Formw202.Text1.Text = DBCombo1(1).Text
Formw202.Show
End Select
End Sub

Private Sub Label3_Click()
khbl = 3
Formw99.Text1 = DBCombo1(1).Text
Formw99.Show
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data1.Recordset.EOF Then Exit Sub
rs = MSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To 11
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next
Command3.Enabled = True
Command4.Enabled = True
Command1.Enabled = False
End Sub
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid1.RowSel
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid1.RowSel
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then Exit Sub

If InStr(Text1.Text, "J") > 0 Then
m = Left(Text1.Text, Len(Text1.Text) - 1)
If Len(m) = 9 Then
Data7.RecordSource = "SELECT * FROM LSFH WHERE 条码='" & m & "'"
Data7.Refresh
If Not Data7.Recordset.EOF Then
DBCombo1(9).Text = Data7.Recordset.Fields(16)
DBCombo1(1).Text = Data7.Recordset.Fields(2)
DBCombo1(2).Text = Data7.Recordset.Fields(3)
DBCombo1(3).Text = Data7.Recordset.Fields(4)
DBCombo1(4).Text = Data7.Recordset.Fields(5)
DBCombo1(5).Text = Data7.Recordset.Fields(6)
DBCombo1(6).Text = Data7.Recordset.Fields(7)
DBCombo1(10).Text = Data7.Recordset.Fields(17)
End If
End If
Text1.Text = ""
Text1.SetFocus
End If
End Sub
