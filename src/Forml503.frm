VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Forml503 
   BackColor       =   &H00C0E0FF&
   Caption         =   "加工入库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form42"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data13 
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data12 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data11 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data10 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   360
      TabIndex        =   70
      Top             =   1920
      Width           =   2775
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "进行"
         Height          =   375
         Left            =   240
         TabIndex        =   72
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1560
         TabIndex        =   71
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command10 
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
      TabIndex        =   69
      Top             =   2880
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Forml503.frx":0000
      Height          =   2295
      Left            =   3600
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   240
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   25
      Left            =   8400
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   24
      Left            =   4920
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   23
      Left            =   7800
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   22
      Left            =   7800
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   21
      Left            =   4800
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   20
      Left            =   6840
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   19
      Left            =   5880
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   18
      Left            =   6840
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
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
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4440
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4440
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4440
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
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
      Height          =   735
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   3600
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   3600
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   11280
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   10080
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   12360
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   8160
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Width           =   4455
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   240
      Width           =   1095
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   10080
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   4800
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   6960
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   12
      Left            =   13440
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   13
      Left            =   6960
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导出"
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
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   14
      Left            =   8400
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   15
      Left            =   11280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command4 
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   16
      Left            =   12360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3960
      Width           =   975
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   17
      Left            =   5880
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   360
      TabIndex        =   27
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Forml503.frx":0014
      Height          =   3855
      Left            =   3600
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5520
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   13680
      TabIndex        =   30
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   36892
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Forml503.frx":0028
      Height          =   330
      Left            =   4800
      TabIndex        =   31
      Top             =   4920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Forml503.frx":003C
      Height          =   330
      Left            =   4800
      TabIndex        =   52
      Top             =   3960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Forml503.frx":0050
      Height          =   330
      Left            =   8400
      TabIndex        =   68
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DBCombo4"
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转入"
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
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6015
      Left            =   360
      TabIndex        =   73
      Top             =   3360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   10610
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1560
      TabIndex        =   74
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1560
      TabIndex        =   75
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   39557
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
      Index           =   28
      Left            =   360
      TabIndex        =   77
      Top             =   1560
      Width           =   1215
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
      Index           =   27
      Left            =   360
      TabIndex        =   76
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "印花单位"
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
      Index           =   26
      Left            =   8400
      TabIndex        =   67
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "印花金额"
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
      Index           =   25
      Left            =   4920
      TabIndex        =   65
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "印花数量"
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
      Index           =   24
      Left            =   7800
      TabIndex        =   63
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "印花单价"
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
      Index           =   23
      Left            =   7800
      TabIndex        =   61
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "织布金额"
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
      Index           =   22
      Left            =   6840
      TabIndex        =   58
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "织布单价"
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
      Index           =   21
      Left            =   5880
      TabIndex        =   56
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染色金额"
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
      Left            =   6840
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "织布单位"
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
      Left            =   4800
      TabIndex        =   51
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "光坯数量"
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
      Left            =   3600
      TabIndex        =   50
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯数量"
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
      Left            =   3600
      TabIndex        =   49
      Top             =   3600
      Width           =   1095
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
      Index           =   9
      Left            =   11280
      TabIndex        =   48
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料"
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
      Left            =   8160
      TabIndex        =   47
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "光坯克重"
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
      Left            =   12360
      TabIndex        =   46
      Top             =   2760
      Width           =   855
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
      Left            =   10080
      TabIndex        =   45
      Top             =   2760
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
      Left            =   3600
      TabIndex        =   44
      Top             =   2760
      Width           =   1695
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
      Left            =   5400
      TabIndex        =   43
      Top             =   2760
      Width           =   1455
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
      Left            =   6960
      TabIndex        =   42
      Top             =   2760
      Width           =   1095
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
      Index           =   0
      Left            =   360
      TabIndex        =   41
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "光坯匹数"
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
      Left            =   10080
      TabIndex        =   40
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染色单位"
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
      Left            =   4800
      TabIndex        =   39
      Top             =   4440
      Width           =   2055
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
      Index           =   12
      Left            =   6960
      TabIndex        =   38
      Top             =   4440
      Width           =   1335
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
      Index           =   13
      Left            =   13440
      TabIndex        =   37
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染厂锅号"
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
      Left            =   6960
      TabIndex        =   36
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染厂色别"
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
      Left            =   8400
      TabIndex        =   35
      Top             =   3600
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
      Index           =   16
      Left            =   11280
      TabIndex        =   34
      Top             =   3600
      Width           =   855
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
      Index           =   17
      Left            =   12360
      TabIndex        =   33
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染色单价"
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
      Index           =   18
      Left            =   5880
      TabIndex        =   32
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Forml503"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Data2.RecordSource = "SELECT 单号,款号,颜色,材料名称,光坯克重,光坯幅宽,材料数量,织布单位,染色单位,印花单位,回厂期限 FROM zbfl WHERE 单号='" & DBCombo1.Text & "'"
Data2.Refresh
If DBCombo1.Text = "" Then
Data4.RecordSource = "select * from rsrk where  日期=cdate('" & Text1(12).Text & "') order by 日期,染色单位,材料名称"
Data4.Refresh
Else
Data4.RecordSource = "select * from rsrk where  单号='" & DBCombo1.Text & "' order by 日期,染色单位,材料名称"
Data4.Refresh
End If
Data9.RecordSource = "select 序号 from rsrk where 单号='" & DBCombo1.Text & "' order by 序号 desc"
Data9.Refresh
Text1(16).Text = 1
If Data9.Recordset.EOF Then
Text1(16).Text = 1
Else
Text1(16).Text = Data9.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command10_Click()
Call tree
Call zk
End Sub

Private Sub Command2_Click()
'On Error Resume Next
Data8.RecordSource = "SELECT MAX(VAL(MID(ckgl.单据号,2))) FROM CKGL WHERE INSTR(单据号,'Z')>0"
Data8.Refresh
l1 = "Z0000001"
If Data8.Recordset.EOF Then
l1 = "Z0000001"
Else
l1 = Left("Z000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If
lo = "d:\数据库\\htgl\2011\ckgl.mdb"
Data3.Database.Execute "insert into ckgl(单号,合约号,库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,库别,供应单位,日期,单据号,序号) in'" & lo & "' select 单号,款号,'1主料库',材料名称,光坯幅宽,'公斤',颜色,光坯克重,光坯重量,'0','0','采购入库',染色单位,日期,'" & l1 & "','1' from rsrk where 转入='N'"
Data3.Database.Execute "update rsrk set 转入='Y' where 转入='N'"
MsgBox ("转入成功！")
End Sub

Private Sub Command3_Click()
Call ZLRK(MSFlexGrid1, "主料入库")
End Sub

Private Sub Command4_Click()
Call rsrk(MSFlexGrid1, "染色入库")
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
On Error Resume Next
If MsgBox("删除不能回复！确定删除吗？", vbYesNo) = vbNo Then Exit Sub
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.Delete
Data4.Refresh
Text1(4).Text = ""
Text1(5).Text = ""
Text1(15).Text = "公斤"
Data9.RecordSource = "select max(序号) from rsrk where 单号='" & DBCombo1.Text & "'"
Data9.Refresh
Text1(16).Text = 1
If Data9.Recordset.EOF Then
Text1(16).Text = 1
Else
Text1(16).Text = Data9.Recordset.Fields(0) + 1
End If
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command7_Click()
On Error Resume Next
If Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(8).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data4.Recordset.AddNew
For i = 0 To 25
Data4.Recordset.Fields(i) = Text1(i).Text
Next
Data4.Recordset.Update
Data4.Refresh
Text1(15).Text = "公斤"
Text1(4).Text = ""
Text1(5).Text = ""
Data9.RecordSource = "select max(序号) from rsrk where 单号='" & DBCombo1.Text & "'"
Data9.Refresh
Text1(16).Text = 1
If Data9.Recordset.EOF Then
Text1(16).Text = 1
Else
Text1(16).Text = Data9.Recordset.Fields(0) + 1
End If
Text1(4).SetFocus
End Sub

Private Sub Command8_Click()
On Error Resume Next
If Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(8).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data4.Recordset.Edit
For i = 0 To 16
Data4.Recordset.Fields(i) = Text1(i).Text
Next
Data4.Recordset.Update
Data4.Refresh

Text1(4).Text = ""
Text1(5).Text = ""
Data9.RecordSource = "select max(序号) from rsrk where 单号='" & DBCombo1.Text & "'"
Data9.Refresh
Text1(16).Text = 1
If Data9.Recordset.EOF Then
Text1(16).Text = 1
Else
Text1(16).Text = Data9.Recordset.Fields(0) + 1
End If
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command9_Click()
On Error Resume Next
If DBCombo1.Text = "" Then
Data4.RecordSource = "select * from rsrk where  日期=cdate('" & Text1(12).Text & "') order by 日期,染色单位,材料名称"
Data4.Refresh
Else
Data4.RecordSource = "select * from rsrk where  单号='" & DBCombo1.Text & "' order by 日期,染色单位,材料名称"
Data4.Refresh
End If
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(15).Text = "公斤"
Text1(4).SetFocus
Data9.RecordSource = "select 序号 from rsrk where 单号='" & DBCombo1.Text & "' order by 序号 desc"
Data9.Refresh
Text1(16).Text = 1
If Data9.Recordset.EOF Then
Text1(16).Text = 1
Else
Text1(16).Text = Data9.Recordset.Fields(0) + 1
End If
End Sub

Private Sub DBCombo1_Change()
On Error Resume Next
Text1(12).Text = Date
Data2.RecordSource = "SELECT 单号,款号,颜色,材料名称,光坯克重,光坯幅宽,材料数量,织布单位,染色单位,印花单位,回厂期限 FROM zbfl WHERE 单号='" & DBCombo1.Text & "'"
Data2.Refresh
If DBCombo1.Text = "" Then
Data4.RecordSource = "select * from rsrk where  日期=cdate('" & Text1(12).Text & "') order by 日期,染色单位,材料名称"
Data4.Refresh
Else
Data4.RecordSource = "select * from rsrk where  单号='" & DBCombo1.Text & "' order by 日期,染色单位,材料名称"
Data4.Refresh
End If
Data9.RecordSource = "select 序号 from rsrk where 单号='" & DBCombo1.Text & "' order by 序号 desc"
Data9.Refresh
Text1(16).Text = 1
If Data9.Recordset.EOF Then
Text1(16).Text = 1
Else
Text1(16).Text = Data9.Recordset.Fields(0) + 1
End If

End Sub

Private Sub DBCombo2_Click(Area As Integer)
Text1(21).Text = DBCombo2.Text
End Sub

Private Sub DBCombo3_Click(Area As Integer)
Text1(10).Text = DBCombo3.Text
End Sub

Private Sub DTPicker1_Change()
Text1(12).Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text1(12).Text = DTPicker1.Value
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""
DBCombo4.Text = ""
For i = 0 To 25
Text1(i).Text = ""
Next
Text1(12).Text = Date
DTPicker1.Value = Date
Text1(15).Text = "公斤"
Text1(17).Text = 0
DTPicker3.Value = Date - 30
DTPicker4.Value = Date
Option4.Value = True
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"

Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data5.RecordSource = "select 简称 from gys where instr(代码,'织')>0 group by 简称"
Data5.Refresh

Data6.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data6.RecordSource = "select 简称 from gys where instr(代码,'染')>0 group by 简称"
Data6.Refresh

Data7.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data8.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data9.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"

Data10.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data11.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data12.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

Data13.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data13.RecordSource = "select 简称 from gys where instr(代码,'印')>0 group by 简称"
Data13.Refresh

Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(4) = 1600
MSFlexGrid1.ColWidth(12) = 1300
For i = 18 To 26
MSFlexGrid1.ColWidth(i) = 0
Next
MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(4) = 1800
End Sub
Private Sub DBCombo4_Click(Area As Integer)
Text1(25).Text = DBCombo4.Text
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data4.Recordset.Move rs - 1
For i = 0 To 25
Text1(i).Text = Data4.Recordset.Fields(i)
Next
DTPicker1.Value = Text1(8).Text
Command7.Enabled = False
Command8.Enabled = True
Command6.Enabled = True
End Sub

Private Sub MSFlexGrid2_dblClick()
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
For i = 0 To 4
Text1(i).Text = Data2.Recordset.Fields(i)
Next
Text1(6).Text = Data2.Recordset.Fields(5)
Text1(10).Text = Data2.Recordset.Fields(8)
Text1(21).Text = Data2.Recordset.Fields(7)
Text1(25).Text = Data2.Recordset.Fields(9)
End Sub
Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 7
If Val(Text1(7).Text) > 0 Then
Text1(11).Text = Format((Val(Text1(7).Text) - Val(Text1(8))) / Val(Text1(7).Text) * 100, "#0.00")
End If
       Case 8
If Val(Text1(7).Text) > 0 Then
Text1(11).Text = Format((Val(Text1(7).Text) - Val(Text1(8))) / Val(Text1(7).Text) * 100, "#0.00")
End If

End Select
End Sub

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
DBCombo1.Text = l1
End If


'DBCombo2.Text = Node.Index
'DBCombo3.Text = TreeView1.Nodes(Node.Index).FullPath
End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub


Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex
   TreeView1.Nodes.Clear
 

If Option4.Value = True Then
    Data12.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data10.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
        Data10.Refresh
        
        If Not Data10.Recordset.EOF Then
        Data10.Recordset.MoveFirst
        Do While Not Data10.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data10.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data11.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data10.Recordset.Fields(0) & "' and 进度='进行'"
        Data11.Refresh
        
        If Not Data11.Recordset.EOF Then
        Data11.Recordset.MoveFirst
        Do While Not Data11.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data11.Recordset.Fields(0))
        Data11.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        
        Data10.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        Data12.Recordset.MoveNext
        m = m + 1
        Loop
    End If
End If


If Option5.Value = True Then
    Data12.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='结束'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(, , Data12.Recordset.Fields(0), Data12.Recordset.Fields(0))
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data10.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='结束'"
        Data10.Refresh
        
        If Not Data10.Recordset.EOF Then
        Data10.Recordset.MoveFirst
        Do While Not Data10.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data10.Recordset.Fields(0))
        intIndex = mNode.Index
        Data11.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data10.Recordset.Fields(0) & "' and 进度='结束'"
        Data11.Refresh
        
        If Not Data11.Recordset.EOF Then
        Data11.Recordset.MoveFirst
        Do While Not Data11.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data11.Recordset.Fields(0))
        Data11.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data10.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data12.Recordset.MoveNext
        Loop
    End If
End If

End Sub



