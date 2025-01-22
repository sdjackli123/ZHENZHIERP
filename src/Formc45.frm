VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formc45 
   BackColor       =   &H00C0E0FF&
   Caption         =   "备料出库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form45"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data13 
      Caption         =   "Data11"
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data12 
      Caption         =   "Data11"
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
      Top             =   10080
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
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   390
      Left            =   3840
      TabIndex        =   61
      Top             =   2040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   688
      _Version        =   393216
      Text            =   "DBCombo2"
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
   Begin VB.CommandButton Command8 
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   12
      Left            =   13320
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Data Data15 
      Caption         =   "Data11"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "备料表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   48
      Top             =   600
      Width           =   1815
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000C0C0&
         Caption         =   "采购料"
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
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C0C0&
         Caption         =   "库存料"
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
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "出库类别"
      Height          =   495
      Left            =   12240
      TabIndex        =   25
      Top             =   480
      Width           =   2655
      Begin VB.OptionButton Option7 
         BackColor       =   &H0000C0C0&
         Caption         =   "强制"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H0000C0C0&
         Caption         =   "正常"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   13320
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   13320
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   13320
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   13320
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   12480
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   11880
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   13320
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   10680
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   10680
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   10680
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   10680
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   10680
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "备料方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   8
      Top             =   600
      Width           =   1815
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "库存备料"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H0000C0C0&
         Caption         =   "采购备料"
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
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   13080
      Top             =   6720
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command10 
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
      Height          =   375
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新单据号"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
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
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "复位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   10560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1111 
      Height          =   495
      Left            =   11160
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   375
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
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
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   10800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3735
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   12240
      TabIndex        =   28
      Top             =   4920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formc45.frx":0000
      Height          =   1695
      Left            =   3120
      TabIndex        =   29
      Top             =   4920
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formc45.frx":0014
      Height          =   2895
      Left            =   3120
      TabIndex        =   30
      Top             =   6720
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5106
      _Version        =   393216
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Bindings        =   "Formc45.frx":0028
      Height          =   330
      Left            =   8640
      TabIndex        =   31
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "车间编号"
      Text            =   "DBCombo3"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc45.frx":003C
      Height          =   1815
      Left            =   3120
      TabIndex        =   32
      Top             =   2880
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   17
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   8640
      TabIndex        =   54
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7455
      Left            =   120
      TabIndex        =   56
      Top             =   2160
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   13150
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1200
      TabIndex        =   57
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
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
      Left            =   1200
      TabIndex        =   58
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   80543745
      CurrentDate     =   39177
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "单号"
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
      Left            =   3120
      TabIndex        =   62
      Top             =   2040
      Width           =   735
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
      Index           =   1
      Left            =   120
      TabIndex        =   60
      Top             =   1200
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
      Index           =   1
      Left            =   120
      TabIndex        =   59
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料来源"
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
      Left            =   7560
      TabIndex        =   53
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "库别"
      Height          =   375
      Index           =   8
      Left            =   12240
      TabIndex        =   52
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "原单"
      Height          =   375
      Index           =   6
      Left            =   12240
      TabIndex        =   47
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
      Height          =   375
      Index           =   7
      Left            =   12240
      TabIndex        =   46
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
      Height          =   375
      Index           =   6
      Left            =   12240
      TabIndex        =   45
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "合计金额"
      Height          =   375
      Index           =   5
      Left            =   11400
      TabIndex        =   44
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "库类"
      Height          =   375
      Index           =   5
      Left            =   12240
      TabIndex        =   43
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "批次"
      Height          =   375
      Index           =   4
      Left            =   10200
      TabIndex        =   42
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "单价"
      Height          =   375
      Index           =   4
      Left            =   11400
      TabIndex        =   41
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "单位"
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   40
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
      Height          =   375
      Index           =   3
      Left            =   10200
      TabIndex        =   39
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料"
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   38
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格"
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   37
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "库存"
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   36
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "数量"
      Height          =   375
      Index           =   0
      Left            =   12240
      TabIndex        =   35
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "领料车间："
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
      Left            =   7560
      TabIndex        =   34
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "单据号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   33
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "Formc45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2, KB As String
Public c, r, BAR, S2, S1 As Integer

Private Sub Command10_Click()
If Data6.Recordset.EOF Then
MsgBox ("此单据号中无记录，不能打印！")
Exit Sub
End If
BAR = 1
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command11_Click()
On Error Resume Next
If MsgBox("确定删除吗？删除将不能恢复！", vbYesNo) = vbNo Then Exit Sub
Data6.Recordset.Delete
Data6.Refresh
Call Command7_Click
Option1.Value = False
Option5.Value = False
For i = 0 To 12
Text1(i).Text = ""
Next

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If DBCombo1.Text = "" Then
MsgBox ("请确认材料来源")
Exit Sub
End If
If DBCombo2.Text = "" Then
MsgBox ("请确认单号")
Exit Sub
End If
Data3.Recordset.MoveFirst
Data3.Recordset.Move S1 - 1
p = S2 - S1 + 1
For II = 1 To p      ''''''''''''''''''''''''''''''''''''''''''''

Text3.Text = Data3.Recordset.Fields(5)
For i = 0 To 6
Text1(i).Text = Data3.Recordset.Fields(i)
Next
Text1(8).Text = Data3.Recordset.Fields(7)
Text1(5).Text = Data3.Recordset.Fields(5)
Text1(7).Text = Format(Data3.Recordset.Fields(5) * Data3.Recordset.Fields(6), "*0.00")
Text1(11).Text = Data3.Recordset.Fields(9)

Text1(9).Text = Date
Text1(12).Text = Data3.Recordset.Fields(8)
Data7.RecordSource = "SELECT MAX(VAL(KPD.序号)) FROM KPD WHERE 单据号='" & Text2.Text & "'"
Data7.Refresh
Text1(10).Text = 1
If Data7.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data7.Recordset.Fields(0) + 1
End If




If Option6.Value = True Then
If DBCombo5.Text = "" Then
MsgBox ("领料车间")
Exit Sub
End If

If Data3.Recordset.EOF Then
MsgBox ("无库存，不能进行！")
Exit Sub
End If

If Text3.Text = "" Then
MsgBox ("无出库量，不能进行！")
Exit Sub
End If
End If


If Option7.Value = True Then
If DBCombo5.Text = "" Then
MsgBox ("领料车间")
Exit Sub
End If

If Data3.Recordset.EOF Then
MsgBox ("无库存，不能进行！")
Exit Sub
End If

If Text3.Text = "" Then
MsgBox ("无出库量，不能进行！")
Exit Sub
End If


End If

Data6.Recordset.AddNew
Data6.Recordset.Fields(0) = ""
Data6.Recordset.Fields(1) = DBCombo2.Text
Data6.Recordset.Fields(2) = DBCombo1.Text
Data6.Recordset.Fields(3) = Text1(0).Text
Data6.Recordset.Fields(4) = Text1(1).Text
Data6.Recordset.Fields(5) = Text1(2).Text
Data6.Recordset.Fields(6) = Text1(3).Text
Data6.Recordset.Fields(7) = Text1(4).Text
Data6.Recordset.Fields(8) = Text1(5).Text
Data6.Recordset.Fields(9) = Text1(6).Text
Data6.Recordset.Fields(10) = Text1(7).Text
Data6.Recordset.Fields(11) = DBCombo3.Text
Data6.Recordset.Fields(12) = DTPicker3.Value
Data6.Recordset.Fields(13) = Text1(10).Text
Data6.Recordset.Fields(14) = Text2.Text
Data6.Recordset.Fields(15) = Text1(11).Text
Data6.Recordset.Fields(16) = Text1(8).Text
Data6.Recordset.Fields(17) = Text1(12).Text         '''''''''库别
Data6.Recordset.Fields(18) = DBCombo5.Text
Data6.Recordset.Fields(19) = "未"
Data6.Recordset.Fields(20) = "未"
Data6.Recordset.Update
Data6.Refresh
Data8.Refresh
If Data6.Recordset.RecordCount = 8 Then
Text2.Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If
Data3.Recordset.MoveNext
Next             '''''''''''''''''''''''''''''''''

If Data6.Recordset.RecordCount = 8 Then
If MsgBox("是否打印本单据？", vbYesNo) = vbNo Then
Text2.Text = "00000001"
If Data8.Recordset.EOF Then
Text2.Text = "00000001"
Else
Text2.Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If
Else       '''''''''''''''''''''''''''
Call Command10_Click
End If
End If
Call Command7_Click
Option1.Value = False
Option5.Value = False
For i = 0 To 12
Text1(i).Text = ""
Next
End Sub

Private Sub Command4_Click()
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
Data6.Recordset.Edit
Data6.Recordset.Fields(0) = ""
Data6.Recordset.Fields(1) = DBCombo2.Text
Data6.Recordset.Fields(3) = Text1(0).Text
Data6.Recordset.Fields(4) = Text1(1).Text
Data6.Recordset.Fields(5) = Text1(2).Text
Data6.Recordset.Fields(6) = Text1(3).Text
Data6.Recordset.Fields(7) = Text1(4).Text
Data6.Recordset.Fields(8) = Text1(5).Text
Data6.Recordset.Fields(9) = Text1(6).Text
Data6.Recordset.Fields(10) = Text1(7).Text
Data6.Recordset.Fields(12) = DTPicker3.Value
Data6.Recordset.Fields(13) = Text1(10).Text
Data6.Recordset.Fields(14) = Text2.Text
Data6.Recordset.Fields(15) = Text1(11).Text
Data6.Recordset.Fields(16) = Text1(8).Text
Data6.Recordset.Fields(17) = Text1(12).Text
Data6.Recordset.Fields(18) = DBCombo5.Text
Data6.Recordset.Update
Data6.Refresh
Data8.Refresh

Call Command7_Click
Option1.Value = False
Option5.Value = False
For i = 0 To 12
Text1(i).Text = ""
Next

End Sub





Private Sub Command7_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option5.Value = False
End Sub

Private Sub Command8_Click()
Call tree
Call zk
End Sub

Private Sub Command9_Click()
On Error Resume Next
Text2.Text = "00000001"

Data8.RecordSource = "SELECT MAX(VAL(KPD.单据号)) FROM KPD"
Data8.Refresh
If Data8.Recordset.EOF Then
Text2.Text = "00000001"
Else
Text2.Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

End Sub



Private Sub Form_Load()
On Error Resume Next
Combo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""
Text2.Text = ""
Text3.Text = ""
DBCombo5.Text = ""
DTPicker3.Value = Date - 30
DTPicker3.Value = Date

For i = 0 To 12
Text1(i).Text = ""
Next
Option6.Value = True

Data10.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

Data3.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data4.RecordSource = "select 简称  from KHZL group by 简称"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data5.RecordSource = "select ct.车间编号  from ct group by ct.车间编号 ORDER BY VAL(CT.车间编号)"
Data5.Refresh


Data7.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
Data7.RecordSource = "SELECT MAX(VAL(KPD.IP)) FROM KPD WHERE KPD.标签='" & DBCombo1.Text & "'"
Data7.Refresh

Data8.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
Data8.RecordSource = "SELECT MAX(VAL(KPD.单据号)) FROM KPD"
Data8.Refresh

Data11.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data12.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data13.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"


ProgressBar1.Visible = False
Timer1.Enabled = False
Text2.Enabled = False
Text2.Text = "00000001"
If Data8.Recordset.EOF Then
Text2.Text = "00000001"
Else
Text2.Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

Data6.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
Data6.RecordSource = "SELECT * FROM KPD WHERE KPD.单据号='" & Text2.Text & "' "
Data6.Refresh

MSFlexGrid4.ColWidth(0) = 400
MSFlexGrid4.ColWidth(1) = 1500
MSFlexGrid4.ColWidth(2) = 1200
MSFlexGrid4.ColWidth(3) = 0
MSFlexGrid4.ColWidth(4) = 0
MSFlexGrid4.ColWidth(6) = 0
MSFlexGrid4.ColWidth(8) = 1500
MSFlexGrid4.ColWidth(9) = 1200
MSFlexGrid4.ColWidth(10) = 0
MSFlexGrid4.ColWidth(11) = 0

MSFlexGrid2.ColWidth(0) = 400
MSFlexGrid2.ColWidth(1) = 1200
MSFlexGrid2.ColWidth(2) = 1200
MSFlexGrid2.ColWidth(3) = 1200
MSFlexGrid2.ColWidth(7) = 0

MSFlexGrid1.ColWidth(0) = 400
MSFlexGrid2.ColWidth(0) = 400
MSFlexGrid3.ColWidth(0) = 400
MSFlexGrid3.ColWidth(1) = 0
MSFlexGrid3.ColWidth(2) = 0
MSFlexGrid3.ColWidth(3) = 1200
MSFlexGrid3.ColWidth(10) = 0
MSFlexGrid3.ColWidth(11) = 0
MSFlexGrid3.ColWidth(12) = 0

For i = 10 To 17
MSFlexGrid2.ColWidth(i) = 0
Next


DBCombo1.Text = ""
End Sub



Private Sub Label3_Click()

End Sub



Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid2.RowSel
End Sub

Private Sub MSFlexGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid2.RowSel
End Sub

Private Sub MSFlexGrid3_DBLClick()
On Error Resume Next
rs = MSFlexGrid3.Row
Data6.Recordset.MoveFirst
Data6.Recordset.Move rs - 1
Text1(5).Text = Data6.Recordset.Fields(8)


DBCombo2.Text = Data6.Recordset.Fields(1)
DBCombo1.Text = Data6.Recordset.Fields(2)
Text1(0).Text = Data6.Recordset.Fields(3)
 Text1(1).Text = Data6.Recordset.Fields(4)
 Text1(2).Text = Data6.Recordset.Fields(5)
 Text1(3).Text = Data6.Recordset.Fields(6)
 Text1(4).Text = Data6.Recordset.Fields(7)
Text1(5).Text = Data6.Recordset.Fields(8)
 Text1(6).Text = Data6.Recordset.Fields(9)
 Text1(7).Text = Data6.Recordset.Fields(10)
 DBCombo3.Text = Data6.Recordset.Fields(11)
 Text1(9).Text = Data6.Recordset.Fields(12)
 Text1(10).Text = Data6.Recordset.Fields(13)
 Text2.Text = Data6.Recordset.Fields(14)
 Text1(11).Text = Data6.Recordset.Fields(15)
 Text1(12).Text = Data6.Recordset.Fields(17)
 Text1(8).Text = Data6.Recordset.Fields(16)
 DBCombo5.Text = Data6.Recordset.Fields(18)

End Sub

Private Sub Option1_Click()
On Error Resume Next
Data8.Database.Execute "DELETE * FROM CLRCZZLS"
Data8.Database.Execute "DELETE * FROM CLRCZZHZLS"
Data8.Database.Execute "INSERT INTO CLRCZZLS(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库型,单号) select KPD.材料名称,KPD.材料规格,KPD.材料单位,KPD.颜色,批次,KPD.数量,KPD.单价,KPD.库类,库别,备注 from KPD WHERE KPD.单号='" & DBCombo2.Text & "' and  标签='库存料'"
Data8.Database.Execute "UPDATE CLRCZZLS SET 库别='出库',数量=-数量 where 库别=NULL"
Data8.Database.Execute "INSERT INTO CLRCZZLS(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库型,单号) select CKBL.材料名称,CKBL.材料规格,CKBL.材料单位,CKBL.颜色,批次,CKBL.数量,CKBL.单价,CKBL.库类,库别,备注 from ckBL WHERE CKBL.单号='" & DBCombo2.Text & "'"
Data8.Database.Execute "UPDATE CLRCZZLS SET 库别='入库' WHERE 库别=NULL"
Data8.Database.Execute "INSERT INTO CLRCZZHZLS(单号,库型,库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价) SELECT 单号,库型,CLRCZZLS.库类,CLRCZZLS.材料名称,CLRCZZLS.材料规格,CLRCZZLS.材料单位,CLRCZZLS.颜色,批次,format(SUM(CLRCZZLS.数量),'#0.00') AS L,CLRCZZLS.单价 FROM CLRCZZLS GROUP BY 单号,库型,CLRCZZLS.库类,CLRCZZLS.材料名称,CLRCZZLS.材料规格,CLRCZZLS.材料单位,CLRCZZLS.颜色,批次,CLRCZZLS.单价"
Data3.RecordSource = "SELECT 材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库型,单号 FROM CLRCZZHZLS WHERE CLRCZZHZLS.数量>0 ORDER BY 库型,库类"
Data3.Refresh
End Sub

Private Sub MSF()
With MSFlexGrid3
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

Private Sub MSFlexGrid3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSF
End If
End Sub


Private Sub Option2_Click()
Data1.Database.Execute "DELETE * FROM JHCK"
Data3.Database.Execute "INSERT INTO JHCK(单号,材料名称,材料规格,材料批号,材料单位,材料颜色,计划量) IN 'd:\数据库\\htgl\2011\SCZYJHD.MDB' SELECT CKBL.单号,CKBL.材料名称,CKBL.材料规格,批次,CKBL.材料单位,CKBL.颜色,SUM(CKBL.数量) AS 计划量 FROM CKBL WHERE CKBL.单号='" & DBCombo2.Text & "'  GROUP BY CKBL.单号,CKBL.材料名称,CKBL.材料规格,批次,CKBL.材料单位,CKBL.颜色"
Data3.Database.Execute "INSERT INTO JHCK(单号,材料名称,材料规格,材料批号,材料单位,材料颜色,实领量) IN 'd:\数据库\\htgl\2011\SCZYJHD.MDB' SELECT KPD.单号,KPD.材料名称,KPD.材料规格,批次,KPD.材料单位,KPD.颜色,SUM(KPD.数量) AS 实领量 FROM KPD WHERE KPD.单号='" & DBCombo2.Text & "' AND 标签='库存料'  GROUP BY KPD.单号,KPD.材料名称,KPD.材料规格,批次,KPD.材料单位,KPD.颜色"
Data1.Database.Execute "UPDATE JHCK SET 计划量=0 WHERE 计划量=null"
Data1.Database.Execute "UPDATE JHCK SET 实领量=0 WHERE 实领量=null"
Data1.Database.Execute "UPDATE JHCK SET 材料颜色='' WHERE 材料颜色=null"
Data1.Database.Execute "UPDATE JHCK SET 材料批号='' WHERE 材料批号=null"
Data1.Database.Execute "UPDATE JHCK SET 材料规格='' WHERE 材料规格=null"
Data1.Database.Execute "UPDATE JHCK SET 类别='1' WHERE 类别=null"
Data1.Database.Execute "INSERT INTO JHCK(材料名称,材料规格,材料批号,材料单位,材料颜色,计划量,实领量,欠领量) SELECT JHCK.材料名称,JHCK.材料规格,材料批号,材料单位,JHCK.材料颜色,format(SUM(JHCK.计划量),'#0.00') AS 计划量,format(SUM(JHCK.实领量),'#0.00') AS 领料量,format(SUM(JHCK.计划量-JHCK.实领量),'#0.00') AS 欠领量 FROM JHCK  GROUP BY JHCK.材料名称,JHCK.材料规格,材料批号,材料单位,JHCK.材料颜色"
Data1.Database.Execute "DELETE * FROM JHCK WHERE 类别='1'"
Data1.RecordSource = "SELECT JHCK.材料名称,JHCK.材料规格,材料批号,材料单位,JHCK.材料颜色,计划量,实领量,欠领量 FROM JHCK WHERE 计划量>0  order by JHCK.材料名称"
Data1.Refresh

Data6.RecordSource = "SELECT * FROM KPD WHERE KPD.单据号='" & Text2.Text & "' "
Data6.Refresh
DBCombo1.Text = Option2.Caption
End Sub

Private Sub Option3_Click()
Data1.Database.Execute "DELETE * FROM JHCK"
'Data3.Database.Execute "INSERT INTO JHCK(单号,材料名称,材料规格,材料批号,材料单位,材料颜色,计划量) IN 'd:\数据库\\htgl\2011\SCZYJHD.MDB' SELECT CKGL.单号,CKGL.材料名称,CKGL.材料规格,CKGL.材料单位,CKGL.颜色,SUM(CKGL.数量) AS 计划量 FROM CKGL WHERE CKGL.单号='" & DBCombo4.Text & "' and 库别<>'清库库存' GROUP BY CKGL.单号,CKGL.材料名称,CKGL.材料规格,CKGL.材料单位,CKGL.颜色"
Data1.Database.Execute "INSERT INTO JHCK(材料名称,材料规格,材料批号,材料单位,材料颜色,计划量) SELECT CGCLB.材料名称,CGCLB.材料规格,材料批号,CGCLB.材料单位,CGCLB.材料颜色,SUM(CGCLB.材料数量) AS 计划量 FROM CGCLB WHERE CGCLB.单号='" & DBCombo2.Text & "'  GROUP BY CGCLB.材料名称,CGCLB.材料规格,材料批号,CGCLB.材料单位,CGCLB.材料颜色"
Data3.Database.Execute "INSERT INTO JHCK(材料名称,材料规格,材料批号,材料单位,材料颜色,实领量) IN 'd:\数据库\\htgl\2011\SCZYJHD.MDB' SELECT KPD.材料名称,KPD.材料规格,批次,KPD.材料单位,KPD.颜色,SUM(KPD.数量) AS 实领量 FROM KPD WHERE KPD.单号='" & DBCombo2.Text & "' AND 标签='采购料'  GROUP BY KPD.材料名称,KPD.材料规格,批次,KPD.材料单位,KPD.颜色"
Data1.Database.Execute "UPDATE JHCK SET 计划量=0 WHERE 计划量=null"
Data1.Database.Execute "UPDATE JHCK SET 实领量=0 WHERE 实领量=null"
Data1.Database.Execute "UPDATE JHCK SET 材料颜色='' WHERE 材料颜色=null"
Data1.Database.Execute "UPDATE JHCK SET 材料批号='' WHERE 材料批号=null"
Data1.Database.Execute "UPDATE JHCK SET 材料规格='' WHERE 材料规格=null"
Data1.Database.Execute "UPDATE JHCK SET 类别='1' WHERE 类别=null"
Data1.Database.Execute "INSERT INTO JHCK(材料名称,材料规格,材料批号,材料单位,材料颜色,计划量,实领量,欠领量) SELECT JHCK.材料名称,JHCK.材料规格,材料批号,材料单位,JHCK.材料颜色,format(SUM(JHCK.计划量),'#0.00') AS 计划量,format(SUM(JHCK.实领量),'#0.00') AS 领料量,format(SUM(JHCK.计划量-JHCK.实领量),'#0.00') AS 欠领量 FROM JHCK  GROUP BY JHCK.材料名称,JHCK.材料规格,材料批号,材料单位,JHCK.材料颜色"
Data1.Database.Execute "DELETE * FROM JHCK WHERE 类别='1'"
Data1.RecordSource = "SELECT JHCK.材料名称,JHCK.材料规格,材料批号,材料单位,JHCK.材料颜色,计划量,实领量,欠领量 FROM JHCK WHERE 计划量>0  order by JHCK.材料名称"
Data1.Refresh

Data6.RecordSource = "SELECT * FROM KPD WHERE KPD.单据号='" & Text2.Text & "' "
Data6.Refresh
DBCombo1.Text = Option3.Caption
End Sub

Private Sub Option5_Click()
On Error Resume Next

Data2.RecordSource = "SELECT * FROM CGCLB WHERE 单号='" & DBCombo2.Text & "' "
Data2.Refresh

'If Not Data2.Recordset.EOF Then
'MsgBox ("采购表生成时没有统一批次，所以不能继续，请统一批次后再继续！")
'Exit Sub
'End If

Data8.Database.Execute "DELETE * FROM KCCXLS"
Data8.Database.Execute "DELETE * FROM KCCXHZLS"
Data10.Database.Execute "INSERT INTO KCCXLS(材料名称,材料规格,材料单位,颜色,批次,数量,库类) IN'd:\数据库\\htgl\2011\CKGL.MDB' select CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,材料批号,材料数量,材料库类 from CGCLB WHERE CGCLB.单号='" & DBCombo2.Text & "'"
Data8.Database.Execute "UPDATE KCCXLS SET 库别='入库' where 库别=NULL"
Data8.Database.Execute "INSERT INTO KCCXLS(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类) select KPD.材料名称,KPD.材料规格,KPD.材料单位,KPD.颜色,KPD.批次,KPD.数量,KPD.单价,KPD.库类 from KPD WHERE KPD.单号='" & DBCombo2.Text & "'  AND 标签='采购料'"
'Data8.Database.Execute "INSERT INTO KCCXLS(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类) select KPD.材料名称,KPD.材料规格,KPD.材料单位,KPD.颜色,KPD.批次,KPD.数量,KPD.单价,KPD.库类 from KPD WHERE KPD.单号='" & DBCombo4.Text & "'  AND AND 标签='采购料' AND KPD.日期 BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') "
Data8.Database.Execute "UPDATE KCCXLS SET 库别='出库',数量=-数量 WHERE 库别=NULL"
Data8.Database.Execute "INSERT INTO KCCXHZLS(库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价) SELECT KCCXLS.库类,KCCXLS.材料名称,KCCXLS.材料规格,KCCXLS.材料单位,KCCXLS.颜色,KCCXLS.批次,format(SUM(KCCXLS.数量),'#0.00') AS L,KCCXLS.单价 FROM KCCXLS GROUP BY KCCXLS.库类,KCCXLS.材料名称,KCCXLS.材料规格,KCCXLS.材料单位,KCCXLS.颜色,KCCXLS.批次,KCCXLS.单价"
Data8.Database.Execute "UPDATE KCCXHZLS SET 库型='采购入库' "
Data3.RecordSource = "SELECT 材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类 FROM KCCXHZLS WHERE KCCXHZLS.数量>0 ORDER BY 库类"
Data3.Refresh
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
Data1.Recordset.FindFirst "材料名称='" & Data3.Recordset.Fields(0) & "'  AND 材料颜色='" & Data3.Recordset.Fields(3) & "' and 材料规格='" & Data3.Recordset.Fields(1) & "'"
If Data1.Recordset.NoMatch Then
Data3.Recordset.Edit
Data3.Recordset.Fields(5) = 0
Data3.Recordset.Update
Else
Data3.Recordset.Edit
Data3.Recordset.Fields(5) = Data1.Recordset.Fields(7)
Data3.Recordset.Update
End If
Data3.Recordset.MoveNext
Loop
Data3.RecordSource = "SELECT 材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库型 FROM KCCXHZLS WHERE KCCXHZLS.数量>0 ORDER BY 库类"
Data3.Refresh
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 5
       Text1(7).Text = Format(Val(Text1(5).Text) * Val(Text1(6).Text), "#0.00")
End Select
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid3.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid3.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid3.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data6.Recordset.MoveFirst
Data6.Recordset.Move r - 1
Data6.Recordset.Edit

Data6.Recordset.Fields(c - 1) = Text1111.Text
Data6.Recordset.Update

Text1111.Visible = False
MSFlexGrid3.SetFocus
End Sub


Private Sub Text2_Change()
Data6.RecordSource = "SELECT * FROM KPD WHERE KPD.单据号='" & Text2.Text & "' "
Data6.Refresh
End Sub

Private Sub Timer1_Timer()
If BAR = 100 Then
DataEnvironment1.Command3 Text2.Text
DataReport9.Show 1
DataEnvironment1.rsCommand3.Close
Timer1.Enabled = False
ProgressBar1.Visible = False
Text2.Text = "00000001"
If Data8.Recordset.EOF Then
Text2.Text = "00000001"
Else
Text2.Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

Exit Sub
End If
BAR = BAR + 1
ProgressBar1.Value = BAR
End Sub

Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex

   TreeView1.Nodes.Clear
 
    Data13.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
    Data13.Refresh
    m = 1
    If Not Data13.Recordset.EOF Then  'make sure there are records in the table
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data13.Recordset.Fields(0)
        intIndex = mNode.Index
        Data11.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data13.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
        Data11.Refresh
        
        If Not Data11.Recordset.EOF Then
        Data11.Recordset.MoveFirst
        Do While Not Data11.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data11.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data12.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data11.Recordset.Fields(0) & "' and 进度='进行'"
        Data12.Refresh
        
        If Not Data12.Recordset.EOF Then
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data12.Recordset.Fields(0))
        Data12.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data11.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data13.Recordset.MoveNext
        Loop
    End If

End Sub


'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") > 0 Then
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
DBCombo2.Text = l1
End If

End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub





