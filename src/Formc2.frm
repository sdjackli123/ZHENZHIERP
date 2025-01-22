VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formc2 
   BackColor       =   &H00C0E0FF&
   Caption         =   "仓库管理之材料入库"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "Formc2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   390
      Left            =   1320
      TabIndex        =   75
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc2.frx":440A
      Height          =   3735
      Left            =   3720
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   32896
      FocusRect       =   2
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
   Begin VB.CommandButton Command3 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8280
      Top             =   8640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "客户信息"
      Height          =   3135
      Index           =   1
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   10935
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":441E
         Height          =   330
         Index           =   6
         Left            =   3360
         TabIndex        =   1
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "YS"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   8
         Left            =   5880
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":4433
         Height          =   330
         Index           =   10
         Left            =   10560
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "XM"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   11
         Left            =   11760
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":4447
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
         Left            =   3360
         TabIndex        =   2
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "仓位"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   12
         Left            =   11640
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   13
         Left            =   11640
         TabIndex        =   0
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":445B
         Height          =   330
         Index           =   14
         Left            =   5880
         TabIndex        =   35
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "MC"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":446F
         Height          =   330
         Index           =   15
         Left            =   840
         TabIndex        =   38
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "MC"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":4483
         Height          =   330
         Index           =   16
         Left            =   840
         TabIndex        =   39
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "MC"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":4498
         Height          =   330
         Index           =   17
         Left            =   11640
         TabIndex        =   42
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "库类"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":44AC
         Height          =   330
         Index           =   18
         Left            =   5880
         TabIndex        =   43
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "XM"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   19
         Left            =   5880
         TabIndex        =   46
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   20
         Left            =   5880
         TabIndex        =   48
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   21
         Left            =   8400
         TabIndex        =   50
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":44C0
         Height          =   330
         Index           =   22
         Left            =   840
         TabIndex        =   52
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
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
         TabIndex        =   60
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":44D4
         Height          =   330
         Index           =   3
         Left            =   3360
         TabIndex        =   61
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "材料名称"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":44E8
         Height          =   330
         Index           =   4
         Left            =   3360
         TabIndex        =   62
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "材料规格"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":44FD
         Height          =   330
         Index           =   5
         Left            =   3360
         TabIndex        =   63
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "MC"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         DataSource      =   "Data3"
         Height          =   330
         Index           =   9
         Left            =   10560
         TabIndex        =   64
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         BoundColumn     =   "XM"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   2
         Left            =   840
         TabIndex        =   65
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formc2.frx":4511
         Height          =   330
         Index           =   0
         Left            =   10920
         TabIndex        =   66
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "简称"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   390
         Index           =   23
         Left            =   7800
         TabIndex        =   67
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单据号"
         Height          =   375
         Index           =   17
         Left            =   7800
         TabIndex        =   68
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "供应商"
         Height          =   375
         Index           =   16
         Left            =   240
         TabIndex        =   53
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "序号"
         Height          =   375
         Index           =   15
         Left            =   7800
         TabIndex        =   51
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "日期"
         Height          =   375
         Index           =   14
         Left            =   5280
         TabIndex        =   49
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "库类"
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   47
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "仓务员"
         Height          =   375
         Index           =   12
         Left            =   5280
         TabIndex        =   45
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "备注"
         Height          =   375
         Index           =   11
         Left            =   5280
         TabIndex        =   44
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "数量"
         Height          =   375
         Index           =   2
         Left            =   5280
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单价"
         Height          =   375
         Index           =   1
         Left            =   10560
         TabIndex        =   40
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "批号"
         Height          =   375
         Index           =   10
         Left            =   2760
         TabIndex        =   36
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单号"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "实领量"
         Height          =   375
         Index           =   6
         Left            =   10560
         TabIndex        =   26
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "库别"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "是否付款"
         Height          =   375
         Index           =   1
         Left            =   10560
         TabIndex        =   23
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "金额"
         Height          =   375
         Index           =   0
         Left            =   10560
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "颜色"
         Height          =   375
         Index           =   9
         Left            =   2760
         TabIndex        =   21
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "规格"
         Height          =   375
         Index           =   6
         Left            =   2760
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "是否开票"
         Height          =   375
         Left            =   10560
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "仓位"
         Height          =   375
         Index           =   7
         Left            =   5280
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "款号"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单位"
         Height          =   375
         Index           =   4
         Left            =   2760
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "名称"
         Height          =   375
         Index           =   3
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "是否含税"
         Height          =   375
         Index           =   2
         Left            =   10560
         TabIndex        =   8
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   5640
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   7200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Formc2.frx":4525
      Left            =   12840
      List            =   "Formc2.frx":452F
      TabIndex        =   57
      Text            =   "Combo1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data17 
      Caption         =   "Data17"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data16 
      Caption         =   "Data16"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data15 
      Caption         =   "Data15"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
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
      Top             =   8640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新单据"
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
      TabIndex        =   55
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
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
      Top             =   10320
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
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
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
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
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
      Left            =   12600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   12120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   8640
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9000
      TabIndex        =   30
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   1320
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Height          =   495
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   12360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
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
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
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
      Left            =   12960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   12600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   3375
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   855
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   975
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4560
      TabIndex        =   31
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9000
      TabIndex        =   32
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39177
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3720
      TabIndex        =   59
      Top             =   5400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7695
      Left            =   240
      TabIndex        =   70
      Top             =   2760
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   13573
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1320
      TabIndex        =   71
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1320
      TabIndex        =   72
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39177
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
      Left            =   240
      TabIndex        =   74
      Top             =   1320
      Width           =   1095
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
      Left            =   240
      TabIndex        =   73
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "确认类别"
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
      Left            =   11880
      TabIndex        =   56
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "备料单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   54
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   34
      Top             =   1320
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   6480
      X2              =   7680
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   33
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "制 衣 材 料 入 库 "
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
      Left            =   6600
      TabIndex        =   6
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Formc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X, BAR As Integer
Dim ba As Database: Dim rr As Integer
Dim rs As Single: Dim rd As Recordset: Dim ba1 As Database: Public ll As Integer
Dim rd1 As Recordset
Dim A As String  '中间变量
Dim B As Double
Dim c, r As Integer
Dim kg As Integer
Dim bb As Long
Dim cc As String
Dim kkf As Integer
Dim N As Integer
Dim DH As Integer
Dim fh As String

Private Sub Command12_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
Data13.Refresh
DBCombo1(23).Text = "00000001"
If Data13.Recordset.EOF Then
DBCombo1(23).Text = "00000001"
Else
DBCombo1(23).Text = Left("00000000", 8 - Len(Trim(Str(Data13.Recordset.Fields(0) + 1)))) + Trim(Str(Data13.Recordset.Fields(0) + 1))
End If

End Sub


Private Sub Command3_Click()
Call tree
Call zk
End Sub

Private Sub Command33_Click()
Unload Me
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

If Combo1.Text = "" Or Combo1.Text = "未" Then

If DBCombo1(0).Text = "" Then
Data5.RecordSource = "select   * from ckgl WHERE CKGL.确认='未' AND CKGL.日期 between CDate(' " & Text4.Text & "' )  and   CDate(' " & Text5.Text & " ') order by CKGL.日期,val(ckgl.序号) desc "
Data5.Refresh
Else
Data5.RecordSource = "select   * from ckgl WHERE CKGL.确认='未' AND CKGL.客户='" & DBCombo1(0).Text & "' AND  CKGL.日期 between CDate(' " & Text4.Text & "' )  and   CDate(' " & Text5.Text & " ') order by CKGL.日期,val(ckgl.序号) desc "
Data5.Refresh
End If

Else

If DBCombo1(0).Text = "" Then
Data5.RecordSource = "select   * from ckgl WHERE CKGL.确认='已' AND CKGL.日期 between CDate(' " & Text4.Text & "' )  and   CDate(' " & Text5.Text & " ') order by CKGL.日期,val(ckgl.序号) desc "
Data5.Refresh
Else
Data5.RecordSource = "select   * from ckgl WHERE CKGL.确认='已' AND CKGL.客户='" & DBCombo1(0).Text & "' AND  CKGL.日期 between CDate(' " & Text4.Text & "' )  and   CDate(' " & Text5.Text & " ') order by CKGL.日期,val(ckgl.序号) desc "
Data5.Refresh
End If

End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
If DBCombo1(18).Text = "" Then
MsgBox ("请选择负责人！")
Exit Sub
End If

If DBCombo1(8).Text = "" Or DBCombo1(9).Text = "" Then
MsgBox ("数量、单价不能为空！！")
Exit Sub
End If

If DBCombo1(22).Text = "" Then
MsgBox ("请选择供应商！")
Exit Sub
End If


rd.AddNew
For i = 0 To rd.Fields.Count - 1
rd.Fields(i) = DBCombo1(i).Text
Next
rd.Fields(24) = "未"
rd.Fields(25) = "未"
rd.Update

For i = 3 To rd.Fields.Count - 7
If i = 18 Then
CWY = DBCombo1(i).Text
End If
DBCombo1(i).Text = ""
Next

Data5.RecordSource = "select   * from ckgl WHERE CKGL.单据号='" & DBCombo1(23).Text & "' order by Val(ckgl.序号)"
Data5.Refresh

If Data5.Recordset.RecordCount = 6 Then
If MsgBox("是否打印本单据？", vbYesNo) = vbNo Then
DBCombo1(23).Text = "00000001"
If Data13.Recordset.EOF Then
DBCombo1(23).Text = "00000001"
Else
DBCombo1(23).Text = Left("00000000", 8 - Len(Trim(Str(Data13.Recordset.Fields(0) + 1)))) + Trim(Str(Data13.Recordset.Fields(0) + 1))
End If
Else       '''''''''''''''''''''''''''
Call Command6_Click
DBCombo1(23).Text = "00000001"
If Data13.Recordset.EOF Then
DBCombo1(23).Text = "00000001"
Else
DBCombo1(23).Text = Left("00000000", 8 - Len(Trim(Str(Data13.Recordset.Fields(0) + 1)))) + Trim(Str(Data13.Recordset.Fields(0) + 1))
End If
End If
End If

Data5.RecordSource = "select   * from ckgl WHERE CKGL.单据号='" & DBCombo1(23).Text & "' order by Val(ckgl.序号)"
Data5.Refresh

Data7.RecordSource = "select MAX(VAL(CKGL.序号)) from ckgl "
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

DBCombo1(11).Text = ""
DBCombo1(16).Text = "采购入库"
DBCombo1(17).Text = 0
DBCombo1(18).Text = CWY
DBCombo1(20).Text = Date
DBCombo1(21).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(0).SetFocus
End Sub

Private Sub Command2_Click()
On Error Resume Next

If DBCombo1(18).Text = "" Then
MsgBox ("请选择负责人！")
Exit Sub
End If

If DBCombo1(22).Text = "" Then
MsgBox ("请选择供应商！")
Exit Sub
End If


If DBCombo1(8).Text = "" Or DBCombo1(9).Text = "" Then
MsgBox ("数量、单价不能为空！！")
Exit Sub
End If
If Data5.Recordset.Fields(24) = "已" Then Exit Sub
Data5.Recordset.Edit
For i = 0 To Data5.Recordset.Fields.Count - 1
Data5.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data5.Recordset.Update


For i = 3 To rd.Fields.Count - 7
If i = 18 Then
CWY = DBCombo1(i).Text
End If
DBCombo1(i).Text = ""
Next

Data5.RecordSource = "select   * from ckgl WHERE CKGL.单据号='" & DBCombo1(23).Text & "' order by Val(ckgl.序号)"
Data5.Refresh


Data7.RecordSource = "select MAX(VAL(CKGL.序号)) from ckgl "
Data7.Refresh

DBCombo1(11).Text = ""
DBCombo1(16).Text = "采购入库"
DBCombo1(17).Text = 0
DBCombo1(18).Text = CWY
DBCombo1(20).Text = Date
DBCombo1(21).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(0).SetFocus
End Sub

Private Sub Command4_Click()
On Error Resume Next

If Data5.Recordset.Fields(24) = "已" Then Exit Sub
Data5.Recordset.Delete

For i = 3 To rd.Fields.Count - 7
If i = 18 Then
CWY = DBCombo1(i).Text
End If
DBCombo1(i).Text = ""
Next


Data5.RecordSource = "select   * from ckgl WHERE CKGL.单据号='" & DBCombo1(23).Text & "' order by Val(ckgl.序号)"
Data5.Refresh
Data7.RecordSource = "select MAX(VAL(CKGL.序号)) from ckgl "
Data7.Refresh

DBCombo1(11).Text = ""
DBCombo1(16).Text = "采购入库"
DBCombo1(17).Text = 0
DBCombo1(18).Text = CWY
DBCombo1(20).Text = Date
DBCombo1(21).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(0).SetFocus
End Sub


Private Sub Command8_Click()
On Error Resume Next
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
DBCombo1(20).Text = Date
DBCombo1(8).Text = 0
Data1.Refresh
Data3.Refresh
Data4.Refresh
Data6.Refresh
Data8.Refresh
Data9.Refresh
Data14.Refresh
Data7.Database.Execute "UPDATe CKGL SET 序号='0'  WHERE 序号=null"
Data7.Database.Execute "UPDATe CKGL SET 单据号='00000000'  WHERE 单据号=null or 单据号=''"
Data7.RecordSource = "select   MAX(VAL(CKGL.序号)) from ckgl "
Data7.Refresh
DBCombo1(21).Text = 1
DBCombo1(21).Text = Data7.Recordset.Fields(0) + 1
Data5.RecordSource = "select   * from ckgl WHERE CKGL.单据号='" & DBCombo1(23).Text & "' order by Val(ckgl.序号)"
Data5.Refresh

End Sub

Private Sub Command6_Click()
If Data5.Recordset.EOF Then
MsgBox ("此单据号中无记录，不能打印！")
Exit Sub
End If

BAR = 10
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub


Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
Case 2


'Data2.RecordSource = "select CKGL.材料名称 from ckgl WHERE CKGL.单号='" & DBCombo1(1).Text & "'  group by ckgl.材料名称"
'Data2.Refresh


Case 3

Case 8
DBCombo1(10).Text = Format(Val(DBCombo1(8).Text) * Val(DBCombo1(9).Text), "#0.00")

Case 9
DBCombo1(10).Text = Format(Val(DBCombo1(8).Text) * Val(DBCombo1(9).Text), "#0.00")


Case 11



Case 15
'If DBCombo1(15).Text = "1主料库" Then
'DBCombo1(7).Text = Format(Date, "YYMMDD")
'Else
'DBCombo1(7).Text = ""
'End If




Case 23
Data5.RecordSource = "select   * from ckgl WHERE CKGL.单据号='" & DBCombo1(23).Text & "'  order by Val(ckgl.序号)"
Data5.Refresh
End Select

End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
Case 2


'Data2.RecordSource = "select CKGL.材料名称 from ckgl WHERE CKGL.单号='" & DBCombo1(1).Text & "' AND CKGL.库类='" & DBCombo1(15).Text & "' group by ckgl.材料名称"
'Data2.Refresh


Case 3
Data16.RecordSource = "select CKGL.材料规格 from ckgl WHERE CKGL.库类='" & DBCombo1(15).Text & "' AND CKGL.材料名称='" & DBCombo1(3).Text & "' group by ckgl.材料规格"
Data16.Refresh

Case 8
DBCombo1(10).Text = Format(Val(DBCombo1(8).Text) * Val(DBCombo1(9).Text), "#0.00")

Case 9
DBCombo1(10).Text = Format(Val(DBCombo1(8).Text) * Val(DBCombo1(9).Text), "#0.00")

Case 15
If DBCombo1(15).Text = "3零件库" Then
For i = 3 To 7
DBCombo1(i).Enabled = True
Next
Else
For i = 3 To 7
DBCombo1(i).Enabled = False
Next
End If
Data2.RecordSource = "select CKGL.材料名称 from ckgl WHERE CKGL.库类='" & DBCombo1(15).Text & "'  group by ckgl.材料名称"
Data2.Refresh

Case 22

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

Private Sub Form_Load()
On Error Resume Next

Combo1.Text = ""

Set ba = OpenDatabase("d:\数据库\\htgl\2011\ckgl.MDB")
Set rd = ba.OpenRecordset("ckgl", dbOpenDynaset)

For i = 0 To rd.Fields.Count - 1
DBCombo1(i).Text = ""
Next
DBCombo1(9).Text = 0
DBCombo1(17).Text = 0
Text4.Text = Date
Text5.Text = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
DBCombo1(20).Text = Date
DTPicker3.Value = Date - 30
DTPicker4.Value = Date
DBCombo2.Text = ""
DBCombo1(23).Enabled = False

Data1.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data1.RecordSource = "select 简称 from KHZL group by 简称"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\CKGL.mdb"
Data2.RecordSource = "select CKGL.材料名称 from ckgl   group by ckgl.材料名称"
Data2.Refresh


Data3.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data3.RecordSource = "select CW.MC from CW group by CW.MC"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data4.RecordSource = "select fzr.xm  from fzr group by fzr.xm"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data5.RecordSource = "select   * from ckgl WHERE CKGL.单据号='" & DBCombo1(23).Text & "' order by Val(ckgl.序号)"
Data5.Refresh

Data6.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data6.RecordSource = "select KL.MC from KL   group by KL.MC"
Data6.Refresh

Data7.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data7.RecordSource = "select   MAX(VAL(CKGL.序号)) from ckgl "
Data7.Refresh

Data8.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data8.RecordSource = "select 简称 from GYS group by 简称"
Data8.Refresh


Data13.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
Data13.RecordSource = "SELECT MAX(VAL(ckgl.单据号)) FROM CKGL"
Data13.Refresh

Data14.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data14.RecordSource = "select KB.MC from KB group by KB.MC"
Data14.Refresh

Data15.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data15.RecordSource = "select YS.YS from YS group by YS.YS"
Data15.Refresh

Data16.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data17.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
ProgressBar1.Visible = False
Timer1.Enabled = False

DBCombo1(23).Text = "00000001"
If Data13.Recordset.EOF Then
DBCombo1(23).Text = "00000001"
Else
DBCombo1(23).Text = Left("00000000", 8 - Len(Trim(Str(Data13.Recordset.Fields(0) + 1)))) + Trim(Str(Data13.Recordset.Fields(0) + 1))
End If




Data9.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data9.RecordSource = "select CLDW.MC from CLDW group by CLDW.MC"
Data9.Refresh

Data10.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data11.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data12.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 0
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 0
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1000
MSFlexGrid1.ColWidth(8) = 1200
MSFlexGrid1.ColWidth(9) = 1200
MSFlexGrid1.ColWidth(10) = 0
MSFlexGrid1.ColWidth(11) = 0
MSFlexGrid1.ColWidth(12) = 0
MSFlexGrid1.ColWidth(13) = 0
MSFlexGrid1.ColWidth(14) = 0
MSFlexGrid1.ColWidth(25) = 0
MSFlexGrid1.ColWidth(26) = 0
MSFlexGrid1.ColWidth(27) = 0
MSFlexGrid1.ColWidth(28) = 0
MSFlexGrid1.ColWidth(29) = 0



DBCombo1(16).Text = "采购入库"
DBCombo1(17).Text = 0
DBCombo1(21).Text = 1
DBCombo1(21).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(20).Text = Date
DBCombo1(0).TabIndex = 0

Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

End Sub

Private Sub Label3_Click(Index As Integer)
Select Case Index
Case 5
       khbl = 2
Formc202.Text1.Text = DBCombo1(2).Text
Formc202.Show

       Case 17
DBCombo1(23).Enabled = False
End Select
End Sub

Private Sub Label3_dblClick(Index As Integer)
Select Case Index
       Case 17
DBCombo1(23).Enabled = True
End Select
End Sub

Private Sub Label5_Click()
If DBCombo2.Text = "" Then Exit Sub
Formc26.DBCombo1.Text = DBCombo2.Text
Formc26.Show
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
Data5.Recordset.MoveFirst
Data5.Recordset.Move rs - 1
For i = 0 To Data5.Recordset.Fields.Count - 4
DBCombo1(i).Text = Data5.Recordset.Fields(i)
Next
Command11.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub









Private Sub Timer1_Timer()
If BAR = 100 Then
Call clrk(Data17, DBCombo1(23).Text)
Timer1.Enabled = False
ProgressBar1.Visible = False
Exit Sub
End If
BAR = BAR + 10
ProgressBar1.Value = BAR

End Sub


Private Sub MSFlex_DBLClick()
With MSFlexGrid1
    c = .Col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex_DBLClick
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid1.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid1.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid1.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data5.Recordset.MoveFirst
Data5.Recordset.Move r - 1
Data5.Recordset.Edit
Data5.Recordset.Fields(c - 1) = Text1111.Text
Data5.Recordset.Update
Text1111.Visible = False
End Sub



Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex

   TreeView1.Nodes.Clear
 
    Data10.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
    Data10.Refresh
    m = 1
    If Not Data10.Recordset.EOF Then  'make sure there are records in the table
        Data10.Recordset.MoveFirst
        Do While Not Data10.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data10.Recordset.Fields(0)
        intIndex = mNode.Index
        Data11.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data10.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
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
        Data10.Recordset.MoveNext
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
DBCombo2.Text = l1
End If

End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub




