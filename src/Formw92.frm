VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formw92 
   BackColor       =   &H00C0E0FF&
   Caption         =   "成品出库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form14"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   12720
      TabIndex        =   64
      Text            =   "Text2"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   12720
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Data Data15 
      Caption         =   "Data15"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
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
      Top             =   9360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
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
      Width           =   3375
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "日期范围"
      Height          =   855
      Left            =   9000
      TabIndex        =   42
      Top             =   5160
      Width           =   5895
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0FF&
         Caption         =   "打印"
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   45
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   3360
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0FF&
         Caption         =   "查询"
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   120
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   840
         TabIndex        =   46
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarTitleBackColor=   8421376
         CalendarTrailingForeColor=   255
         Format          =   81068033
         CurrentDate     =   39177
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   3360
         TabIndex        =   47
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarTitleBackColor=   8421376
         CalendarTrailingForeColor=   255
         Format          =   81068033
         CurrentDate     =   39177
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000C0C0&
         Caption         =   "起始日期"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000C0C0&
         Caption         =   "结束日期"
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   48
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
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
      Top             =   -480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号查询"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库存查询"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5160
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4080
      TabIndex        =   39
      Top             =   6360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   12000
      Top             =   360
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新单据"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5160
      Width           =   975
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
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
      Top             =   -360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
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
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   -360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   -240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   -360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Top             =   -360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   81068033
      CurrentDate     =   39910
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw92.frx":0000
      Height          =   330
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   3720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   2
      Left            =   6840
      TabIndex        =   8
      Top             =   3720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   3
      Left            =   9240
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   4
      Left            =   10920
      TabIndex        =   10
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw92.frx":0014
      Height          =   1335
      Left            =   600
      TabIndex        =   11
      Top             =   1680
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   14
      BackColorFixed  =   12171775
      BackColorBkg    =   45232
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      AllowUserResizing=   3
      FormatString    =   "记录号"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formw92.frx":0028
      Height          =   3255
      Left            =   480
      TabIndex        =   12
      Top             =   6240
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   14
      BackColorFixed  =   12171775
      BackColorBkg    =   45232
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      AllowUserResizing=   3
      FormatString    =   "记录号"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   5
      Left            =   11520
      TabIndex        =   13
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw92.frx":003C
      Height          =   330
      Index           =   6
      Left            =   480
      TabIndex        =   14
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   7
      Left            =   2280
      TabIndex        =   15
      Top             =   4560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   8
      Left            =   4440
      TabIndex        =   16
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw92.frx":0050
      Height          =   330
      Index           =   9
      Left            =   6840
      TabIndex        =   17
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   10
      Left            =   8280
      TabIndex        =   18
      Top             =   4560
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
      Index           =   11
      Left            =   9840
      TabIndex        =   19
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw92.frx":0064
      Height          =   330
      Index           =   12
      Left            =   1320
      TabIndex        =   20
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "xm"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   13
      Left            =   1320
      TabIndex        =   36
      Top             =   5640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Bindings        =   "Formw92.frx":0078
      Height          =   1335
      Left            =   10200
      TabIndex        =   51
      Top             =   1680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   14
      BackColorFixed  =   12171775
      BackColorBkg    =   45232
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      AllowUserResizing=   3
      FormatString    =   "记录号"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   3840
      TabIndex        =   53
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   375
      Left            =   3840
      TabIndex        =   54
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker6 
      Height          =   375
      Left            =   360
      TabIndex        =   58
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   39883
   End
   Begin VB.Label Label1 
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
      Index           =   16
      Left            =   12720
      TabIndex        =   62
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "条码"
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
      Left            =   12720
      TabIndex        =   61
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "型号"
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
      Left            =   10920
      TabIndex        =   60
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
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
      Left            =   360
      TabIndex        =   59
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   57
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   56
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择入库日期范围"
      ForeColor       =   &H000000C0&
      Height          =   855
      Index           =   1
      Left            =   2280
      TabIndex        =   55
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   17
      Left            =   9960
      TabIndex        =   52
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "单据号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   38
      Top             =   5640
      Width           =   855
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
      Index           =   7
      Left            =   11520
      TabIndex        =   35
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
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
      Left            =   4440
      TabIndex        =   34
      Top             =   3360
      Width           =   2175
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
      Index           =   4
      Left            =   6840
      TabIndex        =   33
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "购货单位"
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
      Left            =   2280
      TabIndex        =   32
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "   成   品   出   库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   6600
      TabIndex        =   31
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格"
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
      Left            =   9240
      TabIndex        =   30
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "成品库存信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   6
      Left            =   360
      TabIndex        =   29
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入日期"
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
      Left            =   480
      TabIndex        =   28
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单价"
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
      Left            =   2280
      TabIndex        =   27
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "数量"
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
      Left            =   480
      TabIndex        =   26
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "金额"
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
      Left            =   4440
      TabIndex        =   25
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "箱号"
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
      Left            =   6840
      TabIndex        =   24
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
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
      Left            =   8280
      TabIndex        =   23
      Top             =   4200
      Width           =   1335
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
      Index           =   14
      Left            =   9840
      TabIndex        =   22
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "仓务员"
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
      Left            =   480
      TabIndex        =   21
      Top             =   5160
      Width           =   855
   End
End
Attribute VB_Name = "Formw92"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BAR As Integer
Private Sub Command1_Click()
On Error Resume Next
If DBCombo1(2).Text = "" Then
MsgBox ("输入款号")
Exit Sub
End If

Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = DTPicker1.Value
For i = 0 To 13
Data1.Recordset.Fields(i + 1) = DBCombo1(i).Text
Next
Data1.Recordset.Fields(15) = "未"
Data1.Recordset.Fields(16) = Text1.Text
Data1.Recordset.Fields(17) = Text2.Text

Data1.Recordset.Update
Data1.Refresh
Data2.Refresh
For i = 2 To 12
If i = 9 Then i = i + 1
DBCombo1(i).Text = ""
Next
Text1.Text = ""
Text2.Text = ""
DBCombo1(11).Text = Data1.Recordset.RecordCount + 1
DBCombo1(7).Text = 0
End Sub

Private Sub Command10_Click()
Data1.RecordSource = "SELECT 日期,购货单位,款号,品名,规格,型号,单位,数量,单价,金额,发货地,备注,单据号 FROM LSFH WHERE 日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "') ORDER BY 日期"
Data1.Refresh
Call OutDataToExcel2(MSFlexGrid2, 6, 8, "成品出库")
Data1.RecordSource = "SELECT * from LSFH WHERE 日期 between CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "') ORDER BY 日期"
Data1.Refresh
End Sub




Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
If DBCombo1(2).Text = "" Then
MsgBox ("输入款号")
Exit Sub
End If

Data1.Recordset.Edit
Data1.Recordset.Fields(0) = DTPicker1.Value
For i = 0 To 13
Data1.Recordset.Fields(i + 1) = DBCombo1(i).Text
Next
Data1.Recordset.Fields(16) = Text1.Text
Data1.Recordset.Fields(17) = Text2.Text
Data1.Recordset.Update
Data1.Refresh
Data2.Refresh
For i = 2 To 12
If i = 9 Then i = i + 1
DBCombo1(i).Text = ""
Next
Text1.Text = ""
Text2.Text = ""
DBCombo1(11).Text = Data1.Recordset.RecordCount + 1
DBCombo1(7).Text = 0
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True

End Sub

Private Sub Command4_Click()
If MsgBox("确定删除吗？，删除将不能恢复！", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
Data2.Refresh
For i = 2 To 12
If i = 9 Then i = i + 1
DBCombo1(i).Text = ""
Next
Text1.Text = ""
Text2.Text = ""
DBCombo1(11).Text = Data1.Recordset.RecordCount + 1
DBCombo1(7).Text = 0
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Command5_Click()
If Data1.Recordset.EOF Then
MsgBox ("此单据号中无记录，不能打印！")
Exit Sub
End If
BAR = 1
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command6_Click()
On Error Resume Next
DBCombo1(13).Text = "00000001"

Data1.Database.Execute "UPDATE LSFH SET 单据号='00000000' WHERE 单据号=null"

Data8.DatabaseName = "d:\数据库\\htgl\2011\CPCK"
Data8.RecordSource = "SELECT MAX(VAL(单据号)) FROM LSFH "
Data8.Refresh

If Data8.Recordset.EOF Then
DBCombo1(13).Text = "00000001"
Else
DBCombo1(13).Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

End Sub

Private Sub Command7_Click()
On Error Resume Next
       Data3.Database.Execute "DELETE * FROM LSKC"
       Data3.Database.Execute "INSERT INTO LSKC(款号,品名,规格,型号,单位,数量,单号,条码) SELECT 款号,品名,规格,型号,单位,-数量,单号,条码 FROM LSFH where 日期 BETWEEN CDATE('" & DTPicker4.Value & "') AND CDATE('" & DTPicker5.Value & "')"
       Data3.Database.Execute "INSERT INTO LSKC(款号,品名,规格,型号,单位,数量,单号,条码) SELECT 款号,品名,规格,型号,单位,数量,单号,条码 FROM LSJL where 日期=CDATE('" & DTPicker4.Value & "')"
       Data1.Database.Execute "INSERT INTO LSKC(款号,品名,规格,型号,单位,数量,单号,条码) SELECT 款号,品名,规格,型号,单位,数量,单号,条码 FROM LSRK WHERE 日期 BETWEEN CDATE('" & DTPicker4.Value & "') AND CDATE('" & DTPicker5.Value & "')"
       Data2.RecordSource = "SELECT 款号,品名,规格,型号,单位,SUM(VAL(数量)) AS 库存量,单号,条码 FROM LSKC GROUP BY 款号,品名,规格,型号,单位,单号,条码"
       Data2.Refresh
End Sub

Private Sub Command8_Click()
       Data3.Database.Execute "DELETE * FROM LSKC"
       Data3.Database.Execute "INSERT INTO LSKC(款号,品名,规格,型号,单位,数量,单号,条码) SELECT 款号,品名,规格,型号,单位,-数量,单号,条码 FROM LSFH WHERE 款号='" & DBCombo1(1).Text & "'"
       Data1.Database.Execute "INSERT INTO LSKC(款号,品名,规格,型号,单位,数量,单号,条码) SELECT 款号,品名,规格,型号,单位,数量,单号,条码 FROM LSRK WHERE 款号='" & DBCombo1(1).Text & "'"
       Data1.Database.Execute "INSERT INTO LSKC(款号,品名,规格,型号,单位,数量,单号,条码) SELECT 款号,品名,规格,型号,单位,数量,单号,条码 FROM LSJL WHERE 款号='" & DBCombo1(1).Text & "' AND 日期=cdate('" & DTPicker4.Value & "')"
       Data2.RecordSource = "SELECT 款号,品名,规格,型号,单位,SUM(VAL(数量)) AS 库存量,单号,条码 FROM LSKC GROUP BY 款号,品名,规格,型号,单位,单号,条码"
       Data2.Refresh

End Sub

Private Sub Command9_Click()
Data1.RecordSource = "SELECT * from LSFH WHERE 日期 between CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "') ORDER BY 日期"
Data1.Refresh
End Sub


Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 6
       DBCombo1(8).Text = Format(Val(DBCombo1(7).Text) * Val(DBCombo1(6).Text), "#0.00")
       Case 7
       DBCombo1(8).Text = Format(Val(DBCombo1(7).Text) * Val(DBCombo1(6).Text), "#0.00")
       Case 13
       Data1.RecordSource = "SELECT * from LSFH WHERE 单据号='" & DBCombo1(13).Text & "' ORDER BY VAL(序号) DESC"
       Data1.Refresh
       Data13.RecordSource = "select distinct 客户,款号  FROM zxd WHERE 编号='" & DBCombo1(13).Text & "' "
       Data13.Refresh
End Select
End Sub

Private Sub DTPicker1_Change()
Data1.RecordSource = "SELECT * from LSFH WHERE 日期=CDATE('" & DTPicker1.Value & "') ORDER BY VAL(序号) DESC"
Data1.Refresh
DBCombo1(11).Text = Data1.Recordset.RecordCount + 1
DBCombo1(7).Text = 0
End Sub

Private Sub DTPicker1_CloseUp()
Data1.RecordSource = "SELECT * from LSFH WHERE 日期=CDATE('" & DTPicker1.Value & "') ORDER BY VAL(序号) DESC"
Data1.Refresh
DBCombo1(11).Text = Data1.Recordset.RecordCount + 1
DBCombo1(7).Text = 0
End Sub

Private Sub DTPicker2_Click()
Text4.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text4.Text = DTPicker2.Value
End Sub
Private Sub DTPicker3_Click()
Text5.Text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
Text5.Text = DTPicker3.Value
End Sub


Private Sub DTPicker6_CloseUp()
Data14.DatabaseName = "d:\数据库\\htgl\2011\CJBB.mdb"
Data14.RecordSource = "select * from RQSD where cdate('" & DTPicker6.Value & "') between 起始日期 and 结束日期"
Data14.Refresh
If Data14.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data14.Recordset.Fields(0)
K2 = Data14.Recordset.Fields(1)
End If
DTPicker4.Value = K1
DTPicker5.Value = K2
End Sub

Private Sub DTPicker6_Click()
Data14.DatabaseName = "d:\数据库\\htgl\2011\CJBB.mdb"
Data14.RecordSource = "select * from RQSD where cdate('" & DTPicker6.Value & "') between 起始日期 and 结束日期"
Data14.Refresh
If Data14.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data14.Recordset.Fields(0)
K2 = Data14.Recordset.Fields(1)
End If
DTPicker4.Value = K1
DTPicker5.Value = K2

End Sub

Private Sub Form_Load()
On Error Resume Next
DTPicker1.Value = Date
For i = 0 To 13
DBCombo1(i).Text = ""
Next
Text4.Text = Date
Text5.Text = Date
Text1.Text = ""
Text2.Text = ""
DTPicker3.Value = Date
DTPicker2.Value = Date
DTPicker6.Value = Date
Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data1.RecordSource = "SELECT * from LSFH WHERE 单据号='" & DBCombo1(13).Text & "' ORDER BY VAL(序号) DESC"
Data1.Refresh

Data14.DatabaseName = "d:\数据库\\htgl\2011\CJBB.mdb"
Data14.RecordSource = "select * from RQSD where cdate('" & DTPicker6.Value & "') between 起始日期 and 结束日期"
Data14.Refresh
If Data14.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data14.Recordset.Fields(0)
K2 = Data14.Recordset.Fields(1)
End If
DTPicker4.Value = K1
DTPicker5.Value = K2

ProgressBar1.Visible = False
Timer1.Enabled = False

Data2.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"

Data3.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"

Data4.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data4.RecordSource = "select fzr.xm  from fzr group by fzr.xm"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data5.RecordSource = "select CLDW.MC  FROM CLDW GROUP BY CLDW.MC"
Data5.Refresh

Data6.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"

Data7.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data7.RecordSource = "select FHD.MC  FROM FHD GROUP BY FHD.MC"
Data7.Refresh

DBCombo1(13).Text = "00000001"
DBCombo1(13).Enabled = False
Data1.Database.Execute "UPDATE LSFH SET 单据号='00000000' WHERE 单据号=null"

Data8.DatabaseName = "d:\数据库\\htgl\2011\CPCK"
Data8.RecordSource = "SELECT MAX(VAL(单据号)) FROM LSFH "
Data8.Refresh

If Data8.Recordset.EOF Then
DBCombo1(13).Text = "00000001"
Else
DBCombo1(13).Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

Data9.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data9.RecordSource = "SELECT 简称 FROM KHZL GROUP BY 简称"
Data9.Refresh


Data11.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data12.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data13.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data15.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"


DBCombo1(7).Text = 0
DBCombo1(11).Text = Data1.Recordset.RecordCount + 1
DBCombo1(5).Text = "件"

Data1.Database.Execute "DELETE * FROM LSKC"

Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 1200
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1000
MSFlexGrid1.ColWidth(6) = 1000
MSFlexGrid1.ColWidth(7) = 800



MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(1) = 1200
MSFlexGrid2.ColWidth(2) = 1200
MSFlexGrid2.ColWidth(3) = 1500
MSFlexGrid2.ColWidth(4) = 1200
MSFlexGrid2.ColWidth(5) = 2000
MSFlexGrid2.ColWidth(6) = 600


MSFlexGrid3.ColWidth(0) = 200
MSFlexGrid3.ColWidth(1) = 1200
MSFlexGrid3.ColWidth(2) = 1500
MSFlexGrid3.ColWidth(3) = 1500
MSFlexGrid3.ColWidth(4) = 3000

MSFlexGrid4.ColWidth(0) = 200
MSFlexGrid4.ColWidth(1) = 2200
MSFlexGrid4.ColWidth(2) = 2200

MSFlexGrid5.ColWidth(0) = 200
MSFlexGrid5.ColWidth(1) = 2000
MSFlexGrid5.ColWidth(2) = 1500
MSFlexGrid5.ColWidth(3) = 1500

End Sub

Private Sub Label1_dblClick(Index As Integer)
Select Case Index
       Case 12
Data15.RecordSource = "SELECT * FROM LSFH WHERE 日期=cdate('" & Date & "')"
Data15.Refresh
If Not Data15.Recordset.EOF Then
Data15.RecordSource = "select max(mid(发货地,7)) from lsfh where 日期=cdate('" & Date & "')"
Data15.Refresh
If Len(Data15.Recordset.Fields(0) + 1) < 2 Then
DBCombo1(9).Text = "C" + Format(Date, "mmdd") + "-" + "0" + Trim(Data15.Recordset.Fields(0) + 1)
Else
DBCombo1(9).Text = "C" + Format(Date, "mmdd") + "-" + Trim(Data15.Recordset.Fields(0) + 1)
End If
Else
DBCombo1(9).Text = "C" + Format(Date, "mmdd") + "-" + "01"
End If
End Select
End Sub

Private Sub Label2_Click()
DBCombo1(13).Enabled = False
End Sub

Private Sub Label2_DBLClick()
DBCombo1(13).Enabled = True
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
For i = 0 To 2
DBCombo1(i + 1).Text = Data2.Recordset.Fields(i)
Next
DBCombo1(4).Text = Data2.Recordset.Fields(3)
DBCombo1(5).Text = Data2.Recordset.Fields(4)
DBCombo1(6).Text = Data2.Recordset.Fields(5)
Text1.Text = Data2.Recordset.Fields(6)
Text2.Text = Data2.Recordset.Fields(7)
End Sub

Private Sub MSFlexGrid2_dblClick()
On Error Resume Next
rs = MSFlexGrid2.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
DTPicker1.Value = Data1.Recordset.Fields(0)
For i = 0 To 13
DBCombo1(i).Text = Data1.Recordset.Fields(i + 1)
Next
Text1.Text = Data1.Recordset.Fields(16)
Text2.Text = Data1.Recordset.Fields(17)

Command3.Enabled = True
Command4.Enabled = True
Command1.Enabled = False
End Sub

Private Sub MSFlexGrid3_DBLClick()
On Error Resume Next
If Data10.Recordset.EOF Then Exit Sub
Data10.Recordset.MoveFirst
rs = MSFlexGrid3.Row
Data10.Recordset.Move rs - 1
DBCombo1(1).Text = Data10.Recordset.Fields(2)
End Sub

Private Sub MSFlexGrid4_dblClick()
If Data13.Recordset.EOF Then Exit Sub
Data13.Recordset.MoveFirst
rs = MSFlexGrid4.Row
Data13.Recordset.Move rs - 1
DBCombo1(0).Text = Data13.Recordset.Fields(0)
DBCombo1(1).Text = Data13.Recordset.Fields(1)
End Sub

Private Sub Timer1_Timer()
If BAR = 100 Then
Timer1.Enabled = False
BAR = 1
ProgressBar1.Visible = False
Call fhmxdy(Data2, Data3, DBCombo1(13).Text)
End If
BAR = BAR + 1
ProgressBar1.Value = BAR
End Sub

