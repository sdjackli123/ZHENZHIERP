VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw14 
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
   Begin VB.Data Data16 
      Caption         =   "Data16"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2775
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
      Top             =   9960
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
      Top             =   5760
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
         Format          =   92667905
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
         Format          =   92667905
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
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Top             =   5760
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4080
      TabIndex        =   39
      Top             =   6960
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
      Top             =   5760
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
      Top             =   6240
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
      Top             =   5760
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
      Top             =   6240
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
      Top             =   6240
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
      Top             =   5760
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
      Top             =   4440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   92667905
      CurrentDate     =   39910
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw14.frx":0000
      Height          =   330
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   4440
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
      Top             =   4440
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
      Top             =   4440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   4
      Left            =   11640
      TabIndex        =   10
      Top             =   4440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw14.frx":0014
      Height          =   2175
      Left            =   600
      TabIndex        =   11
      Top             =   1560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3836
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
      Bindings        =   "Formw14.frx":0028
      Height          =   3015
      Left            =   480
      TabIndex        =   12
      Top             =   6840
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   5318
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
      Left            =   13920
      TabIndex        =   13
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw14.frx":003C
      Height          =   330
      Index           =   6
      Left            =   480
      TabIndex        =   14
      Top             =   5280
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
      Top             =   5280
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
      Top             =   5280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   9
      Left            =   6840
      TabIndex        =   17
      Top             =   5280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw14.frx":0050
      Height          =   330
      Index           =   10
      Left            =   9240
      TabIndex        =   18
      Top             =   5280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   11
      Left            =   11640
      TabIndex        =   19
      Top             =   5280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   12
      Left            =   13920
      TabIndex        =   20
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw14.frx":0064
      Height          =   330
      Index           =   13
      Left            =   1320
      TabIndex        =   36
      Top             =   5760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "xm"
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Bindings        =   "Formw14.frx":0078
      Height          =   2175
      Left            =   8520
      TabIndex        =   51
      Top             =   1560
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3836
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
      Format          =   92667905
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
      Format          =   92667905
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
      Format          =   92667905
      CurrentDate     =   39883
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   14
      Left            =   1320
      TabIndex        =   61
      Top             =   6240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "装箱信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   62
      Top             =   1560
      Width           =   255
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
      Left            =   13920
      TabIndex        =   60
      Top             =   4080
      Width           =   975
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
      Caption         =   "销售信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   17
      Left            =   8280
      TabIndex        =   52
      Top             =   1560
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
      Top             =   6240
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
      Left            =   480
      TabIndex        =   35
      Top             =   4920
      Width           =   1575
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
      Left            =   6840
      TabIndex        =   34
      Top             =   4080
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
      Left            =   9240
      TabIndex        =   33
      Top             =   4080
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
      Top             =   4080
      Width           =   1935
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
      Index           =   1
      Left            =   4440
      TabIndex        =   31
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "  成  品  发  货"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   6600
      TabIndex        =   30
      Top             =   360
      Width           =   4935
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
      Left            =   11640
      TabIndex        =   29
      Top             =   4080
      Width           =   2175
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
      Top             =   4080
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
      Left            =   4440
      TabIndex        =   27
      Top             =   4920
      Width           =   2175
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
      Left            =   2280
      TabIndex        =   26
      Top             =   4920
      Width           =   1935
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
      Left            =   6840
      TabIndex        =   25
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发货地"
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
      Left            =   9240
      TabIndex        =   24
      Top             =   4920
      Width           =   2175
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
      Left            =   11640
      TabIndex        =   23
      Top             =   4920
      Width           =   2175
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
      Left            =   13920
      TabIndex        =   22
      Top             =   4920
      Width           =   975
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
      Top             =   5760
      Width           =   855
   End
End
Attribute VB_Name = "Formw14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BAR As Integer
Private Sub Command1_Click()
If DBCombo1(2).text = "" Then
MsgBox ("输入款号")
Exit Sub
End If

Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = DTPicker1.Value
For i = 0 To 14
Data1.Recordset.Fields(i + 1) = DBCombo1(i).text
Next
Data1.Recordset.Fields(16) = "未"
Data1.Recordset.Update
Data1.Refresh
Data2.Refresh
For i = 0 To 12
DBCombo1(i).text = ""
Next
DBCombo1(12).text = Data1.Recordset.RecordCount + 1
DBCombo1(8).text = 0
End Sub

Private Sub Command10_Click()
Data1.RecordSource = "SELECT 日期,购货单位,单号,款号,品名,单位,数量,单价,金额,发货地,备注,单据号 FROM CPFH WHERE 日期 BETWEEN CDATE('" & Text4.text & "') AND CDATE('" & Text5.text & "') ORDER BY 日期"
Data1.Refresh
Call OutDataToExcel2(MSFlexGrid2, 7, 9, "成品出库")
Data1.RecordSource = "SELECT * from CPFH WHERE 日期 between CDATE('" & Text4.text & "') AND CDATE('" & Text5.text & "') ORDER BY 日期"
Data1.Refresh
End Sub




Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
If DBCombo1(2).text = "" Then
MsgBox ("输入款号")
Exit Sub
End If

Data1.Recordset.Edit
Data1.Recordset.Fields(0) = DTPicker1.Value
For i = 0 To 14
Data1.Recordset.Fields(i + 1) = DBCombo1(i).text
Next
Data1.Recordset.Update
Data1.Refresh
Data2.Refresh
For i = 0 To 12
DBCombo1(i).text = ""
Next
DBCombo1(12).text = Data1.Recordset.RecordCount + 1
DBCombo1(7).text = 0
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True

End Sub

Private Sub Command4_Click()
If MsgBox("确定删除吗？，删除将不能恢复！", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
Data2.Refresh
For i = 0 To 12
DBCombo1(i).text = ""
Next
DBCombo1(12).text = Data1.Recordset.RecordCount + 1
DBCombo1(7).text = 0
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
DBCombo1(14).text = "00000001"

Data1.Database.Execute "UPDATE CPFH SET 单据号='00000000' WHERE 单据号=null"

Data8.DatabaseName = "D:\数据库\htgl\2011\CPCK"
Data8.RecordSource = "SELECT MAX(VAL(单据号)) FROM CPFH "
Data8.Refresh

If Data8.Recordset.EOF Then
DBCombo1(14).text = "00000001"
Else
DBCombo1(14).text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

End Sub


Private Sub Command7_Click()
DBCombo1(7).text = 0
DBCombo1(12).text = Data1.Recordset.RecordCount + 1
DBCombo1(6).text = "件"
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Command8_Click()
Data2.RecordSource = "select 客户,款号,箱号,颜色,规格1+规格2+规格3+规格4+规格5+规格6 as 规格,合计件,日期,编号 from zxd where 款号='" & DBCombo1(2).text & "'"
Data2.Refresh
Data13.RecordSource = "select * from cpfh where 款号='" & DBCombo1(2).text & "'"
Data13.Refresh
End Sub

Private Sub Command9_Click()
Data1.RecordSource = "SELECT * from CPFH WHERE 日期 between CDATE('" & Text4.text & "') AND CDATE('" & Text5.text & "') ORDER BY 日期"
Data1.Refresh
End Sub


Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 2
       Data13.RecordSource = "select 款号,日期,品名,单位,数量  FROM CPfh WHERE 款号='" & DBCombo1(2).text & "' ORDER BY 日期,VAL(序号) DESC "
       Data13.Refresh
       Case 7
       DBCombo1(9).text = Format(Val(DBCombo1(7).text) * Val(DBCombo1(8).text), "#0.00")
       Case 8
       DBCombo1(9).text = Format(Val(DBCombo1(7).text) * Val(DBCombo1(8).text), "#0.00")
       Case 14
       Data1.RecordSource = "SELECT * from CPFH WHERE 单据号='" & DBCombo1(14).text & "' ORDER BY VAL(序号) DESC"
       Data1.Refresh
End Select
End Sub

Private Sub DTPicker1_Change()
Data1.RecordSource = "SELECT * from CPfh WHERE 日期=CDATE('" & DTPicker1.Value & "') ORDER BY VAL(序号) DESC"
Data1.Refresh
DBCombo1(12).text = Data1.Recordset.RecordCount + 1
DBCombo1(7).text = 0
End Sub

Private Sub DTPicker1_CloseUp()
Data1.RecordSource = "SELECT * from CPFH WHERE 日期=CDATE('" & DTPicker1.Value & "') ORDER BY VAL(序号) DESC"
Data1.Refresh
DBCombo1(12).text = Data1.Recordset.RecordCount + 1
DBCombo1(7).text = 0
End Sub

Private Sub DTPicker2_Click()
Text4.text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text4.text = DTPicker2.Value
End Sub
Private Sub DTPicker3_Click()
Text5.text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
Text5.text = DTPicker3.Value
End Sub

Private Sub Form_Load()
On Error Resume Next
DTPicker1.Value = Date
For i = 0 To 13
DBCombo1(i).text = ""
Next
Text4.text = Date
Text5.text = Date
DTPicker3.Value = Date
DTPicker2.Value = Date
DTPicker4.Value = Date - 30
DTPicker5.Value = Date
DTPicker6.Value = Date
Data1.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"
Data1.RecordSource = "SELECT * from CPFH WHERE 单据号='" & DBCombo1(14).text & "' ORDER BY VAL(序号) DESC"
Data1.Refresh


ProgressBar1.Visible = False
Timer1.Enabled = False

Data2.DatabaseName = "D:\数据库\htgl\2011\SCJD.MDB"

Data3.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"

Data4.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"
Data4.RecordSource = "select FZR.XM  FROM FZR GROUP BY FZR.XM"
Data4.Refresh

Data5.DatabaseName = "D:\数据库\htgl\2011\CKGL.MDB"
Data5.RecordSource = "select CLDW.MC  FROM CLDW GROUP BY CLDW.MC"
Data5.Refresh

Data6.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"

Data7.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"
Data7.RecordSource = "select FHD.MC  FROM FHD GROUP BY FHD.MC"
Data7.Refresh

DBCombo1(14).text = "00000001"
DBCombo1(14).Enabled = False
Data1.Database.Execute "UPDATE CPFH SET 单据号='00000000' WHERE 单据号=null"

Data8.DatabaseName = "D:\数据库\htgl\2011\CPCK"
Data8.RecordSource = "SELECT MAX(VAL(单据号)) FROM CPFH "
Data8.Refresh

If Data8.Recordset.EOF Then
DBCombo1(14).text = "00000001"
Else
DBCombo1(14).text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

Data9.DatabaseName = "D:\数据库\htgl\2011\SCZYJHD.MDB"
Data9.RecordSource = "SELECT 简称 FROM KHZL GROUP BY 简称"
Data9.Refresh

Data10.DatabaseName = "D:\数据库\htgl\2011\SCZYJHD.MDB"
Data10.RecordSource = "SELECT 客户,工作编号,单号,款式 FROM SCZY_ZDH  WHERE B4=NULL OR B4<>'已' ORDER BY 工作编号"
Data10.Refresh

Data11.DatabaseName = "D:\数据库\htgl\2011\SCZYJHD.MDB"
Data12.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"
Data13.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"
Data15.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"
Data16.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"


DBCombo1(7).text = 0
DBCombo1(12).text = Data1.Recordset.RecordCount + 1
DBCombo1(6).text = "件"
DBCombo1(13).text = "王清德"


Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1400
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 1200
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(5) = 1200


MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(1) = 1200
MSFlexGrid2.ColWidth(2) = 1200
MSFlexGrid2.ColWidth(3) = 1500
MSFlexGrid2.ColWidth(4) = 1200
MSFlexGrid2.ColWidth(5) = 2000
MSFlexGrid2.ColWidth(6) = 600



MSFlexGrid4.ColWidth(0) = 200
MSFlexGrid4.ColWidth(1) = 1200
MSFlexGrid4.ColWidth(2) = 1200
MSFlexGrid4.ColWidth(3) = 1200


End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 5
khbl = 11
Formw202.Text1.text = DBCombo1(2).text
Formw202.Show
End Select
End Sub

Private Sub Label2_Click()
DBCombo1(14).Enabled = False
End Sub

Private Sub Label2_DBLClick()
DBCombo1(14).Enabled = True
End Sub

Private Sub Label3_dblClick()
If DBCombo1(2).text = "" Then
Data2.RecordSource = "SELECT 客户,款号,颜色,合计件,日期,编号 FROM ZXD WHERE 日期 BETWEEN CDATE('" & DTPicker4.Value & "') AND CDATE('" & DTPicker5.Value & "')"
Data2.Refresh
Else
Data2.RecordSource = "SELECT 客户,款号,颜色,合计件,日期,编号 FROM ZXD WHERE 款号='" & DBCombo1(2).text & "' and  日期 BETWEEN CDATE('" & DTPicker4.Value & "') AND CDATE('" & DTPicker5.Value & "')"
Data2.Refresh
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
DBCombo1(0).text = Data2.Recordset.Fields(0)
DBCombo1(2).text = Data2.Recordset.Fields(1)
DBCombo1(4).text = Data2.Recordset.Fields(4)
DBCombo1(7).text = Data2.Recordset.Fields(5)
End Sub

Private Sub MSFlexGrid2_dblClick()
On Error Resume Next
rs = MSFlexGrid2.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
DTPicker1.Value = Data1.Recordset.Fields(0)
For i = 0 To 13
DBCombo1(i).text = Data1.Recordset.Fields(i + 1)
Next
Command3.Enabled = True
Command4.Enabled = True
Command1.Enabled = False
End Sub

Private Sub MSFlexGrid3_DblClick()
On Error Resume Next
If Data10.Recordset.EOF Then Exit Sub
Data10.Recordset.MoveFirst
rs = MSFlexGrid3.Row
Data10.Recordset.Move rs - 1
DBCombo1(1).text = Data10.Recordset.Fields(2)
End Sub

Private Sub Timer1_Timer()
If BAR = 100 Then
Call cpck(Data16, DBCombo1(14).text)
Timer1.Enabled = False
ProgressBar1.Visible = False
Exit Sub
End If
BAR = BAR + 1
ProgressBar1.Value = BAR

End Sub

