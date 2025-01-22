VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy28 
   BackColor       =   &H00C0E0FF&
   Caption         =   "面料分析"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form28"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   13620
      TabIndex        =   21
      Text            =   "Text4"
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   13620
      TabIndex        =   20
      Text            =   "Text4"
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   13620
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   13620
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   13620
      TabIndex        =   17
      Text            =   "Text4"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   13620
      TabIndex        =   16
      Text            =   "Text4"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   13620
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   13620
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   13620
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   13620
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   13620
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   13620
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   1200
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
      Top             =   9240
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
      Top             =   9000
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
      Top             =   8760
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   1200
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
      Top             =   6480
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
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8880
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "Formy28.frx":0000
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Formy28.frx":0006
      Top             =   1680
      Width           =   2535
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
      Top             =   6480
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
      Top             =   6480
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
      Top             =   6480
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
      Top             =   6480
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
      Top             =   6480
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
      Top             =   6480
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
      Top             =   8880
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
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
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
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":000C
      Height          =   330
      Index           =   0
      Left            =   12720
      TabIndex        =   22
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy28.frx":0020
      Height          =   330
      Index           =   0
      Left            =   5640
      TabIndex        =   23
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   0
      Left            =   14160
      TabIndex        =   24
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   0
      Left            =   11760
      TabIndex        =   25
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":0034
      Height          =   330
      Index           =   0
      Left            =   10200
      TabIndex        =   26
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   0
      Left            =   6600
      TabIndex        =   27
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   0
      Left            =   4320
      TabIndex        =   28
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   0
      Left            =   2640
      TabIndex        =   29
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy28.frx":0048
      Height          =   330
      Index           =   0
      Left            =   960
      TabIndex        =   30
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   0
      Left            =   960
      TabIndex        =   31
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy28.frx":005C
      Height          =   2895
      Left            =   600
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7080
      Width           =   14415
      _ExtentX        =   25426
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
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   1
      Left            =   960
      TabIndex        =   33
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy28.frx":0070
      Height          =   330
      Index           =   2
      Left            =   960
      TabIndex        =   34
      Top             =   1200
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
      Left            =   4920
      TabIndex        =   35
      Top             =   720
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
      Left            =   4920
      TabIndex        =   36
      Top             =   1200
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
      Left            =   8520
      TabIndex        =   37
      Top             =   1200
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
      Left            =   8520
      TabIndex        =   38
      Top             =   1680
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
      Left            =   8520
      TabIndex        =   39
      Top             =   2160
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
      Left            =   8520
      TabIndex        =   40
      Top             =   2640
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
      Left            =   8520
      TabIndex        =   41
      Top             =   3120
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
      Left            =   8520
      TabIndex        =   42
      Top             =   3600
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
      Left            =   8520
      TabIndex        =   43
      Top             =   4080
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
      Left            =   8520
      TabIndex        =   44
      Top             =   4560
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
      Left            =   8520
      TabIndex        =   45
      Top             =   5040
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
      Left            =   8520
      TabIndex        =   46
      Top             =   5520
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
      Left            =   8520
      TabIndex        =   47
      Top             =   6000
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
      Left            =   11160
      TabIndex        =   48
      Top             =   0
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
      Left            =   11160
      TabIndex        =   49
      Top             =   0
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
      Left            =   12120
      TabIndex        =   50
      Top             =   8160
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
      Left            =   12120
      TabIndex        =   51
      Top             =   8160
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy28.frx":0084
      Height          =   330
      Index           =   1
      Left            =   960
      TabIndex        =   52
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy28.frx":0098
      Height          =   330
      Index           =   2
      Left            =   960
      TabIndex        =   53
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy28.frx":00AC
      Height          =   330
      Index           =   3
      Left            =   960
      TabIndex        =   54
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   1
      Left            =   2640
      TabIndex        =   55
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   2
      Left            =   2640
      TabIndex        =   56
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   3
      Left            =   2640
      TabIndex        =   57
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy28.frx":00C0
      Height          =   330
      Index           =   4
      Left            =   960
      TabIndex        =   58
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   4
      Left            =   2640
      TabIndex        =   59
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy28.frx":00D4
      Height          =   330
      Index           =   5
      Left            =   960
      TabIndex        =   60
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   5
      Left            =   2640
      TabIndex        =   61
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   6
      Left            =   2640
      TabIndex        =   62
      Top             =   5520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy28.frx":00E8
      Height          =   330
      Index           =   6
      Left            =   960
      TabIndex        =   63
      Top             =   5520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   7
      Left            =   2640
      TabIndex        =   64
      Top             =   6000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy28.frx":00FC
      Height          =   330
      Index           =   7
      Left            =   960
      TabIndex        =   65
      Top             =   6000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   1
      Left            =   4320
      TabIndex        =   66
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   2
      Left            =   4320
      TabIndex        =   67
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   3
      Left            =   4320
      TabIndex        =   68
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   4
      Left            =   4320
      TabIndex        =   69
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   5
      Left            =   4320
      TabIndex        =   70
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   6
      Left            =   4320
      TabIndex        =   71
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   7
      Left            =   4320
      TabIndex        =   72
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   1
      Left            =   6600
      TabIndex        =   73
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   2
      Left            =   6600
      TabIndex        =   74
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   3
      Left            =   6600
      TabIndex        =   75
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   4
      Left            =   6600
      TabIndex        =   76
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   5
      Left            =   6600
      TabIndex        =   77
      Top             =   5040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   6
      Left            =   6600
      TabIndex        =   78
      Top             =   5520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   7
      Left            =   6600
      TabIndex        =   79
      Top             =   6000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":0110
      Height          =   330
      Index           =   1
      Left            =   10200
      TabIndex        =   80
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":0124
      Height          =   330
      Index           =   2
      Left            =   10200
      TabIndex        =   81
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":0138
      Height          =   330
      Index           =   3
      Left            =   10200
      TabIndex        =   82
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":014C
      Height          =   330
      Index           =   4
      Left            =   10200
      TabIndex        =   83
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":0160
      Height          =   330
      Index           =   5
      Left            =   10200
      TabIndex        =   84
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":0174
      Height          =   330
      Index           =   6
      Left            =   10200
      TabIndex        =   85
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":0188
      Height          =   330
      Index           =   7
      Left            =   10200
      TabIndex        =   86
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":019C
      Height          =   330
      Index           =   8
      Left            =   10200
      TabIndex        =   87
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":01B0
      Height          =   330
      Index           =   9
      Left            =   10200
      TabIndex        =   88
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":01C4
      Height          =   330
      Index           =   10
      Left            =   10200
      TabIndex        =   89
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   1
      Left            =   11760
      TabIndex        =   90
      Top             =   1680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   2
      Left            =   11760
      TabIndex        =   91
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   3
      Left            =   11760
      TabIndex        =   92
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   4
      Left            =   11760
      TabIndex        =   93
      Top             =   3120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   5
      Left            =   11760
      TabIndex        =   94
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   6
      Left            =   11760
      TabIndex        =   95
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   7
      Left            =   11760
      TabIndex        =   96
      Top             =   4560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   8
      Left            =   11760
      TabIndex        =   97
      Top             =   5040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   9
      Left            =   11760
      TabIndex        =   98
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   10
      Left            =   11760
      TabIndex        =   99
      Top             =   6000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   1
      Left            =   14160
      TabIndex        =   100
      Top             =   1680
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
      Left            =   14160
      TabIndex        =   101
      Top             =   2160
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
      Left            =   14160
      TabIndex        =   102
      Top             =   2640
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
      Left            =   14160
      TabIndex        =   103
      Top             =   3120
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
      Left            =   14160
      TabIndex        =   104
      Top             =   3600
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
      Left            =   14160
      TabIndex        =   105
      Top             =   4080
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
      TabIndex        =   106
      Top             =   4560
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
      TabIndex        =   107
      Top             =   5040
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
      TabIndex        =   108
      Top             =   5520
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
      TabIndex        =   109
      Top             =   6000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Bindings        =   "Formy28.frx":01D8
      Height          =   330
      Index           =   11
      Left            =   10200
      TabIndex        =   110
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Height          =   330
      Index           =   11
      Left            =   11760
      TabIndex        =   111
      Top             =   6480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo8 
      Height          =   330
      Index           =   11
      Left            =   14160
      TabIndex        =   112
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo8"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy28.frx":01EC
      Height          =   330
      Index           =   1
      Left            =   5640
      TabIndex        =   113
      Top             =   3120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy28.frx":0200
      Height          =   330
      Index           =   2
      Left            =   5640
      TabIndex        =   114
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy28.frx":0214
      Height          =   330
      Index           =   3
      Left            =   5640
      TabIndex        =   115
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy28.frx":0228
      Height          =   330
      Index           =   4
      Left            =   5640
      TabIndex        =   116
      Top             =   4560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy28.frx":023C
      Height          =   330
      Index           =   5
      Left            =   5640
      TabIndex        =   117
      Top             =   5040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy28.frx":0250
      Height          =   330
      Index           =   6
      Left            =   5640
      TabIndex        =   118
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo9 
      Bindings        =   "Formy28.frx":0264
      Height          =   330
      Index           =   7
      Left            =   5640
      TabIndex        =   119
      Top             =   6000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo9"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":0278
      Height          =   330
      Index           =   1
      Left            =   12720
      TabIndex        =   120
      Top             =   1680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":028C
      Height          =   330
      Index           =   2
      Left            =   12720
      TabIndex        =   121
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":02A0
      Height          =   330
      Index           =   3
      Left            =   12720
      TabIndex        =   122
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":02B4
      Height          =   330
      Index           =   4
      Left            =   12720
      TabIndex        =   123
      Top             =   3120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":02C8
      Height          =   330
      Index           =   5
      Left            =   12720
      TabIndex        =   124
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":02DC
      Height          =   330
      Index           =   6
      Left            =   12720
      TabIndex        =   125
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":02F0
      Height          =   330
      Index           =   7
      Left            =   12720
      TabIndex        =   126
      Top             =   4560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":0304
      Height          =   330
      Index           =   8
      Left            =   12720
      TabIndex        =   127
      Top             =   5040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":0318
      Height          =   330
      Index           =   9
      Left            =   12720
      TabIndex        =   128
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":032C
      Height          =   330
      Index           =   10
      Left            =   12720
      TabIndex        =   129
      Top             =   6000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin MSDBCtls.DBCombo DBCombo10 
      Bindings        =   "Formy28.frx":0340
      Height          =   330
      Index           =   11
      Left            =   12720
      TabIndex        =   130
      Top             =   6480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo10"
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   3
      Left            =   13620
      TabIndex        =   172
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料单位"
      Height          =   375
      Index           =   1
      Left            =   12720
      TabIndex        =   171
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料单位"
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   170
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "其它"
      Height          =   375
      Index           =   11
      Left            =   7680
      TabIndex        =   169
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "商标"
      Height          =   375
      Index           =   10
      Left            =   7680
      TabIndex        =   168
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料顶扣"
      Height          =   375
      Index           =   9
      Left            =   7680
      TabIndex        =   167
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料后扣"
      Height          =   375
      Index           =   8
      Left            =   7680
      TabIndex        =   166
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "高头明线"
      Height          =   375
      Index           =   7
      Left            =   7680
      TabIndex        =   165
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "胶条"
      Height          =   375
      Index           =   6
      Left            =   7680
      TabIndex        =   164
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "间线"
      Height          =   375
      Index           =   5
      Left            =   7680
      TabIndex        =   163
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "帽芯"
      Height          =   375
      Index           =   4
      Left            =   7680
      TabIndex        =   162
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣眼"
      Height          =   375
      Index           =   3
      Left            =   7680
      TabIndex        =   161
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前页"
      Height          =   375
      Index           =   2
      Left            =   7680
      TabIndex        =   160
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "汉带"
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   159
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "双针"
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   158
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "主料后扣"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   157
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "主料顶扣"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   156
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前衬布料"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   155
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "下眉布料"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   154
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "上眉布料"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   153
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "后页布料"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   152
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "中页布料"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   151
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "前页布料"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   150
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料数量"
      Height          =   375
      Index           =   34
      Left            =   14160
      TabIndex        =   149
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料规格"
      Height          =   375
      Index           =   33
      Left            =   11760
      TabIndex        =   148
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料名称"
      Height          =   375
      Index           =   32
      Left            =   10200
      TabIndex        =   147
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "辅料明细"
      Height          =   375
      Index           =   31
      Left            =   7680
      TabIndex        =   146
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料名称"
      Height          =   375
      Index           =   30
      Left            =   960
      TabIndex        =   145
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "刀模"
      Height          =   375
      Index           =   29
      Left            =   2640
      TabIndex        =   144
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "主料明细"
      Height          =   375
      Index           =   44
      Left            =   120
      TabIndex        =   143
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料数量"
      Height          =   375
      Index           =   37
      Left            =   6600
      TabIndex        =   142
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料规格"
      Height          =   375
      Index           =   28
      Left            =   4320
      TabIndex        =   141
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "雅冠制帽定量备料分析表"
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
      TabIndex        =   140
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "工作编号"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   139
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "具体要求"
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   138
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "绣花(印刷)说明"
      Height          =   495
      Index           =   12
      Left            =   4080
      TabIndex        =   137
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   375
      Index           =   13
      Left            =   120
      TabIndex        =   136
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "数量"
      Height          =   375
      Index           =   14
      Left            =   4080
      TabIndex        =   135
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   15
      Left            =   11160
      TabIndex        =   134
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "交货期"
      Height          =   375
      Index           =   16
      Left            =   11160
      TabIndex        =   133
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "刀模"
      Height          =   375
      Index           =   18
      Left            =   4080
      TabIndex        =   132
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   19
      Left            =   120
      TabIndex        =   131
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "Formy28"
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

For i = 0 To 7
If DBCombo5(i).Text <> "" Then
rd.AddNew

rd.Fields(0) = DBCombo1(0).Text
rd.Fields(1) = DBCombo1(1).Text
rd.Fields(2) = DBCombo1(2).Text
rd.Fields(3) = Label3(i).Caption
rd.Fields(4) = DBCombo2(i).Text
rd.Fields(5) = DBCombo4(i).Text
rd.Fields(6) = DBCombo9(i).Text
rd.Fields(7) = Text3.Text
rd.Fields(8) = DBCombo5(i).Text
rd.Fields(9) = "1主料库"
rd.Update
End If
Next

For i = 0 To 11
If DBCombo8(i).Text <> "" Then
rd.AddNew
rd.Fields(0) = DBCombo1(0).Text
rd.Fields(1) = DBCombo1(1).Text
rd.Fields(2) = DBCombo1(2).Text
rd.Fields(3) = Label4(i).Caption
rd.Fields(4) = DBCombo6(i).Text
rd.Fields(5) = DBCombo7(i).Text
rd.Fields(6) = DBCombo10(i).Text
rd.Fields(7) = Text4(i).Text
rd.Fields(8) = DBCombo8(i).Text
rd.Fields(9) = "2辅料库"
rd.Update
End If
Next


Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh



End Sub

Private Sub Command2_Click()
On Error Resume Next

If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If


Data4.Recordset.Edit
For i = 0 To 7
If DBCombo5(i).Text <> "" Then
Data4.Recordset.Edit
Data4.Recordset.Fields(0) = DBCombo1(0).Text
Data4.Recordset.Fields(1) = DBCombo1(1).Text
Data4.Recordset.Fields(2) = DBCombo1(2).Text
Data4.Recordset.Fields(3) = Label3(i).Caption
Data4.Recordset.Fields(4) = DBCombo2(i).Text
Data4.Recordset.Fields(5) = DBCombo4(i).Text
Data4.Recordset.Fields(6) = DBCombo9(i).Text
Data4.Recordset.Fields(7) = Text3.Text
Data4.Recordset.Fields(8) = DBCombo5(i).Text
Data4.Recordset.Update
End If
Next

For i = 0 To 11
If DBCombo8(i).Text <> "" Then
Data4.Recordset.Edit
Data4.Recordset.Fields(0) = DBCombo1(0).Text
Data4.Recordset.Fields(1) = DBCombo1(1).Text
Data4.Recordset.Fields(2) = DBCombo1(2).Text
Data4.Recordset.Fields(3) = Label4(i).Caption
Data4.Recordset.Fields(4) = DBCombo6(i).Text
Data4.Recordset.Fields(5) = DBCombo7(i).Text
Data4.Recordset.Fields(6) = DBCombo10(i).Text
Data4.Recordset.Fields(7) = Text4(i).Text
Data4.Recordset.Fields(8) = DBCombo8(i).Text
Data4.Recordset.Update
End If
Next
Data4.Recordset.Update

Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh



End Sub

Private Sub Command4_Click()

On Error Resume Next

If MsgBox("确定删除吗?,删除将不能恢复!", vbYesNo) = vbNo Then Exit Sub


Data4.Recordset.Delete
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh


End Sub

Private Sub Command5_Click()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.EOF
For i = 0 To 7
If Label3(i).Caption = Data4.Recordset.Fields(3) Then
DBCombo2(i).Text = Data4.Recordset.Fields(4)
DBCombo4(i).Text = Data4.Recordset.Fields(5)
DBCombo5(i).Text = Format(Data4.Recordset.Fields(8), "#0.00")
DBCombo9(i).Text = Data4.Recordset.Fields(6)
End If
Next

For i = 0 To 11
If Label4(i).Caption = Data4.Recordset.Fields(3) Then
DBCombo6(i).Text = Data4.Recordset.Fields(4)
DBCombo7(i).Text = Data4.Recordset.Fields(5)
DBCombo8(i).Text = Format(Data4.Recordset.Fields(8), "#0.00")
DBCombo10(i).Text = Data4.Recordset.Fields(6)
Text4(i).Text = Data4.Recordset.Fields(7)
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
       Case 1
Data1.RecordSource = "select * from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "' AND SCZY_X.款号='" & DBCombo1(1).Text & "' "
Data1.Refresh

Data5.RecordSource = "select SCZY_x.颜色 from SCZY_x WHERE SCZY_x.单号='" & DBCombo1(0).Text & "' AND SCZY_X.款号='" & DBCombo1(1).Text & "' GROUP BY SCZY_X.颜色 "
Data5.Refresh

For i = 1 To Data1.Recordset.Fields.Count - 1
If i = 16 Then
Text1.Text = Data1.Recordset.Fields(i)
End If

If i = 17 Then
Text2.Text = Data1.Recordset.Fields(i)
End If

DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next

For i = 0 To 7
DBCombo3(i).Text = DBCombo1(4).Text
Next


Data3.RecordSource = "select * from SCZY_Z WHERE SCZY_Z.单号='" & DBCombo1(0).Text & "' "
Data3.Refresh

For i = 0 To 7
DBCombo2(i).Text = Data3.Recordset.Fields(2)
Next

Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' "
Data4.Refresh

Text3.Text = DBCombo1(2).Text
DBCombo1(3).Text = 12
     
     Case 2
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' "
Data4.Refresh
Text3.Text = DBCombo1(2).Text
End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
     Case 2
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' "
Data4.Refresh
Text3.Text = DBCombo1(2).Text
End Select

End Sub

Private Sub DBCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DBCombo2_Change(Index As Integer)
Select Case Index
       Case 0
 For i = 1 To 7
 DBCombo2(i).Text = DBCombo2(0).Text
 Next
End Select

End Sub

Private Sub DBCombo2_Click(Index As Integer, Area As Integer)
Select Case Index
       Case 0
 For i = 1 To 7
 DBCombo2(i).Text = DBCombo2(0).Text
 Next
End Select
 
End Sub

Private Sub DBCombo4_Change(Index As Integer)
Select Case Index
       Case 0
 For i = 1 To 7
 DBCombo4(i).Text = DBCombo4(0).Text
 Next
End Select
End Sub

Private Sub DBCombo4_Click(Index As Integer, Area As Integer)
Select Case Index
       Case 0
 For i = 1 To 7
 DBCombo4(i).Text = DBCombo4(0).Text
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

For i = 0 To 19
DBCombo1(i).Text = ""
Next

For i = 0 To 7
DBCombo2(i).Text = ""
Next

For i = 0 To 7
DBCombo3(i).Text = ""
Next

For i = 0 To 7
DBCombo4(i).Text = ""
Next

For i = 0 To 7
DBCombo5(i).Text = ""
Next

For i = 0 To 7
DBCombo9(i).Text = "米"
Next

For i = 0 To 11
DBCombo6(i).Text = ""
Next

For i = 0 To 11
DBCombo7(i).Text = ""
Next

For i = 0 To 11
DBCombo8(i).Text = ""
Next

For i = 0 To 11
DBCombo10(i).Text = ""
Next

For i = 0 To 11
Text4(i).Text = ""
Next

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
Data4.RecordSource = "select * from dlclb WHERE DLCLB.单号='" & DBCombo1(0).Text & "' AND DLCLB.款号='" & DBCombo1(1).Text & "' AND DLCLB.订单颜色='" & DBCombo1(2).Text & "' "
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"

Data6.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data6.RecordSource = "SELECT CKGL.材料名称 FROM CKGL WHERE CKGL.库类='1主料库' GROUP BY CKGL.材料名称 "
Data6.Refresh

Data7.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data7.RecordSource = "SELECT cldw.mc FROM cldw  GROUP BY cldw.mc"
Data7.Refresh

Data8.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data8.RecordSource = "SELECT CKGL.材料名称 FROM CKGL WHERE CKGL.库类='2辅料库' GROUP BY CKGL.材料名称 "
Data8.Refresh

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
       Case 11
       Label4(11).Caption = InputBox("", "主辅明细名称", "其它")
End Select
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
Data4.Recordset.MoveFirst
Data4.Recordset.Move rs - 1

For i = 0 To 7
If Data4.Recordset.Fields(3) = Label3(i).Caption Then
DBCombo2(i).Text = Data4.Recordset.Fields(4)
DBCombo4(i).Text = Data4.Recordset.Fields(5)
DBCombo5(i).Text = Format(Data4.Recordset.Fields(8), "#0.00")
DBCombo9(i).Text = Data4.Recordset.Fields(6)
Else
DBCombo2(i).Text = ""
DBCombo4(i).Text = ""
DBCombo5(i).Text = ""
DBCombo9(i).Text = ""
End If
Next

For i = 0 To 11
If Data4.Recordset.Fields(3) = Label4(i).Caption Then
DBCombo6(i).Text = Data4.Recordset.Fields(4)
DBCombo7(i).Text = Data4.Recordset.Fields(5)
DBCombo8(i).Text = Format(Data4.Recordset.Fields(8), "#0.00")
DBCombo10(i).Text = Data4.Recordset.Fields(6)
Text4(i).Text = Data4.Recordset.Fields(7)
Else
DBCombo6(i).Text = ""
DBCombo7(i).Text = ""
DBCombo8(i).Text = ""
DBCombo10(i).Text = ""
Text4(i).Text = ""
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



