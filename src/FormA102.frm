VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FormA102 
   BackColor       =   &H00C0E0FF&
   Caption         =   "生产计划"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15300
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   36
      Text            =   "Text9"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   "ok"
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
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "订单详情"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   33
      Text            =   "Text7"
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10440
      TabIndex        =   32
      Text            =   "Text6"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "打印预览"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   28
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   27
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H008080FF&
      Caption         =   "新锅号"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4680
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
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
      Height          =   360
      ItemData        =   "FormA102.frx":0000
      Left            =   12465
      List            =   "FormA102.frx":000A
      TabIndex        =   25
      Text            =   "Combo2"
      Top             =   5160
      Width           =   990
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   11880
      TabIndex        =   24
      Text            =   "Text3"
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "计划结束"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0FF&
      Caption         =   "计划取消"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4200
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染色"
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯"
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0000C0C0&
      Caption         =   "返修"
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton command2 
      BackColor       =   &H0080FF80&
      Caption         =   "毛坯库存"
      Height          =   1695
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "复制原锅号"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   5280
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "Text13"
      Top             =   5160
      Width           =   1215
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   3480
      Style           =   1  'Simple Combo
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "流卡工序"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "订单查询条件"
      Height          =   1095
      Left            =   4200
      TabIndex        =   0
      Top             =   3480
      Width           =   2055
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "款号"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单号"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "刷新"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc25 
      Height          =   330
      Left            =   7080
      Top             =   10440
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc25"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc24 
      Height          =   330
      Left            =   7560
      Top             =   10440
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc24"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc23 
      Height          =   330
      Left            =   7440
      Top             =   10560
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc23"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc22 
      Height          =   375
      Left            =   7560
      Top             =   10560
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc22"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc21 
      Height          =   330
      Left            =   7800
      Top             =   10680
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc21"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc20 
      Height          =   330
      Left            =   7800
      Top             =   10560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc20"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc19 
      Height          =   375
      Left            =   10800
      Top             =   10320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc19"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc18 
      Height          =   330
      Left            =   8880
      Top             =   10800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc18"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc17 
      Height          =   375
      Left            =   9600
      Top             =   10320
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc17"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc16 
      Height          =   375
      Left            =   7200
      Top             =   10440
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc16"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc15 
      Height          =   375
      Left            =   8040
      Top             =   10560
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc15"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   330
      Left            =   8760
      Top             =   10680
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc14"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   495
      Left            =   7560
      Top             =   10440
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc13"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   8040
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc12"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   8760
      Top             =   10800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc11"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   375
      Left            =   8520
      Top             =   10440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc10"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   9240
      Top             =   10440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   8520
      Top             =   10560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   8880
      Top             =   10440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   10200
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   8040
      Top             =   10560
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8280
      Top             =   10560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8040
      Top             =   10560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8280
      Top             =   10680
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8400
      Top             =   10560
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo8 
      Height          =   330
      Left            =   1560
      TabIndex        =   9
      Top             =   2880
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo6"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "FormA102.frx":0014
      Height          =   330
      Left            =   1680
      TabIndex        =   10
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "xm"
      Text            =   "DataCombo5"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   13560
      TabIndex        =   11
      Top             =   5160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "FormA102.frx":002A
      Height          =   330
      Left            =   8880
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "车台编号"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FormA102.frx":0040
      Height          =   330
      Left            =   1560
      TabIndex        =   13
      Top             =   2400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1440
      TabIndex        =   37
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330760193
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   38
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330760193
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1680
      TabIndex        =   39
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330760193
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1440
      TabIndex        =   40
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330760193
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   375
      Left            =   10440
      TabIndex        =   41
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330760193
      CurrentDate     =   39177
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   1
      Left            =   6840
      TabIndex        =   42
      Top             =   4080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   2
      Left            =   10800
      TabIndex        =   43
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   3
      Left            =   11760
      TabIndex        =   44
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   4
      Left            =   12720
      TabIndex        =   45
      Top             =   4080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   5
      Left            =   13560
      TabIndex        =   46
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   6
      Left            =   9480
      TabIndex        =   47
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   7
      Left            =   6840
      TabIndex        =   48
      Top             =   5160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   8
      Left            =   4200
      TabIndex        =   49
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo4"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "FormA102.frx":0055
      Height          =   2655
      Left            =   4200
      TabIndex        =   50
      Top             =   480
      Width           =   10455
      _cx             =   18441
      _cy             =   4683
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "FormA102.frx":006B
      Height          =   1695
      Left            =   5160
      TabIndex        =   51
      Top             =   5640
      Width           =   9615
      _cx             =   16960
      _cy             =   2990
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
      Bindings        =   "FormA102.frx":0080
      Height          =   2295
      Left            =   480
      TabIndex        =   52
      Top             =   7680
      Width           =   14295
      _cx             =   25215
      _cy             =   4048
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Height          =   330
      Left            =   5280
      TabIndex        =   53
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "工艺编号"
      Text            =   "DataCombo6"
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择负责人"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   78
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请输入锅号"
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   77
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "日期"
      Height          =   375
      Index           =   9
      Left            =   10440
      TabIndex        =   76
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   75
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   74
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户名称"
      Height          =   375
      Index           =   8
      Left            =   480
      TabIndex        =   73
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   72
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   71
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "品名"
      Height          =   375
      Index           =   6
      Left            =   6840
      TabIndex        =   70
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "毛胚幅宽(寸)"
      Height          =   375
      Index           =   1
      Left            =   10800
      TabIndex        =   69
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光胚幅宽(cm)"
      Height          =   375
      Index           =   2
      Left            =   11760
      TabIndex        =   68
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "匹数"
      Height          =   375
      Index           =   3
      Left            =   12720
      TabIndex        =   67
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "重量（公斤）"
      Height          =   375
      Index           =   4
      Left            =   13560
      TabIndex        =   66
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光坯克重"
      Height          =   375
      Index           =   7
      Left            =   4200
      TabIndex        =   65
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   375
      Index           =   8
      Left            =   6840
      TabIndex        =   64
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入单号"
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   63
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "整理类别"
      Height          =   375
      Index           =   11
      Left            =   12480
      TabIndex        =   62
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP"
      Height          =   375
      Index           =   12
      Left            =   11880
      TabIndex        =   61
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请输入合约号"
      Height          =   375
      Index           =   13
      Left            =   480
      TabIndex        =   60
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "毛坯备注"
      Height          =   375
      Index           =   14
      Left            =   13560
      TabIndex        =   59
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请输入原锅号"
      Height          =   375
      Index           =   15
      Left            =   2880
      TabIndex        =   58
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "机台"
      Height          =   375
      Index           =   16
      Left            =   8880
      TabIndex        =   57
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "色号-色别"
      Height          =   375
      Left            =   9480
      TabIndex        =   56
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "工序"
      Height          =   975
      Left            =   2280
      TabIndex        =   55
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "工序"
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   54
      Top             =   4680
      Width           =   1455
   End
End
Attribute VB_Name = "FormA102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public x As Integer: Public BI As Integer ''''BI PANDUAN CHURU KU BIANLIANG
Dim BA As Database: Dim rr As Integer: Public gh, k1, k2 As String: Public hg As Date: Dim BA3 As Database: Dim RD3 As Recordset
Public ZL As Single  ''''''重量变量
Rem ' 中间转换变量
Dim rs As Single: Dim RD1 As Recordset: Dim BA1 As Database: Public ll, c, r As Integer: Public lbj As Long
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command10_Click()
If DataCombo8.Text = "" Then
MsgBox ("请输入订单号")
Exit Sub
End If
If MsgBox("确定订单结束吗？" + DataCombo8.Text, vbYesNo) = vbNo Then Exit Sub
sql1 = "UPDATE SCZY_x SET 排布='Y' WHERE 单号='" & DataCombo8.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End Sub

Private Sub Command11_Click()
Formd332.Text1 = Text7.Text
Formd332.Show
'Forma110.Text1(0) = Text7.Text
'Forma110.Show
End Sub

Private Sub Command12_Click()
On Error Resume Next
If Option1.value = False And Option2.value = False And Option3.value = False Then
MsgBox ("请选择备活信息")
Exit Sub
End If

If Option1.value = True Then
Adodc23.RecordSource = "select MAX(cast(right(锅号,len(锅号)-PATINDEX('%-%',锅号)) as int)) as h  from kpd where 日期=' " & Text6.Text & "' AND 锅号 NOT like '%W%' and 锅号 NOT like '%F%' and 锅号 NOT like '%H%' and left(锅号,1)='" & yhdm & "'"
Adodc23.Refresh

Text7.Text = yhdm + Format(CDate(Text6.Text), "YYMMDD") + "-1"
If Adodc23.Recordset.EOF Then
Text7.Text = yhdm + Format(CDate(Text6.Text), "YYMMDD") + "-1"
Else
Text7.Text = yhdm + Format(CDate(Text6.Text), "YYMMDD") + "-" + Trim(Val(Adodc23.Recordset.Fields(0)) + 1)
End If
  Text3.Text = 1
  Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
End If

If Option2.value = True Then

Adodc23.RecordSource = "select MAX(cast(right(锅号,len(锅号)-PATINDEX('%W%',锅号)) as int)) as h  from kpd where 日期=' " & Text6.Text & "'  AND 锅号 like '%W%' and left(锅号,1)='" & yhdm & "'"
Adodc23.Refresh
Text7.Text = yhdm + Format(CDate(Text6.Text), "YYMMDD") + "W1"
If Adodc23.Recordset.EOF Then
Text7.Text = yhdm + Format(CDate(Text6.Text), "YYMMDD") + "W1"
Else
Text7.Text = yhdm + Format(CDate(Text6.Text), "YYMMDD") + "W" + Trim(Val(Adodc23.Recordset.Fields(0)) + 1)
End If
  Text3.Text = 1
  Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  Text3.Text = Adodc9.Recordset.Fields(0) + 1

End If

If Option3.value = True Then

Adodc23.RecordSource = "select MAX(cast(right(锅号,len(锅号)-PATINDEX('%F%',锅号)) as int)) as h from kpd where 日期=' " & Text6.Text & "'  AND  锅号 like '%F%' and left(锅号,1)='" & yhdm & "'"
Adodc23.Refresh
Text7.Text = yhdm + Format(CDate(Text6.Text), "YYMMDD") + "F1"
If Adodc23.Recordset.EOF Then
Text7.Text = yhdm + Format(CDate(Text6.Text), "YYMMDD") + "F1"
Else
Text7.Text = yhdm + Format(CDate(Text6.Text), "YYMMDD") + "F" + Trim(Val(Adodc23.Recordset.Fields(0)) + 1)
End If
  Text3.Text = 1
  Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  Text3.Text = Adodc9.Recordset.Fields(0) + 1

End If
End Sub


Private Sub Command13_Click()

End Sub

Private Sub Command17_Click()
If DataCombo8.Text = "" Then
MsgBox ("请输入订单号")
Exit Sub
End If
If MsgBox("确定取消结束吗？" + DataCombo8.Text, vbYesNo) = vbNo Then Exit Sub
sql1 = "UPDATE SCZY_x SET 排布='N' WHERE 单号='" & DataCombo8.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End Sub


Private Sub Command2_Click()
Command2.Enabled = False
       Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc6.RecordSource = "select * from  mpckgl3  where 客户名称='" & DataCombo1.Text & "' and  重量<>0"
       Adodc6.Refresh
Command2.Enabled = True
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If Text11.Text = "" Then
MsgBox ("请输入原锅号")
Exit Sub
End If


If Text7.Text = "" Then
MsgBox ("请输入锅号")
Exit Sub
End If

If MsgBox("要复制原锅号" + Text11.Text + "新锅号为" + Text7.Text + "吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "insert into kpd(客户名称,单号,锅号,色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,日期,备注,技术要求,IP,标签,kp,kp1,CKY,负责人,pb,rs,ts,xdx,ddx,fh) select 客户名称,单号,'" & Text7.Text & "',色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,'" & Date & "',备注,技术要求,IP,标签,'N','N',CKY,负责人,'Y','N','N','N','N','N' from kpd where 锅号='" & Text11.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc8.RecordSource = "select 客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,标签 as 合约号,备注,技术要求,类别,CKY as 毛坯备注,车台  from kpd where 锅号='" & Text7.Text & "' order by IP"
Adodc8.Refresh

End Sub

Private Sub Command5_Click()
On Error Resume Next

Call lcd3(Adodc6, Adodc7, Text7.Text, DataCombo4(1))

End Sub

Private Sub Command6_Click()
If DataCombo1.Text = "" Then
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,标签 as 合约号,备注,技术要求,类别,CKY as 毛坯备注,车台,GX AS 工序  from kpd where 日期 between '" & DTPicker3.value & "' and '" & DTPicker4.value & "' and left(锅号,1)='" & yhdm & "'"
Adodc8.Refresh
Else
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,标签 as 合约号,备注,技术要求,类别,CKY as 毛坯备注,车台,GX AS 工序  from kpd where 客户名称='" & DataCombo1.Text & "' and 日期 between '" & DTPicker3.value & "' and '" & DTPicker4.value & "' and left(锅号,1)='" & yhdm & "'"
Adodc8.Refresh
End If
End Sub


Private Sub Command1_Click()
On Error Resume Next
If DataCombo5.Text = "" Then
MsgBox ("请选择负责人！")
Exit Sub
End If

If DataCombo1.Text = "" Then
MsgBox ("请输入客户！")
Exit Sub
End If

If Text7.Text = "" Then
MsgBox ("请输入锅号！")
Exit Sub
End If

If Text3.Text = "" Then Text3.Text = 1

    Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "mpkpd('" & DataCombo1.Text & "','" & DataCombo8.Text & "','" & Text7.Text & "','" & DataCombo4(1).Text & "','" & DataCombo4(2).Text & "','" & DataCombo4(3).Text & "','" & DataCombo4(4).Text & "','" & DataCombo4(5).Text & "','" & DataCombo4(6).Text & "','" & DataCombo4(7).Text & "','" & DataCombo4(8).Text & "','" & Text3.Text & "','" & Text6.Text & "','" & Text9.Text & "','" & DataCombo2.Text & "','" & DataCombo5.Text & "','" & DataCombo3.Text & "','N','N','" & Combo2.Text & "','N','N','N','N','N','N','N','" & Text13.Text & "','N','N')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

Rem 'shuaxin 开票单

Adodc8.RecordSource = "select 客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,标签 as 合约号,备注,技术要求,类别,CKY as 毛坯备注,车台,GX AS 工序  from kpd where 锅号='" & Text7.Text & "' "
Adodc8.Refresh

  
  Text3.Text = 1
  Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  Text3.Text = Adodc9.Recordset.Fields(0) + 1

End Sub


Private Sub Command8_Click()
On Error Resume Next
Rem 'shuaxin 开票单
Adodc8.RecordSource = "select 客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,标签 as 合约号,备注,技术要求,类别,CKY as 毛坯备注,车台,GX AS 工序  from kpd where 锅号='" & Text7.Text & "' "
Adodc8.Refresh

  Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  Text3.Text = Adodc9.Recordset.Fields(0) + 1



Text7.Enabled = True

End Sub


Private Sub Command9_Click()
On Error Resume Next
If Text3.Text = "" Then Exit Sub
If MsgBox("确定删除" + Text3.Text + "吗？", vbYesNo) = vbNo Then Exit Sub
Adodc8.Recordset.Delete
Adodc8.Refresh
Text3.Text = 1
Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
Adodc9.Refresh
Text3.Text = Adodc9.Recordset.Fields(0) + 1
End Sub

Private Sub DataCombo1_Change()
On Error Resume Next
 ww = 0
If Text4.Text = "" Or Text5.Text = "" Then
Exit Sub
End If
RQ = CDate(Text5.Text)
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
 Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
 Adodc9.Refresh
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''订单计划信息
       Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc6.RecordSource = "select * from  mpckgl3  where 客户名称='" & DataCombo1.Text & "' and 重量<>0"
       Adodc6.Refresh

End Sub

Private Sub DataCombo1_Click(Area As Integer)
On Error Resume Next

 ww = 0
If Text4.Text = "" Or Text5.Text = "" Then
End If
RQ = CDate(Text5.Text)
op = 0.5
  Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''订单计划信息
       Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc6.RecordSource = "select * from  mpckgl3  where 客户名称='" & DataCombo1.Text & "' and 重量<>0"
       Adodc6.Refresh

End Sub

Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub dataCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub dataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub dataCombo4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub DataCombo5_Click(Area As Integer)
If DataCombo5.Text = "" Then
Adodc22.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc22.RecordSource = "select 客户,单号,款号,品名,色别,幅宽,克重,备注,计划,累计,交期,序号 from sczykpd where 日期 between '" & Text4.Text & "'  and  '" & Text5.Text & "'  and 排布<>'Y'  order by 客户,日期 "
Adodc22.Refresh
Else
Adodc22.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc22.RecordSource = "select 客户,单号,款号,品名,色别,幅宽,克重,备注,计划,累计,交期,序号 from sczykpd where 负责='" & DataCombo5.Text & "' and 日期 between '" & Text4.Text & "'  and  '" & Text5.Text & "'  and 排布<>'Y'  order by 客户,日期 "
Adodc22.Refresh
End If
End Sub

Private Sub dataCombo5_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub DataCombo6_Click(Area As Integer)
Text13.Text = DataCombo6.Text
End Sub

Private Sub DataCombo6_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo7_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo8_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.value
End Sub

Private Sub DTPicker1_CloseUp()
Text4.Text = DTPicker1.value
Text4.SetFocus
End Sub

Private Sub DTPicker2_Change()
Text5.Text = DTPicker2.value
End Sub

Private Sub DTPicker2_CloseUp()
Text5.Text = DTPicker2.value
Text5.SetFocus
End Sub

Private Sub DTPicker5_Change()
Text6.Text = DTPicker5.value
End Sub

Private Sub DTPicker5_CloseUp()
Text6.Text = DTPicker5.value
End Sub
Private Sub Form_Load()
On Error Resume Next

DataCombo6.Text = ""
DataCombo8.Text = ""
Combo2.Text = "圆筒"
DTPicker1.value = Date - 30
DTPicker2.value = Date
DTPicker3.value = Date
DTPicker4.value = Date
DataCombo1.Text = ""
Text11.Text = ""
Text4.Text = Date - 30
Text5.Text = Date
Text1.Text = ""

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 简称  from khzl  group by 简称"
Adodc3.Refresh

Text3.Text = ""
Text7.Text = ""
Text9.Text = ""
DataCombo3.Text = ""
DataCombo2.Text = ""
DataCombo3.Enabled = False
Text13.Text = ""
Text11.Text = ""
DTPicker5.value = Date
Text6.Text = Date
Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.RecordSource = "select xm from ywf group by xm"
Adodc12.Refresh

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc19.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc19.RecordSource = "SCZY_X"
Adodc19.Refresh

Adodc20.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.RecordSource = "select distinct 工艺编号 from gyshd where 工艺编号 between '0001' and '1000'"
Adodc14.Refresh

Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc17.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc18.RecordSource = "SELECT 车台编号 FROM CT GROUP BY 车台编号"
Adodc18.Refresh


Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc23.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc24.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"



Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select * from  mpckgl3  where 客户名称='" & DataCombo1.Text & "' and 重量<>0"
Adodc6.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc25.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc25.RecordSource = "select MC from JSYQ group by MC"
Adodc25.Refresh


DataCombo5.Text = ""


DataCombo4(4).Enabled = True
DataCombo4(5).Enabled = True



DataCombo7.Text = ""

For i = 1 To 8
DataCombo4(i).Text = ""
Next




VSFlexGrid3.ColWidth(0) = 100
VSFlexGrid3.ColWidth(1) = 1000
VSFlexGrid3.ColWidth(2) = 1800
VSFlexGrid3.ColWidth(3) = 1200
VSFlexGrid3.ColWidth(4) = 1200
VSFlexGrid3.ColWidth(5) = 1600
VSFlexGrid3.ColWidth(6) = 1800


VSFlexGrid2.ColWidth(0) = 100
VSFlexGrid2.ColWidth(1) = 1500
VSFlexGrid2.ColWidth(2) = 1500
VSFlexGrid2.ColWidth(3) = 1500
VSFlexGrid2.ColWidth(4) = 1500
VSFlexGrid2.ColWidth(5) = 1500
VSFlexGrid2.ColWidth(6) = 1200
VSFlexGrid2.ColWidth(7) = 1500


VSFlexGrid4.ColWidth(0) = 100
VSFlexGrid4.ColWidth(2) = 1500
VSFlexGrid4.ColWidth(3) = 500
VSFlexGrid4.ColWidth(4) = 1600
VSFlexGrid4.ColWidth(8) = 1000
VSFlexGrid4.ColWidth(9) = 1800

ZL = 0

Text4.TabIndex = 0
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label10_Click()

End Sub



Private Sub Label1_Click()
ysbl = 2
Forma38.Text1.Text = DataCombo4(6).Text
Forma38.Show
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 6
'Forma113.Text1.Text = DataCombo4(1).Text
'Forma113.Show
       Case 8
beizhu = 11
Forma112.Show
       Case 14
DataCombo3.Enabled = False
End Select
End Sub

Private Sub Label9_Click()
Form18.Text1.Text = DataCombo2.Text
Form18.Show
End Sub

Private Sub Label2_DblClick(Index As Integer)
Select Case Index
       Case 14
DataCombo3.Enabled = True
End Select
End Sub

Private Sub Label4_Click()
FormA101.Show
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
If Adodc20.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc20.Recordset.MoveFirst
Adodc20.Recordset.Move rs - 1
DataCombo4(6).Text = Adodc20.Recordset.Fields(0)

End Sub

Private Sub Label14_DblClick()
DataCombo4(5).Enabled = True
End Sub

Private Sub Label15_DblClick()
Label12.Caption = Format(DataCombo4(5).Text, "###0.00")
       Combo1.Text = "入库"
       
       BI = 0
       Adodc13.RecordSource = "select 客户名称,布类,存放位置,毛胚幅宽,毛胚重量,实际投放量,毛胚匹数,备注,IP from ckgl  where 客户名称='" & DataCombo1.Text & " ' and 布类='" & DataCombo4(1).Text & " ' and 毛胚幅宽='" & DataCombo4(2).Text & " ' and    VAL(毛胚重量)-VAL(实际投放量)>=0 AND VAL(实际投放量)>0 order by 存放位置"
       Adodc13.Refresh
End Sub


Private Sub Label5_Click()
sql1 = ""

If Check2(1).value = 1 Then
sql1 = sql1 + "客户 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "单号 like '%'+'" & DataCombo8.Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "款号 like '%'+'" & Text9.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc22.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc22.RecordSource = "select 客户,单号,款号,品名,色别,幅宽,克重,备注,计划,累计,交期,序号 from sczykpd where (" + sql1 + ")  and 排布<>'Y'  order by 客户,单号,款号,色别"
Adodc22.Refresh

End Sub

Private Sub Text1_Change()
       Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc6.RecordSource = "select * from  mpckgl3  where 客户名称='" & DataCombo1.Text & "' and  布类 like '%'+'" & Text1.Text & "'+'%'  and  重量<>0"
       Adodc6.Refresh
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc22.Recordset.EOF Then
DataCombo2.Text = ""
Exit Sub
End If
rs = VSFlexGrid2.Row
Adodc22.Recordset.MoveFirst
Adodc22.Recordset.Move rs - 1
DataCombo1.Text = Adodc22.Recordset.Fields(0)
DataCombo8.Text = Adodc22.Recordset.Fields(1)
Text9.Text = Adodc22.Recordset.Fields(2)
DataCombo4(1).Text = Adodc22.Recordset.Fields(3)
DataCombo4(6).Text = Adodc22.Recordset.Fields(4)
DataCombo4(3).Text = Adodc22.Recordset.Fields(5)
DataCombo4(8).Text = Adodc22.Recordset.Fields(6)
DataCombo4(7).Text = Adodc22.Recordset.Fields(7)
Text3.Text = Adodc22.Recordset.Fields(11)
End Sub

Private Sub VSFlexGrid3_dblClick()
On Error Resume Next
If Adodc6.Recordset.EOF Then Exit Sub
rs = VSFlexGrid3.Row
Adodc6.Recordset.MoveFirst
Adodc6.Recordset.Move rs - 1
DataCombo4(1).Text = Adodc6.Recordset.Fields(1)
DataCombo4(2).Text = Adodc6.Recordset.Fields(2)
DataCombo3.Text = Adodc6.Recordset.Fields(5)
DataCombo4(4).Text = Adodc6.Recordset.Fields(3)
DataCombo4(5).Text = Adodc6.Recordset.Fields(4)
End Sub

Private Sub VSFlexGrid4_dblClick()
If Adodc8.Recordset.EOF Then Exit Sub
rs = VSFlexGrid4.Row
Adodc8.Recordset.MoveFirst
Adodc8.Recordset.Move rs - 1
Text3.Text = Adodc8.Recordset.Fields(2)
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Text7_Change()
On Error Resume Next
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,标签 as 合约号,备注,技术要求,类别,CKY as 毛坯备注,车台,GX AS 工序  from kpd where 锅号='" & Text7.Text & "' "
Adodc8.Refresh
  Text3.Text = 1
  
  Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub sx()
If Adodc20.Recordset.EOF Then Exit Sub
Adodc20.Recordset.MoveFirst
i = 1
Do While Not Adodc20.Recordset.EOF
VSFlexGrid1.col = 3
VSFlexGrid1.Row = i
VSFlexGrid1.Text = Format(Adodc20.Recordset.Fields(2), "##0.0")
Adodc20.Recordset.MoveNext
i = i + 1
Loop

End Sub

Private Sub MSFlex()
With VSFlexGrid4
    c = .col: r = .Row    '''''C列，，R行
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
End With
End Sub


Private Sub vSFlexGrid4_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid4.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc8.Recordset.MoveFirst
Adodc8.Recordset.Move r - 1
Adodc8.Recordset.Fields(c - 1) = Combo1111.Text
Adodc8.Recordset.Update
Combo1111.Visible = False
VSFlexGrid4.Text = Combo1111.Text
VSFlexGrid4.SetFocus
End If
End Sub


