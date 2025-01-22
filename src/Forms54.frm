VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Forms54 
   BackColor       =   &H00C0E0FF&
   Caption         =   "圆筒定型标签操作"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form54"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   375
      Left            =   10320
      Top             =   9600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   10320
      TabIndex        =   70
      Text            =   "Text7"
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   15480
      TabIndex        =   68
      Text            =   "Text6"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "码单打印选择"
      Height          =   1815
      Left            =   10680
      TabIndex        =   64
      Top             =   5040
      Width           =   1450
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF80&
         Caption         =   "重量"
         Height          =   370
         Left            =   360
         TabIndex        =   66
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFF80&
         Caption         =   "重码"
         Height          =   370
         Left            =   360
         TabIndex        =   65
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.TextBox Text5 
      Height          =   370
      Left            =   9360
      TabIndex        =   63
      Text            =   "Text5"
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "132码单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "日期自动"
      Enabled         =   0   'False
      Height          =   615
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "产量查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   6120
      Top             =   360
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "72码单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
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
      Height          =   615
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "标签打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
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
      Height          =   615
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "校正毛胚"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "触摸扫描"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6600
      Top             =   360
   End
   Begin VB.TextBox TxtSend 
      Height          =   375
      Left            =   10200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Forms54.frx":0000
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxtReceive 
      Height          =   375
      Left            =   10200
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "Forms54.frx":0007
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF80&
      Caption         =   "半自"
      Height          =   375
      Left            =   12480
      TabIndex        =   7
      Top             =   6000
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF80&
      Caption         =   "手动"
      Height          =   375
      Left            =   12480
      TabIndex        =   6
      Top             =   5160
      Width           =   975
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7560
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   360
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6360
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFF80&
      Caption         =   "自动"
      Height          =   375
      Left            =   12480
      TabIndex        =   4
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   9480
      Top             =   10560
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
      Left            =   9840
      Top             =   10440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Left            =   9600
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
      Height          =   330
      Left            =   9600
      Top             =   10320
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
      Left            =   9840
      Top             =   10560
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
      Height          =   495
      Left            =   9720
      Top             =   10320
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Left            =   9960
      Top             =   10320
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
      Height          =   330
      Left            =   9840
      Top             =   10320
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forms54.frx":000E
      Height          =   2055
      Left            =   720
      TabIndex        =   22
      Top             =   720
      Width           =   17535
      _cx             =   30930
      _cy             =   3625
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms54.frx":0023
      Height          =   4095
      Left            =   720
      TabIndex        =   23
      Top             =   5640
      Width           =   7575
      _cx             =   13361
      _cy             =   7223
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   0
      Left            =   3120
      TabIndex        =   24
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   11880
      TabIndex        =   25
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330301441
      CurrentDate     =   39177
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   1
      Left            =   720
      TabIndex        =   26
      Top             =   3360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   2
      Left            =   8280
      TabIndex        =   27
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   3
      Left            =   5160
      TabIndex        =   28
      Top             =   3360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Forms54.frx":0038
      Height          =   330
      Index           =   4
      Left            =   10680
      TabIndex        =   29
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   5
      Left            =   15360
      TabIndex        =   30
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   6
      Left            =   3120
      TabIndex        =   31
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   7
      Left            =   5160
      TabIndex        =   32
      Top             =   4320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   8
      Left            =   720
      TabIndex        =   33
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   9
      Left            =   16920
      TabIndex        =   34
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   10
      Left            =   14040
      TabIndex        =   35
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Forms54.frx":004D
      Height          =   330
      Index           =   11
      Left            =   13440
      TabIndex        =   36
      Top             =   3360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   12
      Left            =   12600
      TabIndex        =   37
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Forms54.frx":0062
      Height          =   330
      Index           =   13
      Left            =   8280
      TabIndex        =   38
      Top             =   4320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "名称"
      Text            =   "DataCombo1"
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   4440
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forms54.frx":0077
      Height          =   615
      Left            =   720
      TabIndex        =   39
      Top             =   9720
      Width           =   7575
      _cx             =   13361
      _cy             =   1085
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   14
      Left            =   16920
      TabIndex        =   40
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   15
      Left            =   720
      TabIndex        =   41
      Top             =   5040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   16
      Left            =   3120
      TabIndex        =   42
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label3 
      Caption         =   "打印份数"
      Height          =   615
      Left            =   9360
      TabIndex        =   69
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作"
      Height          =   255
      Index           =   2
      Left            =   15480
      TabIndex        =   67
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "码数系数"
      Height          =   375
      Index           =   15
      Left            =   9360
      TabIndex        =   62
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      Caption         =   "光胚重量"
      Height          =   255
      Left            =   16920
      TabIndex        =   61
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "班次"
      Height          =   255
      Index           =   0
      Left            =   13440
      TabIndex        =   60
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "锅号扫描"
      Height          =   495
      Index           =   10
      Left            =   720
      TabIndex        =   59
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "布类"
      Height          =   255
      Index           =   9
      Left            =   5160
      TabIndex        =   58
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "类别"
      Height          =   255
      Index           =   1
      Left            =   10680
      TabIndex        =   57
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光胚幅宽"
      Height          =   255
      Index           =   2
      Left            =   15360
      TabIndex        =   56
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "匹数"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   55
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "重量（公斤）"
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   54
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "色别与色号"
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   53
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户名称"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   52
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
      Height          =   255
      Index           =   1
      Left            =   11880
      TabIndex        =   51
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "匹号"
      Height          =   255
      Index           =   0
      Left            =   12600
      TabIndex        =   50
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "克重"
      Height          =   255
      Index           =   6
      Left            =   14040
      TabIndex        =   49
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   255
      Index           =   7
      Left            =   8280
      TabIndex        =   48
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   255
      Index           =   11
      Left            =   8280
      TabIndex        =   47
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "回潮"
      Height          =   375
      Index           =   12
      Left            =   9360
      TabIndex        =   46
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "光胚码数"
      Height          =   255
      Left            =   16920
      TabIndex        =   45
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "缸号"
      Height          =   255
      Index           =   13
      Left            =   720
      TabIndex        =   44
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   255
      Index           =   14
      Left            =   3120
      TabIndex        =   43
      Top             =   4800
      Width           =   1575
   End
End
Attribute VB_Name = "Forms54"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gk As Integer
    Dim SendCount  As Long     '定义已发送字符对应字节数
    Dim ReceiveCount  As Long  '定义已接收字符对应字节数
    Dim PortSwitch As Boolean    '定义串口是否打开标志
    Public L As String
    Dim DisplayFlag As Boolean   '定义接收窗口是否继续显示标志
Dim ZHT  As String
Dim xh As Integer      ''''''循环
Dim sl As Integer      ''''自动打印
Dim dybl As Integer    '''判断加工或销售的合同变量
Dim dzcbl As Integer  '''''电子称变量
Dim sssj, sswd As Single    '''实时传输数据
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim fs As Integer
Dim cdbhf As Integer
Private Sub Command1_Click()
On Error Resume Next

If Val(DataCombo4(9)) < 0.1 Then
DataCombo4(9).SetFocus
Exit Sub
End If

If DataCombo4(9).Text = "0" Then
MsgBox ("请选择班次！")
Exit Sub
End If

If DataCombo4(1).Text = "" Then
MsgBox ("请输入锅号！")
Exit Sub
End If

DataCombo4(6) = Int(Val(DataCombo4(6)))

If Val(DataCombo4(7)) = 0 Then
MsgBox ("请输入正确的毛坯匹数和重量")
Exit Sub
End If

Timer2.Enabled = False

If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If


Adodc8.RecordSource = "select * from bmd where 锅号='" & DataCombo4(1) & "' and 缸号='" & DataCombo4(15) & "' and 匹号='" & DataCombo4(12) & "'"
Adodc8.Refresh
If Not Adodc8.Recordset.EOF Then
MsgBox ("已有此匹号，请注意匹号编号")
Call Command6_Click
If Option1.value = False Then
Timer2.Enabled = True
'Timer3.Enabled = True
sl = 1
dzcbl = 1
End If
Exit Sub
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "mpbmdlr('" & DataCombo4(0).Text & "','" & DataCombo4(1).Text & "','" & DataCombo4(2).Text & "','" & DataCombo4(3).Text & "','" & DataCombo4(4).Text & "','" & DataCombo4(5).Text & "','" & DataCombo4(6).Text & "','" & DataCombo4(7).Text & "','" & DataCombo4(8).Text & "','" & DataCombo4(9).Text & "','" & DataCombo4(10).Text & "','" & DataCombo4(11).Text & "','" & DataCombo4(12).Text & "','" & DTPicker4.value & "','小定型','" & DataCombo4(13).Text & "','" & DataCombo4(14).Text & "','" & DataCombo4(15).Text & "','" & DataCombo4(16).Text & "','" & Text6 & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

Adodc1.RecordSource = "select * from bmd where  锅号='" & DataCombo4(1).Text & "' and 缸号='" & DataCombo4(15).Text & "'  order by 匹号 desc"
Adodc1.Refresh
fs = Val(Text7.Text)    '''''打印份数
If fs <= 0 Then
        MsgBox "打印份数必须大于0"
        Exit Sub
    End If
   
Call xbq(Adodc5, DataCombo4(1).Text, DataCombo4(15).Text, Text7.Text, DataCombo4(5).Text)
 
Adodc2.RecordSource = "select max(匹号) from bmd where 锅号='" & DataCombo4(1).Text & "' and  缸号='" & DataCombo4(15).Text & "' "
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If

Adodc7.RecordSource = "select count(匹号) as 合计匹数,round(sum(光胚重量),2) as 合计重量 from bmd where  锅号='" & DataCombo4(1).Text & "' and 缸号='" & DataCombo4(15).Text & "' "
Adodc7.Refresh
If Option1.value = True Then
DataCombo4(9) = "0.1"
Else
DataCombo4(9) = 0
Timer2.Enabled = True
End If
DataCombo4(13) = ""
DataCombo4(9).SetFocus

End Sub

Private Sub Command10_Click()

If Option4.value = True Then
Call xmd100(Adodc5, DataCombo4(1), DataCombo4(15))
End If

If Option5.value = True Then
Call xmd100ms(Adodc5, DataCombo4(1), DataCombo4(15))
End If

Adodc1.RecordSource = "select * from bmd where  锅号='" & DataCombo4(1).Text & "' and 缸号='" & DataCombo4(15).Text & "' and 定型='小定型'  order by 匹号 desc"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
'Set g_Cmd = New Command
'    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
'    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
'    g_Cmd.CommandText = "MPbmzk('" & DataCombo4(1).Text & "','" & DataCombo4(15).Text & "')"   ' 表示调用哪个存储过程
 '   g_Cmd.Execute           ' 执行存储过程
 '   g_Cmd.Cancel
End If
End Sub

Private Sub Command11_Click()
sql1 = "update bmd_mdxz set 码单重量=毛胚重量,码单匹数=毛胚匹数 where 锅号='" & DataCombo4(1).Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("校正成功！")
Adodc1.Refresh
End Sub

Private Sub Command12_Click()
On Error Resume Next
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "MPbmzk('" & DataCombo4(1).Text & "','" & DataCombo4(15).Text & "')"   ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Forms511.Text12 = bzgrbh
Forms511.Text2 = DataCombo4(1) + "J"
Forms511.Show
End Sub

Private Sub Command13_Click()
On Error Resume Next
Timer2.Enabled = False

If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If

If DataCombo4(9).Text = "0" Then
MsgBox ("请选择班次！")
Exit Sub
End If

If DataCombo4(1).Text = "" Then
MsgBox ("请输入锅号！")
Exit Sub
End If

If DataCombo4(4).Text = "" Then
MsgBox ("请输入整理类别！")
Exit Sub
End If



Adodc8.RecordSource = "select * from bmd where 锅号='" & DataCombo4(1) & "' and 缸号='" & DataCombo4(15) & "' and 匹号='" & DataCombo4(12) & "'"
Adodc8.Refresh
If Not Adodc8.Recordset.EOF Then
MsgBox ("已有此匹号，请注意匹号编号")
Exit Sub
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "mpbmdlr('" & DataCombo4(0).Text & "','" & DataCombo4(1).Text & "','" & DataCombo4(2).Text & "','" & DataCombo4(3).Text & "','" & DataCombo4(4).Text & "','" & DataCombo4(5).Text & "','" & DataCombo4(6).Text & "','" & DataCombo4(7).Text & "','" & DataCombo4(8).Text & "','" & DataCombo4(9).Text & "','" & DataCombo4(10).Text & "','" & DataCombo4(11).Text & "','" & DataCombo4(12).Text & "','" & DTPicker4.value & "','小定型','" & DataCombo4(13).Text & "','" & DataCombo4(14).Text & "','" & DataCombo4(15).Text & "','" & DataCombo4(16).Text & "','" & Text6 & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
    
Adodc1.RecordSource = "select * from bmd where  锅号='" & DataCombo4(1).Text & "' and 缸号='" & DataCombo4(15).Text & "'  order by 匹号 desc"
Adodc1.Refresh

Adodc2.RecordSource = "select max(匹号) from bmd where 锅号='" & DataCombo4(1).Text & "' and  缸号='" & DataCombo4(15).Text & "' "
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If

Adodc7.RecordSource = "select count(匹号) as 合计匹数,round(sum(光胚重量),2) as 合计重量 from bmd where  锅号='" & DataCombo4(1).Text & "' and 品名='" & DataCombo4(3).Text & "' and 光胚幅宽='" & DataCombo4(5) & "' and 克重='" & DataCombo4(10) & "'"
Adodc7.Refresh

DataCombo4(9).SetFocus
Timer2.Enabled = True

End Sub

Private Sub Command2_Click()
On Error Resume Next
If DataCombo4(9).Text = "0" Then
MsgBox ("请选择班次！")
Exit Sub
End If

If DataCombo4(1).Text = "" Then
MsgBox ("请输入锅号！")
Exit Sub
End If

DataCombo4(6) = Int(Val(DataCombo4(6)))

If Val(DataCombo4(7)) = 0 Then
MsgBox ("请输入正确的毛坯匹数和重量")
Exit Sub
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "mpbmdxg('" & DataCombo4(0).Text & "','" & DataCombo4(1).Text & "','" & DataCombo4(2).Text & "','" & DataCombo4(3).Text & "','" & DataCombo4(4).Text & "','" & DataCombo4(5).Text & "','" & DataCombo4(6).Text & "','" & DataCombo4(7).Text & "','" & DataCombo4(8).Text & "','" & DataCombo4(9).Text & "','" & DataCombo4(10).Text & "','" & DataCombo4(11).Text & "','" & DataCombo4(12).Text & "','" & DTPicker4.value & "','小定型','" & DataCombo4(13).Text & "','" & DataCombo4(15).Text & "','" & DataCombo4(16).Text & "','" & Text6 & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel


Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

Call Command6_Click

DataCombo4(13) = ""
DataCombo4(9).SetFocus
End Sub


Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "mpbmdsc1('" & DataCombo4(1).Text & "','" & DataCombo4(12).Text & "','" & DataCombo4(15).Text & "','" & DataCombo4(16) & "')"    ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

Call Command6_Click
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
DataCombo4(9).SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
On Error Resume Next
fs = Val(Text7.Text)    '''''打印份数
If fs <= 0 Then
        MsgBox "打印份数必须大于0"
        Exit Sub
    End If
    
Call xbq(Adodc5, DataCombo4(1).Text, DataCombo4(15).Text, Text7.Text, DataCombo4(5).Text)

 
Adodc2.RecordSource = "select max(匹号) from bmd where 锅号='" & DataCombo4(1).Text & "' and  缸号='" & DataCombo4(15).Text & "' "
Adodc2.Refresh

DataCombo4(9).Text = "0"
DataCombo4(12).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If

Adodc1.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

DataCombo4(9).SetFocus

End Sub

Private Sub Command6_Click()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from bmd where  锅号='" & DataCombo4(1).Text & "' and 缸号='" & DataCombo4(15).Text & "'  order by 匹号 desc"
Adodc1.Refresh

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select max(匹号) from bmd where 锅号='" & DataCombo4(1).Text & "' and  缸号='" & DataCombo4(15).Text & "' "
Adodc2.Refresh

If Option1.value = True Then
DataCombo4(9).Text = "0.1"
Else
DataCombo4(9) = 0
End If
DataCombo4(12).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If
DataCombo4(9).SetFocus
Adodc7.RecordSource = "select count(匹号) as 合计匹数,round(sum(光胚重量),2) as 合计重量 from bmd where  锅号='" & DataCombo4(1).Text & "' and 品名='" & DataCombo4(3).Text & "' and 光胚幅宽='" & DataCombo4(5) & "' and 克重='" & DataCombo4(10) & "'"
Adodc7.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command7_Click()
    ' 查询JGMXKF表，如果Adodc9不为空则退出
    Adodc9.RecordSource = "select * from JGMXKF where 锅号='" & DataCombo4(1).Text & "' and 缸号='" & DataCombo4(15).Text & "'"
    Adodc9.Refresh
    
    ' 检查Adodc9是否不为空
    If Not Adodc9.Recordset.EOF Then
        Exit Sub ' 如果Adodc9有数据，直接退出
    End If
    
    ' 如果Adodc9为空，执行以下代码
    Adodc1.RecordSource = "select * from bmd where 锅号='" & DataCombo4(1).Text & "' and 定型='小定型' order by 匹号 desc"
    Adodc1.Refresh
    
    ' 检查Adodc1是否有记录
    If Not Adodc1.Recordset.EOF Then
        ' 创建并执行存储过程
        Set g_Cmd = New Command
        g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
        g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
        g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
        g_Cmd.CommandText = "MPbmzk('" & DataCombo4(1).Text & "','" & DataCombo4(15).Text & "')"   ' 调用存储过程
        g_Cmd.Execute                           ' 执行存储过程
        g_Cmd.Cancel                            ' 取消命令对象
    End If
End Sub


Private Sub Command8_Click()
Forms509.Text1(4) = DataCombo4(11)
Forms509.Check2(4).value = 1
Forms509.Check2(0).value = 1
Forms509.Timer1.Enabled = True
Forms509.Show
End Sub

Private Sub Command9_Click()
Timer1.Enabled = True
End Sub

Private Sub DataCombo4_Change(Index As Integer)
On Error Resume Next
Select Case Index
Case 1
If InStr(DataCombo4(1).Text, "J") > 0 Or InStr(DataCombo4(1).Text, "j") > 0 Then

DataCombo4(1).Text = Mid(DataCombo4(1), 1, Len(DataCombo4(1).Text) - 1)
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 客户名称,锅号,款号,品名,图案,光胚幅宽,isnull(配缸匹数,0) as 匹数,isnull(配缸重量,0) as 重量,色别+色名 as 色别,克重,类别,'' as 总备注,'' as 合同部门,缸号,序号 from v_kpd_ok where 锅号='" & DataCombo4(1).Text & "' and 类别='圆筒'  ORDER BY 序号"
Adodc3.Refresh

VSFlexGrid2.AutoSizeMode = flexAutoSizeRowHeight
VSFlexGrid2.AutoSize 0, VSFlexGrid2.Cols - 1, False, 30

If Adodc3.Recordset.EOF Then
For i = 0 To 10
If i = 1 Then i = i + 1
DataCombo4(i).Text = ""
Next
Else
For i = 2 To 8
DataCombo4(i).Text = Adodc3.Recordset.Fields(i)
Next
If DataCombo4(6) <> Int(DataCombo4(6)) Then
DataCombo4(6) = Int(Val(DataCombo4(6))) + 1
End If
''DataCombo4(5) = Val(DataCombo4(5)) * 100
DataCombo4(0).Text = Adodc3.Recordset.Fields(0)
DataCombo4(4).Text = Adodc3.Recordset.Fields(10)
DataCombo4(10).Text = Adodc3.Recordset.Fields(9)
DataCombo4(15).Text = Adodc3.Recordset.Fields(13)
DataCombo4(16).Text = Adodc3.Recordset.Fields(14)
If Option1.value = True Then
DataCombo4(9) = "0.1"
Else
DataCombo4(9) = 0
End If
DataCombo4(9).SetFocus
End If

Else
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 客户名称,锅号,款号,品名,图案,光胚幅宽,isnull(配缸匹数,0) as 匹数,isnull(配缸重量,0) as 重量,色别+色名 as 色别,克重,类别,'' as 总备注,'' as 合同部门,缸号,序号 from v_kpd_ok where 锅号='" & DataCombo4(1).Text & "' and 类别='圆筒'  ORDER BY 序号"
Adodc3.Refresh

VSFlexGrid2.AutoSizeMode = flexAutoSizeRowHeight
VSFlexGrid2.AutoSize 0, VSFlexGrid2.Cols - 1, False, 30

End If

Case 16

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from bmd where  锅号='" & DataCombo4(1).Text & "' and 缸号='" & DataCombo4(15).Text & "'  order by 匹号 desc"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select isnull(max(匹号),0) from bmd where 锅号='" & DataCombo4(1).Text & "' and  缸号='" & DataCombo4(15).Text & "' "
Adodc2.Refresh

DataCombo4(12).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If


Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select count(匹号) as 合计匹数,round(sum(光胚重量),2) as 合计重量 from bmd where  锅号='" & DataCombo4(1).Text & "' and 缸号='" & DataCombo4(15).Text & "' "
Adodc7.Refresh

Case 9
If Option3.value = True And Val(DataCombo4(9)) = Val(Text1) Then
sssj = 0
xh = 1
Timer4.Enabled = True
End If

DataCombo4(14) = Format(Val(DataCombo4(9)) * Val(Text5), "#0.00")
End Select

End Sub


Private Sub DataCombo4_GotFocus(Index As Integer)
Select Case Index
       Case 1
       Case 9
DataCombo4(9) = "0.1"
End Select
End Sub

Private Sub DataCombo4_LostFocus(Index As Integer)
Select Case Index
       Case 9
If Val(DataCombo4(9)) = 0 Then
'DataCombo4(9).SetFocus
End If
End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
sql2 = "insert into yhcd(用户,菜单,编号) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.WindowState = 2
Formm1.Adodc1.Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sql2 = "delete from yhcd where 用户='" & yhm & "' and 编号='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.Adodc1.Refresh
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 10
DataCombo4(1) = ""
For i = 0 To 10
If i = 1 Then i = i + 1
DataCombo4(i).Text = ""
Next
DataCombo4(15).Text = ""
DataCombo4(16).Text = ""
Call Command6_Click
DataCombo4(1).SetFocus
End Select
End Sub

Private Sub Option1_Click()
DataCombo4(9) = "0.1"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Option2_Click()
DataCombo4(9) = "0"
DataCombo4(9).SetFocus
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Option3_Click()
DataCombo4(9) = "0"
DataCombo4(9).SetFocus
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Text5_Change()
DataCombo4(14) = Format(Val(DataCombo4(9)) * Val(Text5), "#0.00")
End Sub

Private Sub Timer2_Timer()
On Error Resume Next                           ''''''''''''配套使用效果更好
If Option1.value = True Then
If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''经典没有端口就退出
        End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''半自动
'If jmg = "10E7661011AE6DCF" Or jmg = "10E7665011AE6E0F" Or jmg = "10E7662E11AE6DED" Or jmg = "10E7660411AE6DC3" Then   '''''''''''''''''''''''''''加密狗
If Option2.value = True Then
If MSComm.PortOpen = False Then
            MSComm.PortOpen = True
If Err.Number = 8002 Then Exit Sub
        End If
MSComm.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until MSComm.InBufferCount >= 12
a = MSComm.Input
If Mid(a, 1, 2) = "wn" Then
clsz = Format(Val(Mid(a, 3, 7)) + Val(Text1), "#0.0")
DataCombo4(9) = clsz
End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''全自动
If Option3.value = True Then
If dzcbl = 0 Then DataCombo4(9) = 0
If MSComm.PortOpen = False Then
            MSComm.PortOpen = True
If Err.Number = 8002 Then Exit Sub
        End If
MSComm.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until MSComm.InBufferCount >= 12
a = MSComm.Input
If Mid(a, 1, 2) = "wn" Then
clsz = Format(Val(Mid(a, 3, 7)) + Val(Text1), "#0.0")
DataCombo4(9) = clsz
End If
End If
'End If                  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''加密狗
End Sub

Private Sub Timer3_Timer()
'If jmg = "10E7661011AE6DCF" Or jmg = "10E7665011AE6E0F" Or jmg = "10E7662E11AE6DED" Or jmg = "10E7660411AE6DC3" Then   '''''''''''''''''''''''''''加密狗
If Option3.value = True Then
If Val(DataCombo4(9)) > 1 Then
sl = sl + 1
Else
sl = 1
End If
If sl = 3 And dzcbl = 1 Then
Timer3.Enabled = False
Call Command1_Click
dzcbl = 0
xh = 1
End If
End If
'End If                                                  ''''''''''''''''''''''''''''''加密狗
End Sub

Private Sub Timer4_Timer()
If xh / 2 = Int(xh / 2) And Val(DataCombo4(9)) > 1 Then
sssj = Val(DataCombo4(9))
End If
If sssj = DataCombo4(9) Then
sswd = sswd + 1
Else
sswd = 0
End If
If sswd = 2 Then
Call Command1_Click
Timer4.Enabled = False
End If
xh = xh + 1
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Option3.value = True Or Option2.value = True Then
Exit Sub
End If


If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
For i = 2 To 10
DataCombo4(i).Text = Adodc1.Recordset.Fields(i)
Next
DTPicker4.value = Adodc1.Recordset.Fields(13)
DataCombo4(12).Text = Adodc1.Recordset.Fields(12)
DataCombo4(11).Text = Adodc1.Recordset.Fields(11)
DataCombo4(0).Text = Adodc1.Recordset.Fields(0)
Text6 = Adodc1.Recordset.Fields("操作")
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc3.Recordset.Move rs - 1
For i = 2 To 8
DataCombo4(i).Text = Adodc3.Recordset.Fields(i)
Next
If DataCombo4(6) <> Int(DataCombo4(6)) Then
DataCombo4(6) = Int(Val(DataCombo4(6))) + 1
End If
''DataCombo4(5) = Val(DataCombo4(5)) * 100
DataCombo4(0).Text = Adodc3.Recordset.Fields(0)
DataCombo4(4).Text = Adodc3.Recordset.Fields(10)
DataCombo4(10).Text = Adodc3.Recordset.Fields(9)
DataCombo4(15).Text = Adodc3.Recordset.Fields(13)
DataCombo4(16).Text = Adodc3.Recordset.Fields(14)
If Option1.value = True Then
DataCombo4(9) = "0.1"
Else
DataCombo4(9) = 0
End If
DataCombo4(9).SetFocus
End Sub

Private Sub dataCombo4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub


Private Sub Form_Load()
On Error Resume Next

       TxtReceive.Text = ""
       TxtSend = ""
       MSComm.CommPort = 1
       MSComm.Settings = "9600,n,8,1"
       MSComm.InBufferSize = 1024            ' 设置接收缓冲区为1024字节
       MSComm.OutBufferSize = 1024           ' 设置发送缓冲区为4096字节
       MSComm.InBufferCount = 0              ' 清空输入缓冲区
       MSComm.OutBufferCount = 0             ' 清空输出缓冲区
       MSComm.SThreshold = 1                 ' 发送缓冲区空触发发送事件
       MSComm.RThreshold = 1                 ' 每X个字符到接收缓冲区引起触发接收事件
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ID1 As Long
Dim ID2 As Long
If GetID(ID1, ID2, DevicePath) = 0 Then
jmg = Hex(ID1) + Hex(ID2)           '''''''''''''''加密狗
Else
jmg = ""
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''
Option4.value = True
Option1.value = True
Timer2.Enabled = True
dzcbl = 0
For i = 0 To 16
DataCombo4(i).Text = ""
Next
DTPicker4.value = Date
Text1.Text = "0"
Text7.Text = "1"
DataCombo4(13).Text = ""
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Text5 = 0
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from bmd where  锅号='" & DataCombo4(1).Text & "' and 缸号='" & DataCombo4(15).Text & "' order by 匹号 desc"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select max(匹号) from bmd where 锅号='" & DataCombo4(1).Text & "' and  缸号='" & DataCombo4(15).Text & "' "
Adodc2.Refresh

DataCombo4(12).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If
Text6 = ""
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select distinct mc,xh from bc order by xh"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT distinct 名称  FROM bmdzjyy"
Adodc6.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

dybl = 2
Timer1.Enabled = False
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(12) = 2000
VSFlexGrid2.ColWidth(13) = 1500

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(4) = 0
VSFlexGrid1.ColWidth(5) = 0
VSFlexGrid1.ColWidth(7) = 0
VSFlexGrid1.ColWidth(8) = 0
VSFlexGrid1.ColWidth(12) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(15) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 0
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(21) = 0
VSFlexGrid1.ColWidth(22) = 0

End Sub


Private Sub Timer1_Timer()
DTPicker4.value = Date
End Sub




