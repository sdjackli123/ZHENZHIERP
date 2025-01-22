VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formj31 
   BackColor       =   &H00C0E0FF&
   Caption         =   "车间生产表"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   Icon            =   "FormJ31.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "排产"
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   960
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   3960
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Left            =   6720
      Top             =   10800
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      Left            =   6720
      Top             =   10800
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Left            =   6600
      Top             =   10560
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
      Height          =   375
      Left            =   6840
      Top             =   10560
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      Left            =   6840
      Top             =   10560
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Left            =   6360
      Top             =   10800
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
      Height          =   375
      Left            =   6840
      Top             =   10560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   6840
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
      Left            =   6600
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
      Left            =   6600
      Top             =   10800
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FormJ31.frx":440A
      Height          =   450
      Left            =   1920
      TabIndex        =   46
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "车台编号"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   360
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3480
      Top             =   120
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "整车台删除"
      Height          =   1455
      Left            =   120
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Text1111 
      Height          =   735
      Left            =   3480
      TabIndex        =   42
      Text            =   "Text2"
      Top             =   7680
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H000000FF&
      Caption         =   "正排法"
      Height          =   375
      Left            =   2520
      TabIndex        =   41
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000FF&
      Caption         =   "倒排法"
      Height          =   375
      Left            =   360
      TabIndex        =   40
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   495
      Left            =   10440
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细"
      Height          =   495
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转"
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   16
      Left            =   12120
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   15
      Left            =   12120
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   14
      Left            =   12120
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   13
      Left            =   12120
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   12
      Left            =   6120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   11
      Left            =   6120
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   10
      Left            =   6120
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   9
      Left            =   6120
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   8
      Left            =   6120
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   7
      Left            =   6120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   6
      Left            =   480
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      Height          =   495
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转换车台"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "编号确认"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "操作区"
      Height          =   3135
      Left            =   360
      TabIndex        =   20
      Top             =   1560
      Width           =   18255
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   47
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   30
         Left            =   9720
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   29
         Left            =   9720
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   28
         Left            =   9720
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   27
         Left            =   9720
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   26
         Left            =   4800
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   25
         Left            =   4800
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   24
         Left            =   4800
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   23
         Left            =   4800
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   22
         Left            =   120
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   21
         Left            =   120
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   20
         Left            =   120
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   19
         Left            =   120
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   18
         Left            =   11760
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H008080FF&
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   17
         Left            =   11760
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   2160
         Width           =   375
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   48
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   49
         Top             =   1200
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   4
         Left            =   600
         TabIndex        =   50
         Top             =   1680
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   5
         Left            =   600
         TabIndex        =   51
         Top             =   2160
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   6
         Left            =   600
         TabIndex        =   52
         Top             =   2640
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   7
         Left            =   6240
         TabIndex        =   53
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   8
         Left            =   6240
         TabIndex        =   54
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   9
         Left            =   6240
         TabIndex        =   55
         Top             =   1200
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   10
         Left            =   6240
         TabIndex        =   56
         Top             =   1680
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   11
         Left            =   6240
         TabIndex        =   57
         Top             =   2160
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   12
         Left            =   6240
         TabIndex        =   58
         Top             =   2640
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   13
         Left            =   12240
         TabIndex        =   59
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   14
         Left            =   12240
         TabIndex        =   60
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   15
         Left            =   12240
         TabIndex        =   61
         Top             =   1200
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   16
         Left            =   12240
         TabIndex        =   62
         Top             =   1680
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   17
         Left            =   12240
         TabIndex        =   63
         Top             =   2160
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   18
         Left            =   12240
         TabIndex        =   64
         Top             =   2640
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   19
         Left            =   600
         TabIndex        =   65
         Top             =   3240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   20
         Left            =   600
         TabIndex        =   66
         Top             =   3720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   21
         Left            =   600
         TabIndex        =   67
         Top             =   4200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   22
         Left            =   600
         TabIndex        =   68
         Top             =   4680
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   23
         Left            =   5520
         TabIndex        =   69
         Top             =   3240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   24
         Left            =   5520
         TabIndex        =   70
         Top             =   3720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   25
         Left            =   5520
         TabIndex        =   71
         Top             =   4200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   26
         Left            =   5520
         TabIndex        =   72
         Top             =   4680
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   27
         Left            =   10560
         TabIndex        =   73
         Top             =   3240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   28
         Left            =   10560
         TabIndex        =   74
         Top             =   3720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   29
         Left            =   10560
         TabIndex        =   75
         Top             =   4200
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   330
         Index           =   30
         Left            =   10560
         TabIndex        =   76
         Top             =   4680
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   14640
         Y1              =   3180
         Y2              =   3180
      End
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   77
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "FormJ31.frx":441F
      Height          =   330
      Left            =   8280
      TabIndex        =   79
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "车台编号"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "FormJ31.frx":4434
      Height          =   5415
      Left            =   480
      TabIndex        =   78
      Top             =   4920
      Width           =   18135
      _cx             =   31988
      _cy             =   9551
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
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
      FixedCols       =   0
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
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "请选择要转到的车台："
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请选择车台编号："
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Formj31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public zhct As String: Public c, r As Integer
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public zd As Integer: Public zhci As Integer



Private Sub Command1_Click()
On Error Resume Next

 m = DataCombo1.Text

Adodc11.RecordSource = "select * from jhb where 车台编号= '" & DataCombo1.Text & "' "
Adodc11.Refresh

If Adodc11.Recordset.EOF Then Exit Sub
    zd = 0
    For c = 1 To Adodc11.Recordset.Fields.count - 1
    If Adodc11.Recordset.Fields(c) <> "" Then
    zd = zd + 1
    End If
    Next
    

If zd = 1 Then Exit Sub '一条记录则退出
'判断是否有重号

For i = 2 To zd
If Text1(i).Text = Text1(1).Text Then
MsgBox ("有重号，确认后重输！")
Exit Sub
End If
Next
'判断ZD是否大于编号
For i = 1 To zd
If Val(Text1(i).Text) > zd Then
MsgBox ("wrong number!")
Exit Sub
End If
Next

'tiao zhong biao hao

For i = 1 To zd
Adodc11.Recordset.Fields(Val(Text1(i).Text)) = DataCombo2(i).Text
Next
Adodc11.Recordset.Update
Adodc2.Refresh
DataCombo1.Text = ""
DataCombo1.Text = m

DataCombo1.SetFocus


End Sub


Private Sub Command11_Click()
On Error Resume Next

If MsgBox("确定删除车台" + Adodc2.Recordset.Fields(0) + "的所有排产内容吗？", vbYesNo) = vbYes Then
Adodc2.Recordset.Delete
Adodc2.Refresh
Else
Exit Sub
End If
End Sub

Private Sub Command12_Click()
'Call jhbOutadodcToExcel(VSFlexGrid1, "日期：" + Trim(Now))
Formj18.Show
End Sub

Private Sub Command13_Click()
On Error Resume Next
Dim JCGH As String
Command13.Enabled = False
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select *  from JHB ORDER BY 车台编号"
Adodc8.Refresh

Adodc8.Recordset.MoveFirst
Do While Not Adodc8.Recordset.EOF
For iii = 1 To 6

JCGH = Mid(Adodc8.Recordset.Fields(iii), InStr(Adodc8.Recordset.Fields(iii), "锅号") + 2)

If InStr(JCGH, "单染") > 0 Then
Adodc10.RecordSource = "select 锅号,count(锅号) from RSCL where 锅号=left('" & JCGH & "',len('" & JCGH & "')-2) and 工序 like '%出缸%' group by 锅号 having count(锅号)=1"
Adodc10.Refresh
If Not Adodc10.Recordset.EOF Then
Call jc(Adodc8.Recordset.Fields(0), Trim(iii)) ''''车台编号，序号
End If
End If

If InStr(JCGH, "双染") > 0 Then
Adodc10.RecordSource = "select 锅号,count(锅号) from RSCL where 锅号=left('" & JCGH & "',len('" & JCGH & "')-2) and 工序 like '%出缸%' group by 锅号 having count(锅号)=2"
Adodc10.Refresh
If Not Adodc10.Recordset.EOF Then
Call jc(Adodc8.Recordset.Fields(0), Trim(iii)) ''''车台编号，序号
End If
End If

If InStr(JCGH, "三染") > 0 Then
Adodc10.RecordSource = "select 锅号,count(锅号) from RSCL where 锅号=left('" & JCGH & "',len('" & JCGH & "')-2) and 工序 like '%出缸%' group by 锅号 having count(锅号)=3"
Adodc10.Refresh
If Not Adodc10.Recordset.EOF Then
Call jc(Adodc8.Recordset.Fields(0), Trim(iii))   ''''车台编号，序号
End If
End If

'If InStr(JCGH, "单染") = 0 And InStr(JCGH, "双染") = 0 And InStr(JCGH, "三染") = 0 Then
'Adodc10.RecordSource = "select 锅号,count(锅号) from RSCL where 锅号=left('" & JCGH & "',len('" & JCGH & "')-2) and 工序 like '%出缸%' group by 锅号 having count(锅号)=1"
'Adodc10.Refresh
'If Not Adodc10.Recordset.EOF Then
'Call jc(Adodc8.Recordset.Fields(0), Trim(iii)) ''''车台编号，序号
'End If
'End If

Next
Adodc8.Recordset.MoveNext
Loop
Command13.Enabled = True
Call DataCombo1_Change
Adodc2.RecordSource = "select *  from JHB ORDER BY 车台编号"
Adodc2.Refresh
End Sub

Private Sub Command2_Click()
On Error Resume Next
If zhci = 0 Then Exit Sub
 If zhct = "" Then
 MsgBox ("转换车台内容不能为空,请选择内容!")
 Exit Sub
 End If
  m = DataCombo1.Text
If DataCombo1.Text = "" Or DataCombo3.Text = "" Then
MsgBox ("车台不能为空!")
Exit Sub
End If
If MsgBox("转车的锅为：" + zhct + "从" + DataCombo1.Text + "转至" + DataCombo3.Text + "吗?", vbYesNo) = vbNo Then
Label1.Visible = False
DataCombo3.Visible = False
Exit Sub
End If
If DataCombo1.Text = "" Or DataCombo3.Text = "" Then Exit Sub

Adodc11.RecordSource = "select * from jhb where 车台编号= '" & DataCombo1.Text & "' "
Adodc11.Refresh
If Adodc11.Recordset.EOF Then
  MsgBox ("要转的车台编号错误")
  Exit Sub
  Else
   'shanchu
    zd = 0
    For c = 1 To Adodc11.Recordset.Fields.count - 1
      If Adodc11.Recordset.Fields(c) <> "" Then
      zd = zd + 1
      End If
    Next
      If zd = 1 And zhci = zd Then '只一锅记录
        
        Adodc11.Recordset.Fields(1) = ""
        Adodc11.Recordset.Update
      End If
    If zhci = zd And zd > 1 Then '多锅记录且为最后
       
       For i = 1 To zd - 1
       Adodc11.Recordset.Fields(i) = DataCombo2(i).Text
       Next
       Adodc11.Recordset.Fields(zd) = ""
       Adodc11.Recordset.Update
    Else    '多锅记录不在最后
       
       For i = zhci To zd - 1
       Adodc11.Recordset.Fields(i) = DataCombo2(i + 1)
       Next
       Adodc11.Recordset.Fields(zd) = ""
       Adodc11.Recordset.Update
    End If
  End If

Adodc11.RecordSource = "select * from jhb where 车台编号= '" & DataCombo3.Text & "' "
Adodc11.Refresh
If Adodc11.Recordset.EOF Then
     MsgBox ("转到的车台编号错误")
     Exit Sub
  Else '追加
     zd = 0
      For c = 1 To Adodc11.Recordset.Fields.count - 1
      If Adodc11.Recordset.Fields(c) <> "" Then
      zd = zd + 1
      End If
      Next
      
       Adodc11.Recordset.Fields(zd + 1) = zhct
       Adodc11.Recordset.Update
  End If
Adodc2.Refresh
Label1.Visible = False
DataCombo3.Visible = False
DataCombo1.SetFocus
DataCombo1.Text = DataCombo1.Text
Command2.Enabled = False
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 车台编号  from ct group by 车台编号"
Adodc1.Refresh
DataCombo1.Text = ""
DataCombo1.Text = m
zhci = 0
End Sub


Private Sub Command3_Click()
Formj5.DTPicker3 = Date - 10
Formj5.Check2(4).value = 1
Formj5.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Command2.Enabled = True
Label1.Visible = True
DataCombo3.Visible = True
End Sub

Private Sub Command6_Click()
Forma172.Check2(4).value = 0
Forma172.Check2(15).value = 1
Forma172.Show
End Sub

Private Sub Command7_Click()
On Error Resume Next
If DataCombo1.Text = "" Then
MsgBox ("请选择车台:")
Exit Sub
End If
m = DataCombo1.Text
L = InputBox("请输入车台" + DataCombo1.Text + "编号", , "1")
If MsgBox("确实删除吗?", vbYesNo) = vbYes Then

          Adodc11.RecordSource = "select * from jhb where 车台编号='" & DataCombo1.Text & "'"       '''''''判断此车台排锅情况
          Adodc11.Refresh
           If Adodc11.Recordset.EOF Then
                MsgBox ("此车台不存在！！")
                Exit Sub
           Else
               zd = 0
               For i = 1 To 30
               If Adodc11.Recordset.Fields(i) <> "" Then
               zd = zd + 1
               End If
               Next
               If zd = 0 Then
               MsgBox ("此车台并没有按排活！！！")
               Exit Sub
               End If
           End If
               '''''''''''panduan weizhi
               If L > zd Then        ''''''''''''''判断编号
               MsgBox ("编号有误,请重输!")
               Exit Sub
               End If
               
               If L < zd Then
                   For i = L To zd - 1
                       Adodc11.Recordset.Fields(i) = Adodc11.Recordset.Fields(i + 1)
                   Next
                   Adodc11.Recordset.Fields(zd) = ""
               Else
                   Adodc11.Recordset.Fields(zd) = ""
               End If
               Adodc11.Recordset.Update
               
End If
Adodc2.Refresh
DataCombo1.Text = ""
DataCombo1.Text = m
End Sub



Private Sub DataCombo1_Change()
On Error Resume Next
ww = 1
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select * from jhb where 车台编号='" & DataCombo1.Text & "'"
Adodc11.Refresh
  If Adodc11.Recordset.EOF Then
    For i = 1 To 30 '清除内容
    DataCombo2(i).Text = ""
    DataCombo2(i).Visible = False
    Text1(i).Text = ""
    Text1(i).Visible = False
    Next
    ww = 0
    Exit Sub
  Else
    Command1.Enabled = True
    Command5.Enabled = True
    zd = 0
    For c = 1 To Adodc11.Recordset.Fields.count - 1
    If Adodc11.Recordset.Fields(c) <> "" Then
    zd = zd + 1
    End If
    Next

    If zd = 0 Then
      Command1.Enabled = False
      Command5.Enabled = False
      For i = 1 To 30 '清除内容
      DataCombo2(i).Text = ""
      DataCombo2(i).Visible = False
      Text1(i).Text = ""
      Text1(i).Visible = False
      Next
      ww = 0
      Exit Sub
    End If
    If zd = 1 Then
       Command1.Enabled = False
       Command5.Enabled = True
       DataCombo2(1).Text = Adodc11.Recordset.Fields(1)
       Text1(1).Text = 1
       DataCombo2(1).Visible = True
       Text1(1).Visible = True
       For i = 2 To 30 '清除内容
       DataCombo2(i).Text = ""
       DataCombo2(i).Visible = False
       Text1(i).Text = ""
       Text1(i).Visible = False
       Next
       
       ww = 0
       Exit Sub
    End If
      If zd > 1 Then
          Command1.Enabled = True
          Command5.Enabled = True
      Else
          Command1.Enabled = False
          Command5.Enabled = True
      End If
    For i = 1 To zd
    DataCombo2(i).Visible = True
    Text1(i).Visible = True
    Text1(i).Text = i
    DataCombo2(i).Text = Adodc11.Recordset.Fields(i)
    Next
     
     If i >= 30 Then
     Exit Sub
     ww = 0
     End If
     For i = zd + 1 To 30 '清除内容
     DataCombo2(i).Text = ""
     DataCombo2(i).Visible = False
     Text1(i).Text = ""
     Text1(i).Visible = False
     Next
    
    
    
  End If


asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next
ww = 0


L = "备活"
m = "就绪"


Adodc9.RecordSource = "select * from jhb where 车台编号='" & DataCombo1.Text & "' "
Adodc9.Refresh
For i = 1 To 30


If InStr(Trim(Adodc9.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc9.Recordset.Fields(i)), m) <> 0 Then
DataCombo2(i).ForeColor = vbGreen
End If

If InStr(Trim(Adodc9.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc9.Recordset.Fields(i)), m) <> 0 Then
DataCombo2(i).ForeColor = vbRed
End If

If InStr(Trim(Adodc9.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc9.Recordset.Fields(i)), m) = 0 Then
   DataCombo2(i).ForeColor = vbCyan
End If

If InStr(Trim(Adodc9.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc9.Recordset.Fields(i)), m) = 0 Then
   DataCombo2(i).ForeColor = vbBlack
End If

Next


End Sub

Private Sub DataCombo1_Click(Area As Integer)
On Error Resume Next

ww = 1
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select * from jhb where 车台编号='" & DataCombo1.Text & "'"
Adodc11.Refresh
  If Adodc11.Recordset.EOF Then
    For i = 1 To 30 '清除内容
    DataCombo2(i).Text = ""
    DataCombo2(i).Visible = False
    Text1(i).Text = ""
    Text1(i).Visible = False
    Next
    ww = 0
    Exit Sub
  Else
    Command1.Enabled = True
    Command5.Enabled = True
    zd = 0
    For c = 1 To Adodc11.Recordset.Fields.count - 1
    If Adodc11.Recordset.Fields(c) <> "" Then
    zd = zd + 1
    End If
    Next

    If zd = 0 Then
      Command1.Enabled = False
      Command5.Enabled = False
      For i = 1 To 30 '清除内容
      DataCombo2(i).Text = ""
      DataCombo2(i).Visible = False
      Text1(i).Text = ""
      Text1(i).Visible = False
      Next
      ww = 0
      Exit Sub
    End If
    If zd = 1 Then
       Command1.Enabled = False
       Command5.Enabled = True
       DataCombo2(1).Text = Adodc11.Recordset.Fields(1)
       Text1(1).Text = 1
       DataCombo2(1).Visible = True
       Text1(1).Visible = True
       For i = 2 To 30 '清除内容
       DataCombo2(i).Text = ""
       DataCombo2(i).Visible = False
       Text1(i).Text = ""
       Text1(i).Visible = False
       Next
       
       ww = 0
       Exit Sub
    End If
      If zd > 1 Then
          Command1.Enabled = True
          Command5.Enabled = True
      Else
          Command1.Enabled = False
          Command5.Enabled = True
      End If
    For i = 1 To zd
    DataCombo2(i).Visible = True
    Text1(i).Visible = True
    Text1(i).Text = i
    DataCombo2(i).Text = Adodc11.Recordset.Fields(i)
    Next
     
     If i >= 30 Then
     Exit Sub
     ww = 0
     End If
     For i = zd + 1 To 30 '清除内容
     DataCombo2(i).Text = ""
     DataCombo2(i).Visible = False
     Text1(i).Text = ""
     Text1(i).Visible = False
     Next
    
    
    
  End If


asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next
ww = 0




L = "备活"
m = "就绪"


Adodc9.RecordSource = "select * from jhb where 车台编号='" & DataCombo1.Text & "' "
Adodc9.Refresh
For i = 1 To 30

If InStr(Trim(Adodc9.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc9.Recordset.Fields(i)), m) <> 0 Then
DataCombo2(i).ForeColor = vbGreen
End If

If InStr(Trim(Adodc9.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc9.Recordset.Fields(i)), m) <> 0 Then
DataCombo2(i).ForeColor = vbRed
End If

If InStr(Trim(Adodc9.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc9.Recordset.Fields(i)), m) = 0 Then
   DataCombo2(i).ForeColor = vbCyan
End If

If InStr(Trim(Adodc9.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc9.Recordset.Fields(i)), m) = 0 Then
   DataCombo2(i).ForeColor = vbBlack
End If

Next


End Sub

Private Sub DataCombo2_Click(Index As Integer, Area As Integer)
zhci = 0
Select Case Index
Case 1
 zhct = DataCombo2(1).Text
 zhci = 1
Case 2
 zhct = DataCombo2(2).Text
 zhci = 2
Case 3
 zhct = DataCombo2(3).Text
 zhci = 3
Case 4
 zhct = DataCombo2(4).Text
 zhci = 4
Case 5
 zhct = DataCombo2(5).Text
 zhci = 5
Case 6
 zhct = DataCombo2(6).Text
  zhci = 6
Case 7
 zhct = DataCombo2(7).Text
 zhci = 7
Case 8
 zhct = DataCombo2(8).Text
 zhci = 8
Case 9
 zhct = DataCombo2(9).Text
 zhci = 9
Case 10
 zhct = DataCombo2(10).Text
 zhci = 10
Case 11
 zhct = DataCombo2(11).Text
 zhci = 11
Case 12
 zhct = DataCombo2(12).Text
 zhci = 12
Case 13
 zhct = DataCombo2(13).Text
 zhci = 13
Case 14
 zhct = DataCombo2(14).Text
 zhci = 14
Case 15
 zhct = DataCombo2(15).Text
 zhci = 15
Case 16
 zhct = DataCombo2(16).Text
 zhci = 16
 
 Case 17
 zhct = DataCombo2(17).Text
 zhci = 17
Case 18
 zhct = DataCombo2(18).Text
 zhci = 18
Case 19
 zhct = DataCombo2(19).Text
 zhci = 19
Case 20
 zhct = DataCombo2(20).Text
 zhci = 20
Case 21
 zhct = DataCombo2(21).Text
 zhci = 21
Case 22
 zhct = DataCombo2(22).Text
 zhci = 22
Case 23
 zhct = DataCombo2(23).Text
 zhci = 23
Case 24
 zhct = DataCombo2(24).Text
 zhci = 24
Case 25
 zhct = DataCombo2(25).Text
 zhci = 25
Case 26
 zhct = DataCombo2(26).Text
 zhci = 26
Case 27
 zhct = DataCombo2(27).Text
 zhci = 27
Case 28
 zhct = DataCombo2(28).Text
 zhci = 28
Case 29
 zhct = DataCombo2(29).Text
 zhci = 29
Case 30
 zhct = DataCombo2(30).Text
 zhci = 30
 
End Select
End Sub
Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub VSFlexGrid1_dblClick()
rs = VSFlexGrid1.Row
If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    VSFlexGrid1.Text = Text1111.Text
    Text1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move r - 1


Adodc2.Recordset.Fields(c) = Text1111.Text
Adodc2.Recordset.Update

Text1111.Visible = False
VSFlexGrid1.SetFocus
End Sub


Private Sub Form_Load()


'On Error Resume Next
DataCombo1.Text = ""
DataCombo3.Text = ""
For i = 1 To 30
DataCombo2(i) = ""
Next
Dim zu(60) As String
Dim gk As String
ww = 0
lbj = "0"
Option1.value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Label1.Visible = False
DataCombo3.Visible = False


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 车台编号  from ct group by 车台编号"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select *  from JHB ORDER BY 车台编号"
Adodc2.Refresh

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select *  from JHB ORDER BY 车台编号"
Adodc8.Refresh

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 车台编号  from ct group by 车台编号"
Adodc6.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "JHBF"
Adodc7.Refresh

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

If Adodc1.Recordset.EOF Then
MsgBox ("请先设置车台！")
Exit Sub
End If

Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF  '设备数量
Adodc11.RecordSource = "select * from jhb where 车台编号='" & Adodc1.Recordset.Fields(0) & "'"
Adodc11.Refresh
If Adodc11.Recordset.EOF Then  '无记录就添加
Adodc11.Recordset.AddNew
Adodc11.Recordset.Fields(0) = Adodc1.Recordset.Fields(0)
Adodc11.Recordset.Update
End If
Adodc1.Recordset.MoveNext
Loop

 
For i = 1 To 30 '清除内容
DataCombo2(i).Text = ""
DataCombo2(i).Visible = False
Text1(i).Text = ""
Text1(i).Visible = False
Next
DataCombo1.TabIndex = 0
Adodc2.Refresh



VSFlexGrid1.ColWidth(0) = 1000
For i = 1 To 30
VSFlexGrid1.ColWidth(i) = 3000
Next

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 1800
Next
End If


DataCombo1.Text = ""
Command1.Enabled = False
Command2.Enabled = False
Command5.Enabled = False
End Sub





Private Sub Option1_Click()
bh = 0
End Sub

Private Sub Option2_Click()
bh = 1
End Sub



Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 1
       g = 1
       Case 2
       g = 2
       Case 3
       g = 3
       Case 4
       g = 4
       Case 5
       g = 5
       Case 6
       g = 6
       Case 7
       g = 7
       Case 8
       g = 8
       Case 9
       g = 9
       Case 10
       g = 10
       Case 11
       g = 11
       Case 12
       g = 12
       Case 13
       g = 13
       Case 14
       g = 14
       Case 15
       g = 15
       Case 16
       g = 16
       Case 17
       g = 17
       Case 18
       g = 18
       Case 19
       g = 19
       Case 20
       g = 20
       Case 21
       g = 21
       Case 22
       g = 22
       Case 23
       g = 23
       Case 24
       g = 24
       Case 25
       g = 25
       Case 26
       g = 26
       Case 27
       g = 27
       Case 28
       g = 28
       Case 29
       g = 29
       Case 30
       g = 30
End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
On Error Resume Next
L = 0
If bh = 0 Then


If asd > 1 Then 'shi fou duo tiao ji lu
Select Case Index
       Case 1
       If ww = 1 Then Exit Sub
       If Val(Text1(1).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Text1(1).Text = TT(1)
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(1).Text Then
              TR = i
              End If
           Next
           For L = TR To asd - 1
           If L < 1 Then
          
           Text1(L) = L + 1
           End If
           If L > 1 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量

For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
   
       Case 2
       If ww = 1 Then Exit Sub
       If Val(Text1(2).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(2).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 2 Then
          
           Text1(L) = L + 1
           End If
           If L > 2 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
        
       
       Case 3
       If ww = 1 Then Exit Sub
       If Val(Text1(3).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(3).Text Then
              TR = i
              End If
           Next
           For L = TR To asd - 1
           If L < 3 Then
          
           Text1(L) = L + 1
           End If
           If L > 3 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       Case 4
       If ww = 1 Then Exit Sub
       If Val(Text1(4).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(4).Text Then
              TR = i
              End If
           Next
           For L = TR To asd - 1
           If L < 4 Then
          
           Text1(L) = L + 1
           End If
           If L > 4 Then
            Text1(L) = L
           End If
           Next
          
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       Case 5
       If ww = 1 Then Exit Sub
       If Val(Text1(5).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(5).Text Then
              TR = i
              End If
           Next
           For L = TR To asd - 1
           If L < 5 Then
          
           Text1(L) = L + 1
           End If
           If L > 5 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          L = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       
       
       
       
       
       Case 6
       If ww = 1 Then Exit Sub
       If Val(Text1(6).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(6).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 6 Then
          
           Text1(L) = L + 1
           End If
           If L > 6 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       Case 7
       If ww = 1 Then Exit Sub
       If Val(Text1(7).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(7).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 7 Then
          
           Text1(L) = L + 1
           End If
           If L > 7 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       Case 8
       If ww = 1 Then Exit Sub
       If Val(Text1(8).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(8).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 8 Then
          
           Text1(L) = L + 1
           End If
           If L > 8 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       
       Case 9
       If ww = 1 Then Exit Sub
       If Val(Text1(9).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(9).Text Then
              TR = i
              End If
           Next
           For L = TR To asd - 1
           If L < 9 Then
          
           Text1(L) = L + 1
           End If
           If L > 9 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       Case 10
       If ww = 1 Then Exit Sub
       If Val(Text1(10).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(10).Text Then
              TR = i
              End If
           Next
           For L = TR To asd - 1
           If L < 10 Then
          
           Text1(L) = L + 1
           End If
           If L > 10 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       Case 11
       If ww = 1 Then Exit Sub
       If Val(Text1(11).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(11).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 11 Then
          
           Text1(L) = L + 1
           End If
           If L > 11 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
        
        
        
        
       Case 12
       If ww = 1 Then Exit Sub
       If Val(Text1(12).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(12).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 12 Then
          
           Text1(L) = L + 1
           End If
           If L > 12 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
        
        
        
        
       Case 13
       If ww = 1 Then Exit Sub
       If Val(Text1(13).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(13).Text Then
              TR = i
              End If
           Next
           For L = TR To asd - 1
           If L < 13 Then
          
           Text1(L) = L + 1
           End If
           If L > 13 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       
       
       Case 14
       If ww = 1 Then Exit Sub
       If Val(Text1(14).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(14).Text Then
              TR = i
              End If
           Next
           For L = TR To asd - 1
           If L < 14 Then
          
           Text1(L) = L + 1
           End If
           If L > 14 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
 
 
        Case 15
       If ww = 1 Then Exit Sub
       If Val(Text1(15).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(15).Text Then
              TR = i
              End If
           Next
           For L = TR To asd - 1
           If L < 15 Then
          
           Text1(L) = L + 1
           End If
           If L > 15 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 16
       If ww = 1 Then Exit Sub
       If Val(Text1(16).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(16).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 16 Then
          
           Text1(L) = L + 1
           End If
           If L > 16 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 17
       If ww = 1 Then Exit Sub
       If Val(Text1(17).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(17).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 17 Then
          
           Text1(L) = L + 1
           End If
           If L > 17 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
 

 
        Case 18
       If ww = 1 Then Exit Sub
       If Val(Text1(18).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(18).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 18 Then
          
           Text1(L) = L + 1
           End If
           If L > 18 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 19
       If ww = 1 Then Exit Sub
       If Val(Text1(19).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(19).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 19 Then
          
           Text1(L) = L + 1
           End If
           If L > 19 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 20
       If ww = 1 Then Exit Sub
       If Val(Text1(20).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(20).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 20 Then
          
           Text1(L) = L + 1
           End If
           If L > 20 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 21
       If ww = 1 Then Exit Sub
       If Val(Text1(21).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(21).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 21 Then
          
           Text1(L) = L + 1
           End If
           If L > 21 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 22
       If ww = 1 Then Exit Sub
       If Val(Text1(22).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(22).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 22 Then
          
           Text1(L) = L + 1
           End If
           If L > 22 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 23
       If ww = 1 Then Exit Sub
       If Val(Text1(23).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(23).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 23 Then
          
           Text1(L) = L + 1
           End If
           If L > 23 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 24
       If ww = 1 Then Exit Sub
       If Val(Text1(24).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(24).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 24 Then
          
           Text1(L) = L + 1
           End If
           If L > 24 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 25
       If ww = 1 Then Exit Sub
       If Val(Text1(25).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(25).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 25 Then
          
           Text1(L) = L + 1
           End If
           If L > 25 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 26
       If ww = 1 Then Exit Sub
       If Val(Text1(26).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(26).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 26 Then
          
           Text1(L) = L + 1
           End If
           If L > 26 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 27
       If ww = 1 Then Exit Sub
       If Val(Text1(27).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(27).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 27 Then
          
           Text1(L) = L + 1
           End If
           If L > 27 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 28
       If ww = 1 Then Exit Sub
       If Val(Text1(28).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(28).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 28 Then
          
           Text1(L) = L + 1
           End If
           If L > 28 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 29
       If ww = 1 Then Exit Sub
       If Val(Text1(29).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(29).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 29 Then
          
           Text1(L) = L + 1
           End If
           If L > 29 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 30
       If ww = 1 Then Exit Sub
       If Val(Text1(30).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(30).Text Then
              TR = i
              End If
           Next

           For L = TR To asd - 1
           If L < 30 Then
          
           Text1(L) = L + 1
           End If
           If L > 30 Then
            Text1(L) = L
           End If
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 

 
 
 
 
 End Select
     
 End If
 
 
 
 
 
 
 Else
 L = 0
 
 
 If asd > 1 Then 'shi fou duo tiao ji lu
Select Case Index
       Case 1
       If ww = 1 Then Exit Sub
       If Val(Text1(1).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Text1(1).Text = TT(1)
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(1).Text Then
              TR = i
              End If
           Next
           For L = 1 To Val(Text1(1).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量

For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
   
       Case 2
       If ww = 1 Then Exit Sub
       If Val(Text1(2).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(2).Text Then
              TR = i
              End If
           Next

           For L = 2 To Val(Text1(2).Text) - 1
           Text1(L + 1).Text = L
           Next
           
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
        
       
       Case 3
       If ww = 1 Then Exit Sub
       If Val(Text1(3).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(3).Text Then
              TR = i
              End If
           Next
           
           For L = 3 To Val(Text1(3).Text) - 1
           Text1(L + 1).Text = L
           Next
           
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       Case 4
       If ww = 1 Then Exit Sub
       If Val(Text1(4).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(4).Text Then
              TR = i
              End If
           Next
           For L = 4 To Val(Text1(4).Text) - 1
           Text1(L + 1).Text = L
           Next
          
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       Case 5
       If ww = 1 Then Exit Sub
       If Val(Text1(5).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(5).Text Then
              TR = i
              End If
           Next
           For L = 5 To Val(Text1(5).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          L = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       
       
       
       
       
       Case 6
       If ww = 1 Then Exit Sub
       If Val(Text1(6).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(6).Text Then
              TR = i
              End If
           Next

           For L = 6 To Val(Text1(6).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       Case 7
       If ww = 1 Then Exit Sub
       If Val(Text1(7).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(7).Text Then
              TR = i
              End If
           Next

           For L = 7 To Val(Text1(7).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       Case 8
       If ww = 1 Then Exit Sub
       If Val(Text1(8).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(8).Text Then
              TR = i
              End If
           Next

           For L = 8 To Val(Text1(8).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       
       Case 9
       If ww = 1 Then Exit Sub
       If Val(Text1(9).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(9).Text Then
              TR = i
              End If
           Next
           For L = 9 To Val(Text1(9).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       Case 10
       If ww = 1 Then Exit Sub
       If Val(Text1(10).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(10).Text Then
              TR = i
              End If
           Next
           For L = 10 To Val(Text1(10).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       Case 11
       If ww = 1 Then Exit Sub
       If Val(Text1(11).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(11).Text Then
              TR = i
              End If
           Next

           For L = 11 To Val(Text1(11).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
        
        
        
        
       Case 12
       If ww = 1 Then Exit Sub
       If Val(Text1(12).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(12).Text Then
              TR = i
              End If
           Next

           For L = 12 To Val(Text1(12).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
        
        
        
        
       Case 13
       If ww = 1 Then Exit Sub
       If Val(Text1(13).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(13).Text Then
              TR = i
              End If
           Next
           For L = 13 To Val(Text1(13).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
       
       
       
       
       
       Case 14
       If ww = 1 Then Exit Sub
       If Val(Text1(14).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(14).Text Then
              TR = i
              End If
           Next
           For L = 14 To Val(Text1(14).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
 
 
        Case 15
       If ww = 1 Then Exit Sub
       If Val(Text1(15).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(15).Text Then
              TR = i
              End If
           Next
           For L = 15 To Val(Text1(15).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 16
       If ww = 1 Then Exit Sub
       If Val(Text1(16).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(16).Text Then
              TR = i
              End If
           Next

           For L = 16 To Val(Text1(16).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 17
       If ww = 1 Then Exit Sub
       If Val(Text1(17).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(17).Text Then
              TR = i
              End If
           Next

           For L = 17 To Val(Text1(17).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If
 

 
        Case 18
       If ww = 1 Then Exit Sub
       If Val(Text1(18).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(18).Text Then
              TR = i
              End If
           Next

           For L = 18 To Val(Text1(18).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 19
       If ww = 1 Then Exit Sub
       If Val(Text1(19).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(19).Text Then
              TR = i
              End If
           Next

           For L = 19 To Val(Text1(19).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 20
       If ww = 1 Then Exit Sub
       If Val(Text1(20).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(20).Text Then
              TR = i
              End If
           Next

           For L = 20 To Val(Text1(20).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 21
       If ww = 1 Then Exit Sub
       If Val(Text1(21).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(21).Text Then
              TR = i
              End If
           Next

           For L = 21 To Val(Text1(21).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 22
       If ww = 1 Then Exit Sub
       If Val(Text1(22).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(22).Text Then
              TR = i
              End If
           Next

           For L = 22 To Val(Text1(22).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 23
       If ww = 1 Then Exit Sub
       If Val(Text1(23).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(23).Text Then
              TR = i
              End If
           Next

           For L = 23 To Val(Text1(23).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 24
       If ww = 1 Then Exit Sub
       If Val(Text1(24).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(24).Text Then
              TR = i
              End If
           Next

           For L = 24 To Val(Text1(24).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 25
       If ww = 1 Then Exit Sub
       If Val(Text1(25).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(25).Text Then
              TR = i
              End If
           Next

           For L = 25 To Val(Text1(25).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 26
       If ww = 1 Then Exit Sub
       If Val(Text1(26).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(26).Text Then
              TR = i
              End If
           Next

           For L = 26 To Val(Text1(26).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 27
       If ww = 1 Then Exit Sub
       If Val(Text1(27).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(27).Text Then
              TR = i
              End If
           Next

           For L = 27 To Val(Text1(27).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 28
       If ww = 1 Then Exit Sub
       If Val(Text1(28).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(28).Text Then
              TR = i
              End If
           Next

           For L = 28 To Val(Text1(28).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 29
       If ww = 1 Then Exit Sub
       If Val(Text1(29).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(29).Text Then
              TR = i
              End If
           Next

           For L = 29 To Val(Text1(29).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 
        Case 30
       If ww = 1 Then Exit Sub
       If Val(Text1(30).Text) > asd Then '编号大于范围退出
          MsgBox ("wrong")
          Exit Sub
          Else
         TR = 0
           For i = 1 To asd   '查找同编号要自动编号的记录
              If TT(i) = Text1(30).Text Then
              TR = i
              End If
           Next

           For L = 30 To Val(Text1(30).Text) - 1
           Text1(L + 1).Text = L
           Next
          ww = 0
          asd = 0 '判断不为空白的数量
For i = 1 To 30
If Text1(i).Text <> "" Then
asd = asd + 1
TT(i) = Text1(i).Text
End If
Next

          Exit Sub
        End If

 

 
 
 
 
 End Select
     
 End If

 
 
 End If

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Adodc8.Refresh
p = 1
L = "备活"
m = "就绪"
Adodc8.Recordset.MoveFirst
Do While Not Adodc8.Recordset.EOF
For i = 1 To 30

If InStr(Trim(Adodc8.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc8.Recordset.Fields(i)), m) <> 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i
    VSFlexGrid1.CellForeColor = vbGreen
End If

If InStr(Trim(Adodc8.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc8.Recordset.Fields(i)), m) <> 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i
    VSFlexGrid1.CellForeColor = vbRed
End If

If InStr(Trim(Adodc8.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc8.Recordset.Fields(i)), m) = 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i
    VSFlexGrid1.CellForeColor = vbCyan
End If

If InStr(Trim(Adodc8.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc8.Recordset.Fields(i)), m) = 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i
    VSFlexGrid1.CellForeColor = vbBlack
End If

Next
Adodc8.Recordset.MoveNext
p = p + 1
Loop

End Sub

Private Sub jc(m As String, L As String)

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select * from jhb where 车台编号='" & m & "'"
Adodc11.Refresh
  If Adodc11.Recordset.EOF Then
                MsgBox ("此车台不存在！！")
                Exit Sub
           Else
               zd = 0
               For i = 1 To 30
               If Adodc11.Recordset.Fields(i) <> "" Then
               zd = zd + 1
               End If
               Next
               If zd = 0 Then
              ' MsgBox ("此车台并没有按排活！！！")
               Exit Sub
               End If
           End If
               '''''''''''panduan weizhi
               If L > zd Then        ''''''''''''''判断编号
              ' MsgBox ("编号有误,请重输!")
               Exit Sub
               End If
               
               If L < zd Then
                   For i = L To zd - 1
                       Adodc11.Recordset.Fields(i) = Adodc11.Recordset.Fields(i + 1)
                   Next
                   Adodc11.Recordset.Fields(zd) = ""
               Else
                   Adodc11.Recordset.Fields(zd) = ""
               End If
               Adodc11.Recordset.Update
Adodc2.Refresh
              
End Sub

