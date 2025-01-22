VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formy121 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料入库"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form21"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1200
      TabIndex        =   72
      Text            =   "Text3"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "条码"
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   5400
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc20 
      Height          =   375
      Left            =   7320
      Top             =   10560
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
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
      Left            =   7320
      Top             =   10680
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
      Left            =   7200
      Top             =   10680
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
      Height          =   330
      Left            =   7200
      Top             =   10680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Height          =   330
      Left            =   7440
      Top             =   10680
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
      Height          =   330
      Left            =   7440
      Top             =   10560
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
      Left            =   9600
      Top             =   10440
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
      Height          =   375
      Left            =   7320
      Top             =   10560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Left            =   7800
      Top             =   10800
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
      Left            =   7800
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
      Height          =   330
      Left            =   7800
      Top             =   10680
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
      Left            =   7440
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
      Left            =   7680
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
      Left            =   8040
      Top             =   10440
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
      Height          =   375
      Left            =   7920
      Top             =   10680
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
      Left            =   8160
      Top             =   10440
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   7560
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
      Left            =   8400
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
      Left            =   8160
      Top             =   10440
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
      Left            =   8640
      Top             =   10440
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formy121.frx":0000
      Height          =   5295
      Left            =   360
      TabIndex        =   62
      Top             =   6120
      Width           =   15255
      _cx             =   26908
      _cy             =   9340
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Text2"
      Top             =   11520
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1335
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1335
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
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1455
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
      ItemData        =   "Formy121.frx":0015
      Left            =   11880
      List            =   "Formy121.frx":001F
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4320
      TabIndex        =   29
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   328007681
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8760
      TabIndex        =   30
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330760193
      CurrentDate     =   39177
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "客户信息"
      Height          =   3255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   2040
      Width           =   15135
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   840
         TabIndex        =   71
         Text            =   "Text3"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   615
         Left            =   11040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Text            =   "Formy121.frx":002B
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   14160
         Top             =   1440
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   4320
         TabIndex        =   12
         Top             =   3000
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formy121.frx":0031
         Height          =   330
         Index           =   22
         Left            =   1440
         TabIndex        =   42
         Top             =   2160
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "简称"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formy121.frx":0046
         Height          =   330
         Index           =   3
         Left            =   1440
         TabIndex        =   43
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "材料名称"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formy121.frx":005B
         Height          =   330
         Index           =   4
         Left            =   5400
         TabIndex        =   44
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "材料规格"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formy121.frx":0071
         Height          =   330
         Index           =   5
         Left            =   5400
         TabIndex        =   45
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "MC"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   6
         Left            =   5400
         TabIndex        =   46
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   7
         Left            =   8400
         TabIndex        =   47
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   8
         Left            =   5400
         TabIndex        =   48
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   9
         Left            =   5400
         TabIndex        =   49
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   10
         Left            =   8400
         TabIndex        =   50
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   11
         Left            =   5400
         TabIndex        =   51
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   12
         Left            =   8400
         TabIndex        =   52
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formy121.frx":0086
         Height          =   330
         Index           =   13
         Left            =   8400
         TabIndex        =   53
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "mc"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formy121.frx":009C
         Height          =   330
         Index           =   14
         Left            =   8400
         TabIndex        =   54
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "仓位"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formy121.frx":00B1
         Height          =   330
         Index           =   15
         Left            =   1440
         TabIndex        =   55
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "MC"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formy121.frx":00C6
         Height          =   330
         Index           =   16
         Left            =   1440
         TabIndex        =   56
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "MC"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   17
         Left            =   11040
         TabIndex        =   57
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formy121.frx":00DC
         Height          =   330
         Index           =   18
         Left            =   8400
         TabIndex        =   58
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "XM"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   19
         Left            =   12840
         TabIndex        =   59
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   20
         Left            =   11400
         TabIndex        =   60
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   21
         Left            =   11400
         TabIndex        =   61
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   0
         Left            =   13200
         TabIndex        =   63
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   1
         Left            =   1440
         TabIndex        =   64
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   2
         Left            =   1440
         TabIndex        =   65
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "税率"
         Height          =   375
         Index           =   6
         Left            =   10440
         TabIndex        =   69
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "款号"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   68
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单号"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   67
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "颜色"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   66
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "名称"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "付款"
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   38
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "发票"
         Height          =   375
         Index           =   2
         Left            =   7440
         TabIndex        =   28
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单位"
         Height          =   375
         Index           =   4
         Left            =   4200
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "仓位"
         Height          =   375
         Index           =   7
         Left            =   7440
         TabIndex        =   26
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "规格"
         Height          =   375
         Index           =   6
         Left            =   4200
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "包件"
         Height          =   375
         Index           =   9
         Left            =   4200
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "合计金额"
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "库别"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "批号"
         Height          =   375
         Index           =   10
         Left            =   7440
         TabIndex        =   21
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单价"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   20
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "数量"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   19
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "备注"
         Height          =   375
         Index           =   11
         Left            =   10440
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "仓务员"
         Height          =   375
         Index           =   12
         Left            =   7440
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "库类"
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "日期"
         Height          =   375
         Index           =   14
         Left            =   10440
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "序号"
         Height          =   375
         Index           =   15
         Left            =   10440
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "供应商"
         Height          =   375
         Index           =   16
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   615
      End
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   23
      Left            =   360
      TabIndex        =   41
      Top             =   1560
      Width           =   2300
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 染 整 材 料 入 库 "
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
      Left            =   4680
      TabIndex        =   36
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   35
      Top             =   1320
      Width           =   855
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   6240
      X2              =   7440
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   34
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "合计金额"
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
      Left            =   10200
      TabIndex        =   33
      Top             =   11520
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单据号"
      Height          =   380
      Index           =   17
      Left            =   360
      TabIndex        =   32
      Top             =   1200
      Width           =   2300
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
      Left            =   10920
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Formy121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public x, BAR As Integer
Dim BA As Database: Dim rr As Integer
Dim RD1 As Recordset
Dim a As String  '中间变量
Dim b As Double
Dim c, r As Integer
Dim kg As Integer
Dim bb As Long
Dim cc As String
Dim kkf As Integer
Dim n As Integer
Dim DH As Integer
Dim fh As String
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Private Sub Command12_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "SELECT * FROM clrkdj where 单据编号='" & yhdm & "'"
Adodc13.Refresh

DataCombo1(23).Text = Trim(yhdm) + "0000001"
If Adodc13.Recordset.EOF Then
DataCombo1(23).Text = Trim(yhdm) + "0000001"
Else
uu = Val(Adodc13.Recordset.Fields(1)) + 1
DataCombo1(23).Text = Trim(yhdm) + Left("0000000", 7 - Len(uu)) + Trim(uu)
End If

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select * from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc5.Refresh
Adodc7.RecordSource = "select   MAX(序号) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'"
Adodc7.Refresh
DataCombo1(21).Text = 1
DataCombo1(21).Text = Adodc7.Recordset.Fields(0) + 1
Adodc20.RecordSource = "select round(SUM(合计金额),2) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
Text2.Text = 0
Else
Text2.Text = Adodc20.Recordset.Fields(0)
End If
DataCombo1(20) = Date
DataCombo1(23).Enabled = False
End Sub

Private Sub Command33_Click()
Unload Me
End Sub

Private Sub Command7_Click()
clshxg = 1
Formy146.Check2(0).value = 1
Formy146.Show
End Sub

Private Sub Command11_Click()
On Error Resume Next

If DataCombo1(22).Text = "" Then
MsgBox ("请选择供应商！")
Exit Sub
End If

If DataCombo1(13).Text <> "是" And DataCombo1(13).Text <> "否" Then
MsgBox ("请选择是否付款")
Exit Sub
End If

If Len(DataCombo1(23).Text) <> 8 Then
MsgBox ("单据号编码不符合规则  需要8位")
Exit Sub
End If

If DataCombo1(8).Text = "" Or DataCombo1(9).Text = "" Then
MsgBox ("数量、单价不能为空！！")
Exit Sub
End If

Adodc5.Recordset.AddNew
DataCombo1(19).Text = Text1.Text
For i = 0 To Adodc5.Recordset.Fields.count - 1
Adodc5.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc5.Recordset.Fields(24) = "未"
Adodc5.Recordset.Fields(25) = "未"
Adodc5.Recordset.Update

For i = 3 To Adodc5.Recordset.Fields.count - 7
If i = 13 Then i = 14
If i = 15 Then i = 16
If i = 17 Then i = 18
If i = 18 Then
CWY = DataCombo1(i).Text
End If
If i = 20 Then i = 21
If i = 22 Then i = 24
DataCombo1(i).Text = ""
Next

Adodc5.RecordSource = "select   * from clgl WHERE 单据号='" & DataCombo1(23).Text & "' order by 序号"
Adodc5.Refresh
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, , 9, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, , 11, , vbGreen


Adodc20.RecordSource = "select ROUND(SUM(合计金额),2) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
Text2.Text = 0
Else
Text2.Text = Adodc20.Recordset.Fields(0)
End If

Adodc7.RecordSource = "select MAX(序号) from clgl "
Adodc7.Refresh

DataCombo1(11).Text = ""
DataCombo1(16).Text = "采购入库"
DataCombo1(17).Text = 17
DataCombo1(18).Text = CWY
DataCombo1(21).Text = Adodc7.Recordset.Fields(0) + 1
DataCombo1(0).SetFocus
End Sub

Private Sub Command2_Click()
On Error Resume Next


If DataCombo1(22).Text = "" Then
MsgBox ("请选择供应商！")
Exit Sub
End If

If DataCombo1(13).Text <> "是" And DataCombo1(13).Text <> "否" Then
MsgBox ("请选择是否付款")
Exit Sub
End If

If Len(DataCombo1(23).Text) <> 8 Then
MsgBox ("单据号编码不符合规则  需要8位")
Exit Sub
End If

If DataCombo1(8).Text = "" Or DataCombo1(9).Text = "" Then
MsgBox ("数量、单价不能为空！！")
Exit Sub
End If
If Adodc5.Recordset.Fields(24) = "已" Then Exit Sub
DataCombo1(19).Text = Text1.Text
For i = 0 To Adodc5.Recordset.Fields.count - 1
Adodc5.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc5.Recordset.Update
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, , 11, , vbGreen


For i = 3 To Adodc5.Recordset.Fields.count - 7
If i = 13 Then i = 14
If i = 15 Then i = 16
If i = 17 Then i = 18
If i = 18 Then
CWY = DataCombo1(i).Text
End If
If i = 20 Then i = 21
DataCombo1(i).Text = ""
Next

Adodc5.RecordSource = "select   * from clgl WHERE 单据号='" & DataCombo1(23).Text & "' order by 序号"
Adodc5.Refresh


Adodc7.RecordSource = "select MAX(序号) from clgl "
Adodc7.Refresh
Adodc20.RecordSource = "select ROUND(SUM(合计金额),2) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
Text2.Text = 0
Else
Text2.Text = Adodc20.Recordset.Fields(0)
End If

DataCombo1(11).Text = ""
DataCombo1(16).Text = "采购入库"
DataCombo1(17).Text = 17
DataCombo1(18).Text = CWY
'DataCombo1(20).Text = Date
DataCombo1(21).Text = Adodc7.Recordset.Fields(0) + 1
DataCombo1(0).SetFocus
Command2.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Command4_Click()
On Error Resume Next

If Adodc5.Recordset.Fields(24) = "已" Then Exit Sub
Adodc5.Recordset.Delete

For i = 3 To Adodc5.Recordset.Fields.count - 7
If i = 15 Then i = 16
If i = 18 Then
CWY = DataCombo1(i).Text
End If
If i = 20 Then i = 21
DataCombo1(i).Text = ""
Next


Adodc5.RecordSource = "select   * from clgl WHERE 单据号='" & DataCombo1(23).Text & "' order by 序号"
Adodc5.Refresh
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, , 11, , vbGreen
Adodc7.RecordSource = "select MAX(序号) from clgl "
Adodc7.Refresh
Adodc20.RecordSource = "select ROUND(SUM(合计金额),2) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
Text2.Text = 0
Else
Text2.Text = Adodc20.Recordset.Fields(0)
End If

DataCombo1(11).Text = ""
DataCombo1(16).Text = "采购入库"
DataCombo1(17).Text = 17
DataCombo1(18).Text = CWY
'DataCombo1(20).Text = Date
DataCombo1(21).Text = Adodc7.Recordset.Fields(0) + 1
DataCombo1(0).SetFocus
Command2.Enabled = False
Command4.Enabled = False
Command1.Enabled = True

End Sub


Private Sub Command8_Click()
On Error Resume Next
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
Adodc3.Refresh
Adodc4.Refresh
Adodc6.Refresh
Adodc8.Refresh
Adodc9.Refresh
Adodc14.Refresh

Adodc7.RecordSource = "select   MAX(序号) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'"
Adodc7.Refresh
DataCombo1(21).Text = 1
DataCombo1(21).Text = Adodc7.Recordset.Fields(0) + 1
Adodc5.RecordSource = "select   * from clgl WHERE 单据号='" & DataCombo1(23).Text & "' order by 序号"
Adodc5.Refresh
Adodc20.RecordSource = "select ROUND(SUM(合计金额),2) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
Text2.Text = 0
Else
Text2.Text = Adodc20.Recordset.Fields(0)
End If
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, , 9, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, , 11, , vbGreen

End Sub

Private Sub Command6_Click()
If Adodc5.Recordset.EOF Then
MsgBox ("此单据号中无记录，不能打印！")
Exit Sub
End If
BAR = 1
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub


Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
Case 3
Case 4
DataCombo1(8).Text = Format(Val(DataCombo1(4).Text) * Val(DataCombo1(11).Text), "#0.00")


Case 8
DataCombo1(10).Text = Format(Val(DataCombo1(8).Text) * Val(DataCombo1(9).Text), "#0.00")

Case 9
DataCombo1(10).Text = Format(Val(DataCombo1(8).Text) * Val(DataCombo1(9).Text), "#0.00")
Case 11
DataCombo1(8).Text = Format(Val(DataCombo1(4).Text) * Val(DataCombo1(11).Text), "#0.00")
Case 15
If DataCombo1(15).Text = "染料" Then
DataCombo1(4).Text = "25"
Else
DataCombo1(4).Text = ""
End If
Case 23

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select * from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc5.Refresh
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, , 11, , vbGreen

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select   MAX(序号) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'"
Adodc7.Refresh
DataCombo1(21).Text = 1
DataCombo1(21).Text = Adodc7.Recordset.Fields(0) + 1

Adodc20.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc20.RecordSource = "select ROUND(SUM(合计金额),2) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
Text2.Text = 0
Else
Text2.Text = Adodc20.Recordset.Fields(0)
End If
End Select

End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
Case 1

Case 8
DataCombo1(10).Text = Format(Val(DataCombo1(8).Text) * Val(DataCombo1(9).Text), "#0.00")

Case 9
DataCombo1(10).Text = Format(Val(DataCombo1(8).Text) * Val(DataCombo1(9).Text), "#0.00")

Case 23
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select * from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc5.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select MAX(序号) from clgl "
Adodc7.Refresh
DataCombo1(21).Text = 1
DataCombo1(21).Text = Adodc7.Recordset.Fields(0) + 1
End Select
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub dataCombo1_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
       Case 3
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "select 单价 from clgl where 材料名称='" & DataCombo1(3) & "' order by 日期 desc"
Adodc10.Refresh
If Adodc10.Recordset.EOF Then
DataCombo1(9) = ""
Else
DataCombo1(9) = Adodc10.Recordset.Fields(0)
End If
End Select
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

Private Sub Form_Load()

On Error Resume Next
Text2.Text = ""
Combo1.Text = ""
Text1.Text = ""
Text3.Text = ""
Text6.Text = ""

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "SELECT * FROM clrkdj where 单据编号='" & yhdm & "'"
Adodc13.Refresh

DataCombo1(23).Text = Trim(yhdm) + "0000001"
If Adodc13.Recordset.EOF Then
DataCombo1(23).Text = Trim(yhdm) + "0000001"
Else
uu = Val(Adodc13.Recordset.Fields(1)) + 1
DataCombo1(23).Text = Trim(yhdm) + Left("0000000", 7 - Len(uu)) + Trim(uu)
End If

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select * from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc5.Refresh
cdbhf = cdbh
For i = 0 To Adodc5.Recordset.Fields.count - 1
If i = 23 Then i = 24
DataCombo1(i).Text = ""
Next
DataCombo1(17).Text = 17
DataCombo1(18).Text = ""
DataCombo1(13).Text = "否"
Text4.Text = Date
Text5.Text = Date
DTPicker1.value = Date
DTPicker2.value = Date
DataCombo1(20).Text = Date
cdbhf = cdbh
DataCombo1(23).Enabled = False

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 材料名称 from  clmc  group by 材料名称"
Adodc2.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select distinct 仓位 from  clgl"
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select xm  from CLfzr group by xm"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select * from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc5.Refresh

Adodc20.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc20.RecordSource = "select * from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
Text2.Text = 0
Else
Adodc20.RecordSource = "select round(SUM(合计金额),2) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'  "
Adodc20.Refresh
Text2.Text = Adodc20.Recordset.Fields(0)
End If

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select MC from CLKL where yh='" & yhm & "'  group by MC"
Adodc6.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select MAX(序号) from clgl WHERE 单据号='" & DataCombo1(23).Text & "'"
Adodc7.Refresh

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 简称 from GYS where ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc8.Refresh

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc19.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc19.RecordSource = "select MC from FK group by MC"
Adodc19.Refresh

Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.RecordSource = "select MC from CLKB group by MC"
Adodc14.Refresh

ProgressBar1.Visible = False
Timer1.Enabled = False


Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "select MC from CLDW group by MC"
Adodc9.Refresh

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(4) = 1500
VSFlexGrid1.ColWidth(5) = 1000
VSFlexGrid1.ColWidth(8) = 1200
VSFlexGrid1.ColWidth(9) = 1200
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(15) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(25) = 0
VSFlexGrid1.ColWidth(26) = 0
VSFlexGrid1.ColWidth(27) = 0
VSFlexGrid1.ColWidth(28) = 0
VSFlexGrid1.ColWidth(29) = 0
DataCombo1(16).Text = "采购入库"
DataCombo1(17).Text = 17
DataCombo1(21).Text = 1
DataCombo1(21).Text = Adodc7.Recordset.Fields(0) + 1
DataCombo1(20).Text = Date
DataCombo1(0).TabIndex = 0

Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

If Len(yhdm) <> 1 Then
MsgBox ("这个账户不合适进入这个界面")
Command1.Enabled = False
Command8.Enabled = False
Command11.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Label3(17).Enabled = False
End If

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

Private Sub Label3_Click(Index As Integer)
Select Case Index
       Case 17
DataCombo1(23).Enabled = False
End Select
End Sub

Private Sub Label3_DblClick(Index As Integer)
Select Case Index
       Case 17
DataCombo1(23).Enabled = True
       Case 3
Formy122.Text1 = DataCombo1(1).Text
Formy122.Show
End Select
End Sub

Private Sub Label9_Click()
clbl = 1
Formy58.Text3.Text = DataCombo1(15).Text
Formy58.Text2.Text = DataCombo1(3).Text
Formy58.Show
End Sub

Private Sub Text3_Change()
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 简称 from v_GYS where 简码 like '%'+'" & Text3 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'"
Adodc8.Refresh
End Sub

Private Sub Text6_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from v_clmc where 简码 LIKE '%'+'" & Text6.Text & "'+'%' and  库类='" & DataCombo1(15) & "' and 供应单位 like '%'+'" & DataCombo1(22) & "'+'%' "
Adodc2.Refresh
End Sub

'Private Sub Label5_dblClick()
'If dataCombo1(1).text = "" Then Exit Sub
'Form26.dataCombo1.text = dataCombo1(1).text
'Form26.Show
'End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move rs - 1
For i = 0 To Adodc5.Recordset.Fields.count - 4
DataCombo1(i).Text = Adodc5.Recordset.Fields(i)
Next
Command11.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub

'Private Sub Timer1_Timer()
'If BAR = 100 Then
'adodcEnvironment1.Command1 dataCombo1(23).text
'adodcReport17.Show 1
'adodcEnvironment1.rsCommand1.Close
'Timer1.Enabled = False
'ProgressBar1.Visible = False
'Exit Sub
'End If
'BAR = BAR + 1
'ProgressBar1.Value = BAR

'End Sub


Private Sub MSFlex_DBLClick()
With VSFlexGrid1
    c = .col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex_DBLClick
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
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move r - 1

Adodc5.Recordset.Fields(c - 1) = Text1111.Text
Adodc5.Recordset.Update
Text1111.Visible = False
End Sub




Private Sub Timer1_Timer()
If BAR = 100 Then
Call clrk(Adodc20, DataCombo1(23).Text)
Timer1.Enabled = False
ProgressBar1.Visible = False
BAR = 1
Else
ProgressBar1.value = BAR
BAR = BAR + 1
End If
End Sub
