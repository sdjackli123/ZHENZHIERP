VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forms508 
   BackColor       =   &H00C0E0FF&
   Caption         =   "装卸统计"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Frame1"
      Height          =   495
      Left            =   360
      TabIndex        =   62
      Top             =   120
      Width           =   4575
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF80&
         Caption         =   "按重量核算"
         Height          =   375
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Width           =   2415
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFF80&
         Caption         =   "按匹数核算"
         Height          =   375
         Left            =   2520
         TabIndex        =   63
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   840
      TabIndex        =   61
      Text            =   "Text2"
      Top             =   3840
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc15 
      Height          =   330
      Left            =   6720
      Top             =   10080
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
      Left            =   7080
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
      Height          =   330
      Left            =   7200
      Top             =   10080
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
      Left            =   7440
      Top             =   10200
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
      Height          =   375
      Left            =   6840
      Top             =   10200
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Height          =   495
      Left            =   7560
      Top             =   10080
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Left            =   8280
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
      Height          =   375
      Left            =   7920
      Top             =   10080
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
      Left            =   7440
      Top             =   10200
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
      Left            =   7800
      Top             =   9960
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
      Left            =   7920
      Top             =   10080
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Left            =   7920
      Top             =   10200
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
      Left            =   7680
      Top             =   10200
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
      Top             =   10080
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
      Left            =   8040
      Top             =   9960
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
      Bindings        =   "Forms508.frx":0000
      Height          =   5535
      Left            =   3600
      TabIndex        =   58
      Top             =   3720
      Width           =   11295
      _cx             =   19923
      _cy             =   9763
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
      AllowUserResizing=   4
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
      WordWrap        =   0   'False
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
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "毛坯出库"
      Height          =   375
      Left            =   2640
      TabIndex        =   57
      Top             =   4560
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "光坯出库"
      Height          =   495
      Left            =   2640
      TabIndex        =   55
      Top             =   5160
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "毛坯入库"
      Height          =   495
      Left            =   2640
      TabIndex        =   54
      Top             =   3840
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   5100
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   50
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
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
      Height          =   495
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
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
      Height          =   495
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   9
      Left            =   7680
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Forms508.frx":0015
      Left            =   840
      List            =   "Forms508.frx":0022
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   495
      Index           =   18
      Left            =   7080
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   495
      Index           =   17
      Left            =   7680
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   16
      Left            =   10800
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
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
      Left            =   11280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
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
      Height          =   495
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
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
      Height          =   495
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command5 
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
      Height          =   495
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   15
      Left            =   7680
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   495
      Index           =   14
      Left            =   11400
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   -120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   13
      Left            =   7680
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   495
      Index           =   12
      Left            =   10800
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   840
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   10
      Left            =   4560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   7
      Left            =   840
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   840
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   10800
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   4
      Left            =   4560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   3
      Left            =   4560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   1
      Left            =   4560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   0
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   8
      Left            =   4560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   12960
      TabIndex        =   26
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   327942145
      CurrentDate     =   40055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   840
      TabIndex        =   27
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   327942145
      CurrentDate     =   40055
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   495
      Left            =   12960
      TabIndex        =   28
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   327942145
      CurrentDate     =   40055
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forms508.frx":0038
      Height          =   330
      Left            =   840
      TabIndex        =   59
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "工序其它系数"
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   9840
      TabIndex        =   65
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "负责"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   240
      TabIndex        =   60
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "批加"
      Height          =   495
      Left            =   2640
      TabIndex        =   56
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "刷新"
      Height          =   495
      Left            =   2640
      TabIndex        =   53
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "全清"
      Height          =   495
      Left            =   2640
      TabIndex        =   52
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "全选"
      Height          =   495
      Left            =   2640
      TabIndex        =   51
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
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
      Index           =   17
      Left            =   12960
      TabIndex        =   49
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
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
      Index           =   14
      Left            =   12960
      TabIndex        =   48
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "操作员"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   47
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "总工序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   6120
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "排单号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   6720
      TabIndex        =   45
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "产量系数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   9840
      TabIndex        =   44
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "班次产量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3600
      TabIndex        =   43
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "工序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   42
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "刷新"
      Height          =   495
      Left            =   3120
      TabIndex        =   41
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   240
      TabIndex        =   40
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "工序工资"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   6720
      TabIndex        =   39
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "班次"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   240
      TabIndex        =   38
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   6720
      TabIndex        =   37
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "班次匹数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   9840
      TabIndex        =   36
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "工资系数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   6720
      TabIndex        =   35
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   34
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "车台"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6720
      TabIndex        =   33
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   32
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   31
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3600
      TabIndex        =   30
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3600
      TabIndex        =   29
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "Forms508"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
Text1(6).Text = Combo1.Text
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime) AND 班次='" & Combo1.Text & "' ORDER BY 其它系数 DESC"
Adodc1.Refresh
End Sub

Private Sub Combo1_Click()
Text1(6).Text = Combo1.Text
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime) AND 班次='" & Combo1.Text & "' ORDER BY 其它系数 DESC"
Adodc1.Refresh
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Command1_Click()
If Text1(7) = "" Or Text1(8).Text = "" Or Text1(10).Text = "" Or Text1(13).Text = "" Or Text1(18).Text = "" Or Text1(9).Text = "" Or Text1(2).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If

If Len(Text1(8)) / 4 <> Int(Len(Text1(8)) / 4) Then
MsgBox ("员工编号有误！")
Exit Sub
End If

Adodc1.Recordset.AddNew
For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Fields(16) = ""
Adodc1.Recordset.Update
Adodc1.Refresh

Text1(0).Text = ""
Adodc3.RecordSource = "select max(其它系数) FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh
Text1(14).Text = "1"
If Adodc3.Recordset.EOF Then
Text1(14).Text = "1"
Else
Text1(14).Text = Adodc3.Recordset.Fields(0) + 1
End If
Text1(7).SetFocus

End Sub


Private Sub Command2_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If Text1(7) = "" Or Text1(8).Text = "" Or Text1(10).Text = "" Or Text1(13).Text = "" Or Text1(18).Text = "" Or Text1(9).Text = "" Or Text1(2).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If
If Adodc1.Recordset.EOF Then Exit Sub
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
For i = 0 To 18
Adodc1.Recordset.Fields(i) = Text1(i)
Next
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If MsgBox("确定删除吗?", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Text1(8).Text = ""
Text1(0).Text = ""
Adodc3.RecordSource = "select max(其它系数) FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh
Text1(14).Text = "1"
If Adodc3.Recordset.EOF Then
Text1(14).Text = "1"
Else
Text1(14).Text = Adodc3.Recordset.Fields(0) + 1
End If
Text1(7).SetFocus
Command1.Enabled = True

Command3.Enabled = False
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
On Error Resume Next
Adodc1.RecordSource = "select * FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime) ORDER BY 其它系数 DESC"
Adodc1.Refresh
Adodc3.RecordSource = "select max(其它系数) FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh
Text1(14).Text = "1"
If Adodc3.Recordset.EOF Then
Text1(14).Text = "1"
Else
Text1(14).Text = Adodc3.Recordset.Fields(0) + 1
End If
Command1.Enabled = True

Command3.Enabled = False
End Sub


Private Sub Command6_Click()
If Text1(2).Text = "" Then
Adodc1.RecordSource = "select * FROM ZXCL where cast(CONVERT(varchar,时间, 23) as datetime) between cast('" & DTPicker2.value & "' as datetime) and cast('" & DTPicker3.value & "' as datetime) order by 时间"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select * FROM ZXCL where 锅号='" & Text1(2).Text & "' order by 时间"
Adodc1.Refresh
End If
End Sub


Private Sub Command8_Click()
Call BBDY(VSFlexGrid1, 12, 17, "脱水产量")
End Sub

Private Sub DataCombo1_Change()
Text1(11).Text = DataCombo1.Text
Text1(18).Text = "装卸"
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Text1(11).Text = DataCombo1.Text
Text1(18).Text = "装卸"
End Sub

Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DTPicker1_Change()
Text1(9).Text = DTPicker1.value
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime)"
Adodc1.Refresh
End Sub

Private Sub DTPicker1_CloseUp()
Text1(9).Text = DTPicker1.value
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime)"
Adodc1.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
Option4.value = True
For i = 0 To 18
Text1(i).Text = ""
Next
Text1(9).Text = Date
Text1(12).Text = "0"
Text1(13).Text = "0"
Text1(14).Text = "0"
Text1(15).Text = "0"
Text1(16).Text = ""
DataCombo1.Text = ""
Combo1.Text = ""
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
Text2 = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime) AND 班次='" & Combo1.Text & "' ORDER BY 其它系数 DESC"
Adodc1.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select max(其它系数) FROM ZXCL where 时间=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh
Text1(14).Text = "1"
If Adodc3.Recordset.EOF Then
Text1(14).Text = "1"
Else
Text1(14).Text = Adodc3.Recordset.Fields(0) + 1
End If


Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT 工序其它系数 FROM gyshd WHERE 工艺编号 between  '0001' and '1000' GROUP BY 工序其它系数"
Adodc5.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 800
VSFlexGrid1.ColWidth(4) = 800
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(15) = 0
VSFlexGrid1.ColWidth(17) = 0

Command1.Enabled = True

Command3.Enabled = False
End Sub

Private Sub Label2_Click()
YGBL = 8
If InStr(yhm, "mp") > 0 Then
Forms546.Text1(0) = "白坯"
Forms546.Show
End If

If InStr(yhm, "kf") > 0 Then
Forms546.Text1(0) = "定型"
Forms546.Show
End If

If InStr(yhm, "root") > 0 Then
Forms546.Text1(0) = Text1(18)
Forms546.Show
End If
End Sub

Private Sub Label3_Click()
If Option1.value = True Then
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 单据号,客户名称,布类,毛胚幅宽,毛胚重量,毛胚匹数 from ckgl where 单据号='" & Text1(2).Text & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
Text1(17).Text = ""
Text1(1).Text = ""
Text1(3).Text = ""
Text1(4).Text = ""
Text1(5).Text = ""
Text1(10).Text = ""
Text1(0).Text = ""
Else
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select round(SUM(毛胚重量),2),round(SUM(毛胚匹数),2) from ckgl where 单据号='" & Text1(2).Text & "'"
Adodc11.Refresh
Text1(17).Text = Adodc6.Recordset.Fields(0)
Text1(1).Text = Adodc6.Recordset.Fields(1)
Text1(3).Text = Adodc6.Recordset.Fields(2)
Text1(4).Text = Adodc6.Recordset.Fields(3)
Text1(5).Text = Adodc11.Recordset.Fields(1)
Text1(10).Text = Adodc11.Recordset.Fields(0)
Text1(0).Text = Adodc6.Recordset.Fields(5)
End If
End If


If Option2.value = True Then
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 单号,加工单位,品名,颜色,数量,匹数 from jgmx where 单号='" & Text1(2).Text & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
Text1(17).Text = ""
Text1(1).Text = ""
Text1(3).Text = ""
Text1(4).Text = ""
Text1(5).Text = ""
Text1(10).Text = ""
Text1(0).Text = ""
Else
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select ROUND(SUM(数量),2),ROUND(SUM(匹数),2) from jgmx where 单号='" & Text1(2).Text & "'"
Adodc11.Refresh
Text1(17).Text = Adodc6.Recordset.Fields(0)
Text1(1).Text = Adodc6.Recordset.Fields(1)
Text1(3).Text = Adodc6.Recordset.Fields(2)
Text1(4).Text = Adodc6.Recordset.Fields(3)
Text1(5).Text = Adodc11.Recordset.Fields(1)
Text1(10).Text = Adodc11.Recordset.Fields(0)
Text1(0).Text = Adodc6.Recordset.Fields(5)
End If
End If

If Option3.value = True Then
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 锅号,'' as 客户名称,布类,'' as 毛胚幅宽,毛胚重量,实际匹数 from mpbh where 锅号='" & Text1(2).Text & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
Text1(17).Text = ""
Text1(1).Text = ""
Text1(3).Text = ""
Text1(4).Text = ""
Text1(5).Text = ""
Text1(10).Text = ""
Text1(0).Text = ""
Else
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select round(SUM(毛胚重量),2),round(SUM(实际匹数),2) from mpbh where 锅号='" & Text1(2).Text & "'"
Adodc11.Refresh
Text1(17).Text = Adodc6.Recordset.Fields(0)
Text1(1).Text = Adodc6.Recordset.Fields(1)
Text1(3).Text = Adodc6.Recordset.Fields(2)
Text1(4).Text = Adodc6.Recordset.Fields(3)
Text1(5).Text = Adodc11.Recordset.Fields(1)
Text1(10).Text = Adodc11.Recordset.Fields(0)
Text1(0).Text = Adodc6.Recordset.Fields(5)
End If
End If

End Sub

Private Sub Label4_Click()
GXBL = 8
YGBL = 8
Forms545.Text1.Text = Text1(11).Text
Forms545.Show
End Sub

Private Sub Label5_Click()
List1.Clear
If Option1.value = True Then
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "SELECT distinct 单据号 FROM CKGL where 日期=cast('" & DTPicker1.value & "' as datetime) and 负责人 like '%'+'" & Text2 & "'+'%' and 单据号 not in(select distinct 锅号 from zxcl where 时间=cast('" & DTPicker1.value & "' as datetime))"
Adodc13.Refresh
If Adodc13.Recordset.EOF Then Exit Sub
Adodc13.Recordset.MoveFirst
Do While Not Adodc13.Recordset.EOF
List1.AddItem Adodc13.Recordset.Fields(0)
Adodc13.Recordset.MoveNext
Loop
End If

If Option2.value = True Then
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "SELECT distinct 单号 FROM JGMX where 日期=cast('" & DTPicker1.value & "' as datetime) and 负责 like '%'+'" & Text2 & "'+'%' and 单号 not in(select distinct 锅号 from zxcl where 时间=cast('" & DTPicker1.value & "' as datetime))"
Adodc13.Refresh
If Adodc13.Recordset.EOF Then Exit Sub
Adodc13.Recordset.MoveFirst
Do While Not Adodc13.Recordset.EOF
List1.AddItem Adodc13.Recordset.Fields(0)
Adodc13.Recordset.MoveNext
Loop
End If

If Option3.value = True Then
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "SELECT distinct 锅号 FROM mpbh where 出库日期=cast('" & DTPicker1.value & "' as datetime) and 锅号 like '%TK%' and 配缸负责 like '%'+'" & Text2 & "'+'%' and 单据号 not in(select distinct 锅号 from zxcl where 时间=cast('" & DTPicker1.value & "' as datetime))"
Adodc13.Refresh
If Adodc13.Recordset.EOF Then Exit Sub
Adodc13.Recordset.MoveFirst
Do While Not Adodc13.Recordset.EOF
List1.AddItem Adodc13.Recordset.Fields(0)
Adodc13.Recordset.MoveNext
Loop
End If

End Sub

Private Sub Label6_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub Label7_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Label8_Click()
On Error Resume Next
Label8.Enabled = False
If Text1(7) = "" Or Text1(8).Text = "" Or Text1(6).Text = "" Or Text1(13).Text = "" Or Text1(11).Text = "" Then
MsgBox ("输入不完整！")
Label8.Enabled = True
Exit Sub
End If

If Len(Text1(8)) / 4 <> Int(Len(Text1(8)) / 4) Then
MsgBox ("员工编号有误！")
Exit Sub
End If


For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
Text1(2).Text = List1.List(i)

Call Label3_Click


Adodc8.RecordSource = "select * from ZXcl where 锅号='" & Text1(2).Text & "'"
Adodc8.Refresh

If Adodc8.Recordset.EOF Then
Adodc1.Recordset.AddNew
For L = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(L) = Text1(L).Text
Next
Adodc1.Recordset.Fields(16) = ""
Adodc1.Recordset.Update
Adodc1.Refresh
Text1(14).Text = "1"
Else
If MsgBox(Text1(2).Text + "产量已报!,是否继续？", vbYesNo) = vbYes Then
Adodc1.Recordset.AddNew
For L = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(L) = Text1(L).Text
Next
Adodc1.Recordset.Fields(16) = ""
Adodc1.Recordset.Update
Adodc1.Refresh
Text1(14).Text = "1"
End If
End If

End If
Next

Label8.Enabled = True

End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub


Private Sub Text1_Change(Index As Integer)
Select Case Index

Case 2
If InStr(Text1(2).Text, "J") > 0 Or InStr(Text1(2).Text, "j") > 0 Then
Text1(2).Text = Mid(Text1(2).Text, 1, Len(Text1(2).Text) - 1)
Call Label3_Click
Text1(6).SetFocus
End If

Case 5
If Option5.value = True Then
Text1(15).Text = Format(Val(Text1(5).Text) * Val(Text1(13).Text), "#0.000")
End If

Case 10
If Option4.value = True Then
Text1(15).Text = Format(Val(Text1(10).Text) * Val(Text1(13).Text), "#0.000")
End If

Case 13
If Option4.value = True Then
Text1(15).Text = Format(Val(Text1(10).Text) * Val(Text1(13).Text), "#0.000")
End If
If Option5.value = True Then
Text1(15).Text = Format(Val(Text1(5).Text) * Val(Text1(13).Text), "#0.000")
End If

Case 11
Text1(18).Text = "装卸"
End Select
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub









