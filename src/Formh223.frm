VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formh223 
   BackColor       =   &H00C0E0FF&
   Caption         =   "确认工艺信息"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form23"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   480
      Style           =   1  'Simple Combo
      TabIndex        =   40
      Text            =   "Combo1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "档案查询"
      Height          =   495
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   480
      Width           =   1210
   End
   Begin VB.Data Data10 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   7920
      Top             =   10560
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
      Left            =   8160
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
      Left            =   7920
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
      Height          =   375
      Left            =   7800
      Top             =   10560
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
      Left            =   7680
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
      Left            =   6960
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
      Left            =   6360
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
      Height          =   495
      Left            =   6720
      Top             =   10440
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
      Height          =   615
      Left            =   6360
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
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
      Left            =   6960
      Top             =   10680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Left            =   6480
      Top             =   10440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   6000
      Top             =   10440
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
      Bindings        =   "Formh223.frx":0000
      Height          =   5055
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   13935
      _cx             =   24580
      _cy             =   8916
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
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转入生产"
      Height          =   495
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   6240
      Top             =   120
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新编号"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "配方单"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1560
      Width           =   1210
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "档案保存"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "档案删除"
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "客户信息"
      Height          =   3375
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   9735
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   5520
         TabIndex        =   52
         Text            =   "Text5"
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   1
         Left            =   8400
         TabIndex        =   45
         Text            =   "Text4"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   2
         Left            =   8400
         TabIndex        =   44
         Text            =   "Text4"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   3
         Left            =   8400
         TabIndex        =   43
         Text            =   "Text4"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   4
         Left            =   8400
         TabIndex        =   42
         Text            =   "Text4"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   5
         Left            =   8400
         TabIndex        =   41
         Text            =   "Text4"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   0
         Left            =   8400
         TabIndex        =   38
         Text            =   "Text4"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   600
         TabIndex        =   32
         Text            =   "Text3"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   600
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1680
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   0
         Left            =   5520
         TabIndex        =   26
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5520
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   2640
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   16777215
         Format          =   329777153
         CurrentDate     =   39961
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8421440
         Format          =   329777153
         CurrentDate     =   39961
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh223.frx":0015
         Height          =   330
         Index           =   1
         Left            =   1320
         TabIndex        =   18
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "简称"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh223.frx":002A
         Height          =   330
         Index           =   2
         Left            =   5520
         TabIndex        =   23
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "负责人姓名"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   3
         Left            =   1320
         TabIndex        =   20
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   4
         Left            =   1320
         TabIndex        =   21
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh223.frx":003F
         Height          =   330
         Index           =   5
         Left            =   1320
         TabIndex        =   19
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "pm"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   6
         Left            =   5520
         TabIndex        =   24
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   7
         Left            =   5520
         TabIndex        =   25
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   290
         Index           =   8
         Left            =   4200
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   490
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "预定"
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
         Index           =   17
         Left            =   7800
         TabIndex        =   51
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "工时"
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
         Left            =   7800
         TabIndex        =   50
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "中定"
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
         Left            =   7800
         TabIndex        =   49
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "烘干"
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
         Left            =   7800
         TabIndex        =   48
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "硅油"
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
         Left            =   7800
         TabIndex        =   47
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "成定"
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
         Left            =   7800
         TabIndex        =   46
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "工艺说明"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Index           =   11
         Left            =   4680
         TabIndex        =   39
         Top             =   120
         Width           =   850
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "色号"
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
         TabIndex        =   36
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "色号快捷"
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
         Left            =   4680
         TabIndex        =   30
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1100
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "浴比"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Index           =   0
         Left            =   4200
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   490
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "配方编号"
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
         Left            =   4680
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "日期"
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
         Left            =   4680
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "颜色 "
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
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1100
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
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "客户名称"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "负责人"
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
         Left            =   4680
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "报价"
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
         Left            =   4680
         TabIndex        =   3
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1080
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   10560
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Formh223"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer: Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Dim sz(6) As String: Public JDBAR As Integer

Private Sub Command1_Click()
If DataCombo1(4).Text = "" Then
Adodc5.RecordSource = "SELECT RQ AS 日期,KH AS 客户,fz AS 工艺负责人,YS AS 颜色,SH AS 色号,BL AS 品名,DH AS 编号,ZL AS 生成报价,IP AS 浴比,xs as 工艺说明 FROM ZH2 WHERE RQ BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime)  ORDER BY RQ,DH"
Adodc5.Refresh
Else
Adodc5.RecordSource = "SELECT RQ AS 日期,KH AS 客户,fz AS 工艺负责人,YS AS 颜色,SH AS 色号,BL AS 品名,DH AS 编号,ZL AS 生成报价,IP AS 浴比,xs as 工艺说明 FROM ZH2 WHERE SH='" & DataCombo1(4).Text & "'  ORDER BY RQ,DH"
Adodc5.Refresh
End If
Command5.Enabled = True
Command3.Enabled = False
Command4.Enabled = False

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 500
Next
End If

End Sub

Private Sub Command10_Click()
'On Error Resume Next
If DataCombo1(6).Text = "" Then
MsgBox ("请输入配方编号")
Exit Sub
End If
If MsgBox("确定转入生产吗？", vbYesNo) = vbNo Then Exit Sub
Adodc11.RecordSource = "select * from dpfd WHERE 编号='" & DataCombo1(6).Text & "'"
Adodc11.Refresh
If Adodc11.Recordset.EOF Then
MsgBox ("没有此配方编号")
Exit Sub
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "hysdpfdzr('" & DataCombo1(6).Text & "','" & Date & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel

MsgBox ("转入成功！")
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command3_Click()
On Error Resume Next
If DataCombo1(4).Text = "" Or DataCombo1(6).Text = "" Then
MsgBox ("色号、编号须填完整！")
Exit Sub
End If
If DataCombo1(7).Text = "" Then
MsgBox ("是否报价须填完整！")
Exit Sub
End If

For i = 0 To Adodc5.Recordset.Fields.count - 1
Adodc5.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc5.Recordset.Update
Adodc5.Refresh

End Sub

Private Sub Command4_Click()
On Error Resume Next
If Adodc5.Recordset.EOF Then Exit Sub
If MsgBox("确定删除吗？删除同时配方也被清除！", vbYesNo) = vbNo Then Exit Sub
Adodc5.Recordset.Delete
Adodc5.Refresh
sql1 = "delete  from dPFD where 编号='" & DataCombo1(6).Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End Sub

Private Sub Command5_Click()
If DataCombo1(4).Text = "" Or DataCombo1(6).Text = "" Then
MsgBox ("色号、编号须填完整！")
Exit Sub
End If

If DataCombo1(7).Text = "" Then
MsgBox ("是否报价须填完整！")
Exit Sub
End If

Adodc7.RecordSource = "select * from dbdpfdph"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
If Val(Mid(DataCombo1(6).Text, 3)) <= Adodc7.Recordset.Fields(0) Then
If MsgBox("已存在此配方编号，是否自动新编号？", vbYesNo) = vbNo Then Exit Sub
Adodc1.RecordSource = "SELECT * FROM dbdPFDbh"
Adodc1.Refresh
DataCombo1(6).Text = "PD1"
If Not Adodc1.Recordset.EOF Then
L = Adodc1.Recordset.Fields(0)
DataCombo1(6).Text = yhdm + Trim(L + 1) '''''''''''''OK
Else
DataCombo1(6).Text = yhdm + "1"
End If
End If
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "hyszh2('" & Now & "','" & DataCombo1(1) & "','" & DataCombo1(2) & "','" & DataCombo1(3) & "','" & DataCombo1(4) & "','" & DataCombo1(5) & "','" & DataCombo1(6) & "','" & DataCombo1(7) & "','" & DataCombo1(8) & "','" & Text5 & "','" & Text4(0) & "','" & Text4(1) & "','" & Text4(2) & "','" & Text4(3) & "','" & Text4(4) & "','" & Text4(5) & "')"     ' 表示调用哪个存储过程
Set L = g_Cmd.Execute             ' 执行存储过程
    g_Cmd.Cancel
    

Adodc5.Refresh

Call Command6_Click
End Sub

Private Sub Command6_Click()
'On Error Resume Next

Data10.Database.Execute "delete * from dpfda"

Adodc11.RecordSource = "select * from dpfd where  编号='" & DataCombo1(6).Text & "'"
Adodc11.Refresh
If Not Adodc11.Recordset.EOF Then
Adodc11.Recordset.MoveFirst
mb = 0
For i = 7 To 56
If Adodc11.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next

ProgressBar1.Visible = True
Timer1.Enabled = True
For i = 7 To mb + 7
If Adodc11.Recordset.Fields(i) <> "" Then
sz(0) = Mid(Adodc11.Recordset.Fields(i), 1, InStr(Adodc11.Recordset.Fields(i), "(") - 1)
sz(1) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "(") + 1, InStr(Adodc11.Recordset.Fields(i), ")") - InStr(Adodc11.Recordset.Fields(i), "(") - 1)
sz(2) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), ")") + 1, InStr(Adodc11.Recordset.Fields(i), "-") - InStr(Adodc11.Recordset.Fields(i), ")") - 1)
sz(3) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "-") + 1, InStr(Adodc11.Recordset.Fields(i), "\") - InStr(Adodc11.Recordset.Fields(i), "-") - 1)
sz(4) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "\") + 1, InStr(Adodc11.Recordset.Fields(i), "#") - InStr(Adodc11.Recordset.Fields(i), "\") - 1)
sz(5) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "#") + 1, InStr(Adodc11.Recordset.Fields(i), "^") - InStr(Adodc11.Recordset.Fields(i), "#") - 1)
sz(6) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "^") + 1)
L = i - 6
'原来生产种类为生产类别，与access数据库db对应
Data10.Database.Execute "insert into dpfda(加工单位,品名,颜色,色号,负责人,生产种类,配方编号,配方日期,工序名称,浴比,染化助库,染化助名称,单位,配方,车速,次序号) VALUES('" & DataCombo1(1).Text & "','" & DataCombo1(5).Text & "','" & DataCombo1(3).Text & "','" & DataCombo1(4).Text & "','" & DataCombo1(2).Text & "','" & DataCombo1(7).Text & "','" & DataCombo1(6).Text & "',CDATE('" & DataCombo1(0).Text & "'),'" & sz(0) & "','" & sz(1) & "','" & sz(2) & "',  " & _
                                                                        "'" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & L & "')"
ProgressBar1.value = 100 / mb * L
End If
Next
ProgressBar1.Visible = False
Timer1.Enabled = False
End If

Formh235.DataCombo1(0).Text = DataCombo1(1).Text   ''加工单位
Formh235.DataCombo1(1).Text = DataCombo1(5).Text   ''品名
Formh235.DataCombo1(3).Text = DataCombo1(3).Text   ''颜色
Formh235.DataCombo1(2).Text = DataCombo1(4).Text   ''色号
Formh235.DataCombo1(14).Text = DataCombo1(2).Text  ''负责人
Formh235.DataCombo1(15).Text = DataCombo1(7).Text  ''生产类别  报价
Formh235.DataCombo1(12).Text = DataCombo1(6).Text  ''配方编号
Formh235.DataCombo1(11).Text = DataCombo1(0).Text  ''配方日期
Formh235.Show
End Sub

Private Sub Command7_Click()
On Error Resume Next
DataCombo1(6).Text = "PD1"
Adodc1.RecordSource = "SELECT * FROM dbdPFDbh"
Adodc1.Refresh

If Not Adodc1.Recordset.EOF Then
L = Adodc1.Recordset.Fields(0)
DataCombo1(6).Text = yhdm + Trim(L + 1) '''''''''''''OK
Else
DataCombo1(6).Text = "PD1"
End If

DataCombo1(7).Text = "未"
DataCombo1(0).Text = Date

End Sub

Private Sub Command9_Click()
Formh222.Show
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode

End Sub

Private Sub Form_Load()

On Error Resume Next
Dim L As String

DTPicker1.value = Date - 30
DTPicker2.value = Date
Text1.Text = ""
Text3.Text = ""
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT RQ AS 日期,KH AS 客户,fz AS 工艺负责人,YS AS 颜色,SH AS 色号,BL AS 品名,DH AS 编号,ZL AS 生成报价,IP AS 浴比,xs as 工艺说明 FROM ZH2 WHERE RQ BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) ORDER BY DH DESC"
Adodc5.Refresh

For i = 0 To Adodc5.Recordset.Fields.count - 1
DataCombo1(i).Text = ""
Text4(i) = 0
Next


DataCombo1(0).Text = Date
DataCombo1(7).Text = "未"

Text2.Text = ""
Text5.Text = ""

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset


Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from khZL where ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc2.Refresh


Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT 负责人姓名 FROM GR GROUP BY 负责人姓名"
Adodc4.Refresh

DataCombo1(6).Text = "PD1"
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM dbdPFDbh"
Adodc1.Refresh

If Not Adodc1.Recordset.EOF Then
L = Adodc1.Recordset.Fields(0)
DataCombo1(6).Text = yhdm + Trim(L + 1) '''''''''''''OK
Else
DataCombo1(6).Text = yhdm + "1"
End If

ProgressBar1.Visible = False
Timer1.Enabled = False

Fm = 223

Command5.Enabled = True
Command3.Enabled = False
Command4.Enabled = False

Data10.DatabaseName = App.Path & "\AccessBase\db.mdb"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 1200
VSFlexGrid1.ColWidth(5) = 1000
VSFlexGrid1.ColWidth(6) = 1500
VSFlexGrid1.ColWidth(7) = 1500
VSFlexGrid1.ColWidth(8) = 2000
VSFlexGrid1.ColWidth(9) = 1000

DataCombo1(1).TabIndex = 0

If Len(yhdm) <> 1 Then
MsgBox ("这个账户不合适进入这个界面")
Command1.Enabled = False
Command5.Enabled = False
Command3.Enabled = False
Command7.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Command10.Enabled = False
End If

End Sub

Private Sub Label1_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case 3
       Text4(0) = 1500
       Case 11
       Text4(2) = 1000
       Case 12
       Text4(3) = 1000
       Case 13
       Text4(4) = 300
       Case 16
       Text4(5) = 1000
End Select
End Sub

Private Sub Label1_DblClick(Index As Integer)
On Error Resume Next
Select Case Index
       Case 3
       Text4(0) = 0
       Case 11
       Text4(2) = 0
       Case 12
       Text4(3) = 0
       Case 13
       Text4(4) = 0
       Case 16
       Text4(5) = 0
End Select
End Sub

Private Sub Label2_Click()
Formh226.Text1 = DataCombo1(4)
Formh226.Show
End Sub

Private Sub Text1_Change()
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select PM from PM where pm like '%'+ '" & Text1.Text & "'+'%' group by PM"
Adodc3.Refresh
End Sub

Private Sub Text3_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text3 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'  group by 简称"
Adodc2.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move rs - 1
For i = 0 To Adodc5.Recordset.Fields.count - 1
DataCombo1(i).Text = Adodc5.Recordset.Fields(i)
Next
Text5 = Adodc5.Recordset.Fields(11)
Command5.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Text2_Change()
      Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
      Adodc5.RecordSource = "SELECT RQ AS 日期,KH AS 客户,fz AS 工艺负责人,YS AS 颜色,SH AS 色号,BL AS 品名,DH AS 编号,ZL AS 生成报价,IP AS 浴比,xs as 工艺说明 FROM ZH2 WHERE  SH like '%'+'" & Text2.Text & "'+'%'   ORDER BY DH DESC"
      Adodc5.Refresh
      
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 500
Next
End If

End Sub


Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
    If c <> 7 Then
    
        
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
                
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
    End If
End With
End Sub


Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    VSFlexGrid1.Text = Combo1111.Text
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub

Private Sub Combo1111_LostFocus()
On Error Resume Next
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move r - 1
Adodc5.Recordset.Fields(c - 1) = Combo1111.Text
Adodc5.Recordset.Update
Combo1111.Visible = False
VSFlexGrid4.SetFocus
End Sub


