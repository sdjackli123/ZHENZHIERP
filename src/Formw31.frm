VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw31 
   BackColor       =   &H00C0E0FF&
   Caption         =   "客户账查询---付款"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Height          =   7575
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   14295
      _cx             =   25215
      _cy             =   13361
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
      AllowUserResizing=   0
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   330
      Left            =   8760
      Top             =   10200
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
      Left            =   9360
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
      Left            =   8880
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
      Left            =   8760
      Top             =   10320
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
      Left            =   8640
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
      Top             =   10320
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
      Left            =   9600
      Top             =   10320
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Height          =   330
      Left            =   9480
      Top             =   10320
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
      Left            =   9240
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
      Left            =   9240
      Top             =   10320
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
      Left            =   9840
      Top             =   10560
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
      Height          =   495
      Left            =   9600
      Top             =   10320
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Bindings        =   "Formw31.frx":0000
      Height          =   330
      Left            =   1440
      TabIndex        =   15
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成查询"
      Height          =   855
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "凭证生成"
      Height          =   855
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   423624705
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   423624705
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   11880
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   423624705
      CurrentDate     =   36892
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      Height          =   375
      Index           =   0
      Left            =   11880
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "加工单位"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择日期范围"
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
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Formw31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer

Private Sub Command1_Click()
'On Error Resume Next
'rqq = CDate(Text2.Text) + 1
'Adodc6.Database.Execute "DELETE * FROM JGZCX1"
'Adodc6.Database.Execute "insert into JGZCX1(客户,上期累计应付)  SELECT MID(会计科目,INSTR(会计科目,'-')+1),format(SUM(VAL(余额)),'#0.00') FROM PMMXJZ WHERE 借贷方向='贷' AND 日期=CDATE('" & Text1.Text & "') GROUP BY MID(会计科目,INSTR(会计科目,'-')+1)"
'Adodc3.Database.Execute "insert into JGZCX1(客户,本期应付款) in'd:\数据库\bfrz\" + ljb + "\CW.MDB' SELECT 供应单位,format(SUM(VAL(合计金额)),'#0.00') FROM CKGL WHERE  日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND 库别='采购入库' GROUP BY 供应单位"
'Adodc13.Database.Execute "insert into JGZCX(客户,本期应付款) in'd:\数据库\bfrz\" + ljb + "\cw.mdb' SELECT 客户,format(SUM(VAL(费用)),'#0.00') FROM ZXBZ WHERE  日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) and 类别='应付类' GROUP BY 客户"
'Adodc12.Database.Execute "insert into JGZCX1(客户,本期应付款) in'd:\数据库\bfrz\" + ljb + "\CW.MDB' select 打印模版单位,FORMAT(SUM(VAL(打印模版金额)),'#0.00') from rsrk where 打印模版单位<>'' and 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) GROUP BY 打印模版单位"
'Adodc12.Database.Execute "insert into JGZCX1(客户,本期应付款) in'd:\数据库\bfrz\" + ljb + "\CW.MDB' select 织布单位,FORMAT(SUM(VAL(织布金额)),'#0.00') from rsrk where 织布单位<>'' and 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) GROUP BY 织布单位"
'Adodc12.Database.Execute "insert into JGZCX1(客户,本期应付款) in'd:\数据库\bfrz\" + ljb + "\CW.MDB' select 染色单位,FORMAT(SUM(VAL(金额)),'#0.00') from rsrk where 染色单位<>'' and 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) GROUP BY 染色单位"
'Adodc12.Database.Execute "insert into JGZCX1(客户,本期应付款) in'd:\数据库\bfrz\" + ljb + "\CW.MDB' select 单位,format(SUM(val(数量)*val(单价)),'#0.00') from wxjl where 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) GROUP BY 单位"
'Adodc6.Database.Execute "insert into JGZCX1(客户,本期现付款)  SELECT MID(对方科目,INSTR(对方科目,'-')+1),format(SUM(VAL(贷方金额)),'#0.00') FROM TZJZMX WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND 贷方金额<>'0' and instr(类别,'现金')>0 GROUP BY MID(对方科目,INSTR(对方科目,'-')+1)"
'Adodc6.Database.Execute "insert into JGZCX1(客户,本期银付款)  SELECT MID(对方科目,INSTR(对方科目,'-')+1),format(SUM(VAL(贷方金额)),'#0.00') FROM TZJZMX WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND 贷方金额<>'0' and instr(类别,'银行')>0 GROUP BY MID(对方科目,INSTR(对方科目,'-')+1)"
'Adodc6.Database.Execute "insert into JGZCX1(客户,本期开票)  SELECT 客户,FORMAT(SUM(VAL(开票金额)),'#0.00') FROM JHFP WHERE  开票日期 between cdate('" & Text1 & "') and cdate('" & rqq & "') GROUP BY 客户"
'Adodc6.Database.Execute "insert into JGZCX1(客户,上期累计未开票) SELECT 客户,未开金额 FROM PMJHFP WHERE  结转日期=CDATE('" & Text1.Text & "')"
'Adodc6.Database.Execute "DELETE *  FROM  JGZCX1 WHERE 客户=NULL OR 客户=''"

'Adodc6.RecordSource = "SELECT * FROM JGZCX1"
'Adodc6.Refresh


'If Not Adodc6.Recordset.EOF Then
'Adodc6.Recordset.MoveFirst
'Do While Not Adodc6.Recordset.EOF
'Adodc8.RecordSource = "SELECT * FROM GYS WHERE INSTR(简称,'" & Adodc6.Recordset.Fields(0) & "')>0"
'Adodc8.Refresh
'If Adodc8.Recordset.EOF Then
'Adodc6.Recordset.Delete
'End If
'Adodc6.Recordset.MoveNext
'Loop
'End If

'Adodc6.Database.Execute "UPDATE JGZCX1 SET 类别='1'"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 日期范围='" & Text1.Text & "'+'--'+'" & Text2.Text & "'"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 上期累计应付='0' WHERE 上期累计应付=NULL"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 本期应付款='0' WHERE 本期应付款=NULL"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 本期累计应付款='0' WHERE 本期累计应付款=NULL"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 本期现付款='0' WHERE 本期现付款=NULL"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 本期银付款='0' WHERE 本期银付款=NULL"

'Adodc6.Database.Execute "UPDATE JGZCX1 SET 本期开票='0' WHERE 本期开票=NULL"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 上期累计未开票='0' WHERE 上期累计未开票=NULL"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 本期总付款='0' WHERE 本期总付款=NULL"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 本期未开='0' WHERE 本期未开=NULL"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 本期累计未开='0' WHERE 本期累计未开=NULL"


'Adodc6.Database.Execute "insert into JGZCX1(客户,日期范围,上期累计应付,本期应付款,本期累计应付款,本期现付款,本期银付款,本期开票,上期累计未开票,本期累计未开) SELECT 客户,日期范围,format(SUM(VAL(上期累计应付)),'#0.00'),format(SUM(VAL(本期应付款)),'#0.00'),format(SUM(VAL(本期累计应付款)),'#0.00'),format(SUM(VAL(本期现付款)),'#0.00'),format(SUM(VAL(本期银付款)),'#0.00'),format(SUM(VAL(本期开票)),'#0.00'),format(SUM(VAL(上期累计未开票)),'#0.00'),format(SUM(VAL(本期累计未开)),'#0.00') FROM JGZCX1 GROUP BY 客户,日期范围 "
'Adodc6.Database.Execute "DELETE *  FROM  JGZCX1 WHERE 类别='1'"
'adodc6.database.Execute "UPDATE JGZCX1 SET 本期未开=format(VAL(本期应付款)-VAL(本期开票),'#0.00')"
'Adodc6.Database.Execute "UPDATE JGZCX1 SET 欠款=format(VAL(上期累计应付)+VAL(本期应付款)-VAL(本期现付款)-val(本期银付款),'#0.00'),本期累计应付款=format(VAL(上期累计应付)+VAL(本期应付款),'#0.00'),本期累计未开=format(VAL(上期累计未开票)+VAL(本期应付款)-val(本期开票),'#0.00'),本期总付款=format(val(本期现付款)+val(本期银付款),'#0.00')"


Adodc6.RecordSource = "SELECT 客户,上期累计应付,本期应付款,本期累计应付款,本期现付款,本期银付款,本期总付款,欠款,上期累计未开票,本期应付款 as 本期发生,本期开票,本期累计未开,日期范围 FROM JGZCX1"
Adodc6.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call OutadodcToExcel9(VSFlexGrid1, 2, 3, 4, 5, 6, 7, 8, 9, 10, "客户账目查询--收款" + "截止日期:" + Text2.Text)
End Sub

Private Sub Command4_Click()
Formw332.Combo1.Text = "转账凭证"
Formw332.Show
End Sub

Private Sub Command5_Click()
If MsgBox("操作日期为：" + Trim(DTPicker1.value) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("操作期间为：" + Trim(Month(DTPicker1.value)) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定生成应付系列的凭证吗？", vbYesNo) = vbNo Then Exit Sub
Call CLRKPZ(CDate(Text1.Text), CDate(Text2.Text), CDate(DTPicker1.value))
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If DataCombo1.Text = "" Then
Adodc6.RecordSource = "SELECT 客户,上期累计应付,本期应付款,本期累计应付款,本期现付款,本期银付款,本期总付款,欠款,上期累计未开票,本期应付款 as 本期发生,本期开票,本期累计未开,日期范围 FROM JGZCX1"
Adodc6.Refresh
Else
Adodc6.RecordSource = "SELECT 客户,上期累计应付,本期应付款,本期累计应付款,本期现付款,本期银付款,本期总付款,欠款,上期累计未开票,本期应付款 as 本期发生,本期开票,本期累计未开,日期范围 FROM JGZCX1 WHERE 客户='" & DataCombo1.Text & "'"
Adodc6.Refresh
End If
End Sub

Private Sub DTPicker3_Change()
Text1.Text = DTPicker3.value
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = DTPicker3.value
Text1.SetFocus
End Sub

Private Sub DTPicker4_Change()
Text2.Text = DTPicker4.value
End Sub

Private Sub DTPicker4_CloseUp()
Text2.Text = DTPicker4.value
Text2.SetFocus
End Sub

Private Sub Form_Load()
Text1.Text = Date
Text2.Text = Date
DTPicker1.value = Date
DTPicker3.value = Date
DTPicker4.value = Date
DataCombo1.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 简称 from GYS  GROUP BY 简称"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 供应单位,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,单据号,日期,是否开票,开票,开票日期 from ckgl where 供应单位='" & DataCombo1.Text & "' and 日期 between cdate('" & Text1 & "') and cdate('" & Text2.Text & "') AND 库别='采购入库'"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT 客户,上期累计应付,本期应付款,本期累计应付款,本期现付款,本期银付款,本期总付款,欠款,上期累计未开票,本期应付款 as 本期发生,本期开票,本期累计未开,日期范围 FROM JGZCX1"
Adodc6.Refresh
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "rqsd"
Adodc7.Refresh
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(0) = 300
For i = 1 To 12
VSFlexGrid1.ColWidth(i) = 1300
Next
VSFlexGrid1.ColWidth(13) = 2600

End Sub

Private Sub Label3_DblClick()
DataCombo1.Text = ""
End Sub

Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
S1 = VSFlexGrid1.RowSel
End Sub

Private Sub VSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
S2 = VSFlexGrid1.RowSel
End Sub


Private Sub CLRKPZ(DT1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Adodc10.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & DT1 & "') AND CDATE('" & dt2 & "') and instr(制单,'自动-材料')>0"
Adodc10.Refresh
If Not Adodc10.Recordset.EOF Then
If MsgBox("已有应付生成凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
'Adodc11.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(制单,'自动-材料')>0 and 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Adodc9.RecordSource = "SELECT * FROM JGZCX1 where val(本期应付款)>0"
Adodc9.Refresh


If Adodc9.Recordset.EOF Then Exit Sub
Adodc10.RecordSource = "SELECT * FROM CLZZPZ"
Adodc10.Refresh
Adodc11.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & DT1 & "') AND CDATE('" & dt2 & "')"
Adodc11.Refresh
PZH = "5-1"
If Adodc11.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Adodc11.Recordset.Fields(0) + 1)
End If
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 3
Adodc10.Recordset.AddNew
Adodc10.Recordset.Fields(0) = "购材料"
Adodc10.Recordset.Fields(1) = "原材料"
Adodc10.Recordset.Fields(2) = ""
Adodc10.Recordset.Fields(3) = "应付账款"
Adodc10.Recordset.Fields(4) = Adodc9.Recordset.Fields(0)
Adodc10.Recordset.Fields(5) = Format(Adodc9.Recordset.Fields(2), "#0.00")
Adodc10.Recordset.Fields(6) = PZH
Adodc10.Recordset.Fields(7) = CDate(dt3)
Adodc10.Recordset.Fields(8) = ""
Adodc10.Recordset.Fields(9) = ""
Adodc10.Recordset.Fields(10) = ""
Adodc10.Recordset.Fields(11) = "自动-材料"
Adodc10.Recordset.Update


'adodc10.Recordset.AddNew
'adodc10.Recordset.Fields(0) = "购材料"
'adodc10.Recordset.Fields(1) = "应交税金"
'adodc10.Recordset.Fields(2) = "税金进项"
'adodc10.Recordset.Fields(3) = "应付账款"
'adodc10.Recordset.Fields(4) = adodc9.Recordset.Fields(0)
'adodc10.Recordset.Fields(5) = Format(adodc9.Recordset.Fields(2) * 0.17, "#0.00")
'adodc10.Recordset.Fields(6) = PZH
'adodc10.Recordset.Fields(7) = CDate(dt3)
'adodc10.Recordset.Fields(8) = ""
'adodc10.Recordset.Fields(9) = ""
'adodc10.Recordset.Fields(10) = ""
'adodc10.Recordset.Fields(11) = "自动-材料"
'adodc10.Recordset.Update


Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Adodc11.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & DT1 & "') AND CDATE('" & dt2 & "')"
Adodc11.Refresh
PZH = "5-1"
If Adodc11.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Adodc11.Recordset.Fields(0) + 1)
End If
Loop
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
End Sub

