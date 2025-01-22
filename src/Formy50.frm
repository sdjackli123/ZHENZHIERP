VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formy50 
   BackColor       =   &H00C0E0FF&
   Caption         =   "配料审核"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   260
      Left            =   8520
      Top             =   9840
      Visible         =   0   'False
      Width           =   1930
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
   Begin VB.Data Data6 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   250
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3610
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   250
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   2530
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   250
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   2290
   End
   Begin VB.Data Data4 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   250
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   2170
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "反审"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "审核"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4560
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   5940
      ItemData        =   "Formy50.frx":0000
      Left            =   240
      List            =   "Formy50.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   47
      Top             =   3360
      Width           =   1940
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   15840
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   15840
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   15840
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1575
      Left            =   10680
      TabIndex        =   9
      Top             =   360
      Width           =   4935
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "未审"
         Height          =   255
         Index           =   11
         Left            =   3960
         TabIndex        =   50
         Top             =   1200
         Width           =   850
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "名称"
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "审核日期"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "料单"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "配料日期"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "未称"
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "已称"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "已审"
         Height          =   255
         Index           =   8
         Left            =   2640
         TabIndex        =   13
         Top             =   1200
         Width           =   850
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "机台"
         Height          =   255
         Index           =   9
         Left            =   3960
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色别"
         Height          =   255
         Index           =   10
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10560
      Top             =   0
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   8
      Text            =   "Text8"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Text            =   "Text8"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text8"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   600
      Width           =   495
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
      Height          =   320
      ItemData        =   "Formy50.frx":0004
      Left            =   7200
      List            =   "Formy50.frx":0006
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1090
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Text            =   "Text3"
      Top             =   1560
      Width           =   1330
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   5280
      Top             =   10440
      Visible         =   0   'False
      Width           =   2540
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
      Height          =   380
      Left            =   7560
      Top             =   9840
      Visible         =   0   'False
      Width           =   1820
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
      Width           =   2300
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
      Width           =   2300
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
      Left            =   6840
      Top             =   10560
      Visible         =   0   'False
      Width           =   3260
      _ExtentX        =   5741
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
      Height          =   500
      Left            =   6960
      Top             =   10440
      Visible         =   0   'False
      Width           =   3020
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
      Bindings        =   "Formy50.frx":0008
      Height          =   7455
      Left            =   3240
      TabIndex        =   24
      Top             =   2160
      Width           =   15495
      _cx             =   27331
      _cy             =   13150
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   4080
      TabIndex        =   25
      Top             =   1560
      Width           =   1220
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formy50.frx":001D
      Height          =   330
      Left            =   5400
      TabIndex        =   26
      Top             =   720
      Width           =   1700
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330104833
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330039297
      CurrentDate     =   39177
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formy50.frx":0032
      Height          =   330
      Left            =   5400
      TabIndex        =   29
      Top             =   1560
      Width           =   1700
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Formy50.frx":0047
      Height          =   330
      Left            =   4080
      TabIndex        =   30
      Top             =   720
      Width           =   1220
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "染化助库名"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "Formy50.frx":005C
      Height          =   290
      Left            =   7200
      TabIndex        =   31
      Top             =   720
      Width           =   2530
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo2"
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   840
      TabIndex        =   52
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330039297
      CurrentDate     =   39177
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "审核"
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
      Index           =   1
      Left            =   240
      TabIndex        =   53
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "料单信息"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   49
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料名称"
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
      Left            =   4080
      TabIndex        =   44
      Top             =   1200
      Width           =   1220
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "料单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Index           =   1
      Left            =   5400
      TabIndex        =   43
      Top             =   360
      Width           =   1700
   End
   Begin VB.Label Label6 
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
      Index           =   0
      Left            =   240
      TabIndex        =   42
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label7 
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
      Index           =   0
      Left            =   240
      TabIndex        =   41
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "库类"
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
      Index           =   5
      Left            =   4080
      TabIndex        =   40
      Top             =   360
      Width           =   1220
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
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
      Index           =   1
      Left            =   5400
      TabIndex        =   39
      Top             =   1200
      Width           =   1700
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   38
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   37
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   36
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   35
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "机台"
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
      Index           =   3
      Left            =   7200
      TabIndex        =   34
      Top             =   1200
      Width           =   1090
   End
   Begin VB.Label Label7 
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
      Height          =   380
      Index           =   2
      Left            =   7200
      TabIndex        =   33
      Top             =   360
      Width           =   1700
   End
   Begin VB.Label Label7 
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
      Height          =   380
      Index           =   3
      Left            =   8400
      TabIndex        =   32
      Top             =   1200
      Width           =   1330
   End
End
Attribute VB_Name = "Formy50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim plshsx, yssj As Integer
Dim sz(9) As String: Dim ZS(10) As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
'Call OutadodcToExcel(VSFlexGrid1, 8, "领料车间" + DataCombo1.Text)
Call pfdfj(DataCombo1)
Call plda(Data2, Data3, Data4, DataCombo1, Adodc3)
End Sub
Private Sub Command3_Click()
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "update pldr  set 审核='是',审核日期='" & DTPicker3.value & "' where 料单编号='" & List1.List(i) & "'"
sql2 = "insert into pldsh(审核,审核日期,审核人,料单编号,操作日期) VALUES('是','" & DTPicker3.value & "','" & yhm & "','" & List1.List(i) & "','" & Date & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
Next
Call Command4_Click
End Sub

Private Sub Command4_Click()
On Error Resume Next
sql1 = ""
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")

If Check2(0).value = 1 Then
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text2(0) + ":" + Text2(1) + ":" + Text2(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text8(0) + ":" + Text8(1) + ":" + Text8(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "配料日期 between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "料单编号 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "客户名称='" & DataCombo5 & "' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "锅号='" & DataCombo3.Text & "' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "审核日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "染化助名称 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "isnull(称量标记,'')='Y' and "
End If

If Check2(8).value = 1 Then
sql1 = sql1 + "isnull(审核,'')='是' and "
End If

If Check2(9).value = 1 Then
sql1 = sql1 + "机台 like '%'+'" & Combo2 & "'+'%' and "
End If

If Check2(10).value = 1 Then
sql1 = sql1 + "颜色 like '%'+'" & Text3 & "'+'%' and "
End If

If Check2(11).value = 1 Then
sql1 = sql1 + "isnull(审核,'')<>'是' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "isnull(称量标记,'')<>'Y' and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc1.RecordSource = "SELECT * FROM v_pld_kpd_cx WHERE (" + sql1 + ") order by 料单编号,工序名称,次序号"
Adodc1.Refresh
Adodc5.RecordSource = "SELECT distinct 料单编号 FROM v_pld_kpd_cx WHERE (" + sql1 + ") order by 料单编号"
Adodc5.Refresh


With VSFlexGrid1
    .WordWrap = True
    .MergeCells = 2
    .MergeCol(1 - 5) = True '是否上下列合并
End With

List1.Clear
If Not Adodc5.Recordset.EOF Then
Adodc5.Recordset.MoveFirst
Do While Not Adodc5.Recordset.EOF
List1.AddItem Trim(Adodc5.Recordset.Fields(0))
Adodc5.Recordset.MoveNext
Loop
End If

VSFlexGrid1.ColFormat(11) = "#0.####"
VSFlexGrid1.ColFormat(12) = "#0.####"

VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 0, 11, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 12, , &HC0C0&

End Sub


Private Sub Command5_Click()
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "update pldr  set 审核='',审核日期=null where 料单编号='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
Call Command4_Click
End Sub

Private Sub Command8_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Command9_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub Form_Load()
Combo2 = ""
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
plshsx = 1
Check2(0).value = 1
Text2(0) = "00"
Text2(1) = "00"
Text2(2) = "00"
Text1 = ""
Text8(0) = "00"
Text8(1) = "00"
Text8(2) = "00"
Text3 = ""
yssj = 1

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Data2.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data3.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data4.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data6.DatabaseName = App.Path & "\AccessBase\DB.mdb"

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select distinct 染化助库名 from rhzh"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc6.Refresh

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM v_pld_kpd_cx WHERE 配料日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) order by 编码"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 200
End Sub

Private Sub Text1_Change()
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc6.Refresh
End Sub

Private Sub Timer1_Timer()
If yssj = 5 Then
Timer1.Enabled = False
ddchxmx = 0
yssj = 0
End If
Call Command4_Click
yssj = 5
End Sub

Private Sub pfdfj(ldbh As String)
On Error Resume Next
Data6.Database.Execute "delete * from plda"

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select * from pld where  编号='" & ldbh & "'"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
Adodc7.Recordset.MoveFirst

For i = 0 To 10
ZS(i) = Adodc7.Recordset.Fields(i)
Next

mb = 0
For i = 12 To 61
If Adodc7.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next

For i = 12 To mb + 12
If Adodc7.Recordset.Fields(i) <> "" Then
sz(0) = Mid(Adodc7.Recordset.Fields(i), 1, InStr(Adodc7.Recordset.Fields(i), "(") - 1)
sz(1) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "(") + 1, InStr(Adodc7.Recordset.Fields(i), ")") - InStr(Adodc7.Recordset.Fields(i), "(") - 1)
sz(2) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), ")") + 1, InStr(Adodc7.Recordset.Fields(i), "-") - InStr(Adodc7.Recordset.Fields(i), ")") - 1)
sz(3) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "-") + 1, InStr(Adodc7.Recordset.Fields(i), "\") - InStr(Adodc7.Recordset.Fields(i), "-") - 1)
sz(4) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "\") + 1, InStr(Adodc7.Recordset.Fields(i), "#") - InStr(Adodc7.Recordset.Fields(i), "\") - 1)
sz(5) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "#") + 1, InStr(Adodc7.Recordset.Fields(i), "^") - InStr(Adodc7.Recordset.Fields(i), "#") - 1)
sz(6) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "^") + 1, InStr(Adodc7.Recordset.Fields(i), "[") - InStr(Adodc7.Recordset.Fields(i), "^") - 1)
sz(7) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "[") + 1, InStr(Adodc7.Recordset.Fields(i), "]") - InStr(Adodc7.Recordset.Fields(i), "[") - 1)
sz(8) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "]") + 1, InStr(Adodc7.Recordset.Fields(i), "{") - InStr(Adodc7.Recordset.Fields(i), "]") - 1)
sz(9) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "{") + 1)

L = i - 11
Data6.Database.Execute "insert into plda(审核,锅号,压力,生产类别,配方单,重量,配料打印员,染化助单价,配料日期,生产信息,料单编号,工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,车速,次序号) VALUES('" & ZS(0) & "','" & ZS(1) & "','" & ZS(2) & "','" & ZS(3) & "','" & ZS(4) & "','" & ZS(5) & "','" & ZS(6) & "','" & ZS(7) & "','" & ZS(8) & "','" & ZS(9) & "','" & ZS(10) & "','" & sz(0) & "','" & sz(1) & "','" & sz(2) & "',  " & _
                                                                        "'" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & L & "')"
End If
Next
End If

End Sub





