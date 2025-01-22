VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw114 
   BackColor       =   &H00C0E0FF&
   Caption         =   "特种记账"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   7560
      Top             =   9600
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      Left            =   7800
      Top             =   9840
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
      Left            =   7680
      Top             =   10080
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
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
      Height          =   3015
      Left            =   480
      TabIndex        =   51
      Top             =   5880
      Width           =   13815
      _cx             =   24368
      _cy             =   5318
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   0
      Left            =   1920
      TabIndex        =   39
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      ItemData        =   "Formw114.frx":0000
      Left            =   12600
      List            =   "Formw114.frx":000A
      TabIndex        =   37
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1800
      Width           =   1095
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/M/d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   28
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/M/d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   27
      Top             =   600
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw114.frx":001E
      Height          =   2055
      Left            =   840
      TabIndex        =   14
      Top             =   3360
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   14
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw114.frx":0032
      Height          =   330
      Index           =   0
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   2
      Left            =   3360
      TabIndex        =   8
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   3
      Left            =   3360
      TabIndex        =   9
      Top             =   2280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   4
      Left            =   6720
      TabIndex        =   10
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   5
      Left            =   6720
      TabIndex        =   11
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   6
      Left            =   6720
      TabIndex        =   12
      Top             =   2280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   7
      Left            =   6720
      TabIndex        =   13
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   8
      Left            =   10560
      TabIndex        =   19
      Top             =   1320
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   9
      Left            =   10560
      TabIndex        =   20
      Top             =   1800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   10
      Left            =   10560
      TabIndex        =   21
      Top             =   2280
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   11
      Left            =   10560
      TabIndex        =   22
      Top             =   2760
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   5160
      TabIndex        =   29
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81526785
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1920
      TabIndex        =   30
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81526785
      CurrentDate     =   36892
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   1
      Left            =   1920
      TabIndex        =   40
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   2
      Left            =   1920
      TabIndex        =   41
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   3
      Left            =   1920
      TabIndex        =   42
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   4
      Left            =   5280
      TabIndex        =   43
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   5
      Left            =   5280
      TabIndex        =   44
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   6
      Left            =   5280
      TabIndex        =   45
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   7
      Left            =   5280
      TabIndex        =   46
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   8
      Left            =   9120
      TabIndex        =   47
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   9
      Left            =   9120
      TabIndex        =   48
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   10
      Left            =   9120
      TabIndex        =   49
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   11
      Left            =   9120
      TabIndex        =   50
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "打印选择"
      Height          =   375
      Left            =   11520
      TabIndex        =   38
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "摘要"
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
      Left            =   840
      TabIndex        =   36
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "对方科目"
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
      Left            =   4200
      TabIndex        =   33
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   32
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   840
      TabIndex        =   31
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注2"
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
      Index           =   11
      Left            =   7560
      TabIndex        =   26
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注3"
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
      Index           =   10
      Left            =   7560
      TabIndex        =   25
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注4"
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
      Index           =   9
      Left            =   7560
      TabIndex        =   24
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
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
      Index           =   8
      Left            =   7560
      TabIndex        =   23
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   7
      Left            =   840
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注1"
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
      Index           =   6
      Left            =   4200
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "贷方金额"
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
      Index           =   5
      Left            =   4200
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "借方金额"
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
      Index           =   4
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
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
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "类别"
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
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "Formw114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If MsgBox("确认删除吗", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
For i = 1 To 11
DataCombo1(i).Text = ""
Next
DataCombo1(1).Text = Date
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If DataCombo1(1).Text = "" Or DataCombo1(4).Text = "" Or DataCombo1(5).Text = "" Or DataCombo1(6).Text = "" Then Exit Sub
Adodc1.Recordset.AddNew
For i = 0 To Adodc1.Recordset.Fields.Count - 1
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
For i = 1 To 11
DataCombo1(i).Text = ""
Next
DataCombo1(1).Text = Date
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If DataCombo1(1).Text = "" Or DataCombo1(4).Text = "" Or DataCombo1(5).Text = "" Or DataCombo1(6).Text = "" Then Exit Sub
If MsgBox("确认修改吗", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Edit
For i = 0 To Adodc1.Recordset.Fields.Count - 1
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
For i = 1 To 11
DataCombo1(i).Text = ""
Next
DataCombo1(1).Text = Date
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command5_Click()
If Combo1.Text = "" Then
Adodc1.RecordSource = "SELECT 类别,日期,单据号,摘要,对方科目,借方金额,贷方金额,备注1,备注2,备注3,备注4 FROM TZJZMX WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')  ORDER BY 日期,序号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT 类别,日期,单据号,摘要,对方科目,借方金额,贷方金额,备注1,备注2,备注3,备注4 FROM TZJZMX WHERE 类别='" & Combo1.Text & "' AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')  ORDER BY 日期,序号"
Adodc1.Refresh
End If
Call OutadodcToExcel2(VSFlexGrid1, 6, 7, "特种记账 日期范围： " + Text1.Text + "--" + Text2.Text)
Adodc1.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')  ORDER BY 序号 DESC"
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
On Error Resume Next
Adodc1.Database.Execute "UPDATE TZJZMX SET 序号='0' WHERE 序号=NULL"
Adodc1.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')  ORDER BY 序号 DESC"
Adodc1.Refresh
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
End Sub


Private Sub dataCombo1_LostFocus(Index As Integer)
Select Case Index
       Case 5
       If DataCombo1(5).Text <> "0" Then
       DataCombo1(6).Text = 0
       End If
       Case 6
       If DataCombo1(6).Text <> "0" Then
       DataCombo1(5).Text = 0
       End If
End Select
End Sub

Private Sub DTPicker1_Change()
Text1.Text = DTPicker1.Value
Text1.SetFocus
End Sub

Private Sub DTPicker1_CloseUp()
Text1.Text = DTPicker1.Value
Text1.SetFocus
End Sub
Private Sub DTPicker2_Change()
Text2.Text = DTPicker2.Value
Text2.SetFocus
End Sub

Private Sub DTPicker2_CloseUp()
Text2.Text = DTPicker2.Value
Text2.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = Date
Text2.Text = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc1.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY 序号 DESC"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc2.RecordSource = "SELECT MAX(序号) FROM TZJZMX "
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc3.RecordSource = "SELECT MC FROM TZZL GROUP BY MC "
Adodc3.Refresh
Combo1.Text = ""
For i = 0 To 11
DataCombo1(i).Text = ""
Next
DataCombo1(1).Text = Date
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 1500
VSFlexGrid1.ColWidth(5) = 1500
VSFlexGrid1.ColWidth(8) = 1500
VSFlexGrid1.ColWidth(13) = 0
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
       KMMC = 1
       Formw6.Show
       Case 8
       DataCombo1(11).Enabled = True
End Select
End Sub
Private Sub Label1_DBLClick(Index As Integer)
Select Case Index
       Case 8
       DataCombo1(11).Enabled = False
End Select
End Sub

Private Sub Label3_Click()
KMMC = 6
Formw6.Show
End Sub

Private Sub Label4_Click()
Formw1129.Show
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
If Adodc1.Recordset.Fields(12) = "已" Or Adodc1.Recordset.Fields(13) = "已" Then Exit Sub
For i = 0 To Adodc1.Recordset.Fields.Count - 1
DataCombo1(i).Text = Adodc1.Recordset.Fields(i)
Next
End Sub
