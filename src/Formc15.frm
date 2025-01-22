VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formc15 
   BackColor       =   &H00C0E0FF&
   Caption         =   "成品发货"
   ClientHeight    =   10215
   ClientLeft      =   -435
   ClientTop       =   3810
   ClientWidth     =   15960
   Icon            =   "Formc15.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15735
   ScaleWidth      =   28680
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc31 
      Height          =   495
      Left            =   2520
      Top             =   11400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Adodc31"
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
   Begin MSAdodcLib.Adodc Adodc30 
      Height          =   495
      Left            =   2520
      Top             =   10680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "Adodc30"
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
   Begin MSAdodcLib.Adodc Adodc29 
      Height          =   495
      Left            =   2640
      Top             =   9960
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "Adodc29"
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
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0FF&
      Caption         =   "应收款刷新"
      Height          =   375
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text15 
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
      Left            =   18120
      TabIndex        =   108
      Text            =   "Text13"
      Top             =   4440
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc28 
      Height          =   375
      Left            =   14760
      Top             =   11040
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
      Caption         =   "Adodc28"
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
   Begin MSAdodcLib.Adodc Adodc27 
      Height          =   495
      Left            =   14760
      Top             =   10200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Adodc27"
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
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0FF&
      Caption         =   "锅单打印"
      Height          =   615
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   15960
      TabIndex        =   104
      Text            =   "Text14"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   14640
      TabIndex        =   103
      Text            =   "Text13"
      Top             =   3120
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "光坯打印"
      Height          =   615
      Left            =   8400
      TabIndex        =   98
      Top             =   4320
      Width           =   2775
      Begin VB.OptionButton Option7 
         BackColor       =   &H0080FF80&
         Caption         =   "否"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   100
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H0080FF80&
         Caption         =   "是"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   99
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   120
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formc15.frx":0A7A
      Height          =   3495
      Left            =   240
      TabIndex        =   61
      Top             =   5040
      Width           =   21375
      _cx             =   37703
      _cy             =   6165
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
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0FF&
      Caption         =   "调捡打印"
      Height          =   375
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全部删除"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   360
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo19 
      Bindings        =   "Formc15.frx":0A8F
      Height          =   330
      Left            =   4440
      TabIndex        =   91
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "xm"
      Text            =   "DataCombo19"
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8280
      TabIndex        =   87
      Text            =   "Text7"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "码单信息"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   11280
      TabIndex        =   83
      Text            =   "Text11"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   82
      Text            =   "Text11"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   11160
      TabIndex        =   80
      Text            =   "Text11"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   11160
      TabIndex        =   79
      Text            =   "Text11"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "核算方式"
      Height          =   855
      Left            =   240
      TabIndex        =   74
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton Option6 
         BackColor       =   &H0000C0C0&
         Caption         =   "匹数"
         Height          =   375
         Left            =   2040
         TabIndex        =   86
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H0000C0C0&
         Caption         =   "光坯"
         Height          =   375
         Left            =   1080
         TabIndex        =   76
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0000C0C0&
         Caption         =   "毛坯"
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Text            =   "Text10"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "出库单"
      Height          =   255
      Left            =   4800
      TabIndex        =   70
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "结算单"
      Height          =   255
      Left            =   6000
      TabIndex        =   69
      Top             =   4440
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "欠款条"
      Height          =   255
      Left            =   7200
      TabIndex        =   68
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "出库查询"
      Height          =   375
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6120
      TabIndex        =   66
      Text            =   "Text9"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细进度"
      Height          =   375
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2400
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Formc15.frx":0AA4
      Left            =   9840
      List            =   "Formc15.frx":0AB7
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc26 
      Height          =   330
      Left            =   11160
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
      Caption         =   "Adodc26"
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
      Left            =   10680
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
   Begin MSDataListLib.DataCombo DataCombo17 
      Height          =   330
      Left            =   7560
      TabIndex        =   60
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo17"
   End
   Begin MSDataListLib.DataCombo DataCombo16 
      Height          =   330
      Left            =   7560
      TabIndex        =   59
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo16"
   End
   Begin MSDataListLib.DataCombo DataCombo14 
      Height          =   330
      Left            =   10680
      TabIndex        =   58
      Top             =   6960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo14"
   End
   Begin MSDataListLib.DataCombo DataCombo13 
      Height          =   330
      Left            =   1560
      TabIndex        =   57
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo13"
   End
   Begin MSDataListLib.DataCombo DataCombo12 
      Height          =   330
      Left            =   17520
      TabIndex        =   11
      Top             =   3120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo12"
   End
   Begin MSDataListLib.DataCombo DataCombo11 
      Height          =   330
      Left            =   3000
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo11"
   End
   Begin MSDataListLib.DataCombo DataCombo9 
      Height          =   330
      Left            =   1560
      TabIndex        =   55
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo9"
   End
   Begin MSDataListLib.DataCombo DataCombo8 
      Height          =   330
      Left            =   9840
      TabIndex        =   54
      Top             =   6960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo8"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Height          =   330
      Left            =   12600
      TabIndex        =   13
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo7"
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Height          =   330
      Left            =   9840
      TabIndex        =   53
      Top             =   7320
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo6"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Left            =   9360
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo5"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formc15.frx":0AE3
      Height          =   330
      Left            =   6120
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "YS"
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formc15.frx":0AF8
      Height          =   330
      Left            =   4440
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "PM"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc15.frx":0B0D
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   11280
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Text            =   "Text7"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3480
      Top             =   0
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   16320
      TabIndex        =   49
      Text            =   "Text6"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "下一单据号"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   12720
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "入库"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   21720
      MultiLine       =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5880
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询功能"
      Height          =   735
      Left            =   1080
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   13095
      Begin MSDataListLib.DataCombo DataCombo10 
         Bindings        =   "Formc15.frx":0B22
         Height          =   330
         Left            =   1080
         TabIndex        =   56
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "简称"
         Text            =   "DataCombo10"
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7200
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "客户查询"
         Height          =   375
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4920
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421376
         CalendarTrailingForeColor=   255
         Format          =   328925185
         CurrentDate     =   39181
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7200
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8421376
         CalendarTrailingForeColor=   255
         Format          =   328925185
         CurrentDate     =   39181
      End
      Begin VB.Line Line1 
         X1              =   6480
         X2              =   7080
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "客户名称"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "时间范围："
         Height          =   375
         Left            =   3960
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印确定"
      Height          =   615
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   330
      Left            =   12720
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   328925185
      CurrentDate     =   39181
   End
   Begin MSAdodcLib.Adodc Adodc25 
      Height          =   330
      Left            =   8160
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
      Left            =   8400
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
      Left            =   8520
      Top             =   10200
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
      Left            =   8640
      Top             =   10320
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
      Left            =   8880
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
      Left            =   8880
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc18 
      Height          =   330
      Left            =   9960
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
      Left            =   11400
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
      Left            =   9600
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
      Left            =   9240
      Top             =   10320
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
      Left            =   10200
      Top             =   10200
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
      Left            =   9240
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
      Left            =   9960
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
      Left            =   9720
      Top             =   10200
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
      Left            =   10440
      Top             =   10200
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
      Left            =   9720
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
      Left            =   11520
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
      Left            =   11520
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
      Left            =   9000
      Top             =   10200
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
      Left            =   9360
      Top             =   10320
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
      Left            =   9240
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
      Left            =   9480
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
      Height          =   375
      Left            =   9600
      Top             =   10200
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
   Begin MSDataListLib.DataCombo DataCombo15 
      Height          =   330
      Left            =   8640
      TabIndex        =   71
      Top             =   10560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo12"
   End
   Begin MSDataListLib.DataCombo DataCombo18 
      Height          =   330
      Left            =   11280
      TabIndex        =   89
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo18"
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1800
      TabIndex        =   92
      Top             =   4440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formc15.frx":0B37
      Height          =   735
      Left            =   240
      TabIndex        =   95
      Top             =   8640
      Width           =   18375
      _cx             =   32411
      _cy             =   1296
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
      FormatString    =   $"Formc15.frx":0B4D
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
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   375
      Left            =   21000
      TabIndex        =   106
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   329121793
      CurrentDate     =   45352
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   21000
      TabIndex        =   107
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   329121793
      CurrentDate     =   45352
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "累计欠款"
      Height          =   495
      Left            =   16920
      TabIndex        =   109
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "附加费金额"
      Height          =   255
      Index           =   11
      Left            =   15960
      TabIndex        =   102
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "附加费单价"
      Height          =   255
      Index           =   10
      Left            =   14640
      TabIndex        =   101
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   16200
      TabIndex        =   93
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "负责"
      Height          =   375
      Left            =   4440
      TabIndex        =   90
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "米数"
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   88
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单位"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   11280
      TabIndex        =   84
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "成分"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   81
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   11160
      TabIndex        =   78
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单据"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   11160
      TabIndex        =   77
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光坯重量"
      Height          =   255
      Index           =   1
      Left            =   9360
      TabIndex        =   73
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "收费明细"
      Height          =   255
      Left            =   8640
      TabIndex        =   72
      Top             =   10200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "色号"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   65
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF80&
      Caption         =   "清除"
      Height          =   255
      Left            =   1200
      TabIndex        =   63
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "提货:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   52
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "匹数"
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   51
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   50
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号:"
      Height          =   375
      Index           =   1
      Left            =   15120
      TabIndex        =   48
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "收费明细"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   9840
      TabIndex        =   47
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label13"
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
      Left            =   16320
      TabIndex        =   46
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "当前单号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   15120
      TabIndex        =   45
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "顺序号："
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   255
      Left            =   17520
      TabIndex        =   40
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   39
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   38
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "合计总额（元）："
      Height          =   855
      Left            =   21720
      TabIndex        =   36
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "金额（元）"
      Height          =   255
      Index           =   0
      Left            =   12720
      TabIndex        =   35
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   33
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "品名"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   32
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   31
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "锅号"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   30
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "毛坯重量"
      Height          =   255
      Index           =   5
      Left            =   9360
      TabIndex        =   29
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "日期"
      Height          =   255
      Index           =   0
      Left            =   12720
      TabIndex        =   28
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单价"
      Height          =   255
      Index           =   6
      Left            =   11400
      TabIndex        =   27
      Top             =   2760
      Width           =   975
   End
End
Attribute VB_Name = "Formc15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public gygh As String
Dim JDBAR As Integer
Dim hs, ZS, ps As Integer: Dim fhsl As Single: Dim je As Single: Dim zhy As Integer
Dim cdbhf As Integer
Private Declare Function PRINTDLG Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG) As Long
' 声明 PrintDlg 函数，用于调用 Windows API 打印对话框

Private Type PRINTDLG
    lStructSize As Long            ' 结构体大小
    hwndOwner As Long              ' 拥有窗口的句柄
    hDevMode As Long               ' 打印设备模式的句柄
    hDevNames As Long              ' 打印设备名称的句柄
    hDC As Long                    ' 打印设备的设备上下文句柄
    Flags As Long                  ' 标志位，用于控制对话框的行为
    nFromPage As Integer           ' 打印起始页码
    nToPage As Integer             ' 打印结束页码
    nMinPage As Integer            ' 最小页码
    nMaxPage As Integer            ' 最大页码
    nCopies As Integer             ' 打印份数
    hInstance As Long              ' 应用程序实例句柄
    lCustData As Long              ' 用户自定义数据
    lpfnPrintHook As Long          ' 打印对话框挂钩过程指针
    lpfnSetupHook As Long          ' 打印设置对话框挂钩过程指针
    lpPrintTemplateName As String  ' 打印对话框模板名称
    lpSetupTemplateName As String  ' 打印设置对话框模板名称
    hPrintTemplate As Long         ' 打印对话框模板句柄
    hSetupTemplate As Long         ' 打印设置对话框模板句柄
End Type
' 定义 PRINTDLG 结构体，用于传递和接收打印对话框信息

Private Const PD_RETURNDC = &H100          ' 返回设备上下文标志
Private Const PD_NOSELECTION = &H4         ' 禁止选择标志
Private Const PD_NOPAGENUMS = &H8          ' 禁止页码选择标志
Private Const PD_PRINTSETUP = &H40         ' 显示打印设置对话框标志
' 定义常量，用于设置 PRINTDLG 结构体中的 flags 字段

Private Function ShowPrintDialog() As Boolean
    ' 定义一个私有函数 ShowPrintDialog，返回布尔值，用于显示打印对话框
    Dim pd As PRINTDLG
    ' 声明一个 PRINTDLG 类型的变量 pd，用于存储打印对话框的信息
    Dim result As Long
    ' 声明一个长整型变量 result，用于存储 PrintDlg 函数的返回值

    pd.lStructSize = Len(pd)
    ' 设置 pd 结构体的 lStructSize 字段为 pd 的长度
    pd.hwndOwner = 0
    ' 设置 pd 结构体的 hwndOwner 字段为 0，表示没有拥有窗口
    pd.Flags = PD_RETURNDC Or PD_NOSELECTION Or PD_NOPAGENUMS Or PD_PRINTSETUP
    ' 设置 pd 结构体的 flags 字段，组合多个标志位以控制对话框的行为

    result = PRINTDLG(pd)
    ' 调用 PrintDlg 函数，传递 pd 结构体，并将返回值赋给 result 变量

    If result <> 0 Then
        ' 如果 PrintDlg 函数返回值不为 0，表示用户点击了“打印”
        ShowPrintDialog = True
        ' 将函数的返回值设置为 True
    Else
        ' 如果 PrintDlg 函数返回值为 0，表示用户取消了打印
        ShowPrintDialog = False
        ' 将函数的返回值设置为 False
    End If
End Function

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Command10_Click()
Formc142.Text1 = DataCombo4
Formc142.Text2(0) = Label13.Caption
Formc142.Show
End Sub


Private Sub Command11_Click()
On Error Resume Next
If MsgBox("确定修改司机吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "update jgmx set 跟单='" & DataCombo17 & "' where 单号='" & Label13.Caption & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc9.Refresh
End Sub

Private Sub Command12_Click()
Formc34.DataCombo1(4).Text = DataCombo4.Text
Formc34.Show
End Sub

Private Sub Command13_Click()
On Error Resume Next
If MsgBox("确认全部删除吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "update kpd set zt='发货删除', FH='N' WHERE 锅号 in(select distinct 锅号 from jgmx where 单号='" & Label13.Caption & "')"
sql2 = "delete from jgmx where 单号='" & Label13.Caption & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc9.Refresh
Call Command15_Click
End Sub

Private Sub Command14_Click()
    Dim formattedDate As String
    formattedDate = Format(CDate(Text5.Text), "yymmdd") ' 格式化为当天的日期 yymmdd
    
    ' 连接数据库
    Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    
    ' 查询当天的最大单据号
    Adodc16.RecordSource = "SELECT MAX(CAST(RIGHT(单号, 3) AS INT)) " & _
                           "FROM JGMX " & _
                           "WHERE LEFT(单号, 1) = '" & yhdm & "' " & _
                           "AND SUBSTRING(单号, 2, 6) = '" & formattedDate & "'"
    Adodc16.Refresh
    Debug.Print "SQL 查询语句: " & Adodc16.RecordSource
    Debug.Print "查询结果是否为空: " & Adodc16.Recordset.EOF
    
    Dim L As Integer
    If Not Adodc16.Recordset.EOF And Not IsNull(Adodc16.Recordset.Fields(0).value) Then
        ' 提取当天的最大序号，并递增
        L = Val(Adodc16.Recordset.Fields(0).value) + 1
    Else
        ' 如果当天没有记录，从 1 开始
        L = 1
    End If
    Debug.Print "增加后的序号: " & L
    
    ' 生成新的单据号，确保序号始终是 3 位
    Dim newNumber As String
    newNumber = yhdm & formattedDate & Format(L, "000")
    Debug.Print "生成的单号: " & newNumber

    ' 更新单据号
    Label13.Caption = newNumber
    
    ' 刷新下一个单据的相关数据
    Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
    Adodc9.Refresh
    Debug.Print "查询单号相关数据: " & Label13.Caption

    Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
    Adodc21.Refresh

    If Adodc21.Recordset.EOF Then
        DataCombo9.Text = 1
        DataCombo13.Text = 1
    Else
        DataCombo9.Text = Adodc21.Recordset.Fields(0) + 1
        DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
    End If

    ' 重置其他控件
    Text5.Text = Format(Now, "yyyy-MM-dd HH:mm:ss")
    DTPicker3.value = Now
    Text7.Text = ""
    DataCombo2.Text = ""
    DataCombo3.Text = ""
    DataCombo5.Text = ""
    DataCombo6.Text = ""
    Text8.Text = ""
    DataCombo7.Text = ""
    DataCombo11.Text = ""
    DataCombo12.Text = ""
    DataCombo16.Text = ""
    DataCombo4.SetFocus
End Sub





Private Sub Command15_Click()
On Error Resume Next

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh

Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh

If Adodc21.Recordset.EOF Then
DataCombo13.Text = 1
Else
DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
End If
DataCombo7.Enabled = False

Adodc23.RecordSource = "SELECT SUM(ISNULL(匹数,0)) as 合计匹数,SUM(ISNULL(数量,0)) as 毛坯合计,SUM(ISNULL(光坯,0)) as 光坯合计 FROM JGMX WHERE 单号='" & Label13.Caption & "' and (加工类别='成品布' or 加工类别='染色费' or 加工类别='定型费' or 加工类别='不收费' or 加工类别='外印花' or 加工类别='只磨毛')"
Adodc23.Refresh

 If VSFlexGrid1.Rows > 1 Then
        ' 允许通过键盘和鼠标编辑
        VSFlexGrid1.Editable = flexEDKbdMouse
        ' 将第0列的复选框状态设置为1（选中）
        VSFlexGrid1.Cell(flexcpChecked, 1, CheckboxColumnIndex, VSFlexGrid1.Rows - 1, CheckboxColumnIndex) = 1
    End If
Call gssx
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command17_Click()
    Dim selectedGuoHao As String
    Dim weight As Double
    Dim count As Double

    ' 显示打印机选择界面
    If Not ShowPrintDialog() Then
        Exit Sub ' 用户取消了打印，退出子程序
    End If

    With VSFlexGrid1
        Debug.Print "VSFlexGrid1.Rows: " & .Rows ' 打印总行数
        For i = 1 To .Rows - 1 ' 排除最后一行
            ' 检查复选框是否选中
            If .Cell(flexcpChecked, i, CheckboxColumnIndex) = 1 Then
                selectedGuoHao = .TextMatrix(i, 4) ' 获取第四列的锅号
                weight = CDbl(.TextMatrix(i, 5)) ' 获取第五列的重量，并转换为Double类型
                count = CDbl(.TextMatrix(i, 18)) ' 获取第十三列的匹数，并转换为Double类型

                Call lcd22f3(Adodc27, Adodc28, selectedGuoHao, weight, count) ' 将重量和匹数传递给lcd22f3子程序

                Debug.Print "Row " & i & " processed" ' 调试信息
            End If
        Next i
    End With
End Sub



Private Sub Command18_Click()
On Error Resume Next
Command18.Enabled = False
FP = CDate(DTPicker5.value) + 1
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "YSHZCX('" & DTPicker4.value & "','" & DTPicker5.value & "','" & FP & "','" & yhm & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel
Command17.Enabled = True
'''MsgBox ("刷新成功！")
Adodc29.RecordSource = "SELECT * FROM YSKZCX where 类别='" & yhm & "' order by 客户"
Adodc29.Refresh
Adodc30.RecordSource = "SELECT round(sum(isnull(欠款,0)),2) as 合计欠款 FROM YSKZCX where 类别='" & yhm & "'"
Adodc30.Refresh
Adodc31.RecordSource = "SELECT round(sum(isnull(欠款,0)),2) as 合计欠款 FROM jgzcx where 客户= '" & DataCombo1 & "'"
Adodc31.Refresh
Text15.Text = Adodc31.Recordset.Fields(0)
End Sub

Private Sub Command5_Click()
If Adodc9.Recordset.EOF Then
MsgBox ("无记录，不能打印")
Exit Sub
End If
JDBAR = 10
Timer1.Enabled = True
ProgressBar1.Visible = True
End Sub


Private Sub Command7_Click()
Forma172.DataCombo1 = DataCombo1
Forma172.Show
End Sub

Private Sub Command8_Click()
wwdm = 4
Formc344.Check2(4).value = 1
Formc344.Show
End Sub

Private Sub Command9_Click()
'On Error Resume Next

If Option2.value = True Then
Timer1.Enabled = False
ProgressBar1.Visible = False

Adodc15.RecordSource = "select isnull(count(顺序号),0) from jgmx where 单号='" & Label13.Caption & "'"
Adodc15.Refresh
If Not Adodc15.Recordset.EOF Then
hs = Adodc15.Recordset.Fields(0)
If hs > 0 Then
If hs / 10 = Int(hs / 10) Then
ZS = hs / 10
Else
ZS = Int(hs / 10) + 1
End If
zhy = ZS - 1
End If


For i = 0 To ZS - 1
If i = zhy Then
If Option7(0).value = True Then
Call CPCKTZD(Adodc13, Adodc12, Label13.Caption, i * 10 + 1, i * 10 + 10, i + 1, zhy + 1)
Else
Call CPCKTZDGP(Adodc13, Adodc12, Adodc27, Adodc29, Label13.Caption, i * 10 + 1, i * 10 + 10, i + 1, zhy + 1)
End If
Else
If Option7(0).value = True Then
Call CPCKTZDF(Adodc13, Adodc12, Label13.Caption, i * 10 + 1, i * 10 + 10, i + 1, zhy + 1)
Else
Call CPCKTZDFGP(Adodc13, Adodc12, Label13.Caption, i * 10 + 1, i * 10 + 10, i + 1, zhy + 1)
End If
End If
Next

sql1 = "update jgmx set dy='2' where 单号='" & Label13.Caption & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Exit Sub
End If
End If

Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc16.RecordSource = "SELECT * FROM JGMX where 日期='" & Text5.Text & "'"
Adodc16.Refresh

If Adodc16.Recordset.EOF Then
Label13.Caption = Format(CDate(Text5.Text), "yymmdd") + "001"
Else
Adodc16.RecordSource = "SELECT max(right(单号,len(单号)-6)) FROM JGMX where 日期='" & Text5.Text & "'"
Adodc16.Refresh
L = Val(Adodc16.Recordset.Fields(0))
If Len(L + 1) = 1 Then
Label13.Caption = Format(CDate(Text5.Text), "yymmdd") + "00" + Trim(L + 1)
End If
If Len(L + 1) = 2 Then
Label13.Caption = Format(CDate(Text5.Text), "yymmdd") + "0" + Trim(L + 1)
End If
If Len(L + 1) = 3 Then
Label13.Caption = Format(CDate(Text5.Text), "yymmdd") + Trim(L + 1)
End If
End If


Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh

Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh

If Adodc21.Recordset.EOF Then
DataCombo9.Text = 1
DataCombo13.Text = 1
Else
DataCombo9.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
End If

Text7.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""  ''''''
Text8.Text = ""
DataCombo7.Text = ""
DataCombo11.Text = ""
DataCombo12.Text = ""
DataCombo16.Text = ""
DataCombo4.SetFocus

End Sub



Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub dataCombo10_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub dataCombo11_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub dataCombo12_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub dataCombo14_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub



Private Sub dataCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub



Private Sub dataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub


Private Sub DataCombo4_Change()
'On Error Resume Next
If DataCombo4.Text = "" Then Exit Sub

If InStr(DataCombo4, "J") > 0 Or InStr(DataCombo4, "j") > 0 Then
DataCombo4 = Mid(DataCombo4, 1, Len(DataCombo4) - 1)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 编号 from kpd  WHERE 锅号='" & DataCombo4 & "' order by 编号"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "MPbmzk('" & DataCombo4 & "','" & Adodc2.Recordset.Fields(0) & "')"   ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc2.Recordset.MoveNext
Loop
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Adodc25.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc25.RecordSource = "select distinct '00000000' as 单据,'1' as 序号,客户,缸号,'' as 单号序号,款号,锅号,品名,颜色+色号 as 色别,库存数量,库存匹数 as 发货匹数,光坯数量,收费项目,单价,(case when isnull(核算方式,'')='毛坯' then round(毛坯重量*isnull(单价,0),2) when isnull(核算方式,'')='光坯' then round(光坯数量*isnull(单价,0),2) when isnull(核算方式,'')='匹数' then round(光坯匹数*isnull(单价,0),2) end) as 合计金额,核算方式,库存匹数,备注,图案,日期,附加费单价,(case when isnull(核算方式,'')='毛坯' then round(毛坯重量*isnull(附加费单价,0),2) when isnull(核算方式,'')='光坯' then round(光坯数量*isnull(附加费单价,0),2) when isnull(核算方式,'')='匹数' then round(光坯匹数*isnull(附加费单价,0),2) end) as 附加费金额 from v_kpd_fh  WHERE 锅号='" & DataCombo4 & "' and 库存匹数>0 order by 缸号"
Adodc25.Refresh
If Not Adodc25.Recordset.EOF Then
Adodc25.Recordset.MoveFirst
Do While Not Adodc25.Recordset.EOF
'Adodc28.RecordSource = "SELECT 客户,欠费上限,欠费 FROM yj_qfts WHERE 客户 = Adodc25.Recordset.Fields(2)  "
'Adodc28.Refresh
'If Not Adodc28.Recordset.EOF Then
'If Val(Adodc28.Recordset.Fields(2)) >= Val(Adodc28.Recordset.Fields(1)) Then ' 检查是否>=条件成立
 '       MsgBox ("客户欠费超出预警，不能开发货单")
 '       Exit Sub ' 如果条件成立，弹出提示信息并退出子程序
 '   End If
'End If


Adodc21.RecordSource = "SELECT 顺序号,加工单位 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh

If Not Adodc21.Recordset.EOF Then
ID = Adodc21.Recordset.Fields(0) + 1
SXH = Adodc21.Recordset.Fields(0) + 1
S17 = Adodc21.Recordset.Fields(1)
Else
ID = 1
SXH = 1
S17 = ""
End If
                                                                                        
S1 = Adodc25.Recordset.Fields(0)
S2 = Adodc25.Recordset.Fields(1)
s3 = Adodc25.Recordset.Fields(2)
s4 = Adodc25.Recordset.Fields(3)
s5 = Adodc25.Recordset.Fields(4)
s6 = Adodc25.Recordset.Fields(5)
s7 = Adodc25.Recordset.Fields(6)
s8 = Adodc25.Recordset.Fields(7)
s9 = Adodc25.Recordset.Fields(8)
s10 = Val(Adodc25.Recordset.Fields(9))  '毛坯数量
S11 = Val(Adodc25.Recordset.Fields(10))  ''匹数
S12 = Val(Adodc25.Recordset.Fields(11))  '''光坯数量
S13 = Adodc25.Recordset.Fields(12)      '''加工类别
' 获取 Adodc25 记录的相关字段值
S14 = IIf(IsNull(Adodc25.Recordset.Fields(13).value), 0, Adodc25.Recordset.Fields(13).value)  ' 单价，处理可能的空值
s15 = IIf(IsNull(Adodc25.Recordset.Fields(14).value), 0, Adodc25.Recordset.Fields(14).value)  ' 金额，处理可能的空值
s18 = Val(Adodc25.Recordset.Fields(16))   ''库存匹数
S12 = Format(S12 / s18 * S11, "#0.0")
s19 = IIf(IsNull(Adodc25.Recordset.Fields(20).value), 0, Adodc25.Recordset.Fields(20).value)  ' 附加费单价，处理可能的空值
s20 = IIf(IsNull(Adodc25.Recordset.Fields(21).value), 0, Adodc25.Recordset.Fields(21).value)  ' 附加费金额，处理可能的空值
' 获取 Adodc25 记录的单价字段值
Dim 单价 As Double
If IsNull(Adodc25.Recordset.Fields(13).value) Then
    单价 = 0   ' 如果单价为 NULL，则设置为 0
Else
    单价 = Adodc25.Recordset.Fields(13).value
End If

Dim 附加费单价 As Double
If IsNull(Adodc25.Recordset.Fields(20).value) Then
    附加费单价 = 0   ' 如果单价为 NULL，则设置为 0
Else
    附加费单价 = Adodc25.Recordset.Fields(20).value
End If
' 根据选项判断核算方式
If Option4.value = True Then
    s16 = "毛坯"       ''核算
    s15 = 单价 * Val(Adodc25.Recordset.Fields(9)) ''金额
    s20 = 附加费单价 * Val(Adodc25.Recordset.Fields(9)) ''附加费金额
End If
If Option5.value = True Then
    s16 = "光坯"       ''核算
    s15 = 单价 * Val(Adodc25.Recordset.Fields(11)) ''金额
    s20 = 附加费单价 * Val(Adodc25.Recordset.Fields(11)) ''附加费金额
End If
If Option6.value = True Then
    s16 = "匹数"       ''核算
    s15 = 单价 * Val(Adodc25.Recordset.Fields(10)) ''金额
    s20 = 附加费单价 * Val(Adodc25.Recordset.Fields(10)) ''金额
End If


'If Adodc25.Recordset.Fields(18) = "" Then
's18 = Adodc25.Recordset.Fields(17)  ''''备注
'Else
s18 = Adodc25.Recordset.Fields(18)    ''''图案
'End If
S21 = s15 + s20
If S17 <> s3 And S17 <> "" Then
MsgBox ("不是一个客户的，不能开发货单")
Exit Sub
End If

  
sql1 = "INSERT INTO dbo.jgmx(入库单据,入库序号,加工单位,缸号,ip,和约号,锅号,品名,颜色,数量,匹数,光坯,加工类别,单价,金额,核算,负责,单号,日期,顺序号,单位,备注,跟单,附加费单价,附加费金额,总金额) Values('" & S1 & "','" & S2 & "','" & s3 & "','" & s4 & "','" & s5 & "','" & s6 & "','" & s7 & "','" & s8 & "','" & s9 & "','" & s10 & "','" & S11 & "','" & S12 & "','" & S13 & "','" & S14 & "','" & s15 & "','" & s16 & "','" & DataCombo19 & "','" & Label13.Caption & "','" & Text5 & "','" & SXH & "','公斤','" & s18 & "','" & DataCombo17 & "','" & s19 & "','" & s20 & "','" & S21 & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

sql2 = "update dbo.kpd set FH=convert(nvarchar ,'" & Now & "',120),zt='已发货' WHERE 锅号='" & Adodc25.Recordset.Fields(6) & "' and 编号='" & Adodc25.Recordset.Fields(3) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic


Adodc25.Recordset.MoveNext

Loop
DataCombo4 = ""
End If
End If

Call Command15_Click
End Sub

Private Sub dataCombo4_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DataCombo5_Change()
If Option4.value = True Then
DataCombo7.Text = Format(Val(DataCombo5.Text) * Val(Text8.Text), "#0.00")
End If
End Sub

Private Sub dataCombo5_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub


Private Sub Command1_Click()
On Error Resume Next

If DataCombo1.Text = "" Then
MsgBox ("请输入客户!")
Exit Sub
End If

If DataCombo19.Text = "" Then
MsgBox ("请输入负责!")
Exit Sub
End If


If Label13.Caption = "" Then
MsgBox ("请确认单据号")
Exit Sub
End If

If Text5.Text = "" Then
MsgBox ("请确认日期")
Exit Sub
End If

If Len(Label13.Caption) <> 10 Then
MsgBox ("单据号不正确")
Exit Sub
End If

If Option4.value = True Then
jsfs = "毛坯"
End If

If Option5.value = True Then
jsfs = "光坯"
End If

If Option6.value = True Then
jsfs = "米数"
End If

If DataCombo5.Text = "" Then DataCombo5.Text = 0
If DataCombo6.Text = "" Then DataCombo6.Text = 0
If DataCombo7.Text = "" Then DataCombo7.Text = 0
If Text8.Text = "" Then Text8.Text = 0
If Text12.Text = "" Then Text12.Text = 0

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select * from jgmx where 锅号='" & DataCombo4.Text & "' and 品名='" & DataCombo2.Text & "' and 颜色='" & DataCombo3 & "' and 加工类别='" & Combo1.Text & "'"
Adodc3.Refresh

If Not Adodc3.Recordset.EOF Then
If MsgBox("此锅号已开，请确认，是否继续？", vbYesNo) = vbNo Then Exit Sub
End If


Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cpfhcz1('" & DataCombo1.Text & "','" & DataCombo2.Text & "','" & DataCombo3.Text & "','" & DataCombo4.Text & "','" & DataCombo5.Text & "','" & Text8.Text & "','" & DataCombo7.Text & "','" & Text5.Text & "','" & DataCombo9.Text & "','" & Text9 & "','" & DataCombo11.Text & "','" & DataCombo12.Text & "','" & DataCombo13.Text & "','" & Label13.Caption & "','" & Combo1.Text & "','1','1','" & Text7.Text & "','" & DataCombo16.Text & "','','','','" & DataCombo17.Text & "',null,'" & DataCombo15 & "','" & Text10 & "','" & Text11(2) & "','" & Text11(0) & "','" & Text11(1) & "','" & Text11(3) & "','" & Text12 & "','" & jsfs & "','" & DataCombo19 & "')" ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel

Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh

Adodc23.RecordSource = "SELECT SUM(ISNULL(匹数,0)) as 合计匹数,SUM(ISNULL(数量,0)) as 毛坯合计,SUM(ISNULL(光坯,0)) as 光坯合计 FROM JGMX WHERE 单号='" & Label13.Caption & "' and (加工类别='成品布' or 加工类别='染色费' or 加工类别='定型费' or 加工类别='不收费' or 加工类别='外印花' or 加工类别='只磨毛')"
Adodc23.Refresh

Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh

DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1

Text7.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""  ''''''
Text8.Text = ""
DataCombo7.Text = ""
DataCombo11.Text = ""
DataCombo12.Text = ""
DataCombo16.Text = ""
DataCombo4.SetFocus
Call gssx


End Sub

Private Sub Command2_Click()
On Error Resume Next

If DataCombo1.Text = "" Then
MsgBox ("请输入客户!")
Exit Sub
End If

If DataCombo19.Text = "" Then
MsgBox ("请输入负责!")
Exit Sub
End If


If Text6.Text = "" Then
MsgBox ("请确认单据号")
Exit Sub
End If

If Text5.Text = "" Then
MsgBox ("请确认日期")
Exit Sub
End If

If Len(Label13.Caption) <> 10 Then
MsgBox ("单据号不正确")
Exit Sub
End If

If Option4.value = True Then
jsfs = "毛坯"
End If

If Option5.value = True Then
jsfs = "光坯"
End If

If Option6.value = True Then
jsfs = "米数"
End If


If Adodc9.Recordset.EOF Then Exit Sub

If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
Adodc9.Recordset.Fields(0) = DataCombo1.Text
Adodc9.Recordset.Fields(1) = DataCombo2.Text
Adodc9.Recordset.Fields(2) = DataCombo3.Text
Adodc9.Recordset.Fields(3) = DataCombo4.Text
Adodc9.Recordset.Fields(4) = DataCombo5.Text
Adodc9.Recordset.Fields(5) = Text8.Text
Adodc9.Recordset.Fields(6) = DataCombo7.Text
Adodc9.Recordset.Fields(7) = Text5.Text
Adodc9.Recordset.Fields(8) = DataCombo9.Text
Adodc9.Recordset.Fields(9) = Text9.Text
Adodc9.Recordset.Fields(10) = DataCombo11.Text
Adodc9.Recordset.Fields(11) = DataCombo12.Text
Adodc9.Recordset.Fields(12) = DataCombo13.Text
Adodc9.Recordset.Fields(13) = Text6.Text
Adodc9.Recordset.Fields(14) = Combo1.Text
Adodc9.Recordset.Fields(15) = "1"
Adodc9.Recordset.Fields(16) = "1"
Adodc9.Recordset.Fields(17) = Text7.Text
Adodc9.Recordset.Fields(18) = DataCombo16.Text
Adodc9.Recordset.Fields(22) = DataCombo17.Text
Adodc9.Recordset.Fields(24) = DataCombo15.Text
Adodc9.Recordset.Fields(25) = Val(Text10)    '''''光坯重量
Adodc9.Recordset.Fields(26) = Text11(0)   '''
Adodc9.Recordset.Fields(27) = Text11(1)
Adodc9.Recordset.Fields(28) = Text11(2)
Adodc9.Recordset.Fields(29) = Text11(3)
Adodc9.Recordset.Fields(30) = Text12
Adodc9.Recordset.Fields(34) = DataCombo19
Adodc9.Recordset.Fields(36) = Text13
Adodc9.Recordset.Fields(37) = Text14
Adodc9.Recordset.Fields(38) = Val(Text14.Text) + Val(DataCombo7.Text)
Adodc9.Recordset.Update
Adodc9.Refresh
DataCombo7.Enabled = False

Adodc23.RecordSource = "SELECT SUM(ISNULL(匹数,0)) as 合计匹数,SUM(ISNULL(数量,0)) as 毛坯合计,SUM(ISNULL(光坯,0)) as 光坯合计 FROM JGMX WHERE 单号='" & Label13.Caption & "' and (加工类别='成品布' or 加工类别='染色费' or 加工类别='定型费' or 加工类别='不收费' or 加工类别='外印花' or 加工类别='只磨毛')"
Adodc23.Refresh

Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh
sql1 = "update dbo.kpd set FH=convert(nvarchar ,'" & Now & "',120),zt='已发货' WHERE 锅号= '" & Adodc9.Recordset.Fields(3) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

MsgBox ("修改成功！")
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

Text7.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = "" '''''
Text8.Text = ""
DataCombo7.Text = ""
Text13.Text = ""
Text14.Text = ""
DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo14.Text = ""
DataCombo11.Text = ""
DataCombo12.Text = ""
DataCombo16.Text = ""
DataCombo4.SetFocus
Call gssx

End Sub

Private Sub Command3_Click()
On Error Resume Next
If DataCombo10.Text = "" Then
Adodc9.RecordSource = "select *  from jgmx where  日期 between  '" & Text2.Text & "'  and  '" & Text3.Text & "'  order by 日期,单号,顺序号"
Adodc9.Refresh

Adodc7.RecordSource = "select sum(金额)  from jgmx where  日期 between  '" & Text2.Text & "'  and  '" & Text3.Text & "' "
Adodc7.Refresh

If Adodc7.Recordset.EOF Then
Exit Sub
Else
Text4.Text = Format(Adodc7.Recordset.Fields(0), "###0.00")
End If

Else
Adodc9.RecordSource = "select *  from jgmx where 加工单位='" & DataCombo10.Text & " ' AND 日期 between  '" & Text2.Text & "'  and  '" & Text3.Text & "'  order by 日期,单号,顺序号"
Adodc9.Refresh
Adodc7.RecordSource = "select sum(金额)  from jgmx where  加工单位='" & DataCombo10.Text & " ' and  日期 between  '" & Text2.Text & "'  and  '" & Text3.Text & "' "
Adodc7.Refresh
If Adodc7.Recordset.EOF Then
Exit Sub
Else
Text4.Text = Format(Adodc7.Recordset.Fields(0), "###0.00")
End If
End If

End Sub

Private Sub Command4_Click()
On Error Resume Next

If Adodc9.Recordset.EOF Then Exit Sub
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc9.Recordset.Delete
Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh
MsgBox ("删除成功！")
Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh
DataCombo7.Enabled = False

If Not Adodc9.Recordset.EOF Then
Adodc9.Recordset.MoveFirst
i = 1
Do While Not Adodc9.Recordset.EOF
Adodc9.Recordset.Fields(12) = i
Adodc9.Recordset.Update
Adodc9.Recordset.MoveNext
i = i + 1
Loop
End If
Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh
Adodc23.RecordSource = "SELECT SUM(ISNULL(匹数,0)) as 合计匹数,SUM(ISNULL(数量,0)) as 毛坯合计,SUM(ISNULL(光坯,0)) as 光坯合计 FROM JGMX WHERE 单号='" & Label13.Caption & "' and (加工类别='成品布' or 加工类别='染色费' or 加工类别='定型费' or 加工类别='不收费' or 加工类别='外印花' or 加工类别='只磨毛')"
Adodc23.Refresh


sql1 = "update kpd set zt='发货删除',FH='N' WHERE 锅号='" & DataCombo4 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
Text7.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = "" '''''
Text8.Text = ""
DataCombo7.Text = ""

DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo14.Text = ""
DataCombo11.Text = ""
DataCombo12.Text = ""
DataCombo16.Text = ""
DataCombo4.SetFocus
Call gssx

End Sub

Private Sub Command6_Click()
Unload Me
End Sub



Private Sub DataCombo6_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub


Private Sub DTPicker1_Change()
Text2.Text = DTPicker1.value
End Sub
Private Sub DTPicker1_CloseUp()
Text2.Text = DTPicker1.value
Text2.SetFocus
End Sub

Private Sub DTPicker2_Change()
Text3.Text = DTPicker2.value
End Sub
Private Sub DTPicker2_CloseUp()
Text3.Text = DTPicker2.value
Text3.SetFocus
End Sub

Private Sub DTPicker3_Change()
' 设置带时间的格式
DataCombo8.Text = Format(DTPicker3.value, "yyyy-MM-dd HH:mm:ss")
Text5.Text = Format(DTPicker3.value, "yyyy-MM-dd HH:mm:ss")
End Sub

Private Sub DTPicker3_CloseUp()
' 设置带时间的格式
DataCombo8.Text = Format(DTPicker3.value, "yyyy-MM-dd HH:mm:ss")
Text5.Text = Format(DTPicker3.value, "yyyy-MM-dd HH:mm:ss")
Text5.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
cdbhf = cdbh
Option1.value = True
DataCombo17.Text = ""
' 初始化 Text5 和 DTPicker3 为当前日期时间
Text5.Text = Format(Now, "yyyy-MM-dd HH:mm:ss")
DTPicker3.value = Now
Text1.Text = ""
ProgressBar1.Visible = False
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""
DataCombo7.Text = ""
DataCombo8.Text = ""
DataCombo9.Text = ""
DataCombo9.Text = ""
DataCombo13.Text = ""
DataCombo8.Text = Date
DataCombo10.Text = ""
DataCombo11.Text = ""
DataCombo12.Text = ""
DataCombo14.Text = ""
DataCombo16.Text = ""
DataCombo18.Text = ""
DataCombo19.Text = "朱泽"
Text7.Text = ""
Text12.Text = ""
DataCombo15.Text = ""
Text10.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Option4.value = True
Option2.value = True
For i = 0 To 3
Text11(i) = ""
Next
Text11(3) = "公斤"
Option7(1).value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc28.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc28.RecordSource = "SELECT * from yj_qfts   order by 客户"
Adodc28.Refresh
Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc23.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc27.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc28.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc29.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc30.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc31.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select 简称 from khZL where ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc5.Refresh

Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select xm  from fzr group by xm"
Adodc4.Refresh

 Dim formattedDate As String
    formattedDate = Format(CDate(Text5.Text), "yymmdd") ' 格式化为当天的日期 yymmdd
    
    ' 连接数据库
    Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    
    ' 查询当天的最大单据号
    Adodc16.RecordSource = "SELECT MAX(CAST(RIGHT(单号, 3) AS INT)) " & _
                           "FROM JGMX " & _
                           "WHERE LEFT(单号, 1) = '" & yhdm & "' " & _
                           "AND SUBSTRING(单号, 2, 6) = '" & formattedDate & "'"
    Adodc16.Refresh
    Debug.Print "SQL 查询语句: " & Adodc16.RecordSource
    Debug.Print "查询结果是否为空: " & Adodc16.Recordset.EOF
    
    Dim L As Integer
    If Not Adodc16.Recordset.EOF And Not IsNull(Adodc16.Recordset.Fields(0).value) Then
        ' 提取当天的最大序号，并递增
        L = Val(Adodc16.Recordset.Fields(0).value) + 1
    Else
        ' 如果当天没有记录，从 1 开始
        L = 1
    End If
    Debug.Print "增加后的序号: " & L
    
    ' 生成新的单据号，确保序号始终是 3 位
    Dim newNumber As String
    newNumber = yhdm & formattedDate & Format(L, "000")
    Debug.Print "生成的单号: " & newNumber

    ' 更新单据号
    Label13.Caption = newNumber
    

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh


Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh

If Adodc21.Recordset.EOF Then
DataCombo9.Text = 1
DataCombo13.Text = 1
Else
DataCombo9.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
End If

Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

VSFlexGrid1.ColWidth(0) = 500
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1500
VSFlexGrid1.ColWidth(3) = 1500
VSFlexGrid1.ColWidth(4) = 1200
VSFlexGrid1.ColWidth(8) = 1000
VSFlexGrid1.ColWidth(6) = 1000
VSFlexGrid1.ColWidth(7) = 1000
VSFlexGrid1.ColWidth(9) = 0
VSFlexGrid1.ColWidth(10) = 0
VSFlexGrid1.ColWidth(11) = 1000
VSFlexGrid1.ColWidth(12) = 1000
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(16) = 0      ''''库类
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(20) = 0      '''提取
VSFlexGrid1.ColWidth(21) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0      '''跟单
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0      ''zl
VSFlexGrid1.ColWidth(27) = 0
VSFlexGrid1.ColWidth(28) = 0
VSFlexGrid1.ColWidth(29) = 0
VSFlexGrid1.ColWidth(31) = 0       '''米数
VSFlexGrid1.ColWidth(33) = 0
VSFlexGrid1.ColWidth(34) = 0

'vSFlexGrid1.ColWidth(6) = 0
'vSFlexGrid1.ColWidth(7) = 0
VSFlexGrid7.ColWidth(0) = 100

Combo1.Text = ""

DataCombo8.Text = Text5.Text
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker4.value = "2024-9-28"
DTPicker5.value = Date
Text2.Text = DTPicker1.value
Text3.Text = DTPicker2.value
Timer1.Enabled = False

If Len(yhdm) <> 1 Then
MsgBox ("这个账户不合适进入这个界面")
Command1.Enabled = False
Command15.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command13.Enabled = False
Command14.Enabled = False
Command10.Enabled = False
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

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
       DataCombo7.Enabled = True
       Case 4
       'If DataCombo19 = "" Then
       'MsgBox ("请选择发货人和提货车号")
      'Exit Sub
       'End If
       fhxz = 15
       Formc146.Text1.Text = DataCombo4.Text
       Formc146.Show
End Select
End Sub

Private Sub Label10_Click()
   beizhu = 55
   Forma112.Show
End Sub

Private Sub Label13_Change()
On Error Resume Next

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh

Adodc23.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc23.RecordSource = "SELECT SUM(ISNULL(匹数,0)) as 合计匹数,SUM(ISNULL(数量,0)) as 毛坯合计,SUM(ISNULL(光坯,0)) as 光坯合计 FROM JGMX WHERE 单号='" & Label13.Caption & "' and (加工类别='成品布' or 加工类别='染色费' or 加工类别='定型费' or 加工类别='不收费' or 加工类别='外印花' or 加工类别='只磨毛')"
Adodc23.Refresh

Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh


If Adodc21.Recordset.EOF Then
DataCombo9.Text = 1
DataCombo13.Text = 1
Else
DataCombo9.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
End If
 If VSFlexGrid1.Rows > 1 Then
        ' 允许通过键盘和鼠标编辑
        VSFlexGrid1.Editable = flexEDKbdMouse
        ' 将第0列的复选框状态设置为1（选中）
        VSFlexGrid1.Cell(flexcpChecked, 1, CheckboxColumnIndex, VSFlexGrid1.Rows - 1, CheckboxColumnIndex) = 1
    End If
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

End Sub

Private Sub Label13_dblClick()
On Error Resume Next
Label13.Caption = InputBox("请输入单号", , Label13.Caption)
Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh

Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh

If Adodc21.Recordset.EOF Then
DataCombo9.Text = 1
DataCombo13.Text = 1
Else
DataCombo9.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
End If

End Sub



Private Sub Label16_Click()
DataCombo4.Text = ""
End Sub

Private Sub Label6_Click()
AC1.DT1.Source = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
AC1.DT1.Refresh
AC1.Show
End Sub

Private Sub Option4_Click()
DataCombo7.Text = Format(Val(DataCombo5.Text) * Val(Text8.Text), "#0.00")
Text11(3) = "公斤"
End Sub

Private Sub Option5_Click()
DataCombo7.Text = Format(Val(Text10) * Val(Text8.Text), "#0.00")
Text11(3) = "公斤"
End Sub

Private Sub Option6_Click()
DataCombo7.Text = Format(Val(Text7) * Val(Text8.Text), "#0.00")
'Text11(3) = "米"
End Sub

Private Sub Text1_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'  group by 简称"
Adodc5.Refresh
End Sub

Private Sub Text10_Change()
If Option5.value = True Then
DataCombo7.Text = Format(Val(Text10) * Val(Text8.Text), "#0.00")
End If
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Text12_Change()
If Option6.value = True Then
DataCombo7.Text = Format(Val(Text12) * Val(Text8.Text), "#0.00")
End If
End Sub

Private Sub Text13_Change()
If Option4.value = True Then
Text14.Text = Format(Val(DataCombo5.Text) * Val(Text13.Text), "#0.00")
End If
If Option5.value = True Then
Text14.Text = Format(Val(Text10) * Val(Text13.Text), "#0.00")
End If
If Option6.value = True Then
Text14.Text = Format(Val(Text7) * Val(Text13.Text), "#0.00")
End If
End Sub

Private Sub Text7_Change()
If Option6.value = True Then
DataCombo7.Text = Format(Val(Text7) * Val(Text8.Text), "#0.00")
End If
End Sub

Private Sub Timer2_Timer()    '''''刷新
If fhsx >= 3 Then
Call Command15_Click
Timer2.Enabled = False
End If
fhsx = fhsx + 1
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc9.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc9.Recordset.MoveFirst
Adodc9.Recordset.Move rs - 1
  
If Adodc9.Recordset.Fields(21) = "是" Then     '''''审核则不能修改删除
  Command4.Enabled = False
  Command2.Enabled = False
  Command1.Enabled = False
  
Else
  
     DataCombo1.Text = Adodc9.Recordset.Fields(0)
     DataCombo2.Text = Adodc9.Recordset.Fields(1)
     DataCombo3.Text = Adodc9.Recordset.Fields(2)
     DataCombo4.Text = Adodc9.Recordset.Fields(3)
     DataCombo5.Text = Adodc9.Recordset.Fields(4)
      Text8.Text = Adodc9.Recordset.Fields(5)
      DataCombo7.Text = Adodc9.Recordset.Fields(6)
     DataCombo8.Text = Adodc9.Recordset.Fields(7)
     DataCombo11.Text = Adodc9.Recordset.Fields(10)
     DataCombo12.Text = Adodc9.Recordset.Fields(11)
     DataCombo13.Text = Adodc9.Recordset.Fields(12)   ''顺序号
     DataCombo9.Text = Adodc9.Recordset.Fields(8)    '''ip  单号序号
       DataCombo14.Text = Adodc9.Recordset.Fields(9)
       Text6.Text = Adodc9.Recordset.Fields(13)
       Combo1.Text = Adodc9.Recordset.Fields(14)
       KL = Adodc9.Recordset.Fields(15)
       Text5.Text = Adodc9.Recordset.Fields(7)
       DTPicker3.value = Adodc9.Recordset.Fields(7)
     Text7.Text = Adodc9.Recordset.Fields(17)
     Text9.Text = Adodc9.Recordset.Fields(9)
     DataCombo15.Text = Adodc9.Recordset.Fields(24)
     DataCombo16.Text = Adodc9.Recordset.Fields(18)
     DataCombo17.Text = Adodc9.Recordset.Fields(22)
     Text10 = Adodc9.Recordset.Fields(25)
     Text11(0) = Adodc9.Recordset.Fields(26)
     Text11(1) = Adodc9.Recordset.Fields(27)
     Text11(2) = Adodc9.Recordset.Fields(28)
     Text11(3) = Adodc9.Recordset.Fields(29)
     Text12 = Adodc9.Recordset.Fields(30)
     Text13 = Adodc9.Recordset.Fields(36)
     Text14 = Adodc9.Recordset.Fields(37)
  Command4.Enabled = True
  Command2.Enabled = True
  Command1.Enabled = False
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text8_Change()
If Option4.value = True Then
DataCombo7.Text = Format(Val(DataCombo5.Text) * Val(Text8.Text), "#0.00")
End If
If Option5.value = True Then
DataCombo7.Text = Format(Val(Text10) * Val(Text8.Text), "#0.00")
End If
If Option6.value = True Then
DataCombo7.Text = Format(Val(Text7) * Val(Text8.Text), "#0.00")
End If
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Timer1_Timer()   '''''打印
If JDBAR = 100 Then   '''''''进度条到头就打印
Call Command9_Click
Exit Sub
End If
ProgressBar1.value = JDBAR
JDBAR = JDBAR + 10
End Sub
Private Sub gssx()
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If
End Sub

