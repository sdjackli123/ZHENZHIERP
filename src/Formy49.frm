VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formy49 
   BackColor       =   &H00C0E0FF&
   Caption         =   "配料查询"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   5160
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8760
      TabIndex        =   43
      Text            =   "Text3"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8520
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   360
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
      ItemData        =   "Formy49.frx":0000
      Left            =   7800
      List            =   "Formy49.frx":0002
      TabIndex        =   38
      Text            =   "Combo1"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "汇总打印"
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
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   28
      Text            =   "Text2"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   27
      Text            =   "Text8"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   24
      Text            =   "Text8"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   23
      Text            =   "Text8"
      Top             =   1440
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10440
      Top             =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1695
      Left            =   10200
      TabIndex        =   3
      Top             =   360
      Width           =   5295
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "未审核"
         Height          =   255
         Index           =   12
         Left            =   4080
         TabIndex        =   48
         Top             =   1080
         Visible         =   0   'False
         Width           =   850
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "已审核"
         Height          =   255
         Index           =   11
         Left            =   2760
         TabIndex        =   47
         Top             =   1080
         Visible         =   0   'False
         Width           =   850
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色别"
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   46
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "机台"
         Height          =   255
         Index           =   9
         Left            =   4080
         TabIndex        =   39
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "称量日期"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   1090
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "已称量"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "未称量"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "配料日期"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "料单"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "审核日期"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "名称"
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
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
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
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
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
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
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   5880
      Top             =   10320
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
      Left            =   7440
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
      Left            =   6720
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
      Left            =   6720
      Top             =   10560
      Visible         =   0   'False
      Width           =   3255
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
      Height          =   495
      Left            =   6840
      Top             =   10440
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
      Bindings        =   "Formy49.frx":0004
      Height          =   11175
      Left            =   480
      TabIndex        =   8
      Top             =   2160
      Width           =   17175
      _cx             =   30295
      _cy             =   19711
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
      Left            =   4200
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formy49.frx":0019
      Height          =   330
      Left            =   5760
      TabIndex        =   10
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "名称"
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   720
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   1560
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formy49.frx":002E
      Height          =   330
      Left            =   5760
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Formy49.frx":0043
      Height          =   330
      Left            =   4200
      TabIndex        =   14
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "染化助库名"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formy49.frx":0058
      Height          =   12495
      Left            =   17880
      TabIndex        =   15
      Top             =   840
      Width           =   9255
      _cx             =   16325
      _cy             =   22040
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
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "Formy49.frx":006D
      Height          =   330
      Left            =   7800
      TabIndex        =   44
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo2"
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
      Height          =   375
      Index           =   3
      Left            =   8760
      TabIndex        =   42
      Top             =   1200
      Width           =   735
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
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   40
      Top             =   360
      Width           =   1695
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
      Height          =   375
      Index           =   3
      Left            =   7800
      TabIndex        =   37
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   32
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   31
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   30
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   29
      Top             =   720
      Width           =   255
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
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   21
      Top             =   1200
      Width           =   1695
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
      Height          =   375
      Index           =   5
      Left            =   4200
      TabIndex        =   20
      Top             =   360
      Width           =   1215
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
      Left            =   480
      TabIndex        =   19
      Top             =   1200
      Width           =   1335
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
      Left            =   480
      TabIndex        =   18
      Top             =   360
      Width           =   1335
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
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   17
      Top             =   360
      Width           =   1695
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
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Formy49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim plshsx, yssj As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call OutadodcToExcel(VSFlexGrid1, 8, "领料车间" + DataCombo1.Text)
End Sub


Private Sub Command3_Click()
Call OutadodcToExcel2(VSFlexGrid2, 8, 9, "配料统计")
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

If Check2(9).value = 1 Then
sql1 = sql1 + "机台 like '%'+'" & Combo2 & "'+'%' and "
End If

If Check2(10).value = 1 Then
sql1 = sql1 + "颜色 like '%'+'" & Text3 & "'+'%' and "
End If

If Check2(11).value = 1 Then
sql1 = sql1 + "isnull(审核,'')='是' and "
End If

If Check2(12).value = 1 Then
sql1 = sql1 + "isnull(审核,'')<>'是' and "
End If


If Check2(5).value = 1 Then
sql1 = sql1 + "isnull(称量标记,'')<>'Y' and "
End If

If Check2(8).value = 1 Then
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text2(0) + ":" + Text2(1) + ":" + Text2(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text8(0) + ":" + Text8(1) + ":" + Text8(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "称量日期 between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc1.RecordSource = "SELECT * FROM v_pld_kpd_cx WHERE (" + sql1 + ") order by 料单编号,工序名称,次序号"
Adodc1.Refresh
Adodc5.RecordSource = "SELECT 编码,染化助名称, round(sum(isnull(配料用量,0)),3) as 配料用量, round(sum(isnull(实际称量,0)),3) as 实际称量, 单价, round(sum(isnull(合计金额,0)) , 2) as 合计金额 FROM v_pld_kpd_cx WHERE (" + sql1 + ") group by 编码,染化助名称,单价 order by 编码"
Adodc5.Refresh

With VSFlexGrid1
    .WordWrap = True
    .MergeCells = 2
    .MergeCol(1 - 5) = True '是否上下列合并
End With

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 400
If i / 2 = Int(i / 2) Then
VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &H8000000F
Else
VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &H80000005
End If
Next
End If

If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 600
Next
End If
VSFlexGrid1.ColFormat(11) = "#0.####"
VSFlexGrid1.ColFormat(12) = "#0.####"
VSFlexGrid2.ColFormat(3) = "#0.####"
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(1) = 600
VSFlexGrid2.ColWidth(2) = 2500
VSFlexGrid2.ColWidth(3) = 1500
VSFlexGrid2.ColWidth(4) = 1500
VSFlexGrid2.ColWidth(5) = 800
VSFlexGrid2.ColWidth(6) = 1500
VSFlexGrid2.SubtotalPosition = flexSTBelow
VSFlexGrid2.Subtotal flexSTSum, 0, 3, , &HC0C0&   '''第9列合计
VSFlexGrid2.Subtotal flexSTSum, 0, 4, , &HC0C0&  '''第10列合计
VSFlexGrid2.Subtotal flexSTSum, 0, 6, , &HC0C0&

End Sub


Private Sub Form_Load()
Combo2 = ""
DTPicker1.value = Date
DTPicker2.value = Date
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
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select distinct 染化助库名 from rhzh"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc6.Refresh

Adodc1.CommandTimeout = 10000
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
