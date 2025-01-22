VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw160 
   BackColor       =   &H00C0E0FF&
   Caption         =   "查询打印"
   ClientHeight    =   11955
   ClientLeft      =   -870
   ClientTop       =   1110
   ClientWidth     =   15960
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11955
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   42
      Text            =   "Text6"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "结账"
      Height          =   495
      Left            =   13320
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询方式"
      Height          =   610
      Left            =   7920
      TabIndex        =   36
      Top             =   1560
      Width           =   3250
      Begin VB.OptionButton Option1 
         Caption         =   "汇总"
         Height          =   250
         Index           =   1
         Left            =   1800
         TabIndex        =   38
         Top             =   240
         Width           =   1210
      End
      Begin VB.OptionButton Option1 
         Caption         =   "明细"
         Height          =   250
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   1210
      End
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1320
      TabIndex        =   34
      Text            =   "Text5"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "发货查询"
      Height          =   495
      Left            =   12000
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2880
      Width           =   1330
   End
   Begin VB.ComboBox Combo1111 
      Appearance      =   0  'Flat
      Height          =   300
      ItemData        =   "Formw160.frx":0000
      Left            =   6240
      List            =   "Formw160.frx":0002
      Style           =   1  'Simple Combo
      TabIndex        =   32
      Text            =   "Combo1111"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   30
      Text            =   "Text6"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   1215
      Left            =   14880
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      Width           =   1330
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "对账查询"
      Height          =   495
      Left            =   12000
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      Width           =   1330
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   2520
      Width           =   2050
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导出EXCEL"
      Height          =   495
      Left            =   13320
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1330
      Left            =   7920
      TabIndex        =   2
      Top             =   2160
      Width           =   3975
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "备注"
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "不收费"
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   35
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单据"
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   31
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单号"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "负数"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "款号"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色号"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Text            =   "Text6"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Text            =   "Text6"
      Top             =   2160
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo DataCombo11 
      Height          =   330
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo11"
   End
   Begin MSDataListLib.DataCombo DataCombo10 
      Bindings        =   "Formw160.frx":0004
      Height          =   330
      Left            =   1920
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo10"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw160.frx":0019
      Height          =   9615
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   27855
      _cx             =   49133
      _cy             =   16960
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
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
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw160.frx":002E
      Height          =   330
      Left            =   8160
      TabIndex        =   17
      Top             =   -120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "布类"
      BoundColumn     =   "布类"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw160.frx":0042
      Height          =   330
      Left            =   8400
      TabIndex        =   18
      Top             =   -120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "客户"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   0
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   328859649
      CurrentDate     =   39181
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   328859649
      CurrentDate     =   39181
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5880
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5760
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5400
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   5400
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   5640
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   5160
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   375
      Left            =   5160
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   375
      Left            =   5400
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   375
      Left            =   4920
      Top             =   10440
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
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
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
      TabIndex        =   40
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单据"
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
      Left            =   4680
      TabIndex        =   29
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客 户 加 工 发 货 明 细 查 询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   28
      Top             =   600
      Width           =   6855
   End
   Begin VB.Line Line2 
      X1              =   11880
      X2              =   12720
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "查询打印"
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
      Left            =   12720
      TabIndex        =   27
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户"
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
      Left            =   840
      TabIndex        =   26
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "时间范围"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   840
      TabIndex        =   25
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   2
      Left            =   840
      TabIndex        =   24
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
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
      Left            =   840
      TabIndex        =   23
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   4
      Left            =   4680
      TabIndex        =   22
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号"
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
      Left            =   4680
      TabIndex        =   21
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "Formw160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rr As Integer: Dim KLB1, KLB2 As String '''''库类变量
Dim rs As Single: Dim ll, XXG As Integer  ''''XXG选项变量
Dim mm As Date: Dim ML As Date: Dim YB As Integer ''''''打印变量
Dim c, r As Integer
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim zf As Long
Dim yf As Long
Dim sf As Long
Dim xf As Long
Dim dzgs As Integer    '''''对账格式
Dim sb As RECT
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long




Private Sub Command1_Click()
On Error Resume Next
sql1 = ""
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")

If Check2(4).value = 1 Then
sql1 = sql1 + "日期 between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "加工单位 like '%'+'" & DataCombo10.Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "颜色 like '%'+'" & DataCombo11.Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "和约号 like '%'+'" & Text6.Text & "'+'%' and "
End If

If Check2(0).value = 1 Then
sql1 = sql1 + "锅号='" & Text1.Text & "' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "计划号='" & Text2.Text & "' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "单价<0 and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "单号='" & Text3.Text & "' and "
End If

If Check2(8).value = 1 Then
sql1 = sql1 + "单价=0 and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
If MsgBox("确定结账吗？", vbYesNo) = vbNo Then Exit Sub

    sql2 = "update jgmx set 成分='" & Now & "' where (" + sql1 + ")"
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic
    
 MsgBox ("结账成功！")
End Sub

Private Sub Command13_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "客户: " + DataCombo10 + "  对账查找")
End Sub

Private Sub Command3_Click()
dzgs = 1
On Error Resume Next
sql1 = ""
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")

If Check2(4).value = 1 Then
sql1 = sql1 + "日期 between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "加工单位 ='" & DataCombo10.Text & "' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "颜色 like '%'+'" & DataCombo11.Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "和约号 like '%'+'" & Text6.Text & "'+'%' and "
End If

If Check2(0).value = 1 Then
sql1 = sql1 + "锅号='" & Text1.Text & "' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "计划号='" & Text2.Text & "' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "单价<0 and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "单号='" & Text3.Text & "' and "
End If

If Check2(8).value = 1 Then
sql1 = sql1 + "单价=0 and "
End If

If Check2(9).value = 1 Then
sql1 = sql1 + "备注 like '%'+'" & Text4 & "'+'%' and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

If Option1(0).value = True Then
Adodc1.RecordSource = "select 日期,加工单位,品名,颜色,锅号,匹数,数量 as 投坯数量,单价,金额,备注,单号 as 单据号,和约号 as 款号,结账,克重,织厂,业务  from v_jgmxdz where (" + sql1 + ")  order by 日期,单号,顺序号"
Adodc1.Refresh


VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 0, 6, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 7, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 9, , &HC0C0&
End If

If Option1(1).value = True Then
Adodc1.RecordSource = "select 加工单位,单号 as 单据号,round(sum(匹数),0) as 毛坯匹数,round(sum(数量),2) as 毛坯重量,round(sum(isnull(光坯,0)),2) as 光坯重量,round(sum(isnull(金额,0)),2) as 合计金额 from v_jgmxdz where (" + sql1 + ")  group by 加工单位,单号 order by 单号"
Adodc1.Refresh
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTCount, 0, 2, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 3, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 4, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 5, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 6, , &HC0C0&
End If

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
If i / 2 = Int(i / 2) Then
VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &H80000005
Else
VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &H8000000F
End If
Next
End If

End Sub

Private Sub Command4_Click()
On Error Resume Next
dzgs = 0

sql1 = ""
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")

If Check2(4).value = 1 Then
sql1 = sql1 + "日期 between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "加工单位= '" & DataCombo10.Text & "' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "颜色 like '%'+'" & DataCombo11.Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "和约号 like '%'+'" & Text6.Text & "'+'%' and "
End If

If Check2(0).value = 1 Then
sql1 = sql1 + "锅号='" & Text1.Text & "' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "计划号='" & Text2.Text & "' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "单价<0 and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "单号='" & Text3.Text & "' and "
End If

If Check2(8).value = 1 Then
sql1 = sql1 + "单价=0 and "
End If

If Check2(9).value = 1 Then
sql1 = sql1 + "备注 like '%'+'" & Text4 & "'+'%' and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

If Option1(0).value = True Then
Adodc1.RecordSource = "select 加工单位,日期,单号 as 单据,品名,和约号 as 款号,颜色,色号,锅号,匹数,数量,单价,金额,光坯,备注, 织厂,加工类别,核算,计划号 as 订单号,ip as 订单序号,顺序号,结账,克重,业务   from v_jgmxdz where (" + sql1 + ")  order by 日期,单号,顺序号"
Adodc1.Refresh
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 0, 9, , &HC0C0&   '''第9列合计
VSFlexGrid1.Subtotal flexSTSum, 0, 10, , &HC0C0&  '''第10列合计
VSFlexGrid1.Subtotal flexSTSum, 0, 12, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 13, , &HC0C0&
End If

If Option1(1).value = True Then
Adodc1.RecordSource = "select 加工单位,单号 as 单据,round(sum(匹数),0) as 毛坯匹数,round(sum(数量),2) as 毛坯重量,round(sum(isnull(光坯,0)),2) as 光坯重量,round(sum(isnull(金额,0)),2) as 合计金额 from v_jgmxdz where (" + sql1 + ")  group by 加工单位,单号 order by 单号"
Adodc1.Refresh
VSFlexGrid1.SubtotalPosition = flexSTBelow  '''这行代码将分组统计结果显示在数据网格控件的底部
VSFlexGrid1.Subtotal flexSTCount, 0, 2, , &HC0C0&  ''''统计单据数量,这行代码统计第2列的单据数量，并将结果显示在数据网格控件的底部。&HC0C0&是一个16进制的颜色值，表示分组统计结果的背景颜色。
VSFlexGrid1.Subtotal flexSTSum, 0, 3, , &HC0C0& '''''这行代码统计第3列的总和，并将结果显示在数据网格控件的底部。
VSFlexGrid1.Subtotal flexSTSum, 0, 4, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 5, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 6, , &HC0C0&
End If

If VSFlexGrid1.Rows > 1 Then  ''如果VSFlexGrid1的行数大于1
For i = 1 To VSFlexGrid1.Rows - 1   ' 遍历VSFlexGrid1中的每一行，除了最后一行
VSFlexGrid1.RowHeight(i) = 600   ' 设置当前行的高度为600
If i / 2 = Int(i / 2) Then   ' 判断当前行的行号是否为偶数
VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &H80000005   ' 如果是偶数，则将当前行的单元格背景颜色设置为深灰色
Else
VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &H8000000F   ' 如果是奇数，则将当前行的单元格背景颜色设置为浅灰色
End If
Next  ' 继续处理下一行
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub dataCombo10_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub DataCombo4_Change()
If DataCombo10.Text = "" Then
MsgBox ("请先选择加工单位！")
Exit Sub
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc1.RecordSource = "select *  from v_jgmxdz where 加工单位= '" & DataCombo10.Text & " ' and 开票时间= '" & DataCombo4.Text & " '  order by 顺序号 "
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
VSFlexGrid1.TextMatrix(0, 0) = "记录号"
Text4.Text = Format(0, "###0.00")
Exit Sub
Else
Adodc1.Recordset.MoveLast
VSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To Adodc1.Recordset.RecordCount
VSFlexGrid1.TextMatrix(i, 0) = i
Next
End If
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select sum(金额)  from v_jgmxdz where  开票时间= '" & DataCombo4.Text & " '   "
Adodc7.Refresh
Text4.Text = Format(Adodc7.Recordset.Fields(0), "###0.00")

End Sub

Private Sub dataCombo4_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub


Private Sub Form_Load()

On Error Resume Next

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
DataCombo12.Text = ""
Text5.Text = ""
Check2(4).value = 1
Combo1 = ""
KK = ""
DTPicker1.value = Date
DTPicker2.value = Date
Option1(0).value = True
DataCombo11.Text = ""
DataCombo10.Text = ""
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.CommandTimeout = 10000
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Text4.Text = Format(0, "###0.00")

DataCombo10.TabIndex = 0

VSFlexGrid1.ColWidth(0) = 200


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

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Text5_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select  简称 from KHZL where 代码  like '%'+'" & Text5 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by 简称 "
Adodc5.Refresh
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub
Private Sub VSFlexGrid1_dblClick()
If wwdm = 4 Then
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
Formc15.Label13.Caption = Adodc1.Recordset.Fields(2)
Unload Me
End If
End Sub
Private Sub MSF()
On Error Resume Next
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行

If (c = 10 Or c = 11 Or c = 15 Or c = 14) And InStr(yhm, cw) > 0 And dzgs = 0 And Len(.TextMatrix(r, 19)) < 8 Then
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
Call MSF
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move r - 1
m1 = Adodc1.Recordset.Fields(c - 2)
m2 = Adodc1.Recordset.Fields(c + 1)
Adodc1.Recordset.Fields(c - 1) = Combo1111.Text
blfs = Adodc1.Recordset.Fields(14)
If c = 10 And Adodc1.Recordset.Fields(14) = "毛坯" Then
Adodc1.Recordset.Fields(c) = Format(Adodc1.Recordset.Fields(c - 2) * Val(Combo1111.Text), "#0.00")
End If
If c = 10 And Adodc1.Recordset.Fields(14) = "光坯" Then
Adodc1.Recordset.Fields(c) = Format(Adodc1.Recordset.Fields(c + 1) * Val(Combo1111.Text), "#0.00")
End If
Adodc1.Recordset.Update


    VSFlexGrid1.Text = Combo1111.Text
    If c = 10 And blfs = "毛坯" Then
    VSFlexGrid1.TextMatrix(r, c + 1) = Format(m1 * Val(Combo1111.Text), "#0.00")
    End If
    If c = 10 And blfs = "光坯" Then
    VSFlexGrid1.TextMatrix(r, c + 1) = Format(m2 * Val(Combo1111.Text), "#0.00")
    End If
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub

Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If InStr(yhm, "cw") = 0 And InStr(yhm, "root") = 0 And dzgs <> 0 Then Exit Sub

With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行

If Len(.TextMatrix(r, 19)) > 8 Then Exit Sub
End With

'''''加工单位,日期,单号 as 单据,品名,和约号,颜色,锅号,匹数,数量,单价,金额,光坯,备注,加工类别,核算,订单号，订单序号,顺序号

S1 = VSFlexGrid1.TextMatrix(r, 1)   '''加工单位
S2 = VSFlexGrid1.TextMatrix(r, 2)   '''日期
s3 = VSFlexGrid1.TextMatrix(r, 3)   ''单号
s4 = VSFlexGrid1.TextMatrix(r, 4)   ''品名
s5 = VSFlexGrid1.TextMatrix(r, 5)   '''和约号
s6 = VSFlexGrid1.TextMatrix(r, 6)   '''颜色
s7 = VSFlexGrid1.TextMatrix(r, 7)   '''锅号
s8 = VSFlexGrid1.TextMatrix(r, 8)  '''匹数
s9 = VSFlexGrid1.TextMatrix(r, 9)  '''数量
s10 = VSFlexGrid1.TextMatrix(r, 10)  '''单价
S11 = VSFlexGrid1.TextMatrix(r, 11) '''金额
S12 = VSFlexGrid1.TextMatrix(r, 12) '''光坯
S13 = VSFlexGrid1.TextMatrix(r, 13) '''备注
S14 = VSFlexGrid1.TextMatrix(r, 14) '''加工类别
s15 = VSFlexGrid1.TextMatrix(r, 15) '''核算
s16 = VSFlexGrid1.TextMatrix(r, 16) '''订单
S17 = VSFlexGrid1.TextMatrix(r, 17) '''订单序号
s18 = VSFlexGrid1.TextMatrix(r, 18) '''顺序号
s19 = s18                           ''''顺序号 删除用
s10 = 0 '''单价
S11 = 0 '''金额

Adodc3.RecordSource = "select max(顺序号) from jgmx where 单号='" & s3 & "'"
Adodc3.Refresh
If IsNull(Adodc3.Recordset.Fields(0)) Then
s18 = 1
Else
s18 = Adodc3.Recordset.Fields(0) + 1               ''''''''''顺序号
End If


    If Button = 2 And c = 7 And dzgs = 0 Then
    If MsgBox("确定复制这行的信息吗？" + s7 + S14, vbYesNo) = vbNo Then '''PopupMenu mnu_manager  '这是在窗体中设置的一个顶级菜单名称
    Exit Sub
    Else
    sql2 = "insert into jgmx(加工单位,日期,单号,品名,和约号 as 款号,颜色,锅号,匹数,数量,单价,金额,光坯,备注,加工类别,核算,计划号,ip,顺序号) values('" & S1 & "','" & S2 & "','" & s3 & "','" & s4 & "','" & s5 & "','" & s6 & "','" & s7 & "','" & s8 & "','" & s9 & "','" & s10 & "','" & S11 & "','" & S12 & "','" & S13 & "','" & S14 & "','" & s15 & "','" & s16 & "','" & S17 & "','" & s18 & "')"
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic
    End If
    Call Command4_Click
    End If

    If Button = 2 And c = 1 And dzgs = 0 Then
    If MsgBox("确定删除这行的信息吗？" + s7 + S14, vbYesNo) = vbNo Then '''PopupMenu mnu_manager  '这是在窗体中设置的一个顶级菜单名称
    Exit Sub
    Else
    sql2 = "delete from jgmx where 单号='" & s3 & "' and 顺序号='" & s19 & "' and 锅号='" & s7 & "'"
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic
    End If
    Call Command4_Click
    End If

End Sub
