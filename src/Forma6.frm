VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma6 
   BackColor       =   &H00C0E0FF&
   Caption         =   "毛坯入库详情"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   10440
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17400
      TabIndex        =   41
      Text            =   "Text5"
      Top             =   10560
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   370
      Left            =   10080
      TabIndex        =   37
      Text            =   "Text4"
      Top             =   960
      Width           =   1090
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   375
      Left            =   5400
      Top             =   8760
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6720
      TabIndex        =   26
      Text            =   "Text3"
      Top             =   600
      Width           =   495
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   2640
      Style           =   1  'Simple Combo
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   5280
      Top             =   8760
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
      Left            =   5160
      Top             =   8760
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
      Left            =   5160
      Top             =   8880
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
      Left            =   5520
      Top             =   8760
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
      Left            =   5760
      Top             =   8880
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Forma6.frx":0000
      Height          =   330
      Left            =   3840
      TabIndex        =   20
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "ny"
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Forma6.frx":0015
      Height          =   330
      Left            =   7200
      TabIndex        =   19
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "布类"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma6.frx":002A
      Height          =   330
      Left            =   7200
      TabIndex        =   18
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1695
      Left            =   12120
      TabIndex        =   13
      Top             =   0
      Width           =   2770
      Begin VB.CheckBox Check9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "不含"
         Height          =   255
         Left            =   1080
         TabIndex        =   39
         Top             =   1320
         Width           =   730
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "含"
         Height          =   255
         Left            =   1080
         TabIndex        =   38
         Top             =   960
         Width           =   730
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "单号"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "司机"
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "负责"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   600
         Width           =   730
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "日期"
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   240
         Width           =   730
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "品名"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "客户"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "来料"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
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
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330760193
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   330760193
      CurrentDate     =   36892
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma6.frx":003F
      Height          =   8535
      Left            =   360
      TabIndex        =   17
      Top             =   1800
      Width           =   22575
      _cx             =   39820
      _cy             =   15055
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma6.frx":0054
      Height          =   2535
      Left            =   360
      TabIndex        =   23
      Top             =   10440
      Width           =   11415
      _cx             =   20135
      _cy             =   4471
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
      FormatString    =   $"Forma6.frx":0069
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Left            =   7200
      TabIndex        =   24
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "Forma6.frx":0140
      Height          =   330
      Left            =   1440
      TabIndex        =   28
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "xm"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Bindings        =   "Forma6.frx":0155
      Height          =   330
      Left            =   4320
      TabIndex        =   31
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "业务"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Height          =   330
      Left            =   9600
      TabIndex        =   34
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo2"
   End
   Begin VB.Label Label8 
      Caption         =   "双击单据号进入计划查询"
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
      Left            =   16680
      TabIndex        =   45
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "双击品名进入毛坯入库"
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
      Left            =   16680
      TabIndex        =   44
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "双击客户名进入成品出库查询"
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
      Left            =   16680
      TabIndex        =   43
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "绿色的行代表已下计划"
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
      Left            =   16680
      TabIndex        =   42
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "开单数量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15360
      TabIndex        =   40
      Top             =   10560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "首字"
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
      Left            =   9600
      TabIndex        =   36
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单号"
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
      Left            =   9600
      TabIndex        =   33
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "司机"
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
      Left            =   3840
      TabIndex        =   32
      Top             =   120
      Width           =   495
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
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   29
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "来料"
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
      Left            =   6240
      TabIndex        =   25
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "库类"
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
      Left            =   3840
      TabIndex        =   12
      Top             =   600
      Width           =   1815
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
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   11
      Top             =   1080
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
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择日期范围"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Forma6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Public TM, BAR, c, r As Integer
Private Sub Command1_Click()
On Error Resume Next

sql = ""
If Check1.value = 1 Then
sql = sql + "ny like '%'+ '" & DataCombo4.Text & "' +'%'" + " and "
End If

If Check2.value = 1 Then
sql = sql + "客户名称 like '%'+ '" & DataCombo1.Text & "' +'%'" + " and "
End If

If Check3.value = 1 Then
sql = sql + "布类 like '%'+ '" & DataCombo2.Text & "'+'%'" + " and "
End If

If Check4.value = 1 Then
sql = sql + "日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) and "
End If

If Check5.value = 1 Then
sql = sql + "业务='" & DataCombo6.Text & "' and "
End If

If Check6.value = 1 Then
sql = sql + "负责人='" & DataCombo5.Text & "' and "
End If

If Check7.value = 1 Then
sql = sql + "dh='" & DataCombo7.Text & "' and "
End If

If Check8.value = 1 Then
sql = sql + "left(单据号,1) in(" + Text4 + ") and "
End If

If Check9.value = 1 Then
sql = sql + "left(单据号,1) not in(" + Text4 + ") and "
End If

If Len(sql) > 1 Then
sql = Left$(Trim(sql), Len(Trim(sql)) - 3)
Adodc2.RecordSource = "SELECT 客户名称,布类,毛胚幅宽,克重 as 毛坯克重,毛胚匹数,毛胚重量,和约号,备注,日期,单据号,存放位置,ny as 来料单位,负责人,dh as 订单,幅宽明细,业务 as 司机,颜色,大布重量,领子重量,领子匹数 FROM CKGL WHERE (" + sql + ")   ORDER BY 日期 desc,单据号 desc"
Adodc2.Refresh
Adodc3.RecordSource = "SELECT 业务 as 司机,sum(毛胚匹数) as 合计匹数,round(sum(毛胚重量),2) as 合计重量  FROM CKGL WHERE (" + sql + ") group by 业务"
Adodc3.Refresh
End If

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 500
Next
End If
If VSFlexGrid1.Rows > 1 Then ''合计数量
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTCount, -1, 10, , vbWhite
End If

If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 600
Next
End If

' 填充数据后调用
Call UpdateGridColors
VSFlexGrid2.ColWidth(0) = 100
VSFlexGrid2.ColWidth(1) = 1500
VSFlexGrid2.ColWidth(2) = 2000
VSFlexGrid2.ColWidth(3) = 2000
VSFlexGrid2.SubtotalPosition = flexSTBelow
VSFlexGrid2.Subtotal flexSTSum, 0, 2, , vbGreen
VSFlexGrid2.Subtotal flexSTSum, 0, 3, , vbGreen

End Sub

Private Sub Command2_Click()
Call OutadodcToExcel2(VSFlexGrid1, 5, 6, DataCombo1.Text + "毛坯入库")
End Sub

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Command4_Click()
Call OutadodcToExcel2(VSFlexGrid2, 2, 3, DataCombo1.Text + "毛坯入库")
End Sub

Private Sub DataCombo1_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT 布类 FROM ckgl WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND 客户名称='" & DataCombo1.Text & "' GROUP BY 布类"
Adodc5.Refresh
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT 布类 FROM ckgl WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND 客户名称='" & DataCombo1.Text & "' GROUP BY 布类"
Adodc5.Refresh
End Sub

Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
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
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc4.RecordSource = "select ny from ckgl group by ny"
'Adodc4.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select distinct 业务  from ckgl "
Adodc6.Refresh
Text1.Text = Date
Text2.Text = Date
DTPicker3.value = Date
DTPicker4.value = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""
DataCombo7.Text = ""
Text1.TabIndex = 0
Text3 = ""
Text4 = "'R','F'"
Text5 = ""
Check4.value = 1
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 2500
VSFlexGrid1.ColWidth(6) = 1300
VSFlexGrid1.ColWidth(7) = 1900
VSFlexGrid1.ColWidth(8) = 1200

VSFlexGrid2.ColWidth(0) = 100
VSFlexGrid2.ColWidth(1) = 1500
VSFlexGrid2.ColWidth(2) = 2000
VSFlexGrid2.ColWidth(3) = 2000
VSFlexGrid2.SubtotalPosition = flexSTBelow
VSFlexGrid2.Subtotal flexSTSum, 0, 2, , vbGreen
VSFlexGrid2.Subtotal flexSTSum, 0, 3, , vbGreen

End Sub

Private Sub MSFlex()
With VSFlexGrid1
If InStr(yhm, "ck") > 0 Then
    c = .col: r = .Row    '''''C列，，R行
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

Private Sub Text3_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text3 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'  group by 简称"
Adodc1.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
Dim rs As Integer
Dim col As Integer
rs = VSFlexGrid1.Row ' 获取当前行
col = VSFlexGrid1.col ' 获取当前列
 If col = 10 Then
            If Adodc2.Recordset.EOF Then Exit Sub
            
            Adodc2.Recordset.MoveFirst
            Adodc2.Recordset.Move rs - 1 ' 移动到对应行，因为记录集索引从 0 开始

            ' 假设 Forma172 也使用 DataCombo 控件来显示选中的单据号
            Forma172.DataCombo6.Text = Adodc2.Recordset.Fields(9).value ' 设置单据号
            Forma172.Check2(0).value = 1
            Forma172.Check2(4).value = 0
            Forma172.Show ' 显示 Forma172 表单
            Forma172.Command5_Click ''''''用这个方法必须把forma172中的Command5_Click定义为Public
        End If
     If col = 1 Then
            If Adodc2.Recordset.EOF Then Exit Sub
            
            Adodc2.Recordset.MoveFirst
            Adodc2.Recordset.Move rs - 1 ' 移动到对应行，因为记录集索引从 0 开始

            
            Formc344.DataCombo2.Text = Adodc2.Recordset.Fields(9).value ' 设置单据号
            Formc344.Check2(6).value = 1
            Formc344.Check2(4).value = 0
            Formc344.Show
            Formc344.Command4_Click

        End If
If col = 2 Then
If Adodc2.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
Forma2.DataCombo1(14) = Adodc2.Recordset.Fields(9)
'Unload Me
End If
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
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move r - 1
Adodc2.Recordset.Fields(c - 1) = Combo1111.Text
Adodc2.Recordset.Update
Combo1111.Visible = False
VSFlexGrid1.Text = Combo1111.Text
VSFlexGrid1.SetFocus
End If
End Sub
Private Sub UpdateGridColors()
    Dim rsCheck As New ADODB.Recordset
    Dim sqlCheck As String
    Dim docNum As String
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim greenRowCount As Integer ' 添加变量用来统计绿色行数

    ' 循环遍历 VSFlexGrid1 的每一行
    For rowIndex = 1 To VSFlexGrid1.Rows - 1
        ' 获取单据号，索引注意检查是否正确
        docNum = VSFlexGrid1.TextMatrix(rowIndex, 10) ' 如果列索引从 0 开始，则这里应为 9
        sqlCheck = "SELECT COUNT(*) AS RecCount FROM kpd WHERE 单号 = '" & docNum & "'"
        rsCheck.Open sqlCheck, Adodc2.ConnectionString, adOpenStatic, adLockReadOnly
        
        ' 检查是否有记录
        If Not rsCheck.EOF Then
            If rsCheck.Fields("RecCount").value > 0 Then
                ' 设置整行颜色为绿色
                For colIndex = 0 To VSFlexGrid1.Cols - 1
                    VSFlexGrid1.Row = rowIndex
                    VSFlexGrid1.col = colIndex
                    VSFlexGrid1.CellBackColor = vbGreen
                Next colIndex
                greenRowCount = greenRowCount + 1 ' 增加绿色行计数
            End If
        End If
        
        rsCheck.Close
    Next rowIndex
    
    ' 更新 Text5 文本框显示绿色行数
    Text5.Text = CStr(greenRowCount)
    
    Set rsCheck = Nothing
End Sub



