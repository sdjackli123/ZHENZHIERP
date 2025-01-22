VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma171 
   BackColor       =   &H00C0E0FF&
   Caption         =   "毛坯备活进度"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   18240
      Top             =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   55
      Text            =   "Text2"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   54
      Text            =   "Text8"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   53
      Text            =   "Text2"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   52
      Text            =   "Text2"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   51
      Text            =   "Text8"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   50
      Text            =   "Text8"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Command4 
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
      Height          =   375
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "生产状态"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   14895
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已审核"
         Height          =   375
         Index           =   21
         Left            =   13680
         TabIndex        =   61
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未审核"
         Height          =   375
         Index           =   20
         Left            =   13680
         TabIndex        =   60
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已排产"
         Height          =   375
         Index           =   19
         Left            =   2760
         TabIndex        =   48
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未排产"
         Height          =   375
         Index           =   18
         Left            =   2760
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已烘干"
         Height          =   375
         Index           =   17
         Left            =   7920
         TabIndex        =   46
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未烘干"
         Height          =   375
         Index           =   16
         Left            =   7920
         TabIndex        =   45
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未染色"
         Height          =   375
         Index           =   15
         Left            =   5400
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未配料"
         Height          =   375
         Index           =   14
         Left            =   4080
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已备布"
         Height          =   375
         Index           =   13
         Left            =   1440
         TabIndex        =   42
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未备布"
         Height          =   375
         Index           =   12
         Left            =   1440
         TabIndex        =   41
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000C0C0&
         Caption         =   "清除"
         Height          =   375
         Left            =   120
         MaskColor       =   &H0000C0C0&
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已配料"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   26
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已染色"
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未脱水"
         Height          =   375
         Index           =   2
         Left            =   6720
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已脱水"
         Height          =   375
         Index           =   3
         Left            =   6720
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未圆定"
         Height          =   375
         Index           =   4
         Left            =   9240
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已圆定"
         Height          =   375
         Index           =   5
         Left            =   9240
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未开定"
         Height          =   375
         Index           =   6
         Left            =   10440
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已开定"
         Height          =   375
         Index           =   7
         Left            =   10440
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未入库"
         Height          =   375
         Index           =   8
         Left            =   11640
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已入库"
         Height          =   375
         Index           =   9
         Left            =   11640
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "未发货"
         Height          =   375
         Index           =   10
         Left            =   12600
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C0C0&
         Caption         =   "已发货"
         Height          =   375
         Index           =   11
         Left            =   12600
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "条件"
         Height          =   375
         Left            =   120
         MaskColor       =   &H0000C0C0&
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1335
      Left            =   11160
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "状态"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   63
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色号"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "布类"
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5880
      Top             =   10200
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
      Left            =   5760
      Top             =   10320
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma171.frx":0000
      Height          =   6615
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   14895
      _cx             =   26273
      _cy             =   11668
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   30
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Left            =   8880
      TabIndex        =   7
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   6360
      TabIndex        =   8
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   6360
      TabIndex        =   9
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma171.frx":0015
      Height          =   330
      Left            =   6360
      TabIndex        =   10
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   29
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   424148995
      CurrentDate     =   39961
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1560
      TabIndex        =   30
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   424148995
      CurrentDate     =   39961
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Left            =   8880
      TabIndex        =   31
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Height          =   330
      Left            =   8880
      TabIndex        =   32
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma171.frx":002A
      Height          =   615
      Left            =   360
      TabIndex        =   62
      Top             =   9000
      Width           =   9015
      _cx             =   15901
      _cy             =   1085
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   30
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   5400
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   5400
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   4680
      Top             =   10200
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
   Begin MSDataListLib.DataCombo DataCombo7 
      Height          =   330
      Left            =   1560
      TabIndex        =   64
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo7"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "生产状态"
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
      Left            =   480
      TabIndex        =   65
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   59
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   58
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   57
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   4
      Left            =   3480
      TabIndex        =   56
      Top             =   360
      Width           =   255
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
      Left            =   480
      TabIndex        =   40
      Top             =   360
      Width           =   1095
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
      Left            =   480
      TabIndex        =   39
      Top             =   840
      Width           =   1095
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
      Left            =   5160
      TabIndex        =   38
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   0
      Left            =   5160
      TabIndex        =   37
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   2
      Left            =   5160
      TabIndex        =   36
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Index           =   3
      Left            =   8280
      TabIndex        =   35
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色别"
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
      Left            =   8280
      TabIndex        =   34
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "布类"
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
      Left            =   8280
      TabIndex        =   33
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "Forma171"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public sj As Integer
Private Sub Check1_Click(Index As Integer)
Select Case Index
       Case 0
If Check1(0).value = 1 Then
cxtjsz(0) = "kp<>'N'"
Else
cxtjsz(0) = ""
End If

       Case 1
If Check1(1).value = 1 Then
cxtjsz(1) = "RS<>'N' AND ZT<>'待染色'"
Else
cxtjsz(1) = ""
End If

       Case 2
If Check1(2).value = 1 Then
cxtjsz(2) = "TS='N'"
Else
cxtjsz(2) = ""
End If

       Case 3
If Check1(3).value = 1 Then
cxtjsz(3) = "TS<>'N' and TS<>''"
Else
cxtjsz(3) = ""
End If

       Case 4
If Check1(4).value = 1 Then
cxtjsz(4) = "XDX='N'"
Else
cxtjsz(4) = ""
End If

       Case 5
If Check1(5).value = 1 Then
cxtjsz(5) = "XDX<>'N' and xdx<>''"
Else
cxtjsz(5) = ""
End If

       Case 6
If Check1(6).value = 1 Then
cxtjsz(6) = "DDX='N'"
Else
cxtjsz(6) = ""
End If

       Case 7
If Check1(7).value = 1 Then
cxtjsz(7) = "DDX<>'N' and ddx<>''"
Else
cxtjsz(7) = ""
End If

       Case 8
If Check1(8).value = 1 Then
cxtjsz(8) = "KP1='N'"
Else
cxtjsz(8) = ""
End If
       Case 9
If Check1(9).value = 1 Then
cxtjsz(9) = "KP1<>'N'"
Else
cxtjsz(9) = ""
End If
       Case 10
If Check1(10).value = 1 Then
cxtjsz(10) = "FH='N'"
Else
cxtjsz(10) = ""
End If
       Case 11
If Check1(11).value = 1 Then
cxtjsz(11) = "FH<>'N'"
Else
cxtjsz(11) = ""
End If

       Case 12
If Check1(12).value = 1 Then
cxtjsz(12) = "PB='N'"
Else
cxtjsz(12) = ""
End If

       Case 13
If Check1(13).value = 1 Then
cxtjsz(13) = "PB<>'N' and pb<>''"
Else
cxtjsz(13) = ""
End If

       Case 14
If Check1(14).value = 1 Then
cxtjsz(14) = "kp='N'"
Else
cxtjsz(14) = ""
End If

       Case 15
If Check1(15).value = 1 Then
cxtjsz(15) = "rs='N'"
Else
cxtjsz(15) = ""
End If

       Case 16
If Check1(16).value = 1 Then
cxtjsz(16) = "HG='N'"
Else
cxtjsz(16) = ""
End If

       Case 17
If Check1(17).value = 1 Then
cxtjsz(17) = "HG<>'N' and hg<>''"
Else
cxtjsz(17) = ""
End If

       Case 18
If Check1(18).value = 1 Then
cxtjsz(18) = "ye='N'"
Else
cxtjsz(18) = ""
End If

       Case 19
If Check1(19).value = 1 Then
cxtjsz(19) = "ye<>'N' and ye<>''"
Else
cxtjsz(19) = ""
End If

       Case 20
If Check1(20).value = 1 Then
cxtjsz(20) = "len(sh)< 9"
Else
cxtjsz(20) = ""
End If

       Case 21
If Check1(21).value = 1 Then
cxtjsz(21) = "len(sh)>9"
Else
cxtjsz(21) = ""
End If

End Select
End Sub

Private Sub Command1_Click()

sql1 = ""

If Check2(1).value = 1 Then
sql1 = sql1 + "客户名称 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(0).value = 1 Then
sql1 = sql1 + "生产进度 like '%'+'" & DataCombo7.Text & "'+'%' and "
End If

'If Check2(3).Value = 1 Then
'sql1 = sql1 + "标签 like '%'+'" & DataCombo2.Text & "'+'%' and "
'End If

If Check2(4).value = 1 Then
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text2(0) + ":" + Text2(1) + ":" + Text2(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text8(0) + ":" + Text8(1) + ":" + Text8(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "CONVERT(varchar,日期, 120) between '" & t1 & "' and '" & t2 & "' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "品名 like '%'+'" & DataCombo6.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "锅号 like '%'+'" & DataCombo4.Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "色别 like '%'+'" & DataCombo5.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)


If Option3.value = True Then
Adodc1.RecordSource = "SELECT * FROM v_sczy_x_kpd_jd where (" + sql1 + ")  ORDER BY 日期,锅号"
Adodc1.Refresh
Adodc3.RecordSource = "SELECT sum(isnull(匹数,0)) as 合计匹数,round(sum(isnull(重量,0)),2) as 合计重量 FROM v_sczy_x_kpd_jd where (" + sql1 + ") "
Adodc3.Refresh
End If


If Option1.value = True Then

sql = ""
For i = 0 To 21
If Check1(i).value = 1 Then
sql = sql + Trim(cxtjsz(i)) + " AND "
End If
Next

If sql = "" Then
MsgBox ("请选择生产条件")
Exit Sub
End If

sql = Left$(Trim(sql), Len(Trim(sql)) - 4)


Adodc1.RecordSource = "SELECT 车台 as 库位,xh 排产序号,标签 as 款号,客户名称,锅号,色别,品名,毛胚幅宽 as 主料匹数,光胚幅宽 as 辅料匹数,匹数 as 总匹数,重量,类别,日期,zt as 生产状态,PB AS 排布,ye as 排产,KP AS 称料,dr as 涤入缸,dc as 涤出缸,ds as 染涤时长,mr as 棉入缸,mc as 棉出缸,ms as 染棉时长,TS AS 脱水,HG AS 烘干,XDX as 圆定,DDX AS 开幅,KP1 as 入库,FH AS 发货,sh as 审核 FROM v_sczy_x_kpd_jd where (" + sql + ") AND (" + sql1 + ")  ORDER BY 日期,锅号,ip"
Adodc1.Refresh
Adodc3.RecordSource = "SELECT sum(isnull(匹数,0)) as 合计匹数,round(sum(isnull(重量,0)),2) as 合计重量 FROM v_sczy_x_kpd_jd where (" + sql + ") AND (" + sql1 + ") "
Adodc3.Refresh
End If

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.AutoSizeMode = flexAutoSizeRowHeight
VSFlexGrid1.AutoSize 0, VSFlexGrid1.Cols - 1, False, 30
End Sub


Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Call jdmx(VSFlexGrid1, "进度明细")
End Sub

Private Sub Form_Load()
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""
DataCombo7.Text = ""
DTPicker1.value = Date
DTPicker2.value = Date
Text1.Text = ""

Text2(0) = "00"
Text2(1) = "00"
Text2(2) = "01"

Text8(0) = "23"
Text8(1) = "59"
Text8(2) = "59"

Option3.value = True
sj = 1

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM v_sczy_x_kpd_jd where 日期 between cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) ORDER BY 客户名称,日期,锅号"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from khZL where IP LIKE '%'+'" & yhxx & "'+'%' group by 简称"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(0) = 200
For i = 1 To 11
VSFlexGrid1.ColWidth(i) = 1200
Next

End Sub

Private Sub Option3_Click()
For i = 0 To 11
Check1(i).value = 0
cxtjsz(i) = ""
Next
End Sub

Private Sub Text1_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%' and IP LIKE '%'+'" & yhxx & "'+'%' group by 简称"
Adodc2.Refresh
End Sub

