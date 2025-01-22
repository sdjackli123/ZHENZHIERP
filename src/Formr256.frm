VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formr256 
   BackColor       =   &H00C0E0FF&
   Caption         =   "出库操作"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form22"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formr256.frx":0000
      Height          =   7935
      Left            =   9120
      TabIndex        =   56
      Top             =   1800
      Width           =   6135
      _cx             =   10821
      _cy             =   13996
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formr256.frx":0015
      Height          =   3375
      Left            =   480
      TabIndex        =   55
      Top             =   6360
      Width           =   8415
      _cx             =   14843
      _cy             =   5953
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formr256.frx":002A
      Height          =   330
      Left            =   3000
      TabIndex        =   54
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "染化助库名"
      Text            =   "DataCombo2"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   5880
      TabIndex        =   32
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-M-d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   30
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询功能"
      Height          =   975
      Left            =   480
      TabIndex        =   20
      Top             =   4800
      Width           =   8415
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Left            =   1800
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Left            =   4440
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "普通查询"
         Height          =   375
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "显示全部"
         Height          =   375
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "打印"
         Height          =   375
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0C0FF&
         Caption         =   "染助查询"
         Height          =   375
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarTitleBackColor=   8421376
         CalendarTrailingForeColor=   249
         Format          =   285212673
         CurrentDate     =   36892
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   255
         Left            =   4440
         TabIndex        =   28
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421376
         CalendarTrailingForeColor=   249
         Format          =   285212673
         CurrentDate     =   36892
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "时间范围："
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   4320
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   16
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   15
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   600
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   360
      Top             =   720
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   7560
      TabIndex        =   9
      Text            =   "0"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   2040
      TabIndex        =   8
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   7560
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text1111 
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   5760
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新单据"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   735
   End
   Begin VB.Timer Timer3 
      Interval        =   20
      Left            =   0
      Top             =   1320
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Formr256.frx":003F
      Left            =   4320
      List            =   "Formr256.frx":0049
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   3600
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   5640
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   600
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   285212673
      CurrentDate     =   36892
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7440
      Top             =   10440
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
      Left            =   7440
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
      Left            =   8160
      Top             =   10560
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
      Left            =   7920
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
      Left            =   7800
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
      Left            =   9360
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
      Height          =   330
      Left            =   8400
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   7920
      Top             =   10440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formr256.frx":0059
      Height          =   330
      Left            =   2160
      TabIndex        =   53
      Top             =   3600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Index           =   5
      Left            =   2160
      TabIndex        =   12
      Top             =   3945
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "入库时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   52
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染化助剂名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   51
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "数量(公斤)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   50
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单价(元/公斤)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   49
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "合计金额(元)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   48
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "出库单位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   47
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请先选择染化助剂库："
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
      Left            =   600
      TabIndex        =   46
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "IP号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   45
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "关联库："
      Height          =   255
      Left            =   9120
      TabIndex        =   44
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "染化助剂录入明晰表："
      Height          =   255
      Left            =   480
      TabIndex        =   43
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染化助库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   42
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Height          =   255
      Index           =   7
      Left            =   7560
      TabIndex        =   41
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "染化助出库操作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   2040
      TabIndex        =   40
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "含税额(元)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   39
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "不含税额(元)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   38
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H008080FF&
      Caption         =   "1"
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
      Left            =   2760
      TabIndex        =   37
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H008080FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   36
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H008080FF&
      Caption         =   "5"
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
      Left            =   600
      TabIndex        =   35
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label12 
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
      Left            =   5760
      TabIndex        =   34
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "库别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   4320
      TabIndex        =   33
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "Formr256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Command1_Click()
On Error Resume Next
KL = 0
Dim X As String
If Text1(2).text = "" Then
Text1(2).text = 0
End If
If Text1(3).text = "" Then
Text1(3).text = 0
End If

If Trim(Combo1.text) = "" Then
MsgBox ("请选择库别！")
Exit Sub
End If
If Text1(6).text = "" Then
MsgBox ("请输入染化库名")
Exit Sub
End If
If Text1(5).text = "" Then
MsgBox ("请输入供应商")
Exit Sub
End If

If DataCombo2.text <> Text1(6).text Then
MsgBox ("染化助库有误！")
Exit Sub
End If

If Text1(10).text = "" Then Text1(10).text = 0
If Text1(4).text = "" Then Text1(4).text = 0
If Text1(9).text = "" Then Text1(9).text = 0
Text1(5).text = DataCombo1.text


Adodc1.Recordset.AddNew
For i = 0 To Adodc1.Recordset.Fields.Count - 1
Adodc1.Recordset.Fields(i) = Text1(i).text
Next
Adodc1.Recordset.Fields(5) = DataCombo1.text
Adodc1.Recordset.Fields(13) = Combo1.text
Adodc1.Recordset.Fields(14) = "未"
Adodc1.Recordset.Fields(15) = "未"
Adodc1.Recordset.Update
Adodc1.Refresh


For i = 1 To Adodc1.Recordset.Fields.Count - 1

If i = 6 Then i = 6 + 1
If i = 11 Then i = 12
Text1(i).text = ""
Next

Text1(7).text = Adodc1.Recordset.RecordCount + 1
End Sub


Private Sub Command11_Click()
On Error Resume Next
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
Adodc4.Refresh
Adodc3.Refresh
Adodc1.RecordSource = "select *  from ckmx where  单据号='" & Text1(11).text & "'"
Adodc1.Refresh
Text1(7).text = Adodc1.Recordset.RecordCount + 1
End Sub

Private Sub Command12_Click()
On Error Resume Next
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc6.RecordSource = "SELECT * FROM rlckdj where 单据编号='" & yhdm & "'"
Adodc6.Refresh

Text1(11).text = yhdm + "0000001"
If Adodc6.Recordset.EOF Then
Text1(11).text = yhdm + "0000001"
Else
uu = Val(Adodc6.Recordset.Fields(1)) + 1
Text1(11).text = yhdm + left("0000000", 7 - Len(Trim(Str(uu)))) + Trim(Str(uu))
End If

Adodc1.RecordSource = "select *  from ckmx where  单据号='" & Text1(11).text & "'"
Adodc1.Refresh
Text1(7).text = Adodc1.Recordset.RecordCount + 1
Text1(11).Enabled = False
End Sub

Private Sub Command13_Click()
Call rhlck(Adodc1, Text1(11).text)
End Sub

Private Sub Command14_Click()
If Text1(1).text = "" Then
Adodc1.RecordSource = "select 出库单位,染化助库名,名称,出库数量,单价,合计金额,单据号,出库时间  from ckmx where 库别='" & Combo1.text & "' AND 出库时间 between '" & Text2.text & "'  and   '" & Text3.text & "' order by 出库时间"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select 出库单位,染化助库名,名称,出库数量,单价,合计金额,单据号,出库时间  from ckmx where 库别='" & Combo1.text & "' AND 出库时间 between '" & Text2.text & "'  and   '" & Text3.text & "' AND 名称='" & Text1(1).text & "' order by 出库时间"
Adodc1.Refresh
End If
End Sub

Private Sub Command2_Click()
If DataCombo2.text <> Text1(6).text Then
MsgBox ("染化助库有误！")
Exit Sub
End If
On Error Resume Next
Text1(5).text = DataCombo1.text
For i = 0 To Adodc1.Recordset.Fields.Count - 1
Adodc1.Recordset.Fields(i) = Text1(i).text
Next
Adodc1.Recordset.Fields(5) = DataCombo1.text
Adodc1.Recordset.Fields(13) = Combo1.text
Adodc1.Recordset.Update
Adodc1.Refresh

For i = 1 To Adodc1.Recordset.Fields.Count - 1
If i = 11 Then i = 12
Text1(i).text = ""
Next

Text1(7).text = Adodc1.Recordset.RecordCount + 1

End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text1(5).text = "" Then
Adodc1.RecordSource = "select 出库单位,染化助库名,名称,出库数量,单价,合计金额,单据号,出库时间  from ckmx where 库别='" & Combo1.text & "' AND 出库时间 between '" & Text2.text & "'  and   '" & Text3.text & "' order by 出库时间,单据号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select 出库单位,染化助库名,名称,出库数量,单价,合计金额,单据号,出库时间 from ckmx where 库别='" & Combo1.text & "' AND 出库时间 between '" & Text2.text & "'  and   '" & Text3.text & "' AND 出库单位='" & Text1(5).text & "' order by 出库时间,单据号"
Adodc1.Refresh
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.Delete
Adodc1.Refresh

For i = 1 To Adodc1.Recordset.Fields.Count - 1
If i = 11 Then i = 12
Text1(i).text = ""
Next

Text1(7).text = Adodc1.Recordset.RecordCount + 1

End Sub

Private Sub Command5_Click()
Call OutadodcToExcel(VSFlexGrid1, 6, "出库明细")
End Sub

Private Sub Command6_Click()
On Error Resume Next
BA.Close
Unload Me
End Sub

Private Sub Command7_Click()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc1.RecordSource = "select * from ckmx WHERE 库别='退库' order by 出库时间,VAL(IP)"
Adodc1.Refresh
End Sub

Private Sub DataCombo1_Change()
Text1(5).text = DataCombo1.text
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Text1(5).text = DataCombo1.text
End Sub

Private Sub DataCombo2_Change()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc1.RecordSource = "select *  from ckmx where  单据号='" & Text1(11).text & "'"
Adodc1.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc5.RecordSource = "select *  from RHZH where 染化助库名='" & DataCombo2.text & "' "
Adodc5.Refresh
Text1(7).text = Adodc1.Recordset.RecordCount + 1
Text1(6).text = DataCombo2.text
End Sub

Private Sub dataCombo2_Click(Area As Integer)
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc1.RecordSource = "select *  from ckmx where  单据号='" & Text1(11).text & "'"
Adodc1.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc5.RecordSource = "select *  from RHZH where 染化助库名='" & DataCombo2.text & "' "
Adodc5.Refresh
Text1(6).text = DataCombo2.text
Text1(7).text = Adodc1.Recordset.RecordCount + 1
End Sub

Private Sub DTPicker1_Change()
Text1(0).text = DTPicker1.Value
End Sub


Private Sub DTPicker1_CloseUp()
Text1(0).text = DTPicker1.Value
Text1(0).SetFocus
End Sub

Private Sub DTPicker2_Change()
Text2.text = DTPicker2.Value
End Sub


Private Sub DTPicker3_Change()
Text3.text = DTPicker3.Value
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption + "操作年度 " + ljb
On Error Resume Next

For i = 0 To 16
Text1(i) = ""
Next
Label10.Visible = False
Text1(0).text = Date
DTPicker1.Value = Date
Text1(8).text = ""
Text6.Enabled = False
Text2.text = Date
Text3.text = Date
DTPicker2.Value = Date
DTPicker3.Value = Date
Text1(11).Enabled = False
VSFlexGrid1.ColWidth(1) = 1100
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(6) = 1200
Combo1.text = "退库"
ProgressBar1.Visible = False
Timer3.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False



Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc3.RecordSource = "select GYS.简称  from GYS  GROUP BY GYS.简称 "
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc4.RecordSource = "select 染化助库名  from RHZH GROUP BY 染化助库名 "
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc5.RecordSource = "select *  from RHZH where 染化助库名='" & DataCombo2.text & "' "
Adodc5.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc6.RecordSource = "SELECT * FROM rlckdj where 单据编号='" & yhdm & "'"
Adodc6.Refresh

Text1(11).text = yhdm + "0000001"
If Adodc6.Recordset.EOF Then
Text1(11).text = yhdm + "0000001"
Else
uu = Val(Adodc6.Recordset.Fields(1)) + 1
Text1(11).text = yhdm + left("0000000", 7 - Len(Trim(Str(uu)))) + Trim(Str(uu))
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc1.RecordSource = "select *  from ckmx where  单据号='" & Text1(11).text & "'"
Adodc1.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Text1(0).text = Date
DataCombo1.text = ""
DataCombo2.text = ""
DataCombo2.TabIndex = 0
L = 10
Text1(7).text = Adodc1.Recordset.RecordCount + 1

VSFlexGrid2.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(6) = 1500
VSFlexGrid1.ColWidth(10) = 0
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0

End Sub

Private Sub Label12_Click()
Text1(11).Enabled = True
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To Adodc1.Recordset.Fields.Count - 1
Text1(i).text = Adodc1.Recordset.Fields(i)
Next
Combo1.text = Adodc1.Recordset.Fields(13)
DataCombo1.text = Text1(5).text
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub

Private Sub VSFlexGrid2_Click()
On Error Resume Next
rs = VSFlexGrid2.Row
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move rs - 1
Text1(1).text = Adodc5.Recordset.Fields(0)
Text1(6).text = Adodc5.Recordset.Fields(5)
DataCombo1.text = Adodc5.Recordset.Fields(1)
Adodc8.RecordSource = "SELECT 单价 FROM mx WHERE 名称='" & Text1(1).text & "' ORDER BY 入库时间 DESC"
Adodc8.Refresh
If Adodc8.Recordset.EOF Then
Text1(3).text = 0
Else
Text1(3).text = Format(Adodc8.Recordset.Fields(0), "#0.00")
End If
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next

Select Case Index

Case 2
Text1(10).text = Format(Val(Text1(2).text) * Val(Text1(3).text), "#0.00")

Case 3
Text1(10).text = Format(Val(Text1(2).text) * Val(Text1(3).text), "#0.00")

Case 11
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=tyfw"
Adodc1.RecordSource = "select *  from ckmx where  单据号='" & Text1(11).text & "'"
Adodc1.Refresh
Text1(7).text = Adodc1.Recordset.RecordCount + 1

End Select
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub


Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub VSFlexGrid2_dblClick()
With VSFlexGrid2
    c = .col: r = .Row
        Text1111.left = .left + .ColPos(c)
        Text1111.top = .top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub VSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call VSFlexGrid2_dblClick
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    VSFlexGrid2.text = Text1111.text
    Text1111.Visible = False
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move r - 1

Adodc5.Recordset.Fields(c - 1) = Text1111.text
Adodc5.Recordset.Update
Text1111.Visible = False
End Sub

