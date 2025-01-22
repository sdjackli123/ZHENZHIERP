VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formc344 
   BackColor       =   &H00C0E0FF&
   Caption         =   "发货查询"
   ClientHeight    =   12945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   26445
   LinkTopic       =   "Form39"
   MDIChild        =   -1  'True
   ScaleHeight     =   12945
   ScaleWidth      =   26445
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   735
      Left            =   21960
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   9120
      Width           =   850
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   615
      Left            =   16920
      Top             =   12840
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "Formc344.frx":0000
      Height          =   330
      Left            =   1440
      TabIndex        =   52
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "跟单"
      Text            =   "DataCombo5"
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   11160
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   9120
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   735
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   9120
      Width           =   850
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   6960
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   11160
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   21840
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   5640
      Top             =   10320
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      Left            =   6120
      Top             =   10320
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
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3600
      TabIndex        =   30
      Text            =   "Text2"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   9240
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   6840
      Top             =   10560
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1455
      Left            =   12840
      TabIndex        =   16
      Top             =   360
      Width           =   5055
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "品名"
         Height          =   375
         Index           =   13
         Left            =   3960
         TabIndex        =   49
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色号"
         Height          =   375
         Index           =   12
         Left            =   3960
         TabIndex        =   46
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "颜色"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "核算"
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   42
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "总备注"
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   35
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "司机"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "类别"
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "备注"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单号"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "负责"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "款号"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Formc344.frx":0015
      Height          =   6975
      Left            =   720
      TabIndex        =   15
      Top             =   2040
      Width           =   25935
      _cx             =   45746
      _cy             =   12303
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
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
      MergeCells      =   110
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7200
      Top             =   10320
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
      Height          =   330
      Left            =   7320
      Top             =   10560
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Left            =   6960
      TabIndex        =   14
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formc344.frx":002A
      Height          =   330
      Left            =   6960
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "xm"
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   4080
      TabIndex        =   12
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc344.frx":003F
      Height          =   330
      Left            =   4080
      TabIndex        =   11
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   19200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   19200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   19200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   423886849
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   423886849
      CurrentDate     =   36892
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formc344.frx":0054
      Height          =   3615
      Left            =   720
      TabIndex        =   23
      Top             =   9120
      Width           =   9855
      _cx             =   17383
      _cy             =   6376
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
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
      MergeCells      =   110
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
      Bindings        =   "Formc344.frx":0069
      Height          =   3495
      Left            =   12840
      TabIndex        =   31
      Top             =   9120
      Width           =   9135
      _cx             =   16113
      _cy             =   6165
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
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
      MergeCells      =   110
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Formc344.frx":007E
      Left            =   4080
      List            =   "Formc344.frx":008B
      TabIndex        =   44
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
      Height          =   375
      Index           =   12
      Left            =   10680
      TabIndex        =   50
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号"
      Height          =   375
      Index           =   11
      Left            =   8760
      TabIndex        =   48
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
      Height          =   375
      Index           =   10
      Left            =   6480
      TabIndex        =   41
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "核算"
      Height          =   375
      Index           =   9
      Left            =   3240
      TabIndex        =   39
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "总备注"
      Height          =   375
      Index           =   8
      Left            =   10680
      TabIndex        =   37
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "司机"
      Height          =   375
      Index           =   7
      Left            =   720
      TabIndex        =   33
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
      Height          =   375
      Index           =   6
      Left            =   8760
      TabIndex        =   25
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "类别"
      Height          =   375
      Index           =   5
      Left            =   8760
      TabIndex        =   24
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
      Height          =   375
      Index           =   4
      Left            =   6480
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "负责"
      Height          =   375
      Index           =   3
      Left            =   6480
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号（单号）"
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户"
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Formc344"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer

Private Sub Combo1_Change()
Text1(4) = Combo1
End Sub

Private Sub Combo1_Click()
Text1(4) = Combo1
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command11_Click()
Call MXOutadodcToExcel(VSFlexGrid3, "")
End Sub

Private Sub Command2_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "发货报表")
End Sub

Private Sub Command3_Click()
Call MXOutadodcToExcel(VSFlexGrid2, "发货报表")
End Sub

Public Sub Command4_Click()
sql1 = ""

If Check2(1).value = 1 Then
sql1 = sql1 + "客户名称='" & DataCombo1.Text & "' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "备注 like '%'+'" & Text1(0) & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "和约号 like '%'+'" & DataCombo4.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "锅号 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
t1 = Format(Trim(DTPicker3.value), "yyyy-mm-dd")
t2 = Format(Trim(DTPicker4.value), "yyyy-mm-dd")
sql1 = sql1 + "CONVERT(varchar,日期, 23) between '" & t1 & "' and '" & t2 & "' and "
End If

If Check2(0).value = 1 Then
sql1 = sql1 + "单号 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "负责 like '%'+'" & DataCombo3.Text & "'+'%' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "加工类别 like '%'+'" & Text1(1) & "'+'%' and "
End If

If Check2(8).value = 1 Then
sql1 = sql1 + "跟单 like '%'+'" & DataCombo5.Text & "'+'%' and "
End If

If Check2(9).value = 1 Then
sql1 = sql1 + "总备注 like '%'+'" & Text1(3) & "'+'%' and "
End If

If Check2(10).value = 1 Then
sql1 = sql1 + "核算 like '%'+'" & Text1(4) & "'+'%' and "
End If

If Check2(11).value = 1 Then
sql1 = sql1 + "颜色 like '%'+'" & Text1(5) & "'+'%' and "
End If

If Check2(12).value = 1 Then
sql1 = sql1 + "发票已开 like '%'+'" & Text1(6) & "'+'%' and "
End If
If Check2(13).value = 1 Then
sql1 = sql1 + "品名 like '%'+'" & Text1(7) & "'+'%' and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc1.RecordSource = "select 日期,单号,客户名称,锅号,品名,颜色,匹数,大布重量,领子重量,数量 as 总重量,光坯,核算,加工类别,备注,负责,跟单 as 提货,单价,金额,入库日期,来料单位 from v_jgmx where (" + sql1 + ")  order by 日期,单号"
Adodc1.Refresh
Adodc3.RecordSource = "select 客户名称,单号,round(sum(匹数),0) as 毛坯匹数,round(sum(数量),2) as 毛坯重量,round(sum(isnull(光坯,0)),2) as 光坯重量,round(sum(isnull(金额,0)),2) as 总金额 from v_jgmx where (" + sql1 + ")  group by 客户名称,单号 order by 客户名称,单号"
Adodc3.Refresh
Adodc5.RecordSource = "select 跟单 as 司机, round(sum(匹数),0) as 毛坯匹数,round(sum(数量),2) as 毛坯重量,round(sum(isnull(光坯,0)),2) as 光坯重量,round(sum(isnull(金额,0)),2) as 总金额 from v_jgmx where (" + sql1 + ") group by 跟单"
Adodc5.Refresh

VSFlexGrid3.AutoSizeMode = flexAutoSizeRowHeight
VSFlexGrid3.AutoSize 0, VSFlexGrid3.Cols - 1, False, 30

If VSFlexGrid3.Rows > 1 Then
For i = 1 To VSFlexGrid3.Rows - 1
VSFlexGrid3.RowHeight(i) = 500
Next
End If
VSFlexGrid3.ColWidth(0) = 300
VSFlexGrid3.ColWidth(1) = 1300
VSFlexGrid3.ColWidth(2) = 1300
VSFlexGrid3.ColWidth(3) = 1200
VSFlexGrid3.ColWidth(4) = 1200
VSFlexGrid3.ColWidth(5) = 0
VSFlexGrid3.ColWidth(6) = 1500
VSFlexGrid3.ColWidth(7) = 1500
VSFlexGrid3.ColWidth(8) = 0
VSFlexGrid3.ColWidth(9) = 1200
VSFlexGrid3.ColWidth(10) = 1200
VSFlexGrid3.ColWidth(11) = 1200
VSFlexGrid3.ColWidth(12) = 1200
VSFlexGrid3.ColWidth(13) = 1200
VSFlexGrid3.ColWidth(14) = 1200
VSFlexGrid3.ColWidth(15) = 1200
VSFlexGrid3.ColWidth(16) = 2000
VSFlexGrid3.ColWidth(17) = 1000
VSFlexGrid3.ColWidth(18) = 1000
VSFlexGrid3.ColWidth(19) = 1500
VSFlexGrid3.ColWidth(20) = 1500

VSFlexGrid3.SubtotalPosition = flexSTBelow
VSFlexGrid3.Subtotal flexSTSum, 0, 7, , &HC0C0&
VSFlexGrid3.Subtotal flexSTSum, 0, 8, , &HC0C0&
VSFlexGrid3.Subtotal flexSTSum, 0, 9, , &HC0C0&
VSFlexGrid3.Subtotal flexSTSum, 0, 10, , &HC0C0&
VSFlexGrid3.Subtotal flexSTSum, 0, 11, , &HC0C0&
VSFlexGrid3.Subtotal flexSTSum, 0, 17, , &HC0C0&

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1500
VSFlexGrid1.ColWidth(3) = 1500
VSFlexGrid1.ColWidth(4) = 1500
VSFlexGrid1.ColWidth(5) = 1500
VSFlexGrid1.ColWidth(6) = 1500

If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 600
Next
End If
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(1) = 1200
VSFlexGrid2.ColWidth(2) = 1500
VSFlexGrid2.ColWidth(3) = 1500
VSFlexGrid2.ColWidth(4) = 1500
VSFlexGrid2.ColWidth(5) = 1500
VSFlexGrid2.SubtotalPosition = flexSTBelow
VSFlexGrid2.Subtotal flexSTSum, 0, 2, , &HC0C0&
VSFlexGrid2.Subtotal flexSTSum, 0, 3, , &HC0C0&
VSFlexGrid2.Subtotal flexSTSum, 0, 4, , &HC0C0&
VSFlexGrid2.Subtotal flexSTSum, 0, 5, , &HC0C0&

VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 1, 3, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 1, 4, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 1, 5, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 1, 6, , &HC0C0&
End Sub

Private Sub Form_Load()
DTPicker3.value = Date
DTPicker4.value = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = yhxm
DataCombo4.Text = ""
DataCombo5.Text = ""
For i = 0 To 7
Text1(i) = ""
Next
Text2 = ""
cdbhf = cdbh
Check2(4).value = 1
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 日期,单号,客户名称,锅号,和约号 as 款号,品名,颜色,发票已开 as 色号,匹数,大布重量,领子重量,数量 as 总重量,光坯,核算,加工类别,备注,负责,跟单 as 提货,单价,金额,附加费单价,附加费金额,总金额,入库日期,来料单位 from v_jgmx where 日期=cast('" & DTPicker3.value & "' as datetime)  order by 日期,单号"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from khZL where ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select xm  from fzr group by xm"
Adodc4.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select distinct 跟单  from v_jgmx "
Adodc6.Refresh
VSFlexGrid3.ColWidth(0) = 300
VSFlexGrid3.ColWidth(1) = 1200
VSFlexGrid3.ColWidth(2) = 1200
VSFlexGrid3.ColWidth(3) = 1200
VSFlexGrid3.ColWidth(4) = 1200
VSFlexGrid3.ColWidth(5) = 0
VSFlexGrid3.ColWidth(6) = 1500
VSFlexGrid3.ColWidth(7) = 1500
VSFlexGrid3.ColWidth(8) = 0
VSFlexGrid3.ColWidth(9) = 1200
VSFlexGrid3.ColWidth(10) = 1200
VSFlexGrid3.ColWidth(11) = 1200
VSFlexGrid3.ColWidth(12) = 1200
VSFlexGrid3.ColWidth(13) = 1200
VSFlexGrid3.ColWidth(14) = 1200
VSFlexGrid3.ColWidth(15) = 1200
VSFlexGrid3.ColWidth(16) = 1800
VSFlexGrid3.ColWidth(17) = 1000
VSFlexGrid3.ColWidth(18) = 1000
VSFlexGrid3.ColWidth(19) = 800
VSFlexGrid3.ColWidth(20) = 1200
VSFlexGrid3.ColWidth(21) = 0
VSFlexGrid3.ColWidth(22) = 0
VSFlexGrid3.ColWidth(23) = 1200
VSFlexGrid3.ColWidth(24) = 1200



VSFlexGrid3.ColWidth(19) = 1500
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

Private Sub Text2_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from khZL where ip like '%'+'" & yhxx & "'+'%' and 代码 like '%'+'" & Text2 & "'+'%' group by 简称"
Adodc2.Refresh
End Sub

Private Sub VSFlexGrid3_dblClick()
If wwdm = 4 Then
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid3.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
Formc15.Label13 = Adodc1.Recordset.Fields(1)
Unload Me
End If
End Sub
