VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formc39 
   BackColor       =   &H00C0E0FF&
   Caption         =   "成品发货"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16020
   LinkTopic       =   "Form39"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   16020
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印确定"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   4440
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询功能"
      Height          =   735
      Left            =   1440
      TabIndex        =   42
      Top             =   5040
      Width           =   13095
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "客户查询"
         Height          =   375
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7200
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DataCombo10 
         Bindings        =   "Formc39.frx":0000
         Height          =   330
         Left            =   1080
         TabIndex        =   43
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "简称"
         Text            =   "DataCombo10"
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4920
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421376
         CalendarTrailingForeColor=   255
         Format          =   130416641
         CurrentDate     =   39181
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7200
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8421376
         CalendarTrailingForeColor=   255
         Format          =   130416641
         CurrentDate     =   39181
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   10200
         TabIndex        =   49
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "时间范围："
         Height          =   375
         Left            =   3960
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "客户名称"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   6480
         X2              =   7080
         Y1              =   360
         Y2              =   360
      End
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
      Left            =   14760
      MultiLine       =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6600
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "入库"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   14040
      TabIndex        =   38
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "下一单据号"
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   13440
      TabIndex        =   36
      Text            =   "Text6"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4440
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3720
      Top             =   240
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8640
      TabIndex        =   34
      Text            =   "Text7"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   11160
      TabIndex        =   33
      Top             =   3360
      Width           =   975
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
      ItemData        =   "Formc39.frx":0015
      Left            =   12120
      List            =   "Formc39.frx":0022
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2640
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细进度"
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Text            =   "Text9"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "出库查询"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5760
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "欠款条"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "结算单"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "出库单"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      Text            =   "Text10"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "核算方式"
      Height          =   855
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   3015
      Begin VB.OptionButton Option4 
         BackColor       =   &H0000C0C0&
         Caption         =   "毛坯"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H0000C0C0&
         Caption         =   "光坯"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Text            =   "Text11"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Text            =   "Text11"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Text            =   "Text11"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   11160
      TabIndex        =   1
      Text            =   "Text11"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "发货信息"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc26 
      Height          =   330
      Left            =   11400
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
      Left            =   10920
      Top             =   10800
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formc39.frx":0038
      Height          =   2895
      Left            =   480
      TabIndex        =   17
      Top             =   6360
      Width           =   13935
      _cx             =   24580
      _cy             =   5106
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
   Begin MSDataListLib.DataCombo DataCombo17 
      Height          =   330
      Left            =   5160
      TabIndex        =   18
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo17"
   End
   Begin MSDataListLib.DataCombo DataCombo16 
      Height          =   330
      Left            =   5160
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo16"
   End
   Begin MSDataListLib.DataCombo DataCombo14 
      Height          =   330
      Left            =   10920
      TabIndex        =   20
      Top             =   7200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo14"
   End
   Begin MSDataListLib.DataCombo DataCombo13 
      Height          =   330
      Left            =   1800
      TabIndex        =   21
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo13"
   End
   Begin MSDataListLib.DataCombo DataCombo12 
      Height          =   330
      Left            =   14040
      TabIndex        =   22
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo12"
   End
   Begin MSDataListLib.DataCombo DataCombo11 
      Height          =   330
      Left            =   3240
      TabIndex        =   23
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo11"
   End
   Begin MSDataListLib.DataCombo DataCombo9 
      Height          =   330
      Left            =   1800
      TabIndex        =   24
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo9"
   End
   Begin MSDataListLib.DataCombo DataCombo8 
      Height          =   330
      Left            =   10080
      TabIndex        =   25
      Top             =   7200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo8"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Height          =   330
      Left            =   12240
      TabIndex        =   26
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo7"
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Height          =   330
      Left            =   10080
      TabIndex        =   27
      Top             =   7560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo6"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Left            =   9600
      TabIndex        =   28
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo5"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Left            =   480
      TabIndex        =   29
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formc39.frx":004D
      Height          =   330
      Left            =   6360
      TabIndex        =   30
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "YS"
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formc39.frx":0062
      Height          =   330
      Left            =   4680
      TabIndex        =   31
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "PM"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc39.frx":0077
      Height          =   330
      Left            =   2040
      TabIndex        =   32
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   330
      Left            =   14040
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   130416641
      CurrentDate     =   39181
   End
   Begin MSAdodcLib.Adodc Adodc25 
      Height          =   330
      Left            =   8400
      Top             =   10680
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
      Left            =   8640
      Top             =   10440
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
      Left            =   8760
      Top             =   10440
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
      Left            =   8880
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
      Left            =   9120
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
      Left            =   9120
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
      Left            =   10200
      Top             =   10680
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
      Left            =   11640
      Top             =   10560
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
      Left            =   9840
      Top             =   10680
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
      Left            =   9480
      Top             =   10560
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
      Left            =   10200
      Top             =   10560
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
      Left            =   10440
      Top             =   10440
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
      Left            =   9480
      Top             =   10680
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
      Left            =   10200
      Top             =   10800
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
      Left            =   9960
      Top             =   10440
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
      Left            =   10680
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
      Left            =   9960
      Top             =   10560
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
      Left            =   11760
      Top             =   10680
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
      Left            =   11760
      Top             =   10800
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
      Left            =   9240
      Top             =   10440
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
      Left            =   9600
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
      Left            =   9480
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
      Left            =   9720
      Top             =   10560
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
      Left            =   9840
      Top             =   10440
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
      Left            =   5880
      TabIndex        =   58
      Top             =   10200
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo12"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单价"
      Height          =   255
      Index           =   6
      Left            =   11160
      TabIndex        =   88
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户加工明细表"
      Height          =   255
      Left            =   480
      TabIndex        =   87
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "日期"
      Height          =   255
      Index           =   0
      Left            =   14040
      TabIndex        =   86
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "毛坯重量"
      Height          =   255
      Index           =   5
      Left            =   9600
      TabIndex        =   85
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "锅号"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   84
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   83
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "品名"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   82
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   81
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "金额（元）"
      Height          =   255
      Index           =   0
      Left            =   12240
      TabIndex        =   80
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "合计总额（元）："
      Height          =   855
      Left            =   14760
      TabIndex        =   79
      Top             =   6120
      Width           =   255
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
      Left            =   480
      TabIndex        =   78
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   77
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   255
      Left            =   14040
      TabIndex        =   76
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "顺序号："
      Height          =   375
      Left            =   480
      TabIndex        =   75
      Top             =   1920
      Width           =   1215
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
      Left            =   12240
      TabIndex        =   74
      Top             =   2040
      Width           =   1215
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
      Left            =   13440
      TabIndex        =   73
      Top             =   2040
      Width           =   1935
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
      Left            =   12120
      TabIndex        =   72
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号:"
      Height          =   375
      Index           =   1
      Left            =   12120
      TabIndex        =   71
      Top             =   1440
      Width           =   1335
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
      Left            =   4560
      TabIndex        =   70
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H008080FF&
      Caption         =   "4"
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
      Left            =   2760
      TabIndex        =   69
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "匹数"
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   68
      Top             =   2640
      Width           =   855
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
      Left            =   4560
      TabIndex        =   67
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF80&
      Caption         =   "清除"
      Height          =   255
      Left            =   1440
      TabIndex        =   66
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "色号"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   65
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "收费明细"
      Height          =   255
      Left            =   5880
      TabIndex        =   64
      Top             =   9840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光坯重量"
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   63
      Top             =   3000
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
      Left            =   480
      TabIndex        =   62
      Top             =   3480
      Width           =   1455
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
      Left            =   480
      TabIndex        =   61
      Top             =   4200
      Width           =   1455
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
      Left            =   4680
      TabIndex        =   60
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单位"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   11160
      TabIndex        =   59
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "Formc39"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public gygh As String
Dim JDBAR As Integer
Dim hs, ZS, ps As Integer: Dim fhsl As Single: Dim je As Single
Dim cdbhf As Integer


Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Command10_Click()
Formc151.Text1(0) = Label13.Caption
Formc151.Show
End Sub

Private Sub Command12_Click()
Formc34.DataCombo1(4).Text = DataCombo4.Text
Formc34.Show
End Sub

Private Sub Command14_Click()

Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc16.RecordSource = "SELECT * FROM JGMX where 日期=cast('" & Text5.Text & "' as datetime) and left(单号,1)='" & yhdm & "' and len(单号)=10"
Adodc16.Refresh

If Adodc16.Recordset.EOF Then
Label13.Caption = yhdm + Format(CDate(Text5.Text), "yymmdd") + "001"
Else
Adodc16.RecordSource = "SELECT max(right(单号,len(单号)-7)) FROM JGMX where 日期=cast('" & Text5.Text & "' as datetime) and left(单号,1)='" & yhdm & "' and len(单号)=10"
Adodc16.Refresh
L = Val(Adodc16.Recordset.Fields(0))
If Len(L + 1) = 1 Then
Label13.Caption = yhdm + Format(CDate(Text5.Text), "yymmdd") + "00" + Trim(L + 1)
End If
If Len(L + 1) = 2 Then
Label13.Caption = yhdm + Format(CDate(Text5.Text), "yymmdd") + "0" + Trim(L + 1)
End If
If Len(L + 1) = 3 Then
Label13.Caption = yhdm + Format(CDate(Text5.Text), "yymmdd") + Trim(L + 1)
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
Text5.Text = Date
DTPicker3.Value = Date
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


Private Sub Command15_Click()
On Error Resume Next

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
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
DataCombo7.Enabled = False

Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
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
Forma171.DataCombo1 = DataCombo1
Forma171.Show
End Sub

Private Sub Command8_Click()
wwdm = 4
Formc344.Show
End Sub

Private Sub Command9_Click()
'On Error Resume Next
If Option1.Value = True Then
Timer1.Enabled = False
ProgressBar1.Visible = False

Adodc14.RecordSource = "select isnull(count(顺序号),0) from jgmx where 单号='" & Label13.Caption & "'"
Adodc14.Refresh
If Not Adodc14.Recordset.EOF Then
hs = Adodc14.Recordset.Fields(0)
If hs > 0 Then
If hs / 5 = Int(hs / 5) Then
ZS = hs / 5
Else
ZS = Int(hs / 5) + 1
End If
Adodc14.RecordSource = "select isnull(sum(匹数),0),isnull(sum(数量),0) from jgmx where 单号='" & Label13.Caption & "'"
Adodc14.Refresh
ps = Adodc14.Recordset.Fields(0)
fhsl = Format(Adodc14.Recordset.Fields(1), "#0.00")
For i = 0 To ZS - 1
Call CPCKOutadodcToExcel(Adodc13, Adodc12, Label13.Caption, i * 5 + 1, i * 5 + 5, ps, fhsl)
Next

End If
End If
sql1 = "update jgmx set dy='2' where 单号='" & Label13.Caption & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Exit Sub
End If


If Option2.Value = True Then
Timer1.Enabled = False
ProgressBar1.Visible = False

Adodc15.RecordSource = "select isnull(count(顺序号),0) from jgmx where 单号='" & Label13.Caption & "'"
Adodc15.Refresh
If Not Adodc15.Recordset.EOF Then
hs = Adodc15.Recordset.Fields(0)
If hs > 0 Then
If hs / 5 = Int(hs / 5) Then
ZS = hs / 5
Else
ZS = Int(hs / 5) + 1
End If
zhy = ZS - 1
Adodc15.RecordSource = "select isnull(sum(金额),0) from jgmx where 单号='" & Label13.Caption & "'"
Adodc15.Refresh
je = Format(Adodc15.Recordset.Fields(0), "#0.00")
For i = 0 To ZS - 1
If i = zhy Then
Call CPCKDJ(Adodc13, Adodc12, Label13.Caption, i * 5 + 1, i * 5 + 5)
Else
Call CPCKDJF(Adodc13, Adodc12, Label13.Caption, i * 5 + 1, i * 5 + 5)
End If
Next

End If
End If
sql1 = "update jgmx set dy='2' where 单号='" & Label13.Caption & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Exit Sub
Option3.Value = True
End If

If Option3.Value = True Then
Timer1.Enabled = False
ProgressBar1.Visible = False
Adodc15.RecordSource = "select isnull(sum(金额),0) from jgmx where 单号='" & Label13.Caption & "'"
Adodc15.Refresh
je = Format(Adodc15.Recordset.Fields(0), "#0.00")
Call CPCKQKT(Adodc14, Label13.Caption, je)
Option1.Value = True
End If

sql1 = "update jgmx set dy='2' where 单号='" & Label13.Caption & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic


Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
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
On Error Resume Next
If DataCombo4.Text = "" Then Exit Sub

gygh = DataCombo4

Adodc25.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc25.RecordSource = "SELECT 加工单位,锅号,品名,颜色,匹数,数量,单价,金额,日期,单号 FROM JGMX WHERE 锅号='" & DataCombo4.Text & "'"
Adodc25.Refresh
If DataCombo4.Text = "" Then Exit Sub


'            Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
'            Adodc18.RecordSource = "select max(重量) as zl from kpd where 锅号='" & DataCombo4.Text & "'"
'            Adodc18.Refresh
'                 If Adodc18.Recordset.EOF Then
'                   ' MsgBox ("计划部或下活处有失误！！")
'                    Exit Sub
'                 End If
'             A = Adodc18.Recordset.Fields("zl")    '把最大重量复制给变量A
'             Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
'
 '            Adodc18.RecordSource = "select SUM(重量) as zl1,SUM(匹数) as zl2 from kpd where 锅号='" & DataCombo4.Text & "'"   '''统计重量
 '            Adodc18.Refresh
 '               If Adodc18.Recordset.EOF Then
 '                  '  MsgBox ("计划部或下活处有失误！！")
 '                    Exit Sub
 '               End If
 '           c1 = Adodc18.Recordset.Fields("zl1")    '把总重量复制给变量C
 '           c2 = Adodc18.Recordset.Fields("zl2")    '把总匹数复制给变量C
 '           Adodc19.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
'
'            Adodc19.RecordSource = "select * from kpd where 锅号='" & DataCombo4.Text & "' And 重量 = '" & A & "'"
'            Adodc19.Refresh
'                  If Adodc19.Recordset.EOF Then
'                    '  MsgBox ("下活处有失误！！")
''                      Exit Sub
'                  End If
'
'            DataCombo1.Text = Adodc19.Recordset.Fields(0)   ''客户
'            DataCombo5.Text = Format(c1, "#0.0")            ''重量
'            DataCombo3.Text = Adodc19.Recordset.Fields(8)   ''颜色
'            DataCombo2.Text = Adodc19.Recordset.Fields(3)   ''品名
'            DataCombo11.Text = Adodc19.Recordset.Fields(13) ''款号
'            DataCombo12.Text = Adodc19.Recordset.Fields(10) ''技术要求
'            DataCombo16.Text = Adodc19.Recordset.Fields(1)  ''单号
'            Text7.Text = c2
'            Adodc19.RecordSource = "select 布区 from v_kpd_bq WHERE 锅号='" & DataCombo4 & "'"
'            Adodc19.Refresh
'            If Not Adodc19.Recordset.EOF Then
'            Text9.Text = Adodc19.Recordset.Fields(0)       ''布区
'            Else
'            Text9.Text = ""       ''布区
'            End If

End Sub

Private Sub DataCombo4_Click(Area As Integer)
On Error Resume Next
If DataCombo4.Text = "" Then Exit Sub
gygh = DataCombo4

Adodc25.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc25.RecordSource = "SELECT 加工单位,锅号,品名,颜色,匹数,数量,单价,金额,日期,单号 FROM JGMX WHERE 锅号='" & DataCombo4.Text & "'"
Adodc25.Refresh
'If DataCombo4.Text = "" Then Exit Sub


'            Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
'            Adodc18.RecordSource = "select max(重量) as zl from kpd where 锅号='" & DataCombo4.Text & "'"
'            Adodc18.Refresh
'                 If Adodc18.Recordset.EOF Then
'                   ' MsgBox ("计划部或下活处有失误！！")
'                    Exit Sub
'                 End If
'             A = Adodc18.Recordset.Fields("zl")    '把最大重量复制给变量A
'             Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
'
'             Adodc18.RecordSource = "select SUM(重量) as zl1,SUM(匹数) as zl2 from kpd where 锅号='" & DataCombo4.Text & "'"   '''统计重量
'             Adodc18.Refresh
'                If Adodc18.Recordset.EOF Then
'                   '  MsgBox ("计划部或下活处有失误！！")
'                     Exit Sub
'                End If
'            c1 = Adodc18.Recordset.Fields("zl1")    '把总重量复制给变量C
'            c2 = Adodc18.Recordset.Fields("zl2")    '把总匹数复制给变量C
'            Adodc19.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
'
'            Adodc19.RecordSource = "select * from kpd where 锅号='" & DataCombo4.Text & "' And 重量 = '" & A & "'"
'            Adodc19.Refresh
'                  If Adodc19.Recordset.EOF Then
'                    '  MsgBox ("下活处有失误！！")
'                      Exit Sub
'                  End If
'
'            DataCombo1.Text = Adodc19.Recordset.Fields(0)
'            DataCombo5.Text = Format(c1, "#0.0")
'            DataCombo3.Text = Adodc19.Recordset.Fields(8)
'            DataCombo2.Text = Adodc19.Recordset.Fields(3)
'            DataCombo11.Text = Adodc19.Recordset.Fields(13)
'            DataCombo12.Text = ""
'            DataCombo16.Text = Adodc19.Recordset.Fields(1)
'            Text7.Text = c2
'            Adodc19.RecordSource = "select 布区 from v_kpd_bq WHERE 锅号='" & DataCombo4 & "'"
'            Adodc19.Refresh
 '           If Not Adodc19.Recordset.EOF Then
'            Text9.Text = Adodc19.Recordset.Fields(0)       ''布区
'            Else
'            Text9.Text = ""       ''布区
'            End If

End Sub

Private Sub dataCombo4_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub DataCombo5_Change()
If Option4.Value = True Then
DataCombo7.Text = Format(Val(DataCombo5.Text) * Val(Text8.Text), "#0.00")
End If
End Sub

Private Sub dataCombo5_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub


Private Sub Command1_Click()
'On Error Resume Next

If DataCombo1.Text = "" Then
MsgBox ("请输入客户!")
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


If DataCombo5.Text = "" Then DataCombo5.Text = 0
If DataCombo6.Text = "" Then DataCombo6.Text = 0
If DataCombo7.Text = "" Then DataCombo7.Text = 0
If Text8.Text = "" Then Text8.Text = 0

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc3.RecordSource = "select * from jgmx where 锅号='" & DataCombo4.Text & "' and 品名='" & DataCombo2.Text & "' and 颜色='" & DataCombo3 & "' and 加工类别='" & Combo1.Text & "'"
Adodc3.Refresh

If Not Adodc3.Recordset.EOF Then
If MsgBox("此锅号已开，请确认，是否继续？", vbYesNo) = vbNo Then Exit Sub
End If

Adodc23.RecordSource = "select * from yj_qfts where 客户='" & DataCombo1.Text & "'"
Adodc23.Refresh

If Not Adodc23.Recordset.EOF Then
If Adodc23.Recordset.Fields(3) > Adodc23.Recordset.Fields(2) And Adodc23.Recordset.Fields(3) < Adodc23.Recordset.Fields(1) Then
MsgBox ("客户" + DataCombo1.Text + "已达欠费上限预警,请及时付款,否则将不能出库！")
End If

If Adodc23.Recordset.Fields(3) >= Adodc23.Recordset.Fields(1) Then
MsgBox ("客户" + DataCombo1.Text + "已达欠费上限预警,不能出库！")
Exit Sub
End If
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cpfhcz1('" & DataCombo1.Text & "','" & DataCombo2.Text & "','" & DataCombo3.Text & "','" & DataCombo4.Text & "','" & DataCombo5.Text & "','" & Text8.Text & "','" & DataCombo7.Text & "','" & Text5.Text & "','" & DataCombo9.Text & "','" & Text9 & "','" & DataCombo11.Text & "','" & DataCombo12.Text & "','" & DataCombo13.Text & "','" & Label13.Caption & "','" & Combo1.Text & "','1','1','" & Text7.Text & "','" & DataCombo16.Text & "','','','','" & DataCombo17.Text & "',null,'" & DataCombo15 & "','" & Text10 & "','" & Text11(2) & "','" & Text11(0) & "','" & Text11(1) & "','" & Text11(3) & "')" ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel

Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc13.RecordSource = "SELECT 锅号 FROM JGMX WHERE 单号='" & Label13.Caption & "' "
Adodc13.Refresh


'If Val(Adodc13.Recordset.RecordCount) > 5 Then
'If MsgBox("将进入下一单号,是否打印本单号?", vbYesNo) = vbYes Then
'Call Command9_Click
'End If

'Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
'Adodc16.RecordSource = "SELECT * FROM JGMX where 日期='" & Text5.Text & "'"
'Adodc16.Refresh

'If Adodc16.Recordset.EOF Then
'Label13.Caption = Format(CDate(Text5.Text), "yymmdd") + "001"
'Else
'Adodc16.RecordSource = "SELECT max(right(单号,len(单号)-6)) FROM JGMX where 日期='" & Text5.Text & "'"
'Adodc16.Refresh
'L = Val(Adodc16.Recordset.Fields(0))
'If Len(L + 1) = 1 Then
'Label13.Caption = Format(CDate(Text5.Text), "yymmdd") + "00" + Trim(L + 1)
'End If
'If Len(L + 1) = 2 Then
'Label13.Caption = Format(CDate(Text5.Text), "yymmdd") + "0" + Trim(L + 1)
'End If
'If Len(L + 1) = 3 Then
'Label13.Caption = Format(CDate(Text5.Text), "yymmdd") + Trim(L + 1)
'End If
'End If

'Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
'Adodc9.Refresh

'Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
'Adodc21.Refresh

'If Adodc21.Recordset.EOF Then
'DataCombo9.Text = 1
'DataCombo13.Text = 1
'Else
'DataCombo9.Text = Adodc21.Recordset.Fields(0) + 1
'DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
'End If

'Else

Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh
DataCombo9.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
'End If

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

Private Sub Command2_Click()
On Error Resume Next

If DataCombo1.Text = "" Then
MsgBox ("请输入客户!")
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
Adodc9.Recordset.Update
Adodc9.Refresh
DataCombo7.Enabled = False

Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh

MsgBox ("修改成功！")
Text7.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = "" '''''
Text8.Text = ""
DataCombo7.Text = ""
DataCombo9.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo14.Text = ""
DataCombo11.Text = ""
DataCombo12.Text = ""
DataCombo16.Text = ""
DataCombo4.SetFocus

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
Adodc9.Refresh
MsgBox ("删除成功！")
Adodc21.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc21.Refresh
DataCombo7.Enabled = False

sql1 = "update kpd set FH='N',zt='光坯已入库' WHERE 锅号='" & DataCombo4 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

Text7.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = "" '''''
Text8.Text = ""
DataCombo7.Text = ""
DataCombo9.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo13.Text = Adodc21.Recordset.Fields(0) + 1
DataCombo14.Text = ""
DataCombo11.Text = ""
DataCombo12.Text = ""
DataCombo16.Text = ""
DataCombo4.SetFocus

End Sub

Private Sub Command6_Click()
Unload Me
End Sub



Private Sub DataCombo6_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub


Private Sub DTPicker1_Change()
Text2.Text = DTPicker1.Value
End Sub
Private Sub DTPicker1_CloseUp()
Text2.Text = DTPicker1.Value
Text2.SetFocus
End Sub

Private Sub DTPicker2_Change()
Text3.Text = DTPicker2.Value
End Sub
Private Sub DTPicker2_CloseUp()
Text3.Text = DTPicker2.Value
Text3.SetFocus
End Sub

Private Sub DTPicker3_Change()
DataCombo8.Text = DTPicker3.Value
Text5.Text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
DataCombo8.Text = DTPicker3.Value
Text5.Text = DTPicker3.Value
Text5.SetFocus
End Sub

Private Sub Form_Load()

On Error Resume Next

Option1.Value = True
DataCombo17.Text = ""
Text5.Text = Date
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
Text7.Text = ""
DataCombo15.Text = ""
Text10.Text = ""
Option4.Value = True
Option2.Value = True
For i = 0 To 3
Text11(i) = ""
Next
Text11(3) = "公斤"
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Set RD = New ADODB.Recordset

Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc23.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc5.RecordSource = "select 简称 from khZL where ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc5.Refresh

Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"

Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc16.RecordSource = "SELECT * FROM JGMX where 日期=cast('" & Text5.Text & "' as datetime) and left(单号,1)='" & yhdm & "' and len(单号)=10"
Adodc16.Refresh

If Adodc16.Recordset.EOF Then
Label13.Caption = yhdm + Format(CDate(Text5.Text), "yymmdd") + "001"
Else
Adodc16.RecordSource = "SELECT max(right(单号,len(单号)-7)) FROM JGMX where 日期=cast('" & Text5.Text & "' as datetime) and left(单号,1)='" & yhdm & "' and len(单号)=10"
Adodc16.Refresh
L = Val(Adodc16.Recordset.Fields(0))
If Len(L + 1) = 1 Then
Label13.Caption = yhdm + Format(CDate(Text5.Text), "yymmdd") + "00" + Trim(L + 1)
End If
If Len(L + 1) = 2 Then
Label13.Caption = yhdm + Format(CDate(Text5.Text), "yymmdd") + "0" + Trim(L + 1)
End If
If Len(L + 1) = 3 Then
Label13.Caption = yhdm + Format(CDate(Text5.Text), "yymmdd") + Trim(L + 1)
End If
End If

'Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
'Adodc14.RecordSource = "SELECT * FROM JGMX WHERE dy='1' ORDER BY 顺序号"
'Adodc14.Refresh
'If Adodc14.Recordset.EOF Then
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh
'Else
'Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
'Adodc9.RecordSource = "SELECT * FROM JGMX WHERE dy='1' ORDER BY 顺序号"
'Adodc9.Refresh
'End If

Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
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

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1000
VSFlexGrid1.ColWidth(2) = 1000
VSFlexGrid1.ColWidth(3) = 1000
VSFlexGrid1.ColWidth(4) = 1000
VSFlexGrid1.ColWidth(8) = 1000
VSFlexGrid1.ColWidth(10) = 0
VSFlexGrid1.ColWidth(11) = 1000
VSFlexGrid1.ColWidth(12) = 1000
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(16) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 0
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(21) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0

'vSFlexGrid1.ColWidth(6) = 0
'vSFlexGrid1.ColWidth(7) = 0
VSFlexGrid7.ColWidth(0) = 100

Combo1.Text = ""
DTPicker3.Value = Date
DataCombo8.Text = Text5.Text
DTPicker1.Value = Date
DTPicker2.Value = Date
Text2.Text = DTPicker1.Value
Text3.Text = DTPicker2.Value
Timer1.Enabled = False

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
       If Len(gygh) > 0 Then
       fhxz = 39
       Formc143.Text1.Text = gygh
       Formc143.Text2(0).Text = Label13.Caption
       Formc143.Show
       End If
End Select
End Sub

Private Sub Label10_Click()
   beizhu = 55
   Forma112.Show
End Sub

Private Sub Label13_Change()
On Error Resume Next

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc9.RecordSource = "SELECT * FROM JGMX WHERE 单号='" & Label13.Caption & "' ORDER BY 顺序号"
Adodc9.Refresh

Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
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

Private Sub Text1_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=ERP\SQL2008"
Adodc5.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'  group by 简称"
Adodc5.Refresh
End Sub

Private Sub Text10_Change()
If Option5.Value = True Then
DataCombo7.Text = Format(Val(Text10) * Val(Text8.Text), "#0.00")
End If
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub VSFlexGrid1_DblClick()
On Error Resume Next
If Adodc9.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc9.Recordset.MoveFirst
Adodc9.Recordset.Move rs - 1
  
If Adodc9.Recordset.Fields(21) = "是" Then
  Command4.Enabled = False
  Command2.Enabled = False
  Command1.Enabled = False
  
Else
  
     DataCombo1.Text = Adodc9.Recordset.Fields(0)
     DataCombo2.Text = Adodc9.Recordset.Fields(1)
      DataCombo3.Text = Adodc9.Recordset.Fields(2)
     DataCombo4.Text = Adodc9.Recordset.Fields(3)
       DataCombo5.Text = Adodc9.Recordset.Fields(4)
     '''''''dataCombo6.text = Adodc9.RECORDSET.Fields(5)
      Text8.Text = Adodc9.Recordset.Fields(5)
      DataCombo7.Text = Adodc9.Recordset.Fields(6)
     DataCombo8.Text = Adodc9.Recordset.Fields(7)
     DataCombo11.Text = Adodc9.Recordset.Fields(10)
     DataCombo12.Text = Adodc9.Recordset.Fields(11)
         DataCombo13.Text = Adodc9.Recordset.Fields(12)
         DataCombo9.Text = Adodc9.Recordset.Fields(12)
       DataCombo14.Text = Adodc9.Recordset.Fields(9)
       Text6.Text = Adodc9.Recordset.Fields(13)
       Combo1.Text = Adodc9.Recordset.Fields(14)
       KL = Adodc9.Recordset.Fields(15)
       Text5.Text = Adodc9.Recordset.Fields(7)
       DTPicker3.Value = Adodc9.Recordset.Fields(7)
 Text7.Text = Adodc9.Recordset.Fields(17)
 Text9.Text = Adodc9.Recordset.Fields(9)
     DataCombo15.Text = Adodc9.Recordset.Fields(24)
     DataCombo16.Text = Adodc9.Recordset.Fields(18)
     DataCombo17.Text = Adodc9.Recordset.Fields(22)
     Text10 = Adodc9.Recordset.Fields(25)
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
If Option4.Value = True Then
DataCombo7.Text = Format(Val(DataCombo5.Text) * Val(Text8.Text), "#0.00")
End If
If Option5.Value = True Then
DataCombo7.Text = Format(Val(Text10) * Val(Text8.Text), "#0.00")
End If
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Timer1_Timer()
If JDBAR = 100 Then
Call Command9_Click
Exit Sub
End If
ProgressBar1.Value = JDBAR
JDBAR = JDBAR + 10
End Sub


