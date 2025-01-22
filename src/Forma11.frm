VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma11 
   BackColor       =   &H00C0E0FF&
   Caption         =   "生产计划"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18960
   Icon            =   "Forma11.frx":0000
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc27 
      Height          =   375
      Left            =   12000
      Top             =   12240
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
   Begin MSAdodcLib.Adodc Adodc26 
      Height          =   375
      Left            =   9600
      Top             =   12120
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
   Begin VB.CommandButton Command24 
      BackColor       =   &H008080FF&
      Caption         =   "修改保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   13560
      TabIndex        =   125
      Text            =   "Text20"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   19200
      TabIndex        =   123
      Text            =   "Text19"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      Height          =   2655
      Left            =   13440
      MultiLine       =   -1  'True
      TabIndex        =   121
      Text            =   "Forma11.frx":440A
      Top             =   4560
      Width           =   10095
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   12120
      TabIndex        =   120
      Text            =   "Text17"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00C0C0FF&
      Caption         =   "毛坯库存"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00C0C0FF&
      Caption         =   "配缸"
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
      Left            =   21840
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0FF&
      Caption         =   "取消"
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
      Left            =   21840
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认"
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
      Left            =   21840
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "来布方式"
      Height          =   735
      Left            =   13080
      TabIndex        =   111
      Top             =   1080
      Visible         =   0   'False
      Width           =   3495
      Begin VB.OptionButton Option4 
         Caption         =   "一次"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   113
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "多次"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   112
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Index           =   2
      Left            =   11280
      TabIndex        =   106
      Text            =   "Text16"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Index           =   1
      Left            =   9120
      TabIndex        =   104
      Text            =   "Text16"
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   103
      Text            =   "Text16"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0FF&
      Caption         =   "取消"
      Height          =   375
      Left            =   27720
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0FF&
      Caption         =   "完成"
      Height          =   375
      Left            =   27240
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   25800
      TabIndex        =   97
      Text            =   "Text15"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   6960
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "Forma11.frx":4411
      Left            =   5760
      List            =   "Forma11.frx":442D
      TabIndex        =   95
      Text            =   "Combo1"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0FF&
      Caption         =   "卡号复制"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   5520
      Width           =   1335
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
      Bindings        =   "Forma11.frx":4451
      Height          =   4575
      Left            =   480
      TabIndex        =   41
      Top             =   7320
      Width           =   23055
      _cx             =   40666
      _cy             =   8070
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
      MergeCells      =   1
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
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   15000
      TabIndex        =   93
      Text            =   "Text14"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   6000
      TabIndex        =   91
      Text            =   "Text12"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H008080FF&
      Caption         =   "印花打印"
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H008080FF&
      Caption         =   "返修打印"
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
      Left            =   21600
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H008080FF&
      Caption         =   "返修"
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
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H008080FF&
      Caption         =   "最新锅号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   3360
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Forma11.frx":4466
      Left            =   20280
      List            =   "Forma11.frx":4482
      TabIndex        =   86
      Text            =   "Combo1"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "委外加工信息"
      Height          =   2295
      Left            =   4440
      TabIndex        =   76
      Top             =   7440
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox Text10 
         Height          =   1095
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Text            =   "Forma11.frx":44A6
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   1080
         TabIndex        =   78
         Text            =   "Text8"
         Top             =   360
         Width           =   2655
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
         Bindings        =   "Forma11.frx":44AC
         Height          =   1695
         Left            =   5160
         TabIndex        =   81
         Top             =   360
         Width           =   4215
         _cx             =   7435
         _cy             =   2990
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
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF80&
         Caption         =   "委外取消"
         Height          =   495
         Left            =   3960
         TabIndex        =   84
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "备注信息"
         Height          =   1095
         Index           =   19
         Left            =   120
         TabIndex        =   82
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF80&
         Caption         =   "委外打印"
         Height          =   495
         Left            =   3960
         TabIndex        =   80
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF80&
         Caption         =   "委外确认"
         Height          =   495
         Left            =   3960
         TabIndex        =   79
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "委外客户"
         Height          =   495
         Index           =   18
         Left            =   120
         TabIndex        =   77
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   75
      Text            =   "Text2"
      Top             =   1440
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9960
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "流程更新"
      Height          =   540
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9120
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   9480
      Top             =   240
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   25
      Text            =   "Text9"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   "传票打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "毛坯信息"
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Text            =   "Text7"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   16920
      TabIndex        =   22
      Text            =   "Text6"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "排缸卡打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
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
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H008080FF&
      Caption         =   "新锅号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4080
      Width           =   1455
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
      ItemData        =   "Forma11.frx":44C2
      Left            =   18720
      List            =   "Forma11.frx":44D5
      TabIndex        =   16
      Text            =   "Combo2"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   20640
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command9 
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
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "流卡工序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染色"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0000C0C0&
      Caption         =   "定型"
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0000C0C0&
      Caption         =   "返修"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "锅号复制"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   10920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "Forma11.frx":44F7
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H008080FF&
      Caption         =   "计划信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc25 
      Height          =   330
      Left            =   7080
      Top             =   10200
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
      Left            =   7560
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
      Left            =   7440
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
      Left            =   7560
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
      Left            =   7800
      Top             =   10440
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
      Left            =   7800
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
   Begin MSAdodcLib.Adodc Adodc19 
      Height          =   375
      Left            =   10800
      Top             =   10080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc Adodc18 
      Height          =   330
      Left            =   8880
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
      Left            =   9600
      Top             =   10080
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
      Left            =   7200
      Top             =   10200
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
      Left            =   8040
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
      Left            =   8760
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
      Left            =   7560
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
      Left            =   8040
      Top             =   10200
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
      Left            =   8760
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
      Left            =   8520
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
      Left            =   9240
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
      Left            =   8520
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
      Left            =   8880
      Top             =   10200
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
      Left            =   10200
      Top             =   10320
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
      Left            =   8040
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
      Left            =   8280
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
      Left            =   8040
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
      Left            =   8280
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
      Left            =   8400
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
   Begin MSDataListLib.DataCombo DataCombo8 
      Height          =   330
      Left            =   22680
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo6"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   13320
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Forma11.frx":44FE
      Height          =   330
      Left            =   5520
      TabIndex        =   5
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "车台编号"
      BoundColumn     =   ""
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma11.frx":4514
      Height          =   360
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5400
      TabIndex        =   26
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330825729
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5400
      TabIndex        =   27
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330891265
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   375
      Left            =   16920
      TabIndex        =   28
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330891265
      CurrentDate     =   39177
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   1
      Left            =   4200
      TabIndex        =   29
      Top             =   2520
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   2
      Left            =   10680
      TabIndex        =   32
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   3
      Left            =   12360
      TabIndex        =   33
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   4
      Left            =   16080
      TabIndex        =   35
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   5
      Left            =   17400
      TabIndex        =   36
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   6
      Left            =   7080
      TabIndex        =   30
      Top             =   2520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   7
      Left            =   6600
      TabIndex        =   39
      Top             =   3960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Forma11.frx":4529
      Height          =   330
      Index           =   8
      Left            =   14280
      TabIndex        =   34
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Forma11.frx":453F
      Height          =   330
      Index           =   9
      Left            =   4200
      TabIndex        =   42
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "hx"
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Forma11.frx":4555
      Height          =   330
      Index           =   10
      Left            =   22680
      TabIndex        =   43
      Top             =   2760
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "Forma11.frx":456B
      Height          =   330
      Left            =   1440
      TabIndex        =   44
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "xm"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma11.frx":4580
      Height          =   2175
      Left            =   22920
      TabIndex        =   45
      Top             =   4320
      Visible         =   0   'False
      Width           =   4095
      _cx             =   7223
      _cy             =   3836
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forma11.frx":4596
      Height          =   255
      Left            =   4680
      TabIndex        =   40
      Top             =   7440
      Visible         =   0   'False
      Width           =   7815
      _cx             =   13785
      _cy             =   450
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
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   3360
      TabIndex        =   107
      Top             =   8400
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   257753089
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   3360
      TabIndex        =   108
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   257753089
      CurrentDate     =   36892
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "附加费单价"
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
      Index           =   29
      Left            =   13560
      TabIndex        =   124
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "布头"
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
      Index           =   28
      Left            =   19320
      TabIndex        =   122
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "加工单价"
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
      Index           =   27
      Left            =   12120
      TabIndex        =   119
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFF00&
      Caption         =   "客户名称"
      Height          =   495
      Left            =   480
      TabIndex        =   118
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   110
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   109
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "库存"
      Height          =   375
      Index           =   26
      Left            =   10680
      TabIndex        =   105
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单据"
      Height          =   375
      Index           =   25
      Left            =   5640
      TabIndex        =   102
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   24
      Left            =   8520
      TabIndex        =   101
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "缸号"
      Height          =   375
      Index           =   23
      Left            =   24720
      TabIndex        =   98
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "原卡号"
      Height          =   375
      Index           =   22
      Left            =   5760
      TabIndex        =   96
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "收费项目"
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
      Index           =   21
      Left            =   15000
      TabIndex        =   92
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "打印卡号"
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
      Index           =   20
      Left            =   20280
      TabIndex        =   85
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "委外加工"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   73
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "色号"
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
      Index           =   17
      Left            =   9120
      TabIndex        =   72
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF80&
      Caption         =   "标签打印"
      Height          =   375
      Left            =   15240
      TabIndex        =   71
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "锅  号"
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
      Left            =   480
      TabIndex        =   70
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   9
      Left            =   16920
      TabIndex        =   69
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   68
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   67
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   6
      Left            =   4200
      TabIndex        =   66
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "毛胚幅宽(寸)"
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
      Left            =   10680
      TabIndex        =   65
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光胚幅宽(cm)"
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
      Left            =   12360
      TabIndex        =   64
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "匹数"
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
      Left            =   16080
      TabIndex        =   63
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "重量（公斤）"
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
      Left            =   17400
      TabIndex        =   62
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "克重"
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
      Left            =   14280
      TabIndex        =   61
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "染色要求"
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
      Left            =   6600
      TabIndex        =   60
      Top             =   3480
      Width           =   5415
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入单号"
      Height          =   375
      Index           =   3
      Left            =   21600
      TabIndex        =   59
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "整理类别"
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
      Left            =   18720
      TabIndex        =   58
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP"
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
      Index           =   12
      Left            =   20640
      TabIndex        =   57
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款  号"
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
      Index           =   13
      Left            =   480
      TabIndex        =   56
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "毛坯备注"
      Height          =   375
      Index           =   14
      Left            =   15120
      TabIndex        =   55
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请输入原锅号"
      Height          =   375
      Index           =   15
      Left            =   4200
      TabIndex        =   54
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "机台"
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
      Index           =   16
      Left            =   5520
      TabIndex        =   53
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
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
      Left            =   7080
      TabIndex        =   52
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "工序"
      Height          =   255
      Left            =   10680
      TabIndex        =   51
      Top             =   15
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "是否合缝"
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
      Left            =   4200
      TabIndex        =   50
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF80&
      Caption         =   "产量统计"
      Height          =   375
      Left            =   17160
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      Caption         =   "毛坯码单"
      Height          =   375
      Left            =   16200
      TabIndex        =   48
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "坯布类型"
      Height          =   375
      Index           =   5
      Left            =   21600
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "调度员"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   46
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Forma11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public x As Integer: Public BI As Integer ''''BI PANDUAN CHURU KU BIANLIANG
Dim BA As Database: Dim rr As Integer: Public gh, k1, k2 As String: Public hg As Date: Dim BA3 As Database: Dim RD3 As Recordset
Public ZL As Single  ''''''重量变量
Rem ' 中间转换变量
Dim rs As Single: Dim RD1 As Recordset: Dim BA1 As Database: Dim c, r As Long: Dim lbj As Long
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim ll As String: Dim cdbhf As Integer
Dim plshsx As Integer
'''''''''''''''''''''''''''''''''
Dim zf As Long
Dim yf As Long
Dim sf As Long
Dim xf As Long

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



Private Sub Command10_Click()
ghcx = 1
Forma172.Show
End Sub

Private Sub Command11_Click()
Forma117.Text1(0) = Text7
Forma117.Show
End Sub

Private Sub Command12_Click()
On Error Resume Next
If Option1.value = False And Option2.value = False And Option3.value = False Then
MsgBox ("请选择备活信息")
Exit Sub
End If

If Option1.value = True Then
Adodc22.RecordSource = "select * from zxgh"
Adodc22.Refresh
If Adodc22.Recordset.EOF Then
MsgBox ("请设置最新锅号信息")
Exit Sub
End If

Adodc23.RecordSource = "select isnull(max(cast(SUBSTRING(锅号,5,5) as int)),0) as h  from kpd where substring(锅号,1,4)='" & Adodc22.Recordset.Fields(1) & "' and 锅号 NOT like '%F%' and 锅号 NOT like '%H%' and len(锅号)>4 and 锅号 not like '%-%'"
Adodc23.Refresh

Text7.Text = Adodc22.Recordset.Fields(1) + "00001"
If Adodc23.Recordset.EOF Then
Text7.Text = Adodc22.Recordset.Fields(1) + "00001"
Else
Text7.Text = Adodc22.Recordset.Fields(1) + Mid("00000", 1, 4 - Len(Trim(Val(Adodc23.Recordset.Fields(0)) + 1))) + Trim(Val(Adodc23.Recordset.Fields(0)) + 1)
End If
End If



If Option2.value = True Then

Adodc23.RecordSource = "select MAX(right(锅号,len(锅号)-6)) as h   from kpd where month(日期)=month(' " & Text6.Text & "') and year(日期)=year(' " & Text6.Text & "')  AND 锅号 like 'D%'  and 锅号 not like '%H%'"
Adodc23.Refresh
Text7.Text = "D" + Format(CDate(Text6.Text), "YYMM") + "0001"
If Adodc23.Recordset.EOF Then
Text7.Text = "D" + Format(CDate(Text6.Text), "YYMM") + "0001"
Else
Text7.Text = "D" + Format(CDate(Text6.Text), "YYMM") + Mid("0000", 1, 4 - Len(Trim(Val(Adodc23.Recordset.Fields(0)) + 1))) + Trim(Val(Adodc23.Recordset.Fields(0)) + 1)
End If
End If

If Option3.value = True Then

Adodc23.RecordSource = "select MAX(right(锅号,len(锅号)-6)) as h  from kpd where month(日期)=month(' " & Text6.Text & "') and year(日期)=year(' " & Text6.Text & "')  AND  锅号 like 'F%' and 锅号 not like '%H%'"
Adodc23.Refresh
Text7.Text = "F" + Format(CDate(Text6.Text), "YYMM") + "0001"
If Adodc23.Recordset.EOF Then
Text7.Text = "F" + Format(CDate(Text6.Text), "YYMM") + "0001"
Else
Text7.Text = "F" + Format(CDate(Text6.Text), "YYMM") + Mid("0000", 1, 4 - Len(Trim(Val(Adodc23.Recordset.Fields(0)) + 1))) + Trim(Val(Adodc23.Recordset.Fields(0)) + 1)
End If
End If

  Adodc9.RecordSource = "select isnull(max(Ip),0) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  If Adodc9.Recordset.EOF Then
  Text3 = 1
  Else
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
  End If
Option4(0).value = True
End Sub


Private Sub Command13_Click()
Formd332.Text1 = Text7
Formd332.Show
End Sub

Private Sub Command14_Click()
Forma104.Show
End Sub

Private Sub Command15_Click()
Call lcd22fx(Adodc14, Text7.Text, Combo1)
End Sub

Private Sub Command16_Click()
Call lcd22yh(Adodc14, Text7.Text)
End Sub

Private Sub Command17_Click()
On Error Resume Next
If Combo3 = "" Then
MsgBox ("请输入原卡号")
Exit Sub
End If

If Text7.Text = "" Then
MsgBox ("请输入锅号")
Exit Sub
End If

If MsgBox("要复制卡号" + Combo1 + "吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "insert into kpd(客户名称,单号,锅号,色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,日期,备注,技术要求,IP,标签,kp,kp1,CKY,负责人,pb,rs,ts,xdx,ddx,fh,色名,卡号,dr,gx,zt,hx,mr,编号) select 客户名称,单号,锅号,色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,'" & Date & "',备注,技术要求,IP,标签,'N','N',CKY,负责人,'Y','N','N','N','N','N',色名,'" & Combo1 & "',dr,gx,'计划',hx,mr,编号 from kpd where 锅号='" & Text7.Text & "' and 卡号='" & Combo3 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc8.RecordSource = "select 日期,客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,色名 as 色号,标签 as 合约号,备注 as 染色要求,技术要求 AS 克重,类别,CKY as 毛坯备注,车台,GX AS 工序,HX AS 是否合缝,单号,DR AS 特别注明,卡号,图案  from kpd where 锅号='" & Text7.Text & "' order by IP"
Adodc8.Refresh

End Sub

Private Sub Command18_Click()
Command18.Enabled = False
sql1 = "update sczy_x set 排布='Y' where 缸号='" & Text15.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Command18.Enabled = True
End Sub

Private Sub Command19_Click()
Command19.Enabled = False
sql1 = "update sczy_x set 排布='N' where 缸号='" & Text15.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Command19.Enabled = True
End Sub

Private Sub Command20_Click()
If Option4(1).value = True Then
If Len(Text16(0)) <> 8 Then
MsgBox ("毛坯库入库单据号错误")
Exit Sub
End If
DataCombo4(4) = Val(DataCombo4(4))
Text16(2) = Val(Text16(2))
mpgh = Trim(Text7) + Trim(Text3)
sql1 = "insert into mpbh(单据号,入库序号,布类,毛胚匹数,毛胚重量,锅号,序号,客户,出库日期,缸号) VALUES('" & Text16(0) & "','" & Text16(1) & "','" & DataCombo4(1) & "','" & DataCombo4(4) & "','" & Text16(2) & "','" & Text7 & "','" & Text3 & "','" & DataCombo1 & "','" & Text6 & "','" & mpgh & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("毛坯出库成功！")
End If
Timer1.Enabled = True
End Sub

Private Sub Command21_Click()
If MsgBox("确定取消 备布出库吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete from mpbh where 锅号='" & Text7 & "' and 序号='" & Text3 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("返库成功！")
End Sub

Private Sub Command22_Click()
Forma103.DataCombo1(1) = Text7
Forma103.Show
End Sub

Private Sub Command23_Click()
mmkc = 1
Forma17.Check1(0).value = 0  '''日期选上
Forma17.Check1(2).value = 1  ''客户选上
Forma17.Check1(4).value = 1  ''库存选上
Forma17.DataCombo1 = DataCombo1
Forma17.Show
End Sub


Private Sub Command24_Click()
If Text7.Text = "" Then
MsgBox ("请输入锅号")
Exit Sub
End If
' 检查Text8是否为空，如果为空，设置一个默认值
    If Trim(Text20.Text) = "" Then
        Text20.Text = "0"  ' 根据字段类型和业务需求设置合适的默认值
    End If
    ' 检查Text8是否为空，如果为空，设置一个默认值
    If Trim(Text17.Text) = "" Then
        Text17.Text = "0"  ' 根据字段类型和业务需求设置合适的默认值
    End If
 Adodc8.Recordset.Fields(10) = Text1.Text
  Adodc8.Recordset.Fields(3) = Text3.Text
  Adodc8.Recordset.Fields(0) = Text6.Text
Adodc8.Recordset.Fields(2) = Text7.Text
Adodc8.Recordset.Fields(11) = Text9.Text
 Adodc8.Recordset.Fields(21) = Combo1.Text
  Adodc8.Recordset.Fields(14) = Combo2.Text
 Adodc8.Recordset.Fields(1) = DataCombo1.Text

 Adodc8.Recordset.Fields(4) = DataCombo4(1).Text
 Adodc8.Recordset.Fields(5) = DataCombo4(2).Text
 Adodc8.Recordset.Fields(6) = DataCombo4(3).Text
 Adodc8.Recordset.Fields(7) = DataCombo4(4).Text
 Adodc8.Recordset.Fields(8) = DataCombo4(5).Text
 Adodc8.Recordset.Fields(9) = DataCombo4(6).Text
 Adodc8.Recordset.Fields(12) = DataCombo4(7).Text
 Adodc8.Recordset.Fields(13) = DataCombo4(8).Text
 Adodc8.Recordset.Fields(18) = DataCombo4(9).Text
 Adodc8.Recordset.Fields(16) = DataCombo2.Text
Adodc8.Recordset.Fields(19) = Text16(0).Text
Adodc8.Recordset.Fields(25) = Text19.Text
Adodc8.Recordset.Fields(24) = Text18.Text
Adodc8.Recordset.Fields(23) = Text17.Text
Adodc8.Recordset.Fields(26) = Text20.Text
Adodc8.Recordset.Update
Adodc8.Refresh
Adodc8.RecordSource = "select 日期,客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,色名 as 色号,标签 as 合约号,备注 as 染色要求,技术要求 AS 克重,类别,CKY as 毛坯备注,车台,GX AS 工序,HX AS 是否合缝,单号,DR AS 特别注明,卡号,图案,单价,幅宽明细,布头,附加费单价  from kpd where 锅号='" & Text7.Text & "' order by 日期"
Adodc8.Refresh
Adodc9.RecordSource = "select isnull(max(Ip),0) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  If Adodc9.Recordset.EOF Then
  Text3 = 1
  Else
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
  End If
 Adodc8.RecordSource = "select 日期,客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,色名 as 色号,标签 as 合约号,备注 as 染色要求,技术要求 AS 克重,类别,CKY as 毛坯备注,车台,GX AS 工序,HX AS 是否合缝,单号,DR AS 特别注明,卡号,图案,单价,幅宽明细,布头,附加费单价  from kpd where 锅号='" & Text7.Text & "' order by 日期"
Adodc8.Refresh
Call gssx
End Sub


Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Text11.Text = "" Then
MsgBox ("请输入原锅号")
Exit Sub
End If


If Text7.Text = "" Then
MsgBox ("请输入锅号")
Exit Sub
End If


If MsgBox("要复制原锅号" + Text11.Text + "新锅号为" + Text7.Text + "吗？", vbYesNo) = vbNo Then Exit Sub

If Combo1 = "" Then
sql1 = "insert into kpd(客户名称,单号,锅号,色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,日期,备注,技术要求,IP,标签,kp,kp1,CKY,负责人,pb,rs,ts,xdx,ddx,fh,色名,卡号,dr,gx,zt,hx,mr,编号) select 客户名称,单号,'" & Text7.Text & "',色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,'" & Date & "',备注,技术要求,IP,标签,'N','N',CKY,负责人,'Y','N','N','N','N','N',色名,卡号,dr,'','计划',hx,'','" & Text7.Text & "'+ cast(IP as nvarchar(2)) from kpd where 锅号='" & Text11.Text & "'"
'sql2 = "insert into ghgx(锅号,序号,工序,倍数) select '" & Text7 & "',序号,工序,倍数 from ghgx where 锅号='" & Text11.Text & "'"
sql3 = "insert into dhjgxm(缸号,序号,项目,单价) select 编号,IP,'染色费',0 from kpd where 锅号='" & Text7.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
'RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
Else
sql1 = "insert into kpd(客户名称,单号,锅号,色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,日期,备注,技术要求,IP,标签,kp,kp1,CKY,负责人,pb,rs,ts,xdx,ddx,fh,色名,卡号,dr,gx,zt,hx,mr,幅宽明细,布头,附加费单价,编号) select 客户名称,单号,'" & Text7.Text & "',色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,'" & Date & "',备注,技术要求,IP,标签,'N','N',CKY,负责人,'Y','N','N','N','N','N',色名,卡号,dr,'','计划',hx,'',幅宽明细,布头,附加费单价 ,'" & Text7.Text & "'+ cast(IP as nvarchar(2)) from kpd where 锅号='" & Text11.Text & "' and 卡号='" & Combo1 & "'"
'sql2 = "insert into ghgx(锅号,序号,工序,倍数) select '" & Text7 & "',序号,工序,倍数 from ghgx where 锅号='" & Text11.Text & "'"
sql3 = "insert into dhjgxm(缸号,序号,项目,单价) select 编号,IP,'染色费',0 from kpd where 锅号='" & Text7.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
'RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
End If

Adodc8.RecordSource = "select 日期,客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,色名 as 色号,标签 as 合约号,备注 as 染色要求,技术要求 AS 克重,类别,CKY as 毛坯备注,车台,GX AS 工序,HX AS 是否合缝,单号,DR AS 特别注明,卡号,图案,幅宽明细,布头,附加费单价  from kpd where 锅号='" & Text7.Text & "' order by IP"
Adodc8.Refresh
Call Text7_Change
Call gssx
End Sub

Private Sub Command5_Click()
On Error Resume Next

If Combo1 = "" Then
Call pgk1(Adodc14, Text7.Text)
Else
Call pgk1(Adodc14, Text7.Text)
End If

End Sub

Private Sub Command6_Click()
If DataCombo1.Text = "" Then
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 日期,客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,色名 as 色号,标签 as 合约号,备注 as 染色要求,技术要求 AS 克重,类别,CKY as 毛坯备注,车台,GX AS 工序,HX AS 是否合缝,单号,DR AS 特别注明,卡号,图案,幅宽明细,布头,附加费单价  from kpd where cast(convert(nvarchar,日期,23) as datetime) between cast('" & DTPicker3.value & "' as datetime) and cast('" & DTPicker4.value & "' as datetime)  order by 锅号,日期"
Adodc8.Refresh
Else
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 日期,客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,色名 as 色号,标签 as 合约号,备注 as 染色要求,技术要求 AS 克重,类别,CKY as 毛坯备注,车台,GX AS 工序,HX AS 是否合缝,单号,DR AS 特别注明,卡号,图案,单价,幅宽明细,布头,附加费单价  from kpd where 客户名称='" & DataCombo1.Text & "' and cast(convert(nvarchar,日期,23) as datetime) between cast('" & DTPicker3.value & "' as datetime) and cast('" & DTPicker4.value & "' as datetime)  order by 锅号,日期"
Adodc8.Refresh
End If
End Sub

Private Sub Command7_Click()
mmkc = 1
Formc25.Check1(0).value = 0
Formc25.Check1(2).value = 1
Formc25.Check1(4).value = 1
Formc25.DataCombo1 = DataCombo1
Formc25.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next

'If DataCombo5.Text = "" Then
'MsgBox ("请选择负责人！")
'Exit Sub
'End If

If DataCombo1.Text = "" Then
MsgBox ("请输入客户！")
Exit Sub
End If


If Text7.Text = "" Then
MsgBox ("请输入锅号！")
Exit Sub
End If

If Text3.Text = "" Then
MsgBox ("请输入序号！")
Exit Sub
End If

Adodc7.RecordSource = "select * from kpd where 锅号='" & Text7 & "' and ip='" & Text3 & "'"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
MsgBox ("已有此锅号和ip信息，程序禁止这样操作，建议一个缸号的信息是一个订单内的信息，否则请按照新缸号操作，才有并缸处理？")
Exit Sub
End If

Adodc11.RecordSource = "select 锅号,count(distinct 色别+色名) from kpd  where 日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) group by 锅号  having(count(distinct 色别+色名))>1"
Adodc11.Refresh
If Not Adodc11.Recordset.EOF Then
MsgBox ("锅号" + Adodc11.Recordset.Fields(0) + "有不同的色别  不应该这样 请检查修改")
End If

If Option4(0).value = True Then
If Val(Text16(2)) < Val(DataCombo4(5)) Then
MsgBox ("计划量大约库存量，禁止操作")
Exit Sub
End If
End If

If Option4(0).value = True Then
    Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "mpkpd1('" & DataCombo1.Text & "','" & DataCombo8.Text & "','" & Text7.Text & "','" & DataCombo4(1).Text & "','" & DataCombo4(2).Text & "','" & DataCombo4(3).Text & "','" & DataCombo4(4).Text & "','" & DataCombo4(5).Text & "','" & DataCombo4(6).Text & "','" & DataCombo4(7).Text & "','" & DataCombo4(8).Text & "','" & Text3.Text & "','" & Now & "','" & Text9.Text & "','" & DataCombo2.Text & "','" & DataCombo5.Text & "','" & DataCombo3.Text & "','N','N','" & Combo2.Text & "','N','N','N','','','','N','" & Text13.Text & "','N','N','" & DataCombo4(9).Text & "','" & DataCombo4(10).Text & "','" & Text1 & "','" & Combo1 & "','" & Text16(0) & "','" & Text16(1) & "','" & Text14 & "','" & Text17 & "','" & Text18 & "','" & Text19 & "','" & Text20 & "')"  ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
End If

If Option4(1).value = True Then
    Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "mpkpd('" & DataCombo1.Text & "','" & DataCombo8.Text & "','" & Text7.Text & "','" & DataCombo4(1).Text & "','" & DataCombo4(2).Text & "','" & DataCombo4(3).Text & "','" & DataCombo4(4).Text & "','" & DataCombo4(5).Text & "','" & DataCombo4(6).Text & "','" & DataCombo4(7).Text & "','" & DataCombo4(8).Text & "','" & Text3.Text & "','" & Now & "','" & Text9.Text & "','" & DataCombo2.Text & "','" & DataCombo5.Text & "','" & DataCombo3.Text & "','N','N','" & Combo2.Text & "','N','N','N','','','','N','" & Text13.Text & "','N','N','" & DataCombo4(9).Text & "','" & DataCombo4(10).Text & "','" & Text1 & "','" & Combo1 & "','" & Text16(0) & "','" & Text16(1) & "','" & Text14 & "')"  ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
End If

Adodc8.RecordSource = "select 日期,客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,色名 as 色号,标签 as 合约号,备注 as 染色要求,技术要求 AS 克重,类别,CKY as 毛坯备注,车台,GX AS 工序,HX AS 是否合缝,单号,DR AS 特别注明,卡号,图案,单价,幅宽明细,布头,附加费单价  from kpd where 锅号='" & Text7.Text & "' order by 日期 "
Adodc8.Refresh

If VSFlexGrid4.Rows > 1 Then
For i = 1 To VSFlexGrid4.Rows - 1
VSFlexGrid4.RowHeight(i) = 600
Next
End If

  Adodc9.RecordSource = "select isnull(max(Ip),0) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  If Adodc9.Recordset.EOF Then
  Text3 = 1
  Else
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
  End If
Call gssx
'DataCombo1.Text = ""
Text1.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text14.Text = ""
Text15.Text = ""
Text17.Text = 0
For i = 0 To 2
Text16(i) = ""
Next
Text2.Text = ""
For i = 0 To 8
DataCombo4(i).Text = ""
Next
Text8.Text = ""
Text9.Text = ""
DataCombo3.Text = ""
DataCombo2.Text = ""
Text13.Text = ""
Text11.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = 0
End Sub


Private Sub Command8_Click()
''''Call lcd22(Adodc14, Text7.Text)
If Combo1 = "" Then
Call lcd2(Adodc14, Text7.Text)
Else
Call lcd222f(Adodc14, Adodc20, Text7.Text, Combo1)
End If
End Sub


Private Sub Command9_Click()
On Error Resume Next
If Text3.Text = "" Then Exit Sub
If MsgBox("确定删除 ip " + Text3.Text + " 吗？", vbYesNo) = vbNo Then Exit Sub
Adodc4.RecordSource = "select * from v_cjcl where 锅号='" & Text7 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
MsgBox ("车间生产正在进行  不能删除！")
Exit Sub
End If
Adodc8.Recordset.Delete
Adodc8.Refresh
'sql1 = "delete from ghgx where 锅号='" & Text7 & "' and 序号='" & Text3 & "'"
sql2 = "delete from mpbh where 锅号='" & Text7 & "' and 序号='" & Text3 & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
''Call Command2_Click
Call gssx
End Sub

Private Sub DataCombo1_Change()
On Error Resume Next
 ww = 0
If Text4.Text = "" Or Text5.Text = "" Then
Exit Sub
End If
RQ = CDate(Text5.Text)
  Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  Adodc9.RecordSource = "select isnull(max(Ip),0) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  If Adodc9.Recordset.EOF Then
  Text3 = 1
  Else
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
  End If

End Sub

Private Sub DataCombo1_Click(Area As Integer)
On Error Resume Next

 ww = 0
If Text4.Text = "" Or Text5.Text = "" Then
End If
RQ = CDate(Text5.Text)
op = 0.5
  Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  Adodc9.RecordSource = "select max(Ip) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  '''  Text3.Text = Adodc9.Recordset.Fields(0) + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''订单计划信息

       Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc6.RecordSource = "select * from  mpckgl3  where 客户名称='" & DataCombo1.Text & "' and 重量<>0"
       Adodc6.Refresh

End Sub

Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub dataCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub dataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo4_Change(Index As Integer)
    Select Case Index
    Case 1
    Adodc27.RecordSource = "select 备注,技术要求 from kpd where 客户名称= '" & DataCombo1.Text & "' and 品名= '" & DataCombo4(1).Text & "' "
      Adodc27.Refresh
      If Not Adodc27.Recordset.EOF Then
      DataCombo4(7).Text = Adodc27.Recordset.Fields(0)
      DataCombo4(8).Text = Adodc27.Recordset.Fields(1)
      End If
       
      
        Case 6
            ' 构建SQL查询，选择最新日期的记录
            Adodc26.RecordSource = "SELECT 色名, 单价, 附加费单价 " & _
                                   "FROM kpd " & _
                                   "WHERE 客户名称= '" & DataCombo1.Text & "' " & _
                                   "AND 色别= '" & DataCombo4(6).Text & "' " & _
                                   "AND 品名= '" & DataCombo4(1).Text & "' " & _
                                   "AND 日期 = (SELECT MAX(日期) FROM kpd WHERE 客户名称= '" & DataCombo1.Text & "' AND 色别= '" & DataCombo4(6).Text & "' and 品名 = '" & DataCombo4(1).Text & "')"
            ' 刷新Adodc26以应用新的查询
            Adodc26.Refresh
            ' 检查查询结果是否为空
            If Not Adodc26.Recordset.EOF Then
                ' 如果有记录，将结果赋值给对应的文本框
                Text1.Text = Adodc26.Recordset.Fields(0)
                
                ' 检查单价是否为Null，为Null则设置为0
                If IsNull(Adodc26.Recordset.Fields(1)) Then
                    Text17.Text = 0
                Else
                    Text17.Text = Adodc26.Recordset.Fields(1)
                End If
                
                ' 检查附加费单价是否为Null，为Null则设置为0
                If IsNull(Adodc26.Recordset.Fields(2)) Then
                    Text20.Text = 0
                Else
                    Text20.Text = Adodc26.Recordset.Fields(2)
                End If
            End If
    End Select
End Sub


Private Sub dataCombo4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub dataCombo5_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo6_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub DataCombo7_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub



Private Sub DataCombo8_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.value
End Sub


Private Sub DTPicker1_CloseUp()
Text4.Text = DTPicker1.value
Text4.SetFocus
End Sub

Private Sub DTPicker2_Change()
Text5.Text = DTPicker2.value
End Sub

Private Sub DTPicker2_CloseUp()
Text5.Text = DTPicker2.value
Text5.SetFocus
End Sub


Private Sub DTPicker5_Change()
Text6.Text = DTPicker5.value
End Sub

Private Sub DTPicker5_CloseUp()
Text6.Text = DTPicker5.value
End Sub
Private Sub Form_Load()
On Error Resume Next
cdbhf = cdbh
DataCombo8.Text = ""
Combo2.Text = "圆筒"
DTPicker1.value = Date - 10
DTPicker2.value = Date
DTPicker3.value = Date
DTPicker4.value = Date
DataCombo1.Text = ""
Text1.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text14.Text = ""
Text15.Text = ""
Text17.Text = 0
Text4.Text = Date - 10
Text5.Text = Date
Text8 = ""
Combo1 = "甲"
Combo3 = "甲"
DataCombo5.Text = ""
Option4(0).value = True
Option1.value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
plshsx = 1
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 简称  from khzl where  ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc3.Refresh
For i = 0 To 2
Text16(i) = ""
Next
Text2.Text = ""
Text3.Text = ""
Text7.Text = ""
Text9.Text = ""
DataCombo3.Text = ""
DataCombo2.Text = ""
DataCombo3.Enabled = False
Text13.Text = ""
Text11.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = 0
DTPicker5.value = Date
Text6.Text = Date
Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.RecordSource = "select xm  from fzr group by xm"
Adodc12.Refresh

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc19.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc19.RecordSource = "select distinct 车台编号 from ct order by 车台编号"
Adodc19.Refresh

Adodc20.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc26.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc27.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc21.RecordSource = "select distinct dr from kpd"
Adodc21.Refresh

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc17.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc18.RecordSource = "SELECT distinct hx FROM kpd"
Adodc18.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select xm  from ddy group by xm"
Adodc2.Refresh

Adodc22.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc23.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc24.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"



Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

  Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  Adodc9.RecordSource = "select isnull(max(Ip),0) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  If Adodc9.Recordset.EOF Then
  Text3 = 1
  Else
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
  End If


Adodc25.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc25.RecordSource = "select MC from JSYQ group by MC"
Adodc25.Refresh


DataCombo5.Text = ""


DataCombo4(4).Enabled = True
DataCombo4(5).Enabled = True




DataCombo7.Text = ""

For i = 1 To 10
DataCombo4(i).Text = ""
Next




VSFlexGrid3.ColWidth(0) = 100
VSFlexGrid3.ColWidth(1) = 1000
VSFlexGrid3.ColWidth(2) = 1200
VSFlexGrid3.ColWidth(3) = 1000
VSFlexGrid3.ColWidth(4) = 1000
VSFlexGrid3.ColWidth(5) = 1000
VSFlexGrid3.ColWidth(6) = 1000
VSFlexGrid3.ColWidth(7) = 1800


VSFlexGrid2.ColWidth(0) = 100
VSFlexGrid2.ColWidth(1) = 1500
VSFlexGrid2.ColWidth(2) = 1500
VSFlexGrid2.ColWidth(3) = 1500
VSFlexGrid2.ColWidth(4) = 1500
VSFlexGrid2.ColWidth(5) = 1500
VSFlexGrid2.ColWidth(6) = 1200
VSFlexGrid2.ColWidth(7) = 1500


VSFlexGrid4.ColWidth(0) = 100
VSFlexGrid4.ColWidth(2) = 1500
VSFlexGrid4.ColWidth(3) = 1500
VSFlexGrid4.ColWidth(4) = 500
VSFlexGrid4.ColWidth(8) = 1000
VSFlexGrid4.ColWidth(10) = 1800

ZL = 0

Text4.TabIndex = 0
End Sub

Private Sub Label11_Click()
If MsgBox("确定委外加工吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "insert into kpdwwjg(委外锅号,委外单位,委外信息) VALUES('" & Text7 & "','" & Text8 & "','" & Text10 & "')"
sql2 = "update kpd  set gz=convert(nvarchar ,getdate(),120),zt='委外' where 锅号='" & Text7 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
  Adodc15.RecordSource = "select * from kpdwwjg where 委外锅号= '" & Text7.Text & "' "
  Adodc15.Refresh
End Sub

Private Sub Label1_Click()
ysbl = 2
Forma38.Text1.Text = DataCombo4(6).Text
Forma38.Text2.Text = DataCombo1.Text
Forma38.Show
End Sub

Private Sub Label12_Click()
Call wtlcd22(Adodc14, Text7.Text)
End Sub

Private Sub Label13_Click()
If MsgBox("确定取消委外加工吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete from kpdwwjg where 委外锅号='" & Text7 & "'"
sql2 = "update kpd  set zt='计划' where 锅号='" & Text7 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
  Adodc15.RecordSource = "select * from kpdwwjg where 委外锅号= '" & Text7.Text & "' "
  Adodc15.Refresh
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 6
Forma113.Text3(0).Text = DataCombo4(1).Text
Forma113.Show
       Case 8
beizhu = 11
Forma112.Show
       Case 14
DataCombo3.Enabled = False
       Case 21
beizhu = 13
Forma112.Show
       End Select
End Sub

Private Sub Label9_Click()
If Text7 = "" Or Val(Text3) = 0 Then Exit Sub
Call mpbq(Adodc14, Text7, Text3)
End Sub

Private Sub Label2_DblClick(Index As Integer)
Select Case Index
       Case 14
DataCombo3.Enabled = True
End Select
End Sub

Private Sub Label4_Click()
FormA101.Show
End Sub

Private Sub Text12_Change()
Adodc19.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc19.RecordSource = "select distinct 车台编号 from ct where 车台编号 like '%'+'" & Text12 & "'+'%' order by 车台编号"
Adodc19.Refresh
End Sub

Private Sub Text2_Change()
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text2 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc3.Refresh
End Sub

Private Sub Timer1_Timer()
'If Text16(0) = "" Or Text16(1) = "" Then Exit Sub
'Adodc16.RecordSource = "select 库存匹数,库存重量 from v_mp_kc  where 单据号='" & Text16(0) & "' and 序号='" & Text16(1) & "' and 库存重量>0"
'Adodc16.Refresh

Adodc16.RecordSource = "select 重量 from MPCKGL3 where 客户名称='" & DataCombo1 & "' and 布类='" & DataCombo4(1) & "' and 重量>0"
Adodc16.Refresh
If Not Adodc16.Recordset.EOF Then
Text16(2) = Adodc16.Recordset.Fields(0)
Else
Text16(2) = 0
End If

'If Not Adodc16.Recordset.EOF Then
'Text16(2) = Adodc16.Recordset.Fields(1)
'Else
'Text16(2) = 0
'End If
'If Option4(1).value = True Then
'Adodc11.RecordSource = "select isnull(sum(毛胚重量),0) from mpbh  where 锅号='" & Text7 & "' and 序号='" & Text3 & "'"
'Adodc11.Refresh
'DataCombo4(5) = Adodc11.Recordset.Fields(0)
'End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If plshsx = 180 Then
plshsx = 1
Else
plshsx = plshsx + 1
End If
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
If Adodc20.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc20.Recordset.MoveFirst
Adodc20.Recordset.Move rs - 1
DataCombo4(6).Text = Adodc20.Recordset.Fields(0)

End Sub

Private Sub Label14_DblClick()
On Error Resume Next
If MsgBox("确定修改客户名称吗", vbYesNo) = vbNo Then Exit Sub
If DataCombo1 = "" Then
MsgBox ("请选择客户名称")
Exit Sub
End If
sql1 = "update kpd set 客户名称='" & DataCombo1 & "'  where 锅号='" & Text7 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc8.Refresh
End Sub

Private Sub Label5_Click()
Forms501.Text1(2) = Text7
Forms501.Show
End Sub

Private Sub Label8_Click()
Forma110.Text1(0) = Text7
Forma110.Text1(2) = Text9
Forma110.Text1(1) = DataCombo4(1)
Forma110.Show
End Sub

Private Sub VSFlexGrid2_Click()
On Error Resume Next
If Adodc15.Recordset.EOF Then Exit Sub
Text8 = Adodc15.Recordset.Fields(1)
Text10 = Adodc15.Recordset.Fields(2)
End Sub

Private Sub VSFlexGrid3_dblClick()
On Error Resume Next
If Adodc6.Recordset.EOF Then Exit Sub
rs = VSFlexGrid3.Row
Adodc6.Recordset.MoveFirst
Adodc6.Recordset.Move rs - 1
DataCombo4(1).Text = Adodc6.Recordset.Fields(1)
DataCombo4(2).Text = Adodc6.Recordset.Fields(2)
DataCombo3.Text = Adodc6.Recordset.Fields(6)
'DataCombo4(4).Text = Adodc6.Recordset.Fields(3)
DataCombo4(5).Text = Adodc6.Recordset.Fields(5)
End Sub


Private Sub VSFlexGrid4_dblClick()
If Adodc8.Recordset.EOF Then Exit Sub
rs = VSFlexGrid4.Row
cl = VSFlexGrid4.col
Adodc8.Recordset.MoveFirst
Adodc8.Recordset.Move rs - 1
If cl = 4 Then
Text1.Text = IIf(IsNull(Adodc8.Recordset.Fields(10)), "", Adodc8.Recordset.Fields(10)) ''色号
Text3.Text = IIf(IsNull(Adodc8.Recordset.Fields(3)), "", Adodc8.Recordset.Fields(3)) '''IP
Text6.Text = IIf(IsNull(Adodc8.Recordset.Fields(0)), "", Adodc8.Recordset.Fields(0)) ''日期
Text7.Text = IIf(IsNull(Adodc8.Recordset.Fields(2)), "", Adodc8.Recordset.Fields(2)) '''锅号
Text9.Text = IIf(IsNull(Adodc8.Recordset.Fields(11)), "", Adodc8.Recordset.Fields(11)) '''款号
Combo1.Text = IIf(IsNull(Adodc8.Recordset.Fields(21)), "", Adodc8.Recordset.Fields(21)) '''打印卡号
Combo2.Text = IIf(IsNull(Adodc8.Recordset.Fields(14)), "", Adodc8.Recordset.Fields(14)) ''整理类别

DataCombo1.Text = IIf(IsNull(Adodc8.Recordset.Fields(1)), "", Adodc8.Recordset.Fields(1)) ''''客户
DataCombo4(1).Text = IIf(IsNull(Adodc8.Recordset.Fields(4)), "", Adodc8.Recordset.Fields(4)) '''品名
DataCombo4(2).Text = IIf(IsNull(Adodc8.Recordset.Fields(5)), "", Adodc8.Recordset.Fields(5)) ''毛坯幅宽
DataCombo4(3).Text = IIf(IsNull(Adodc8.Recordset.Fields(6)), "", Adodc8.Recordset.Fields(6)) ''光坯幅宽
DataCombo4(4).Text = IIf(IsNull(Adodc8.Recordset.Fields(7)), "", Adodc8.Recordset.Fields(7)) ''匹数
DataCombo4(5).Text = IIf(IsNull(Adodc8.Recordset.Fields(8)), "", Adodc8.Recordset.Fields(8)) ''重量
DataCombo4(6).Text = IIf(IsNull(Adodc8.Recordset.Fields(9)), "", Adodc8.Recordset.Fields(9)) ''颜色
DataCombo4(7).Text = IIf(IsNull(Adodc8.Recordset.Fields(12)), "", Adodc8.Recordset.Fields(12)) ''染色要求
DataCombo4(8).Text = IIf(IsNull(Adodc8.Recordset.Fields(13)), "", Adodc8.Recordset.Fields(13)) ''克重
DataCombo4(9).Text = IIf(IsNull(Adodc8.Recordset.Fields(18)), "", Adodc8.Recordset.Fields(18)) ''合缝
DataCombo2.Text = IIf(IsNull(Adodc8.Recordset.Fields(16)), "", Adodc8.Recordset.Fields(16)) ''机台
Text16(0).Text = IIf(IsNull(Adodc8.Recordset.Fields(19)), "", Adodc8.Recordset.Fields(19)) ''单据
Text19.Text = IIf(IsNull(Adodc8.Recordset.Fields(25)), "", Adodc8.Recordset.Fields(25)) '''来料单位
Text18.Text = IIf(IsNull(Adodc8.Recordset.Fields(24)), "", Adodc8.Recordset.Fields(24)) ''幅宽明细
Text17.Text = IIf(IsNull(Adodc8.Recordset.Fields(23)), "", Adodc8.Recordset.Fields(23)) ''单价
Text20.Text = IIf(IsNull(Adodc8.Recordset.Fields(26)), "", Adodc8.Recordset.Fields(26)) ''附加费单价
End If

If cl = 5 Then
Forma110.Text1(0) = Adodc8.Recordset.Fields(1)
Forma110.Text1(2) = Adodc8.Recordset.Fields(9)
Forma110.Text1(1) = Adodc8.Recordset.Fields(3)
Forma110.Show
End If

End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Text7_Change()
On Error Resume Next
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 日期,客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,色名 as 色号,标签 as 合约号,备注 as 染色要求,技术要求 AS 克重,类别,CKY as 毛坯备注,车台,GX AS 工序,HX AS 是否合缝,单号,DR AS 特别注明,卡号,图案,单价,幅宽明细,布头,附加费单价  from kpd where 锅号='" & Text7.Text & "' order by 日期"
Adodc8.Refresh
  
  Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  Adodc15.RecordSource = "select * from kpdwwjg where 委外锅号= '" & Text7.Text & "' "
  Adodc15.Refresh
 
  Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  Adodc9.RecordSource = "select isnull(max(Ip),0) as bj from kpd where 锅号= '" & Text7.Text & "' "
  Adodc9.Refresh
  If Adodc9.Recordset.EOF Then
  Text3 = 1
  Else
  Text3.Text = Adodc9.Recordset.Fields(0) + 1
  End If
  
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "select 序号,工序,倍数 from ghgx where 锅号='" & Text7 & "' order by 序号,工序"
Adodc13.Refresh

Call gssx
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub sx()
If Adodc20.Recordset.EOF Then Exit Sub
Adodc20.Recordset.MoveFirst
i = 1
Do While Not Adodc20.Recordset.EOF
VSFlexGrid1.col = 3
VSFlexGrid1.Row = i
VSFlexGrid1.Text = Format(Adodc20.Recordset.Fields(2), "##0.0")
Adodc20.Recordset.MoveNext
i = i + 1
Loop
End Sub

Private Sub MSFlex()
With VSFlexGrid4
    c = .col: r = .Row    '''''C列，，R行
    If c <> 4 And c <> 1 And c <> 2 Then
    
        
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


Private Sub vSFlexGrid4_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid4.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    VSFlexGrid4.Text = Combo1111.Text
    Combo1111.Visible = False
    VSFlexGrid4.SetFocus
End If
End Sub

Private Sub Combo1111_LostFocus()
On Error Resume Next
Adodc8.Recordset.MoveFirst
Adodc8.Recordset.Move r - 1
Adodc8.Recordset.Fields(c - 1) = Combo1111.Text
Adodc8.Recordset.Update
If c = 9 And Val(Combo1111.Text) > 0 Then
sql2 = "update kpd set zt='计划',pb='" & Now & "' where 锅号='" & Text7 & "' and pb='N' and rs='N'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
Combo1111.Visible = False
VSFlexGrid4.SetFocus
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

Private Sub gssx()
If VSFlexGrid4.Rows > 1 Then
For i = 1 To VSFlexGrid4.Rows - 1
VSFlexGrid4.RowHeight(i) = 600
Next
End If
End Sub

