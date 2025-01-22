VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms509 
   BackColor       =   &H00C0E0FF&
   Caption         =   "车间报表"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "产量打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   22200
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   1920
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   18000
      Top             =   13200
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
      Bindings        =   "Forms509.frx":0000
      Height          =   8415
      Left            =   19320
      TabIndex        =   54
      Top             =   2520
      Width           =   5655
      _cx             =   9975
      _cy             =   14843
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
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "脱水统计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "染色统计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "排布统计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "产量打印"
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
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "能耗登记"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   24120
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   11400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "报表类别"
      Height          =   855
      Left            =   10560
      TabIndex        =   45
      Top             =   1320
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "生产"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   47
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "染色"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6360
      Top             =   0
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8280
      Top             =   10800
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "产量打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   8520
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "个人打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1680
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4440
      Top             =   10920
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
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   32
      Text            =   "Text8"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   1
      Left            =   7440
      TabIndex        =   31
      Text            =   "Text8"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   30
      Text            =   "Text2"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   7440
      TabIndex        =   29
      Text            =   "Text2"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   28
      Text            =   "Text8"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1095
      Left            =   8880
      TabIndex        =   13
      Top             =   120
      Width           =   4095
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "分工序"
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   39
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "总工序"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "班次"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色号"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单号"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "操作"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5160
      Top             =   10920
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
   Begin VB.TextBox Text1111 
      Height          =   270
      Left            =   4560
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Forms509.frx":0015
      Left            =   480
      List            =   "Forms509.frx":002B
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
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
      Height          =   735
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
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
      Height          =   735
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   423034881
      CurrentDate     =   40055
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   423034881
      CurrentDate     =   40055
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms509.frx":0053
      Height          =   5895
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   17895
      _cx             =   31565
      _cy             =   10398
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
      Bindings        =   "Forms509.frx":0068
      Height          =   5175
      Left            =   240
      TabIndex        =   37
      Top             =   9120
      Width           =   17895
      _cx             =   31565
      _cy             =   9128
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forms509.frx":007D
      Height          =   3135
      Left            =   19320
      TabIndex        =   43
      Top             =   11040
      Width           =   4575
      _cx             =   8070
      _cy             =   5530
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
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "正在染色："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   58
      Top             =   8640
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "已染色完成"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   57
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "产量明细"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   20280
      TabIndex        =   56
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "班次"
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
      Left            =   6960
      TabIndex        =   48
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "分工序"
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
      Left            =   8520
      TabIndex        =   42
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   4
      Left            =   7200
      TabIndex        =   36
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   35
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   2
      Left            =   7200
      TabIndex        =   34
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   3
      Left            =   7920
      TabIndex        =   33
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "操作"
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
      Left            =   3720
      TabIndex        =   22
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   3
      Left            =   3720
      TabIndex        =   21
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
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
      Left            =   480
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
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
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择工序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
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
      Index           =   14
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      Index           =   17
      Left            =   2760
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Forms509"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2, R1, R2 As Integer
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Dim cxbl As Integer
Private Sub Command1_Click()
Call CLBB(VSFlexGrid2, 2, Combo1.Text + "产量")
End Sub

Private Sub Command10_Click()
Forms500.Show
End Sub

Private Sub Command11_Click()
Call CLBB(VSFlexGrid4, 2, Combo1.Text + "产量")
End Sub

Private Sub Command2_Click()
Call CLBB(VSFlexGrid3, 2, Combo1.Text + "产量")
End Sub

Private Sub Command3_Click()
If Adodc2.Recordset.EOF And Option1(0).value = True Then
MsgBox ("没有产量汇总 不能登记能耗")
Exit Sub
End If
Forms519.Text1(1) = VSFlexGrid2.TextMatrix(2, 2)
Forms519.Text1(2) = VSFlexGrid2.TextMatrix(2, 3)
Forms519.Text1(3) = VSFlexGrid2.TextMatrix(2, 4)
Forms519.Text1(4) = VSFlexGrid2.TextMatrix(2, 5)
Forms519.Text1(5) = VSFlexGrid2.TextMatrix(2, 6)
Forms519.Text1(6) = VSFlexGrid2.TextMatrix(2, 7)
Forms519.Text1(10) = VSFlexGrid2.TextMatrix(2, 1)
Forms519.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Call CLDY(VSFlexGrid1, Text1(4) + "班次产量", VSFlexGrid2)
End Sub

Private Sub Command6_Click()
    On Error Resume Next

    Dim sql1 As String
    sql1 = ""

    If Check2(1).value = 1 Then
        sql1 = sql1 + "锅号 like '%'+'" & Text1(1).Text & "'+'%' and "
    End If

    If Check2(2).value = 1 Then
        sql1 = sql1 + "单号 like '%'+'" & Text1(0) & "'+'%' and "
    End If

    If Check2(3).value = 1 Then
        sql1 = sql1 + "操作员 like '%'+'" & Text1(3).Text & "'+'%' and "
    End If

    If Check2(4).value = 1 Then
        Dim t1 As String
        Dim t2 As String
        t1 = Format(Trim(DTPicker2.value) + Space(2) + Text2(0) + ":" + Text2(1) + ":" + Text2(2), "yyyy-MM-dd hh:mm:ss")
        t2 = Format(Trim(DTPicker3.value) + Space(2) + Text8(0) + ":" + Text8(1) + ":" + Text8(2), "yyyy-MM-dd hh:mm:ss")
        sql1 = sql1 + "cast(CONVERT(varchar,时间, 120) as datetime) between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
    End If

    If Check2(0).value = 1 Then
        sql1 = sql1 + "班次 like '%'+'" & Text1(4).Text & "'+'%' and "
    End If

    If Check2(5).value = 1 Then
        sql1 = sql1 + "色别 like '%'+'" & Text1(2).Text & "'+'%' and "
    End If

    If Check2(6).value = 1 Then
        sql1 = sql1 + "总工序 like '%'+'" & Combo1 & "'+'%' and "
    End If

    If Check2(7).value = 1 Then
        sql1 = sql1 + "工序 like '%'+'" & Text1(5).Text & "'+'%' and "
    End If

    If sql1 = "" Then
        MsgBox ("请选择查询条件")
        Exit Sub
    End If
    sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

    If Option1(1).value = True Then
        Adodc1.RecordSource = "select 车台,客户,锅号,品名,色别,班次,工艺编号,工序,操作员,时间, 匹数,工序产量,工资系数,工序工资,(case when 总工序 like '%染色%' then '染色' else 总工序 end) as 总工序 FROM v_cjcl where (" + sql1 + ") and 工艺编号 not between '0080' and '0200'  order by 时间"
        Adodc1.Refresh
        Adodc4.RecordSource = "select distinct 锅号,重量,单价,Round(cast(isnull(重量,0)as real)* 单价,2) as 收费金额, (case when 总工序 like '%染色%' then '染色' else 总工序 end) as 总工序 FROM v_cjcl_dj where (" + sql1 + ") and 工艺编号 not between '0080' and '0200' order by 锅号"
        Adodc4.Refresh
        Adodc2.RecordSource = "select 客户,锅号,色别,班次,操作员,时间,工序,总工序,round(sum(cast(isnull(工序产量,0) as real)),2) as 总产量,round(sum(cast(isnull(工序产量,0) as real))/len(操作员)*4,2) as 分产量,round(sum(cast(isnull(工序工资,0) as real)),2) as 总工资,round(round(sum(cast(isnull(工序工资,0) as real)),2)/len(操作员)*4,2) as 个人工资 FROM v_cjcl where (" + sql1 + ") and len(操作员)<>0 and 工艺编号 not between '0080' and '0200' group by 客户,锅号,色别,班次,工序,操作员,时间,总工序"
        Adodc2.Refresh
        Adodc3.RecordSource = "select (case when 总工序 like '%染色%' then '染色' else 总工序 end),round(sum(cast(isnull(工序产量,0) as real)),2) as 班次产量,round(sum(cast(isnull(工序工资,0) as real)),2) as 总工资 FROM v_cjcl where (" + sql1 + ") and 工艺编号 not between '0080' and '0200'   group by (case when 总工序 like '%染色%' then '染色' else 总工序 end)"
        Adodc3.Refresh
    End If

    If Option1(0).value = True Then
        Adodc1.RecordSource = "select distinct 车台,客户,锅号,品名,色别,班次,工艺编号,工序,操作员,时间,匹数,班次产量,工资系数,工序工资,总工序 FROM v_cjcl where (" + sql1 + ") and 工艺编号 between '1001' and '6000' and (工序 like '%后处理%'  or 工序 like '%出缸%') order by 时间"
        Adodc1.Refresh
        
        Adodc2.RecordSource = "SELECT DISTINCT t1.车台, t1.客户, t1.锅号, t1.品名, t1.色别, t1.班次, t1.工艺编号, t1.工序, t1.操作员, t1.时间, t1.匹数, t1.班次产量, t1.工资系数, t1.工序工资, t1.总工序 " & _
                              "FROM v_cjcl t1 " & _
                              "INNER JOIN (SELECT 锅号, MAX(工艺编号) AS max_工艺编号 " & _
                                          "FROM v_cjcl " & _
                                          "WHERE (" & sql1 & ") " & _
                                          "AND 工艺编号 BETWEEN '1001' AND '6000' " & _
                                          "AND 工序 NOT LIKE '%后处理%' " & _
                                          "AND 工序 NOT LIKE '%出缸%' " & _
                                          "AND 总工序 LIKE '%染色%' " & _
                                          "GROUP BY 锅号) t2 " & _
                              "ON t1.锅号 = t2.锅号 AND t1.工艺编号 = t2.max_工艺编号 " & _
                              "ORDER BY t1.时间"
        Adodc2.Refresh
        
        Adodc3.RecordSource = "select (case when 总工序 like '%染色%' then '染色' else 总工序 end),round(sum(cast(isnull(工序产量,0) as real)),2) as 班次产量,round(sum(cast(isnull(工序工资,0) as real)),2) as 总工资 FROM v_cjcl where (" + sql1 + ") and 工艺编号 between '1001' and '6000'   group by (case when 总工序 like '%染色%' then '染色' else 总工序 end)"
        Adodc3.Refresh
        Adodc4.RecordSource = "select distinct 锅号,重量,单价,Round(cast(isnull(重量,0)as real)* 单价,2) as 收费金额, (case when 总工序 like '%染色%' then '染色' else 总工序 end) as 总工序 FROM v_cjcl_dj where (" + sql1 + ") and 工艺编号 not between '0080' and '0200'  and (工序 like '%后处理%'  or 工序 like '%出缸%') order by 锅号"
        Adodc4.Refresh
    End If


VSFlexGrid1.ColWidth(1) = 500
VSFlexGrid1.ColWidth(2) = 800
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(5) = 1200
VSFlexGrid1.ColWidth(8) = 2000
VSFlexGrid1.ColWidth(10) = 2000
VSFlexGrid1.ColWidth(15) = 2000



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
VSFlexGrid2.RowHeight(i) = 400
Next
End If
VSFlexGrid2.ColWidth(1) = 500
VSFlexGrid2.ColWidth(3) = 1200
VSFlexGrid2.ColWidth(5) = 1200
VSFlexGrid2.ColWidth(8) = 2000
VSFlexGrid2.ColWidth(10) = 2000
VSFlexGrid2.ColWidth(15) = 2000

If VSFlexGrid3.Rows > 1 Then
For i = 1 To VSFlexGrid3.Rows - 1
VSFlexGrid3.RowHeight(i) = 400
Next
End If

If VSFlexGrid4.Rows > 1 Then
For i = 1 To VSFlexGrid4.Rows - 1
VSFlexGrid4.RowHeight(i) = 400
Next
End If
VSFlexGrid4.ColWidth(2) = 1000
VSFlexGrid4.ColWidth(4) = 1500
VSFlexGrid2.SubtotalPosition = flexSTAbove
VSFlexGrid2.Subtotal flexSTSum, 0, 9, , &HC0C0&
VSFlexGrid2.Subtotal flexSTSum, 0, 10, , &HC0C0&
VSFlexGrid2.Subtotal flexSTSum, 0, 12, , &HC0C0&
VSFlexGrid2.Subtotal flexSTCount, 0, 2, , &HC0C0&

VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 0, 14, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 12, , &HC0C0&
VSFlexGrid1.Subtotal flexSTCount, 0, 3, , &HC0C0&

VSFlexGrid4.SubtotalPosition = flexSTBelow
VSFlexGrid4.Subtotal flexSTSum, 0, 2, , &HC0C0&
VSFlexGrid4.Subtotal flexSTSum, 0, 4, , &HC0C0&
VSFlexGrid4.Subtotal flexSTCount, 0, 1, , &HC0C0&

End Sub

Private Sub Command7_Click()
Forms501.Show
End Sub

Private Sub Command8_Click()
 Dim t1 As String
    Dim t2 As String
    
   t1 = Format(Trim(DTPicker2.value) + Space(2) + Text2(0) + ":" + Text2(1) + ":" + Text2(2), "yyyy-MM-dd hh:mm:ss")
        t2 = Format(Trim(DTPicker3.value) + Space(2) + Text8(0) + ":" + Text8(1) + ":" + Text8(2), "yyyy-MM-dd hh:mm:ss")

    Call CLBB(VSFlexGrid1, 11, Combo1.Text & "产量" & t1 & "  ---- " & t2 & " ")
End Sub

Private Sub Command9_Click()
Forms504.Show
End Sub

Private Sub Form_Load()
Combo1.Text = ""
For i = 0 To 5
Text1(i).Text = ""
Next
DTPicker3.value = Date
DTPicker2.value = Date
Option1(1).value = True
Text2(0) = "00"
Text2(1) = "00"
Text2(2) = "00"
cxbl = 1
Text8(0) = "23"
Text8(1) = "59"
Text8(2) = "59"
Text1(4) = "染色甲班"
Check2(4).value = 1
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Command3.Visible = False
Adodc1.CommandTimeout = 10000
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 800
For i = 2 To 5
VSFlexGrid1.ColWidth(i) = 1500
Next

VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid3.ColWidth(0) = 200

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
       Case 4
       YGBL = 11
       Forms546.Show
End Select
End Sub

Private Sub Label2_Click()
Text1(4) = "染色甲班"
End Sub

Private Sub Label2_DblClick()
Text1(4) = "染色乙班"
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
       Case 0
Command3.Visible = True
Command5.Visible = True
Label5.Visible = True
Label6.Visible = True
       Case 1
Command3.Visible = False
Command5.Visible = False
Label5.Visible = False
Label6.Visible = False
End Select
End Sub

Private Sub Timer1_Timer()
If cxbl >= 3 Then
Call Command6_Click
Timer1.Enabled = False
cxbl = 1
Else
cxbl = cxbl + 1
End If
End Sub

Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
S1 = VSFlexGrid1.RowSel
R1 = VSFlexGrid1.ColSel
End Sub

Private Sub VSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
S2 = VSFlexGrid1.RowSel
L = 0
For i = S1 To S2
L = L + Val(VSFlexGrid1.TextMatrix(i, R1))
Next
End Sub

