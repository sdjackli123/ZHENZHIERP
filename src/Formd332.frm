VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formd332 
   BackColor       =   &H00C0E0FF&
   Caption         =   "流卡工序设置"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
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
      Left            =   3360
      TabIndex        =   34
      Text            =   "Text3"
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "校正"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6240
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   3360
      Top             =   9960
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "追加"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "工序追加"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "工序确定"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   1110
      ItemData        =   "Formd332.frx":0000
      Left            =   480
      List            =   "Formd332.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      ItemData        =   "Formd332.frx":0004
      Left            =   11280
      List            =   "Formd332.frx":0006
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text13 
      Height          =   2895
      Left            =   10800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Formd332.frx":0008
      Top             =   6000
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   6120
      Width           =   1210
   End
   Begin VB.ListBox List3 
      Height          =   1950
      ItemData        =   "Formd332.frx":000F
      Left            =   5880
      List            =   "Formd332.frx":0011
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   6720
      Width           =   1935
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      ItemData        =   "Formd332.frx":0013
      Left            =   8040
      List            =   "Formd332.frx":0015
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   4440
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formd332.frx":0017
      Height          =   390
      Left            =   10920
      TabIndex        =   3
      Top             =   5280
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "工序编号"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formd332.frx":002C
      Height          =   2415
      Left            =   480
      TabIndex        =   17
      Top             =   1320
      Width           =   9615
      _cx             =   16960
      _cy             =   4260
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1320
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
      Left            =   1440
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1680
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8160
      Top             =   10680
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   8280
      Top             =   10680
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8520
      Top             =   10680
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formd332.frx":0041
      Height          =   2655
      Left            =   480
      TabIndex        =   18
      Top             =   6240
      Width           =   2895
      _cx             =   5106
      _cy             =   4683
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
      FormatString    =   $"Formd332.frx":0056
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
   Begin VB.Label Label6 
      Caption         =   "倍数"
      Height          =   375
      Left            =   3360
      TabIndex        =   35
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "全选"
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
      Left            =   6960
      TabIndex        =   30
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入锅号"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   29
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号信息"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   28
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号序号信息"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   27
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "工序信息"
      Height          =   495
      Index           =   3
      Left            =   10560
      TabIndex        =   26
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "工序内容"
      Height          =   2895
      Left            =   10560
      TabIndex        =   25
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "工序编号"
      Height          =   495
      Index           =   4
      Left            =   10560
      TabIndex        =   24
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染色工序信息"
      Height          =   375
      Index           =   5
      Left            =   5640
      TabIndex        =   23
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "染色工序"
      Height          =   2175
      Left            =   5640
      TabIndex        =   22
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
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
      Height          =   495
      Left            =   7440
      TabIndex        =   21
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "所选工序"
      Height          =   375
      Index           =   7
      Left            =   8040
      TabIndex        =   20
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "工序信息"
      Height          =   375
      Index           =   6
      Left            =   480
      TabIndex        =   19
      Top             =   5880
      Width           =   1935
   End
End
Attribute VB_Name = "Formd332"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim dhxx As String
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command10_Click()
If MsgBox("工序已选择，确认追加此类工序吗？", vbYesNo) = vbNo Then Exit Sub

If Text13.Text = "" Then
MsgBox ("请选择工序")
Exit Sub
End If

For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "update kpd set gx=isnull(gx,'')+'" & Text13.Text & "' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

For Q = 0 To List4.ListCount - 1

If List4.Selected(Q) = True Then
gxmc = Mid(List4.List(Q), InStr(List4.List(Q), "-") + 1) ' 取出 '-' 之后的数据
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "ghgxlr('" & Text1.Text & "','" & List1.List(i) & "','" & List4.List(Q) & "','1','" & gxmc & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
End If
Next
End If
Next
Adodc1.Refresh
Adodc6.Refresh

sql2 = "insert into czrz(日期,锅号,操作,内容,功能) VALUES('" & Now & "','" & Text1.Text & "','" & yhm & "','" & Text13.Text & "','工序追加')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

MsgBox ("追加成功！")
End Sub

Private Sub Command11_Click()
If Text13.Text = "" Then
ll = Text13.Text
Else
ll = Text13.Text + "-"
End If
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
If InStr(ll, Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1)) = 0 Then
ll = ll + Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1) + "-"
End If
End If
Next
Text13.Text = Mid(ll, 1, Len(ll) - 1)
End Sub

Private Sub Command12_Click()
If MsgBox("确定校正吗？", vbYesNo) = vbNo Then Exit Sub
For i = 1 To VSFlexGrid2.Rows - 1
If VSFlexGrid2.Cell(flexcpChecked, i, 2) = 1 Then
bs = Val(Text3)
sql1 = "UPDATE ghgx SET 倍数='" & bs & "' WHERE 锅号='" & Text1 & "' and 序号='" & VSFlexGrid2.TextMatrix(i, 1) & "' and 工序='" & VSFlexGrid2.TextMatrix(i, 2) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
Adodc6.RecordSource = "select 序号,工序,倍数 from ghgx where 锅号='" & Text1.Text & "' order by 序号,工序"
Adodc6.Refresh

    With VSFlexGrid2
        .Editable = flexEDKbdMouse
        .Cell(flexcpChecked, 1, 2, .Rows - 1, 2) = 2
        End With
End Sub

Private Sub Command2_Click()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 锅号,ip as 序号,品名,色别,匹数,重量,备注,gx as 工序 from kpd where 锅号='" & Text1.Text & "' order by ip"
Adodc1.Refresh

Adodc6.RecordSource = "select 序号,工序,倍数 from ghgx where 锅号='" & Text1.Text & "' order by 序号,工序"
Adodc6.Refresh
If Adodc1.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc1.Recordset.MoveFirst
List1.Clear
Do While Not Adodc1.Recordset.EOF
List1.AddItem Trim(Adodc1.Recordset.Fields(1))
Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub Command3_Click()
If MsgBox("工序已选择，确认此类设置吗？", vbYesNo) = vbNo Then Exit Sub

If Text13.Text = "" Then
MsgBox ("请选择工序")
Exit Sub
End If
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql1 = "update kpd set pb='',yd='',mr='',ts='',zd='',hg='',qr='',xdx='',ddx='',scbh='0000',ztbh='0000',zt='计划' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

sql2 = "delete from ghgx  where 锅号='" & Text1.Text & "' and 序号='" & List1.List(i) & "' and  工序 not between '1000' and '6000'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

sclc = ""   '''''''''''''''''''''''''''''清楚流程
gx = ""
For Q = 0 To List4.ListCount - 1
If List4.Selected(Q) = True Then
gxbh = Mid(List4.List(Q), 1, InStr(List4.List(Q), "-") - 1)
gxmc = Mid(List4.List(Q), InStr(List4.List(Q), "-") + 1) ' 取出 '-' 之后的数据
If Len(gx) > 0 Then
gx = gx + "-" + Mid(List4.List(Q), 1, InStr(List4.List(Q), "-") - 1)
Else
gx = gx + Mid(List4.List(Q), 1, InStr(List4.List(Q), "-") - 1)
End If

If Val(Mid(List4.List(Q), 1, InStr(List4.List(Q), "-") - 1)) < 1000 Or Val(Mid(List4.List(Q), 1, InStr(List4.List(Q), "-") - 1)) > 6000 Then
If Len(sclc) > 0 Then
sclc = sclc + "-" + Mid(List4.List(Q), InStr(List4.List(Q), "-") + 1)
Else
sclc = sclc + Mid(List4.List(Q), InStr(List4.List(Q), "-") + 1)
End If
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "ghgxlr('" & Text1.Text & "','" & List1.List(i) & "','" & gxbh & "','1','" & gxmc & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

If Val(gxbh) > 1 And Val(gxbh) < 500 Then
sql1 = "update kpd set PB='N' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

If Val(gxbh) > 500 And Val(gxbh) < 1000 Then
sql1 = "update kpd set YD='N' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

'If Val(gxbh) = 1001 Then
'sql1 = "update kpd set dr='N',mr='' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
'MsgBox (1001)
'End If

'If Val(gxbh) = 1002 Then
'sql1 = "update kpd set mr='N',dr='' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
'End If

'If Val(gxbh) = 1003 Then
'sql1 = "update kpd set dr='N',mr='N' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
'End If

If Val(gxbh) > 6000 And Val(gxbh) <= 6500 Then
sql1 = "update kpd set TS='N' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If


If Val(gxbh) > 6500 And Val(gxbh) <= 7000 Then
sql1 = "update kpd set ZD='N' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

If Val(gxbh) > 7000 And Val(gxbh) <= 7500 Then
sql1 = "update kpd set HG='N' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

If Val(gxbh) >= 7500 And Val(gxbh) <= 8000 Then
sql1 = "update kpd set QR='N' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

If Val(gxbh) > 8000 And Val(gxbh) <= 9000 Then
sql1 = "update kpd set XDX='N' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

If Val(gxbh) > 9000 And Val(gxbh) < 9999 Then
sql1 = "update kpd set DDX='N' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
End If
Next
If Len(sclc) > 0 Then

sql1 = "update kpd set mr='" & sclc & "' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"

'sql3 = "insert into dhgx(单号,序号,工序,倍数) select '" & dhxx & "',序号,工序,'1' from ghgx where 锅号='" & Text1.Text & "' and 序号='" & List1.List(i) & "'"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic

'RD.Open sql3, conn, adOpenStatic, adLockOptimistic
End If

If Len(gx) > 0 Then
sql1 = "update kpd set gx='" & gx & "' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

End If
Next

Adodc7.RecordSource = "select 工序 from ghgx where 锅号='" & Text1 & "' order by 工序"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
sql1 = "update ghgx  set 起始='" & Now & "' where 锅号='" & Text1 & "' and 工序='" & Adodc7.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

sql2 = "insert into czrz(日期,锅号,操作,内容,功能) VALUES('" & Now & "','" & Text1.Text & "','" & yhm & "','" & gx & "','工序设置')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(8) = 3200
MsgBox ("设置成功！")
Adodc6.Refresh
End Sub

Private Sub Command4_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = False
Next
End Sub

Private Sub Command5_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = True
Next
End Sub

Private Sub Command6_Click()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 工艺编号,工序名称 from GYSHD where 工序其它系数<>'0' and 工艺编号 not between '1001' and  '6000' GROUP BY 工艺编号,工序名称 order by 工艺编号"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List2.Clear
Do While Not Adodc2.Recordset.EOF
List2.AddItem Adodc2.Recordset.Fields(0) + "-" + Trim(Adodc2.Recordset.Fields(1))
Adodc2.Recordset.MoveNext
Loop
End Sub

Private Sub Command7_Click()
On Error Resume Next
ll = ""
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
ll = ll + Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1) + "-"
End If
Next
Text13.Text = Mid(ll, 1, Len(ll) - 1)
End Sub

Private Sub Command8_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Command9_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If DataCombo1.Text = "" Then Exit Sub
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 工序内容 from gybh  where 工序编号='" & DataCombo1.Text & "'"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
Text13.Text = ""
Else
Text13.Text = Adodc4.Recordset.Fields(0)
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

Label2.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = 1
Text13.Text = ""
DataCombo1 = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 锅号,ip as 序号,品名,色别,匹数,重量,备注,gx as 工序 from kpd where 锅号='" & Text1.Text & "' order by ip"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 工序编号,工序内容 from gybh "
Adodc3.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(8) = 3200

End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 5
GXBL = 33
FormS4.Show
End Select
End Sub

Private Sub Label3_Click()
On Error Resume Next
ll = ""
For i = 0 To List3.ListCount - 1
If List3.Selected(i) = True Then
ll = ll + Mid(List3.List(i), 1, InStr(List3.List(i), "-") - 1) + "-"
End If
Next
If Text13 = "" Then
Text13.Text = Mid(ll, 1, Len(ll) - 1)
Else
Text13.Text = Text13.Text + "-" + Mid(ll, 1, Len(ll) - 1)
End If
For i = 0 To List3.ListCount - 1
List3.Selected(i) = False
Next
End Sub

Private Sub Label5_Click()
For i = 0 To List3.ListCount - 1
List3.Selected(i) = True
Next
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 4 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 锅号,ip as 序号,品名,色别,匹数,重量,备注,gx as 工序,单号 from kpd where 锅号='" & Text1.Text & "' order by ip"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
dhxx = ""
List1.Clear
Else
dhxx = Adodc1.Recordset.Fields(8)
Adodc1.Recordset.MoveFirst
List1.Clear
Do While Not Adodc1.Recordset.EOF
List1.AddItem Trim(Adodc1.Recordset.Fields(1))
Adodc1.Recordset.MoveNext
Loop
End If
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 序号,工序,倍数 from ghgx where 锅号='" & Text1.Text & "' order by 序号,工序"
Adodc6.Refresh

    With VSFlexGrid2
        .Editable = flexEDKbdMouse
        .Cell(flexcpChecked, 1, 2, .Rows - 1, 2) = 2
        End With
End If

Call Command8_Click
End Sub

Private Sub Text13_Change()
List4.Clear
i = 1
For L = 0 To Int(Len(Text13.Text) / 5)
''''''''''''''''''''''''''''''''''''''''''''''''''''
gxbh = Mid(Text13.Text, L * 4 + i, 4)
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select 工艺编号,工序名称 from GYSHD where 工艺编号='" & gxbh & "'"
Adodc5.Refresh
If Not Adodc5.Recordset.EOF Then
List4.AddItem Adodc5.Recordset.Fields(0) + "-" + Trim(Adodc5.Recordset.Fields(1))
End If
i = i + 1
Next
For i = 0 To List4.ListCount - 1
List4.Selected(i) = True
Next
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then Exit Sub
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 工艺编号,工序名称 from GYSHD where 工序其它系数 like '%'+'" & Text2.Text & "'+'%' and 工艺编号 between '1001' and  '6000' GROUP BY 工艺编号,工序名称 order by 工艺编号"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
List3.Clear
Exit Sub
End If
Adodc4.Recordset.MoveFirst
List3.Clear
Do While Not Adodc4.Recordset.EOF
List3.AddItem Adodc4.Recordset.Fields(0) + "-" + Trim(Adodc4.Recordset.Fields(1))
Adodc4.Recordset.MoveNext
Loop
For i = 0 To List3.ListCount - 1
List3.Selected(i) = True
Next
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc6.Recordset.EOF Then Exit Sub
rs = VSFlexGrid2.Row
Adodc6.Recordset.MoveFirst
Adodc6.Recordset.Move rs - 1
If MsgBox("删除" + "序号：" + Trim(Adodc6.Recordset.Fields(0)) + "工序：" + Adodc6.Recordset.Fields(1) + "吗？", vbYesNo) = vbNo Then Exit Sub
sql2 = "delete from ghgx  where 锅号='" & Text1.Text & "' and 序号='" & Adodc6.Recordset.Fields(0) & "' and 工序='" & Adodc6.Recordset.Fields(1) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc6.Recordset.Delete
Adodc6.Refresh
End Sub
