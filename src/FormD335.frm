VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FormD335 
   BackColor       =   &H00C0E0FF&
   Caption         =   "染色流程设置"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   14685
   WindowState     =   2  'Maximized
   Begin VB.ListBox List4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      ItemData        =   "FormD335.frx":0000
      Left            =   7320
      List            =   "FormD335.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   15
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ListBox List3 
      Height          =   1950
      ItemData        =   "FormD335.frx":0004
      Left            =   3960
      List            =   "FormD335.frx":0006
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认"
      Height          =   495
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text13 
      Height          =   2895
      Left            =   10680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "FormD335.frx":0008
      Top             =   6000
      Width           =   3495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
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
      ItemData        =   "FormD335.frx":000F
      Left            =   10440
      List            =   "FormD335.frx":0011
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1110
      ItemData        =   "FormD335.frx":0013
      Left            =   360
      List            =   "FormD335.frx":0015
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "工序确定"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   3255
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FormD335.frx":0017
      Height          =   390
      Left            =   10800
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
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
      Bindings        =   "FormD335.frx":002C
      Height          =   2415
      Left            =   360
      TabIndex        =   17
      Top             =   1440
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
      Left            =   1200
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1320
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1560
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8040
      Top             =   10800
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
      Left            =   8160
      Top             =   10800
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
      Left            =   8400
      Top             =   10800
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
      Bindings        =   "FormD335.frx":0041
      Height          =   2655
      Left            =   360
      TabIndex        =   18
      Top             =   6360
      Width           =   3015
      _cx             =   5318
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
      FormatString    =   $"FormD335.frx":0056
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
      Left            =   5040
      TabIndex        =   30
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "工序信息"
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   29
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "所选工序"
      Height          =   375
      Index           =   7
      Left            =   7320
      TabIndex        =   28
      Top             =   4200
      Width           =   2055
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
      Left            =   5520
      TabIndex        =   27
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "染色工序"
      Height          =   2175
      Left            =   3720
      TabIndex        =   26
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染色工序信息"
      Height          =   375
      Index           =   5
      Left            =   3720
      TabIndex        =   25
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "工序编号"
      Height          =   495
      Index           =   4
      Left            =   10440
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "工序内容"
      Height          =   2895
      Left            =   10440
      TabIndex        =   23
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "工序信息"
      Height          =   495
      Index           =   3
      Left            =   10440
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号序号信息"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   21
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号信息"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入锅号"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   19
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "FormD335"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Len(Text1.Text) > 4 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 锅号,ip as 序号,品名,色别,匹数,重量,备注,gx as 工序 from kpd where 锅号='" & Text1.Text & "' order by ip"
Adodc1.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 序号,工序 from ghgx where 锅号='" & Text1.Text & "' order by 序号,工序"
Adodc6.Refresh
End If

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

'sql2 = "delete from ghgx  where 锅号='" & Text1.text & "' and 序号='" & List1.List(i) & "' and 工序 between '1004' and '6000'"
'RD.Open sql2, conn, adOpenStatic, adLockOptimistic

For Q = 0 To List4.ListCount - 1
If List4.Selected(Q) = True Then
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "ghgxlr('" & Text1.Text & "','" & List1.List(i) & "','" & List4.List(Q) & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
End If
Next

End If
Next
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

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(8) = 3200

End Sub

Private Sub Label3_Click()
On Error Resume Next

ll = ""
For i = 0 To List3.ListCount - 1
If List3.Selected(i) = True Then
ll = ll + Mid(List3.List(i), 1, InStr(List3.List(i), "-") - 1) + "-"
End If
Next
Text13.Text = Text13.Text + Mid(ll, 1, Len(ll) - 1)
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
If Len(Text1.Text) > 4 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 锅号,ip as 序号,品名,色别,匹数,重量,备注,gx as 工序 from kpd where 锅号='" & Text1.Text & "' order by ip"
Adodc1.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 序号,工序 from ghgx where 锅号='" & Text1.Text & "' order by 序号,工序"
Adodc6.Refresh
End If
End Sub

Private Sub Text13_Change()
List4.Clear
i = 1
For L = 0 To Int(Len(Text13.Text) / 5)
List4.AddItem Mid(Text13.Text, L * 4 + i, 4)
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
Adodc6.Refresh
End Sub
