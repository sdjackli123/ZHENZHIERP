VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formy134 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料退库"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form33"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc15 
      Height          =   375
      Left            =   4920
      Top             =   11280
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4680
      Width           =   1335
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formy134.frx":0000
      Height          =   2175
      Left            =   720
      TabIndex        =   53
      Top             =   5640
      Width           =   22335
      _cx             =   39396
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
      Bindings        =   "Formy134.frx":0015
      Height          =   5295
      Left            =   720
      TabIndex        =   51
      Top             =   8040
      Width           =   22335
      _cx             =   39396
      _cy             =   9340
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   0
      Left            =   12480
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formy134.frx":002A
      Height          =   330
      Left            =   1920
      TabIndex        =   31
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   4320
      TabIndex        =   30
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formy134.frx":003F
      Height          =   330
      Left            =   1920
      TabIndex        =   29
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "材料名称"
      Text            =   "DataCombo1"
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   11040
      Top             =   2520
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新单据号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "材料查询"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库类查询"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   10320
      TabIndex        =   25
      Top             =   5160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8400
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7920
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8280
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
      Left            =   7320
      Top             =   10560
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
      Height          =   375
      Left            =   7920
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   375
      Left            =   7800
      Top             =   10560
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   7920
      Top             =   10560
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Left            =   7560
      Top             =   10560
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
      Height          =   330
      Left            =   7200
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   7680
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   7560
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
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   7680
      Top             =   10800
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
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   375
      Left            =   7200
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   330
      Left            =   9480
      Top             =   10560
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   1
      Left            =   11400
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   2
      Left            =   12720
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   3
      Left            =   2040
      TabIndex        =   35
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   4
      Left            =   2040
      TabIndex        =   36
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   5
      Left            =   2040
      TabIndex        =   37
      Top             =   3480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   6
      Left            =   2040
      TabIndex        =   38
      Top             =   3960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   7
      Left            =   5640
      TabIndex        =   39
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   8
      Left            =   5640
      TabIndex        =   40
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   9
      Left            =   5640
      TabIndex        =   41
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   10
      Left            =   5640
      TabIndex        =   42
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   11
      Left            =   2040
      TabIndex        =   43
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   12
      Left            =   9120
      TabIndex        =   44
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   13
      Left            =   9120
      TabIndex        =   45
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   390
      Index           =   14
      Left            =   6960
      TabIndex        =   46
      Top             =   5160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   688
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   15
      Left            =   9120
      TabIndex        =   47
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   16
      Left            =   9120
      TabIndex        =   48
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   17
      Left            =   12120
      TabIndex        =   49
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Formy134.frx":0054
      Height          =   360
      Index           =   18
      Left            =   6960
      TabIndex        =   50
      Top             =   4680
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo4"
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   19
      Left            =   5640
      TabIndex        =   52
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   20
      Left            =   12840
      TabIndex        =   54
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   21
      Left            =   9120
      TabIndex        =   56
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo4"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Formy134.frx":0069
      Height          =   3375
      Left            =   14040
      TabIndex        =   58
      Top             =   1200
      Visible         =   0   'False
      Width           =   6015
      _cx             =   10610
      _cy             =   5953
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
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
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单据号"
      Height          =   375
      Index           =   6
      Left            =   4440
      TabIndex        =   59
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "仓位"
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   57
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "批号"
      Height          =   375
      Index           =   9
      Left            =   4440
      TabIndex        =   28
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "库别"
      Height          =   375
      Index           =   16
      Left            =   10800
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "供应单位"
      Height          =   375
      Index           =   15
      Left            =   720
      TabIndex        =   26
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "退回客户"
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
      Left            =   5760
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "库类"
      Height          =   375
      Index           =   12
      Left            =   7920
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择材料"
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
      Left            =   720
      TabIndex        =   16
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择颜色"
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
      Left            =   4320
      TabIndex        =   15
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择库类"
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
      Left            =   720
      TabIndex        =   14
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "金额"
      Height          =   375
      Index           =   11
      Left            =   4440
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   10
      Left            =   7920
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "日期"
      Height          =   375
      Index           =   8
      Left            =   7920
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料名称"
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料规格"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单价"
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "数量"
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "材料单位"
      Height          =   375
      Index           =   6
      Left            =   720
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "Formy134"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BAR, c, r As Integer

Private Sub Combo1_Change()
DataCombo4(17).Text = Combo1.Text
End Sub

Private Sub Combo1_Click()
DataCombo4(17).Text = Combo1.Text
End Sub

Private Sub Command1_Click()
Adodc2.RecordSource = "select * from v_cltk WHERE 库类='" & DataCombo3.Text & "' and 材料名称 like '%'+'" & DataCombo1.Text & "'+'%' and 库存数量>0"
Adodc2.Refresh
End Sub

Private Sub Command10_Click()
If Adodc6.Recordset.EOF Then
MsgBox ("此单据号中无记录，不能打印！")
Exit Sub
End If
BAR = 1
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command11_Click()
On Error Resume Next
If MsgBox("确定删除吗？删除将不能恢复！", vbYesNo) = vbNo Then Exit Sub
Adodc6.Recordset.Delete
Adodc6.Refresh

Call Command1_Click

Adodc7.RecordSource = "SELECT 供应单位,单据号,材料名称,材料规格,材料单位,退库数量,单价,合计金额,日期,序号,单据 FROM cltk WHERE 单据号='" & DataCombo4(14).Text & "' "
Adodc7.Refresh
Adodc8.Refresh

For i = 3 To 10
DataCombo4(i).Text = ""
Next

DataCombo4(13).Text = 1
DataCombo4(13).Text = Adodc7.Recordset.Fields(9) + 1

End Sub

Private Sub Command12_Click()
On Error Resume Next
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT * FROM cltkdj"
Adodc8.Refresh

DataCombo4(14).Enabled = False
If Adodc8.Recordset.EOF Then
DataCombo4(14).Text = "00000001"
Else
uu = Val(Adodc8.Recordset.Fields(1)) + 1
DataCombo4(14).Text = Left("00000000", 8 - Len(Trim(Str(uu)))) + Trim(Str(uu))
End If

Adodc6.RecordSource = "SELECT * FROM cltk WHERE 单据号='" & DataCombo4(14).Text & "' ORDER BY 序号 desc"
Adodc6.Refresh

DataCombo4(13).Text = 1
DataCombo4(13).Text = Adodc6.Recordset.RecordCount + 1
DataCombo4(14).Enabled = False
End Sub




Private Sub Command3_Click()
Call OutadodcToExcel(VSFlexGrid1, 11, DataCombo1.Text)
End Sub

Private Sub Command4_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "材料库存")
End Sub

Private Sub Command5_Click()
Adodc2.RecordSource = "select * from v_cltk WHERE  库类='" & DataCombo3.Text & "' and 库存数量>0 ORDER BY 库类,材料名称,材料规格,材料单位"
Adodc2.Refresh
End Sub

Private Sub Command7_Click()
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub

If DataCombo4(18).Text = "" Then
MsgBox ("车间")
Exit Sub
End If

If DataCombo4(3).Text = "" Then
MsgBox ("材料名称不能为空")
Exit Sub
End If
If DataCombo4(8).Text = "" Then
MsgBox ("材料单价不能为空")
Exit Sub
End If
If DataCombo4(9).Text = "" Then
MsgBox ("材料数量不能为空")
Exit Sub
End If

If DataCombo4(16).Text = "" Then
MsgBox ("无库类！")
Exit Sub
End If

If Len(DataCombo4(14).Text) <> 8 Then
MsgBox ("单据号编码不符合规则  需要8位")
Exit Sub
End If

Adodc6.Recordset.Fields(3) = DataCombo4(3).Text
Adodc6.Recordset.Fields(4) = DataCombo4(4).Text
Adodc6.Recordset.Fields(5) = DataCombo4(5).Text
Adodc6.Recordset.Fields(8) = DataCombo4(8).Text
Adodc6.Recordset.Fields(9) = DataCombo4(9).Text
Adodc6.Recordset.Fields(10) = DataCombo4(10).Text
Adodc6.Recordset.Fields(15) = DataCombo4(16).Text
Adodc6.Recordset.Fields(16) = DataCombo4(17).Text
Adodc6.Recordset.Fields(20) = DataCombo4(12).Text
Adodc6.Recordset.Fields(21) = DataCombo4(13).Text
Adodc6.Recordset.Fields(22) = DataCombo4(18).Text
Adodc6.Recordset.Fields(23) = DataCombo4(19).Text
Adodc6.Recordset.Fields(29) = DataCombo4(14).Text ''退库单据
'Adodc6.Recordset.Fields(17) = Combo1.Text
Adodc6.Recordset.Update
Adodc6.Refresh

Call Command1_Click

Adodc7.RecordSource = "SELECT 供应单位,单据号,材料名称,材料规格,材料单位,退库数量,单价,合计金额,日期,序号,单据 FROM cltk WHERE 单据='" & DataCombo4(14).Text & "' "
Adodc7.Refresh
Adodc8.Refresh

'Call Command6_Click
'Call Command1_Click
For i = 3 To 10
DataCombo4(i).Text = ""
Next

DataCombo4(13).Text = 1
DataCombo4(13).Text = Adodc7.Recordset.Fields(9) + 1

End Sub

Private Sub Command8_Click()
On Error Resume Next
If DataCombo4(18).Text = "" Then
MsgBox ("车间")
Exit Sub
End If

If DataCombo4(3).Text = "" Then
MsgBox ("材料名称不能为空")
Exit Sub
End If
If DataCombo4(8).Text = "" Then
MsgBox ("材料单价不能为空")
Exit Sub
End If
If DataCombo4(9).Text = "" Then
MsgBox ("材料数量不能为空")
Exit Sub
End If

If DataCombo4(16).Text = "" Then
MsgBox ("无库类！")
Exit Sub
End If

If Len(DataCombo4(14).Text) <> 8 Then
MsgBox ("单据号编码不符合规则  需要8位")
Exit Sub
End If

Adodc6.Recordset.AddNew

Adodc6.Recordset.Fields(3) = DataCombo4(3).Text
Adodc6.Recordset.Fields(4) = DataCombo4(4).Text
Adodc6.Recordset.Fields(5) = DataCombo4(5).Text
Adodc6.Recordset.Fields(8) = DataCombo4(8).Text
Adodc6.Recordset.Fields(9) = DataCombo4(9).Text
Adodc6.Recordset.Fields(10) = DataCombo4(10).Text
Adodc6.Recordset.Fields(15) = DataCombo4(16).Text
Adodc6.Recordset.Fields(16) = DataCombo4(17).Text
Adodc6.Recordset.Fields(20) = DataCombo4(12).Text
Adodc6.Recordset.Fields(21) = DataCombo4(13).Text
Adodc6.Recordset.Fields(22) = DataCombo4(18).Text
Adodc6.Recordset.Fields(23) = DataCombo4(19).Text
Adodc6.Recordset.Fields(29) = DataCombo4(14).Text ''退库单据
Adodc6.Recordset.Update
Adodc6.Refresh

Call Command1_Click
Adodc7.RecordSource = "SELECT 供应单位,单据号,材料名称,材料规格,材料单位,退库数量,单价,合计金额,日期,序号,单据 FROM cltk WHERE 单据='" & DataCombo4(14).Text & "'ORDER BY 序号 desc "
Adodc7.Refresh
Adodc8.Refresh

If Adodc6.Recordset.RecordCount = 8 Then
If MsgBox("是否打印本单据？", vbYesNo) = vbNo Then
DataCombo4(14).Text = "00000001"
If Adodc8.Recordset.EOF Then
DataCombo4(14).Text = "00000001"
Else
DataCombo4(14).Text = Left("00000000", 8 - Len(Trim(Str(Adodc8.Recordset.Fields(0) + 1)))) + Trim(Str(Adodc8.Recordset.Fields(0) + 1))
End If
End If
End If
For i = 3 To 10
DataCombo4(i).Text = ""
Next
DataCombo4(13).Text = 1
DataCombo4(13).Text = Adodc7.Recordset.Fields(9) + 1
End Sub

Private Sub Command9_Click()
Unload Me
End Sub

Private Sub DataCombo3_Click(Area As Integer)
Adodc5.RecordSource = "SELECT 材料名称 FROM v_cltk WHERE 库存数量>0 AND 库类='" & DataCombo3.Text & "' GROUP BY 材料名称"
Adodc5.Refresh
End Sub

Private Sub DataCombo4_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 1
      ' adodc2.RecordSource = "SELECT * FROM KCCXHZ WHERE KCCXHZ.数量>0 AND KCCXHZ.数量>0 AND KCCXHZ.库类='" & dataCombo3.Text & "'"
       'adodc2.Refresh
      ' Call SX(adodc2, vSFlexGrid1, 7)
      ' Call SX(adodc2, vSFlexGrid1, 6)
       Case 3
       Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc15.RecordSource = "SELECT 单价,材料单位 FROM clgl WHERE 材料名称 ='" & DataCombo4(3).Text & "' ORDER BY 序号 desc"
       Adodc15.Refresh
       Case 8
       DataCombo4(10).Text = Format(Val(DataCombo4(8).Text) * Val(DataCombo4(9).Text), "#0.00")
       Case 9
       DataCombo4(10).Text = Format(Val(DataCombo4(8).Text) * Val(DataCombo4(9).Text), "#0.00")
       Case 14
       Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc6.RecordSource = "SELECT * FROM cltk WHERE 单据='" & DataCombo4(14).Text & "' ORDER BY 序号 desc"
       Adodc6.Refresh
       Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc7.RecordSource = "SELECT 供应单位,单据号,材料名称,材料规格,材料单位,退库数量,单价,合计金额,日期,序号,单据 FROM cltk WHERE 单据='" & DataCombo4(14).Text & "'ORDER BY 序号 desc "
       Adodc7.Refresh
       DataCombo4(13).Text = 1
       DataCombo4(13).Text = Adodc7.Recordset.Fields(9) + 1
End Select
End Sub


Private Sub Form_Load()
Me.Caption = Me.Caption + "操作年度 " + ljb
On Error Resume Next
Combo1.Text = ""
DTPicker1 = Date
DTPicker2 = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
For i = 0 To 21
DataCombo4(i).Text = ""
Next
DataCombo4(17).Text = ""
DataCombo4(19).Text = ""
DataCombo4(20).Text = "未"
DataCombo4(12).Text = Date
DataCombo4(13).Text = 1
ProgressBar1.Visible = False
Timer1.Enabled = False

ProgressBar1.Visible = False
Timer1.Enabled = False


Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from v_cltk WHERE 库存数量>0 order BY 库类,材料名称,材料规格,材料单位"
Adodc2.Refresh


Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select MC from CLKL  GROUP BY MC"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "select 简称 from GYS where ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc9.Refresh
Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 2600
VSFlexGrid1.ColWidth(4) = 1600
VSFlexGrid1.ColWidth(5) = 1500
DataCombo1.Text = ""



VSFlexGrid2.ColWidth(1) = 0
VSFlexGrid2.ColWidth(2) = 0
VSFlexGrid2.ColWidth(3) = 0
VSFlexGrid2.ColWidth(12) = 0
VSFlexGrid2.ColWidth(20) = 0
VSFlexGrid2.ColWidth(21) = 0
VSFlexGrid2.ColWidth(17) = 0
VSFlexGrid2.ColWidth(18) = 0

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT * FROM cltkdj"
Adodc8.Refresh

DataCombo4(14).Enabled = False
If Adodc8.Recordset.EOF Then
DataCombo4(14).Text = "00000001"
Else
uu = Val(Adodc8.Recordset.Fields(1)) + 1
DataCombo4(14).Text = Left("00000000", 8 - Len(Trim(Str(uu)))) + Trim(Str(uu))
End If


Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT * FROM cltk WHERE 单据='" & DataCombo4(14).Text & "' ORDER BY 序号 desc"
Adodc6.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "SELECT 供应单位,单据号,材料名称,材料规格,材料单位,退库数量,单价,合计金额,日期,序号,单据 FROM cltk WHERE 单据='" & DataCombo4(14).Text & "'ORDER BY 序号 desc "
Adodc7.Refresh

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

DataCombo4(13).Text = 1
DataCombo4(13).Text = Adodc7.Recordset.Fields(9) + 1

End Sub


Private Sub Label2_DblClick(Index As Integer)
Select Case Index
       Case 1
       DataCombo4(14).Enabled = True
End Select
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 1
       DataCombo4(14).Enabled = False
End Select
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
DataCombo4(11).Text = Adodc2.Recordset.Fields(0)
DataCombo4(3).Text = Adodc2.Recordset.Fields(2)
DataCombo4(4).Text = Adodc2.Recordset.Fields(3)
DataCombo4(5).Text = Adodc2.Recordset.Fields(4)
DataCombo4(8).Text = Adodc2.Recordset.Fields(5)
DataCombo4(10).Text = Adodc2.Recordset.Fields(7)
DataCombo4(9).Text = Adodc2.Recordset.Fields(6)
DataCombo4(13).Text = Adodc2.Recordset.Fields(13)
DataCombo4(16).Text = Adodc2.Recordset.Fields(8)
DataCombo4(18).Text = Adodc2.Recordset.Fields(0)
DataCombo4(19).Text = Adodc2.Recordset.Fields(1)
DataCombo1 = DataCombo4(3)
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
rs = VSFlexGrid2.Row
Adodc6.Recordset.MoveFirst
Adodc6.Recordset.Move rs - 1
DataCombo4(3).Text = Adodc6.Recordset.Fields(3)
DataCombo4(4).Text = Adodc6.Recordset.Fields(4)
DataCombo4(5).Text = Adodc6.Recordset.Fields(5)
DataCombo4(8).Text = Adodc6.Recordset.Fields(8)
DataCombo4(9).Text = Adodc6.Recordset.Fields(9)
DataCombo4(10).Text = Adodc6.Recordset.Fields(10)
DataCombo4(16).Text = Adodc6.Recordset.Fields(15)
DataCombo4(17).Text = Adodc6.Recordset.Fields(16)
DataCombo4(12).Text = Adodc6.Recordset.Fields(20)
DataCombo4(13).Text = Adodc6.Recordset.Fields(21)
DataCombo4(18).Text = Adodc6.Recordset.Fields(22)
DataCombo4(19).Text = Adodc6.Recordset.Fields(23)
DataCombo4(14).Text = Adodc6.Recordset.Fields(29)
DataCombo4(11).Text = Adodc6.Recordset.Fields(22) '
DataCombo1 = Adodc6.Recordset.Fields(3)
DataCombo3 = Adodc6.Recordset.Fields(15)

End Sub


Private Sub Timer1_Timer()
If BAR = 100 Then
Call clck(Adodc6, DataCombo4(14).Text)
Timer1.Enabled = False
ProgressBar1.Visible = False
Exit Sub
End If
BAR = BAR + 1
ProgressBar1.value = BAR
End Sub
