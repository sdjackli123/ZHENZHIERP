VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma132 
   BackColor       =   &H00C0E0FF&
   Caption         =   "毛坯退库"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   8520
      TabIndex        =   43
      Text            =   "Text4"
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5880
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   1320
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   6480
      Top             =   9360
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
      Left            =   6960
      Top             =   9240
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
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   615
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5880
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   615
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      Height          =   615
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   615
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
      Height          =   615
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新单据"
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   7320
      Top             =   9240
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
      Height          =   330
      Left            =   7920
      Top             =   9360
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo2"
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8640
      Top             =   9360
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
      Left            =   8640
      Top             =   9480
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
      Left            =   8640
      Top             =   9360
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
      Left            =   8640
      Top             =   9480
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma132.frx":0000
      Height          =   330
      Index           =   0
      Left            =   4800
      TabIndex        =   7
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   2
      Left            =   3840
      TabIndex        =   9
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma132.frx":0015
      Height          =   5775
      Left            =   360
      TabIndex        =   10
      Top             =   7080
      Width           =   19695
      _cx             =   34740
      _cy             =   10186
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
      FormatString    =   $"Forma132.frx":002A
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
      Bindings        =   "Forma132.frx":027B
      Height          =   3015
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Width           =   19575
      _cx             =   34528
      _cy             =   5318
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
      FormatString    =   $"Forma132.frx":0290
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   2
      Left            =   4920
      TabIndex        =   13
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   3
      Left            =   7320
      TabIndex        =   14
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   4
      Left            =   9720
      TabIndex        =   15
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   5
      Left            =   10920
      TabIndex        =   16
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   6
      Left            =   12120
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   7
      Left            =   13320
      TabIndex        =   18
      Top             =   2760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   8
      Left            =   15120
      TabIndex        =   19
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   16680
      TabIndex        =   36
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   327876609
      CurrentDate     =   39177
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma132.frx":04E1
      Height          =   330
      Index           =   3
      Left            =   4800
      TabIndex        =   39
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   6720
      Top             =   9000
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Forma132.frx":04F6
      Height          =   330
      Index           =   10
      Left            =   1320
      TabIndex        =   41
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "xm"
      Text            =   "DataCombo2"
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
      Height          =   375
      Index           =   4
      Left            =   7680
      TabIndex        =   44
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "退库负责"
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
      Left            =   360
      TabIndex        =   42
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "退回客户"
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
      Left            =   4800
      TabIndex        =   40
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "出库日期"
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
      Left            =   16680
      TabIndex        =   37
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   33
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯备注"
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
      Left            =   13320
      TabIndex        =   32
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯重量"
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
      Left            =   12120
      TabIndex        =   31
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯匹数"
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
      Left            =   10920
      TabIndex        =   30
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯幅宽"
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
      Left            =   9720
      TabIndex        =   29
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯布类"
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
      Left            =   7320
      TabIndex        =   28
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯简码"
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   27
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯库存"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   26
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
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
      TabIndex        =   25
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单据"
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
      Left            =   360
      TabIndex        =   24
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯入库序号"
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
      Left            =   2640
      TabIndex        =   23
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯入库单据"
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
      Left            =   360
      TabIndex        =   22
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯款号"
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
      Left            =   4920
      TabIndex        =   21
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "存放位置"
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
      Left            =   15120
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "Forma132"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BAR As Integer
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Private Sub Command1_Click()
On Error Resume Next
If MsgBox("确定删除吗？，锅号" + DataCombo1(1) + "" + 序号 + DataCombo1(2), vbYesNo) = vbNo Then Exit Sub
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.Delete
'Adodc2.RecordSource = "select * from v_mp_kc where 库存重量<>0 and 客户名称='" & DataCombo1(0) & "' and 简码 like '%'+'" & Text2 & "'+'%' order by 日期,客户名称,布类"
Adodc2.Refresh
Adodc3.RecordSource = "select * from mpbh where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc3.Refresh

DataCombo1(2).Text = 1
If Adodc3.Recordset.EOF Then
DataCombo1(2).Text = 1
Else
DataCombo1(2).Text = Adodc3.Recordset.RecordCount + 1
End If
DataCombo2(5) = ""
DataCombo2(6) = ""
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗？，锅号" + DataCombo1(1) + "" + 序号 + DataCombo1(2), vbYesNo) = vbNo Then Exit Sub
If Adodc3.Recordset.EOF Then Exit Sub
If DataCombo2(10) = "" Then
MsgBox ("请输入负责人")
Exit Sub
End If
If DataCombo2(5) = "" Or DataCombo2(6) = "" Then Exit Sub
For i = 0 To 8
Adodc3.Recordset.Fields(i) = DataCombo2(i)
Next
Adodc3.Recordset.Fields(9) = DataCombo1(1)
Adodc3.Recordset.Fields(10) = DataCombo1(2)
Adodc3.Recordset.Fields(11) = DataCombo1(3)
Adodc3.Recordset.Fields(12) = DTPicker1.value
Adodc3.Recordset.Fields(13) = DataCombo2(5)
Adodc3.Recordset.Fields(14) = DataCombo2(10)

Adodc3.Recordset.Update

Adodc2.RecordSource = "select * from v_mp_kc where 库存重量<0 and 单据号='" & DataCombo2(1) & "' and 序号='" & DataCombo2(0) & "'"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
MsgBox ("入库单据： " + DataCombo2(1) + "入库序号： " + Trim(DataCombo2(0)) + "出现负库存" + "请确认原因")
End If

Adodc3.RecordSource = "select * from mpbh where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc3.Refresh


DataCombo1(2).Text = 1
If Adodc3.Recordset.EOF Then
DataCombo1(2).Text = 1
Else
DataCombo1(2).Text = Adodc3.Recordset.RecordCount + 1
End If
DataCombo2(5) = ""
DataCombo2(6) = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
If DataCombo1(3) = "" Or DataCombo1(1) = "" Or DataCombo1(2) = "" Then
MsgBox ("输入不完整，不能保存")
Exit Sub
End If

'If MsgBox("确定结转吗？结转到的日期为" + Trim(DTPicker3.Value), vbYesNo) = vbNo Then Exit Sub
    Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "mpbhtk('" & DataCombo2(0) & "','" & DataCombo2(1) & "','" & DataCombo2(2) & "','" & DataCombo2(3) & "','" & DataCombo2(4) & "','" & DataCombo2(5) & "','" & DataCombo2(6) & "','" & DataCombo2(7) & "','" & DataCombo2(8) & "','" & DataCombo1(1) & "','" & DataCombo1(2) & "','" & DTPicker1.value & "','" & DataCombo2(5) & "','" & DataCombo2(10) & "','" & DataCombo1(3) & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

Adodc2.RecordSource = "select * from v_mp_kc where 库存重量<0 and 单据号='" & DataCombo2(1) & "' and 序号='" & DataCombo2(0) & "'"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
MsgBox ("入库单据： " + DataCombo2(1) + "入库序号： " + Trim(DataCombo2(0)) + "出现负库存" + "请确认原因")
End If

Adodc3.RecordSource = "select * from mpbh where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc3.Refresh

DataCombo1(2).Text = 1
If Adodc3.Recordset.EOF Then
DataCombo1(2).Text = 1
Else
DataCombo1(2).Text = Adodc3.Recordset.RecordCount + 1
End If

Adodc8.RecordSource = "select isnull(count(序号),0) from mpbh where 锅号='" & DataCombo1(1).Text & "'"
Adodc8.Refresh

If Adodc8.Recordset.Fields(0) > 4 Then
Call Command6_Click
Call Command5_Click
End If

End Sub


Private Sub Command5_Click()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT isnull(max(锅号),0) FROM mpbh where 锅号 like 'TK%' AND len(锅号)=8"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
DataCombo1(1).Text = "TK000001"
Else
uu = Val(Mid(Adodc5.Recordset.Fields(0), 3)) + 1
DataCombo1(1).Text = Left("TK000000", 8 - Len(Trim(uu))) + Trim(Str(uu))
End If

Adodc3.RecordSource = "SELECT * FROM mpbh WHERE 锅号='" & DataCombo1(1).Text & "'"
Adodc3.Refresh
DataCombo1(2).Text = 1
If Adodc6.Recordset.EOF Then
DataCombo1(2).Text = 1
Else
DataCombo1(2).Text = Adodc3.Recordset.RecordCount + 1
End If

End Sub

Private Sub Command6_Click()
Call mpck(Adodc7, DataCombo1(1))
End Sub

Private Sub DataCombo1_Change(Index As Integer)
Select Case Index
      Case 0
      DataCombo1(3).Text = DataCombo1(0).Text
       Case 1
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select *  from mpbh where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc3.Refresh

DataCombo1(2).Text = 1
If Adodc3.Recordset.EOF Then
DataCombo1(2).Text = 1
Else
DataCombo1(2).Text = Adodc3.Recordset.RecordCount + 1
End If

End Select
End Sub
Private Sub Form_Load()
On Error Resume Next
DTPicker1.value = Date
DTPicker3.value = Date
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT isnull(max(锅号),0) FROM mpbh where 锅号 like 'TK%' AND len(锅号)=8"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
DataCombo1(1).Text = "TK000001"
Else
uu = Val(Mid(Adodc5.Recordset.Fields(0), 3)) + 1
DataCombo1(1).Text = Left("TK000000", 8 - Len(Trim(uu))) + Trim(Str(uu))
End If

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT * FROM mpbh WHERE 锅号='" & DataCombo1(1).Text & "'"
Adodc6.Refresh
DataCombo1(2).Text = 1
If Adodc6.Recordset.EOF Then
DataCombo1(2).Text = 1
Else
DataCombo1(2).Text = Adodc6.Recordset.Fields(0) + 1
End If

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select xm  from fzr group by xm"
Adodc7.Refresh
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
For i = 0 To 10
DataCombo1(i).Text = ""
DataCombo2(i).Text = ""
Next
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 1700
VSFlexGrid1.ColWidth(5) = 1700
Text1.TabIndex = 0
Call Command5_Click
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

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 0
Text1 = ""
End Select
End Sub

Private Sub Text1_Change()
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'  group by 简称"
Adodc4.Refresh
End Sub

Private Sub Text2_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from v_mp_kc where 库存匹数<>0 and 客户名称='" & DataCombo1(0) & "' and 简码 like '%'+'" & Text2 & "'+'%' order by 日期,客户名称,布类"
Adodc2.Refresh
If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 400
Next
End If
End Sub

Private Sub VSFlexGrid1_dblClick()
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
For i = 0 To 3
If i = 1 Then i = 2
DataCombo1(i) = Adodc1.Recordset.Fields(i)
Next
DataCombo2(5) = Adodc1.Recordset.Fields(15)
DataCombo2(6) = Adodc1.Recordset.Fields(16)
End Sub

Private Sub Text3_Change()
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text3 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc9.Refresh
End Sub

Private Sub Text4_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from v_mp_kc where  库存匹数<>0 and 客户名称='" & DataCombo1(0) & "' and 和约号 like '%'+'" & Text4 & "'+'%' order by 日期,单据号,序号"
Adodc2.Refresh
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc2.Recordset.Move rs - 1
For i = 0 To 1
DataCombo2(i) = Adodc2.Recordset.Fields(i)
Next
DataCombo2(2) = Adodc2.Recordset.Fields(4)
DataCombo2(3) = Adodc2.Recordset.Fields(5)
DataCombo2(4) = Adodc2.Recordset.Fields(6)
DataCombo2(5) = Adodc2.Recordset.Fields(14)
DataCombo2(6) = Adodc2.Recordset.Fields(15)
End Sub

Private Sub VSFlexGrid3_dblClick()
On Error Resume Next
If Adodc3.Recordset.EOF Then Exit Sub
rs = VSFlexGrid3.Row
Adodc3.Recordset.MoveFirst
Adodc3.Recordset.Move rs - 1
For i = 0 To 8
DataCombo2(i) = Adodc3.Recordset.Fields(i)
Next
DataCombo1(1) = Adodc3.Recordset.Fields(9)
DataCombo1(2) = Adodc3.Recordset.Fields(10)
DataCombo1(3) = Adodc3.Recordset.Fields(11)
End Sub



