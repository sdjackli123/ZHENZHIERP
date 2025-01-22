VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formc19 
   BackColor       =   &H00C0E0FF&
   Caption         =   "委外入库"
   ClientHeight    =   9570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   15465
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新单据号"
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   10560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Formc19.frx":0000
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   14400
      TabIndex        =   0
      Text            =   "Text3"
      Top             =   2880
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formc19.frx":0006
      Height          =   330
      Left            =   12480
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "委外单位"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   1
      Left            =   3000
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   2
      Left            =   5280
      TabIndex        =   13
      Top             =   1920
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   3
      Left            =   8760
      TabIndex        =   14
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   4
      Left            =   10560
      TabIndex        =   15
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   5
      Left            =   12480
      TabIndex        =   16
      Top             =   1920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   6
      Left            =   720
      TabIndex        =   17
      Top             =   3000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   7
      Left            =   3000
      TabIndex        =   18
      Top             =   3000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5280
      TabIndex        =   19
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   329187329
      CurrentDate     =   39921
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc19.frx":001B
      Height          =   330
      Index           =   8
      Left            =   7200
      TabIndex        =   20
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "负责"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Formc19.frx":0030
      Height          =   1815
      Left            =   720
      TabIndex        =   21
      Top             =   4320
      Width           =   14655
      _cx             =   25850
      _cy             =   3201
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
      Bindings        =   "Formc19.frx":0045
      Height          =   1935
      Left            =   720
      TabIndex        =   22
      Top             =   6360
      Width           =   14655
      _cx             =   25850
      _cy             =   3413
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   1920
      Top             =   9360
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
      Height          =   495
      Left            =   2760
      Top             =   9360
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Height          =   375
      Left            =   2040
      Top             =   9480
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
      Left            =   2400
      Top             =   9600
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
      Height          =   495
      Left            =   3000
      Top             =   9240
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
      Left            =   2280
      Top             =   9480
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
      Left            =   3600
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
      Height          =   375
      Left            =   2880
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
      Left            =   3480
      Top             =   9360
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
      Left            =   3840
      Top             =   9360
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc19.frx":005A
      Height          =   330
      Index           =   9
      Left            =   8760
      TabIndex        =   23
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "类别"
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "单据号"
      Height          =   495
      Left            =   720
      TabIndex        =   38
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Enabled         =   0   'False
      Height          =   375
      Index           =   11
      Left            =   14400
      TabIndex        =   37
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "日期"
      Height          =   495
      Index           =   10
      Left            =   5280
      TabIndex        =   36
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   255
      Index           =   8
      Left            =   10560
      TabIndex        =   35
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "重量"
      Height          =   495
      Index           =   7
      Left            =   3000
      TabIndex        =   34
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "匹数"
      Height          =   495
      Index           =   6
      Left            =   720
      TabIndex        =   33
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   495
      Index           =   5
      Left            =   12480
      TabIndex        =   32
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "品名"
      Height          =   495
      Index           =   4
      Left            =   5280
      TabIndex        =   31
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "克重"
      Height          =   495
      Index           =   0
      Left            =   8760
      TabIndex        =   30
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "门幅"
      Height          =   495
      Index           =   3
      Left            =   10560
      TabIndex        =   29
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "锅号"
      Height          =   495
      Index           =   2
      Left            =   3000
      TabIndex        =   28
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户"
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   27
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "负责"
      Height          =   495
      Index           =   9
      Left            =   7200
      TabIndex        =   26
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "委外类别"
      Height          =   495
      Index           =   13
      Left            =   8760
      TabIndex        =   25
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "委外单位"
      Height          =   495
      Index           =   14
      Left            =   12480
      TabIndex        =   24
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "Formc19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text2 = "" Then
MsgBox ("输入有误！")
Exit Sub
End If

Adodc10.RecordSource = "select distinct 单据 from wwkpdr where 锅号='" & DataCombo1(1) & "' and 品名='" & DataCombo1(2) & "'"
Adodc10.Refresh
If Not Adodc10.Recordset.EOF Then
If MsgBox("已经入库单据是  " + Adodc10.Recordset.Fields(0) + "确定重复入库吗？", vbYesNo) = vbNo Then Exit Sub
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "wwkpdrlr('" & DataCombo1(0).Text & "','" & DataCombo1(1).Text & "','" & DataCombo1(2).Text & "','" & DataCombo1(3).Text & "','" & DataCombo1(4).Text & "','" & DataCombo1(5).Text & "','" & DataCombo1(6).Text & "','" & DataCombo1(7).Text & "','" & DataCombo1(8) & "','" & DataCombo1(9) & "','" & DTPicker1.value & "','" & Text1 & "','" & DataCombo2.Text & "','" & Text2 & "','" & Text3 & "')"        ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel
Adodc1.RecordSource = "SELECT * FROM wwkpdr where 单据='" & Text2 & "' order by 序号"
Adodc1.Refresh

Adodc5.RecordSource = "SELECT 序号 FROM wwkpdr WHERE 单据='" & Text2 & "' ORDER BY 序号 DESC"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
Text3 = 1
Else
Text3 = Adodc5.Recordset.Fields(0) + 1
End If

DataCombo1(6).Text = ""
DataCombo1(7).Text = ""
End Sub


Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
For i = 0 To 9
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc1.Recordset.Fields(10) = DTPicker1.value
Adodc1.Recordset.Fields(11) = Text1.Text
Adodc1.Recordset.Fields(12) = DataCombo2.Text
Adodc1.Recordset.Fields(14) = Text3
Adodc1.Recordset.Update
Adodc1.RecordSource = "SELECT * FROM wwkpdr where 单据='" & Text2 & "' order by 序号"
Adodc1.Refresh

Adodc5.RecordSource = "SELECT 序号 FROM wwkpdr WHERE 单据='" & Text2 & "' ORDER BY 序号 DESC"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
Text3 = 1
Else
Text3 = Adodc5.Recordset.Fields(0) + 1
End If

DataCombo1(6).Text = ""
DataCombo1(7).Text = ""

End Sub

Private Sub Command3_Click()
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT * FROM wwkpdr where 日期=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh

If Adodc3.Recordset.EOF Then
Text2 = Format(CDate(DTPicker1.value), "yymmdd") + "001"
Else
Adodc3.RecordSource = "SELECT max(right(单据,len(单据)-6)) FROM wwkpdr where 日期=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh
L = Val(Adodc3.Recordset.Fields(0))
If Len(L + 1) = 1 Then
Text2 = Format(CDate(DTPicker1.value), "yymmdd") + "00" + Trim(L + 1)
End If
If Len(L + 1) = 2 Then
Text2 = Format(CDate(DTPicker1.value), "yymmdd") + "0" + Trim(L + 1)
End If
If Len(L + 1) = 3 Then
Text2 = Format(CDate(DTPicker1.value), "yymmdd") + Trim(L + 1)
End If
End If


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM wwkpdr WHERE 单据='" & Text2 & "' ORDER BY 序号"
Adodc1.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT 序号 FROM wwkpdr WHERE 单据='" & Text2 & "' ORDER BY 序号 DESC"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
Text3 = 1
Else
Text3 = Adodc5.Recordset.Fields(0) + 1
End If

End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.RecordSource = "SELECT * FROM wwkpdr where 单据='" & Text2 & "' order by 序号"
Adodc1.Refresh

Adodc5.RecordSource = "SELECT 序号 FROM wwkpdr WHERE 单据='" & Text2 & "' ORDER BY 序号 DESC"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
Text3 = 1
Else
Text3 = Adodc5.Recordset.Fields(0) + 1
End If

DataCombo1(6).Text = ""
DataCombo1(7).Text = ""

End Sub

Private Sub Command5_Click()
'Call DXDY(Adodc4, Text2)
End Sub

Private Sub Command6_Click()
Unload Me
End Sub


Private Sub Command7_Click()
wwdm = 2
Formc21.DataCombo1 = DataCombo1(0)
Formc21.Show
End Sub

Private Sub Command9_Click()
For i = 0 To 9
DataCombo1(i).Text = ""
Next
DTPicker1.value = Date
Text1.Text = ""
Text3.Text = ""
DataCombo2 = ""

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM wwkpdr where 单据='" & Text2 & "' order by 序号"
Adodc1.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT 序号 FROM wwkpdr WHERE 单据='" & Text2 & "' ORDER BY 序号 DESC"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
Text3 = 1
Else
Text3 = Adodc5.Recordset.Fields(0) + 1
End If

Adodc6.Refresh
Adodc7.Refresh
Adodc8.Refresh
End Sub

Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 1
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT  * FROM wwkpd where 锅号= '" & DataCombo1(1).Text & "'"
Adodc2.Refresh
       Case 9
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM wwkpdr where 单据='" & Text2 & "' order by 序号"
Adodc1.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT 序号 FROM wwkpdr WHERE 单据='" & Text2 & "' ORDER BY 序号 DESC"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
Text3 = 1
Else
Text3 = Adodc5.Recordset.Fields(0) + 1
End If
End Select
End Sub


Private Sub Form_Load()
On Error Resume Next
For i = 0 To 9
DataCombo1(i).Text = ""
Next
DTPicker1.value = Date
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
DataCombo2 = ""
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT * FROM wwkpdr where 日期=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh

If Adodc3.Recordset.EOF Then
Text2 = Format(CDate(DTPicker1.value), "yymmdd") + "001"
Else
Adodc3.RecordSource = "SELECT max(right(单据,len(单据)-6)) FROM wwkpdr where 日期=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh
L = Val(Adodc3.Recordset.Fields(0))
If Len(L + 1) = 1 Then
Text2 = Format(CDate(DTPicker1.value), "yymmdd") + "00" + Trim(L + 1)
End If
If Len(L + 1) = 2 Then
Text2 = Format(CDate(DTPicker1.value), "yymmdd") + "0" + Trim(L + 1)
End If
If Len(L + 1) = 3 Then
Text2 = Format(CDate(DTPicker1.value), "yymmdd") + Trim(L + 1)
End If
End If


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM wwkpdr WHERE 单据='" & Text2 & "' ORDER BY 序号"
Adodc1.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT 序号 FROM wwkpdr WHERE 单据='" & Text2 & "' ORDER BY 序号 DESC"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
Text3 = 1
Else
Text3 = Adodc5.Recordset.Fields(0) + 1
End If
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT distinct 负责 FROM wwkpd"
Adodc6.Refresh
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "SELECT distinct 委外单位 FROM wwkpd"
Adodc7.Refresh
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT distinct 类别 FROM wwkpd"
Adodc8.Refresh
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
End Sub

Private Sub Label2_Click()
DataCombo1(9).Enabled = False
End Sub

Private Sub Label2_DblClick()
DataCombo1(9).Enabled = True
End Sub

Private Sub VSFlexGrid1_dblClick()
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To 9
DataCombo1(i).Text = Adodc1.Recordset.Fields(i)
Next
DTPicker1.value = Adodc1.Recordset.Fields(10)
Text1.Text = Adodc1.Recordset.Fields(11)
DataCombo2.Text = Adodc1.Recordset.Fields(12)
Text3 = Adodc1.Recordset.Fields(14)

End Sub

Private Sub VSFlexGrid3_dblClick()
If Adodc2.Recordset.EOF Then Exit Sub
rs = VSFlexGrid3.Row
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
For i = 0 To 9
DataCombo1(i).Text = Adodc2.Recordset.Fields(i)
Next
DTPicker1.value = Adodc2.Recordset.Fields(10)
Text1.Text = Adodc2.Recordset.Fields(11)
DataCombo2.Text = Adodc2.Recordset.Fields(12)
End Sub


