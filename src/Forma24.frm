VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma24 
   BackColor       =   &H00C0E0FF&
   Caption         =   "条码打印"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14490
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   14490
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "委外打印"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   12
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查漏"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   600
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   3360
      Top             =   8520
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
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   11
      Left            =   4920
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   10
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印2"
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   5160
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   3720
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   3960
      Top             =   8520
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   2640
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   11400
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   4920
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   5160
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   5160
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1800
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6480
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
      Left            =   6480
      Top             =   9840
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
      Left            =   6480
      Top             =   9960
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
      Height          =   330
      Left            =   6240
      Top             =   10080
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma24.frx":0000
      Height          =   3855
      Left            =   960
      TabIndex        =   5
      Top             =   4320
      Width           =   12615
      _cx             =   22251
      _cy             =   6800
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
      GridLines       =   2
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "厂内打印"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma24.frx":0015
      Height          =   3015
      Left            =   8040
      TabIndex        =   12
      Top             =   840
      Width           =   5535
      _cx             =   9763
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   2
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   330563585
      CurrentDate     =   39961
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma24.frx":002A
      Height          =   330
      Index           =   20
      Left            =   4920
      TabIndex        =   33
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "车台"
      Text            =   "DataCombo1"
   End
   Begin VB.Label lblLabels 
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
      Height          =   405
      Index           =   9
      Left            =   1080
      TabIndex        =   36
      Top             =   1200
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
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
      Height          =   405
      Index           =   8
      Left            =   4200
      TabIndex        =   32
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "车间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   4200
      TabIndex        =   29
      Top             =   1080
      Width           =   705
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   4200
      X2              =   4680
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "条码范围:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   1440
      TabIndex        =   23
      Top             =   3720
      Width           =   1185
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "换批"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   4200
      TabIndex        =   22
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "打印日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   1080
      TabIndex        =   19
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "匹号范围:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   13
      Left            =   1080
      TabIndex        =   11
      Top             =   3120
      Width           =   1545
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4200
      X2              =   4680
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "织号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "计量标重"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   600
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "计划量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   1080
      TabIndex        =   8
      Top             =   2520
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "理论匹数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   3960
      TabIndex        =   7
      Top             =   2520
      Width           =   1185
   End
End
Attribute VB_Name = "Forma24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(1).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Or Text1(8).Text = "" Or Text1(9).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If

If Val(Text1(8)) < 0 Then
MsgBox ("请输入条码范围")
Exit Sub
End If

If (Val(Text1(9)) - Val(Text1(8))) <> (Val(Text1(5)) - Val(Text1(4))) Then
MsgBox ("匹号数与条码数不符")
Exit Sub
End If

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 当前编号 from scbqbh where cast('" & Text1(8) & "' as int) > 当前编号"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
MsgBox ("已有此编号！")
Exit Sub
End If


TM = Val(Text1(4))
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select * from jhscbq where '" & TM & "' between 起号 and 结号 and 织号='" & Text1(0) & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
MsgBox ("已有此匹号保存记录！")
Exit Sub
End If

TM = Val(Text1(4))
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select * from TMFJ where 匹号='" & TM & "' and 织号='" & Text1(0) & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
MsgBox ("已有此匹号！")
Exit Sub
End If


Adodc1.Recordset.AddNew
For i = 0 To 12
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Fields(7) = DTPicker1.value
Adodc1.Recordset.Update
Adodc1.Refresh

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "bqtmcz('" & Text1(0) & "','" & Text1(1) & "','" & Text1(2) & "','" & Text1(3) & "','" & Text1(4) & "','" & Text1(5) & "','" & Text1(6) & "','" & DTPicker1.value & "','" & Text1(8) & "','" & Text1(9) & "','" & yhm & "','" & Now & "','录入','" & Text1(10) & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjtmsc2('" & Text1(4) & "','" & Text1(5) & "','" & Text1(8) & "','" & Text1(0) & "','" & Text1(6) & "','" & Text1(11) & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel

Adodc2.RecordSource = "SELECT * FROM scbqzj WHERE 织号='" & Text1(0).Text & "'"
Adodc2.Refresh

End Sub

Private Sub Command10_Click()
On Error Resume Next
If MsgBox("确定打印吗？", vbYesNo) = vbNo Then Exit Sub
Dim L As Long
If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(1).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If
If Val(Text1(5).Text) >= Val(Text1(4).Text) Then
L = Val(Text1(8).Text)
For i = Text1(4).Text To Text1(5).Text
Call dbqww(Adodc3, Adodc4, Text1(0), Val(i), L, Text1(10), Text1(11), Text1(6))
L = L + 1
Next
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定修改吗", vbYesNo) = vbNo Then Exit Sub
If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(1).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If
For i = 0 To 12
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Fields(7) = DTPicker1.value
Adodc1.Recordset.Update
Adodc1.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "bqtmcz('" & Text1(0) & "','" & Text1(1) & "','" & Text1(2) & "','" & Text1(3) & "','" & Text1(4) & "','" & Text1(5) & "','" & Text1(6) & "','" & DTPicker1.value & "','" & Text1(8) & "','" & Text1(9) & "','" & yhm & "','" & Now & "','修改','" & Text1(10) & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.CancelAdodc2.RecordSource = "SELECT * FROM scbqzj WHERE 织号='" & Text1(0).Text & "'"
Adodc2.Refresh

End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除吗", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "bqtmcz('" & Text1(0) & "','" & Text1(1) & "','" & Text1(2) & "','" & Text1(3) & "','" & Text1(4) & "','" & Text1(5) & "','" & Text1(6) & "','" & DTPicker1.value & "','" & Text1(8) & "','" & Text1(9) & "','" & yhm & "','" & Now & "','删除')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.CancelAdodc2.RecordSource = "SELECT * FROM scbqzj WHERE 织号='" & Text1(0).Text & "'"
Adodc2.Refresh
End Sub

Private Sub Command5_Click()
On Error Resume Next
If MsgBox("确定打印吗？", vbYesNo) = vbNo Then Exit Sub
Dim L As Long
If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(1).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If
If Val(Text1(5).Text) >= Val(Text1(4).Text) Then
L = Val(Text1(8).Text)
For i = Text1(4).Text To Text1(5).Text
'Call dbq(Adodc3, Adodc4, Text1(0), Val(i), L, Text1(10), Text1(11), Text1(6))
L = L + 1
Next
End If
End Sub

Private Sub Command6_Click()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM jhscbq WHERE 织号='" & Text1(0).Text & "' order by 日期"
Adodc1.Refresh
       
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT * FROM scbqzj WHERE 织号='" & Text1(0).Text & "'"
Adodc2.Refresh
End Sub

Private Sub Command7_Click()
On Error Resume Next
If Text1(4).Text = "" Or Text1(5).Text = "" Then
MsgBox ("请输入打印匹号")
Text1(7).Text = ""
Exit Sub
End If

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select 当前编号 from scbqbh"
Adodc5.Refresh

If Adodc5.Recordset.Fields(0) = Null Then
Text1(8).Text = 1

Text1(9).Text = Val(Text1(8).Text) + Val(Text1(5).Text) - Val(Text1(4).Text)
Else
Text1(8).Text = Val(Adodc5.Recordset.Fields(0)) + 1
Text1(9).Text = Val(Text1(8).Text) + Val(Text1(5).Text) - Val(Text1(4).Text)
End If
End Sub

Private Sub Command8_Click()
On Error Resume Next
If MsgBox("确定打印吗？", vbYesNo) = vbNo Then Exit Sub
Dim L As Long
If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(1).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If
If Val(Text1(5).Text) >= Val(Text1(4).Text) Then
L = Val(Text1(8).Text)
For i = Text1(4).Text To Text1(5).Text
'Call dbq1(Adodc3, Adodc4, Text1(0).Text, Val(i), L)
L = L + 1
Next
End If

End Sub

Private Sub Command9_Click()
'FormA104.Text3 = Text1(0)  原是批次
'FormA104.Show
Forma100.Show
End Sub

Private Sub DataCombo1_Change(Index As Integer)
Select Case Index
       Case 20
Text1(10) = DataCombo1(20)
End Select
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
Select Case Index
       Case 20
Text1(10) = DataCombo1(20)
End Select
End Sub

Private Sub Form_Load()
For i = 0 To 12
Text1(i).Text = ""
Next
DTPicker1.value = Date
DataCombo1(20) = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM jhscbq WHERE 织号='" & Text1(0).Text & "' order by 日期"
Adodc1.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT 车台 FROM ZBCT  GROUP BY 车台"
Adodc6.Refresh
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1600
VSFlexGrid1.ColWidth(2) = 1600
VSFlexGrid1.ColWidth(3) = 1000
VSFlexGrid1.ColWidth(4) = 1000
VSFlexGrid1.ColWidth(5) = 1600
VSFlexGrid1.ColWidth(6) = 1600
End Sub


Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM jhscbq WHERE 织号='" & Text1(0).Text & "' order by 日期"
Adodc1.Refresh
       
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT * FROM scbqzj WHERE 织号='" & Text1(0).Text & "'"
Adodc2.Refresh
       
       
       Case 1, 2
If Val(Text1(2).Text) > 0 Then
If Int(Val(Text1(1).Text) / Val(Text1(2).Text)) = Val(Text1(1).Text) / Val(Text1(2).Text) Then
Text1(3).Text = Val(Text1(1).Text) / Val(Text1(2).Text)
Else
Text1(3).Text = Int(Val(Text1(1).Text) / Val(Text1(2).Text)) + 1
End If
End If
End Select

End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To 12
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
DTPicker1.value = Adodc1.Recordset.Fields(7)
End Sub
