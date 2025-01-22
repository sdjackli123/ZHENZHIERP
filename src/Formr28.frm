VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formr28 
   BackColor       =   &H00C0E0FF&
   Caption         =   "手工统计"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   LinkTopic       =   "Form10"
   ScaleHeight     =   10185
   ScaleWidth      =   12060
   StartUpPosition =   2  '屏幕中心
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
      Left            =   1080
      TabIndex        =   29
      Text            =   "Combo3"
      Top             =   120
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4920
      Top             =   9720
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
      ItemData        =   "Formr28.frx":0000
      Left            =   1080
      List            =   "Formr28.frx":0019
      TabIndex        =   27
      Text            =   "Combo2"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单位刷新"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3600
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5400
      Top             =   9720
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
      Left            =   5760
      Top             =   9600
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
      Left            =   6120
      Top             =   9720
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
      Bindings        =   "Formr28.frx":0057
      Height          =   4215
      Left            =   360
      TabIndex        =   22
      Top             =   5160
      Width           =   11295
      _cx             =   19923
      _cy             =   7435
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
      Bindings        =   "Formr28.frx":006C
      Height          =   2775
      Left            =   5040
      TabIndex        =   21
      Top             =   600
      Width           =   6615
      _cx             =   11668
      _cy             =   4895
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
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
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3000
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   249
      Format          =   423034881
      CurrentDate     =   36892
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Formr28.frx":0081
      Left            =   1080
      List            =   "Formr28.frx":008B
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2775
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Formr28.frx":0096
      Height          =   360
      Left            =   1080
      TabIndex        =   26
      Top             =   4560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "车间编号"
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
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "编号"
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
      Left            =   360
      TabIndex        =   28
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "信息"
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
      Left            =   360
      TabIndex        =   25
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   24
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "简码"
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
      Left            =   4080
      TabIndex        =   13
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
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
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "库类"
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
      TabIndex        =   8
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单位"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "名称"
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
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "数量"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "Formr28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim rr As Integer
Public rs As Single: Public rlbl As Integer
Dim cdbhf As Integer
Private Sub Combo1_Click()
Text1(1).Text = Combo1.Text
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Text1(2).Text = "" Then
MsgBox ("请输入数量")
Exit Sub
End If

Adodc1.Recordset.AddNew
For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Fields(6) = Combo2.Text
Adodc1.Recordset.Fields(7) = DataCombo4
Adodc1.Recordset.Fields(8) = Combo3
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc3.RecordSource = "SELECT max(序号) FROM SGTJ where 日期='" & DTPicker2.value & "'"
Adodc3.Refresh
Text1(5).Text = 1
If Adodc3.Recordset.EOF Then
Text1(5).Text = 1
Else
Text1(5).Text = Adodc3.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Adodc3.RecordSource = "SELECT max(序号) FROM SGTJ where 日期='" & DTPicker2.value & "'"
Adodc3.Refresh
Text1(5).Text = 1
If Adodc3.Recordset.EOF Then
Text1(5).Text = 1
Else
Text1(5).Text = Adodc3.Recordset.Fields(0) + 1
End If
Text1(2).Text = 0
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
On Error Resume Next
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Adodc3.Refresh
Adodc1.RecordSource = "SELECT * FROM SGTJ where 日期='" & DTPicker2.value & "' ORDER BY 序号 DESC"
Adodc1.Refresh
Text1(5).Text = 1
If Adodc3.Recordset.EOF Then
Text1(5).Text = 1
Else
Text1(5).Text = Adodc3.Recordset.Fields(0) + 1
End If
Text1(2).Text = 0
End Sub

Private Sub Command6_Click()
If MsgBox("确定单位刷新吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "update sgtj set 数量=round(数量/1000,5),单位='kg' where 单位='g' and 日期='" & Text1(4).Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("单位刷新成功！")
End Sub

Private Sub DTPicker2_Change()
Text1(4).Text = DTPicker2.value
End Sub

Private Sub DTPicker2_CloseUp()
Text1(4).Text = DTPicker2.value
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

Private Sub VSFlexGrid1_Click()
On Error Resume Next
rlbl = 2
rs = VSFlexGrid1.Row
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
Text1(0).Text = Adodc2.Recordset.Fields(0)
Text1(3).Text = Adodc2.Recordset.Fields(1)
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text1(2).Text = "" Then
MsgBox ("请输入数量")
Exit Sub
End If
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub

For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Fields(6) = Combo2.Text
Adodc1.Recordset.Fields(7) = DataCombo4
Adodc1.Recordset.Fields(8) = Combo3
Adodc1.Recordset.Update
Adodc1.Refresh

Adodc3.RecordSource = "SELECT max(序号) FROM SGTJ where 日期='" & DTPicker2.value & "'"
Adodc3.Refresh
Text1(5).Text = 1
If Adodc3.Recordset.EOF Then
Text1(5).Text = 1
Else
Text1(5).Text = Adodc3.Recordset.Fields(0) + 1
End If
Text1(2).Text = 0
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Form_Load()

On Error Resume Next
Text2.Text = ""
Combo1.Text = "kg"

DTPicker2.value = Date
DataCombo4 = ""
Combo2 = ""
Combo3 = ""
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM SGTJ where 日期='" & DTPicker2.value & "' ORDER BY 序号 DESC"
Adodc1.Refresh

For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = ""
Next
Text1(4).Text = Date
Text1(1).Text = Combo1.Text

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 染料名称,染化助库名 from rhzh where 简码 like '%'+'" & Text2.Text & "'+'%'"
Adodc2.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 车间编号 from cj  GROUP BY 车间编号"
Adodc4.Refresh


Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT max(序号) FROM SGTJ where 日期='" & DTPicker2.value & "'"
Adodc3.Refresh
Text1(5).Text = 1
If Adodc3.Recordset.EOF Then
Text1(5).Text = 1
Else
Text1(5).Text = Adodc3.Recordset.Fields(0) + 1
End If

VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(1) = 1800
VSFlexGrid1.ColWidth(2) = 1800
VSFlexGrid2.ColWidth(1) = 1500
VSFlexGrid2.ColWidth(2) = 1500
VSFlexGrid2.ColWidth(3) = 1500
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False

End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
rlbl = 1
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid2.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To 5
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
Combo2.Text = Adodc1.Recordset.Fields(6)
DataCombo4 = Adodc1.Recordset.Fields(7)
Combo3 = Adodc1.Recordset.Fields(8)

DTPicker2.value = Text1(4).Text
Combo1.Text = Text1(1).Text
Command1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 3
If rlbl = 2 Then
If InStr(Text1(3).Text, "分散") > 0 Or InStr(Text1(3).Text, "活性") > 0 Then
Combo1.Text = "g"
Text1(1).Text = "g"
End If
If InStr(Text1(3).Text, "分散") = 0 And InStr(Text1(3).Text, "活性") = 0 Then
Combo1.Text = "kg"
Text1(1).Text = "kg"
End If
End If
End Select
End Sub

Private Sub Text2_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 染料名称,染化助库名 from rhzh where 简码 like '%'+'" & Text2.Text & "'+'%'"
Adodc2.Refresh
End Sub
