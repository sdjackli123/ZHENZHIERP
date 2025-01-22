VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formr334 
   BackColor       =   &H00C0E0FF&
   Caption         =   "配料单"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form34"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0FFFF&
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
      ItemData        =   "Formr334.frx":0000
      Left            =   8280
      List            =   "Formr334.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Formr334.frx":0018
      Left            =   10080
      List            =   "Formr334.frx":0022
      TabIndex        =   22
      Text            =   "Combo2"
      Top             =   840
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   5160
      Top             =   9840
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
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
      Left            =   5880
      Top             =   9720
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
      Left            =   5520
      Top             =   9720
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
      Left            =   4680
      Top             =   9600
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
      Left            =   5280
      Top             =   9840
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Height          =   375
      Left            =   5640
      Top             =   9480
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5760
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
      Bindings        =   "Formr334.frx":0036
      Height          =   2535
      Left            =   960
      TabIndex        =   19
      Top             =   1560
      Width           =   14295
      _cx             =   25215
      _cy             =   4471
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   390
      Left            =   5520
      TabIndex        =   18
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成确定"
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "批量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   5640
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
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
      ItemData        =   "Formr334.frx":004B
      Left            =   5520
      List            =   "Formr334.frx":0055
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330235905
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   330235905
      CurrentDate     =   36892
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   13200
      TabIndex        =   16
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   330235905
      CurrentDate     =   36892
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formr334.frx":0063
      Height          =   4695
      Left            =   360
      TabIndex        =   20
      Top             =   4680
      Width           =   14895
      _cx             =   26273
      _cy             =   8281
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
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   23
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "库类"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   9360
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "确认日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   12000
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "确认意见"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "料单信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "料单编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Formr334"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public c, r As Integer:  Dim ZS(10) As String: Dim sz(61) As String: Dim pfsz(6) As String: Dim pfdsz(6) As String
Dim cdbhf As Integer

Private Sub Command1_Click()
If DataCombo1.Text = "" Then Exit Sub
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "DBPLDSH('" & DataCombo1.Text & "','" & Combo1.Text & "','" & DTPicker3.value & "','" & Combo2 & "')"   ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel
MsgBox ("审核完成！")
Adodc2.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
FormR336.DTPicker1.value = DTPicker1.value
FormR336.DTPicker2.value = DTPicker2.value
FormR336.Show
End Sub

Private Sub Command4_Click()
If DataCombo1.Text = "" Then
Adodc2.RecordSource = "SELECT 客户,锅号,品名,颜色,色号,数量,操作,车台,日期,信息,编号,审核 FROM pLd WHERE  cast(CONVERT(varchar(120),日期, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND 审核='" & Combo3 & "' ORDER BY 日期,编号"
Adodc2.Refresh
Else
Adodc2.RecordSource = "SELECT 客户,锅号,品名,颜色,色号,数量,操作,车台,日期,信息,编号,审核 FROM pLd WHERE 编号='" & DataCombo1.Text & "' ORDER BY 日期,编号"
Adodc2.Refresh
End If

Adodc6.RecordSource = "select * from pldc where  料单编号='" & DataCombo1.Text & "'"
Adodc6.Refresh

'If Not Adodc6.Recordset.EOF Then
'If MsgBox("已存在此料单，是否重新生成？", vbYesNo) = vbNo Then
'Adodc1.RecordSource = "select 工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,次序号 from pldc where  料单编号='" & DataCombo1.Text & "'"
'Adodc1.Refresh
'Exit Sub
'End If
'End If

sql1 = "delete  from pldc WHERE 料单编号='" & DataCombo1.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

Adodc4.RecordSource = "select * from pld where  编号='" & DataCombo1.Text & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.Recordset.MoveFirst

For i = 0 To 10
ZS(i) = Adodc4.Recordset.Fields(i)
Next

mb = 0
For i = 12 To 61
If Adodc4.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next

ProgressBar1.Visible = True
For i = 12 To mb + 12
If Adodc4.Recordset.Fields(i) <> "" Then
sz(0) = Mid(Adodc4.Recordset.Fields(i), 1, InStr(Adodc4.Recordset.Fields(i), "(") - 1)
sz(1) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "(") + 1, InStr(Adodc4.Recordset.Fields(i), ")") - InStr(Adodc4.Recordset.Fields(i), "(") - 1)
sz(2) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), ")") + 1, InStr(Adodc4.Recordset.Fields(i), "-") - InStr(Adodc4.Recordset.Fields(i), ")") - 1)
sz(3) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "-") + 1, InStr(Adodc4.Recordset.Fields(i), "\") - InStr(Adodc4.Recordset.Fields(i), "-") - 1)
sz(4) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "\") + 1, InStr(Adodc4.Recordset.Fields(i), "#") - InStr(Adodc4.Recordset.Fields(i), "\") - 1)
sz(5) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "#") + 1, InStr(Adodc4.Recordset.Fields(i), "^") - InStr(Adodc4.Recordset.Fields(i), "#") - 1)
sz(6) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "^") + 1, InStr(Adodc4.Recordset.Fields(i), "[") - InStr(Adodc4.Recordset.Fields(i), "^") - 1)
sz(7) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "[") + 1, InStr(Adodc4.Recordset.Fields(i), "]") - InStr(Adodc4.Recordset.Fields(i), "[") - 1)
sz(8) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "]") + 1, InStr(Adodc4.Recordset.Fields(i), "{") - InStr(Adodc4.Recordset.Fields(i), "]") - 1)
sz(9) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "{") + 1)

L = i - 11
sql1 = "insert into pldc(审核,锅号,压力,生产类别,配方单,重量,配料打印员,染化助单价,配料日期,生产信息,料单编号,工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,车速,次序号) VALUES('" & ZS(0) & "','" & ZS(1) & "','" & ZS(2) & "','" & ZS(3) & "','" & ZS(4) & "','" & ZS(5) & "','" & ZS(6) & "','" & ZS(7) & "','" & ZS(8) & "','" & ZS(9) & "','" & ZS(10) & "','" & sz(0) & "','" & sz(1) & "','" & sz(2) & "',  " & _
                                                                        "'" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & L & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
ProgressBar1.value = 100 / mb * L
End If
Next
ProgressBar1.Visible = False
End If

Adodc1.RecordSource = "select 工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,次序号 from pldc where  料单编号='" & DataCombo1.Text & "'"
Adodc1.Refresh

End Sub

Private Sub Command5_Click()
If DataCombo1.Text = "" Then
MsgBox ("请输入料单编号")
Exit Sub
End If
Call scpfd(DataCombo1.Text)
MsgBox ("生成成功！")
End Sub

Private Sub DataCombo1_Change()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,次序号 FROM pldc WHERE 料单编号='" & DataCombo1.Text & "' ORDER BY 工序名称,次序号"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
L = ""
Else
L = Adodc1.Recordset.Fields(0)
End If

End Sub

Private Sub DataCombo1_Click(Area As Integer)
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,次序号 FROM pldc WHERE 料单编号='" & DataCombo1.Text & "' ORDER BY 工序名称,次序号"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
L = ""
Else
L = Adodc1.Recordset.Fields(0)
End If

End Sub


Private Sub Form_Load()
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
DataCombo1.Text = ""
Combo1.Text = "确认"
Combo2 = ""
Combo3.Text = "未"
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,次序号 FROM pldc WHERE 料单编号='" & DataCombo1.Text & "' ORDER BY 工序名称,次序号"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 客户,锅号,品名,颜色,色号,数量,操作,车台,日期,信息,编号,审核 from pLd WHERE  日期 between cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) and 审核='未' ORDER BY 日期,编号"
Adodc2.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"




Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(1) = 0
VSFlexGrid2.ColWidth(4) = 2500
VSFlexGrid2.ColWidth(6) = 1500
VSFlexGrid2.ColWidth(8) = 1500
VSFlexGrid2.ColWidth(6) = 2200

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(5) = 1500

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

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc2.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
DataCombo1.Text = Adodc2.Recordset.Fields(10)

Adodc6.RecordSource = "select * from pldc where  料单编号='" & DataCombo1.Text & "'"
Adodc6.Refresh

'If Not Adodc6.Recordset.EOF Then
'If MsgBox("已存在此料单，是否重新生成？", vbYesNo) = vbNo Then
'Adodc1.RecordSource = "select 工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,次序号 from pldc where  料单编号='" & DataCombo1.Text & "'"
'Adodc1.Refresh
'Exit Sub
'End If
'End If

sql1 = "delete  from pldc WHERE 料单编号='" & DataCombo1.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

Adodc4.RecordSource = "select * from pld where  编号='" & DataCombo1.Text & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.Recordset.MoveFirst

For i = 0 To 10
ZS(i) = Adodc4.Recordset.Fields(i)
Next

mb = 0
For i = 12 To 61
If Adodc4.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next

ProgressBar1.Visible = True
For i = 12 To mb + 12
If Adodc4.Recordset.Fields(i) <> "" Then
sz(0) = Mid(Adodc4.Recordset.Fields(i), 1, InStr(Adodc4.Recordset.Fields(i), "(") - 1)
sz(1) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "(") + 1, InStr(Adodc4.Recordset.Fields(i), ")") - InStr(Adodc4.Recordset.Fields(i), "(") - 1)
sz(2) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), ")") + 1, InStr(Adodc4.Recordset.Fields(i), "-") - InStr(Adodc4.Recordset.Fields(i), ")") - 1)
sz(3) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "-") + 1, InStr(Adodc4.Recordset.Fields(i), "\") - InStr(Adodc4.Recordset.Fields(i), "-") - 1)
sz(4) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "\") + 1, InStr(Adodc4.Recordset.Fields(i), "#") - InStr(Adodc4.Recordset.Fields(i), "\") - 1)
sz(5) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "#") + 1, InStr(Adodc4.Recordset.Fields(i), "^") - InStr(Adodc4.Recordset.Fields(i), "#") - 1)
sz(6) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "^") + 1, InStr(Adodc4.Recordset.Fields(i), "[") - InStr(Adodc4.Recordset.Fields(i), "^") - 1)
sz(7) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "[") + 1, InStr(Adodc4.Recordset.Fields(i), "]") - InStr(Adodc4.Recordset.Fields(i), "[") - 1)
sz(8) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "]") + 1, InStr(Adodc4.Recordset.Fields(i), "{") - InStr(Adodc4.Recordset.Fields(i), "]") - 1)
sz(9) = Mid(Adodc4.Recordset.Fields(i), InStr(Adodc4.Recordset.Fields(i), "{") + 1)

L = i - 11
sql1 = "insert into pldc(审核,锅号,压力,生产类别,配方单,重量,配料打印员,染化助单价,配料日期,生产信息,料单编号,工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,车速,次序号) VALUES('" & ZS(0) & "','" & ZS(1) & "','" & ZS(2) & "','" & ZS(3) & "','" & ZS(4) & "','" & ZS(5) & "','" & ZS(6) & "','" & ZS(7) & "','" & ZS(8) & "','" & ZS(9) & "','" & ZS(10) & "','" & sz(0) & "','" & sz(1) & "','" & sz(2) & "',  " & _
                                                                        "'" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & L & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
ProgressBar1.value = 100 / mb * L
End If
Next
ProgressBar1.Visible = False
End If

Adodc1.RecordSource = "select 工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,次序号 from pldc where  料单编号='" & DataCombo1.Text & "'"
Adodc1.Refresh

End Sub
Private Sub VSFlexGrid2_DblClick()
With VSFlexGrid2
    c = .col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub VSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call VSFlexGrid2_DblClick
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid2.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move r - 1

Adodc1.Recordset.Fields(c - 1) = Text1111.Text
Adodc1.Recordset.Update
    VSFlexGrid2.Text = Text1111.Text
    Text1111.Visible = False
    VSFlexGrid2.SetFocus
End If
End Sub

Private Sub scpfd(bh As String)
Adodc6.RecordSource = "select distinct 工序名称 from pldc where 料单编号='" & bh & "' order by 工序名称"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then Exit Sub
Adodc6.Recordset.MoveFirst

i = 12

Do While Not Adodc6.Recordset.EOF
Adodc7.RecordSource = "select 工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,配料用量,配料单位,车速 from pldc where 料单编号='" & bh & "' and 工序名称='" & Adodc6.Recordset.Fields(0) & "' order by 次序号"
Adodc7.Refresh


If Not Adodc7.Recordset.EOF Then

Adodc7.Recordset.MoveFirst
Do While Not Adodc7.Recordset.EOF
If IsNull(Adodc7.Recordset.Fields(9)) Then
L = ""
Else
L = Trim(Adodc7.Recordset.Fields(9))
End If
sz(i) = Adodc7.Recordset.Fields(0) + "(" + Adodc7.Recordset.Fields(1) + ")" + Adodc7.Recordset.Fields(2) + "-" + Adodc7.Recordset.Fields(3) + "\" + Adodc7.Recordset.Fields(4) + "#" + Trim(Adodc7.Recordset.Fields(5)) + "^" + Adodc7.Recordset.Fields(6) + "[" + Trim(Adodc7.Recordset.Fields(7)) + "]" + Adodc7.Recordset.Fields(8) + "{" + L
i = i + 1
Adodc7.Recordset.MoveNext
Loop
End If

Adodc6.Recordset.MoveNext
Loop

If i < 62 Then
For L = i To 61
sz(L) = ""
Next
End If


Adodc6.RecordSource = "select distinct 审核,锅号,压力,生产类别,配方单,重量,配料打印员,染化助单价,配料日期,生产信息,料单编号 from pldc where 料单编号='" & bh & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then Exit Sub
Adodc6.Recordset.MoveFirst
i = 0
For i = 0 To 10
sz(i) = Adodc6.Recordset.Fields(i)
Next
sz(11) = "确认"

Adodc4.RecordSource = "select 编号 from pld where 编号='" & bh & "'"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
sql1 = "INSERT INTO pld(客户,锅号,品名,颜色,色号,数量,操作,车台,日期,信息,编号,审核,N1,N2,N3,N4,N5,N6,N7,N8,N9,N10,N11,N12,N13,N14,N15,N16,N17,N18,N19,N20,N21,N22,N23,N24,N25,N26,N27,N28,N29,N30,N31,N32,N33,N34,N35,N36,N37,N38,N39,N40,N41,N42,N43,N44,N45,N46,N47,N48,N49,N50) VALUES('" & sz(0) & "','" & sz(1) & "','" & sz(2) & "','" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & sz(10) & "',  " & _
                                                                        "'" & sz(11) & "','" & sz(12) & "','" & sz(13) & "','" & sz(14) & "','" & sz(15) & "','" & sz(16) & "','" & sz(17) & "','" & sz(18) & "','" & sz(19) & "','" & sz(20) & "','" & sz(21) & "', " & _
                                                                        "'" & sz(22) & "','" & sz(23) & "','" & sz(24) & "','" & sz(25) & "','" & sz(26) & "','" & sz(27) & "','" & sz(28) & "','" & sz(29) & "','" & sz(30) & "','" & sz(31) & "','" & sz(32) & "', " & _
                                                                        "'" & sz(33) & "','" & sz(34) & "','" & sz(35) & "','" & sz(36) & "','" & sz(37) & "','" & sz(38) & "','" & sz(39) & "','" & sz(40) & "','" & sz(41) & "','" & sz(42) & "','" & sz(43) & "', " & _
                                                                        "'" & sz(44) & "','" & sz(45) & "','" & sz(46) & "','" & sz(47) & "','" & sz(48) & "','" & sz(49) & "','" & sz(50) & "','" & sz(51) & "','" & sz(52) & "','" & sz(53) & "','" & sz(54) & "', " & _
                                                                        "'" & sz(55) & "','" & sz(56) & "','" & sz(57) & "','" & sz(58) & "','" & sz(59) & "','" & sz(60) & "','" & sz(61) & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Else
sql1 = "delete  from pld where 编号='" & bh & "'"
sql2 = "INSERT INTO pld(客户,锅号,品名,颜色,色号,数量,操作,车台,日期,信息,编号,审核,N1,N2,N3,N4,N5,N6,N7,N8,N9,N10,N11,N12,N13,N14,N15,N16,N17,N18,N19,N20,N21,N22,N23,N24,N25,N26,N27,N28,N29,N30,N31,N32,N33,N34,N35,N36,N37,N38,N39,N40,N41,N42,N43,N44,N45,N46,N47,N48,N49,N50) VALUES('" & sz(0) & "','" & sz(1) & "','" & sz(2) & "','" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & sz(10) & "',  " & _
                                                                        "'" & sz(11) & "','" & sz(12) & "','" & sz(13) & "','" & sz(14) & "','" & sz(15) & "','" & sz(16) & "','" & sz(17) & "','" & sz(18) & "','" & sz(19) & "','" & sz(20) & "','" & sz(21) & "', " & _
                                                                        "'" & sz(22) & "','" & sz(23) & "','" & sz(24) & "','" & sz(25) & "','" & sz(26) & "','" & sz(27) & "','" & sz(28) & "','" & sz(29) & "','" & sz(30) & "','" & sz(31) & "','" & sz(32) & "', " & _
                                                                        "'" & sz(33) & "','" & sz(34) & "','" & sz(35) & "','" & sz(36) & "','" & sz(37) & "','" & sz(38) & "','" & sz(39) & "','" & sz(40) & "','" & sz(41) & "','" & sz(42) & "','" & sz(43) & "', " & _
                                                                        "'" & sz(44) & "','" & sz(45) & "','" & sz(46) & "','" & sz(47) & "','" & sz(48) & "','" & sz(49) & "','" & sz(50) & "','" & sz(51) & "','" & sz(52) & "','" & sz(53) & "','" & sz(54) & "', " & _
                                                                        "'" & sz(55) & "','" & sz(56) & "','" & sz(57) & "','" & sz(58) & "','" & sz(59) & "','" & sz(60) & "','" & sz(61) & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If

End Sub





