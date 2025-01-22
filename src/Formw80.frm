VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw80 
   BackColor       =   &H00C0E0FF&
   Caption         =   "客户账目查询---收款"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   375
      Left            =   13080
      Top             =   10560
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   855
      Left            =   11160
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   9000
      TabIndex        =   22
      Top             =   360
      Width           =   2055
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "客户"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "司机"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6360
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   960
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   8160
      Top             =   10680
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
      Left            =   8400
      Top             =   10560
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Left            =   8400
      Top             =   10680
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw80.frx":0000
      Height          =   7575
      Left            =   600
      TabIndex        =   19
      Top             =   1560
      Width           =   21015
      _cx             =   37068
      _cy             =   13361
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formw80.frx":0015
      Height          =   330
      Left            =   6960
      TabIndex        =   18
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw80.frx":002A
      Height          =   330
      Left            =   6960
      TabIndex        =   17
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "司机"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成查询"
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "凭证生成"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "结转下期"
      Height          =   735
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新刷"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   329056257
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   329056257
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12480
      TabIndex        =   11
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   329056257
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   329056257
      CurrentDate     =   36892
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9000
      Top             =   10440
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
      Left            =   9000
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8520
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
      Left            =   8640
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
      Height          =   330
      Left            =   8760
      Top             =   10680
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
      Height          =   330
      Left            =   8280
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   375
      Left            =   8520
      Top             =   10560
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formw80.frx":0040
      Height          =   3255
      Left            =   600
      TabIndex        =   20
      Top             =   9240
      Width           =   7695
      _cx             =   13573
      _cy             =   5741
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
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "下期起初"
      Height          =   375
      Index           =   0
      Left            =   12480
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "司机"
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
      Left            =   5880
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期范围"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Formw80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public S1, S2 As Integer
Dim cdbhf As Integer
Private Sub Command1_Click()
'On Error Resume Next
Command1.Enabled = False
FP = CDate(DTPicker4.value) + 1
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "YSHZCX('" & DTPicker3.value & "','" & DTPicker4.value & "','" & FP & "','" & yhm & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel
Command1.Enabled = True
MsgBox ("刷新成功！")
Adodc6.RecordSource = "SELECT * FROM YSKZCX where 类别='" & yhm & "' order by 客户"
Adodc6.Refresh
Adodc5.RecordSource = "SELECT 司机, round(sum(isnull(本期累计未开,0)),2) as 合计本期累计未开,round(sum(isnull(欠款,0)),2) as 合计欠款 FROM YSKZCX where 类别='" & yhm & "' group by 司机"
Adodc5.Refresh

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid2.ColWidth(0) = 300
VSFlexGrid2.ColWidth(1) = 1200
VSFlexGrid2.ColWidth(2) = 1500
VSFlexGrid2.ColWidth(3) = 1500
VSFlexGrid2.SubtotalPosition = flexSTBelow
VSFlexGrid2.Subtotal flexSTSum, 0, 2, , vbGreen
VSFlexGrid2.Subtotal flexSTSum, 0, 3, , vbGreen
For i = 1 To 11
VSFlexGrid1.ColWidth(i) = 1300
Next
VSFlexGrid1.ColWidth(12) = 2000
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call OutadodcToExcel11(VSFlexGrid1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, "衡水广兴服饰有限公司 客户账目查询--收款" + "截止日期:" + Trim(DTPicker4.value))
End Sub



Private Sub Command4_Click()
On Error Resume Next

sql1 = ""

If Check4.value = 1 Then
sql1 = sql1 + "客户='" & DataCombo2 & "' and "
End If

If Check6.value = 1 Then
sql1 = sql1 + "司机 like '%'+'" & DataCombo1 & "'+'%' and "
End If


If Len(sql1) > 1 Then
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 3)
Adodc6.RecordSource = "SELECT * FROM YSKZCX where (" + sql1 + ") and 类别='" & yhm & "' order by 客户"
Adodc6.Refresh
Adodc5.RecordSource = "SELECT round(sum(isnull(本期累计未开,0)),2) as 合计本期累计未开,round(sum(isnull(欠款,0)),2) as 合计欠款 FROM YSKZCX where (" + sql1 + ") and 类别='" & yhm & "'"
Adodc5.Refresh
End If
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If


VSFlexGrid1.ColWidth(0) = 300
For i = 1 To 11
VSFlexGrid1.ColWidth(i) = 1300
Next
VSFlexGrid1.ColWidth(12) = 2000

VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 1, 4, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 5, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 6, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 7, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 8, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 9, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 10, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 11, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 12, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 13, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 14, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 15, , vbGreen

End Sub

Private Sub Command5_Click()
If MsgBox("确定结转下期吗，下期起初为：" + Trim(DTPicker1.value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定结转下期吗?", vbYesNo) = vbNo Then Exit Sub

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "yshzjz('" & DTPicker1.value & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel

MsgBox ("结转成功！，在期初设置中可以查询！")

End Sub

Private Sub Command6_Click()
If MsgBox("操作日期为：" + Trim(DTPicker2.value) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("操作期间为：" + Trim(Month(DTPicker2.value)) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定生成应收系列的凭证吗？", vbYesNo) = vbNo Then Exit Sub
Call CPFHPZ(DTPicker3.value, DTPicker4.value, DTPicker2.value)
End Sub

Private Sub Command7_Click()
Formw1132.DTPicker1.value = DTPicker2.value
Formw1132.Show
Unload Me
End Sub

Private Sub DataCombo1_Click(Area As Integer)
'Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc6.RecordSource = "SELECT * FROM YSKZCX WHERE 业务 like '%'+'" & DataCombo1 & "'+'%' order by 客户"
'Adodc6.Refresh
'Adodc5.RecordSource = "SELECT round(sum(isnull(上期累计应收,0)),2) as 合计上期累计应收,round(sum(isnull(本期应收款,0)),2) as 合计本期应收款,round(sum(isnull(本期累计应收款,0)),2) as 合计本期累计应收款,round(sum(isnull(本期已收款,0)),2) as 合计本期已收款,round(sum(isnull(上期累计开票,0)),2) as 合计上期累计开票,round(sum(isnull(本期开票,0)),2) as 合计本期开票,round(sum(isnull(本期累计开票,0)),2) as 合计本期累计开票,round(sum(isnull(上期累计未开票,0)),2) as 合计上期累计未开票,round(sum(isnull(本期未开,0)),2) as 合计本期未开,round(sum(isnull(本期累计未开,0)),2) as 合计本期累计未开,round(sum(isnull(欠款,0)),2) as 合计欠款 FROM YSKZCX WHERE 业务 like '%'+'" & DataCombo1 & "'+'%'"
'Adodc5.Refresh
End Sub


Private Sub Form_Load()
On Error Resume Next


DTPicker1.value = Date
DTPicker3.value = Date
DTPicker2.value = Date
DTPicker4.value = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
Text1 = ""
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 代码,简称 from KHZL  GROUP BY 代码,简称"
Adodc1.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT * FROM YSKZCX WHERE 客户='" & DataCombo2.Text & "'  order by 客户"
Adodc6.Refresh
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select ip from KHZL  GROUP BY ip"
Adodc8.Refresh
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select distinct 司机 from YSKZCX  "
Adodc11.Refresh
VSFlexGrid1.ColWidth(0) = 300
For i = 1 To 11
VSFlexGrid1.ColWidth(i) = 1300
Next
VSFlexGrid1.ColWidth(12) = 2000

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

Private Sub Text1_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%'  group by 简称"
Adodc1.Refresh
End Sub

Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
S1 = VSFlexGrid1.RowSel
End Sub

Private Sub VSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
S2 = VSFlexGrid1.RowSel
End Sub

Private Sub CPFHPZ(DT1 As Date, dt2 As Date, dt3 As Date)
'On Error Resume Next
Dim djs As Integer

Adodc4.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "' and 制单 like '%自动-发货%'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
If MsgBox("已有应收生成凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CLZZPZ WHERE 制单 like '%自动-发货%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If


Adodc9.RecordSource = "SELECT * FROM YSKZCX where 本期应收款<>0 order by 客户"
Adodc9.Refresh

If Adodc9.Recordset.EOF Then Exit Sub

Adodc10.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc10.Refresh
If Adodc10.Recordset.EOF Then
PZH = "5-1"
Else
Adodc10.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-2)) FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc10.Refresh
PZH = "5-" + Trim(Adodc10.Recordset.Fields(0) + 1)
End If
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 7

sql1 = "INSERT INTO CLZZPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('主营收入','应收账款','" & Adodc9.Recordset.Fields(1) & "','" & Adodc9.Recordset.Fields(3) & "','','" & PZH & "','" & dt3 & "','','','','自动-发货','','','')"
sql2 = "INSERT INTO CLZZPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('主营收入','主营业务收入','','','" & Adodc9.Recordset.Fields(3) & "','" & PZH & "','" & dt3 & "','','','','自动-发货','','','')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
Adodc10.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc10.Refresh
If Adodc10.Recordset.EOF Then
PZH = "5-1"
Else
Adodc10.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-2)) FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc10.Refresh
PZH = "5-" + Trim(Adodc10.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("成品发货单转账成功！" + "生成" + Str(KLLLL) + "凭证")
End Sub





