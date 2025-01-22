VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formh100 
   BackColor       =   &H00C0E0FF&
   Caption         =   "模板设置"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   14805
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转入"
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Index           =   6
      Left            =   10200
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Index           =   5
      Left            =   10200
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Index           =   4
      Left            =   10200
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Index           =   3
      Left            =   10200
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Index           =   2
      Left            =   10200
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选取"
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   7680
      Top             =   9960
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   8160
      Top             =   10080
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   6960
      Top             =   10080
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
      Left            =   7080
      Top             =   10200
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
      Left            =   6720
      Top             =   10080
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
      Height          =   375
      Left            =   7080
      Top             =   9840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   6840
      Top             =   9840
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
      Height          =   375
      Left            =   6840
      Top             =   10080
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
      Left            =   7320
      Top             =   9960
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
      Left            =   6840
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
      Left            =   7440
      Top             =   9840
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
      Left            =   7560
      Top             =   9840
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
      Bindings        =   "Formh100.frx":0000
      Height          =   4455
      Left            =   360
      TabIndex        =   11
      Top             =   4680
      Width           =   13935
      _cx             =   24580
      _cy             =   7858
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "常规工艺"
      Height          =   3495
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   13935
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   6
         Left            =   3720
         TabIndex        =   38
         Text            =   "Text5"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   5
         Left            =   3720
         TabIndex        =   37
         Text            =   "Text5"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   4
         Left            =   3720
         TabIndex        =   36
         Text            =   "Text5"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   3
         Left            =   3720
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   2
         Left            =   3720
         TabIndex        =   34
         Text            =   "Text5"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   1
         Left            =   3720
         TabIndex        =   33
         Text            =   "Text5"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   0
         Left            =   3720
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Index           =   6
         Left            =   11160
         TabIndex        =   31
         Text            =   "Text3"
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Index           =   5
         Left            =   11160
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Index           =   4
         Left            =   11160
         TabIndex        =   29
         Text            =   "Text3"
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Index           =   3
         Left            =   11160
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Index           =   2
         Left            =   11160
         TabIndex        =   27
         Text            =   "Text3"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Index           =   1
         Left            =   11160
         TabIndex        =   26
         Text            =   "Text3"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Index           =   0
         Left            =   11160
         TabIndex        =   25
         Text            =   "Text3"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Index           =   1
         Left            =   9840
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Index           =   0
         Left            =   9840
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   0
         Left            =   8400
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   1
         Left            =   8400
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   2
         Left            =   8400
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   3
         Left            =   8400
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   4
         Left            =   8400
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   5
         Left            =   8400
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   6
         Left            =   8400
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh100.frx":0015
         Height          =   330
         Index           =   0
         Left            =   6960
         TabIndex        =   39
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh100.frx":002A
         Height          =   330
         Index           =   0
         Left            =   4440
         TabIndex        =   40
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh100.frx":003F
         Height          =   330
         Index           =   0
         Left            =   1440
         TabIndex        =   41
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "工艺工序"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh100.frx":0054
         Height          =   330
         Index           =   1
         Left            =   1440
         TabIndex        =   42
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "模板编号"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh100.frx":0069
         Height          =   330
         Index           =   2
         Left            =   1440
         TabIndex        =   43
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "染化助库名"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh100.frx":007E
         Height          =   330
         Index           =   3
         Left            =   1440
         TabIndex        =   44
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "标志"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh100.frx":0094
         Height          =   330
         Index           =   4
         Left            =   1080
         TabIndex        =   45
         Top             =   2880
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "模板编号"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   5
         Left            =   720
         TabIndex        =   46
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   6
         Left            =   120
         TabIndex        =   47
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   7
         Left            =   2040
         TabIndex        =   48
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh100.frx":00A9
         Height          =   330
         Index           =   1
         Left            =   4440
         TabIndex        =   49
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh100.frx":00BE
         Height          =   330
         Index           =   2
         Left            =   4440
         TabIndex        =   50
         Top             =   1560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh100.frx":00D3
         Height          =   330
         Index           =   3
         Left            =   4440
         TabIndex        =   51
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh100.frx":00E8
         Height          =   330
         Index           =   4
         Left            =   4440
         TabIndex        =   52
         Top             =   2280
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh100.frx":00FD
         Height          =   330
         Index           =   5
         Left            =   4440
         TabIndex        =   53
         Top             =   2640
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh100.frx":0112
         Height          =   330
         Index           =   6
         Left            =   4440
         TabIndex        =   54
         Top             =   3000
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh100.frx":0127
         Height          =   330
         Index           =   1
         Left            =   6960
         TabIndex        =   55
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh100.frx":013C
         Height          =   330
         Index           =   2
         Left            =   6960
         TabIndex        =   56
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh100.frx":0151
         Height          =   330
         Index           =   3
         Left            =   6960
         TabIndex        =   57
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh100.frx":0166
         Height          =   330
         Index           =   4
         Left            =   6960
         TabIndex        =   58
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh100.frx":017B
         Height          =   330
         Index           =   5
         Left            =   6960
         TabIndex        =   59
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh100.frx":0190
         Height          =   330
         Index           =   6
         Left            =   6960
         TabIndex        =   60
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "模板编号"
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
         Left            =   360
         TabIndex        =   70
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "编号"
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
         Index           =   21
         Left            =   3720
         TabIndex        =   69
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "备注"
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
         Left            =   11160
         TabIndex        =   68
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "序号"
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
         Left            =   9840
         TabIndex        =   67
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "染化助库"
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
         Index           =   13
         Left            =   360
         TabIndex        =   66
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "染化助代码"
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
         Index           =   12
         Left            =   360
         TabIndex        =   65
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "工艺名称"
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
         Index           =   9
         Left            =   360
         TabIndex        =   64
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "配方"
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
         Index           =   7
         Left            =   8400
         TabIndex        =   63
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "单位"
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
         Index           =   6
         Left            =   6960
         TabIndex        =   62
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "染化助名称"
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
         Index           =   5
         Left            =   4440
         TabIndex        =   61
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Formh100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public c, r As Integer
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Command1_Click()
On Error Resume Next
Call MXOutadodcToExcel(VSFlexGrid1, DataCombo1(0).Text + DataCombo1(1).Text)
End Sub

Private Sub Command2_Click()
On Error Resume Next
If DataCombo1(0).Text = "" Or DataCombo1(1).Text = "" Then
MsgBox ("工艺名称、编号须填完整！")
Exit Sub
End If

L = 1
For i = 0 To 6     '''''''''''''''''''''''''
If Text1(i).Text <> "" Then
Adodc7.Recordset.AddNew
Adodc7.Recordset.Fields(0) = DataCombo1(0).Text
Adodc7.Recordset.Fields(1) = DataCombo1(1).Text
Adodc7.Recordset.Fields(2) = DataCombo1(2).Text
Adodc7.Recordset.Fields(3) = DataCombo1(3).Text
Adodc7.Recordset.Fields(4) = DataCombo2(i).Text
Adodc7.Recordset.Fields(5) = DataCombo3(i).Text
Adodc7.Recordset.Fields(6) = Text1(i).Text
Adodc7.Recordset.Fields(7) = Text2(i).Text
Adodc7.Recordset.Fields(8) = Text3(i).Text
Adodc7.Recordset.Update
End If
Next
       Adodc7.RecordSource = "SELECT * FROM CGGYMB WHERE  模板编号='" & DataCombo1(1).Text & "' ORDER BY 工艺名称,序号"
       Adodc7.Refresh
                '''''''''''''''''''''''
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text3(i).Text = ""
Text5(i).Text = ""
Next

End Sub

Private Sub Command3_Click()
If DataCombo1(0).Text = "" Or DataCombo1(1).Text = "" Then
MsgBox ("工艺名称、编号须填完整！")
Exit Sub
End If

Adodc7.Recordset.Fields(0) = DataCombo1(0).Text
Adodc7.Recordset.Fields(1) = DataCombo1(1).Text
Adodc7.Recordset.Fields(2) = DataCombo1(2).Text
Adodc7.Recordset.Fields(3) = DataCombo1(3).Text
Adodc7.Recordset.Fields(4) = DataCombo2(0).Text
Adodc7.Recordset.Fields(5) = DataCombo3(0).Text
Adodc7.Recordset.Fields(6) = Text1(0).Text
Adodc7.Recordset.Fields(7) = Text2(0).Text
Adodc7.Recordset.Fields(8) = Text3(0).Text
Adodc7.Recordset.Update
       Adodc7.RecordSource = "SELECT * FROM CGGYMB WHERE  模板编号='" & DataCombo1(1).Text & "' ORDER BY 工艺名称,序号"
       Adodc7.Refresh
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text3(i).Text = ""
Text5(i).Text = ""
Next
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False

End Sub

Private Sub Command4_Click()
On Error Resume Next
If Adodc7.Recordset.EOF Then Exit Sub
Adodc7.Recordset.Delete
Adodc7.RecordSource = "SELECT * FROM CGGYMB WHERE  模板编号='" & DataCombo1(1).Text & "' ORDER BY 工艺名称,序号"
Adodc7.Refresh

For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text3(i).Text = ""
Text5(i).Text = ""
Next
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False

End Sub

Private Sub Command5_Click()
If MsgBox("确定把外部常规工艺转入内部吗？", vbYesNo) = vbNo Then Exit Sub
lo = "d:\数据库\bfrz\" + ljb + "\HYS.mdb"
'Adodc12.Database.Execute "delete * from CGGYMB"
'Adodc11.Database.Execute "insert into CGGYMB in'" & lo & "' select * from CGGYMB"
MsgBox ("转入成功！")
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Formh233.DataCombo7 = DataCombo1(1)
Unload Me
End Sub

Private Sub Command8_Click()
Adodc1.Refresh
Adodc4.Refresh
Adodc8.Refresh
Adodc7.RecordSource = "SELECT * FROM CGGYMB WHERE 模板编号='" & DataCombo1(1).Text & "' ORDER BY 工艺名称,序号"
Adodc7.Refresh
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text3(i).Text = ""
Text2(i) = i + 1
Next
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub



Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
       Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc4.RecordSource = "select 模板编号 from CGGYMB where 工艺名称='" & DataCombo1(0).Text & "'group by 模板编号 ORDER BY 模板编号"
       Adodc4.Refresh

       Case 1
       Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc7.RecordSource = "SELECT * FROM CGGYMB WHERE  模板编号='" & DataCombo1(1).Text & "' ORDER BY 工艺名称,序号"
       Adodc7.Refresh
       Case 2
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH where 染化助库名='" & DataCombo1(2).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc10.RecordSource = "SELECT 标志 FROM RHZH where 染化助库名='" & DataCombo1(2).Text & "' GROUP BY 标志 "
       Adodc10.Refresh
       
       
       Case 3
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH  where 染化助库名='" & DataCombo1(2).Text & "' AND 标志 LIKE '%'+'" & DataCombo1(3).Text & "'+'%' GROUP BY 染料名称"
       Adodc8.Refresh
End Select

End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 0
       Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc4.RecordSource = "select 模板编号 from CGGYMB where 工艺名称='" & DataCombo1(0).Text & "'group by 模板编号 ORDER BY 模板编号"
       Adodc4.Refresh

       Case 1
       Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc7.RecordSource = "SELECT * FROM CGGYMB WHERE  模板编号='" & DataCombo1(1).Text & "' ORDER BY 工艺名称,序号"
       Adodc7.Refresh
       Case 2
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH where 染化助库名='" & DataCombo1(2).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc10.RecordSource = "SELECT 标志 FROM RHZH where 染化助库名='" & DataCombo1(2).Text & "' GROUP BY 标志 "
       Adodc10.Refresh
       If InStr(DataCombo1(2), "染料") > 0 Then
       For i = 0 To 6
       DataCombo3(i).Text = "%"
       Next
       Else
       For i = 0 To 6
       DataCombo3(i).Text = "g/l"
       Next
       End If

       Case 3
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH  where 染化助库名='" & DataCombo1(2).Text & "' AND 标志 LIKE '%'+'" & DataCombo1(3).Text & "'+'%' GROUP BY 染料名称"
       Adodc8.Refresh
End Select
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode

End Sub

Private Sub Form_Load()

'On Error Resume Next
Dim L As String



For i = 0 To 4
DataCombo1(i) = ""
Next

For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i) = i + 1
Text3(i).Text = ""
Text5(i).Text = ""
Next

Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 工艺工序 from gx group by 工艺工序 ORDER BY 工艺工序"
Adodc1.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 模板编号 from CGGYMB group by 模板编号 ORDER BY 模板编号"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select dw,IP from dw group by dw,IP ORDER BY IP"
Adodc5.Refresh


Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "pfda"
Adodc6.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "SELECT * FROM CGGYMB WHERE 模板编号='" & DataCombo1(1).Text & "' ORDER BY 工艺名称,序号"
Adodc7.Refresh

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT 染化助库名 FROM RHZH GROUP BY 染化助库名"
Adodc8.Refresh

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "SELECT 染化助库名 FROM RHZH GROUP BY 染化助库名"
Adodc9.Refresh

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "SELECT 负责人姓名 FROM GR GROUP BY 负责人姓名"
Adodc11.Refresh

Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

DataCombo1(0).TabIndex = 0

VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 1200
VSFlexGrid1.ColWidth(5) = 2000
VSFlexGrid1.ColWidth(6) = 1200
VSFlexGrid1.ColWidth(7) = 1000
VSFlexGrid1.ColWidth(9) = 2600



End Sub

Private Sub Label1_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case 16
       DataCombo1(16).Enabled = True
       Case 12
       DataCombo1(12).Enabled = True
       Case 11
       DataCombo1(10).Enabled = True
       Case 8
       DataCombo1(11).Enabled = True
       Case 9
       DataCombo1(12).Enabled = True
End Select
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text5_Change(Index As Integer)
Select Case Index
       Case Index
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT 染料名称 FROM rhzh where 简码 like '%'+'" & Text5(Index) & "'+'%' and 染化助库名='" & DataCombo1(2) & "'"
Adodc8.Refresh
If Not Adodc8.Recordset.EOF Then
DataCombo2(Index) = Adodc8.Recordset.Fields(0)
Else
DataCombo2(Index) = ""
End If
End Select
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc7.Recordset.MoveFirst
Adodc7.Recordset.Move rs - 1
DataCombo1(0).Text = Adodc7.Recordset.Fields(0)
DataCombo1(1).Text = Adodc7.Recordset.Fields(1)
DataCombo1(2).Text = Adodc7.Recordset.Fields(2)
DataCombo1(3).Text = Adodc7.Recordset.Fields(3)
DataCombo2(0).Text = Adodc7.Recordset.Fields(4)
DataCombo3(0).Text = Adodc7.Recordset.Fields(5)
Text1(0).Text = Adodc7.Recordset.Fields(6)
Text2(0).Text = Adodc7.Recordset.Fields(7)
Text3(0).Text = Adodc7.Recordset.Fields(8)
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
End Sub


Private Sub MSF()
On Error Resume Next
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSF
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc7.Recordset.MoveFirst
Adodc7.Recordset.Move r - 1
Adodc7.Recordset.Fields(c - 1) = Text1111.Text
Adodc7.Recordset.Update
    Text1111.Visible = False
    VSFlexGrid1.Text = Text1111.Text
    VSFlexGrid1.SetFocus
End If
End Sub






