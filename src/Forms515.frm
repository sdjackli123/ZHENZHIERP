VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms515 
   BackColor       =   &H00C0E0FF&
   Caption         =   "光坯检验"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   14730
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1920
      Top             =   7200
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
   Begin VB.TextBox Text2 
      Height          =   735
      Index           =   3
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "Text2"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   5160
      TabIndex        =   47
      Text            =   "Text2"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   360
      Width           =   855
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
      Height          =   615
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1080
      Width           =   855
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
      Height          =   615
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Text2"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Text2"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   7200
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   6480
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   5760
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Forms515.frx":0000
      Left            =   7080
      List            =   "Forms515.frx":000D
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   5040
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   4320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms515.frx":0023
      Height          =   1695
      Left            =   1440
      TabIndex        =   50
      Top             =   5160
      Width           =   12375
      _cx             =   21828
      _cy             =   2990
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
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "疵点合计"
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
      Index           =   10
      Left            =   8040
      TabIndex        =   48
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "检验员"
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
      Index           =   9
      Left            =   5160
      TabIndex        =   46
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "缸号"
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
      Index           =   15
      Left            =   1440
      TabIndex        =   41
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "匹号"
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
      Index           =   17
      Left            =   3840
      TabIndex        =   40
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   7200
      TabIndex        =   37
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   7200
      TabIndex        =   36
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "其它"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   8
      Left            =   7200
      TabIndex        =   34
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   6480
      TabIndex        =   33
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   6480
      TabIndex        =   32
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   5760
      TabIndex        =   30
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   5760
      TabIndex        =   29
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "三丝"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   6480
      TabIndex        =   27
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "水折"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   5760
      TabIndex        =   26
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "等级"
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
      Index           =   12
      Left            =   7080
      TabIndex        =   25
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   5040
      TabIndex        =   23
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   4320
      TabIndex        =   22
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3600
      TabIndex        =   21
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   2880
      TabIndex        =   20
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2160
      TabIndex        =   19
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1440
      TabIndex        =   18
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   5040
      TabIndex        =   17
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   4320
      TabIndex        =   16
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3600
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   2880
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2160
      TabIndex        =   13
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1440
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "棉结"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   5040
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "油点"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   4320
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "色沾"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   3600
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "沾污"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   2880
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "破洞"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "断丝"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   1440
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "Forms515"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Text2(2) = "" Then
MsgBox ("请输入验布编号")
Exit Sub
End If

If Combo1 = "" Then
MsgBox ("请输入等级")
Exit Sub
End If

Adodc1.Recordset.AddNew
For i = 0 To 3
Adodc1.Recordset.Fields(i) = Text2(i)
Next
Adodc1.Recordset.Fields(4) = Combo1
For i = 0 To 8
Adodc1.Recordset.Fields(5 + i) = Text1(i)
Next
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text2(2) = "" Then
MsgBox ("请输入验布编号")
Exit Sub
End If

If Combo1 = "" Then
MsgBox ("请输入等级")
Exit Sub
End If
For i = 0 To 3
Adodc1.Recordset.Fields(i) = Text2(i)
Next
Adodc1.Recordset.Fields(4) = Combo1
For i = 0 To 8
Adodc1.Recordset.Fields(5 + i) = Text1(i)
Next
Adodc1.Recordset.Update
Adodc1.Refresh

End Sub

Private Sub Command4_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
For i = 0 To 17
Text1(i) = ""
Text2(i) = ""
Next
Combo1 = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from ybbb where 缸号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "'"
Adodc1.Refresh
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case Index
Text1(Index) = Val(Text1(Index)) + 1
End Select
End Sub

Private Sub Label3_Click(Index As Integer)
Select Case Index
       Case Index
If Val(Text1(Index)) <= 0 Then Exit Sub
Text1(Index) = Val(Text1(Index)) - 1
End Select
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
If Index < 11 Then
bh = 0
For i = 0 To 10
bh = bh + Val(Text1(i))
Next
End If
Text2(3) = bh
End Select
End Sub

Private Sub Text2_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 1
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from ybbb where 缸号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "'"
Adodc1.Refresh
End Select
End Sub

Private Sub VSFlexGrid1_dblClick()
For i = 0 To 3
Text2(i) = Adodc1.Recordset.Fields(i)
Next
Combo1 = Adodc1.Recordset.Fields(4)
For i = 0 To 8
Text1(i) = Adodc1.Recordset.Fields(5 + i)
Next
End Sub
