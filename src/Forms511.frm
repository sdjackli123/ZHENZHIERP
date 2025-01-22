VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms511 
   BackColor       =   &H00C0E0FF&
   Caption         =   "产量扫描"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "产量查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   229
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "操作模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   225
      Top             =   5040
      Width           =   4215
      Begin VB.OptionButton Option4 
         Caption         =   "手动"
         Height          =   495
         Left            =   2640
         TabIndex        =   227
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "自动"
         Height          =   495
         Left            =   840
         TabIndex        =   226
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   223
      Text            =   "Forms511.frx":0000
      Top             =   6240
      Width           =   6135
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   4680
      Top             =   10560
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
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
      Left            =   4800
      Top             =   10680
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
   Begin VB.OptionButton Option2 
      Caption         =   "操作输入"
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
      Left            =   6840
      TabIndex        =   220
      Top             =   4320
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "机台输入"
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
      Left            =   6840
      TabIndex        =   219
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   218
      Text            =   "Text11"
      Top             =   1800
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   4080
      Top             =   10680
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
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4200
      ItemData        =   "Forms511.frx":0007
      Left            =   12000
      List            =   "Forms511.frx":0009
      Style           =   1  'Checkbox
      TabIndex        =   207
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "操作员或锅号或班次"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   8520
      TabIndex        =   103
      Top             =   2880
      Width           =   6615
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   4320
         TabIndex        =   228
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFF00&
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         TabIndex        =   224
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   216
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   215
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   1800
         TabIndex        =   211
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   120
         TabIndex        =   210
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   960
         TabIndex        =   209
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "/"
         Height          =   495
         Index           =   13
         Left            =   5760
         TabIndex        =   206
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   5040
         TabIndex        =   116
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "清除"
         Height          =   495
         Left            =   5760
         TabIndex        =   115
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   4320
         TabIndex        =   114
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   3480
         TabIndex        =   113
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   3480
         TabIndex        =   112
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   2640
         TabIndex        =   111
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   2640
         TabIndex        =   110
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   1800
         TabIndex        =   109
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   1800
         TabIndex        =   108
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   960
         TabIndex        =   107
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   960
         TabIndex        =   106
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   105
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   104
         Top             =   360
         Width           =   615
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms511.frx":000B
      Height          =   1095
      Left            =   360
      TabIndex        =   62
      Top             =   2400
      Width           =   6255
      _cx             =   11033
      _cy             =   1931
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
      FormatString    =   $"Forms511.frx":0020
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
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   102
      Text            =   "Text10"
      Top             =   1080
      Width           =   2655
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forms511.frx":00F5
      Height          =   1335
      Left            =   360
      TabIndex        =   63
      Top             =   3600
      Width           =   6255
      _cx             =   11033
      _cy             =   2355
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
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Forms511.frx":010A
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "报表"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   98
      Text            =   "Text1"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   360
      TabIndex        =   64
      Top             =   7800
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   116
         Left            =   8640
         TabIndex        =   205
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   115
         Left            =   8640
         TabIndex        =   204
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   114
         Left            =   8640
         TabIndex        =   203
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   113
         Left            =   8640
         TabIndex        =   202
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   112
         Left            =   8640
         TabIndex        =   201
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   111
         Left            =   8640
         TabIndex        =   200
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   110
         Left            =   8640
         TabIndex        =   199
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   109
         Left            =   8640
         TabIndex        =   198
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   108
         Left            =   7920
         TabIndex        =   197
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   107
         Left            =   7920
         TabIndex        =   196
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   106
         Left            =   7920
         TabIndex        =   195
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   105
         Left            =   7920
         TabIndex        =   194
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   104
         Left            =   7920
         TabIndex        =   193
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   103
         Left            =   7920
         TabIndex        =   192
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   102
         Left            =   7920
         TabIndex        =   191
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   101
         Left            =   7920
         TabIndex        =   190
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   100
         Left            =   7920
         TabIndex        =   189
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   99
         Left            =   7200
         TabIndex        =   188
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   98
         Left            =   7200
         TabIndex        =   187
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   97
         Left            =   7200
         TabIndex        =   186
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   96
         Left            =   7200
         TabIndex        =   185
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   95
         Left            =   7200
         TabIndex        =   184
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   94
         Left            =   7200
         TabIndex        =   183
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   93
         Left            =   7200
         TabIndex        =   182
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   92
         Left            =   7200
         TabIndex        =   181
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   91
         Left            =   7200
         TabIndex        =   180
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   90
         Left            =   6480
         TabIndex        =   179
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   8520
         X2              =   8520
         Y1              =   120
         Y2              =   5520
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   89
         Left            =   6480
         TabIndex        =   178
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   88
         Left            =   6480
         TabIndex        =   177
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   87
         Left            =   6480
         TabIndex        =   176
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   86
         Left            =   6480
         TabIndex        =   175
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   85
         Left            =   6480
         TabIndex        =   174
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   84
         Left            =   6480
         TabIndex        =   173
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   83
         Left            =   6480
         TabIndex        =   172
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   7800
         X2              =   7800
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   82
         Left            =   6480
         TabIndex        =   171
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   81
         Left            =   5760
         TabIndex        =   170
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   80
         Left            =   5760
         TabIndex        =   169
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   79
         Left            =   5760
         TabIndex        =   168
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   78
         Left            =   5760
         TabIndex        =   167
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   77
         Left            =   5760
         TabIndex        =   166
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   76
         Left            =   5760
         TabIndex        =   165
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   7080
         X2              =   7080
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   75
         Left            =   5760
         TabIndex        =   164
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   74
         Left            =   5760
         TabIndex        =   163
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   73
         Left            =   5760
         TabIndex        =   162
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   72
         Left            =   5040
         TabIndex        =   161
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   71
         Left            =   5040
         TabIndex        =   160
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   70
         Left            =   5040
         TabIndex        =   159
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   69
         Left            =   5040
         TabIndex        =   158
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   6360
         X2              =   6360
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   68
         Left            =   5040
         TabIndex        =   157
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   67
         Left            =   5040
         TabIndex        =   156
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   66
         Left            =   5040
         TabIndex        =   155
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   65
         Left            =   5040
         TabIndex        =   154
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   64
         Left            =   5040
         TabIndex        =   153
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   63
         Left            =   4320
         TabIndex        =   152
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   62
         Left            =   4320
         TabIndex        =   151
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   5640
         X2              =   5640
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   61
         Left            =   4320
         TabIndex        =   150
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   60
         Left            =   4320
         TabIndex        =   149
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   59
         Left            =   4320
         TabIndex        =   148
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   58
         Left            =   4320
         TabIndex        =   147
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   57
         Left            =   4320
         TabIndex        =   146
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   56
         Left            =   4320
         TabIndex        =   145
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   55
         Left            =   4320
         TabIndex        =   144
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   4920
         X2              =   4920
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   54
         Left            =   3600
         TabIndex        =   143
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   53
         Left            =   3600
         TabIndex        =   142
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   52
         Left            =   3600
         TabIndex        =   141
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   51
         Left            =   3600
         TabIndex        =   140
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   50
         Left            =   3600
         TabIndex        =   139
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   49
         Left            =   3600
         TabIndex        =   138
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   48
         Left            =   3600
         TabIndex        =   137
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   4200
         X2              =   4200
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   47
         Left            =   3600
         TabIndex        =   136
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   46
         Left            =   3600
         TabIndex        =   135
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   45
         Left            =   2880
         TabIndex        =   134
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   44
         Left            =   2880
         TabIndex        =   133
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   43
         Left            =   2880
         TabIndex        =   132
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   42
         Left            =   2880
         TabIndex        =   131
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   41
         Left            =   2880
         TabIndex        =   130
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   3480
         X2              =   3480
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   40
         Left            =   2880
         TabIndex        =   129
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   39
         Left            =   2880
         TabIndex        =   128
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   38
         Left            =   2880
         TabIndex        =   127
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   37
         Left            =   2880
         TabIndex        =   126
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   36
         Left            =   2160
         TabIndex        =   125
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   35
         Left            =   2160
         TabIndex        =   124
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   34
         Left            =   2160
         TabIndex        =   123
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   3
         Visible         =   0   'False
         X1              =   2760
         X2              =   2760
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   33
         Left            =   2160
         TabIndex        =   122
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   32
         Left            =   2160
         TabIndex        =   121
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   31
         Left            =   2160
         TabIndex        =   120
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   30
         Left            =   2160
         TabIndex        =   119
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   29
         Left            =   2160
         TabIndex        =   118
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   8640
         TabIndex        =   117
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   600
         X2              =   600
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   92
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   91
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   90
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   89
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   88
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   0
         TabIndex        =   87
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   0
         TabIndex        =   86
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   0
         TabIndex        =   85
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   0
         TabIndex        =   84
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   720
         TabIndex        =   83
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   720
         TabIndex        =   82
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   720
         TabIndex        =   81
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   720
         TabIndex        =   80
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   720
         TabIndex        =   79
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1320
         X2              =   1320
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   720
         TabIndex        =   78
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   720
         TabIndex        =   77
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   720
         TabIndex        =   76
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   720
         TabIndex        =   75
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   1440
         TabIndex        =   74
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   20
         Left            =   1440
         TabIndex        =   73
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   21
         Left            =   1440
         TabIndex        =   72
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   22
         Left            =   1440
         TabIndex        =   71
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   23
         Left            =   1440
         TabIndex        =   70
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   24
         Left            =   1440
         TabIndex        =   69
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   25
         Left            =   1440
         TabIndex        =   68
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   26
         Left            =   1440
         TabIndex        =   67
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   27
         Left            =   1440
         TabIndex        =   66
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   5520
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   28
         Left            =   2160
         TabIndex        =   65
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   375
      Left            =   6240
      Top             =   10680
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
      Left            =   4800
      Top             =   10680
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
   Begin VB.Data Data7 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10440
      Top             =   0
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "班次量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8520
      TabIndex        =   1
      Top             =   1320
      Width           =   6615
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   5040
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   2160
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   3120
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   4080
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "清除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Data Data8 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6960
      Top             =   10680
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6960
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   6720
      Top             =   10680
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
      Left            =   7680
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
      Left            =   7680
      Top             =   10680
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
      Left            =   6240
      Top             =   10680
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
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
   Begin VB.Label Label22 
      Caption         =   "操作员"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   222
      Top             =   5280
      Width           =   6135
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFC0&
      Caption         =   "键盘选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7080
      TabIndex        =   221
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0E0FF&
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
      Height          =   495
      Left            =   6120
      TabIndex        =   217
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   3840
      TabIndex        =   214
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   3240
      TabIndex        =   213
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   4440
      TabIndex        =   212
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFC0&
      Caption         =   "班组选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   208
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      Caption         =   "当前工序"
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
      Left            =   360
      TabIndex        =   101
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "当前操作"
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
      Left            =   3360
      TabIndex        =   99
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   35
      Left            =   11280
      TabIndex        =   97
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   34
      Left            =   11280
      TabIndex        =   96
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   33
      Left            =   11280
      TabIndex        =   95
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   32
      Left            =   11280
      TabIndex        =   94
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "水洗"
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
      Left            =   13560
      TabIndex        =   93
      Top             =   10560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
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
      Height          =   495
      Left            =   2880
      TabIndex        =   61
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "缸号"
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
      Left            =   360
      TabIndex        =   60
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "排布量"
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
      Left            =   8520
      TabIndex        =   59
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "已报量"
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
      Left            =   8520
      TabIndex        =   58
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "未报量"
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
      Left            =   12120
      TabIndex        =   57
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "班次量"
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
      Left            =   12120
      TabIndex        =   56
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   6960
      TabIndex        =   55
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   6960
      TabIndex        =   54
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   6960
      TabIndex        =   53
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   6960
      TabIndex        =   52
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   6960
      TabIndex        =   51
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   7680
      TabIndex        =   50
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   7680
      TabIndex        =   49
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   7680
      TabIndex        =   48
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   7680
      TabIndex        =   47
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   7680
      TabIndex        =   46
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   8400
      TabIndex        =   45
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   8400
      TabIndex        =   44
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   8400
      TabIndex        =   43
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   8400
      TabIndex        =   42
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   8400
      TabIndex        =   41
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   9120
      TabIndex        =   40
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   9120
      TabIndex        =   39
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   18
      Left            =   9120
      TabIndex        =   38
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   19
      Left            =   9120
      TabIndex        =   37
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   20
      Left            =   9120
      TabIndex        =   36
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   21
      Left            =   9840
      TabIndex        =   35
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   22
      Left            =   9840
      TabIndex        =   34
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   23
      Left            =   9840
      TabIndex        =   33
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   24
      Left            =   9840
      TabIndex        =   32
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   25
      Left            =   9840
      TabIndex        =   31
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   26
      Left            =   10560
      TabIndex        =   30
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   27
      Left            =   10560
      TabIndex        =   29
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   28
      Left            =   10560
      TabIndex        =   28
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   29
      Left            =   10560
      TabIndex        =   27
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   30
      Left            =   10560
      TabIndex        =   26
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "水洗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   31
      Left            =   11280
      TabIndex        =   25
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0E0FF&
      Caption         =   "当前缸号"
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
      Left            =   360
      TabIndex        =   24
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      Caption         =   "工序"
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
      Left            =   4560
      TabIndex        =   23
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "Forms511"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l1, L2, hgxx As String
Dim gybb, gybh As String
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf, sjsx As Integer
Dim dqgx As Integer
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End Sub


Private Sub Command2_Click()
On Error Resume Next
If Text9.Text = "" Then
MsgBox ("请输入操作员信息！")
Exit Sub
End If

If Adodc1.Recordset.EOF Then
Label11.Caption = "请扫描流程卡"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
Else

bz = ""       '班组
yg = Text9.Text          '员工

If Text3.Text <> "" And Len(Text3.Text) = 2 Then

If Val(Text3.Text) >= 1 And Val(Text3.Text) <= 1000 Then       ''''''''''''''''''排布
Adodc3.RecordSource = "select * from PBCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmpb('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If


If Val(Text3.Text) >= 1001 And Val(Text3.Text) <= 6000 Then    '''''''''''''''''染色
Adodc3.RecordSource = "select * from RSCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmrs('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If



If Val(Text3.Text) >= 6001 And Val(Text3.Text) <= 7000 Then    '''''''''''''''''脱水
Adodc3.RecordSource = "select * from TSCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmts('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If


If Val(Text3.Text) >= 7001 And Val(Text3.Text) <= 8000 Then    '''''''''''''''''烘干
Adodc3.RecordSource = "select * from HGCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmhg('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If

If Val(Text3.Text) >= 8001 And Val(Text3.Text) <= 9000 Then    '''''''''''''''''小定型
Adodc3.RecordSource = "select * from XDCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmxd('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If

If Val(Text3.Text) >= 9001 And Val(Text3.Text) <= 9999 Then    '''''''''''''''''大定型
Adodc3.RecordSource = "select * from DDCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmdd('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If

Else
Text2.Text = ""
Text2.SetFocus
Exit Sub
End If


Label11.Caption = "请扫描流程卡"
Text2.Text = ""
Exit Sub
End If


End Sub

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Command4_Click()
  Dim text12Value As String
    Dim parts() As String
    Dim result As String

    ' 假设 Text12 是当前表单中的一个文本框控件
    text12Value = Text12.Text  ' 使用控件的 Text 属性获取其值

    ' 检查 Text12 的值是否为空
    If text12Value = "" Then
        ' 如果为空，直接显示 Forms509
        Forms509.Show
    Else
        parts = Split(text12Value, "/")  ' 使用斜杠分割字符串
        If UBound(parts) >= 1 Then
            result = parts(1)  ' 获取分割后数组的第二个元素
        Else
            result = "No valid data after '/'"  ' 或者设置为合适的默认值或错误消息
        End If

        ' 接下来的代码，例如赋值给 Forms509 的控件等
        ' 确保 Forms509 和 Text1(3) 控件存在
        If Not Forms509 Is Nothing Then
            Forms509.Text1(3).Text = result
            Forms509.Show
        Else
            MsgBox "Forms509 is not loaded or does not exist."
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Label17.Caption = Format(Month(Date), "0#")
Label20.Caption = Format(Month(Date) - 1, "0#")
For i = 1 To 35
Label1(i).Caption = ""
Label1(i).Visible = False
Next

If yhxm <> "" Then
Option3.value = True
Option4.Visible = False
Else
Option4.value = True
Option3.Visible = False
End If

Option2.value = True
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
sjsx = 0
If InStr(yhdm, "2") > 0 Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select 工序其它系数 from gyshd where 工序其它系数<>'0' group by 工序其它系数 order by 工序其它系数"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
i = 1
Adodc7.Recordset.MoveFirst
Do While Not Adodc7.Recordset.EOF
Label13(i).Caption = Adodc7.Recordset.Fields(0)
Label13(i).Visible = True
i = i + 1
Adodc7.Recordset.MoveNext
Loop
End If


Label11.Caption = "请扫描流程卡"
Text2.TabIndex = 0
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 1600
VSFlexGrid1.ColWidth(2) = 1600

ActivateKeyboardLayout 134481924, 1
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
On Error Resume Next
sql2 = "delete from yhcd where 用户='" & yhm & "' and 编号='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.Adodc1.Refresh
End Sub

Private Sub Label1_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
Text8.Text = Mid(Label1(Index).Caption, InStr(Label1(Index).Caption, "-") + 1, InStr(Label1(Index).Caption, "/") - InStr(Label1(Index).Caption, "-") - 1)
Text3.Text = Mid(Label1(Index).Caption, 1, InStr(Label1(Index).Caption, "-") - 1)
l1 = Mid(Label1(Index).Caption, InStr(Label1(Index).Caption, "-") + 1, InStr(Label1(Index).Caption, "/") - InStr(Label1(Index).Caption, "-") - 1)
L2 = Mid(Label1(Index).Caption, InStr(Label1(Index).Caption, "/") + 1)
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select 锅号 from v_lv_sczd where 工序='" & l1 & "' order by 起始"
Adodc11.Refresh
If Not Adodc11.Recordset.EOF And Adodc11.Recordset.Fields(0) = Text1 Then
For i = 0 To List1.ListCount - 1
If InStr(List1.List(i), Text3.Text) > 0 Then
List1.Selected(i) = True
End If
Next
End If
If InStr(yhdm, "2") > 0 Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If
End Select
End Sub


Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label1(Index).BackColor = &H8080FF
End Select
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label1(Index).BackColor = &HC0FFC0
End Select
End Sub

Private Sub Label10_Click()
'If Val(Text3.Text) > 1000 And Val(Text3.Text) < 6000 Then Exit Sub
Text7.Text = ""
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.BackColor = &H8080FF
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.BackColor = &HC0FFC0
End Sub

Private Sub Label13_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
Text9.Text = ""
Text8.Text = ""
Text3.Text = ""
Label3.Caption = Label13(Index).Caption
Text10.Text = Label13(Index).Caption
For i = 1 To 36
Label1(i).Caption = ""
Label1(i).Visible = False
Next
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select * from gyshd where 工序其它系数='" & Label13(Index).Caption & "'  order by 工艺编号"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
i = 1
Adodc7.Recordset.MoveFirst
Do While Not Adodc7.Recordset.EOF
Label1(i).Caption = Adodc7.Recordset.Fields(0) + "-" + Adodc7.Recordset.Fields(1) + "/" + Adodc7.Recordset.Fields(2)
Label1(i).Visible = True
i = i + 1
Adodc7.Recordset.MoveNext
Loop
End If
End Select
End Sub

Private Sub Label13_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label13(Index).BackColor = &H8080FF
End Select
End Sub

Private Sub Label13_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label13(Index).BackColor = &HFFFF00
End Select
End Sub

Private Sub Label15_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
If Option1.value = True Then
Text11.Text = Text11.Text + Label15(Index).Caption
End If
If Option2.value = True Then
Text2.Text = Text2.Text + Label15(Index).Caption
End If
End Select
End Sub

Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label15(Index).BackColor = &H8080FF
End Select
End Sub

Private Sub Label15_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label15(Index).BackColor = &HFFFFC0
End Select
End Sub

Private Sub Label16_Click()
On Error Resume Next
If Option1.value = True Then
Text11.Text = Mid(Text11, 1, Len(Text11) - 1)
End If
If Option2.value = True Then
Text2.Text = Mid(Text2, 1, Len(Text2) - 1)
End If
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label16.BackColor = &H8080FF
End Sub

Private Sub Label16_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label16.BackColor = &HFFFFC0
End Sub

Private Sub Label17_Click()
Text2 = "A" + Mid(Format(Date, "YYYY"), 3) + Label17.Caption
End Sub






Private Sub Label18_Click()
pmbl = 7
Formr440.Show
End Sub

Private Sub Label19_Click()
Forms512.Show
End Sub

Private Sub Label20_Click()
Text2 = "A" + Mid(Format(Date, "YYYY"), 3) + Label20.Caption
End Sub

Private Sub Label23_Click()
Text2 = Text12 + "J"
End Sub

Private Sub Label3_Click()
YGBL = 9
Forms546.Text1(0) = Label3.Caption
Forms546.Show
End Sub

Private Sub Label9_Click(Index As Integer)
Select Case Index
       Case Index
Text7.Text = Text7.Text + Label9(Index).Caption
End Select
End Sub


Private Sub Label9_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label9(Index).BackColor = &H8080FF
End Select
End Sub

Private Sub Label9_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label9(Index).BackColor = &HC0FFC0
End Select
End Sub

Private Sub Text1_Change()
On Error Resume Next ' 如果代码遇到错误，继续执行，不显示错误信息

' 设置 Adodc8 的数据库连接字符串
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

' 设置 Adodc1 的数据库连接字符串并查询 kpd 表中与 Text1.Text 匹配的记录
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 锅号,品名,色别,匹数,重量,车台,ztbh as 生产编号 from kpd where 锅号='" & Text1.Text & "' "
Adodc1.Refresh ' 刷新数据源

' 设置 Adodc5 的数据库连接字符串并查询 ghgx 表，筛选特定工序范围的记录
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select distinct 工序 from ghgx where 锅号='" & Text1.Text & "' and 工序 not between '0080' and '0200'  order by 工序"
Adodc5.Refresh ' 刷新数据源

sjsx = 0 ' 初始化变量 sjsx
List1.Clear ' 清空 List1 列表框
For i = 1 To 35 ' 遍历 1 到 35 号标签
    Label1(i).Caption = "" ' 清空标签的显示内容
    Label1(i).Visible = False ' 隐藏标签
Next

If Adodc5.Recordset.EOF Then ' 如果 Adodc5 结果集为空
    For i = 1 To 35 ' 再次清空所有标签
        Label1(i).Caption = ""
        Label1(i).Visible = False
    Next
Else ' 如果 Adodc5 结果集中有数据
    Adodc5.Recordset.MoveFirst ' 移动到第一条记录
    i = 1 ' 初始化计数器 i
    Do While Not Adodc5.Recordset.EOF ' 遍历所有记录
        ' 设置 Adodc7 的数据库连接字符串并查询 gyshd 表中与工序匹配的记录
        Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
        Adodc7.RecordSource = "select * from gyshd where 工艺编号='" & Adodc5.Recordset.Fields(0) & "'"
        Adodc7.Refresh ' 刷新数据源

        If Not Adodc7.Recordset.EOF Then ' 如果 Adodc7 结果集中有数据
            If InStr(yhdm, "1") > 0 Then ' 如果 yhdm 字符串中包含 '1'
                ' 检查工艺编号是否在特定范围内
                If Adodc7.Recordset.Fields(0) > 0 And Adodc7.Recordset.Fields(0) < 900 Then
                    ' 设置标签的显示内容为工艺编号、名称加上 '1'
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True ' 显示标签

                    ' 查询 pbcl 表中与锅号和工艺编号匹配的记录
                    Adodc8.RecordSource = "select * from pbcl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh ' 刷新数据源

                    If Adodc8.Recordset.EOF Then ' 如果 pbcl 表中没有匹配记录
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1" ' 将记录添加到列表框
                        Label1(i).Enabled = True ' 启用标签
                    Else
                        ' 查询 kpd 表中与锅号和工序匹配的重量
                        Adodc6.RecordSource = "select round(sum(重量),2) from kpd where 锅号='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
                        Adodc6.Refresh ' 刷新数据源

                        ' 查询 pbcl 表中与锅号和工艺编号匹配的班次产量
                        Adodc8.RecordSource = "select round(sum(班次产量),2) from pbcl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh ' 刷新数据源

                        ' 如果 kpd 表中的重量小于或等于 pbcl 表中的班次产量，禁用标签
                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else ' 否则，启用标签并将记录添加到列表框
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1 ' 增加计数器
                End If
            End If

            ' 检查 yhdm 中是否包含 "2"
            If InStr(yhdm, "2") > 0 Then
                ' 检查工艺编号是否在特定范围内
                If Adodc7.Recordset.Fields(0) > 1000 And Adodc7.Recordset.Fields(0) < 6000 Then
                    ' 设置标签内容并查询 rscl 表
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True
                    Adodc8.RecordSource = "select * from rscl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh

                    If Adodc8.Recordset.EOF Then
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                        Label1(i).Enabled = True
                    Else
                        ' 查询 pld 表的数量和 rscl 表的班次产量
                        Adodc6.RecordSource = "select round(sum(数量),2) from pld where 锅号='" & Text1.Text & "' and 信息 like '%正常%'"
                        Adodc6.Refresh
                        Adodc8.RecordSource = "select round(sum(班次产量),2) from rscl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh

                        ' 对比 pld 和 rscl 表中的数量和产量
                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1
                End If
            End If

            ' 检查 yhdm 中是否包含 "3" (例如用于脱水工序)
            If InStr(yhdm, "3") > 0 Then
                If Adodc7.Recordset.Fields(0) > 6001 And Adodc7.Recordset.Fields(0) < 7000 Then
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True
                    Adodc8.RecordSource = "select * from tscl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh

                    If Adodc8.Recordset.EOF Then
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                        Label1(i).Enabled = True
                    Else
                        Adodc6.RecordSource = "select round(sum(重量),2) from kpd where 锅号='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
                        Adodc6.Refresh
                        Adodc8.RecordSource = "select round(sum(班次产量),2) from tscl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh

                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1
                End If
            End If

            ' 检查 yhdm 中是否包含 "4" (用于烘干工序)
            If InStr(yhdm, "4") > 0 Then
                If Adodc7.Recordset.Fields(0) > 7000 And Adodc7.Recordset.Fields(0) < 8000 Then
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True
                    Adodc8.RecordSource = "select * from hgcl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh

                    If Adodc8.Recordset.EOF Then
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                        Label1(i).Enabled = True
                    Else
                        Adodc6.RecordSource = "select round(sum(重量),2) from kpd where 锅号='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
                        Adodc6.Refresh
                        Adodc8.RecordSource = "select round(sum(班次产量),2) from hgcl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh

                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1
                End If
           
            
            If Adodc7.Recordset.Fields(0) > 8000 And Adodc7.Recordset.Fields(0) < 9000 Then
         Label1(i).Caption = Adodc7.Recordset.Fields(0) + "-" + Adodc7.Recordset.Fields(1) + "/" + "1"
         Label1(i).Visible = True

         Adodc8.RecordSource = "select * from xdcl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
         Adodc8.Refresh
         If Adodc8.Recordset.EOF Then
        List1.AddItem Adodc7.Recordset.Fields(0) + "-" + Adodc7.Recordset.Fields(1) + "/" + "1"
        Label1(i).Enabled = True

         Else
        Adodc6.RecordSource = "select round(sum(重量),2) from kpd where 锅号='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
         Adodc6.Refresh
       Adodc8.RecordSource = "select round(sum(班次产量),2) from xdcl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
        Adodc8.Refresh

       If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
        Label1(i).Enabled = False
    Else
       List1.AddItem Adodc7.Recordset.Fields(0) + "-" + Adodc7.Recordset.Fields(1) + "/" + "1"
        Label1(i).Enabled = True

        End If
       End If
       i = i + 1
      End If
      End If

            ' 检查 yhdm 中是否包含 "5" (用于大定型工序)
            If InStr(yhdm, "5") > 0 Then
                If Adodc7.Recordset.Fields(0) > 9000 And Adodc7.Recordset.Fields(0) < 9999 Then
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True
                    Adodc8.RecordSource = "select * from ddcl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh

                    If Adodc8.Recordset.EOF Then
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                        Label1(i).Enabled = True
                    Else
                        Adodc6.RecordSource = "select round(sum(重量),2) from kpd where 锅号='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
                        Adodc6.Refresh
                        Adodc8.RecordSource = "select round(sum(班次产量),2) from ddcl where 锅号='" & Text1.Text & "' and 工艺编号='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh

                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1
                End If
            End If
        End If

        Adodc5.Recordset.MoveNext ' 移动到下一条记录
    Loop
End If

' 检查标签是否启用并处理工艺编号的自动扫描和状态控制
For i = 1 To 35
    If Label1(i).Enabled = True Then
        L = i
        dqgx = L ' 当前工序为 L

        If Option3.value = True Then
            If Mid(List1.List(0), 1, 4) = yhxm And Text12 <> "" Then
                ' 如果工艺编号等于用户姓名，并且班组或员工信息不为空，则自动扫描产量
                Label1_Click (dqgx)
                Text2 = Text12 & "J"
            End If
        End If

        GoTo 100
    End If
Next

100:

' 如果标签的第一个字符是 "12345"，则启用相应的后续标签
If InStr("12345", Left(Label1(L).Caption, 1)) > 0 Then
    For m = L + 1 To 35
        If InStr("12345", Left(Label1(m).Caption, 1)) > 0 Then
            Label1(m).Enabled = True
        Else
            Label1(m).Enabled = True '''''这里改成flase扫完上一个工序才能扫下面的工序
        End If
    Next
Else
    ' 否则，禁用后续标签
    For m = L + 1 To 35
        Label1(m).Enabled = True '''''这里改成flase扫完上一个工序才能扫下面的工序
    Next
End If
End Sub



Private Sub Text2_Change()
On Error Resume Next

If InStr(Text2.Text, "J") > 0 Or InStr(Text2.Text, "j") > 0 Then

m = Mid(Text2.Text, 1, Len(Text2.Text) - 1)

'If Len(M) = 7 And Mid(M, 1, 1) = "8" Then
If (InStr(m, ".") > 0 And Len(m) / 4 = Int(Len(m) / 4)) Or (InStr(m, ".") > 0 And InStr(m, "/") > 0) Then
If Adodc1.Recordset.EOF Then
Label11.Caption = "请扫描流程卡"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
Else

If InStr(m, "/") > 0 Then

bz = Mid(m, 1, InStr(m, "/") - 1)       '班组
'BZ = yhdm      '班组
yg = Mid(m, InStr(m, "/") + 1)          '员工
'yg = M          '员工

Else
'bz = Mid(M, 1, InStr(M, "/") - 1)       '班组
bz = bzdm       '班组
'yg = Mid(M, InStr(M, "/") + 1)          '员工
yg = m          '员工
End If

If Text3.Text <> "" And Len(Text3.Text) = 4 Then

If Val(Text7.Text) <= 0 Then
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
Exit Sub
End If

djsj = Now

For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then

Text8.Text = Mid(List1.List(i), InStr(List1.List(i), "-") + 1, InStr(List1.List(i), "/") - InStr(List1.List(i), "-") - 1)
Text3.Text = Mid(List1.List(i), 1, InStr(List1.List(i), "-") - 1)    ''''''工序编号


Adodc10.RecordSource = "select 工序 from ghgx where 锅号='" & Text1 & "' and 工序>'" & Text3 & "' order by 工序"
Adodc10.Refresh
If Not Adodc10.Recordset.EOF Then
If Val(Adodc10.Recordset.Fields(0)) < 1000 Or Val(Adodc10.Recordset.Fields(0)) > 6000 Then
gybb = Adodc10.Recordset.Fields(0)
sql1 = "update ghgx  set 起始='" & Now & "' where 锅号='" & Text1 & "' and 工序='" & Adodc10.Recordset.Fields(0) & "'"
sql2 = "update ghgx  set 结束='" & Now & "' where 锅号='" & Text1 & "' and 工序='" & Text3 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
Else
gybb = "9999"
sql1 = "update ghgx  set 结束='" & Now & "' where 锅号='" & Text1 & "' and 工序='" & Text3 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

l1 = Mid(List1.List(i), InStr(List1.List(i), "-") + 1, InStr(List1.List(i), "/") - InStr(List1.List(i), "-") - 1)
L2 = Mid(List1.List(i), InStr(List1.List(i), "/") + 1)
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select 工序其它系数,工序工资系数,工序名称 from gyshd where 工艺编号='" & Text3 & "'"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
l1 = Adodc7.Recordset.Fields(2)   ''''工序名称
L2 = Adodc7.Recordset.Fields(1)   ''''工序其它系数
L3 = Adodc7.Recordset.Fields(0)   ''''工序工资系数
Else
l1 = l1
L2 = 0
L3 = ""
End If

If Text8 <> l1 Then
MsgBox ("请确认工序选择 是否正确")
Exit Sub
End If

'If Val(Text3) < 1000 And Val(Adodc10.Recordset.Fields(0)) > 1000 And Val(Adodc10.Recordset.Fields(0)) < 6000 Then
'gybb = Adodc10.Recordset.Fields(0)
'l1 = "染缸待排"
'End If

If Val(Text3.Text) >= 1 And Val(Text3.Text) <= 1000 Then       ''''''''''''''''''排布

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmpb1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"    ' 表示调用哪个存储过程
    g_Cmd.Execute  ' 执行存储过程        锅号               工序编号          工序名称      班组         员工          时间       工序其它系数     未报量                班次量         工序工资系数       机台            工序编号
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If


If Val(Text3.Text) >= 1001 And Val(Text3.Text) <= 6000 Then    '''''''''''''''''染色
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''增加倍数
Adodc9.RecordSource = "select distinct 倍数 from ghgx where 锅号='" & Text1 & "' and 工序='" & Text3 & "'"
Adodc9.Refresh
If Not Adodc9.Recordset.EOF Then
bs = Val(Adodc9.Recordset.Fields(0))
End If

If InStr(l1, "出缸") > 0 Then
Adodc10.RecordSource = "select 工序 from ghgx where 锅号='" & Text1 & "' and 工序>'6000' order by 工序"
Adodc10.Refresh
If Not Adodc10.Recordset.EOF Then
gybb = Adodc10.Recordset.Fields(0)
Else
gybb = Text3
End If
End If

If Val(L2) > 0.2 Then '''''如果工资系数大于2毛钱
L2 = L2 * bs                '''' 装换倍数

If hgxx = "1" Then                 '''''''''''''''''''''没有合缸的执行cjsmrsa
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmrsa('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus

Else    ''''''有合缸的执行cjsmrsc,如果扫码的工序工资少了，是因为合缸了,又把合缸的料单删除了

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmrsc('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & hgxx & "','" & L3 & "','" & Text11 & "','" & gybb & "')"    ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If



Else '''''工资系数小于2毛钱

L2 = L2 * bs                '''' 装换倍数
If hgxx = "1" Then                 '''''''''''''''''''''合缸信息
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmrsb('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
Else
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmrsd('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & hgxx & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If

End If
End If


If Val(Text3.Text) >= 6001 And Val(Text3.Text) <= 7000 Then    '''''''''''''''''脱水
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmts1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If


If Val(Text3.Text) >= 7001 And Val(Text3.Text) <= 8000 Then    '''''''''''''''''烘干
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmhg1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If

If Val(Text3.Text) >= 8001 And Val(Text3.Text) <= 9000 Then    '''''''''''''''''小定型
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmxd1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If

If Val(Text3.Text) >= 9001 And Val(Text3.Text) <= 9999 Then    '''''''''''''''''大定型
If InStr(Text8.Text, "光坯发货") > 0 Then
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 锅号 from jgmx where 锅号='" & Text1 & "'"
Adodc2.Refresh

If Adodc2.Recordset.EOF Then
MsgBox ("没有开具发货单据，不能出库")
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
Exit Sub
Else
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmdd1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If
Else
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjsmdd1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If
End If

End If     'list选取用
Next
Call Text1_Change

Else
Text2.Text = ""
Text2.SetFocus
Exit Sub
End If


Label11.Caption = "请扫描流程卡"
Text2.Text = ""
Exit Sub
End If
End If

If Len(m) > 3 Then
Text2.Text = ""
Text3.Text = ""
' 检查m中是否存在"+", 如果存在，则将其替换为空字符，即去除所有的"+"
If InStr(m, "+") > 0 Then
    Text1.Text = Replace(m, "+", "")
Else
    Text1.Text = m
End If
Label11.Caption = "请选择工序"
End If

End If

End Sub

Private Sub Text3_Change()
On Error Resume Next
If Val(Text3.Text) >= 1 And Val(Text3.Text) <= 1000 Then
Adodc3.RecordSource = "select * from PBCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 1)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "smccjc1"       ' 表示调用哪个存储过程"
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

Text4.Text = Val(g_Cmd.Parameters("pb").value)
Text5.Text = Val(g_Cmd.Parameters("cl").value) '''
Text6.Text = Val(Text4.Text) - Val(Text5.Text)
Text7.Text = Val(Text4.Text) - Val(Text5.Text)

End If

' 如果Text3控件的文本值大于等于1001且小于等于6000，则执行以下操作
If Val(Text3.Text) >= 1001 And Val(Text3.Text) <= 6000 Then

    ' 设置Adodc4控件的数据源为一个SQL查询语句，查询bgxx表中所有配料编号，条件是并缸锅号等于Text1控件的文本值
    Adodc4.RecordSource = "select 配料编号 from bgxx where 并缸锅号='" & Text1.Text & "'"
    ' 刷新Adodc4控件，执行查询
    Adodc4.Refresh
    ' 如果Adodc4的记录集到达了末尾，说明没有找到记录
    If Adodc4.Recordset.EOF Then
        ' 设置hgxx变量的值为"1"
        hgxx = "1"
        ' 设置Adodc3控件的数据源为一个SQL查询语句，查询RSCL表的所有列，条件是锅号等于Text1控件的文本值，工艺编号等于Text3控件的文本值
        Adodc3.RecordSource = "select * from RSCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
        ' 刷新Adodc3控件，执行查询
        Adodc3.Refresh
    Else
        ' 如果Adodc4的记录集没有到达末尾，说明找到了记录
        ' 获取Adodc4记录集的第一个字段值，赋给hgxx变量
        hgxx = Adodc4.Recordset.Fields(0)
        ' 重新设置Adodc3控件的数据源，与上面相同的查询
        Adodc3.RecordSource = "select * from RSCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
        ' 刷新Adodc3控件，执行查询
        Adodc3.Refresh
    End If

    ' 创建一个新的Command对象，用于执行数据库命令
    Set g_Cmd = New Command
    ' 设置数据库连接字符串，包含数据库提供者、密码、用户ID、数据库名、数据源地址
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    ' 将Command对象的活动连接设置为上面的连接字符串
    g_Cmd.ActiveConnection = g_Con

    ' 以下是创建并追加多个参数给Command对象，以供存储过程使用
    ' 创建一个参数，表示锅号，输入类型，长度40，值为Text1的文本值去除两端空格
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text)) '''缸号
    g_Cmd.Parameters.Append param

    ' 创建一个参数，表示工艺编号，输入类型，长度4，值为Text3的文本值去除两端空格
    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text)) ''''工艺编号
    g_Cmd.Parameters.Append param
    
    ' 根据Adodc4是否找到记录，设置不同的条件值给“tj”参数
    ' 如果Adodc4没有找到记录，"tj"参数的值为2；否则为8
    If Adodc4.Recordset.EOF Then
        Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 2)
    Else
        Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 8)
    End If
    g_Cmd.Parameters.Append param

    ' 创建三个输出参数，分别是"pb"、"cl"、"pb1"，它们的类型都是单精度浮点数
    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    ' 设置Command对象的类型为存储过程，并指定存储过程的名称为"smccjc"
    g_Cmd.CommandType = adCmdStoredProc
    g_Cmd.CommandText = "smccjc"
    ' 执行存储过程
    g_Cmd.Execute
    ' 取消Command对象的执行，用于清理
    g_Cmd.Cancel

    ' 从Command对象的参数集合中获取输出参数的值，显示在相应的文本框中
    Text4.Text = Val(g_Cmd.Parameters("pb").value)  ' "pb"参数的值显示在Text4中
    Text5.Text = Val(g_Cmd.Parameters("cl").value)  ' "cl"参数的值显示在Text5中
    ' 计算并显示Text4和Text5的差值在Text6和Text7中
    Text6.Text = Val(Text4.Text) - Val(Text5.Text)
    Text7.Text = Val(Text4.Text) - Val(Text5.Text)
End If


If Val(Text3.Text) >= 6001 And Val(Text3.Text) <= 7000 Then
Adodc3.RecordSource = "select * from TSCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 3)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "smccjc1"       ' 表示调用哪个存储过程"
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

Text4.Text = Val(g_Cmd.Parameters("pb").value)
Text5.Text = Val(g_Cmd.Parameters("cl").value)
Text6.Text = Val(Text4.Text) - Val(Text5.Text)
Text7.Text = Val(Text4.Text) - Val(Text5.Text)
End If


If Val(Text3.Text) >= 7001 And Val(Text3.Text) <= 8000 Then
Adodc3.RecordSource = "select * from HGCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh


Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 4)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "smccjc1"       ' 表示调用哪个存储过程"
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel



Text4.Text = Val(g_Cmd.Parameters("pb").value) ''排布量
Text5.Text = Val(g_Cmd.Parameters("cl").value) ''已报量
Text6.Text = Val(Text4.Text) - Val(Text5.Text) '' 未报量
Text7.Text = Val(Text4.Text) - Val(Text5.Text) ''班次量

End If

If Val(Text3.Text) >= 8001 And Val(Text3.Text) <= 9000 Then
Adodc3.RecordSource = "select * from XDCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 5)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
  
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "smccjc1"       ' 表示调用哪个存储过程"
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

Text4.Text = Val(g_Cmd.Parameters("pb").value) ''排布量
Text5.Text = Val(g_Cmd.Parameters("cl").value) ''已报量
Text6.Text = Val(Text4.Text) - Val(Text5.Text) '' 未报量
Text7.Text = Val(Text4.Text) - Val(Text5.Text) ''班次量
End If

If Val(Text3.Text) >= 9001 And Val(Text3.Text) <= 9999 Then
Adodc3.RecordSource = "select * from DDCL where 锅号='" & Text1.Text & "' and 工艺编号='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 6)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
     Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
   
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "smccjc1"       ' 表示调用哪个存储过程"
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

Text4.Text = Val(g_Cmd.Parameters("pb").value)
Text5.Text = Val(g_Cmd.Parameters("cl").value)
Text6.Text = Val(Text4.Text) - Val(Text5.Text)
Text7.Text = Val(Text4.Text) - Val(Text5.Text)
End If

VSFlexGrid3.ColWidth(0) = 100
VSFlexGrid3.ColWidth(1) = 0
VSFlexGrid3.ColWidth(2) = 0
VSFlexGrid3.ColWidth(6) = 0
VSFlexGrid3.ColWidth(10) = 0
VSFlexGrid3.ColWidth(12) = 0
VSFlexGrid3.ColWidth(13) = 0
VSFlexGrid3.ColWidth(14) = 0
VSFlexGrid3.ColWidth(15) = 0
VSFlexGrid3.ColWidth(16) = 0
VSFlexGrid3.ColWidth(17) = 0
VSFlexGrid3.ColWidth(18) = 0

VSFlexGrid1.ColWidth(0) = 200
Label11.Caption = "请扫描工号条码"

End Sub

Private Sub Timer1_Timer()
If sjsx >= 1 Then
Text12 = bzgrbh
Timer1.Enabled = False
sjsx = 1
Else
sjsx = sjsx + 1
End If
End Sub
