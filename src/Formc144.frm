VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formc144 
   BackColor       =   &H00C0E0FF&
   Caption         =   "光坯入库码单"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   15150
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   29
      Text            =   "Text3"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   28
      Text            =   "Text3"
      Top             =   3240
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   9360
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   4
      Left            =   10680
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   720
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   10200
      Top             =   240
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选择入库"
      Height          =   495
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   3
      Left            =   8040
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   2
      Left            =   5760
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text2 
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
      Height          =   375
      Index           =   1
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text2 
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
      Height          =   375
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "码单打印"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      ItemData        =   "Formc144.frx":0000
      Left            =   9960
      List            =   "Formc144.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   3600
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   855
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选取确认"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询信息"
      Height          =   1095
      Left            =   600
      TabIndex        =   2
      Top             =   4320
      Width           =   4815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "已入"
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "未入"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "取消确认"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   855
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10800
      Top             =   240
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
      Bindings        =   "Formc144.frx":0004
      Height          =   1815
      Left            =   600
      TabIndex        =   11
      Top             =   1320
      Width           =   13215
      _cx             =   23310
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   10800
      Top             =   240
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   11040
      Top             =   240
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
      Bindings        =   "Formc144.frx":0019
      Height          =   1935
      Left            =   600
      TabIndex        =   23
      Top             =   5880
      Width           =   9255
      _cx             =   16325
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
      FormatString    =   $"Formc144.frx":002E
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
      Caption         =   "计划序号"
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
      Left            =   4680
      TabIndex        =   27
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "计划单号"
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
      Left            =   600
      TabIndex        =   26
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "光坯米数"
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
      Index           =   5
      Left            =   9600
      TabIndex        =   25
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "光坯重量"
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
      Index           =   4
      Left            =   6960
      TabIndex        =   22
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "光坯匹数"
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
      Index           =   3
      Left            =   4680
      TabIndex        =   21
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "入库序号"
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
      Left            =   4680
      TabIndex        =   16
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "入库单号"
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
      Left            =   600
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
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
      Index           =   0
      Left            =   600
      TabIndex        =   13
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Formc144"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public gygh As String

Private Sub Command1_Click()
On Error Resume Next
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
ph = Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1)
gpps = gpps + 1
gpzl = gpzl + Val(Mid(List2.List(i), InStr(List2.List(i), "-") + 1))
sql1 = "update bmd set 入库单据='',入库序号='' where 锅号='" & Text1.Text & "' and 匹号='" & ph & "' and 单号='" & Text3(0) & "' and 序号='" & Text3(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
MsgBox ("取消成功！")
Call Command6_Click
End Sub

Private Sub Command11_Click()
Call rkdmd(Adodc3, Text1, Text2(0), Text2(1), Text3(0), Text3(1))
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command3_Click()
Formc34.DataCombo1(10) = Text2(3)
Formc34.DataCombo1(11) = Text2(2)
Formc34.DataCombo1(9) = Text2(4)
Formc34.DataCombo1(10) = Text2(3)
Unload Me
End Sub

Private Sub Command4_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = False
Next
End Sub

Private Sub Command5_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = True
Next
End Sub

Private Sub Command6_Click()
If Option1(1).value = True Then
Adodc2.RecordSource = "select 匹号,光胚重量,码数 from bmd where 锅号='" & Text1 & "' and 单号='" & Text3(0) & "' and 序号='" & Text3(1) & "' and len(isnull(入库单据,''))=0 order by 匹号"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List2.Clear
Do While Not Adodc2.Recordset.EOF
List2.AddItem Trim(Adodc2.Recordset.Fields(0)) + "-" + Trim(Adodc2.Recordset.Fields(1)) + "/" + Trim(Adodc2.Recordset.Fields(2))
Adodc2.Recordset.MoveNext
Loop
End If
If Option1(0).value = True Then
Adodc2.RecordSource = "select 匹号,光胚重量,码数 from bmd where 锅号='" & Text1 & "' and isnull(入库单据,'')='" & Text2(0) & "' and 入库序号='" & Text2(1) & "' and  单号='" & Text3(0) & "' and 序号='" & Text3(1) & "' order by 匹号"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List2.Clear
gpps = 0
gpzl = 0
gpms = 0
Do While Not Adodc2.Recordset.EOF
List2.AddItem Trim(Adodc2.Recordset.Fields(0)) + "-" + Trim(Adodc2.Recordset.Fields(1)) + "/" + Trim(Adodc2.Recordset.Fields(2))
gpps = gpps + 1
gpzl = gpzl + Val(Adodc2.Recordset.Fields(1))
gpms = gpms + Val(Adodc2.Recordset.Fields(2))
Adodc2.Recordset.MoveNext
Loop
Text2(2) = gpps
Text2(3) = gpzl
Text2(4) = gpms
Command1.Enabled = True
End If
End Sub

Private Sub Command7_Click()
gpps = 0
gpzl = 0
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
ph = Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1)
gpps = gpps + 1
gpzl = gpzl + Val(Mid(List2.List(i), InStr(List2.List(i), "-") + 1))
sql1 = "update bmd set 入库单据='" & Text2(0) & "',入库序号='" & Text2(1) & "' where 锅号='" & Text1.Text & "' and 匹号='" & ph & "' and 单号='" & Text3(0) & "' and 序号='" & Text3(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
Text2(2) = gpps
Text2(3) = gpzl
MsgBox ("确认成功！")
Call Command6_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Option1(1).value = True
For i = 1 To 4
Text2(i) = ""
Text3(i) = ""
Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(11) = 1500
VSFlexGrid1.ColWidth(12) = 0
VSFlexGrid1.ColWidth(13) = 0
End Sub
Private Sub Text1_Change()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 日期,锅号,客户名称,色别,色名 as 色号,品名,毛胚幅宽,重量,匹数,技术要求 as 克重,光胚幅宽,备注,IP as 单号序号,KP1 AS 光坯入库,单号,标签 AS 款号,单号,ip as 序号 from kpd  WHERE 锅号='" & Text1.Text & "'"
Adodc1.Refresh
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select distinct 日期,单据,序号,客户,品名,成分,颜色,色号,光坯匹数,光坯数量 from jgmxkf  WHERE 锅号='" & Text1.Text & "'"
Adodc4.Refresh
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc4.Recordset.EOF Then Exit Sub
Adodc4.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc4.Recordset.Move rs - 1
Text2(0) = Adodc4.Recordset.Fields(1)
Text2(1) = Adodc4.Recordset.Fields(2)
Command1.Enabled = False
Command7.Enabled = False
End Sub
