VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw55 
   BackColor       =   &H00C0E0FF&
   Caption         =   "加工客户资料"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "Formw55.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9600
      TabIndex        =   32
      Text            =   "Text2"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "简码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   840
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Left            =   600
      TabIndex        =   29
      Top             =   2280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw55.frx":440A
      Height          =   5415
      Left            =   600
      TabIndex        =   28
      Top             =   4200
      Width           =   13935
      _cx             =   24580
      _cy             =   9551
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
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3240
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "单个"
      Height          =   375
      Left            =   13560
      TabIndex        =   26
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "成批"
      Height          =   375
      Left            =   13560
      TabIndex        =   25
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Index           =   8
      Left            =   3600
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   600
      TabIndex        =   3
      Top             =   3480
      Width           =   2895
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
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   9600
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   7800
      TabIndex        =   6
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Index           =   4
      Left            =   5880
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Index           =   3
      Left            =   7680
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   5880
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3600
      TabIndex        =   0
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
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
      Left            =   600
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9480
      Top             =   10320
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
      Left            =   9960
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
      Left            =   9120
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   9960
      Top             =   10440
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "简称快捷"
      Height          =   495
      Left            =   8520
      TabIndex        =   31
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户简称"
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
      Index           =   8
      Left            =   3600
      TabIndex        =   23
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户代码"
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
      Index           =   7
      Left            =   600
      TabIndex        =   22
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "客 户 资 料 信 息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   4560
      TabIndex        =   21
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "地区号"
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
      Index           =   6
      Left            =   9600
      TabIndex        =   20
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "传真"
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
      Left            =   7800
      TabIndex        =   17
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "联系手机"
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
      Left            =   5880
      TabIndex        =   16
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "联系电话"
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
      Left            =   7680
      TabIndex        =   15
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "联系人"
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
      Left            =   5880
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户地址"
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
      Index           =   1
      Left            =   3600
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户全称"
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
      TabIndex        =   12
      Top             =   1560
      Width           =   2895
   End
End
Attribute VB_Name = "Formw55"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command12_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh

For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = ""
Next
DataCombo1.Text = ""
DataCombo1.SetFocus
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If

   For i = 0 To Adodc1.Recordset.Fields.count - 1
   Adodc1.Recordset.Fields(i) = Text1(i).Text
   Next
Adodc1.Recordset.Update
Adodc1.Refresh

If Adodc1.Recordset.EOF Then
VSFlexGrid1.TextMatrix(0, 0) = "记录号"
Exit Sub
End If
Adodc1.Recordset.MoveLast
VSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To Adodc1.Recordset.RecordCount
VSFlexGrid1.TextMatrix(i, 0) = i
Next
For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = ""
Next
DataCombo1.Text = ""
DataCombo1.SetFocus
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定删除吗?,删除将不能恢复!", vbYesNo) = vbNo Then
Exit Sub
End If

Adodc1.Recordset.Delete
Adodc1.Refresh

For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = ""
Next
DataCombo1.Text = ""
DataCombo1.SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub


Private Sub Command6_Click()
'On Error Resume Next
If Option1.value = True Then
If MsgBox("确定成批导入吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CWMC WHERE 科目编号 like '1131%' AND LEN(科目编号)>4"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.RecordSource = "CWMC"
Adodc3.Refresh
Adodc1.Refresh
If Adodc1.Recordset.EOF Then Exit Sub
i = 1
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
Adodc3.Recordset.AddNew
If Len(i) = 1 Then
Adodc3.Recordset.Fields(0) = "1131" + "00" + Trim(i)
End If
If Len(i) = 2 Then
Adodc3.Recordset.Fields(0) = "1131" + "0" + Trim(i)
End If
If Len(i) = 3 Then
Adodc3.Recordset.Fields(0) = "1131" + Trim(i)
End If
Adodc3.Recordset.Fields(1) = Adodc1.Recordset.Fields("简称")
Adodc3.Recordset.Fields(2) = "资产"
Adodc3.Recordset.Fields(3) = "借"
Adodc3.Recordset.Fields(4) = "2"
Adodc3.Recordset.Fields(5) = "是"
Adodc3.Recordset.Update
i = i + 1
Adodc1.Recordset.MoveNext
Loop
End If

If Option2.value = True Then
If MsgBox(Text1(8).Text + "   确定导入吗？", vbYesNo) = vbNo Then Exit Sub
Adodc3.RecordSource = "SELECT MAX(right(科目编号,len(科目编号)-1)) FROM CWMC WHERE  科目编号 like '1131%' AND LEN(科目编号)>4"
Adodc3.Refresh
MC = 0
If Not Adodc3.Recordset.EOF Then
MC = Adodc3.Recordset.Fields(0)
End If
If MC = Null Then MC = 0
MC = Val(MC)
Adodc4.Recordset.AddNew
If Len(MC) = 1 Then
Adodc4.Recordset.Fields(0) = "1131" + "00" + Trim(MC + 1)
End If
If Len(MC) = 2 Then
Adodc4.Recordset.Fields(0) = "1131" + "0" + Trim(MC + 1)
End If
If Len(MC) = 3 Then
Adodc4.Recordset.Fields(0) = "1131" + Trim(MC + 1)
End If
Adodc4.Recordset.Fields(1) = Adodc1.Recordset.Fields("简称")
Adodc4.Recordset.Fields(2) = "资产"
Adodc4.Recordset.Fields(3) = "借"
Adodc4.Recordset.Fields(4) = "2"
Adodc4.Recordset.Fields(5) = "是"
Adodc4.Recordset.Update
End If
MsgBox ("导入成功！")
End Sub


Private Sub Command8_Click()
On Error Resume Next
If MsgBox("重新刷新简码吗？", vbYesNo) = vbNo Then Exit Sub
Adodc1.RecordSource = "select * from KHZL where ip like '%'+'" & yhxx & "'+'%' order by 简称"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF

    If Adodc1.Recordset.Fields(8) = "" Then Adodc1.Recordset.Fields(7) = ""
     Dim a As Integer
     Adodc1.Recordset.Fields(7) = ""
     a = Len(Adodc1.Recordset.Fields(8))
     For i = 1 To a
         Adodc1.Recordset.Fields(7) = Adodc1.Recordset.Fields(7) & py(Mid(Adodc1.Recordset.Fields(8), i, 1))
     Next
Adodc1.Recordset.Update
Adodc1.Recordset.MoveNext
Loop
MsgBox ("刷新成功！")
Adodc1.Refresh
End Sub



Private Sub Command5_Click()
Adodc1.RecordSource = "select * from KHZL where ip like '%'+'" & yhxx & "'+'%' order by 简称"
Adodc1.Refresh

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

End Sub

Private Sub Command7_Click()
Call PCOutadodcToExcel(VSFlexGrid1)
End Sub

Private Sub DataCombo1_Change()
Text1(0).Text = DataCombo1.Text
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Text1(0).Text = DataCombo1.Text
End Sub

Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Form_Load()

On Error Resume Next

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from KHZL where ip like '%'+'" & yhxx & "'+'%' order by 简称"
Adodc1.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT * FROM CWMC"
Adodc4.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Text2 = ""
DataCombo1.Text = ""
DataCombo1.TabIndex = 0
VSFlexGrid1.ColWidth(1) = 2600

VSFlexGrid1.ColWidth(2) = 2600
VSFlexGrid1.ColWidth(4) = 1200
VSFlexGrid1.ColWidth(5) = 1200
VSFlexGrid1.ColWidth(6) = 1200

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

End Sub

Private Sub Label1_DblClick(Index As Integer)
Select Case Index
       Case 6
       Text2.Enabled = True
End Select
End Sub


Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 8
           If Text1(8).Text = "" Then Text1(7).Text = ""
     Dim a As Integer
     Text1(7).Text = ""
     a = Len(Text1(8).Text)
     For i = 1 To a
         Text1(7).Text = Text1(7).Text & py(Mid(Text1(8).Text, i, 1))
     Next
End Select
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To Adodc1.Recordset.Fields.count - 1
If i = 6 Then i = 7
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
DataCombo1.Text = Text1(0).Text
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text1_LostFocus(Index As Integer)
On Error Resume Next
         
         
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then Exit Sub
Adodc1.RecordSource = "select * from KHZL where 代码  like '%'+'" & Text2 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' order by 简称"
Adodc1.Refresh

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub
