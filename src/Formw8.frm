VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw8 
   BackColor       =   &H00C0E0FF&
   Caption         =   "特种记账"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "应付类别"
      Height          =   495
      Left            =   10080
      TabIndex        =   43
      Top             =   600
      Width           =   3735
      Begin VB.OptionButton Option3 
         Caption         =   "五金"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   45
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "染料"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   44
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF80&
      Caption         =   "应收账款"
      Height          =   495
      Left            =   8520
      TabIndex        =   42
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF80&
      Caption         =   "应付账款"
      Height          =   495
      Left            =   6960
      TabIndex        =   41
      Top             =   600
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   5520
      Top             =   10320
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4080
      TabIndex        =   40
      Text            =   "Text3"
      Top             =   1320
      Width           =   615
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw8.frx":0000
      Height          =   6735
      Left            =   240
      TabIndex        =   39
      Top             =   3360
      Width           =   14535
      _cx             =   25638
      _cy             =   11880
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   6120
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6720
      Top             =   10320
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Left            =   6480
      Top             =   10560
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
      Bindings        =   "Formw8.frx":0015
      Height          =   330
      Index           =   0
      Left            =   1440
      TabIndex        =   27
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      ItemData        =   "Formw8.frx":002A
      Left            =   11880
      List            =   "Formw8.frx":0034
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/M/d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/M/d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      Enabled         =   0   'False
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Enabled         =   0   'False
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   4560
      TabIndex        =   15
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   329711617
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1320
      TabIndex        =   16
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   329711617
      CurrentDate     =   36892
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   1
      Left            =   1440
      TabIndex        =   28
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   2
      Left            =   1440
      TabIndex        =   29
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   3
      Left            =   1440
      TabIndex        =   30
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw8.frx":0044
      Height          =   330
      Index           =   4
      Left            =   4800
      TabIndex        =   31
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   5
      Left            =   4800
      TabIndex        =   32
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   6
      Left            =   4800
      TabIndex        =   33
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   7
      Left            =   4800
      TabIndex        =   34
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   8
      Left            =   8400
      TabIndex        =   35
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   9
      Left            =   8400
      TabIndex        =   36
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   10
      Left            =   8400
      TabIndex        =   37
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   11
      Left            =   8400
      TabIndex        =   38
      Top             =   2760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "类别"
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
      Left            =   240
      TabIndex        =   25
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "打印选择"
      Height          =   375
      Left            =   10800
      TabIndex        =   24
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "摘要"
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
      Left            =   240
      TabIndex        =   22
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "对方科目"
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
      Left            =   3600
      TabIndex        =   19
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3480
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注2"
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
      Index           =   11
      Left            =   6960
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注3"
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
      Index           =   10
      Left            =   6960
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注4"
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
      Index           =   9
      Left            =   6960
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   8
      Left            =   6960
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单据号"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注1"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "贷方金额"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "借方金额"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "Formw8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Private Sub Command1_Click()
On Error Resume Next
If MsgBox("确认删除吗", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT MAX(序号) FROM TZJZMX "
Adodc2.Refresh
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If DataCombo1(1).Text = "" Or DataCombo1(4).Text = "" Or DataCombo1(5).Text = "" Or DataCombo1(6).Text = "" Then Exit Sub
Adodc1.Recordset.AddNew
For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT MAX(序号) FROM TZJZMX "
Adodc2.Refresh
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If DataCombo1(1).Text = "" Or DataCombo1(4).Text = "" Or DataCombo1(5).Text = "" Or DataCombo1(6).Text = "" Then Exit Sub
If MsgBox("确认修改吗", vbYesNo) = vbNo Then Exit Sub

For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT MAX(序号) FROM TZJZMX "
Adodc2.Refresh
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command5_Click()
Call OutadodcToExcel2(VSFlexGrid1, 6, 7, "特种记账 日期范围： " + Text1.Text + "--" + Text2.Text)
If Combo1.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)  ORDER BY 日期,序号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * from TZJZMX WHERE 类别='" & Combo1.Text & "' AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)  ORDER BY 日期,序号"
Adodc1.Refresh
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next

If Combo1.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)  ORDER BY 日期,序号"
Adodc1.Refresh
Else
If Combo1.Text = "应付" Then
If Option3(0).value = True Then
Adodc1.RecordSource = "SELECT * from TZJZMX WHERE  cast(贷方金额 as real)>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) and right(对方科目,len(对方科目)-PATINDEX('%-%',对方科目)) in(select distinct 简称 from gys where ip like '%染料%')  ORDER BY 日期,序号"
Adodc1.Refresh
End If

If Option3(1).value = True Then
Adodc1.RecordSource = "SELECT * from TZJZMX WHERE  cast(贷方金额 as real)>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) and right(对方科目,len(对方科目)-PATINDEX('%-%',对方科目)) in(select distinct 简称 from gys where ip like '%五金%')  ORDER BY 日期,序号"
Adodc1.Refresh
End If

End If

If Combo1.Text = "应收" Then
Adodc1.RecordSource = "SELECT * from TZJZMX WHERE  cast(借方金额 as real)>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)  ORDER BY 日期,序号"
Adodc1.Refresh
End If

End If
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT MAX(序号) FROM TZJZMX "
Adodc2.Refresh
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If

VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 1, 6, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 1, 7, , &HC0C0&

End Sub


Private Sub Command7_Click()
Formw119.DTPicker1 = Text1
Formw119.DTPicker2 = Text2
Formw119.Text1 = Text1
Formw119.Text2 = Text2
Formw119.Show
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 0
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT MAX(序号) FROM TZJZMX "
Adodc2.Refresh
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
End Select
End Sub


Private Sub dataCombo1_LostFocus(Index As Integer)
Select Case Index
       Case 4
       If Option1.value Then
       If InStr(DataCombo1(4), "应付账款") = 0 Then
       DataCombo1(4) = "应付账款" + "-" + DataCombo1(4)
       End If
       End If
       
       If Option2.value Then
       If InStr(DataCombo1(4), "应收账款") = 0 Then
       DataCombo1(4) = "应收账款" + "-" + DataCombo1(4)
       End If
       End If
       
       Case 5
       If DataCombo1(5).Text <> "0" Then
       DataCombo1(6).Text = 0
       End If
       Case 6
       If DataCombo1(6).Text <> "0" Then
       DataCombo1(5).Text = 0
       End If
       
End Select
End Sub

Private Sub DTPicker1_Change()
Text1.Text = DTPicker1.value
Text1.SetFocus
End Sub

Private Sub DTPicker1_CloseUp()
Text1.Text = DTPicker1.value
Text1.SetFocus
End Sub
Private Sub DTPicker2_Change()
Text2.Text = DTPicker2.value
Text2.SetFocus
End Sub

Private Sub DTPicker2_CloseUp()
Text2.Text = DTPicker2.value
Text2.SetFocus
End Sub

Private Sub Form_Load()

On Error Resume Next
Option1.value = True
Text1.Text = Date
Text2.Text = Date
DTPicker1.value = Date
DTPicker2.value = Date
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Text3 = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) ORDER BY 序号 DESC"
Adodc1.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT MC FROM TZZL GROUP BY MC "
Adodc3.Refresh
Combo1.Text = ""
For i = 0 To 11
DataCombo1(i).Text = ""
Next
Option3(0).value = True
DataCombo1(1).Text = Date
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT MAX(序号) FROM TZJZMX "
Adodc2.Refresh
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 1500
VSFlexGrid1.ColWidth(5) = 1500
VSFlexGrid1.ColWidth(8) = 1500
VSFlexGrid1.ColWidth(13) = 0
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

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 8
       DataCombo1(11).Enabled = True
End Select
End Sub
Private Sub Label1_DblClick(Index As Integer)
Select Case Index
       Case 8
       DataCombo1(11).Enabled = False
End Select
End Sub

Private Sub Label3_Click()
KMMC = 9
Formw6.Show
End Sub

Private Sub Label4_Click()
PZZY = 1
Formw1129.Show
End Sub

Private Sub Label6_Click()
KMMC = 14
Formw6.Show
End Sub

Private Sub Text3_Change()
If Option1.value = True Then
If Option3(0).value = True Then
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 简称 from gys where 代码  like '%'+'" & Text3 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' and ip like '%染料%'  group by 简称"
Adodc4.Refresh
End If
If Option3(1).value = True Then
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 简称 from gys where 代码  like '%'+'" & Text3 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' and ip like '%五金%' group by 简称"
Adodc4.Refresh
End If
End If
If Option2.value = True Then
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text3 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'  group by 简称"
Adodc4.Refresh
End If
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
If Adodc1.Recordset.Fields(12) = "已" Or Adodc1.Recordset.Fields(13) = "已" Then
Command1.Enabled = False
Command4.Enabled = False
Exit Sub
Else
For i = 0 To Adodc1.Recordset.Fields.count - 1
DataCombo1(i).Text = Adodc1.Recordset.Fields(i)
Next
Command1.Enabled = True
Command4.Enabled = True
End If
End Sub
