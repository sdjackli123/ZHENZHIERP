VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forms510 
   BackColor       =   &H00C0E0FF&
   Caption         =   "员工考勤"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   7560
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   8280
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8640
      Top             =   10560
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
      Height          =   375
      Left            =   8160
      Top             =   10560
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
      Height          =   330
      Left            =   8160
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms510.frx":0000
      Height          =   7095
      Left            =   240
      TabIndex        =   44
      Top             =   2640
      Width           =   14655
      _cx             =   25850
      _cy             =   12515
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
      Bindings        =   "Forms510.frx":0015
      Height          =   330
      Left            =   1320
      TabIndex        =   43
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "工种"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8280
      Top             =   10560
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
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   9720
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   120
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   600
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1080
      Width           =   1095
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1080
      Width           =   1095
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   8520
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   5520
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入考勤"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "考勤刷新"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   10920
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   6480
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   855
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "编号查询"
      Height          =   855
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   10920
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "汇总查询"
      Height          =   855
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导出"
      Height          =   855
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "班次查询"
      Height          =   855
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   12120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认"
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
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1200
      TabIndex        =   25
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   331087873
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   8640
      TabIndex        =   26
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   331087873
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8640
      TabIndex        =   27
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   331087873
      CurrentDate     =   36892
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "车台"
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
      Left            =   9720
      TabIndex        =   42
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "系数"
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
      Left            =   4200
      TabIndex        =   41
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "生产系数"
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
      Left            =   5520
      TabIndex        =   40
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "姓名:"
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
      Left            =   1440
      TabIndex        =   39
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "编号:"
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
      Left            =   240
      TabIndex        =   38
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "工种:"
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
      Left            =   2880
      TabIndex        =   37
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "班次:"
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
      Left            =   8520
      TabIndex        =   36
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
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
      Left            =   6840
      TabIndex        =   35
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "工种"
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
      Left            =   240
      TabIndex        =   34
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "总系数"
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
      Left            =   10920
      TabIndex        =   33
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
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
      Left            =   240
      TabIndex        =   32
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "日期范围"
      Height          =   855
      Left            =   7920
      TabIndex        =   31
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "编号"
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
      Left            =   10920
      TabIndex        =   30
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "班次"
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
      Left            =   12120
      TabIndex        =   29
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "校正系数"
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
      Left            =   3480
      TabIndex        =   28
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Forms510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Command10_Click()
'On Error Resume Next
Adodc1.RecordSource = "SELECT 工种,编号,姓名,sum(cast(出勤时间 as int)) as 出勤,sum(cast(系数时间 as int)) as 系数 FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "'  group BY 工种,编号,姓名 order by 工种,编号"
Adodc1.Refresh
End Sub

Private Sub Command11_Click()
Call kqbb(VSFlexGrid1, 4, "考勤信息")
End Sub

Private Sub Command12_Click()
If Text2.Text = "" Then
If DataCombo1.Text = "" Then
If Text3.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "'  ORDER BY 日期,工种,编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' and 班次='" & Text3.Text & "' ORDER BY 日期,工种,编号"
Adodc1.Refresh
End If

Else

If Text3.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' AND 工种='" & DataCombo1.Text & "'  ORDER BY 日期,工种,编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' and 班次='" & Text3.Text & "'  AND 工种='" & DataCombo1.Text & "' ORDER BY 日期,工种,编号"
Adodc1.Refresh
End If
End If

Else

If DataCombo1.Text = "" Then
If Text3.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' AND 编号='" & Text2.Text & "'  ORDER BY 日期,工种,编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' and 班次='" & Text3.Text & "' AND 编号='" & Text2.Text & "' ORDER BY 日期,工种,编号"
Adodc1.Refresh
End If

Else

If Text3.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' AND 工种='" & DataCombo1.Text & "' AND 编号='" & Text2.Text & "' ORDER BY 日期,工种,编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' and 班次='" & Text3.Text & "'  AND 工种='" & DataCombo1.Text & "' AND 编号='" & Text2.Text & "' ORDER BY 日期,工种,编号"
Adodc1.Refresh
End If
End If
End If


End Sub

Private Sub Command13_Click()
If Text4.Text = "" Then
MsgBox ("请输入系数")
Exit Sub
End If
If Text1(0).Text = "" Then
MsgBox ("请输入编号")
Exit Sub
End If

If MsgBox("要校正的员工编号：" + Text1(0).Text + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("要校正的日期：" + Trim(DTPicker1.value) + "--" + Trim(DTPicker4.value) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "update sjkq set 系数='" & Text4.Text & "' where 编号='" & Text1(0).Text & "' and 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "'"
sql2 = "update sjkq set 系数时间=cast(系数 as int)*cast(出勤时间 as int) where  编号='" & Text1(0).Text & "' and 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("校正成功！")
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确认修改吗", vbYesNo) = vbNo Then Exit Sub

For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh

End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确认删除吗", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
End Sub

Private Sub Command5_Click()
If MsgBox("确定导入吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete * from sjkq where 日期='" & DTPicker3.value & "'"
sql2 = "insert into sjkq(编号,姓名,工种,系数,出勤时间,日期,班次) select 编号,姓名,工种,核算系数1,核算系数1,'" & DTPicker3.value & "',班次 from works1"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 工种='" & DataCombo1.Text & "' and 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
End If
End Sub


Private Sub Command7_Click()
sql1 = "update sjkq set 系数时间=cast(系数 as int)*cast(出勤时间 as int) where 日期='" & DTPicker3.value & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 工种='" & DataCombo1.Text & "' and 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
End If
End Sub

Private Sub Command8_Click()
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' ORDER BY 日期,工种,编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' and 工种='" & DataCombo1.Text & "' ORDER BY 日期,工种,编号"
Adodc1.Refresh
End If
End Sub

Private Sub Command9_Click()
If Text2.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "'  ORDER BY 日期,工种,编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期 between '" & DTPicker1.value & "' and '" & DTPicker4.value & "' and 编号='" & Text2.Text & "' ORDER BY 日期,工种,编号"
Adodc1.Refresh
End If
End Sub

Private Sub DataCombo1_Change()
If DataCombo1.Text = "" Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
Else
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 工种='" & DataCombo1.Text & "' and 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
End If
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 工种='" & DataCombo1.Text & "' and 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
End If
End Sub

Private Sub DTPicker3_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
End Sub

Private Sub DTPicker3_CloseUp()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
End Sub

Private Sub Form_Load()
DataCombo1.Text = ""
DTPicker3.value = Date
For i = 0 To 9
Text1(i).Text = ""
Next
Text1(5).Text = Date
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
DTPicker1.value = Date
DTPicker4.value = Date

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM sjkq WHERE 日期='" & DTPicker3.value & "' ORDER BY 工种,编号"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT * FROM WORKS1"
Adodc2.Refresh
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT 工种 FROM WORKS1 group by 工种"
Adodc4.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
For i = 1 To 3
VSFlexGrid1.ColWidth(i) = 1500
Next
For i = 4 To 5
VSFlexGrid1.ColWidth(i) = 1500
Next
VSFlexGrid1.ColWidth(6) = 1500
VSFlexGrid1.ColWidth(7) = 1500
End Sub

Private Sub vSFlexGrid1_dbClickl()
On Error Resume Next
rs = VSFlexGrid1.Row
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To 9
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To 9
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT * FROM WORKS1 where 编号='" & Text1(0).Text & "'"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
Else
Text1(1).Text = Adodc2.Recordset.Fields(2)
Text1(2).Text = Adodc2.Recordset.Fields(5)
Text1(3).Text = Adodc2.Recordset.Fields(12)
End If
End Select
End Sub

