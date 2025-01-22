VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formh123 
   BackColor       =   &H00C0E0FF&
   Caption         =   "色号报价"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15555
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15555
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   370
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   370
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   380
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data4 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data5 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   370
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1340
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "编号"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色号"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   5520
      Top             =   10080
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
      Left            =   5520
      Top             =   10200
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
      Height          =   330
      Left            =   5760
      Top             =   10320
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
      Left            =   5760
      Top             =   9720
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
      Left            =   5400
      Top             =   9960
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formh123.frx":0000
      Height          =   5055
      Left            =   4440
      TabIndex        =   9
      Top             =   240
      Width           =   18975
      _cx             =   33470
      _cy             =   8916
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
      AllowUserResizing=   4
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
      Height          =   330
      Index           =   0
      Left            =   1560
      TabIndex        =   10
      Top             =   2280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   424083457
      CurrentDate     =   39961
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   424083457
      CurrentDate     =   39961
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formh123.frx":0015
      Height          =   330
      Index           =   1
      Left            =   1560
      TabIndex        =   14
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   4
      Left            =   1560
      TabIndex        =   15
      Top             =   1800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formh123.frx":002A
      Height          =   7575
      Left            =   480
      TabIndex        =   16
      Top             =   6000
      Width           =   23055
      _cx             =   40666
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
      AllowUserResizing=   4
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
      Begin MSAdodcLib.Adodc Adodc6 
         Height          =   375
         Left            =   7800
         Top             =   3240
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
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
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
      Left            =   360
      TabIndex        =   21
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
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
      Left            =   360
      TabIndex        =   20
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号"
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
      Left            =   360
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户名称"
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
      Left            =   360
      TabIndex        =   18
      Top             =   1320
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
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "Formh123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sz(6) As String: Dim ZS(6) As String
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim bh As String

Private Sub Command1_Click()
Call OutadodcToExcel(VSFlexGrid1, 10, "色号报价 日期范围： " + Trim(DTPicker1.value) + "--" + Trim(DTPicker2.value))
End Sub

Private Sub Command2_Click()
Call shbj(Adodc1, Adodc6, DataCombo1(0))
End Sub

Private Sub Command3_Click()
On Error Resume Next
sql1 = ""


If Check2(0).value = 1 Then
sql1 = sql1 + "rq between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "kh like '%'+'" & DataCombo1(1) & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "dh like '%'+'" & DataCombo1(0) & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "sh like '%'+'" & DataCombo1(4) & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc3.RecordSource = "SELECT RQ AS 日期,KH AS 客户,fz AS 工艺负责人,YS AS 颜色,SH AS 色号,BL AS 品名,DH AS 配方编号,IP AS 浴比,染助,预定,工时,中定,烘干,硅油,成定,税后 FROM v_zh2_bj WHERE (" + sql1 + ") ORDER BY RQ DESC"
Adodc3.Refresh
End Sub

Private Sub Command7_Click()
Unload Me
End Sub


Private Sub Form_Load()
DTPicker1.value = Date - 30
DTPicker2.value = Date

DataCombo1(0).Text = ""
DataCombo1(1).Text = ""
DataCombo1(4).Text = ""

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from khZL  group by 简称"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT RQ AS 日期,KH AS 客户,fz AS 工艺负责人,YS AS 颜色,SH AS 色号,BL AS 品名,DH AS 配方编号,IP AS 浴比,染助,预定,工时,中定,烘干,硅油,成定,税后 FROM v_zh2_bj where RQ BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) ORDER BY RQ DESC"
Adodc3.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid2.ColWidth(0) = 200
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc3.Recordset.Move rs - 1
DataCombo1(0).Text = Adodc3.Recordset.Fields(6)
DataCombo1(1).Text = Adodc3.Recordset.Fields(1)
DataCombo1(4).Text = Adodc3.Recordset.Fields(4)


Call dpf(DataCombo1(0).Text)

Adodc1.RecordSource = "SELECT * FROM v_dpfd_sh_bj WHERE 配方编号='" & DataCombo1(0).Text & "'ORDER BY 工序名称,次序号"
Adodc1.Refresh


VSFlexGrid2.ColFormat(9) = "#0.#####"

End Sub

Private Sub dpf(bh As String)
    ' 定义sql1语句，删除dpfd1表中配方编号为bh的记录
    sql1 = "delete from dpfd1 where 配方编号='" & bh & "'"
    ' 执行sql1语句
    RD.Open sql1, conn, adOpenStatic, adLockOptimistic

    ' 设置Adodc5的数据源为dpfd表中编号为bh的记录
    Adodc5.RecordSource = "select * from pfd where 编号='" & bh & "'"   '''''如果是从确认配方中色号报价，改成dpfd
    ' 刷新Adodc5的数据
    Adodc5.Refresh

    ' 如果记录集不为空
    If Not Adodc5.Recordset.EOF Then
        ' 移动到记录集的第一条记录
        Adodc5.Recordset.MoveFirst
        ' 初始化变量mb
        mb = 0

        ' 将前7个字段值赋给数组ZS
        For i = 0 To 6
            ZS(i) = Adodc5.Recordset.Fields(i)
        Next

        ' 遍历从第8个字段到第57个字段
        For i = 7 To 56
            ' 如果字段不为空，mb加1
            If Adodc5.Recordset.Fields(i) <> "" Then
                mb = mb + 1
            End If
        Next

        ' 显示进度条
        ProgressBar1.Visible = True
        ' 遍历从第8个字段到mb+8个字段
        For i = 7 To mb + 7
            ' 如果字段不为空
            If Adodc5.Recordset.Fields(i) <> "" Then
                ' 按特定格式解析字段内容并赋值给数组sz
                sz(0) = Mid(Adodc5.Recordset.Fields(i), 1, InStr(Adodc5.Recordset.Fields(i), "(") - 1)
                sz(1) = Mid(Adodc5.Recordset.Fields(i), InStr(Adodc5.Recordset.Fields(i), "(") + 1, InStr(Adodc5.Recordset.Fields(i), ")") - InStr(Adodc5.Recordset.Fields(i), "(") - 1)
                sz(2) = Mid(Adodc5.Recordset.Fields(i), InStr(Adodc5.Recordset.Fields(i), ")") + 1, InStr(Adodc5.Recordset.Fields(i), "-") - InStr(Adodc5.Recordset.Fields(i), ")") - 1)
                sz(3) = Mid(Adodc5.Recordset.Fields(i), InStr(Adodc5.Recordset.Fields(i), "-") + 1, InStr(Adodc5.Recordset.Fields(i), "\") - InStr(Adodc5.Recordset.Fields(i), "-") - 1)
                sz(4) = Mid(Adodc5.Recordset.Fields(i), InStr(Adodc5.Recordset.Fields(i), "\") + 1, InStr(Adodc5.Recordset.Fields(i), "#") - InStr(Adodc5.Recordset.Fields(i), "\") - 1)
                sz(5) = Mid(Adodc5.Recordset.Fields(i), InStr(Adodc5.Recordset.Fields(i), "#") + 1, InStr(Adodc5.Recordset.Fields(i), "^") - InStr(Adodc5.Recordset.Fields(i), "#") - 1)
                sz(6) = Mid(Adodc5.Recordset.Fields(i), InStr(Adodc5.Recordset.Fields(i), "^") + 1)

                ' 计算L的值，L为次序号
                L = i - 6

                ' 检查浴比和车速是否为空
                If sz(1) = "" Then sz(1) = "NULL"
               If sz(6) = "" Then sz(6) = "NULL"
                
                ' 定义sql2语句，插入解析后的数据到dpfd1表
                sql2 = "insert into dpfd1(加工单位,品名,色号,颜色,配方编号,负责人,配方日期,工序名称,浴比,染化助库,染化助名称,单位,配方,车速,次序号,校正值) VALUES('" & ZS(0) & "','" & ZS(1) & "','" & ZS(2) & "','" & ZS(3) & "','" & ZS(4) & "','" & ZS(5) & "','" & ZS(6) & "','" & sz(0) & "'," & _
                        IIf(sz(1) = "NULL", "NULL", "'" & sz(1) & "'") & ",'" & sz(2) & "','" & sz(3) & "','" & sz(4) & "','" & sz(5) & "'," & _
                       IIf(sz(6) = "NULL", "NULL", "'" & sz(6) & "'") & "," & L & ",1)"
                ' 执行sql2语句
                RD.Open sql2, conn, adOpenStatic, adLockOptimistic
                ' 更新进度条的值
                ProgressBar1.value = 100 / mb * L
            End If
        Next
        ' 隐藏进度条
        ProgressBar1.Visible = False
    End If
End Sub





Private Sub VSFlexGrid2_DblClick()
rs = VSFlexGrid2.Row
cl = VSFlexGrid2.col
If cl = 7 Then
Formr29.DataCombo1 = VSFlexGrid2.TextMatrix(rs, 7)
Formr29.Text1(0) = VSFlexGrid2.TextMatrix(rs, 8)
Formr29.Show
End If
End Sub




