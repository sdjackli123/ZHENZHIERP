VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw1140 
   BackColor       =   &H00C0E0FF&
   Caption         =   "总类账"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form40"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   4800
      TabIndex        =   27
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   375
      Left            =   7680
      Top             =   8760
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw1140.frx":0000
      Height          =   4095
      Left            =   360
      TabIndex        =   26
      Top             =   3720
      Width           =   14895
      _cx             =   26273
      _cy             =   7223
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
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1920
      TabIndex        =   25
      Text            =   "Text4"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细打印"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "期末结转"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   9840
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按总账汇总"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细账"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "总明细账"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按明细汇总"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "余额打印"
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw1140.frx":0015
      Height          =   1695
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   8421631
      BackColorBkg    =   34952
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   209321985
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   9720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   209321985
      CurrentDate     =   39883
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw1140.frx":0029
      Height          =   330
      Left            =   6120
      TabIndex        =   15
      Top             =   1200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "科目名称"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   9840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   209321985
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   209321985
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker6 
      Height          =   375
      Left            =   2280
      TabIndex        =   22
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   209321985
      CurrentDate     =   39883
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8520
      Top             =   9480
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
      Left            =   8400
      Top             =   9120
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
      Left            =   8040
      Top             =   9600
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
      Left            =   8040
      Top             =   9240
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
      Left            =   8280
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
      Left            =   7800
      Top             =   9720
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
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   24
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   23
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "总账科目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   13
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作月份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "上期结转日"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   9600
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Formw1140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2, BT As String: Public c, r As Integer


Private Sub Command1_Click()
Adodc4.Refresh
Adodc1.RecordSource = "SELECT * FROM ZFLZ WHERE ZFLZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY ZFLZ.日期,ZFLZ.凭证号"
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
Adodc1.Database.Execute "DELETE * FROM ZLCX"
Adodc1.Database.Execute "INSERT INTO ZLCX SELECT * FROM ZFLZ WHERE ZFLZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Adodc1.Database.Execute "UPDATE ZLCX set 序号='2'"
Adodc1.Database.Execute "INSERT INTO ZLCX(会计科目,借方金额,贷方金额) SELECT ZFLZ.会计科目,FORMAT(SUM(VAL(ZFLZ.借方金额)),'#0.00'),FORMAT(SUM(VAL(ZFLZ.贷方金额)),'#0.00') FROM ZFLZ WHERE  ZFLZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') GROUP BY ZFLZ.会计科目"
Adodc1.Database.Execute "UPDATE ZLCX set 摘要='本月合计',序号='3' WHERE 摘要=NULL"

Adodc3.RecordSource = "ZLCX"
Adodc3.Refresh

Adodc1.Database.Execute "INSERT INTO ZLCX SELECT * FROM PMZJZ WHERE 日期='" & DTPicker1.Value & "' "
Adodc1.Database.Execute "UPDATE ZLCX set 序号='1' WHERE 摘要='期初余额'"

Adodc2.RecordSource = "SELECT ZLCX.会计科目 FROM ZLCX GROUP BY ZLCX.会计科目"
Adodc2.Refresh

If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
L = Right(Adodc2.Recordset.Fields(0), Len(Adodc2.Recordset.Fields(0)) - InStr(Adodc2.Recordset.Fields(0), "-"))

Adodc4.Recordset.FindFirst "科目名称='" & L & "' AND LEN(科目编号)=4"
If Adodc4.Recordset.NoMatch Then
MsgBox (L + "科目设置有错")
Exit Sub
End If

Adodc3.RecordSource = "SELECT * FROM ZLCX WHERE INSTR(ZLCX.会计科目,'" & L & "')>0"
Adodc3.Refresh
Adodc3.Recordset.FindFirst "序号='1'"
If Adodc3.Recordset.NoMatch Then
Adodc3.RecordSource = "SELECT * FROM ZLCX WHERE ZLCX.摘要='本月合计' AND ZLCX.会计科目='" & Adodc2.Recordset.Fields(0) & "'"
Adodc3.Refresh
Adodc3.Recordset.Edit
If Adodc4.Recordset.Fields(3) = "贷" Then
Adodc3.Recordset.Fields(7) = Format(Format(Val(Adodc3.Recordset.Fields(5)), "#0.00") - Format(Val(Adodc3.Recordset.Fields(4)), "#0.00"), "#0.00")
Adodc3.Recordset.Fields(6) = "贷"
Else
Adodc3.Recordset.Fields(7) = Format(Format(Val(Adodc3.Recordset.Fields(4)), "#0.00") - Format(Val(Adodc3.Recordset.Fields(5)), "#0.00"), "#0.00")
Adodc3.Recordset.Fields(6) = "借"
End If
Adodc3.Recordset.Update

Else

KKK = Format(Adodc3.Recordset.Fields(7), "#0.00")
Adodc3.RecordSource = "SELECT * FROM ZLCX WHERE ZLCX.摘要='本月合计' AND ZLCX.会计科目='" & Adodc2.Recordset.Fields(0) & "'"
Adodc3.Refresh
Adodc3.Recordset.Edit
If Adodc4.Recordset.Fields(3) = "贷" Then
Adodc3.Recordset.Fields(7) = Format(Format(Val(Adodc3.Recordset.Fields(5)), "#0.00") - Format(Val(Adodc3.Recordset.Fields(4)), "#0.00") + KKK, "#0.00")
Adodc3.Recordset.Fields(6) = "贷"
Else
Adodc3.Recordset.Fields(7) = Format(Format(Val(Adodc3.Recordset.Fields(4)), "#0.00") - Format(Val(Adodc3.Recordset.Fields(5)), "#0.00") + KKK, "#0.00")
Adodc3.Recordset.Fields(6) = "借"
End If
Adodc3.Recordset.Update
End If
Adodc2.Recordset.MoveNext
Loop




Adodc1.Database.Execute "UPDATE ZLCX SET 凭证号='结-'+'" & Text3.Text & "' WHERE 凭证号=NULL"
Adodc1.RecordSource = "SELECT * FROM ZLCX ORDER BY ZLCX.会计科目,VAL(ZLCX.序号),ZLCX.日期"
Adodc1.Refresh
BT = "按明细账汇总"
End Sub

Private Sub Command3_Click()
Call ZYEBDOutadodcToExcelSZ(Adodc2, Adodc3, Text3.Text)
End Sub

Private Sub Command4_Click()
Adodc4.Refresh
Adodc1.RecordSource = "SELECT * FROM ZFLZ WHERE INSTR(ZFLZ.会计科目,'" & DataCombo2.Text & "')>0 AND  ZFLZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY ZFLZ.日期,ZFLZ.凭证号"
Adodc1.Refresh
End Sub



Private Sub Command5_Click()
If MsgBox("确定本期期末余额结转吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("本期为: " + Text3.Text + " 期间" + "，正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("结转为次月的: " + Str(DTPicker2.Value) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
Adodc3.RecordSource = "SELECT * FROM PMZJZ WHERE 日期='" & DTPicker2.Value & "'"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
If MsgBox("已有此时间内的记录，即已结转过，覆盖原先记录吗？", vbYesNo) = vbNo Then
Exit Sub
Else
Adodc1.Database.Execute "DELETE * FROM PMZJZ WHERE 日期='" & DTPicker2.Value & "'"
End If
End If
Adodc1.Database.Execute "INSERT INTO PMZJZ(凭证号,摘要,会计科目,借方金额,贷方金额,借贷方向,余额,类别,序号) SELECT 凭证号,摘要,会计科目,借方金额,贷方金额,借贷方向,余额,类别,序号 FROM ZLCX WHERE 序号='3'"
Adodc1.Database.Execute "UPDATE PMZJZ SET 摘要='期初余额',日期='" & DTPicker2.Value & "' WHERE 日期=null"
MsgBox ("结转成功！")
End Sub

Private Sub Command6_Click()
Adodc4.Refresh
On Error Resume Next
Adodc1.Database.Execute "DELETE * FROM ZLCX"
Adodc1.Database.Execute "INSERT INTO ZLCX(会计科目,借方金额,贷方金额) SELECT ZFLZ.会计科目,FORMAT(SUM(VAL(ZFLZ.借方金额)),'#0.00'),FORMAT(SUM(VAL(ZFLZ.贷方金额)),'#0.00') FROM ZFLZ WHERE  ZFLZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') GROUP BY ZFLZ.会计科目"
Adodc3.RecordSource = "ZLCX"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
L = Adodc3.Recordset.Fields(3)
m = 4

Adodc4.Recordset.FindFirst "科目名称='" & L & "' AND LEN(科目编号)='" & m & "'"
If Adodc4.Recordset.NoMatch Then
MsgBox (L + "科目设置有错")
Exit Sub
Else
If Adodc4.Recordset.Fields(3) = "借" Then
Adodc3.Recordset.Edit
Adodc3.Recordset.Fields(0) = Text2.Text
Adodc3.Recordset.Fields(1) = "汇总"
Adodc3.Recordset.Fields(2) = "本期发生额"
Adodc3.Recordset.Fields(6) = "借"
Adodc3.Recordset.Fields(9) = "2"
Adodc3.Recordset.Update
End If
If Adodc4.Recordset.Fields(3) = "贷" Then
Adodc3.Recordset.Edit
Adodc3.Recordset.Fields(0) = Text2.Text
Adodc3.Recordset.Fields(1) = "汇总"
Adodc3.Recordset.Fields(2) = "本期发生额"
Adodc3.Recordset.Fields(6) = "贷"
Adodc3.Recordset.Fields(9) = "2"
Adodc3.Recordset.Update
End If
End If
Adodc3.Recordset.MoveNext
Loop
Adodc1.Database.Execute "INSERT INTO ZLCX(日期,凭证号,摘要,会计科目,借贷方向,余额,类别) SELECT 日期,凭证号,摘要,会计科目,借贷方向,余额,类别 FROM PMZJZ WHERE 日期='" & DTPicker1.Value & "' "
Adodc1.Database.Execute "UPDATE ZLCX set 序号='1' WHERE 摘要='期初余额'"

Adodc2.RecordSource = "SELECT ZLCX.会计科目 FROM ZLCX GROUP BY ZLCX.会计科目"
Adodc2.Refresh

If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
Adodc3.RecordSource = "SELECT * FROM ZLCX WHERE ZLCX.会计科目='" & Adodc2.Recordset.Fields(0) & "' ORDER BY VAL(ZLCX.序号)"
Adodc3.Refresh
Adodc3.Recordset.FindFirst "序号='1'"
If Adodc3.Recordset.NoMatch Then
M3 = Adodc3.Recordset.Fields(3)
M4 = Val(Adodc3.Recordset.Fields(4))
M5 = Val(Adodc3.Recordset.Fields(5))
M6 = Adodc3.Recordset.Fields(6)

Adodc3.Recordset.Edit
If Adodc3.Recordset.Fields(6) = "贷" Then
Adodc3.Recordset.Fields(7) = Str(Format(Val(Adodc3.Recordset.Fields(5)) - Val(Adodc3.Recordset.Fields(4)), "#0.00"))
M7 = Adodc3.Recordset.Fields(7)
Else
Adodc3.Recordset.Fields(7) = Str(Format(Val(Adodc3.Recordset.Fields(4)) - Val(Adodc3.Recordset.Fields(5)), "#0.00"))
M7 = Adodc3.Recordset.Fields(7)
End If
Adodc3.Recordset.Update

Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields(2) = "本期发生额及余额"
Adodc3.Recordset.Fields(3) = M3
Adodc3.Recordset.Fields(4) = M4
Adodc3.Recordset.Fields(5) = M5
Adodc3.Recordset.Fields(6) = M6
Adodc3.Recordset.Fields(7) = M7
Adodc3.Recordset.Fields(9) = "3"
Adodc3.Recordset.Update
Else
L = Format(Val(Adodc3.Recordset.Fields(7)), "#0.00")
Adodc3.Recordset.MoveNext
M3 = Adodc3.Recordset.Fields(3)
M4 = Val(Adodc3.Recordset.Fields(4))
M5 = Val(Adodc3.Recordset.Fields(5))
M6 = Adodc3.Recordset.Fields(6)
Adodc3.Recordset.Edit
If Adodc3.Recordset.Fields(6) = "贷" Then
Adodc3.Recordset.Fields(7) = Str(Format(Val(Adodc3.Recordset.Fields(5)) - Val(Adodc3.Recordset.Fields(4)) + L, "#0.00"))
M7 = Adodc3.Recordset.Fields(7)
Else
Adodc3.Recordset.Fields(7) = Str(Format(Val(Adodc3.Recordset.Fields(4)) - Val(Adodc3.Recordset.Fields(5)) + L, "#0.00"))
M7 = Adodc3.Recordset.Fields(7)
End If
Adodc3.Recordset.Update

Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields(2) = "本期发生额及余额"
Adodc3.Recordset.Fields(3) = M3
Adodc3.Recordset.Fields(4) = M4
Adodc3.Recordset.Fields(5) = M5
Adodc3.Recordset.Fields(6) = M6
Adodc3.Recordset.Fields(7) = M7
Adodc3.Recordset.Fields(9) = "3"
Adodc3.Recordset.Update

End If

Adodc2.Recordset.MoveNext
Loop


Adodc2.RecordSource = "SELECT * FROM ZLCX WHERE ZLCX.序号='1'"
Adodc2.Refresh

Adodc3.RecordSource = "SELECT * FROM ZLCX "
Adodc3.Refresh


Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
Adodc3.Recordset.FindFirst "序号='2' AND 会计科目='" & Adodc2.Recordset.Fields(3) & "'"
If Adodc3.Recordset.NoMatch Then
Adodc3.Database.Execute "INSERT INTO ZLCX(日期,摘要,会计科目,借方金额,贷方金额,借贷方向,余额,序号) VALUES('" & DTPicker3.Value & "','本月合计','" & Adodc2.Recordset.Fields(3) & "','" & Adodc2.Recordset.Fields(4) & "','" & Adodc2.Recordset.Fields(5) & "','" & Adodc2.Recordset.Fields(6) & "','" & Adodc2.Recordset.Fields(7) & "','3')"
End If
Adodc2.Recordset.MoveNext
Loop

Adodc1.Database.Execute "UPDATE ZLCX SET 借方金额=FORMAT(借方金额,'#0.00'),贷方金额=FORMAT(贷方金额,'#0.00'),余额=FORMAT(余额,'#0.00')"
Adodc1.Database.Execute "DELETE * FROM ZLCX WHERE 会计科目=NULL"
Adodc1.Database.Execute "UPDATE ZLCX SET 凭证号='结-'+'" & Text3.Text & "' WHERE 凭证号=NULL"
Adodc1.RecordSource = "SELECT * FROM ZLCX ORDER BY ZLCX.会计科目,VAL(ZLCX.序号),ZLCX.日期"
Adodc1.Refresh


BT = "按总账汇总"

End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
Call OutadodcToExcel3(VSFlexGrid1, 5, 6, 8, "明细打印")
End Sub

Private Sub DTPicker3_Change()
Text3.Text = Month(DTPicker3.Value)
Text4.Text = Val(Text3.Text) + 1
End Sub

Private Sub DTPicker3_CloseUp()
Text3.Text = Month(DTPicker3.Value)
Text4.Text = Val(Text3.Text) + 1
End Sub

Private Sub DTPicker5_Change()
Text1.Text = DTPicker5.Value
End Sub

Private Sub DTPicker5_CloseUp()
Text1.Text = DTPicker5.Value
End Sub
Private Sub DTPicker6_Change()
Text2.Text = DTPicker6.Value
End Sub

Private Sub DTPicker6_CloseUp()
Text2.Text = DTPicker6.Value
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption + "操作年度 " + ljb
On Error Resume Next
DTPicker3 = Date
DTPicker2 = Date
DTPicker1 = Date
Text3.Text = Month(Date)
Text4.Text = Val(Text3.Text) + 1
DataCombo2.Text = ""
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc7.RecordSource = "rqsd"
Adodc7.Refresh
Adodc7.Recordset.FindFirst "月份='" & Text3.Text & "'"
If Adodc7.Recordset.NoMatch Then
Exit Sub
End If

Text1.Text = Adodc7.Recordset.Fields(0)
Text2.Text = Adodc7.Recordset.Fields(1)

DTPicker5.Value = Adodc7.Recordset.Fields(0)
DTPicker6.Value = Adodc7.Recordset.Fields(1)



Combo1.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc1.RecordSource = "SELECT * FROM ZFLZ WHERE ZFLZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY ZFLZ.日期,ZFLZ.凭证号"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc2.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc4.RecordSource = "CWMC"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc5.RecordSource = "SELECT 科目名称 FROM CWMC WHERE LEN(科目编号)=4 GROUP BY 科目名称"
Adodc5.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc6.RecordSource = "SELECT 科目名称 FROM CWMC WHERE LEN(科目编号)=4 GROUP BY 科目名称"
Adodc6.Refresh

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(3) = 2000
VSFlexGrid1.ColWidth(4) = 2500
VSFlexGrid1.ColWidth(7) = 700
VSFlexGrid1.ColWidth(8) = 700
VSFlexGrid1.ColWidth(9) = 700
End Sub

Private Sub Text3_Change()
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc7.RecordSource = "rqsd"
Adodc7.Refresh
Adodc7.Recordset.FindFirst "月份='" & Text3.Text & "'"
If Adodc7.Recordset.NoMatch Then
Exit Sub
End If

Text1.Text = Adodc7.Recordset.Fields(0)
Text2.Text = Adodc7.Recordset.Fields(1)
DTPicker1.Value = Adodc7.Recordset.Fields(0)
DTPicker5.Value = Adodc7.Recordset.Fields(0)
DTPicker6.Value = Adodc7.Recordset.Fields(1)

End Sub

Private Sub VSFlexGrid1_dblClick()
With VSFlexGrid1
    c = .Col: r = .Row
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
    Call VSFlexGrid1_dblClick
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid1.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    VSFlexGrid1.Text = Text1111.Text
    Text1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move r - 1
Adodc1.Recordset.Edit
Adodc1.Recordset.Fields(c - 1) = Text1111.Text
Adodc1.Recordset.Update
Text1111.Visible = False
End Sub

Private Sub Text4_Change()
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc7.RecordSource = "rqsd"
Adodc7.Refresh
Adodc7.Recordset.FindFirst "月份='" & Text4.Text & "'"
If Adodc7.Recordset.NoMatch Then
Exit Sub
End If
DTPicker2.Value = Adodc7.Recordset.Fields(0)
End Sub
