VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw19 
   BackColor       =   &H00C0E0FF&
   Caption         =   "客户账目查询---发票"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   7200
      Top             =   9000
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
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
      Height          =   330
      Left            =   7200
      Top             =   9240
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
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
      Left            =   7080
      Top             =   10080
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
      Height          =   330
      Left            =   7080
      Top             =   9480
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
      Left            =   7560
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
      Left            =   7320
      Top             =   9720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   330
      Left            =   7560
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7560
      Top             =   9360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   330
      Left            =   1440
      TabIndex        =   14
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Height          =   3135
      Left            =   480
      TabIndex        =   13
      Top             =   4560
      Width           =   14415
      _cx             =   25426
      _cy             =   5530
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
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw19.frx":0000
      Height          =   2295
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   10790143
      BackColorBkg    =   44718
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw19.frx":0014
      Height          =   330
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81788929
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   81788929
      CurrentDate     =   36892
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择日期范围"
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
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "加工单位"
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
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Formw19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer

Private Sub Command1_Click()
On Error Resume Next
rqq = CDate(Text2.Text) + 1
Adodc6.Database.Execute "DELETE * FROM JGZCX"
Adodc6.Database.Execute "insert into JGZCX(客户,上期累计应收)  SELECT MID(会计科目,INSTR(会计科目,'-')+1),format(SUM(VAL(余额)),'#0.00') FROM PMMXJZ WHERE 借贷方向='借' AND 日期=CDATE('" & Text1.Text & "') GROUP BY MID(会计科目,INSTR(会计科目,'-')+1)"
Adodc5.Database.Execute "insert into JGZCX(客户,本期应收款) in'd:\数据库\bfrz\" + ljb + "\cw.mdb' SELECT 购货单位,format(SUM(VAL(金额)),'#0.00') FROM cpfh WHERE  日期 between cdate('" & Text1.Text & "') and cdate('" & Text2.Text & "')  GROUP BY 购货单位"
Adodc5.Database.Execute "insert into JGZCX(客户,本期应收款) in'd:\数据库\bfrz\" + ljb + "\cw.mdb' SELECT 购货单位,format(SUM(VAL(金额)),'#0.00') FROM LSFH WHERE  日期 between cdate('" & Text1.Text & "') and cdate('" & Text2.Text & "')  GROUP BY 购货单位"
Adodc6.Database.Execute "insert into JGZCX(客户,本期开票)  SELECT 客户,format(SUM(VAL(开票金额)),'#0.00') FROM FHFP WHERE 开票日期 between cdate('" & Text1.Text & "') and cdate('" & rqq & "') GROUP BY 客户"
Adodc6.Database.Execute "insert into JGZCX(客户,上期累计开票) SELECT 客户,开票金额 FROM PMFHFP WHERE  结转日期=CDATE('" & Text1.Text & "')"
Adodc6.Database.Execute "insert into JGZCX(客户,上期累计未开票) SELECT 客户,未开金额 FROM PMFHFP WHERE  结转日期=CDATE('" & Text1.Text & "')"

Adodc6.RecordSource = "SELECT * FROM JGZCX"
Adodc6.Refresh

If Adodc6.Recordset.EOF Then Exit Sub
Adodc6.Recordset.MoveFirst
Do While Not Adodc6.Recordset.EOF
Adodc8.RecordSource = "SELECT * FROM KHZL WHERE INSTR(简称,'" & Adodc6.Recordset.Fields(0) & "')>0"
Adodc8.Refresh
If Adodc8.Recordset.EOF Then
Adodc6.Recordset.Delete
End If
Adodc6.Recordset.MoveNext
Loop


Adodc6.Database.Execute "UPDATE JGZCX SET 类别='1'"
Adodc6.Database.Execute "UPDATE JGZCX SET 日期范围='" & Text1.Text & "'+'--'+'" & Text2.Text & "'"
Adodc6.Database.Execute "UPDATE JGZCX SET 上期累计应收='0' WHERE 上期累计应收=NULL"
Adodc6.Database.Execute "UPDATE JGZCX SET 本期应收款='0' WHERE 本期应收款=NULL"
Adodc6.Database.Execute "UPDATE JGZCX SET 本期累计应收款='0' WHERE 本期累计应收款=NULL"
Adodc6.Database.Execute "UPDATE JGZCX SET 本期已收款='0' WHERE 本期已收款=NULL"
Adodc6.Database.Execute "UPDATE JGZCX SET 上期累计开票='0' WHERE 上期累计开票=NULL"
Adodc6.Database.Execute "UPDATE JGZCX SET 本期开票='0' WHERE 本期开票=NULL"
Adodc6.Database.Execute "UPDATE JGZCX SET 上期累计未开票='0' WHERE 上期累计未开票=NULL"
Adodc6.Database.Execute "UPDATE JGZCX SET 本期累计开票='0' WHERE 本期累计开票=NULL"
Adodc6.Database.Execute "UPDATE JGZCX SET 本期未开='0' WHERE 本期未开=NULL"
Adodc6.Database.Execute "UPDATE JGZCX SET 本期累计未开='0' WHERE 本期累计未开=NULL"


Adodc6.Database.Execute "insert into JGZCX(客户,日期范围,上期累计应收,本期应收款,本期累计应收款,本期已收款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期未开,本期累计未开) SELECT 客户,日期范围,format(SUM(VAL(上期累计应收)),'#0.00'),format(SUM(VAL(本期应收款)),'#0.00'),format(SUM(VAL(本期累计应收款)),'#0.00'),format(SUM(VAL(本期已收款)),'#0.00'),format(SUM(VAL(上期累计开票)),'#0.00'),format(SUM(VAL(本期开票)),'#0.00'),format(SUM(VAL(本期累计开票)),'#0.00'),format(SUM(VAL(上期累计未开票)),'#0.00'),format(SUM(VAL(本期未开)),'#0.00'),format(SUM(VAL(本期累计未开)),'#0.00') FROM JGZCX GROUP BY 客户,日期范围 "
Adodc6.Database.Execute "DELETE *  FROM  JGZCX WHERE 类别='1'"
Adodc6.Database.Execute "UPDATE JGZCX SET 本期未开=format(VAL(本期应收款)-VAL(本期开票),'#0.00')"
Adodc6.Database.Execute "UPDATE JGZCX SET 欠款=format(VAL(上期累计应收)+VAL(本期应收款)-VAL(本期已收款),'#0.00'),本期累计应收款=format(VAL(上期累计应收)+VAL(本期应收款),'#0.00'),本期累计开票=format(VAL(上期累计开票)+VAL(本期开票),'#0.00'),本期累计未开=format(VAL(上期累计未开票)+VAL(本期未开),'#0.00')"


Adodc6.RecordSource = "SELECT 客户,上期累计未开票,本期应收款,本期开票,本期累计未开,日期范围 FROM JGZCX"
Adodc6.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If DataCombo1.Text = "" Then
Adodc6.RecordSource = "SELECT 客户,本期应收款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期未开,本期累计未开 FROM JGZCX"
Adodc6.Refresh
Call OutadodcToExcel6(VSFlexGrid1, 3, 4, 5, 6, 7, 8, "客户账目查询--收款" + "截止日期:" + Text2.Text)
Adodc6.RecordSource = "SELECT 客户,本期应收款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期未开,本期累计未开,日期范围 FROM JGZCX"
Adodc6.Refresh
Else
Adodc6.RecordSource = "SELECT 客户,本期应收款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期未开,本期累计未开 FROM JGZCX WHERE 客户='" & DataCombo1.Text & "'"
Adodc6.Refresh
Call OutadodcToExcel6(VSFlexGrid1, 3, 4, 5, 6, 7, 8, "客户账目查询--收款" + "截止日期:" + Text2.Text)
Adodc6.RecordSource = "SELECT 客户,本期应收款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期未开,本期累计未开,日期范围 FROM JGZCX WHERE 客户='" & DataCombo1.Text & "'"
Adodc6.Refresh
End If
End Sub

Private Sub dataCombo1_Click(Area As Integer)
If DataCombo1.Text = "" Then
Adodc6.RecordSource = "SELECT 客户,上期累计未开票,本期应收款,本期开票,本期累计未开,日期范围 FROM JGZCX"
Adodc6.Refresh
Else
Adodc6.RecordSource = "SELECT 客户,上期累计未开票,本期应收款,本期开票,本期累计未开,日期范围 FROM JGZCX WHERE 客户='" & DataCombo1.Text & "'"
Adodc6.Refresh
End If
End Sub

Private Sub DTPicker3_Change()
Text1.Text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = DTPicker3.Value
Text1.SetFocus
End Sub

Private Sub DTPicker4_Change()
Text2.Text = DTPicker4.Value
End Sub

Private Sub DTPicker4_CloseUp()
Text2.Text = DTPicker4.Value
Text2.SetFocus
End Sub

Private Sub Form_Load()
Text1.Text = Date
Text2.Text = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
DataCombo1.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc1.RecordSource = "select 简称 from KHZL  GROUP BY 简称"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc6.RecordSource = "SELECT 客户,上期累计未开票,本期应收款,本期开票,本期累计未开,日期范围 FROM JGZCX"
Adodc6.Refresh
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc7.RecordSource = "rqsd"
Adodc7.Refresh
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"

For i = 2 To 5
VSFlexGrid1.ColWidth(i) = 1700
Next
VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(6) = 2200

End Sub

Private Sub Label3_DblClick()
DataCombo1.Text = ""
End Sub

Private Sub vSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
S1 = VSFlexGrid1.RowSel
End Sub

Private Sub vSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
S2 = VSFlexGrid1.RowSel
End Sub





