VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form Formw216 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料调账"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   0
      Left            =   2040
      TabIndex        =   40
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8520
      Top             =   10440
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
      Left            =   7800
      Top             =   9720
      Width           =   3495
      _ExtentX        =   6165
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
      Left            =   7800
      Top             =   9840
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
      Height          =   375
      Left            =   7320
      Top             =   9000
      Width           =   3135
      _ExtentX        =   5530
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
      Left            =   8280
      Top             =   9600
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
      Left            =   8160
      Top             =   9720
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
      Height          =   2895
      Left            =   480
      TabIndex        =   39
      Top             =   5520
      Width           =   14415
      _cx             =   25426
      _cy             =   5106
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
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8520
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8520
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Enabled         =   0   'False
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1095
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
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1215
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "调账查询"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw216.frx":0000
      Height          =   1815
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw216.frx":0014
      Height          =   390
      Index           =   0
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw216.frx":0028
      Height          =   390
      Index           =   1
      Left            =   3720
      TabIndex        =   10
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   2
      Left            =   3720
      TabIndex        =   11
      Top             =   2040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw216.frx":003C
      Height          =   390
      Index           =   3
      Left            =   3720
      TabIndex        =   12
      Top             =   2640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw216.frx":0050
      Height          =   390
      Index           =   4
      Left            =   8880
      TabIndex        =   13
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   5
      Left            =   8880
      TabIndex        =   14
      Top             =   2040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "T"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   6
      Left            =   8880
      TabIndex        =   15
      Top             =   2640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   7
      Left            =   3720
      TabIndex        =   16
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   8
      Left            =   14040
      TabIndex        =   17
      Top             =   1440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   9
      Left            =   14040
      TabIndex        =   18
      Top             =   2040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   688
      _Version        =   393216
      Locked          =   -1  'True
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw216.frx":0064
      Height          =   390
      Index           =   10
      Left            =   14040
      TabIndex        =   19
      Top             =   2640
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8520
      TabIndex        =   33
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   245432321
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8520
      TabIndex        =   34
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   245432321
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   11880
      TabIndex        =   38
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   245432321
      CurrentDate     =   39883
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   1
      Left            =   2040
      TabIndex        =   41
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   2
      Left            =   2040
      TabIndex        =   42
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   3
      Left            =   2040
      TabIndex        =   43
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   4
      Left            =   7200
      TabIndex        =   44
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   5
      Left            =   7200
      TabIndex        =   45
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   6
      Left            =   7200
      TabIndex        =   46
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   7
      Left            =   2040
      TabIndex        =   47
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   8
      Left            =   12360
      TabIndex        =   48
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   9
      Left            =   12360
      TabIndex        =   49
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   10
      Left            =   12360
      TabIndex        =   50
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "导入日期"
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
      Left            =   11880
      TabIndex        =   37
      Top             =   120
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
      Left            =   8520
      TabIndex        =   36
      Top             =   120
      Width           =   1815
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
      Left            =   8520
      TabIndex        =   35
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料名称"
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
      Left            =   480
      TabIndex        =   30
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "供应单位"
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
      Left            =   480
      TabIndex        =   29
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料单位"
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
      Left            =   480
      TabIndex        =   28
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料规格"
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
      Left            =   480
      TabIndex        =   27
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "批次"
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
      Left            =   5640
      TabIndex        =   26
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
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
      Left            =   5640
      TabIndex        =   25
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Index           =   3
      Left            =   480
      TabIndex        =   24
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单价"
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
      Index           =   3
      Left            =   5640
      TabIndex        =   23
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "金额"
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
      Left            =   10680
      TabIndex        =   22
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "数量"
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
      Left            =   10680
      TabIndex        =   21
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "库类"
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
      Left            =   10680
      TabIndex        =   20
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "Formw216"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If DataCombo1(7).Text = "" Or DataCombo1(10).Text = "" Then
Exit Sub
End If
Data1.Recordset.AddNew
For i = 0 To 10
Data1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If DataCombo1(7).Text = "" Or DataCombo1(10).Text = "" Then
Exit Sub
End If
If MsgBox("确定修改吗", vbYesNo) = vbNo Then Exit Sub
If DataCombo1(7).Text = "" Then Exit Sub
Data1.Recordset.Edit
For i = 0 To 10
Data1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

End Sub

Private Sub Command4_Click()
If MsgBox("确定删除吗", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub Command5_Click()
'Call OutDataToExcel(VSFlexGrid1, 10, DataCombo1(7).Text)
End Sub

Private Sub Command6_Click()
If DataCombo1(10).Text = "" Then
MsgBox ("请选择库类!")
Exit Sub
End If
If MsgBox("确认转入日期为：" + Text1.Text + "----" + Text2.Text + " 正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确认转入日期为：" + Text1.Text + "----" + Text2.Text + " 再次确认", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确认导入日期为：" + Trim(DTPicker3.value) + " 最后确认？", vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "DELETE * FROM CKGL WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND 库别='采购入库' AND 单据号='00000000' AND 库类='" & DataCombo1(10).Text & "'"
Data1.Database.Execute "INSERT INTO CKGL(供应单位,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,库类) SELECT 单号,材料名称,材料规格,材料单位,颜色,批次,数量,单价,金额,BL FROM CLTZH WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND BL='" & DataCombo1(10).Text & "'"
Data1.Database.Execute "UPDATE CKGL SET 供应单位='调账',库别='采购入库',单据号='00000000',序号='0',日期='" & DTPicker3.value & "' WHERE 库别=NULL AND 库类='" & DataCombo1(10).Text & "'"
MsgBox ("导入成功！")
End Sub

Private Sub Command7_Click()
Data1.RecordSource = "SELECT * FROM CLTZH WHERE 材料名称='" & DataCombo1(1).Text & "' AND 日期=CDATE('" & DataCombo1(7).Text & "') ORDER BY BL"
Data1.Refresh
End Sub

Private Sub Command8_Click()
If DataCombo1(10).Text = "" Then
       Data1.RecordSource = "SELECT * FROM CLTZH WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
       Data1.Refresh
Else
       Data1.RecordSource = "SELECT * FROM CLTZH WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND BL='" & DataCombo1(10).Text & "'"
       Data1.Refresh
End If
End Sub

Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 7
Data1.RecordSource = "SELECT * FROM CLTZH WHERE 日期=CDATE('" & DataCombo1(7).Text & "') ORDER BY BL"
Data1.Refresh
       Case 6
       If DataCombo1(6).Text <> "" And DataCombo1(8).Text <> "" Then
       DataCombo1(9).Text = Format(DataCombo1(6).Text * DataCombo1(8).Text, "#0.00")
       End If
       Case 8
       If DataCombo1(6).Text <> "" And DataCombo1(8).Text <> "" Then
       DataCombo1(9).Text = Format(DataCombo1(6).Text * DataCombo1(8).Text, "#0.00")
       End If
       Case 10
       Data1.RecordSource = "SELECT * FROM CLTZH WHERE 日期=CDATE('" & DataCombo1(7).Text & "') AND BL='" & DataCombo1(10).Text & "'"
       Data1.Refresh
Data5.RecordSource = "select 材料名称 from KCJL WHERE BL='" & DataCombo1(10).Text & "' GROUP BY 材料名称"
Data5.Refresh
End Select
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 1
       Data1.RecordSource = "SELECT * FROM CLTZH WHERE 材料名称='" & DataCombo1(1).Text & "' AND 日期=CDATE('" & DataCombo1(7).Text & "') AND BL='" & DataCombo1(10).Text & "'"
       Data1.Refresh
       Case 7
Data1.RecordSource = "SELECT * FROM CLTZH WHERE 日期=CDATE('" & DataCombo1(7).Text & "') ORDER BY BL"
Data1.Refresh
       Case 10
Data1.RecordSource = "SELECT * FROM CLTZH WHERE 日期=CDATE('" & DataCombo1(7).Text & "') AND BL='" & DataCombo1(10).Text & "'"
Data1.Refresh
Data5.RecordSource = "select 材料名称 from KCJL WHERE BL='" & DataCombo1(10).Text & "' GROUP BY 材料名称"
Data5.Refresh

End Select
End Sub



Private Sub DTPicker1_Change()
Text1.Text = DTPicker1.value
End Sub

Private Sub DTPicker1_CloseUp()
Text1.Text = DTPicker1.value
End Sub
Private Sub DTPicker2_Change()
Text2.Text = DTPicker2.value
End Sub

Private Sub DTPicker2_CloseUp()
Text2.Text = DTPicker2.value
End Sub

Private Sub Form_Load()
For i = 0 To 10
DataCombo1(i).Text = ""
Next
DataCombo1(7).Text = Date
Text1.Text = Date
Text2.Text = Date
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\MDB"
Data1.RecordSource = "SELECT * FROM CLTZH WHERE 日期=CDATE('" & DataCombo1(7).Text & "')"
Data1.Refresh
Data2.DatabaseName = "d:\数据库\bfrz\" + ljb + "\mdb"
Data2.RecordSource = "select MC from KL  GROUP BY MC"
Data2.Refresh
Data3.DatabaseName = "d:\数据库\bfrz\" + ljb + "\SCZYJHD.mdb"
Data3.RecordSource = "select 简称 from GYS GROUP BY 简称"
Data3.Refresh
Data4.DatabaseName = "d:\数据库\bfrz\" + ljb + "\mdb"
Data4.RecordSource = "select MC from CLDW GROUP BY MC"
Data4.Refresh
Data5.DatabaseName = "d:\数据库\bfrz\" + ljb + "\mdb"
Data5.RecordSource = "select 材料名称 from KCJL GROUP BY 材料名称"
Data5.Refresh
Data6.DatabaseName = "d:\数据库\bfrz\" + ljb + "\SCZYJHD.mdb"
Data6.RecordSource = "select YS from YS GROUP BY YS"
Data6.Refresh

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 1600
For i = 3 To 11
VSFlexGrid1.ColWidth(i) = 1200
Next
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
If Data1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To 10
DataCombo1(i).Text = Data1.Recordset.Fields(i)
Next
End Sub

