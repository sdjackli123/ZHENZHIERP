VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formd44 
   BackColor       =   &H00C0E0FF&
   Caption         =   "����������Ϣ"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   29
      Text            =   "Text8"
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formd44.frx":0000
      Left            =   1320
      List            =   "Formd44.frx":0019
      TabIndex        =   26
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   7560
      Top             =   10200
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
      Caption         =   "Adodc7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   7440
      Top             =   10200
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
      Caption         =   "Adodc6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   7440
      Top             =   10200
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
      Caption         =   "Adodc5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   7920
      Top             =   10200
      Visible         =   0   'False
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
         Name            =   "����"
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
      Left            =   8280
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
         Name            =   "����"
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
      Top             =   10200
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
         Name            =   "����"
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
      Left            =   7680
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ɾ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   240
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   0
   End
   Begin VB.Data Data7 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   22320
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   21360
      TabIndex        =   19
      Text            =   "Text5"
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�䷽"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   19800
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formd44.frx":0057
      Left            =   1320
      List            =   "Formd44.frx":0061
      TabIndex        =   13
      Text            =   "Combo2"
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   12600
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   600
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   17880
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   17880
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   17880
      TabIndex        =   6
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   423886849
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   17880
      TabIndex        =   7
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   423886849
      CurrentDate     =   36892
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VSFlex8DAOCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formd44.frx":0071
      Height          =   9135
      Left            =   360
      TabIndex        =   24
      Top             =   4680
      Width           =   21975
      _cx             =   38761
      _cy             =   16113
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Begin MSAdodcLib.Adodc Adodc8 
         Height          =   375
         Left            =   9000
         Top             =   6480
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
         Caption         =   "Adodc8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formd44.frx":0085
      Height          =   330
      Left            =   1320
      TabIndex        =   25
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formd44.frx":009A
      Height          =   2775
      Left            =   360
      TabIndex        =   23
      Top             =   1680
      Width           =   21975
      _cx             =   38761
      _cy             =   4895
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      FormatString    =   $"Formd44.frx":00AF
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
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   28
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ϵ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   12600
      TabIndex        =   17
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ͷ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "ɫ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   9480
      TabIndex        =   12
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ͻ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   1
      Left            =   16920
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
      Height          =   375
      Index           =   1
      Left            =   16920
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "Formd44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public c, r As Integer: Public l9 As String: Dim sz(9) As String: Dim ZS(10) As String
Private Sub Command1_Click()
Formd331.Data13.Refresh
Formd331.VSFlexGrid1.ColWidth(0) = 400
Formd331.VSFlexGrid1.ColWidth(1) = 0
Formd331.VSFlexGrid1.ColWidth(2) = 0
Formd331.VSFlexGrid1.ColWidth(5) = 400
Formd331.VSFlexGrid1.ColWidth(7) = 2000
Formd331.VSFlexGrid1.ColWidth(8) = 800
Formd331.VSFlexGrid1.ColWidth(10) = 600
Formd331.VSFlexGrid1.ColWidth(13) = 0
Formd331.VSFlexGrid1.ColWidth(14) = 0
Formd331.VSFlexGrid1.ColWidth(17) = 0
Formd331.VSFlexGrid1.ColWidth(18) = 0
Formd331.VSFlexGrid1.ColWidth(19) = 2600
Formd331.VSFlexGrid1.ColWidth(20) = 0
Formd331.VSFlexGrid1.ColWidth(22) = 0
Formd331.VSFlexGrid1.ColWidth(23) = 0
Formd331.VSFlexGrid1.ColWidth(24) = 0
Formd331.VSFlexGrid1.ColWidth(25) = 0
Unload Me
End Sub


Private Sub Command2_Click()
On Error Resume Next

'Adodc2.RecordSource = "select * from pld where ���='" & Text1 & "'"
Adodc2.Refresh
'If Not Adodc2.Recordset.EOF Then
'MsgBox ("�ϵ��Ѿ���ӡ�� ��ֹ��ӡ")
'Exit Sub
'End If

If Data1.Recordset.EOF Then
MsgBox ("��ѡ���ϵ���Ϣ")
Else
Call plda(Data2, Data3, Data4, Text1.Text, Adodc2)
End If
End Sub

Private Sub Command3_Click()

Adodc4.RecordSource = "SELECT �ͻ�,����,Ʒ��,��ɫ,ɫ��,����,����,��̨,����,��Ϣ,���,��ֵ,���׹��� FROM pld where �ͻ� like '%'+'" & DataCombo2 & "'+'%' and ���� like '%'+'" & Text8 & "'+'%' and ��Ϣ like '%'+'" & Combo1 & "'+'%' and ɫ�� like '%'+'" & Text2 & "'+'%' and cast(CONVERT(varchar,����, 23) as datetime) BETWEEN cast('" & Text7.Text & "' as datetime) AND cast('" & Text3.Text & "' as datetime)  ORDER BY ���� DESC,��� desc"
Adodc4.Refresh

End Sub


Private Sub Command4_Click()
Adodc4.RecordSource = "SELECT �ͻ�,����,Ʒ��,��ɫ,ɫ��,����,����,��̨,����,��Ϣ,���,��ֵ,���׹��� FROM pld where ɫ�� like '%'+'" & Text2.Text & "'+'%' AND cast(CONVERT(varchar,����, 23) as datetime) BETWEEN cast('" & Text7.Text & "' as datetime) AND cast('" & Text3.Text & "' as datetime)  ORDER BY ���� DESC,��� desc"
Adodc4.Refresh
End Sub

Private Sub Command5_Click()
On Error Resume Next
If Text1.Text = "" Then
MsgBox ("��Ϣ���룬���ܸ��ƹ���")
Exit Sub
End If

If Formd331.Option1(0).value = True Then
Adodc3.RecordSource = "SELECT * FROM DBPLDBH where ����='" & yhdm & "'"
Adodc3.Refresh

KLL = yhdm + "1" ''''''''''''OK
If Adodc3.Recordset.EOF Then
KLL = yhdm + "1" ''''''''''''OK
Else
L = Val(Adodc3.Recordset.Fields(0))
KLL = yhdm + Trim(L + 1) '''''''''''''OK
End If
End If

If Formd331.Option1(1).value = True Then
KLL = Text1
End If

Data1.Database.Execute "insert into pldd(��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,�����,����,����,��������) in'" & lo & "' SELECT ��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,�����,����,����,��������  FROM plda WHERE �ϵ����='" & Text1.Text & "'"
Data1.Database.Execute "UPDATE pldd SET �ϵ����='" & KLL & "' WHERE �ϵ����=NULL"
Formd331.Text2.Text = KLL
Formd331.Data13.Refresh
Formd331.VSFlexGrid1.ColWidth(0) = 400
Formd331.VSFlexGrid1.ColWidth(1) = 0
Formd331.VSFlexGrid1.ColWidth(2) = 0
Formd331.VSFlexGrid1.ColWidth(5) = 800
Formd331.VSFlexGrid1.ColWidth(7) = 6000
Formd331.VSFlexGrid1.ColWidth(8) = 800
Formd331.VSFlexGrid1.ColWidth(10) = 600
Formd331.VSFlexGrid1.ColWidth(13) = 0
Formd331.VSFlexGrid1.ColWidth(14) = 0
Formd331.VSFlexGrid1.ColWidth(17) = 0
Formd331.VSFlexGrid1.ColWidth(18) = 0
Formd331.VSFlexGrid1.ColWidth(19) = 2600
Formd331.VSFlexGrid1.ColWidth(20) = 0
Formd331.VSFlexGrid1.ColWidth(22) = 0
Formd331.VSFlexGrid1.ColWidth(23) = 0
Formd331.VSFlexGrid1.ColWidth(24) = 0
Formd331.VSFlexGrid1.ColWidth(25) = 0
Unload Me
End Sub

Private Sub Command6_Click()
    On Error GoTo errorhandler
    
    ' ������Ƿ�Ϊ��
    If Text1.Text = "" Then
        MsgBox "��������"
        Exit Sub
    End If

    ' �����ݿ��л�ȡ����
    Adodc8.RecordSource = "SELECT ���� FROM pld WHERE ���='" & Text1.Text & "'"
    Adodc8.Refresh
    
    ' ����¼�Ƿ����
    If Adodc8.Recordset.EOF Then
        MsgBox "��Ų�����"
        Exit Sub
    End If
    
    ' ��ȡ����
    Dim recordDate As Date
    recordDate = CDate(Adodc8.Recordset.Fields("����").value)
    
    
    ' �ж��Ƿ�����ɾ��
    If yhm <> "root" And DateDiff("d", recordDate, Date) > 7 Then
        MsgBox "�ϵ��ѳ���7�첻����ɾ��"
        Exit Sub
    End If

    ' ȷ��ɾ������
    If MsgBox("ȷ��ɾ�����ϱ��Ϊ" & Text1.Text & "��", vbYesNo) = vbNo Then Exit Sub

    ' ɾ������
    Dim sql1 As String
    Dim sql2 As String
    Dim sql3 As String
    
    sql1 = "DELETE FROM pld WHERE ���='" & Text1.Text & "'"
    sql2 = "DELETE FROM pldr WHERE �ϵ����='" & Text1.Text & "'"
    sql3 = "DELETE FROM pldb WHERE �ϵ����='" & Text1.Text & "'"
    
    ' ִ��ɾ�����
    RD.Open sql1, conn, adOpenStatic, adLockOptimistic
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic
    RD.Open sql3, conn, adOpenStatic, adLockOptimistic
    
    ' ɾ���ɹ���ʾ
    MsgBox "ɾ���ɹ���"
    
    ' ˢ������
    Adodc4.Refresh
    Exit Sub

errorhandler:
    ' ������
    MsgBox "��������: " & Err.Description
End Sub




Private Sub DTPicker3_Change()
Text7.Text = DTPicker3.value
End Sub

Private Sub DTPicker3_CloseUp()
Text7.Text = DTPicker3.value
End Sub

Private Sub DTPicker4_Change()
Text3.Text = DTPicker4.value
End Sub

Private Sub DTPicker4_CloseUp()
Text3.Text = DTPicker4.value
End Sub

Private Sub Form_Load()

On Error Resume Next
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Combo1.Text = ""
DataCombo2.Text = ""
DTPicker3.value = Date - 60
DTPicker4.value = Date
Text3.Text = Date
Text7.Text = Date - 60

If InStr(yhmk, "����") > 0 Then
Command6.Visible = True
Else
Command6.Visible = False
End If

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Data1.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data2.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data3.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data4.DatabaseName = App.Path & "\AccessBase\DB.mdb"

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select ��� from khZL  group by ���"
Adodc5.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Data6.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid2.ColWidth(0) = 200


End Sub


Private Sub VSFlexGrid1_dblClick()
If Text1.Text = "" Then
MsgBox ("��Ϣ���룬���ܸ��ƹ���")
Exit Sub
End If

Adodc3.RecordSource = "SELECT * FROM DBPLDBH where ����='" & yhdm & "'"
Adodc3.Refresh

KLL = yhdm + "1" ''''''''''''OK
If Adodc3.Recordset.EOF Then
KLL = yhdm + "1" ''''''''''''OK
Else
L = Val(Adodc3.Recordset.Fields(0))
KLL = yhdm + Trim(L + 1) '''''''''''''OK
End If

If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Data1.Recordset.Move rs - 1
If MsgBox(Data1.Recordset.Fields(4) + "  ת�����ϵ���", vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "insert into pldd(�ϵ����,��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,�����,����,����)  VALUES('" & KLL & "','" & Data1.Recordset.Fields(1) & "','" & Data1.Recordset.Fields(2) & "','" & Data1.Recordset.Fields(3) & "','" & Data1.Recordset.Fields(4) & "','" & Data1.Recordset.Fields(5) & "','" & Data1.Recordset.Fields(6) & "','" & Data1.Recordset.Fields(7) & "', '" & Data1.Recordset.Fields(8) & "','" & Data1.Recordset.Fields(9) & "','')"

End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc4.Recordset.EOF Then
Text2.Text = ""
Exit Sub
End If
Text1.Text = ""
rs = VSFlexGrid2.Row
Adodc4.Recordset.MoveFirst
Adodc4.Recordset.Move rs - 1
Text1.Text = Adodc4.Recordset.Fields(10)
If Formd331.Option1(1).value = True Then
Formd331.Text5 = Adodc4.Recordset.Fields(1)
Formd331.Text4 = Adodc4.Recordset.Fields(7)
Formd331.DataCombo20 = Adodc4.Recordset.Fields(6)
Formd331.Text11 = Adodc4.Recordset.Fields(11)
Formd331.Combo2 = Adodc4.Recordset.Fields(9)
Formd331.DataCombo9(2) = Adodc4.Recordset.Fields(5)
End If
End Sub


Private Sub Text1_Change()
On Error Resume Next
Call pfdfj
Data1.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data1.RecordSource = "SELECT �ϵ����,��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,��������,���ϵ�λ,�������� FROM plda WHERE �ϵ����='" & Text1.Text & "' ORDER BY plda.��������,plda.�����"
Data1.Refresh
VSFlexGrid1.ColFormat(9) = "#0.####"
End Sub

Private Sub pfdfj()
On Error Resume Next
Data6.Database.Execute "delete * from plda"

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select * from pld where  ���='" & Text1.Text & "'"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
Adodc7.Recordset.MoveFirst

For i = 0 To 10
ZS(i) = Adodc7.Recordset.Fields(i)
Next

mb = 0
For i = 12 To 61
If Adodc7.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next

ProgressBar1.Visible = True
For i = 12 To mb + 12
If Adodc7.Recordset.Fields(i) <> "" Then
sz(0) = Mid(Adodc7.Recordset.Fields(i), 1, InStr(Adodc7.Recordset.Fields(i), "(") - 1)
sz(1) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "(") + 1, InStr(Adodc7.Recordset.Fields(i), ")") - InStr(Adodc7.Recordset.Fields(i), "(") - 1)
sz(2) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), ")") + 1, InStr(Adodc7.Recordset.Fields(i), "-") - InStr(Adodc7.Recordset.Fields(i), ")") - 1)
sz(3) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "-") + 1, InStr(Adodc7.Recordset.Fields(i), "\") - InStr(Adodc7.Recordset.Fields(i), "-") - 1)
sz(4) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "\") + 1, InStr(Adodc7.Recordset.Fields(i), "#") - InStr(Adodc7.Recordset.Fields(i), "\") - 1)
sz(5) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "#") + 1, InStr(Adodc7.Recordset.Fields(i), "^") - InStr(Adodc7.Recordset.Fields(i), "#") - 1)
sz(6) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "^") + 1, InStr(Adodc7.Recordset.Fields(i), "[") - InStr(Adodc7.Recordset.Fields(i), "^") - 1)
sz(7) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "[") + 1, InStr(Adodc7.Recordset.Fields(i), "]") - InStr(Adodc7.Recordset.Fields(i), "[") - 1)
sz(8) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "]") + 1, InStr(Adodc7.Recordset.Fields(i), "{") - InStr(Adodc7.Recordset.Fields(i), "]") - 1)
sz(9) = Mid(Adodc7.Recordset.Fields(i), InStr(Adodc7.Recordset.Fields(i), "{") + 1)

L = i - 11
Data6.Database.Execute "insert into plda(���,����,ѹ��,�������,�䷽��,����,���ϴ�ӡԱ,Ⱦ��������,��������,������Ϣ,�ϵ����,��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,��������,���ϵ�λ,����,�����) VALUES('" & ZS(0) & "','" & ZS(1) & "','" & ZS(2) & "','" & ZS(3) & "','" & ZS(4) & "','" & ZS(5) & "','" & ZS(6) & "','" & ZS(7) & "','" & ZS(8) & "','" & ZS(9) & "','" & ZS(10) & "','" & sz(0) & "','" & sz(1) & "','" & sz(2) & "',  " & _
                                                                        "'" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & L & "')"
ProgressBar1.value = 100 / mb * L
End If
Next
ProgressBar1.Visible = False
End If

End Sub



