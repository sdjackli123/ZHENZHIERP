VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formh221 
   BackColor       =   &H00C0E0FF&
   Caption         =   "��������䷽����Ϣ"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form21"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "���뷽ʽ"
      Height          =   2415
      Left            =   10200
      TabIndex        =   42
      Top             =   360
      Width           =   855
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   720
      Style           =   1  'Simple Combo
      TabIndex        =   41
      Text            =   "Combo1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "������ѯ"
      Height          =   495
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1560
      Width           =   1330
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ƻ���ѯ"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2760
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   375
      Left            =   7680
      Top             =   9720
      Visible         =   0   'False
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
      Caption         =   "Adodc11"
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   5760
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
      Caption         =   "Adodc10"
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   5280
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
      Caption         =   "Adodc9"
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   5520
      Top             =   9720
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   5040
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
      Left            =   5400
      Top             =   9840
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
      Height          =   375
      Left            =   6000
      Top             =   9600
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
      Height          =   375
      Left            =   5520
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   5760
      Top             =   9840
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
      Left            =   5640
      Top             =   9960
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
      Left            =   6120
      Top             =   9840
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formh221.frx":0000
      Height          =   5175
      Left            =   720
      TabIndex        =   8
      Top             =   4080
      Width           =   13815
      _cx             =   24368
      _cy             =   9128
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
   Begin VB.Data Data11 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   6000
      Top             =   120
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�±��"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   960
      Width           =   1575
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�䷽��"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��������"
      Height          =   495
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�޸ı���"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ɾ��"
      Height          =   495
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�ͻ���Ϣ"
      Height          =   3375
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   9375
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   6600
         TabIndex        =   38
         Text            =   "Text4"
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   600
         TabIndex        =   35
         Text            =   "Text3"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   600
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1800
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   0
         Left            =   6600
         TabIndex        =   29
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6600
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   2760
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8421440
         Format          =   424869889
         CurrentDate     =   39961
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8421440
         Format          =   424869889
         CurrentDate     =   39961
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh221.frx":0015
         Height          =   330
         Index           =   1
         Left            =   1320
         TabIndex        =   21
         Top             =   1320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "���"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh221.frx":002A
         Height          =   330
         Index           =   2
         Left            =   6600
         TabIndex        =   27
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "����������"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   3
         Left            =   1320
         TabIndex        =   23
         Top             =   2280
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   4
         Left            =   1320
         TabIndex        =   24
         Top             =   2760
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh221.frx":003F
         Height          =   330
         Index           =   5
         Left            =   1320
         TabIndex        =   22
         Top             =   1800
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "pm"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   6
         Left            =   6600
         TabIndex        =   26
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh221.frx":0054
         Height          =   330
         Index           =   7
         Left            =   6600
         TabIndex        =   28
         Top             =   1800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "mc"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   8
         Left            =   4440
         TabIndex        =   25
         Top             =   2040
         Visible         =   0   'False
         Width           =   620
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "����˵��"
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
         Index           =   11
         Left            =   5400
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "ɫ�� "
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
         Index           =   8
         Left            =   120
         TabIndex        =   33
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "��ʼ����"
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
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "��������"
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
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "ԡ��"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "�䷽���"
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
         Index           =   9
         Left            =   5400
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
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
         Index           =   4
         Left            =   5400
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "��ɫ "
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
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "ɫ�ſ��"
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
         Index           =   3
         Left            =   5400
         TabIndex        =   9
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "Ʒ��"
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
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "�ͻ�����"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "����/����"
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
         Index           =   14
         Left            =   5400
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "�������"
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
         Index           =   15
         Left            =   5400
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      Height          =   495
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   480
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   11280
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   375
      Left            =   3120
      Top             =   9840
      Visible         =   0   'False
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
      Caption         =   "Adodc11"
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
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ȡ���"
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
      Index           =   7
      Left            =   3840
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "Formh221"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c, r As Integer: Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Dim sz(6) As String
Dim shbh As Integer
Dim cdbhf As Integer
Private Sub Command1_Click()
If DataCombo1(4).Text = "" Then
Adodc5.RecordSource = "SELECT RQ AS ����,KH AS �ͻ�,fz AS ���ո�����,YS AS ��ɫ,SH AS ɫ��,BL AS Ʒ��,DH AS ���,ZL AS �������,IP AS ԡ��,qr as ���,qs as �����,xs as ��ˮ�� FROM ZH where rq between cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime)  ORDER BY DH DESC"
Adodc5.Refresh
Else
Adodc5.RecordSource = "SELECT RQ AS ����,KH AS �ͻ�,fz AS ���ո�����,YS AS ��ɫ,SH AS ɫ��,BL AS Ʒ��,DH AS ���,ZL AS �������,IP AS ԡ��,qr as ���,qs as �����,xs as ��ˮ�� FROM ZH where sh like '%'+'" & DataCombo1(4).Text & "'+'%'  ORDER BY DH DESC"
Adodc5.Refresh
End If
shbh = 1
Command5.Enabled = True
Command4.Enabled = False
Command3.Enabled = False

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 500
Next
End If

End Sub

Private Sub Command2_Click()
gyhys = 0
Unload Me
End Sub


Private Sub Command3_Click()
On Error Resume Next
If MsgBox("ȷ���޸ĵ�����Ϣ��", vbYesNo) = vbNo Then Exit Sub

If DataCombo1(4).Text = "" Or DataCombo1(6).Text = "" Then
MsgBox ("ɫ�š��������������")
Exit Sub
End If

For i = 0 To Adodc5.Recordset.Fields.count - 1
Adodc5.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc5.Recordset.Fields(11) = Text4
Adodc5.Recordset.Update
Adodc5.Refresh

sql1 = "update PFD set �ͻ�='" & DataCombo1(1) & "',Ʒ��='" & DataCombo1(5) & "',ɫ��='" & DataCombo1(4) & "',��ɫ='" & DataCombo1(3) & "',����='" & DataCombo1(2) & "' where ���='" & DataCombo1(6).Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

shbh = 1
Command5.Enabled = True
Command4.Enabled = False
Command3.Enabled = False

End Sub

Private Sub Command4_Click()
On Error Resume Next
If Adodc5.Recordset.EOF Then Exit Sub
If MsgBox("ȷ��ɾ����ɾ��ͬʱ�䷽Ҳ�������", vbYesNo) = vbNo Then Exit Sub
Adodc5.Recordset.Delete
Adodc5.Refresh
sql1 = "delete  from PFD where ���='" & DataCombo1(6).Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
shbh = 1
Command5.Enabled = True
Command4.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command5_Click()
If DataCombo1(4).Text = "" Or DataCombo1(6).Text = "" Then
MsgBox ("ɫ�š��������������")
Exit Sub
End If

Adodc7.RecordSource = "select * from zh where sh='" & DataCombo1(4) & "' and bl='" & DataCombo1(5) & "' and qr<>'����'"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
If MsgBox("���д�ɫ�ŵ���������ɫ�ſ�������ѯ  �Ƿ�������棿", vbYesNo) = vbNo Then Exit Sub
End If

Adodc7.RecordSource = "SELECT ���� FROM dbPFDbh where ����='" & yhdm & "'"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
If Val(Mid(DataCombo1(6).Text, 2)) <= Adodc7.Recordset.Fields(0) Then
If MsgBox("�Ѵ��ڴ��䷽��ţ��Ƿ��Զ��±�ţ�", vbYesNo) = vbNo Then Exit Sub
Adodc1.RecordSource = "SELECT ���� FROM dbPFDbh"
Adodc1.Refresh
DataCombo1(6).Text = yhdm + "1"
If Not Adodc1.Recordset.EOF Then
L = Adodc1.Recordset.Fields(0)
DataCombo1(6).Text = yhdm + Trim(L + 1) '''''''''''''OK
Else
DataCombo1(6).Text = yhdm + "1"
End If
End If
End If

'Adodc7.RecordSource = "select * from v_zh_pfd where ɫ��='" & DataCombo1(4) & "' and ȷ��='���'  and Ʒ��='" & DataCombo1(5) & "'"
'Adodc7.Refresh
'If Not Adodc7.Recordset.EOF Then
'MsgBox ("���д�ɫ�ż�¼������ɫ�ſ�������ѯ")
'Exit Sub
'Else
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "hyszh('" & Now & "','" & DataCombo1(1) & "','" & DataCombo1(2) & "','" & DataCombo1(3) & "','" & DataCombo1(4) & "','" & DataCombo1(5) & "','" & DataCombo1(6) & "','" & DataCombo1(7) & "','" & DataCombo1(8) & "','','" & Text4 & "')"     ' ��ʾ�����ĸ��洢����
Set L = g_Cmd.Execute             ' ִ�д洢����
    g_Cmd.Cancel
'End If

Adodc5.Refresh
Call Command6_Click
End Sub

Private Sub Command6_Click()
On Error Resume Next
Data10.Database.Execute "delete * from pfda"

Adodc11.RecordSource = "select * from pfd where  ���='" & DataCombo1(6).Text & "'"
Adodc11.Refresh
If Not Adodc11.Recordset.EOF Then
Adodc11.Recordset.MoveFirst
mb = 0
For i = 7 To 56
If Adodc11.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next

ProgressBar1.Visible = True
Timer1.Enabled = True
For i = 7 To mb + 7
If Adodc11.Recordset.Fields(i) <> "" Then
sz(0) = Mid(Adodc11.Recordset.Fields(i), 1, InStr(Adodc11.Recordset.Fields(i), "(") - 1)
sz(1) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "(") + 1, InStr(Adodc11.Recordset.Fields(i), ")") - InStr(Adodc11.Recordset.Fields(i), "(") - 1)
sz(2) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), ")") + 1, InStr(Adodc11.Recordset.Fields(i), "-") - InStr(Adodc11.Recordset.Fields(i), ")") - 1)
sz(3) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "-") + 1, InStr(Adodc11.Recordset.Fields(i), "\") - InStr(Adodc11.Recordset.Fields(i), "-") - 1)
sz(4) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "\") + 1, InStr(Adodc11.Recordset.Fields(i), "#") - InStr(Adodc11.Recordset.Fields(i), "\") - 1)
sz(5) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "#") + 1, InStr(Adodc11.Recordset.Fields(i), "^") - InStr(Adodc11.Recordset.Fields(i), "#") - 1)
sz(6) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "^") + 1)
L = i - 6
Data10.Database.Execute "insert into pfda(�ӹ���λ,Ʒ��,��ɫ,ɫ��,������,��������,�䷽���,�䷽����,��������,ԡ��,Ⱦ������,Ⱦ��������,��λ,�䷽,����,�����) VALUES('" & DataCombo1(1).Text & "','" & DataCombo1(5).Text & "','" & DataCombo1(3).Text & "','" & DataCombo1(4).Text & "','" & DataCombo1(2).Text & "','" & DataCombo1(7).Text & "','" & DataCombo1(6).Text & "',CDATE('" & DataCombo1(0).Text & "'),'" & sz(0) & "','" & sz(1) & "','" & sz(2) & "',  " & _
                                                                        "'" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & L & "')"
ProgressBar1.value = 100 / mb * L
End If
Next
ProgressBar1.Visible = False
Timer1.Enabled = False
End If


Formh233.DataCombo1(0).Text = DataCombo1(1).Text
Formh233.DataCombo1(1).Text = DataCombo1(5).Text
Formh233.DataCombo1(3).Text = DataCombo1(3).Text
Formh233.DataCombo1(2).Text = DataCombo1(4).Text
Formh233.DataCombo1(14).Text = DataCombo1(2).Text
Formh233.DataCombo1(15).Text = DataCombo1(7).Text
Formh233.DataCombo1(12).Text = DataCombo1(6).Text
If gyhys = 1 Then
Unload Me
Formh233.Show
Else
Formh233.Show
End If
End Sub

Private Sub Command7_Click()
On Error Resume Next

DataCombo1(6).Text = yhdm + "1"
Adodc1.RecordSource = "SELECT ���� FROM dbpfdbh where ����='" & yhdm & "'"
Adodc1.Refresh

If Not Adodc1.Recordset.EOF Then
L = Adodc1.Recordset.Fields(0)
DataCombo1(6).Text = yhdm + Trim(L + 1)  '''''''''''''OK
Else
DataCombo1(6).Text = yhdm + "1"
End If
shbh = 1
DataCombo1(0).Text = Date

End Sub

Private Sub Command8_Click()
hysbl = 1
Forma172.Command2.Visible = True
Forma172.Show
End Sub

Private Sub Command9_Click()
Formh224.Show
End Sub

Private Sub DataCombo1_Change(Index As Integer)
Select Case Index
        Case 4
If Option1(0).value = True Then
If Len(DataCombo1(4)) > 0 Then
Text2 = DataCombo1(4)
End If
End If
End Select
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub dataCombo1_LostFocus(Index As Integer)
Select Case Index
       Case 4
If Option1(0).value = True Then
If Len(DataCombo1(4)) > 0 Then
Adodc12.RecordSource = "select ys from khy where sh='" & DataCombo1(4) & "'"
Adodc12.Refresh
If Not Adodc12.Recordset.EOF Then
DataCombo1(3) = Adodc12.Recordset.Fields(0)
End If
End If
End If
End Select
End Sub

Private Sub Form_Load()

On Error Resume Next
Dim L As String

cdbhf = cdbh
DTPicker1.value = Date - 30
DTPicker2.value = Date
Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT RQ AS ����,KH AS �ͻ�,fz AS ���ո�����,YS AS ��ɫ,SH AS ɫ��,BL AS Ʒ��,DH AS ���,ZL AS �������,IP AS ԡ��,qr as ���,qs as �����,xs as ��ˮ�� FROM ZH WHERE RQ BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) ORDER BY DH DESC"
Adodc5.Refresh

shbh = 1 ''''''''''''''''''''ɫ�ű仯����
For i = 0 To Adodc5.Recordset.Fields.count - 1
DataCombo1(i).Text = ""
Next
DataCombo1(0).Text = Date
DataCombo1(7).Text = ""
Text2.Text = ""

Option1(0).value = True

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

If InStr(yhm, "db") > 0 Or InStr(yhm, "scy") > 0 Then
Command3.Visible = True
Else
Command3.Visible = False
End If

DataCombo1(6).Text = yhdm + "1"
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT ���� FROM dbPFDbh where ����='" & yhdm & "'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
L = Adodc1.Recordset.Fields(0)
DataCombo1(6).Text = yhdm + Trim(Val(L) + 1) '''''''''''''OK
Else
L = yhdm + "1"
End If

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select ��� from khZL where ip like '%'+'" & yhxx & "'+'%' group by ���"
Adodc2.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select MC from RSFS group by MC"
Adodc6.Refresh


Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT ���������� FROM GR GROUP BY ����������"
Adodc4.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Data10.DatabaseName = App.Path & "\AccessBase\db.mdb"

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Command5.Enabled = True
Command4.Enabled = False
Command3.Enabled = False

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 1200
VSFlexGrid1.ColWidth(5) = 1000
VSFlexGrid1.ColWidth(6) = 1500
VSFlexGrid1.ColWidth(7) = 1500
VSFlexGrid1.ColWidth(8) = 2000
VSFlexGrid1.ColWidth(9) = 1000
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(1) = 0
VSFlexGrid2.ColWidth(2) = 0
VSFlexGrid2.ColWidth(5) = 1500
VSFlexGrid2.ColWidth(8) = 0
DataCombo1(1).TabIndex = 0

If Len(yhdm) <> 1 Then
MsgBox ("����˻������ʽ����������")
Command1.Enabled = False
Command5.Enabled = False
Command3.Enabled = False
Command7.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
sql2 = "insert into yhcd(�û�,�˵�,���) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.WindowState = 2
Formm1.Adodc1.Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sql2 = "delete from yhcd where �û�='" & yhm & "' and ���='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.Adodc1.Refresh
End Sub

Private Sub Text1_Change()
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select PM from PM where pm like '%'+ '" & Text1.Text & "'+'%' group by PM"
Adodc3.Refresh
End Sub

Private Sub Text3_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select ��� from KHZL where ����  like '%'+'" & Text3 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'  group by ���"
Adodc2.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
shbh = 0 '''''''''''''''''''''''''''''ɫ�ű仯
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move rs - 1
For i = 0 To Adodc5.Recordset.Fields.count - 1
DataCombo1(i).Text = Adodc5.Recordset.Fields(i)
Next
Text4 = Adodc5.Recordset.Fields(11)
Command5.Enabled = False
Command4.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Text2_Change()
If shbh = 1 Then
       Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc5.RecordSource = "SELECT RQ AS ����,KH AS �ͻ�,fz AS ���ո�����,YS AS ��ɫ,SH AS ɫ��,BL AS Ʒ��,DH AS ���,ZL AS �������,IP AS ԡ��,qr as ���,qs as �����,xs as ��ˮ�� FROM ZH WHERE  SH='" & Text2.Text & "'  ORDER BY DH DESC"
       Adodc5.Refresh
End If
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 500
Next
End If
End Sub

Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
    If c <> 7 Then
    
        
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
                
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
    End If
End With
End Sub


Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    VSFlexGrid1.Text = Combo1111.Text
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub

Private Sub Combo1111_LostFocus()
On Error Resume Next
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move r - 1
Adodc5.Recordset.Fields(c - 1) = Combo1111.Text
Adodc5.Recordset.Update
Combo1111.Visible = False
VSFlexGrid4.SetFocus
End Sub

