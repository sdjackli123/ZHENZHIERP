VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formd47 
   BackColor       =   &H00C0E0FF&
   Caption         =   "������Ϣ"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   DrawWidth       =   4684
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      Height          =   1575
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      Height          =   495
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formd47.frx":0000
      Left            =   1200
      List            =   "Formd47.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   600
      Width           =   1580
   End
   Begin VB.TextBox Text1 
      Height          =   370
      Left            =   840
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   120
      Width           =   370
   End
   Begin VB.TextBox Text5 
      Height          =   370
      Left            =   4320
      TabIndex        =   25
      Text            =   "Text5"
      Top             =   120
      Width           =   490
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   8280
      TabIndex        =   22
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   8280
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   18
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ѯ����"
      Height          =   1335
      Left            =   10200
      TabIndex        =   11
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ʒ��"
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   28
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ͷ��"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ɫ��"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "��̨"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�ͻ�"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   5880
      Top             =   10200
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      Height          =   330
      Left            =   5880
      Top             =   10560
      Visible         =   0   'False
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
      Left            =   6240
      Top             =   10560
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
      Height          =   375
      Left            =   6960
      Top             =   10200
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
      Left            =   7320
      Top             =   10440
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formd47.frx":006E
      Height          =   4935
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   15255
      _cx             =   26908
      _cy             =   8705
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
      Cols            =   11
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ѯ"
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3375
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   422969345
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   422969345
      CurrentDate     =   36892
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Bindings        =   "Formd47.frx":0083
      Height          =   290
      Left            =   2880
      TabIndex        =   26
      Top             =   600
      Width           =   1930
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "����"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formd47.frx":0098
      Height          =   290
      Left            =   1200
      TabIndex        =   30
      Top             =   120
      Width           =   1570
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formd47.frx":00AD
      Height          =   2415
      Left            =   240
      TabIndex        =   32
      Top             =   6840
      Width           =   15255
      _cx             =   26908
      _cy             =   4260
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
      Cols            =   11
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
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ʒ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   2880
      TabIndex        =   27
      Top             =   120
      Width           =   1450
   End
   Begin VB.Label Label2 
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
      Index           =   3
      Left            =   7440
      TabIndex        =   23
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "��̨"
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
      Left            =   7440
      TabIndex        =   21
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Index           =   1
      Left            =   7440
      TabIndex        =   19
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   970
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   7
      Top             =   600
      Width           =   855
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
      Height          =   380
      Index           =   8
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   970
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
      Left            =   4920
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Formd47"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer
Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Command2_Click()
Call MXOutadodcToExcel(VSFlexGrid2, "ͶȾ��Ϣ����")
End Sub

Private Sub Command3_Click()
On Error Resume Next
sql1 = ""

If Check2(0).value = 1 Then
sql1 = sql1 + "��Ϣ like '%'+'" & Combo1.Text & "'+'%' and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "�ͻ� like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "���� like '%'+'" & Text2(1).Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "��̨ like '%'+'" & Text2(2).Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
t1 = Format(Trim(DTPicker3.value), "yyyy-MM-dd")
t2 = Format(Trim(DTPicker4.value), "yyyy-MM-dd")
sql1 = sql1 + "���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "���� like '%'+'" & Text2(3).Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "��ɫ like '%'+'" & Text2(0).Text & "'+'%' and "
End If


If Check2(7).value = 1 Then
sql1 = sql1 + "Ʒ��='" & DataCombo6 & "' and "
End If


If sql1 = "" Then
MsgBox ("��ѡ���ѯ����")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)


Adodc4.RecordSource = "SELECT �ͻ�,����,Ʒ��,��ɫ,ɫ��,����,����,��̨,���� as ��������,��Ϣ,���,���� FROM pld where (" + sql1 + ") order BY ���� DESC,��� desc"
Adodc4.Refresh

Adodc2.RecordSource = "SELECT * FROM v_pld_fx_hz_ok where ���� between cast('" & DTPicker3.value & "' as datetime) and cast('" & DTPicker4.value & "' as datetime) ORDER BY ���� DESC"
Adodc2.Refresh

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 600
Next
End If

End Sub



Private Sub Command4_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "ͶȾ���ܱ���")
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Caption = Me.Caption + "������� " + ljb
For i = 0 To 3
Text2(i).Text = ""
Next
'Combo1.Text = ""
DataCombo2.Text = ""
DTPicker3.value = Date - 7
DTPicker4.value = Date
'Combo2.Text = "���"
Check2(4).value = 1
Text1 = ""
Text5 = ""
DataCombo6 = ""
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT * FROM v_pld_fx_hz_ok where ���� between cast('" & DTPicker3.value & "' as datetime) and cast('" & DTPicker4.value & "' as datetime) ORDER BY ���� DESC"
Adodc2.Refresh

Adodc4.CommandTimeout = 10000
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT �ͻ�,����,Ʒ��,��ɫ,ɫ��,����,����,��̨,���� as ��������,��Ϣ,���,����,����Ա,��������  FROM v_pld_plcl_cx where ���� BETWEEN '" & DTPicker3 & "' AND '" & DTPicker4 & "'  ORDER BY ���� DESC,��� desc"
Adodc4.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select ��� from khZL where ip like '%'+'" & yhxx & "'+'%' and ���� like '%'+'" & Text1 & "'+'%' group by ���"
Adodc5.Refresh


VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid2.ColWidth(0) = 400
For i = 1 To 11
VSFlexGrid2.ColWidth(i) = 1100
Next
VSFlexGrid2.ColWidth(12) = 1500
VSFlexGrid2.BackColorAlternate = &HCDEEC6
VSFlexGrid2.SelectionMode = flexSelectionListBox
End Sub


Private Sub Text1_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select ��� from khZL where ip like '%'+'" & yhxx & "'+'%' and ���� like '%'+'" & Text1 & "'+'%' group by ���"
Adodc5.Refresh
End Sub


Private Sub Text5_Change()
If Text5 = "" Then Exit Sub
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select distinct Ʒ�� as ����  from v_sczy_x_pmb where ����  like '%'+'" & Text5 & "'+'%' and �ͻ� like '%'+'" & DataCombo2 & "' +'%' order by Ʒ��"
Adodc5.Refresh
End Sub
