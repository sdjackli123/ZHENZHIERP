VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formd331 
   BackColor       =   &H00C0E0FF&
   Caption         =   "������ϵ�"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Formd331.frx":0000
      Left            =   13560
      List            =   "Formd331.frx":005E
      TabIndex        =   90
      Text            =   "Combo1"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   24480
      TabIndex        =   87
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Formd331.frx":00EF
      Height          =   1815
      Left            =   21960
      TabIndex        =   85
      Top             =   2640
      Width           =   3015
      _cx             =   5318
      _cy             =   3201
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
      FormatString    =   $"Formd331.frx":0105
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "������ʽ"
      Height          =   735
      Left            =   14280
      TabIndex        =   81
      Top             =   2040
      Width           =   4575
      Begin VB.OptionButton Option1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   83
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   82
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   2880
      Style           =   1  'Simple Combo
      TabIndex        =   80
      Text            =   "Combo1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ģ��ȷ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   74
      Text            =   "Text8"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����䷽"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18360
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   14520
      TabIndex        =   72
      Text            =   "Text15"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   70
      Text            =   "Text14"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1920
      TabIndex        =   69
      Text            =   "Text13"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   0
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   13440
      TabIndex        =   65
      Text            =   "Text12"
      Top             =   10200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   24840
      TabIndex        =   63
      Top             =   2280
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1440
      TabIndex        =   61
      Text            =   "Text9"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formd331.frx":01DB
      Height          =   450
      Left            =   8760
      TabIndex        =   54
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "��������"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8DAOCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formd331.frx":01F0
      Height          =   9615
      Left            =   480
      TabIndex        =   52
      Top             =   4560
      Width           =   22815
      _cx             =   40243
      _cy             =   16960
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
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
      Begin MSAdodcLib.Adodc Adodc32 
         Height          =   375
         Left            =   12240
         Top             =   4920
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Adodc32"
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
      Begin MSAdodcLib.Adodc Adodc31 
         Height          =   375
         Left            =   13800
         Top             =   4080
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
         Caption         =   "Adodc31"
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
      Begin MSAdodcLib.Adodc Adodc30 
         Height          =   615
         Left            =   13440
         Top             =   7920
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1085
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
         Caption         =   "Adodc30"
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
      Begin MSAdodcLib.Adodc Adodc29 
         Height          =   615
         Left            =   13320
         Top             =   6960
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1085
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
         Caption         =   "Adodc29"
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
   Begin MSAdodcLib.Adodc Adodc28 
      Height          =   330
      Left            =   8160
      Top             =   11640
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
      Caption         =   "Adodc28"
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
   Begin MSAdodcLib.Adodc Adodc27 
      Height          =   375
      Left            =   7080
      Top             =   11280
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Adodc27"
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
   Begin MSAdodcLib.Adodc Adodc26 
      Height          =   495
      Left            =   7440
      Top             =   10320
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
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
      Caption         =   "Adodc26"
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
   Begin MSAdodcLib.Adodc Adodc25 
      Height          =   330
      Left            =   7440
      Top             =   10440
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
      Caption         =   "Adodc25"
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
   Begin MSAdodcLib.Adodc Adodc24 
      Height          =   330
      Left            =   6840
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
      Caption         =   "Adodc24"
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
   Begin MSAdodcLib.Adodc Adodc23 
      Height          =   330
      Left            =   6720
      Top             =   10440
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
      Caption         =   "Adodc23"
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
   Begin MSAdodcLib.Adodc Adodc22 
      Height          =   330
      Left            =   6600
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
      Caption         =   "Adodc22"
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
   Begin MSAdodcLib.Adodc Adodc21 
      Height          =   330
      Left            =   8400
      Top             =   11160
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
      Caption         =   "Adodc21"
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
   Begin MSAdodcLib.Adodc Adodc20 
      Height          =   330
      Left            =   9000
      Top             =   10680
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
      Caption         =   "Adodc20"
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
   Begin MSAdodcLib.Adodc Adodc19 
      Height          =   330
      Left            =   8280
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
      Caption         =   "Adodc19"
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
   Begin MSAdodcLib.Adodc Adodc18 
      Height          =   375
      Left            =   10200
      Top             =   10080
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
      Caption         =   "Adodc18"
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
   Begin MSAdodcLib.Adodc Adodc17 
      Height          =   330
      Left            =   9960
      Top             =   10440
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
      Caption         =   "Adodc17"
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
   Begin MSAdodcLib.Adodc Adodc16 
      Height          =   330
      Left            =   9000
      Top             =   10560
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
      Caption         =   "Adodc16"
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
   Begin MSAdodcLib.Adodc Adodc15 
      Height          =   330
      Left            =   9600
      Top             =   11400
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
      Caption         =   "Adodc15"
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
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   330
      Left            =   5640
      Top             =   11520
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
      Caption         =   "Adodc14"
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
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   330
      Left            =   6000
      Top             =   11280
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "Adodc13"
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
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   495
      Left            =   6240
      Top             =   10320
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      Caption         =   "Adodc12"
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   5880
      Top             =   10440
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
      Left            =   6240
      Top             =   11520
      Visible         =   0   'False
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
      Left            =   6360
      Top             =   10560
      Visible         =   0   'False
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
      Height          =   375
      Left            =   6720
      Top             =   10320
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
      Height          =   375
      Left            =   6840
      Top             =   10440
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
      Left            =   7200
      Top             =   10440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   6360
      Top             =   11160
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
      Left            =   6120
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
      Left            =   7200
      Top             =   10680
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
      Left            =   7320
      Top             =   10320
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
      Left            =   7440
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
   Begin MSDataListLib.DataCombo DataCombo9 
      Height          =   450
      Index           =   0
      Left            =   10800
      TabIndex        =   51
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   794
      _Version        =   393216
      Text            =   "DataCombo3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo20 
      Bindings        =   "Formd331.frx":0205
      Height          =   450
      Left            =   17640
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "����������"
      Text            =   "DataCombo3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   10800
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   49
      Text            =   "Formd331.frx":021B
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ϵ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   20040
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ת�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   20040
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   20640
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   15120
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H008080FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   20640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   9360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�������ϵ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   2655
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   10560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   10560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
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
      Width           =   3015
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
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
      Top             =   10560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18360
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   0
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
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
      Top             =   10560
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data15 
      Caption         =   "Data15"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data16 
      Caption         =   "Data16"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data17 
      Caption         =   "Data17"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data18 
      Caption         =   "Data18"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   -240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data19 
      Caption         =   "Data19"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9840
      Top             =   0
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Formd331.frx":0221
      Left            =   600
      List            =   "Formd331.frx":0252
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3240
      Width           =   2535
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Formd331.frx":02CD
      Left            =   17640
      List            =   "Formd331.frx":02DA
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ϵ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16080
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13800
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�䷽����"
      Height          =   375
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Data Data20 
      Caption         =   "Data20"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data21 
      Caption         =   "Data21"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data22 
      Caption         =   "Data22"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data23 
      Caption         =   "Data23"
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data24 
      Caption         =   "Data24"
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
      Top             =   10560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data25 
      Caption         =   "Data25"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data26 
      Caption         =   "Data26"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   17
      Text            =   "Text8"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ת�����"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data27 
      Caption         =   "Data27"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Text            =   "Text10"
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data Data28 
      Caption         =   "Data28"
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDataListLib.DataCombo DataCombo9 
      Bindings        =   "Formd331.frx":02F0
      Height          =   450
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo9 
      Height          =   450
      Index           =   2
      Left            =   15240
      TabIndex        =   6
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   794
      _Version        =   393216
      Text            =   "DataCombo3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo9 
      Bindings        =   "Formd331.frx":0305
      Height          =   450
      Index           =   3
      Left            =   11520
      TabIndex        =   4
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "YS"
      Text            =   "DataCombo3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo9 
      Bindings        =   "Formd331.frx":031A
      Height          =   450
      Index           =   5
      Left            =   6120
      TabIndex        =   2
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "pm"
      Text            =   "DataCombo3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8DAOCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formd331.frx":032F
      Height          =   1455
      Left            =   360
      TabIndex        =   53
      Top             =   9120
      Visible         =   0   'False
      Width           =   14895
      _cx             =   26273
      _cy             =   2566
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
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formd331.frx":0343
      Height          =   450
      Left            =   15120
      TabIndex        =   71
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "��̨���"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Bindings        =   "Formd331.frx":0359
      Height          =   450
      Left            =   1920
      TabIndex        =   77
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "ģ����"
      Text            =   "DataCombo7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ͶȾ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label24 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   89
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label23 
      BackColor       =   &H0080FFFF&
      Caption         =   "�Ÿ׼ƻ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   360
      TabIndex        =   88
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�пع���"
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
      Left            =   23520
      TabIndex        =   86
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ⱦɫ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   20640
      TabIndex        =   84
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ģ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   600
      TabIndex        =   78
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackColor       =   &H0000C0C0&
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
      Height          =   495
      Left            =   11040
      TabIndex        =   76
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label18 
      BackColor       =   &H0000C0C0&
      Caption         =   "ԡ�ȵ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   75
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�������"
      Height          =   375
      Left            =   11640
      TabIndex        =   68
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�Ÿ׿�"
      Height          =   375
      Left            =   11640
      TabIndex        =   67
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ⱦɫ����"
      Height          =   375
      Left            =   11640
      TabIndex        =   66
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "���ű��"
      Height          =   495
      Left            =   12720
      TabIndex        =   64
      Top             =   10200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����ֵ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   23880
      TabIndex        =   62
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ⱦɫ����"
      Height          =   375
      Left            =   480
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF80&
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   58
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "��������"
      Height          =   375
      Left            =   8760
      TabIndex        =   57
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ⱦɫ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   19080
      TabIndex        =   56
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ɫ����"
      Height          =   375
      Left            =   10800
      TabIndex        =   55
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ƥ  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   9600
      TabIndex        =   50
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ע��Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   5520
      TabIndex        =   48
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label55 
      BackColor       =   &H00C0FFC0&
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   19080
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label49 
      BackColor       =   &H008080FF&
      Caption         =   "�����䷽"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   19080
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�ͻ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   46
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   45
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����(kg)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   13920
      TabIndex        =   44
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��  ɫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   10440
      TabIndex        =   43
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ʒ ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4560
      TabIndex        =   42
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ɫ ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   5520
      TabIndex        =   41
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��̨"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   40
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label20 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   16680
      TabIndex        =   39
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label51 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����Ա"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   16680
      TabIndex        =   38
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label20 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�䷽��ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   37
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label20 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   2
      Left            =   8760
      TabIndex        =   36
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ϵ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   35
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "������Ϣ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   600
      TabIndex        =   34
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "1��У��ֵΪ0.1  0.5��У��ֵΪ0.05"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   33
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ȫ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11760
      TabIndex        =   32
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000C0C0&
      Caption         =   "ԡ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   31
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000C0C0&
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
      Height          =   495
      Left            =   7560
      TabIndex        =   30
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ϵ�ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   29
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "Formd331"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r, JDBAR As Integer: Public rhl As String: Dim sz(61) As String:: Dim pfsz(6) As String: Dim pfdsz(6) As String
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Dim ba55 As Database: Dim rd55 As Recordset: Dim ba15 As Database: Dim rd15 As Recordset: Dim BA2 As Database: Dim RD2 As Recordset: Dim ba5 As Database: Dim rd5 As Recordset: Dim ba6 As Database: Dim rd6 As Recordset: Dim rd11 As Recordset: Dim ba11 As Database: Dim rd12 As Recordset: Dim ba12 As Database: Dim BA13 As Database: Dim rd13 As Recordset
Dim ZS(10) As String: Dim ysl As Single: Dim yqz As Single
Dim plshsx As Integer
Dim cdbhf As Integer
''''''''''''''''''''''''''''''''''''''''''''''''
Dim sb As RECT
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long




Private Sub Combo2_Change()
If InStr(Combo2.Text, "����") > 0 Or InStr(Combo2.Text, "����") > 0 Then
Label6.Visible = False
Text11 = ""
Else
Label6.Visible = True
Text11 = "0.01"
End If
End Sub

Private Sub Combo2_Click()
If InStr(Combo2.Text, "����") > 0 Or InStr(Combo2.Text, "����") > 0 Then
Label6.Visible = False
Text11 = ""
Else
Label6.Visible = True
Text11 = "0.01"
End If
End Sub

Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Command1_Click()
On Error Resume Next
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 400
VSFlexGrid1.ColWidth(7) = 6000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0
VSFlexGrid1.ColFormat(11) = "#0"

If Data13.Recordset.EOF Then
MsgBox ("�������ϵ�������")
Exit Sub
End If

Text5.Enabled = False
Call scpfd(Text2.Text)
Call clxt(Text2.Text)

sql2 = "insert into czrz(����,����,����,����,����) VALUES('" & Now & "','" & Text5.Text & "','" & yhm & "','�ϵ���ӡ','��������')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

ProgressBar1.Visible = True   ''��ʾ������
Timer2.Enabled = True
End Sub

Private Sub Command10_Click()
On Error Resume Next
If Text1.Text = "" Then
MsgBox ("�������䷽���")
Exit Sub
End If

Data27.Database.Execute "delete * from pfda"

Adodc24.RecordSource = "select * from pfd where  ���='" & Text1.Text & "'"
Adodc24.Refresh
If Not Adodc24.Recordset.EOF Then
Adodc24.Recordset.MoveFirst

For i = 0 To 6
pfdsz(i) = Adodc24.Recordset.Fields(i)
Next

mb = 0
For i = 7 To 56
If Adodc24.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next
ProgressBar1.Visible = True
For i = 7 To mb + 7
If Adodc24.Recordset.Fields(i) <> "" Then
pfsz(0) = Mid(Adodc24.Recordset.Fields(i), 1, InStr(Adodc24.Recordset.Fields(i), "(") - 1)
pfsz(1) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "(") + 1, InStr(Adodc24.Recordset.Fields(i), ")") - InStr(Adodc24.Recordset.Fields(i), "(") - 1)
pfsz(2) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), ")") + 1, InStr(Adodc24.Recordset.Fields(i), "-") - InStr(Adodc24.Recordset.Fields(i), ")") - 1)
pfsz(3) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "-") + 1, InStr(Adodc24.Recordset.Fields(i), "\") - InStr(Adodc24.Recordset.Fields(i), "-") - 1)
pfsz(4) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "\") + 1, InStr(Adodc24.Recordset.Fields(i), "#") - InStr(Adodc24.Recordset.Fields(i), "\") - 1)
pfsz(5) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "#") + 1, InStr(Adodc24.Recordset.Fields(i), "^") - InStr(Adodc24.Recordset.Fields(i), "#") - 1)
pfsz(6) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "^") + 1)
L = i - 6
Data27.Database.Execute "insert into pfda(�ӹ���λ,Ʒ��,ɫ��,��ɫ,�䷽���,������,�䷽����,��������,ԡ��,Ⱦ������,Ⱦ��������,��λ,�䷽,����,�����) VALUES('" & pfdsz(0) & "','" & pfdsz(1) & "','" & pfdsz(2) & "','" & pfdsz(3) & "','" & pfdsz(4) & "','" & pfdsz(5) & "',CDATE('" & pfdsz(6) & "'),'" & pfsz(0) & "','" & pfsz(1) & "','" & pfsz(2) & "',  " & _
                                                                        "'" & pfsz(3) & "','" & pfsz(4) & "','" & pfsz(5) & "','" & pfsz(6) & "','" & L & "')"
ProgressBar1.value = 100 / mb * L
End If
Next
ProgressBar1.Visible = False
End If
Formh233.DataCombo1(12).Text = Text1.Text
'Form111111.dataCombo1(14).Text = Text7.Text
End Sub

Private Sub Command11_Click()
On Error Resume Next
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh
End Sub

Private Sub Command12_Click()
Data13.Database.Execute "delete * from pldd"
Adodc21.Refresh
End Sub

Private Sub Command13_Click()
On Error Resume Next

If Text2.Text = "" Then Exit Sub
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "DBPLDSH('" & Text2.Text & "','ȷ��','" & Date & "')"    ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
g_Cmd.Cancel
       
       
      adwww6.DE = ""
       adwww6.W12 = ""
       adwww6.W11 = ""
       adwww6.a = ""  '����
       adwww6.b = ""  'ע1"
       adwww6.c = ""  'ע2
       adwww6.d = ""  'ע3
       adwww6.e = ""  '��������
       adwww6.F = "" '����
       adwww6.g = "" 'ԡ��
       adwww6.h = ""  'ˮ��
       adwww6.i = "" '�䷽��
       adwww6.i2 = "" '����
       adwww6.i3 = ""
       adwww6.i4 = ""
       adwww6.u = ""
       adwww6.x = ""
For i% = 1 To 15
adwww6.j(i%) = ""
adwww6.k(i%) = ""
adwww6.L(i%) = ""
adwww6.l1(i%) = ""
adwww6.m(i%) = ""
adwww6.n(i%) = ""
adwww6.o(i%) = ""
Next



Data27.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' and Ⱦ������<>'������' and ѹ��='y' order by ��������,�����"
Data27.Refresh
If Data27.Recordset.EOF Then
Data27.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' and Ⱦ������<>'������' order by ��������,�����"
Data27.Refresh
If Data27.Recordset.EOF Then
Call scpfd(Text2.Text)
Call Command2_Click
Exit Sub
End If

Data27.Recordset.MoveFirst
'adwww6.A = Data27.Recordset.Fields(2) '
'adwww6.e = Data27.Recordset.Fields(0)
'adwww6.F = Data27.Recordset.Fields(1)
'adwww6.g = Data27.Recordset.Fields(4) '
'adwww6.h = Data27.Recordset.Fields(20)

       If Len(adwww6.DE) < 1 Then
       adwww6.DE = Space(1)
       Else
       adwww6.DE = Mid(adwww6.DE, 1, 1)
       End If
       
       If Len(adwww6.W12) < 12 Then
       adwww6.W12 = adwww6.W12 + Space(12 - Len(adwww6.W12))
       Else
       adwww6.W12 = Mid(adwww6.W12, 1, 12)
       End If
       
       adwww6.a = Data27.Recordset.Fields(2)   '����
       If Len(adwww6.a) < 30 Then
       adwww6.a = adwww6.a + Space(30 - Len(adwww6.a))
       Else
       adwww6.a = Mid(adwww6.a, 1, 30)
       End If
       
       adwww6.b = Space(6)  'ע1
       adwww6.c = Space(6)  'ע2
       adwww6.d = Space(6)  'ע3
       adwww6.e = Mid(Data27.Recordset.Fields(0), 5) '��������
       If Len(adwww6.e) < 30 Then
       adwww6.e = adwww6.e + Space(30 - Len(adwww6.e))
       Else
       adwww6.e = Mid(adwww6.e, 1, 30)
       End If
       
       adwww6.F = Data27.Recordset.Fields(1) '����
       If Len(adwww6.F) < 10 Then
       adwww6.F = adwww6.F + Space(10 - Len(adwww6.F))
       Else
       adwww6.F = Mid(adwww6.F, 1, 10)
       End If
       
       adwww6.g = Data27.Recordset.Fields(4) 'ԡ��
       If Len(adwww6.g) < 6 Then
       adwww6.g = adwww6.g + Space(6 - Len(adwww6.g))
       Else
       adwww6.g = Mid(adwww6.g, 1, 6)
       End If
       
       adwww6.h = Data27.Recordset.Fields(20)  'ˮ��
       If Len(adwww6.g) < 10 Then
       adwww6.h = adwww6.h + Space(10 - Len(adwww6.h))
       Else
       adwww6.h = Mid(adwww6.h, 1, 10)
       End If
       
       adwww6.i = Space(12)
       
       adwww6.i2 = Trim(Format(Date, "yymmdd")) '����
       adwww6.i3 = Space(1)
       adwww6.i4 = Space(1)
For i% = 1 To 15
If Not Data27.Recordset.EOF Then
Do While Not Data27.Recordset.EOF

       adwww6.j(i%) = Data27.Recordset.Fields(6)
       If Len(adwww6.j(i%)) < 12 Then
       adwww6.j(i%) = adwww6.j(i%) + Space(12 - Len(adwww6.j(i%)))
       Else
       adwww6.j(i%) = Mid(adwww6.j(i%), 1, 12)
       End If
       
       adwww6.k(i%) = Data27.Recordset.Fields(8)
       If Len(adwww6.k(i%)) < 8 Then
       adwww6.k(i%) = adwww6.k(i%) + Space(8 - Len(adwww6.k(i%)))
       Else
       adwww6.k(i%) = Mid(adwww6.k(i%), 1, 8)
       End If
       
       adwww6.L(i%) = Format(Data27.Recordset.Fields(10), "#0.0")
       If Len(adwww6.L(i%)) < 9 Then
       adwww6.L(i%) = adwww6.L(i%) + Space(9 - Len(adwww6.L(i%)))
       Else
       adwww6.L(i%) = Mid(adwww6.L(i%), 1, 9)
       End If
       
       adwww6.l1(i%) = "00000.000"
       
       adwww6.m(i%) = "D"
       adwww6.n(i%) = "%"
       adwww6.o(i%) = "100 "

Data27.Recordset.MoveNext
i% = i% + 1
Loop
Else
adwww6.j(i%) = Space(12)
adwww6.k(i%) = Space(8)
adwww6.L(i%) = Space(9)
adwww6.m(i%) = Space(1)
adwww6.n(i%) = Space(1)
adwww6.o(i%) = Space(4)
adwww6.l1(i%) = "00000.000"
End If

Next
      da$ = Format(Year(Now), "0000") + Format(Month(Now), "00")
      namep$ = "\\ad1\c\adcc\DAT3\G" + da$ + ".666"
      op1% = FreeFile: Open namep$ For Random As #op1% Len = Len(adwww6)
      n& = LOF(op1%) / Len(adwww6) + 1
      adwww6.W11 = Format(Trim(n&), "00000")
      Close #op1%
      adwww6.u = Format(Date, "mm") + Space(1) + Format(Hour(Now), "hh") + ":" + Format(Minute(Now), "mm")
           '( 05 08:33 )
           adwww6.x = Chr(13) + Chr(10)
           da$ = Format(Date, "yyyy") + Format(Date, "mm")
           
      Call bpww666(2, da$)
     
     '---------------------------------------------------------------------------------------------------------------------------------------

MsgBox ("ת��ɹ���")
Data27.Database.Execute "update pldd set ѹ��='y' where �ϵ����='" & Text2.Text & "'"
Call scpfd(Text2.Text)
Call Command2_Click
Exit Sub

Else
If MsgBox("�����ϵ���ת�룬�Ƿ�����ת�룿", vbYesNo) = vbNo Then Exit Sub

       If Len(adwww6.DE) < 1 Then
       adwww6.DE = Space(1)
       Else
       adwww6.DE = Mid(adwww6.DE, 1, 1)
       End If
       
       If Len(adwww6.W12) < 12 Then
       adwww6.W12 = adwww6.W12 + Space(12 - Len(adwww6.W12))
       Else
       adwww6.W12 = Mid(adwww6.W12, 1, 12)
       End If
       
       adwww6.a = Data27.Recordset.Fields(2)   '����
       If Len(adwww6.a) < 30 Then
       adwww6.a = adwww6.a + Space(30 - Len(adwww6.a))
       Else
       adwww6.a = Mid(adwww6.a, 1, 30)
       End If
       
       adwww6.b = Space(6)  'ע1
       adwww6.c = Space(6)  'ע2
       adwww6.d = Space(6)  'ע3
       adwww6.e = Mid(Data27.Recordset.Fields(0), 5) '��������
       If Len(adwww6.e) < 30 Then
       adwww6.e = adwww6.e + Space(30 - Len(adwww6.e))
       Else
       adwww6.e = Mid(adwww6.e, 1, 30)
       End If
       
       adwww6.F = Data27.Recordset.Fields(1) '����
       If Len(adwww6.F) < 10 Then
       adwww6.F = adwww6.F + Space(10 - Len(adwww6.F))
       Else
       adwww6.F = Mid(adwww6.F, 1, 10)
       End If
       
       adwww6.g = Data27.Recordset.Fields(4) 'ԡ��
       If Len(adwww6.g) < 6 Then
       adwww6.g = adwww6.g + Space(6 - Len(adwww6.g))
       Else
       adwww6.g = Mid(adwww6.g, 1, 6)
       End If
       
       adwww6.h = Data27.Recordset.Fields(20)  'ˮ��
       If Len(adwww6.g) < 10 Then
       adwww6.h = adwww6.h + Space(10 - Len(adwww6.h))
       Else
       adwww6.h = Mid(adwww6.h, 1, 10)
       End If
       
       adwww6.i = Space(12)
       
       adwww6.i2 = Trim(Format(Date, "yymmdd")) '����
       adwww6.i3 = Space(1)
       adwww6.i4 = Space(1)
For i% = 1 To 15
If Not Data27.Recordset.EOF Then
Do While Not Data27.Recordset.EOF

       adwww6.j(i%) = Data27.Recordset.Fields(6)
       If Len(adwww6.j(i%)) < 12 Then
       adwww6.j(i%) = adwww6.j(i%) + Space(12 - Len(adwww6.j(i%)))
       Else
       adwww6.j(i%) = Mid(adwww6.j(i%), 1, 12)
       End If
       
       adwww6.k(i%) = Data27.Recordset.Fields(8)
       If Len(adwww6.k(i%)) < 8 Then
       adwww6.k(i%) = adwww6.k(i%) + Space(8 - Len(adwww6.k(i%)))
       Else
       adwww6.k(i%) = Mid(adwww6.k(i%), 1, 8)
       End If
       
       adwww6.L(i%) = Format(Data27.Recordset.Fields(10), "#0.0")
       If Len(adwww6.L(i%)) < 9 Then
       adwww6.L(i%) = adwww6.L(i%) + Space(9 - Len(adwww6.L(i%)))
       Else
       adwww6.L(i%) = Mid(adwww6.L(i%), 1, 9)
       End If
       
       adwww6.l1(i%) = "00000.000"
       
       adwww6.m(i%) = "D"
       adwww6.n(i%) = "%"
       adwww6.o(i%) = "100 "

Data27.Recordset.MoveNext
i% = i% + 1
Loop
Else
adwww6.j(i%) = Space(12)
adwww6.k(i%) = Space(8)
adwww6.L(i%) = Space(9)
adwww6.m(i%) = Space(1)
adwww6.n(i%) = Space(1)
adwww6.o(i%) = Space(4)
adwww6.l1(i%) = "00000.000"
End If

Next
      da$ = Format(Year(Now), "0000") + Format(Month(Now), "00")
      namep$ = "\\ad1\c\adcc\DAT3\G" + da$ + ".666"
      op1% = FreeFile: Open namep$ For Random As #op1% Len = Len(adwww6)
      n& = LOF(op1%) / Len(adwww6) + 1
      adwww6.W11 = Format(Trim(n&), "00000")
      Close #op1%
      adwww6.u = Format(Date, "mm") + Space(1) + Format(Hour(Now), "hh") + ":" + Format(Minute(Now), "mm")
           '( 05 08:33 )
           adwww6.x = Chr(13) + Chr(10)
           da$ = Format(Date, "yyyy") + Format(Date, "mm")
           
      Call bpww666(2, da$)
     
     '---------------------------------------------------------------------------------------------------------------------------------------

MsgBox ("ת��ɹ���")
Call scpfd(Text2.Text)
Call Command2_Click
Exit Sub
End If

End Sub

Private Sub Command14_Click()
On Error Resume Next
If MsgBox("ȷ�����������䷽��", vbYesNo) = vbNo Then Exit Sub
Data13.Recordset.MoveFirst
Do While Not Data13.Recordset.EOF
Data13.Recordset.Edit
If Trim(Data13.Recordset.Fields(5)) = "ӡ����" And Val(Data13.Recordset.Fields(20)) > 0 Then
Data13.Recordset.Fields(7) = "g/l"
Data13.Recordset.Fields(8) = Format(Val(Data13.Recordset.Fields(10)) / Val(Data13.Recordset.Fields(20)), "#0.00000")
End If
Data13.Recordset.Update
Data13.Recordset.MoveNext
Loop

Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 400
VSFlexGrid1.ColWidth(7) = 6000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0

End Sub

Private Sub Command15_Click()
'On Error Resume Next
If MsgBox("����ģ�� " + DataCombo7 + " �������ϵ���", vbYesNo) = vbNo Then Exit Sub
If DataCombo7 = "" Then
MsgBox ("��ѡ��ģ��!")
Exit Sub
End If
Adodc2.RecordSource = "select * from CGGYMB where ģ����='" & DataCombo7 & "'"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
pf = Format(Adodc2.Recordset.Fields(6), "#0.00000")
Data7.Database.Execute "insert into pldd(�ϵ����,��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,�����,����,����) values('" & Text2 & "','" & Adodc2.Recordset.Fields(0) & "','" & Text16 & "','" & Adodc2.Recordset.Fields(2) & "','" & Adodc2.Recordset.Fields(4) & "','" & Adodc2.Recordset.Fields(5) & "','" & pf & "','1','" & Adodc2.Recordset.Fields(7) & "','','" & Adodc2.Recordset.Fields(8) & "')"
Adodc2.Recordset.MoveNext
Loop
End If

Call Label14_Click

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Len(Text5) < 7 Then
MsgBox ("������Ų�����")
Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�޸�״̬
Adodc12.RecordSource = "SELECT * FROM PLD WHERE ���='" & Text2 & "'"
Adodc12.Refresh
If Adodc12.Recordset.EOF Then
MsgBox ("û������������Ϣ �����ϵ��䷽������ ���ܴ�������")
Exit Sub
End If

Adodc12.RecordSource = "SELECT * FROM pldr WHERE �ϵ����='" & Text2 & "'"
Adodc12.Refresh
If Adodc12.Recordset.EOF Then
MsgBox ("û��ת�� ����ϵͳ �����ϵ��䷽������ ���ܴ�������")
Exit Sub
End If

Adodc12.RecordSource = "SELECT * FROM ghgx WHERE ����='" & Text5 & "' and ���� between '1001' and '6000'"   ''������Ⱦɫ������ת��
Adodc12.Refresh
If Adodc12.Recordset.EOF Then
MsgBox ("û������Ⱦɫ���� ������Ⱦɫ����")
Exit Sub
End If

sql1 = "UPDATE pld SET ��ӡ='1' WHERE ���='" & Text2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

sql2 = "insert into czrz(����,����,����,����,����) VALUES('" & Now & "','" & Text5.Text & "','" & yhm & "','ת�����','��������')"  ''������־
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM DBPLDBH where ����='" & yhdm & "'"
Adodc1.Refresh

Text2.Text = yhdm + "1" ''''''''''''OK
If Adodc1.Recordset.EOF Then
Text2.Text = yhdm + "1" ''''''''''''OK
Else
L = Val(Adodc1.Recordset.Fields(0))
Text2.Text = yhdm + Trim(L + 1) '''''''''''''OK
End If

Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text5.Enabled = True
For i = 1 To 5
If i = 4 Then i = 5
DataCombo9(i).Text = ""
Next
Text6.Text = ""
Text7.Text = ""
Text5.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Command4_Click()
On Error Resume Next

If Text6.Text = "" Then
MsgBox ("������ɫ��")
Exit Sub
End If

If Trim(Text4.Text) = "" Then
MsgBox ("�������̨")
Exit Sub
End If

If Text5.Text = "" Then
MsgBox ("���������")
Exit Sub
End If

If DataCombo9(1).Text = "" Then
MsgBox ("������ͻ�����")
Exit Sub
End If

Adodc32.RecordSource = "select ���׹��� from bgxx where ���׹���='" & Text5 & "'"
Adodc32.Refresh
If Not Adodc32.Recordset.EOF Then
MsgBox ("ע�⣺�����Ѳ���")
End If


Adodc13.RecordSource = "select ���,��Ϣ from pld where ����='" & Text5 & "'"
Adodc13.Refresh

Data17.RecordSource = "SELECT distinct Ⱦ�������� FROM pldd where �ϵ����='" & Text2.Text & "' "
Data17.Refresh
If Not Data17.Recordset.EOF Then
Data17.Recordset.MoveFirst
Do While Not Data17.Recordset.EOF
Adodc17.RecordSource = "select * from rhzh where Ⱦ������='" & Data17.Recordset.Fields(0) & "' and ��־='��'"
Adodc17.Refresh
If Adodc17.Recordset.EOF Then
MsgBox ("������" + Data17.Recordset.Fields(0))
Exit Sub
End If
Data17.Recordset.MoveNext
Loop
End If

If Combo2 = "��������" Then
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text5.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("jc", adChar, adParamOutput, 1)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "bgjc"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel


If Val(g_Cmd.Parameters("jc").value) = 1 Then
MsgBox ("��ֹ�˹���������Ϣ  �������� ��Σ�ϵͳֻ����һ�Σ����д˲�����Ϣ  ")
Exit Sub
End If
End If

If Option1(0).value = True Then
If Combo2 = "����" Then
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text5.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("jc", adChar, adParamOutput, 1)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "bzcjc"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel


If Val(g_Cmd.Parameters("jc").value) = 1 Then
MsgBox ("��ֹ�˹���������Ϣ  ���� ��Σ�ϵͳֻ����һ�Σ����д˹���Ϣ  ")
Exit Sub
End If
End If
End If

If Option1(1).value = True Then
Adodc4.RecordSource = "select ��Ϣ from pld where ���='" & Text2 & "'"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
MsgBox ("�����ڴ��ϵ���Ϣ")
Exit Sub
Else
Combo1 = Adodc4.Recordset.Fields(0)
End If
End If

If InStr(Combo2, "����") > 0 Then
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text5.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("jc", adChar, adParamOutput, 1)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "btmjc"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel


If Val(g_Cmd.Parameters("jc").value) = 1 Then
MsgBox ("��ֹ�˹���������Ϣ  Ⱦ���� ��Σ�ϵͳֻ����һ�Σ����д˹���Ϣ ��û�л���Ⱦɫ���ղ������ޣ�")
Exit Sub
End If
End If

If InStr(Combo2, "����") > 0 Then
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text5.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("jc", adChar, adParamOutput, 1)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "btmjc2"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel


If Val(g_Cmd.Parameters("jc").value) = 1 Then
MsgBox ("��ֹ�˹���������Ϣ  Ⱦ���� ��Σ�ϵͳֻ����һ�Σ����д˹���Ϣ ��û�л���Ⱦɫ���ղ������ޣ�")
Exit Sub
End If
End If

''''''''''''''''''''''������Ϣ¼��
If Option1(0).value = True Then
If Combo2.Text = "����" Then
Data27.Database.Execute "UPDATE pldd SET ����='" & Text5.Text & "',����='" & DataCombo9(2).Text & "',������Ϣ='" & Combo2.Text & "',���ϴ�ӡԱ='" & Combo3 & "'+'/'+'" & DataCombo20.Text & "',��������='" & Now & "',�䷽��='" & Text6.Text & "',�������='" & DataCombo9(3).Text & "',ѹ��='" & DataCombo9(5).Text & "',���='" & DataCombo9(1).Text & "',Ⱦ��������='" & Text4.Text & "',���ȷ��='" & Text7.Text & "' WHERE �ϵ����='" & Text2.Text & "'"
Else
Data27.Database.Execute "UPDATE pldd SET ����='" & Text5.Text & "',����='" & DataCombo9(2).Text & "',������Ϣ='" & Combo2.Text & "',���ϴ�ӡԱ='" & Combo3 & "'+'/'+'" & DataCombo20.Text & "',��������='" & Now & "',�䷽��='" & Text6.Text & "',�������='" & DataCombo9(3).Text & "',ѹ��='" & DataCombo9(5).Text & "',���='" & DataCombo9(1).Text & "',Ⱦ��������='" & Text4.Text & "',���ȷ��='" & Text7.Text & "' WHERE �ϵ����='" & Text2.Text & "'"
End If
End If

If Option1(1).value = True Then
If Combo2.Text = "����" Then
Data27.Database.Execute "UPDATE pldd SET ����='" & Text5.Text & "',����='" & DataCombo9(2).Text & "',������Ϣ='" & Combo2.Text & "',���ϴ�ӡԱ='" & Combo3 & "'+'/'+'" & DataCombo20.Text & "',�䷽��='" & Text6.Text & "',�������='" & DataCombo9(3).Text & "',ѹ��='" & DataCombo9(5).Text & "',���='" & DataCombo9(1).Text & "',Ⱦ��������='" & Text4.Text & "',���ȷ��='" & Text7.Text & "' WHERE �ϵ����='" & Text2.Text & "'"
Else
Data27.Database.Execute "UPDATE pldd SET ����='" & Text5.Text & "',����='" & DataCombo9(2).Text & "',������Ϣ='" & Combo2.Text & "',���ϴ�ӡԱ='" & Combo3 & "'+'/'+'" & DataCombo20.Text & "',�䷽��='" & Text6.Text & "',�������='" & DataCombo9(3).Text & "',ѹ��='" & DataCombo9(5).Text & "',���='" & DataCombo9(1).Text & "',Ⱦ��������='" & Text4.Text & "',���ȷ��='" & Text7.Text & "' WHERE �ϵ����='" & Text2.Text & "'"
End If
End If
Data27.Database.Execute "delete * from pldd WHERE �ϵ����='" & Text2.Text & "' and trim(�䷽)='0'"

'''''''''''''''''''''Ⱦ��������
Data13.Recordset.MoveFirst
Do While Not Data13.Recordset.EOF
Data13.Recordset.Edit

If Data13.Recordset.Fields(7) = "%" Or Data13.Recordset.Fields(7) = "g/l" Then

If Trim(Data13.Recordset.Fields(5)) <> "����" Then
k = 10
Data13.Recordset.Fields(11) = "g"
Data13.Recordset.Fields(10) = Format(Val(Data13.Recordset.Fields(9)) * Val(Data13.Recordset.Fields(8)) * Val(Data13.Recordset.Fields(1)) * k, "#0.0")
End If

If Trim(Data13.Recordset.Fields(5)) = "����" Then
k = Data13.Recordset.Fields(4)      ''''''''''ԡ��
Data13.Recordset.Fields(11) = "kg"
If Val(Data13.Recordset.Fields(1)) > 10 Then
Data13.Recordset.Fields(10) = Format(Val(Data13.Recordset.Fields(9)) * Val(Data13.Recordset.Fields(8)) * Val(Data13.Recordset.Fields(1)) * k / 1000, "#0.0")
Else
Data13.Recordset.Fields(10) = Format(Val(Data13.Recordset.Fields(9)) * Val(Data13.Recordset.Fields(8)) * 10 * k / 1000, "#0.0")
End If
End If


Else

Data13.Recordset.Fields(10) = Format(Val(Data13.Recordset.Fields(8)) * Val(Data13.Recordset.Fields(9)), "#0.0")
Data13.Recordset.Fields(11) = Data13.Recordset.Fields(7)
End If

Data13.Recordset.Fields(20) = Format(Val(Data13.Recordset.Fields(1)) * Val(Data13.Recordset.Fields(4)), "#0.0")
Data13.Recordset.Update
Data13.Recordset.MoveNext
Loop

Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh


If Combo2 = "��������" Then
sql1 = "UPDATE KPD SET kp=convert(nvarchar(120),getdate(),120),zt='Ⱦɫ��',gz=convert(nvarchar ,'" & Now & "',120) WHERE ���� in(select ���׹��� from bgxx where ���ϱ�� ='" & Text2.Text & "') and isnull(zt.'')<>'��������'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
If Combo2 = "����" Then
sql1 = "UPDATE KPD SET kp=convert(nvarchar(120),getdate(),120),zt='Ⱦɫ��',gz=convert(nvarchar ,'" & Now & "',120) WHERE ����='" & Text5.Text & "' and isnull(zt,'')<>'��������'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
If Combo2 = "����" Then
sql1 = "UPDATE KPD SET kp=convert(nvarchar(120),getdate(),120),zt='Ⱦɫ��',gz=convert(nvarchar ,'" & Now & "',120) WHERE ����='" & Text5.Text & "' and isnull(zt,'')<>'��������'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
If Combo2 = "��������" Then
sql1 = "UPDATE KPD SET kp=convert(nvarchar(120),getdate(),120),zt='Ⱦɫ��',gz=convert(nvarchar ,'" & Now & "',120) WHERE ���� in(select ���׹��� from bgxx where ���ϱ�� ='" & Text2.Text & "') and isnull(zt,'')<>'������������'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh


If Data13.Recordset.EOF Then
MsgBox ("���ղ�����")
Exit Sub
End If

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 600
VSFlexGrid1.ColWidth(7) = 6000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0
VSFlexGrid1.ColFormat(11) = "#0.0"

End Sub



Private Sub Command5_Click()
Formd337.Show
End Sub

Private Sub Command6_Click()
On Error Resume Next
If Text1.Text = "" Then
MsgBox ("��ѡ���յ�")
Exit Sub
End If

If DBCombo1.Text = "" Then
MsgBox ("��ѡ����")
Exit Sub
End If

If DataCombo9(2).Text = "" Or Text5.Text = "" Or Text1.Text = "" Or Text2.Text = "" Then
If MsgBox("���š��������䷽���š����ϵ�����д������", vbYesNo) = vbNo Then
Exit Sub
End If
End If

'Data22.Database.Execute "DELETE * FROM pldd WHERE �ϵ����='" & Text2.Text & "'"
If DBCombo1.Text = "ȫ��" Then
Data20.Database.Execute "INSERT INTO  pldd(��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,�����,����,����) SELECT ��������,ԡ��,Ⱦ������,Ⱦ��������,��λ,�䷽,У��ֵ,�����,����,���� From pfda WHERE �䷽���='" & Text1.Text & "' ORDER BY VAL(��������)"
Else
Data22.RecordSource = "select * from pldd where ��������='" & DataCombo1 & "'"
Data22.Refresh
If Not Data22.Recordset.EOF Then
MsgBox ("�˹����Ѿ����ڣ�")
Exit Sub
Else
Data20.Database.Execute "INSERT INTO  pldd(��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,�����,����,����) SELECT ��������,ԡ��,Ⱦ������,Ⱦ��������,��λ,�䷽,У��ֵ,�����,����,���� From pfda WHERE �䷽���='" & Text1.Text & "' AND ��������='" & DBCombo1.Text & "' "
End If
End If
Data20.Database.Execute "UPDATE pldd SET �ϵ����='" & Text2.Text & "',У��ֵ='1'  WHERE �ϵ����=NULL"
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 600
VSFlexGrid1.ColWidth(7) = 3000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0
End Sub


Private Sub Command7_Click()
On Error Resume Next

If MsgBox("ȷ��������", vbYesNo) = vbNo Then
Exit Sub
End If

'sql2 = "delete  from pldr where �ϵ����='" & Text1.Text & "'"
'sql3 = "delete  from pldb where �ϵ����='" & Text1.Text & "'"

'RD.Open sql2, conn, adOpenStatic, adLockOptimistic
'RD.Open sql3, conn, adOpenStatic, adLockOptimistic

Data22.Database.Execute "DELETE * FROM pldd WHERE �ϵ����='" & Text2.Text & "'"
Data13.Refresh
Call Label14_Click
End Sub

Private Sub Command8_Click()
On Error Resume Next
If Text5.Text = "" Or DataCombo9(2).Text = "" Or Text2.Text = "" Then
MsgBox ("���������?����?���")
Exit Sub
End If
Data27.Database.Execute "delete * from pldb where �ϵ����='" & Text2.Text & "'"
Data27.Database.Execute "INSERT INTO pldb SELECT * FROM pldd WHERE �ϵ����='" & Text2.Text & "'"
Formd11111.DataCombo1(0).Text = Text5.Text
Formd11111.DataCombo1(1).Text = DataCombo9(2).Text
Formd11111.DataCombo1(2).Text = Text2.Text
End Sub

Private Sub Command9_Click()
Data22.Database.Execute "delete * from pfda"

Adodc24.RecordSource = "select * from pfd where  ���='" & Text1.Text & "'"
Adodc24.Refresh
If Not Adodc24.Recordset.EOF Then
Adodc24.Recordset.MoveFirst

For i = 0 To 6
pfsz(i) = Adodc24.Recordset.Fields(i)
Next

mb = 0
For i = 7 To 56
If Adodc24.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next

ProgressBar1.Visible = True
For i = 7 To mb + 7
If Adodc24.Recordset.Fields(i) <> "" Then
sz(0) = Mid(Adodc24.Recordset.Fields(i), 1, InStr(Adodc24.Recordset.Fields(i), "(") - 1)
sz(1) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "(") + 1, InStr(Adodc24.Recordset.Fields(i), ")") - InStr(Adodc24.Recordset.Fields(i), "(") - 1)
sz(2) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), ")") + 1, InStr(Adodc24.Recordset.Fields(i), "-") - InStr(Adodc24.Recordset.Fields(i), ")") - 1)
sz(3) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "-") + 1, InStr(Adodc24.Recordset.Fields(i), "\") - InStr(Adodc24.Recordset.Fields(i), "-") - 1)
sz(4) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "\") + 1, InStr(Adodc24.Recordset.Fields(i), "#") - InStr(Adodc24.Recordset.Fields(i), "\") - 1)
sz(5) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "#") + 1, InStr(Adodc24.Recordset.Fields(i), "^") - InStr(Adodc24.Recordset.Fields(i), "#") - 1)
sz(6) = Mid(Adodc24.Recordset.Fields(i), InStr(Adodc24.Recordset.Fields(i), "^") + 1)
L = i - 6
Data22.Database.Execute "insert into pfda(�ӹ���λ,Ʒ��,ɫ��,��ɫ,�䷽���,������,�䷽����,��������,ԡ��,Ⱦ������,Ⱦ��������,��λ,�䷽,����,�����) VALUES('" & pfsz(0) & "','" & pfsz(1) & "','" & pfsz(2) & "','" & pfsz(3) & "','" & pfsz(4) & "','" & pfsz(5) & "',CDATE('" & pfsz(6) & "'),'" & sz(0) & "','" & sz(1) & "','" & sz(2) & "',  " & _
                                                                        "'" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & L & "')"
ProgressBar1.value = 100 / mb * L
End If
Next
ProgressBar1.Visible = False
End If

Data7.RecordSource = "SELECT * FROM pfda WHERE �䷽���='" & Text1.Text & "'"
Data7.Refresh
Data12.RecordSource = "SELECT �������� FROM pfda where �䷽���='" & Text1.Text & "'GROUP BY ��������"
Data12.Refresh
End Sub

Private Sub DataCombo1_Change()
Text4 = DataCombo1
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Text4 = DataCombo1
End Sub

Private Sub DataCombo9_Change(Index As Integer)
Select Case Index
       Case 3
If InStr(DataCombo9(3), "��") > 0 Or InStr(DataCombo9(3), "ϴ") > 0 Then
Text11 = "0.33"
Else
Text11 = ""
End If
End Select
End Sub

Private Sub DBCombo1_Change()
On Error Resume Next
If DBCombo1.Text = "ȫ��" Then
Data7.RecordSource = "SELECT * FROM pfda WHERE �䷽���='" & Text1.Text & "' ORDER BY VAL(��������),�����"
Data7.Refresh
Data12.RecordSource = "SELECT �������� FROM pfda where �䷽���='" & Text1.Text & "'GROUP BY ��������"
Data12.Refresh
Else
Data7.RecordSource = "SELECT * FROM pfda WHERE �䷽���='" & Text1.Text & "' AND ��������='" & DBCombo1.Text & "' ORDER BY �����"
Data7.Refresh
Data12.RecordSource = "SELECT �������� FROM pfda where �䷽���='" & Text1.Text & "'GROUP BY ��������"
Data12.Refresh
End If
End Sub

Private Sub DBCombo1_Click(Area As Integer)
On Error Resume Next
Data7.RecordSource = "SELECT * FROM pfda WHERE �䷽���='" & Text1.Text & "' AND ��������='" & DBCombo1.Text & "' ORDER BY VAL(��������),�����"
Data7.Refresh
End Sub

Private Sub dataCombo20_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode

End Sub



Private Sub dataCombo9_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Form_Load()
On Error Resume Next

TZBL = "1"
DataCombo1.Text = ""
DataCombo7.Text = ""
DBCombo1.Text = "ȫ��"
Text1.Text = ""
Text3.Text = ""
Text2.Text = ""
Text8.Text = ""
Text9.Text = ""
Text7.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text17.Text = ""
Combo1.Text = ""
Combo1111.Text = ""
scbl = "���"
Combo2.Text = "����"
plshsx = 1
Text15.Text = ""
Text16.Text = ""
cdbhf = cdbh
For i = 0 To 5
DataCombo9(i).Text = ""
Next
Option1(0).value = True
Data20.DatabaseName = App.Path & "\AccessBase\DB.mdb"
plshsx = 1              '''''''''''''''''''�������ˢ��

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc21.RecordSource = "SELECT * FROM kpD WHERE ����='" & Text5.Text & "'"
Adodc21.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select DISTINCT ģ���� from CGGYMB ORDER by ģ����"
Adodc3.Refresh


Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc17.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc32.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM DBPLDBH where ����='" & yhdm & "'"
Adodc1.Refresh

Text2.Text = yhdm + "1" ''''''''''''OK
If Adodc1.Recordset.EOF Then
Text2.Text = yhdm + "1" ''''''''''''OK
Else
L = Val(Adodc1.Recordset.Fields(0))
Text2.Text = yhdm + Trim(L + 1) '''''''''''''OK
End If

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select ��� from khZL  group by ���"
Adodc2.Refresh



Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc4.RecordSource = "select ���չ���,��� from gx group by ���չ���,��� ORDER BY ���"
'Adodc4.Refresh

Adodc20.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc20.RecordSource = "select distinct ��̨��� from ct order by ��̨���"
Adodc20.Refresh


Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Data7.DatabaseName = App.Path & "\AccessBase\DB.mdb"

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "SELECT ���������� FROM dbGR GROUP BY ����������"
Adodc11.Refresh

Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Data11.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data12.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data17.DatabaseName = App.Path & "\AccessBase\DB.mdb"

Data13.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh

Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc19.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Data22.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data23.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Adodc24.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc25.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc25.RecordSource = "SELECT ����Ҫ�� from GYYQ group by ����Ҫ�� "
Adodc25.Refresh
Adodc29.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc30.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Data27.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data28.DatabaseName = App.Path & "\AccessBase\DB.mdb"

VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(1) = 0
VSFlexGrid2.ColWidth(2) = 0
VSFlexGrid2.ColWidth(3) = 0
VSFlexGrid2.ColWidth(4) = 0
VSFlexGrid2.ColWidth(5) = 1000
VSFlexGrid2.ColWidth(6) = 1500
VSFlexGrid2.ColWidth(7) = 1500
VSFlexGrid2.ColWidth(8) = 2000
VSFlexGrid2.ColWidth(9) = 1000
VSFlexGrid2.ColWidth(10) = 1500
VSFlexGrid2.ColWidth(11) = 1000
VSFlexGrid2.ColWidth(12) = 1200

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 400
VSFlexGrid1.ColWidth(7) = 6000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0
Adodc31.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc31.RecordSource = "select ���� from yhb  where �û� ='" & yhm & "'"
Adodc31.Refresh
DataCombo20.Text = Adodc31.Recordset.Fields(0)
DataCombo9(0).Text = Date
Text5.TabIndex = 0
szh = "����"

If Len(yhdm) <> 1 Then
MsgBox ("����˻������ʽ����������")
Command1.Enabled = False
Command4.Enabled = False
Command14.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command2.Enabled = False
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

Private Sub Label11_Click()
Text12 = ""
Text12.SetFocus
End Sub


Private Sub Label13_dblClick()
On Error Resume Next
If MsgBox("ȷ������" + DBCombo1.Text + "ԡ����", vbYesNo) = vbNo Then Exit Sub
If DBCombo1.Text = "ȫ��" Then
Data13.Database.Execute "UPDATE pldd SET ˮ��=val('" & Text8.Text & "'),ԡ��=format(val('" & Text8.Text & "')/val('" & DataCombo9(2) & "'),'#0.000') where �ϵ����='" & Text2.Text & "' and val('" & DataCombo9(2) & "')<>0"
Else
Data13.Database.Execute "UPDATE pldd SET ˮ��=val('" & Text8.Text & "'),ԡ��=format(val('" & Text8.Text & "')/val('" & DataCombo9(2) & "'),'#0.000') WHERE ��������='" & DBCombo1.Text & "' AND �ϵ����='" & Text2.Text & "' and val('" & DataCombo9(2) & "')<>0"
End If
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 400
VSFlexGrid1.ColWidth(7) = 2000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0
End Sub

Private Sub Label14_Click()
On Error Resume Next
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 400
VSFlexGrid1.ColWidth(7) = 6000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0
VSFlexGrid1.ColFormat(11) = "#0.####"
End Sub




Private Sub Label15_Click()
FormS499.Show
End Sub

Private Sub Label16_Click()
Formc140.Text1 = Text5
Formc140.Show
End Sub

Private Sub Label17_Click()
Formr334.DataCombo1 = Text2
Formr334.Show
End Sub

Private Sub Label19_Click()
On Error Resume Next
If MsgBox("ȷ������" + DBCombo1.Text + "ԡ����", vbYesNo) = vbNo Then Exit Sub
If DBCombo1.Text = "ȫ��" Then
Data13.Database.Execute "UPDATE pldd SET ԡ��='" & Text16 & "' WHERe �ϵ����='" & Text2.Text & "'"
Else
Data13.Database.Execute "UPDATE pldd SET ԡ��='" & Text16 & "' WHERE ��������='" & DBCombo1.Text & "' AND �ϵ����='" & Text2.Text & "'"
End If
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 400
VSFlexGrid1.ColWidth(7) = 6000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0
End Sub

Private Sub Label2_Click()
Text2.Enabled = True
End Sub

Private Sub Label2_DblClick()
Text2.Enabled = False
End Sub

Private Sub Label20_Click(Index As Integer)
Select Case Index
       Case 1
       Text1.Enabled = True
End Select
End Sub

Private Sub Label20_DblClick(Index As Integer)
Select Case Index
       Case 1
       Text1.Enabled = False
End Select
End Sub

Private Sub Label21_Click()
FormS499.Show
End Sub

Private Sub Label23_Click()
pfjh = 2
FormJ8.Check2(4).value = 1
FormJ8.Show
End Sub

Private Sub Label3_Click()
'FormA115.Text11 = Text5
'FormA115.Show
End Sub

Private Sub Label4_dblClick()
'FormS4.Show
End Sub

Private Sub Label49_Click()
Formd221.Show
Formd221.Text1.Text = Text6.Text
End Sub

Private Sub Label5_Click()
On Error Resume Next
Data11.RecordSource = "SELECT sum(val(�䷽)) FROM pldd where �ϵ����='" & Text2.Text & "' and instr(Ⱦ������,'Ⱦ��')>0"
Data11.Refresh
pfyl = 0
pfyl = Val(Data11.Recordset.Fields(0))
'pfyljt = Text4''''''�ѳ�̨����Ⱦɫ����ѡ���text3
Formd334.Text1 = Text5
Formd334.Show
End Sub

Private Sub Label55_Click()
'Formd44.DataCombo2.Text = DataCombo9(1).Text
Formd44.Text2.Text = Text6.Text
Formd44.Combo1.Text = ""
Formd44.Combo2.Text = "���"
Formd44.Text5.Text = Text5.Text
Formd44.Text6.Text = DataCombo9(2).Text
Formd44.Show
End Sub


Private Sub Label7_dblClick()
DBCombo1.Text = ""
DBCombo1.Text = Label7.Caption
End Sub

Private Sub Label8_Click()
FormD335.Text1 = Text5.Text
FormD335.Show
End Sub

Private Sub Label9_dblClick()
Formd336.Text1 = Text2
Formd336.Text2 = Text5
Formd336.Show
End Sub

Private Sub Text13_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select ��� from khZL where ���� like '%'+'" & Text13 & "' +'%' group by ���"
Adodc2.Refresh
End Sub


Private Sub Text15_Change()
Adodc20.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc20.RecordSource = "select distinct ��̨��� from ct where ��̨���  like '%'+'" & Text15 & "'+'%' order by ��̨���"
Adodc20.Refresh
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
If plshsx = 300 Then

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "pldrzdhs"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

plshsx = 1
Else
plshsx = plshsx + 1
End If
End Sub


Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
End With

S1 = VSFlexGrid1.TextMatrix(r, 3)   '''�ϵ����
S2 = VSFlexGrid1.TextMatrix(r, 4)   '''��������
s3 = VSFlexGrid1.TextMatrix(r, 5)   ''ԡ��
s4 = VSFlexGrid1.TextMatrix(r, 6)   ''Ⱦ������
s5 = VSFlexGrid1.TextMatrix(r, 7)   '''Ⱦ��������
s6 = VSFlexGrid1.TextMatrix(r, 8)   '''�䷽��λ
s7 = VSFlexGrid1.TextMatrix(r, 9)   '''�䷽
s8 = VSFlexGrid1.TextMatrix(r, 10)  '''У��ֵ
s9 = Val(VSFlexGrid1.TextMatrix(r, 16)) + 1  ''�����


    If Button = 2 Then
    If MsgBox("ȷ���������е���Ϣ��" + s5, vbYesNo) = vbNo Then  '''PopupMenu mnu_manager  '�����ڴ��������õ�һ�������˵�����
    Exit Sub
    Else
    Data7.Database.Execute "insert into pldd(�ϵ����,��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,�����,����,����) values('" & S1 & "','" & S2 & "','" & s3 & "','" & s4 & "','" & s5 & "','" & s6 & "','" & s7 & "','1','" & s9 & "','','')"
    End If
    Call Label14_Click
    End If
End Sub
Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Data13.Recordset.EOF Then Exit Sub
Data13.Recordset.MoveFirst
rs = VSFlexGrid1.Row
rc = VSFlexGrid1.col
Data13.Recordset.Move rs - 1
If rc = 1 Then
Data13.Recordset.Delete
Data13.Refresh
Else
DBCombo1 = Data13.Recordset.Fields(3)
End If
Call Label14_Click
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
rs = VSFlexGrid2.Row
If Data7.Recordset.EOF Then
Exit Sub
End If

Data7.Recordset.MoveFirst
Data7.Recordset.Move rs - 1

Data13.Recordset.AddNew
Data13.Recordset.Fields(0) = Text5.Text
Data13.Recordset.Fields(1) = DataCombo9(2).Text
Data13.Recordset.Fields(2) = Text2.Text
Data13.Recordset.Fields(3) = Data7.Recordset.Fields(4)
Data13.Recordset.Fields(4) = Data7.Recordset.Fields(5)
Data13.Recordset.Fields(5) = Data7.Recordset.Fields(6)
Data13.Recordset.Fields(6) = Data7.Recordset.Fields(7)
Data13.Recordset.Fields(7) = Data7.Recordset.Fields(8)
Data13.Recordset.Fields(8) = Data7.Recordset.Fields(9)
Data13.Recordset.Fields(9) = Data7.Recordset.Fields(10)
Data13.Recordset.Fields(15) = Data7.Recordset.Fields(16)
Data13.Recordset.Fields(25) = Data7.Recordset.Fields(19)
Data13.Recordset.Update
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh
End Sub
Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text10_Change()
If Text10.Text = "" Then
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT Ⱦ������ FROM RHZH WHERE Ⱦ������='" & Text10.Text & "'"
Adodc8.Refresh
Else
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT Ⱦ������ FROM RHZH WHERE ���� LIKE '%'+'" & Text10.Text & "'+'%'"
Adodc8.Refresh
End If
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text1_Change()
On Error Resume Next
Data7.RecordSource = "SELECT * FROM pfda WHERE �䷽���='" & Text1.Text & "'"
Data7.Refresh
Data12.RecordSource = "SELECT �������� FROM pfda where �䷽���='" & Text1.Text & "'GROUP BY ��������"
Data12.Refresh
End Sub

Private Sub Text2_Change()
On Error Resume Next
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "' ORDER BY ��������,�����"
Data13.Refresh

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 400
VSFlexGrid1.ColWidth(7) = 6000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0
End Sub



Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode

End Sub

Private Sub Text5_Change()
On Error Resume Next

If InStr(Text5, "j") > 0 Or InStr(Text5, "J") > 0 Then
Text5.Text = Mid(Text5, 1, Len(Text5) - 1)
End If

 If Len(Text5.Text) < 4 Then Exit Sub
          
               '�����������
            Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
            Adodc18.RecordSource = "select max(����) as zl from kpd where ����='" & Text5.Text & "'"
            Adodc18.Refresh
                 If Adodc18.Recordset.EOF Then
                   ' MsgBox ("�ƻ������»��ʧ�󣡣�")
                    Exit Sub
                 End If
             a = Adodc18.Recordset.Fields("zl")    '������������Ƹ�����A
             Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

             Adodc18.RecordSource = "select SUM(isnull(�������,0)) as zl1,SUM(isnull(���ƥ��,0)) as zl2 from v_kpd_ok where ����='" & Text5.Text & "'"   '''ͳ������
             Adodc18.Refresh
                If Adodc18.Recordset.EOF Then
                   '  MsgBox ("�ƻ������»��ʧ�󣡣�")
                     Exit Sub
                End If
            c1 = Adodc18.Recordset.Fields("zl1")    '�����������Ƹ�����C
            c2 = Adodc18.Recordset.Fields("zl2")    '����ƥ�����Ƹ�����C
            Adodc19.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

            Adodc19.RecordSource = "select * from kpd where ����='" & Text5.Text & "' And ���� = cast('" & a & "' as real)"
            Adodc19.Refresh
                  If Adodc19.Recordset.EOF Then
                    '  MsgBox ("�»��ʧ�󣡣�")
                      Exit Sub
                  End If
            
            d = Adodc19.Recordset.Fields(52)
            e = Adodc19.Recordset.Fields(8)
            DH = Adodc19.Recordset.Fields(1)
'            Text3.Text = dh
            DataCombo9(0).Text = Date
            DataCombo9(1).Text = Adodc19.Recordset.Fields(0)
            DataCombo9(2).Text = Format(c1, "#0.0")
            DataCombo9(3).Text = e
            DataCombo9(5).Text = Adodc19.Recordset.Fields(3)
            DataCombo9(6).Text = d
            Text6.Text = d
            Text7.Text = c2
            Text4 = Adodc19.Recordset.Fields(14)
            Text11 = Adodc19.Recordset.Fields(30)
            Adodc18.RecordSource = "select �ܱ�ע from sczy_z  where ���� in(select distinct ���� from kpd where ����='" & Text5.Text & "' and len(isnull(����,0))>0)"
            Adodc18.Refresh
            If Not Adodc18.Recordset.EOF Then
            Text3 = Adodc18.Recordset.Fields(0)
            Else
            Text3 = ""
            End If
           ' Text4.Text = Adodc19.Recordset.Fields(14)
            
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select * from pld where ����='" & Text5.Text & "'"
Adodc5.Refresh

If Not Adodc5.Recordset.EOF Then
            
            DataCombo9(2).Text = Format(Adodc5.Recordset.Fields(5), "#0.0")
            Text4.Text = Adodc5.Recordset.Fields(7)
End If
End Sub


Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub vSFlexGrid1_DbClick()
On Error Resume Next
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
    If c > 4 Then
    If c = 19 Then
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
    Else
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
    End If
    End If
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
    Call MSF
End If
End Sub
Private Sub MSF()
On Error Resume Next
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
    If c = 7 Or c = 8 Or c = 9 Or c = 10 Or c = 11 Or c = 12 Or c = 16 Or c = 19 Then
    
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111.Text = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
   End If
End With
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
If c = 10 Then
If Val(Combo1111.Text) > 1 Then
If MsgBox("У��ֵ��Լ������1�����ϣ���ȷ���Ƿ������", vbYesNo) = vbNo Then Exit Sub
End If
End If



Data13.Recordset.MoveFirst
Data13.Recordset.Move r - 1
Data13.Recordset.Edit

If c = 7 Then
Adodc14.RecordSource = "select distinct Ⱦ������,Ⱦ�������� from rhzh where ����='" & Combo1111 & "'"
Adodc14.Refresh
If Not Adodc14.Recordset.EOF Then
Combo1111 = Adodc14.Recordset.Fields(0)
Data13.Recordset.Fields(c - 2) = Adodc14.Recordset.Fields(1)
VSFlexGrid1.TextMatrix(r, c - 1) = Adodc14.Recordset.Fields(1)
End If
End If

Data13.Recordset.Fields(c - 1) = Combo1111.Text
Data13.Recordset.Update
VSFlexGrid1.Text = Combo1111.Text
Combo1111.Visible = False
VSFlexGrid1.SetFocus
End If

If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode

End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Timer2_Timer()
If JDBAR = 10 Then
      Call pldd(Data22, Data23, Data20, Text2.Text, Text3.Text, Text11, Adodc16, Combo1)
      ProgressBar1.Visible = False
      Timer2.Enabled = False
      JDBAR = 0
Exit Sub
End If
ProgressBar1.value = JDBAR * 10
JDBAR = JDBAR + 1

End Sub

Private Sub Lwx()
On Error Resume Next
Data13.RecordSource = "SELECT * FROM pldd where �ϵ����='" & Text2.Text & "'  ORDER BY ��������,�����"
Data13.Refresh
End Sub

Private Sub scpfd(bh As String)
On Error Resume Next
ysl = 0
Data22.RecordSource = "select distinct ��������,ˮ�� from pldd where �ϵ����='" & bh & "' order by ��������"
Data22.Refresh
If Data22.Recordset.EOF Then Exit Sub
Data22.Recordset.MoveFirst
i = 12

Do While Not Data22.Recordset.EOF
ysl = ysl + Val(Data22.Recordset.Fields(1))
Data23.RecordSource = "select ��������,ԡ��,Ⱦ������,Ⱦ��������,�䷽��λ,�䷽,У��ֵ,��������,���ϵ�λ,���� from pldd where �ϵ����='" & bh & "' and ��������='" & Data22.Recordset.Fields(0) & "' order by �����"
Data23.Refresh
If Not Data23.Recordset.EOF Then
Data23.Recordset.MoveFirst
Do While Not Data23.Recordset.EOF
If IsNull(Data23.Recordset.Fields(9)) Then
L = ""
Else
L = Trim(Data23.Recordset.Fields(9))
End If
sz(i) = Data23.Recordset.Fields(0) + "(" + Data23.Recordset.Fields(1) + ")" + Data23.Recordset.Fields(2) + "-" + Data23.Recordset.Fields(3) + "\" + Data23.Recordset.Fields(4) + "#" + Trim(Data23.Recordset.Fields(5)) + "^" + Data23.Recordset.Fields(6) + "[" + Trim(Data23.Recordset.Fields(7)) + "]" + Data23.Recordset.Fields(8) + "{" + L
i = i + 1
Data23.Recordset.MoveNext
Loop
End If

Data22.Recordset.MoveNext
Loop

If i < 62 Then
For L = i To 61
sz(L) = ""
Next
End If


Data22.RecordSource = "select distinct ���,����,ѹ��,�������,�䷽��,����,���ϴ�ӡԱ,Ⱦ��������,��������,������Ϣ,�ϵ���� from pldd where �ϵ����='" & bh & "'"
Data22.Refresh
If Data22.Recordset.EOF Then Exit Sub
Data22.Recordset.MoveFirst
i = 0
For i = 0 To 10
sz(i) = Data22.Recordset.Fields(i)
Next
sz(11) = "δ"

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
g_Cmd.CommandText = "dbpld('" & sz(0) & "','" & sz(1) & "','" & sz(2) & "','" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & sz(10) & "','" & sz(11) & "','" & sz(12) & "','" & sz(13) & "','" & sz(14) & "','" & sz(15) & "','" & sz(16) & "','" & sz(17) & "','" & sz(18) & "','" & sz(19) & "','" & sz(20) & "','" & sz(21) & "','" & sz(22) & "','" & sz(23) & "','" & sz(24) & "','" & sz(25) & "','" & sz(26) & "','" & sz(27) & "','" & sz(28) & "','" & sz(29) & "','" & sz(30) & "','" & sz(31) & "','" & sz(32) & "','" & sz(33) & "','" & sz(34) & "','" & sz(35) & "','" & sz(36) & "','" & sz(37) & "','" & sz(38) & "','" & sz(39) & "','" & sz(40) & "','" & sz(41) & "','" & sz(42) & "','" & sz(43) & "','" & sz(44) & "','" & sz(45) & "','" & sz(46) & "','" & sz(47) & "','" & sz(48) & "','" & sz(49) & "','" & sz(50) & "','" & sz(51) & "','" & sz(52) & "','" & sz(53) & "','" & sz(54) & "','" & sz(55) & "','" & sz(56) & "','" & sz(57) & "'," & _
                           "'" & sz(58) & "','" & sz(59) & "','" & sz(60) & "','" & sz(61) & "')"
g_Cmd.Execute           ' ִ�д洢����
g_Cmd.Cancel

yqz = Val(Text11)
Text7 = Val(Text7)
Text16 = Val(Text16)
sql1 = "update pld set ˮ��='" & ysl & "',��ֵ='" & yqz & "',ƥ��='" & Text7 & "',ԡ��='" & Text16 & "',����='" & Text17 & "' WHERE ���='" & bh & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

End Sub

Private Sub clxt(bh As String)
On Error GoTo errorhandler ' ��ʼ������
sql1 = "delete  from pldr WHERE �ϵ����='" & bh & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

Adodc15.RecordSource = "select * from pld where  ���='" & bh & "'"
Adodc15.Refresh
If Not Adodc15.Recordset.EOF Then
Adodc15.Recordset.MoveFirst

For i = 0 To 10
ZS(i) = Adodc15.Recordset.Fields(i)
Next

mb = 0
For i = 12 To 61
If Adodc15.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next

For i = 12 To mb + 12
If Adodc15.Recordset.Fields(i) <> "" Then
sz(0) = Mid(Adodc15.Recordset.Fields(i), 1, InStr(Adodc15.Recordset.Fields(i), "(") - 1)
sz(1) = Mid(Adodc15.Recordset.Fields(i), InStr(Adodc15.Recordset.Fields(i), "(") + 1, InStr(Adodc15.Recordset.Fields(i), ")") - InStr(Adodc15.Recordset.Fields(i), "(") - 1)
sz(2) = Mid(Adodc15.Recordset.Fields(i), InStr(Adodc15.Recordset.Fields(i), ")") + 1, InStr(Adodc15.Recordset.Fields(i), "-") - InStr(Adodc15.Recordset.Fields(i), ")") - 1)
sz(3) = Mid(Adodc15.Recordset.Fields(i), InStr(Adodc15.Recordset.Fields(i), "-") + 1, InStr(Adodc15.Recordset.Fields(i), "\") - InStr(Adodc15.Recordset.Fields(i), "-") - 1)
sz(4) = Mid(Adodc15.Recordset.Fields(i), InStr(Adodc15.Recordset.Fields(i), "\") + 1, InStr(Adodc15.Recordset.Fields(i), "#") - InStr(Adodc15.Recordset.Fields(i), "\") - 1)
sz(5) = Mid(Adodc15.Recordset.Fields(i), InStr(Adodc15.Recordset.Fields(i), "#") + 1, InStr(Adodc15.Recordset.Fields(i), "^") - InStr(Adodc15.Recordset.Fields(i), "#") - 1)
sz(6) = Mid(Adodc15.Recordset.Fields(i), InStr(Adodc15.Recordset.Fields(i), "^") + 1, InStr(Adodc15.Recordset.Fields(i), "[") - InStr(Adodc15.Recordset.Fields(i), "^") - 1)
sz(7) = Mid(Adodc15.Recordset.Fields(i), InStr(Adodc15.Recordset.Fields(i), "[") + 1, InStr(Adodc15.Recordset.Fields(i), "]") - InStr(Adodc15.Recordset.Fields(i), "[") - 1)
sz(8) = Mid(Adodc15.Recordset.Fields(i), InStr(Adodc15.Recordset.Fields(i), "]") + 1, InStr(Adodc15.Recordset.Fields(i), "{") - InStr(Adodc15.Recordset.Fields(i), "]") - 1)
sz(9) = Mid(Adodc15.Recordset.Fields(i), InStr(Adodc15.Recordset.Fields(i), "{") + 1)

L = i - 11

If Trim(sz(8)) = "g" Then
sz(8) = "kg"
sz(7) = Format(Val(sz(7)) / 1000, "#0.00000")
End If

sql1 = "insert into pldr(����,����,������Ϣ,�ϵ����,��������,��������,ԡ��,Ⱦ������,Ⱦ��������,���ϵ�λ,��������,ʵ�ʳ���,�����,��̨) VALUES('" & ZS(1) & "','" & ZS(5) & "','" & ZS(9) & "','" & ZS(10) & "','" & ZS(8) & "','" & sz(0) & "','" & sz(1) & "','" & sz(2) & "',  " & _
                                                                        "'" & sz(3) & "','" & sz(8) & "','" & sz(7) & "','','" & L & "','" & ZS(7) & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
End If
Adodc30.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    Adodc30.RecordSource = "SELECT ���׹��� from BGXX WHERE ���ϱ��='" & Text2.Text & "'"
    Adodc30.Refresh
    
    If Not Adodc30.Recordset.EOF Then
        Adodc30.Recordset.MoveFirst
        Dim updateValue As String
        Dim isFirst As Boolean
        isFirst = True
        
        ' �������м�¼�������Ǻϲ���һ���ַ���
        Do Until Adodc30.Recordset.EOF
            If isFirst Then
                updateValue = Adodc30.Recordset.Fields(0).value
                isFirst = False
            Else
                updateValue = updateValue & "/" & Adodc30.Recordset.Fields(0).value
            End If
            Adodc30.Recordset.MoveNext
        Loop
        ' ����������
        Dim strSQL As String
        strSQL = "UPDATE pld SET ���׹��� = '" & updateValue & "' WHERE ���='" & Text2.Text & "'"
        
        ' ִ�и��²���
        conn.Execute strSQL
    End If
   Exit Sub

errorhandler:
    ' �������������ʾ������Ϣ�����ر����ݿ�����
    MsgBox "��������" & Err.Description & "������ţ�" & Err.Number & "��", vbCritical, "����pldr����"
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub

