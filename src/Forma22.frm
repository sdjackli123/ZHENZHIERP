VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma22 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�ƻ�����"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   Icon            =   "Forma22.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ɾ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�޸�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   4800
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   330
      Left            =   3360
      Top             =   9840
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
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ƻ���ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   4800
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   3960
      Top             =   9960
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
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���մ�ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   9720
      Style           =   1  'Simple Combo
      TabIndex        =   75
      Text            =   "Combo1111"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma22.frx":440A
      Height          =   3255
      Left            =   360
      TabIndex        =   6
      Top             =   6240
      Width           =   18255
      _cx             =   32200
      _cy             =   5741
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
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�µ��ݺ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   1095
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   91
      Text            =   "Forma22.frx":441F
      Top             =   3360
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF80&
      Caption         =   "ί��"
      Height          =   375
      Left            =   2160
      TabIndex        =   90
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF80&
      Caption         =   "����"
      Height          =   375
      Left            =   360
      TabIndex        =   89
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   5400
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   4800
      Top             =   9960
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
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���Ų�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ɾ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����׼��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   73
      Text            =   "Text6"
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   72
      Text            =   "Text3"
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "֯�Ÿ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   7200
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma22.frx":4425
      Height          =   360
      Index           =   0
      Left            =   5520
      TabIndex        =   3
      Top             =   6720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5040
      Top             =   10080
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
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
      Left            =   5520
      Top             =   9960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
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
      Left            =   5400
      Top             =   9960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
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
      Left            =   6000
      Top             =   9840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
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
      Left            =   6000
      Top             =   9840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
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
      Left            =   5640
      Top             =   10080
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
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
      Left            =   5520
      Top             =   10080
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
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
      Left            =   6120
      Top             =   9960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
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
      Left            =   6120
      Top             =   9960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
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
      Left            =   5280
      Top             =   9840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   1
      Left            =   2160
      TabIndex        =   7
      Top             =   6720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   2
      Left            =   9000
      TabIndex        =   8
      Top             =   6720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   3
      Left            =   1080
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma22.frx":443A
      Height          =   360
      Index           =   4
      Left            =   2160
      TabIndex        =   10
      Top             =   7200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "pm"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   5
      Left            =   2160
      TabIndex        =   11
      Top             =   7800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   6
      Left            =   2160
      TabIndex        =   12
      Top             =   8400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   7
      Left            =   1080
      TabIndex        =   13
      Top             =   2880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   8
      Left            =   15240
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   9
      Left            =   5280
      TabIndex        =   15
      Top             =   9000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma22.frx":444F
      Height          =   330
      Index           =   10
      Left            =   1080
      TabIndex        =   16
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "��������"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   11
      Left            =   7800
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   12
      Left            =   7800
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   13
      Left            =   6000
      TabIndex        =   19
      Top             =   7560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   14
      Left            =   5280
      TabIndex        =   20
      Top             =   7920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   15
      Left            =   11760
      TabIndex        =   21
      Top             =   9000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   16
      Left            =   8400
      TabIndex        =   22
      Top             =   7800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   17
      Left            =   8400
      TabIndex        =   23
      Top             =   8400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   19
      Left            =   12360
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma22.frx":4464
      Height          =   330
      Index           =   20
      Left            =   1080
      TabIndex        =   26
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "��̨"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   21
      Left            =   12360
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   22
      Left            =   7800
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   23
      Left            =   4320
      TabIndex        =   29
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   24
      Left            =   12120
      TabIndex        =   30
      Top             =   7200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   25
      Left            =   12120
      TabIndex        =   31
      Top             =   7800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   26
      Left            =   12120
      TabIndex        =   32
      Top             =   8400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   27
      Left            =   12120
      TabIndex        =   33
      Top             =   9000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   28
      Left            =   2160
      TabIndex        =   34
      Top             =   9000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5280
      TabIndex        =   35
      Top             =   7200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   330825729
      CurrentDate     =   39961
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8400
      TabIndex        =   36
      Top             =   7200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   330825729
      CurrentDate     =   39961
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma22.frx":4479
      Height          =   360
      Index           =   29
      Left            =   8400
      TabIndex        =   37
      Top             =   9000
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "xm"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma22.frx":448E
      Height          =   330
      Index           =   30
      Left            =   3120
      TabIndex        =   67
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "���ձ��"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   31
      Left            =   12240
      TabIndex        =   68
      Top             =   6720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   32
      Left            =   7800
      TabIndex        =   84
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma22.frx":44A3
      Height          =   360
      Index           =   18
      Left            =   4320
      TabIndex        =   24
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   38
      Left            =   4200
      TabIndex        =   86
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   4320
      TabIndex        =   94
      Top             =   2400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   423034881
      CurrentDate     =   39961
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   4320
      TabIndex        =   96
      Top             =   2880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   423034881
      CurrentDate     =   39961
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��̨"
      Height          =   375
      Index           =   22
      Left            =   3600
      TabIndex        =   99
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   21
      Left            =   3600
      TabIndex        =   97
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   20
      Left            =   3600
      TabIndex        =   93
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ע"
      Height          =   1095
      Index           =   4
      Left            =   360
      TabIndex        =   92
      Top             =   3360
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4575
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   8895
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   23
      Left            =   3480
      TabIndex        =   87
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      Height          =   375
      Left            =   3600
      TabIndex        =   85
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�ʼ�"
      Height          =   375
      Index           =   9
      Left            =   7080
      TabIndex        =   83
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ҫ���Ƶ���֯��"
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   74
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ϻ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   12840
      TabIndex        =   71
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      Height          =   375
      Index           =   7
      Left            =   2400
      TabIndex        =   70
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ɴ"
      Height          =   375
      Index           =   0
      Left            =   11280
      TabIndex        =   69
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      Height          =   375
      Left            =   1200
      TabIndex        =   66
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   6
      Left            =   1080
      TabIndex        =   65
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      Height          =   375
      Index           =   4
      Left            =   7920
      TabIndex        =   64
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "֯��"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   63
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ƿ֯"
      Height          =   375
      Index           =   2
      Left            =   11520
      TabIndex        =   62
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�ͻ�"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   61
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�ۼ�"
      Height          =   375
      Index           =   6
      Left            =   11520
      TabIndex        =   60
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   8
      Left            =   11520
      TabIndex        =   59
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�ò�"
      Height          =   375
      Index           =   5
      Left            =   11520
      TabIndex        =   58
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   57
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ʒ��"
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   56
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ת��"
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   55
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   54
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   53
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�ƻ�"
      Height          =   375
      Index           =   10
      Left            =   360
      TabIndex        =   52
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�߳�"
      Height          =   375
      Index           =   11
      Left            =   14520
      TabIndex        =   51
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      Height          =   375
      Index           =   5
      Left            =   4560
      TabIndex        =   50
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   12
      Left            =   360
      TabIndex        =   49
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "֯��"
      Height          =   375
      Index           =   13
      Left            =   7080
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�۷�"
      Height          =   375
      Index           =   14
      Left            =   7080
      TabIndex        =   47
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ë��"
      Height          =   375
      Index           =   7
      Left            =   5280
      TabIndex        =   46
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      Height          =   375
      Index           =   15
      Left            =   4560
      TabIndex        =   45
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   16
      Left            =   11040
      TabIndex        =   44
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ע"
      Height          =   375
      Index           =   17
      Left            =   7680
      TabIndex        =   43
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ҫ"
      Height          =   375
      Index           =   8
      Left            =   7680
      TabIndex        =   42
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   18
      Left            =   4560
      TabIndex        =   41
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   19
      Left            =   7680
      TabIndex        =   40
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   9
      Left            =   11640
      TabIndex        =   39
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ҵ��"
      Height          =   375
      Index           =   10
      Left            =   7680
      TabIndex        =   38
      Top             =   9000
      Width           =   735
   End
End
Attribute VB_Name = "Forma22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs, r, c  As Integer
Private strname As String
Dim Stm As New ADODB.Stream
Dim StrPicTemp As String
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private strn As String

Private Sub Command1_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If Adodc1.Recordset.Fields(6) = Null Or Val(Adodc1.Recordset.Fields(6)) <= 0 Then
MsgBox ("���ܸ��ƣ�û���Ų���������Ҫ���ƣ�")
Exit Sub
End If

If Adodc1.Recordset.Fields(5) <= Val(Adodc1.Recordset.Fields(6)) Then
MsgBox ("���ܸ��ƣ��Ų��Ѿ������ƻ�����")
Exit Sub
End If


If MsgBox("ȷ��Ҫ���Ƶ���֯��Ϊ" + Text6.Text + "��", vbYesNo) = vbNo Then Exit Sub
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select ֯�� from zbkpd where ֯��='" & Text6.Text & "'"
Adodc5.Refresh

If Adodc5.Recordset.EOF Then
sql1 = "insert into zbkpd(�ͻ�,����,���,֯��,Ʒ��,����,Ͳ��,����,�ƻ�,����,����,����,��ע,������,ë��,ҵ��,ɴ��,֯��,�۷�,���,����,���,����,ת��,���) select �ͻ�,����,���,'" & Text6.Text & "',Ʒ��,����,Ͳ��,����,�ƻ�,����,����,����,��ע,������,ë��,ҵ��,ɴ��,֯��,�۷�,0,0,���,����,ת��,���+1 from zbkpd where ֯��='" & DataCombo1(3).Text & "'"
sql2 = "insert into sxpb(֯��,ɴ֧,֯��,���,����) select '" & Text6.Text & "',ɴ֧,֯��,���,���� from sxpb where ֯��='" & DataCombo1(3).Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("���Ƴɹ���")
Else
MsgBox ("Ҫ���Ƶ���֯���Ѵ��ڣ�")
End If
End Sub


Private Sub Command10_Click()
jhbl = 1
Formj13.DataCombo6 = DataCombo1(3)
Formj13.Show
End Sub

Private Sub Command11_Click()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select MAX(right(����,2)) as h  from kpd_jtjh where ����=cast('" & DTPicker3.value & "' as datetime) AND left(����,1)='" & yhdm & "'"
Adodc2.Refresh
DataCombo1(10).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "01"
If Adodc2.Recordset.EOF Then
DataCombo1(10).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "01"
Else
uu = Val(Adodc2.Recordset.Fields(0)) + 1
Select Case Len(uu)
       Case "1"
DataCombo1(10).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "0" + Trim(uu)
       Case "2"
DataCombo1(10).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + Trim(uu)
End Select
End If

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select ��� from kpd_jtjh where ����='" & DataCombo1(10).Text & "' AND left(����,1)='" & yhdm & "' order by ��� desc"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
DataCombo1(23).Text = 1
Else
DataCombo1(23).Text = Adodc3.Recordset.Fields(0) + 1
End If


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM kpd_jtjh where  ����='" & DataCombo1(10).Text & "' order by ��� desc"
Adodc1.Refresh
DataCombo1(3) = ""
Text4 = ""
DataCombo1(7) = ""
DataCombo1(20) = ""
End Sub

Private Sub Command12_Click()
Adodc1.RecordSource = "SELECT * FROM kpd_jtjh where  ����='" & DataCombo1(10).Text & "' order by ��� desc"
Adodc1.Refresh

Adodc3.Refresh
If Adodc3.Recordset.EOF Then
DataCombo1(23).Text = 1
Else
DataCombo1(23).Text = Adodc3.Recordset.Fields(0) + 1
End If

Adodc3.RecordSource = "select ��� from kpd_jtjh where ����='" & DataCombo1(10).Text & "' AND left(����,1)='" & yhdm & "' order by ��� desc"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
DataCombo1(23).Text = 1
Else
DataCombo1(23).Text = Adodc3.Recordset.Fields(0) + 1
End If

DataCombo1(7) = ""
Text4 = ""
End Sub

Private Sub Command13_Click()
Call jhlcd(Adodc10, Adodc12, DataCombo1(10))
End Sub

Private Sub Command14_Click()
Call ddlcdjh(Adodc10, Adodc12, DataCombo1(10))
End Sub

Private Sub Command15_Click()
On Error Resume Next
If DataCombo1(3).Text = "" Then
MsgBox ("������֯��")
Exit Sub
End If

If Val(DataCombo1(7).Text) <= 0 Then
MsgBox ("�������Ų���")
Exit Sub
End If


If DataCombo1(20).Text = "" Then
MsgBox ("�����복��")
Exit Sub
End If

If DataCombo1(3).Text = "" Or DataCombo1(7).Text = "" Then
MsgBox ("�������Ų����ƻ���ת����Ϣ")
Exit Sub
End If

If Option1.value = True Then
lb = "����"
End If

If Option2.value = True Then
lb = "ί��"
End If

If Adodc1.Recordset.EOF Then Exit Sub

If MsgBox("ȷ���޸���" + Adodc1.Recordset.Fields(1), vbYesNo) = vbNo Then Exit Sub
sql1 = "update kpd_jtjh set ����='" & DataCombo1(20) & "',�ƻ�='" & DataCombo1(7) & "',����='" & DTPicker3.value & "',����='" & DTPicker4.value & "',��ע='" & Text4 & "',���='" & lb & "',��̨='" & DataCombo1(18) & "' where ֯��='" & DataCombo1(3).Text & "' and ����='" & DataCombo1(10) & "' and ����='" & Adodc1.Recordset.Fields(1) & "'"
sql2 = "update zbkpd set ��̨=replace(��̨,'" & Adodc1.Recordset.Fields(1) & "','" & DataCombo1(20) & "') where ֯��='" & DataCombo1(3).Text & "'"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

'sql1 = "update zbkpd set ֯��='" & DataCombo1(11) & "',�۷�='" & DataCombo1(12).Text & "',ת��='" & DataCombo1(22).Text & "',���='" & DataCombo1(30) & "',�ʼ�='" & DataCombo1(32).Text & "',����='" & DataCombo1(38).Text & "'  where ֯��='" & DataCombo1(3).Text & "'"
'sql2 = "update zbkpdgybh set ����=����ͼƬ  where ֯��='" & DataCombo1(3).Text & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
'RD.Open sql2, conn, adOpenStatic, adLockOptimistic
'MsgBox ("����ɹ���")

'Else
'If MsgBox("ȷ��������", vbYesNo) = vbNo Then Exit Sub

'sql1 = "update zbkpd set ֯��='" & DataCombo1(11) & "',�۷�='" & DataCombo1(12).Text & "',ת��='" & DataCombo1(22).Text & "',���='" & DataCombo1(30) & "',�ʼ�='" & DataCombo1(32).Text & "',����='" & DataCombo1(38).Text & "'  where ֯��='" & DataCombo1(3).Text & "'"
'sql2 = "update zbkpdgybh set ����=����ͼƬ  where ֯��='" & DataCombo1(3).Text & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
'RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("����ɹ���")
Adodc1.RecordSource = "SELECT * FROM kpd_jtjh where  ����='" & DataCombo1(10).Text & "' order by ��� desc"
Adodc1.Refresh
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
DataCombo1(23).Text = 1
Else
DataCombo1(23).Text = Adodc3.Recordset.Fields(0) + 1
End If

Adodc13.RecordSource = "SELECT isnull(Ƿ����,0) FROM v_zbkpd_kpd_jtjh where  ֯��='" & DataCombo1(3).Text & "'"
Adodc13.Refresh
If Not Adodc13.Recordset.EOF Then
DataCombo1(7) = Adodc13.Recordset.Fields(0)
Else
DataCombo1(7) = 0
End If
Text4 = ""

End Sub

Private Sub Command16_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If MsgBox("ȷ��ɾ����" + Adodc1.Recordset.Fields(1), vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Adodc1.RecordSource = "SELECT * FROM kpd_jtjh where  ����='" & DataCombo1(10).Text & "' order by ��� desc"
Adodc1.Refresh
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
DataCombo1(23).Text = 1
Else
DataCombo1(23).Text = Adodc3.Recordset.Fields(0) + 1
End If

Adodc13.RecordSource = "SELECT isnull(Ƿ����,0) FROM v_zbkpd_kpd_jtjh where  ֯��='" & DataCombo1(3).Text & "'"
Adodc13.Refresh
If Not Adodc13.Recordset.EOF Then
DataCombo1(7) = Adodc13.Recordset.Fields(0)
Else
DataCombo1(7) = 0
End If
Text4 = ""

End Sub

Private Sub Command2_Click()
On Error Resume Next
If DataCombo1(3).Text = "" Then
MsgBox ("������֯��")
Exit Sub
End If

If Val(DataCombo1(7).Text) <= 0 Then
MsgBox ("�������Ų���")
Exit Sub
End If


If DataCombo1(20).Text = "" Then
MsgBox ("�����복��")
Exit Sub
End If

If DataCombo1(3).Text = "" Or DataCombo1(7).Text = "" Then
MsgBox ("�������Ų����ƻ���ת����Ϣ")
Exit Sub
End If

If Option1.value = True Then
lb = "����"
End If

If Option2.value = True Then
lb = "ί��"
End If


If MsgBox("ȷ��������", vbYesNo) = vbNo Then Exit Sub
sql1 = "insert into  kpd_jtjh(֯��,����,�ƻ�,����,֯��,�ʼ�,�۷�,ת��,����,����,����,��ע,����,���,���,��̨) select ֯��,'" & DataCombo1(20) & "','" & DataCombo1(7) & "',����,֯��,�ʼ�,�۷�,ת��,����,'" & DTPicker3.value & "','" & DTPicker4.value & "','" & Text4 & "','" & DataCombo1(10) & "','" & DataCombo1(23) & "','" & lb & "','" & DataCombo1(18) & "' from zbzbkpd  where ֯��='" & DataCombo1(3).Text & "'"
sql2 = "update zbkpd set ��̨=��̨+'-'+'" & DataCombo1(20) & "' where ֯��='" & DataCombo1(3).Text & "' and  ��̨ not like '%'+'" & DataCombo1(20) & "'+'%'"
sql3 = "update zbkpd set ��̨=right(��̨,len(��̨)-1) where ֯��='" & DataCombo1(3).Text & "' and left(��̨,1)='-'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic

'sql1 = "update zbkpd set ֯��='" & DataCombo1(11) & "',�۷�='" & DataCombo1(12).Text & "',ת��='" & DataCombo1(22).Text & "',���='" & DataCombo1(30) & "',�ʼ�='" & DataCombo1(32).Text & "',����='" & DataCombo1(38).Text & "'  where ֯��='" & DataCombo1(3).Text & "'"
'sql2 = "update zbkpdgybh set ����=����ͼƬ  where ֯��='" & DataCombo1(3).Text & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
'RD.Open sql2, conn, adOpenStatic, adLockOptimistic
'MsgBox ("����ɹ���")

'Else
'If MsgBox("ȷ��������", vbYesNo) = vbNo Then Exit Sub

'sql1 = "update zbkpd set ֯��='" & DataCombo1(11) & "',�۷�='" & DataCombo1(12).Text & "',ת��='" & DataCombo1(22).Text & "',���='" & DataCombo1(30) & "',�ʼ�='" & DataCombo1(32).Text & "',����='" & DataCombo1(38).Text & "'  where ֯��='" & DataCombo1(3).Text & "'"
'sql2 = "update zbkpdgybh set ����=����ͼƬ  where ֯��='" & DataCombo1(3).Text & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
'RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("����ɹ���")
Adodc1.RecordSource = "SELECT * FROM kpd_jtjh where  ����='" & DataCombo1(10).Text & "' order by ��� desc"
Adodc1.Refresh
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
DataCombo1(23).Text = 1
Else
DataCombo1(23).Text = Adodc3.Recordset.Fields(0) + 1
End If

Adodc13.RecordSource = "SELECT isnull(Ƿ����,0) FROM v_zbkpd_kpd_jtjh where  ֯��='" & DataCombo1(3).Text & "'"
Adodc13.Refresh
If Not Adodc13.Recordset.EOF Then
DataCombo1(7) = Adodc13.Recordset.Fields(0)
Else
DataCombo1(7) = 0
End If
Text4 = ""
End Sub


Private Sub Command3_Click()
If InStr(Text6.Text, "_") > 0 Then
If MsgBox("ȷ��ɾ��֯�ţ�" + Text6.Text, vbYesNo) = vbNo Then Exit Sub
sql1 = "delete from zbkpd where ֯��='" & Text6.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("ɾ���ɹ���")
End If
End Sub

Private Sub Command33_Click()
Unload Me
End Sub


Private Sub Command4_Click()
Text6.Text = DataCombo1(3).Text + "_"
End Sub

Private Sub Command5_Click()
On Error Resume Next
Forma24.Text1(0) = DataCombo1(3).Text
Forma24.Text1(1) = DataCombo1(7).Text
Forma24.Text1(2) = 20
Forma24.Text1(10) = DataCombo1(20).Text
Forma24.Text1(11) = DataCombo1(18).Text
Forma24.Text1(12) = DataCombo1(10).Text
Forma24.Show
End Sub

Private Sub Command6_Click()
'FormA101.Text1(0).Text = DataCombo1(3).Text
'FormA101.Text1(1).Text = DataCombo1(23).Text
'FormA101.Text1(2).Text = DataCombo1(10).Text
'FormA101.Text2.Text = DataCombo1(1).Text
'FormA101.Show
End Sub

Private Sub Command7_Click()
'Formy133.DataCombo5 = DataCombo1(1).Text
'Formy133.DataCombo4(18).Text = DataCombo1(3).Text
'Formy133.Show
End Sub

Private Sub Command8_Click()
If MsgBox("ȷ����������", vbYesNo) = vbNo Then Exit Sub

sql1 = "update zbkpd set ��̨='',�Ų�=null where ֯��='" & DataCombo1(3).Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("������ɹ���")

End Sub

Private Sub Command9_Click()
FormA102.Show
End Sub

Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 3

       Case 30
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "SELECT * FROM gytp where ���ձ��='" & DataCombo1(30).Text & "'"
Adodc9.Refresh

       
    Image1.Picture = Nothing

If Adodc9.Recordset.Fields(3).Type = 205 Then
     StrPicTemp = "c:\temp.tmp"     '��ʱ�ļ�,�������������ͼƬ
     With Stm
        .Type = adTypeBinary
        .Open
        .Write Adodc9.Recordset.Fields(3)        'д�����ݿ��е�������Stream��
        .SaveToFile StrPicTemp, adSaveCreateOverWrite   '��Stream������д����ʱ�ļ���
        .Close
    End With
    
    Image1.Picture = LoadPicture(StrPicTemp)

End If

End Select
End Sub


Private Sub Form_Load()
On Error Resume Next
For i = 0 To 38
DataCombo1(i).Text = ""
Next

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
DTPicker4.value = Date
Option1.value = True
DataCombo1(10).Enabled = False


Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT ��� FROM ZBCT  GROUP BY ���"
Adodc6.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT ��̨ FROM ZBCT  GROUP BY ��̨"
Adodc4.Refresh

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT ���ձ�� FROM gytp  GROUP BY ���ձ��"
Adodc8.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select MAX(right(����,2)) as h  from kpd_jtjh where ����=cast('" & DTPicker3.value & "' as datetime) AND left(����,1)='" & yhdm & "'"
Adodc2.Refresh
DataCombo1(10).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "01"
If Adodc2.Recordset.EOF Then
DataCombo1(10).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "01"
Else
uu = Val(Adodc2.Recordset.Fields(0)) + 1
Select Case Len(uu)
       Case "1"
DataCombo1(10).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "0" + Trim(uu)
       Case "2"
DataCombo1(10).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + Trim(uu)
End Select
End If

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select ��� from kpd_jtjh where ����='" & DataCombo1(10).Text & "' AND left(����,1)='" & yhdm & "' order by ��� desc"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
DataCombo1(23).Text = 1
Else
DataCombo1(23).Text = Adodc3.Recordset.Fields(0) + 1
End If

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM kpd_jtjh where  ����='" & DataCombo1(10).Text & "' order by ��� desc"
Adodc1.Refresh

VSFlexGrid1.ColWidth(0) = 300
For i = 1 To 5
VSFlexGrid1.ColWidth(i) = 1500
Next

End Sub


Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 12
DataCombo1(10).Enabled = False
End Select
End Sub

Private Sub Label2_DblClick(Index As Integer)
Select Case Index
       Case 12
DataCombo1(10).Enabled = True
End Select
End Sub

Private Sub Label3_Click(Index As Integer)
Select Case Index
       Case 6
DataCombo1(1).Enabled = True
       Case 3
DataCombo1(3).Enabled = True
End Select
End Sub

Private Sub Label3_DblClick(Index As Integer)
Select Case Index
       Case 6
DataCombo1(1).Enabled = False
       Case 3
DataCombo1(3).Enabled = False
End Select
End Sub

Private Sub Label5_Click()
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select * from v_pcl where �ƺ�=LEFT('" & DataCombo1(3).Text & "', 9)"
Adodc11.Refresh
If Adodc11.Recordset.EOF Then
MsgBox ("֯������")
Exit Sub
Else
DataCombo1(23).Text = Format(Adodc11.Recordset.Fields(1) - Adodc11.Recordset.Fields(2), "#0.00")
End If
End Sub

Private Sub Option1_Click()
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT ��̨ FROM ZBCT  GROUP BY ��̨"
Adodc4.Refresh
End Sub

Private Sub Option2_Click()
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT distinct ��� as ��̨ FROM GYS where IP like '%Z%'"
Adodc4.Refresh
DataCombo1(18) = ""
End Sub

Private Sub Text3_Change()
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "SELECT �������� FROM CLMC where �������� like '%'+'" & Text3.Text & "'+'%' GROUP BY ��������"
Adodc7.Refresh
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
DataCombo1(3).Text = Adodc1.Recordset.Fields(0)
DataCombo1(7) = Adodc1.Recordset.Fields(3)
End Sub

Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
End With
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If MsgBox("ȷ��ɾ����?", vbYesNo) = vbNo Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
Adodc1.Recordset.Delete
Adodc1.Refresh
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move r - 1
Adodc1.Recordset.Fields(c - 1) = Combo1111.Text
Adodc1.Recordset.Update

    VSFlexGrid1.Text = Combo1111.Text
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub


