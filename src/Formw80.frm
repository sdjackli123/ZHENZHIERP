VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw80 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�ͻ���Ŀ��ѯ---�տ�"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   375
      Left            =   13080
      Top             =   10560
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ѯ"
      Height          =   855
      Left            =   11160
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   9000
      TabIndex        =   22
      Top             =   360
      Width           =   2055
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�ͻ�"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "˾��"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6360
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   960
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   8160
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
      Left            =   8400
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
      Left            =   8400
      Top             =   10680
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw80.frx":0000
      Height          =   7575
      Left            =   600
      TabIndex        =   19
      Top             =   1560
      Width           =   21015
      _cx             =   37068
      _cy             =   13361
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formw80.frx":0015
      Height          =   330
      Left            =   6960
      TabIndex        =   18
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw80.frx":002A
      Height          =   330
      Left            =   6960
      TabIndex        =   17
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "˾��"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ɲ�ѯ"
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ƾ֤����"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ת����"
      Height          =   735
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ˢ"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   329056257
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   329056257
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12480
      TabIndex        =   11
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   329056257
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   329056257
      CurrentDate     =   36892
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9000
      Top             =   10440
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
      Left            =   9000
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8520
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8640
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   8760
      Top             =   10680
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
      Left            =   8280
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   375
      Left            =   8520
      Top             =   10560
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formw80.frx":0040
      Height          =   3255
      Left            =   600
      TabIndex        =   20
      Top             =   9240
      Width           =   7695
      _cx             =   13573
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   30
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
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "�������"
      Height          =   375
      Index           =   0
      Left            =   12480
      TabIndex        =   12
      Top             =   600
      Width           =   1455
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
      Index           =   1
      Left            =   5880
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "˾��"
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
      Index           =   0
      Left            =   5880
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���ڷ�Χ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Formw80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public S1, S2 As Integer
Dim cdbhf As Integer
Private Sub Command1_Click()
'On Error Resume Next
Command1.Enabled = False
FP = CDate(DTPicker4.value) + 1
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "YSHZCX('" & DTPicker3.value & "','" & DTPicker4.value & "','" & FP & "','" & yhm & "')"      ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
g_Cmd.Cancel
Command1.Enabled = True
MsgBox ("ˢ�³ɹ���")
Adodc6.RecordSource = "SELECT * FROM YSKZCX where ���='" & yhm & "' order by �ͻ�"
Adodc6.Refresh
Adodc5.RecordSource = "SELECT ˾��, round(sum(isnull(�����ۼ�δ��,0)),2) as �ϼƱ����ۼ�δ��,round(sum(isnull(Ƿ��,0)),2) as �ϼ�Ƿ�� FROM YSKZCX where ���='" & yhm & "' group by ˾��"
Adodc5.Refresh

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid2.ColWidth(0) = 300
VSFlexGrid2.ColWidth(1) = 1200
VSFlexGrid2.ColWidth(2) = 1500
VSFlexGrid2.ColWidth(3) = 1500
VSFlexGrid2.SubtotalPosition = flexSTBelow
VSFlexGrid2.Subtotal flexSTSum, 0, 2, , vbGreen
VSFlexGrid2.Subtotal flexSTSum, 0, 3, , vbGreen
For i = 1 To 11
VSFlexGrid1.ColWidth(i) = 1300
Next
VSFlexGrid1.ColWidth(12) = 2000
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call OutadodcToExcel11(VSFlexGrid1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, "��ˮ���˷������޹�˾ �ͻ���Ŀ��ѯ--�տ�" + "��ֹ����:" + Trim(DTPicker4.value))
End Sub



Private Sub Command4_Click()
On Error Resume Next

sql1 = ""

If Check4.value = 1 Then
sql1 = sql1 + "�ͻ�='" & DataCombo2 & "' and "
End If

If Check6.value = 1 Then
sql1 = sql1 + "˾�� like '%'+'" & DataCombo1 & "'+'%' and "
End If


If Len(sql1) > 1 Then
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 3)
Adodc6.RecordSource = "SELECT * FROM YSKZCX where (" + sql1 + ") and ���='" & yhm & "' order by �ͻ�"
Adodc6.Refresh
Adodc5.RecordSource = "SELECT round(sum(isnull(�����ۼ�δ��,0)),2) as �ϼƱ����ۼ�δ��,round(sum(isnull(Ƿ��,0)),2) as �ϼ�Ƿ�� FROM YSKZCX where (" + sql1 + ") and ���='" & yhm & "'"
Adodc5.Refresh
End If
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If


VSFlexGrid1.ColWidth(0) = 300
For i = 1 To 11
VSFlexGrid1.ColWidth(i) = 1300
Next
VSFlexGrid1.ColWidth(12) = 2000

VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 1, 4, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 5, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 6, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 7, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 8, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 9, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 10, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 11, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 12, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 13, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 14, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 1, 15, , vbGreen

End Sub

Private Sub Command5_Click()
If MsgBox("ȷ����ת�������������Ϊ��" + Trim(DTPicker1.value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("ȷ����ת������?", vbYesNo) = vbNo Then Exit Sub

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "yshzjz('" & DTPicker1.value & "')"       ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
g_Cmd.Cancel

MsgBox ("��ת�ɹ��������ڳ������п��Բ�ѯ��")

End Sub

Private Sub Command6_Click()
If MsgBox("��������Ϊ��" + Trim(DTPicker2.value) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("�����ڼ�Ϊ��" + Trim(Month(DTPicker2.value)) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("ȷ������Ӧ��ϵ�е�ƾ֤��", vbYesNo) = vbNo Then Exit Sub
Call CPFHPZ(DTPicker3.value, DTPicker4.value, DTPicker2.value)
End Sub

Private Sub Command7_Click()
Formw1132.DTPicker1.value = DTPicker2.value
Formw1132.Show
Unload Me
End Sub

Private Sub DataCombo1_Click(Area As Integer)
'Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc6.RecordSource = "SELECT * FROM YSKZCX WHERE ҵ�� like '%'+'" & DataCombo1 & "'+'%' order by �ͻ�"
'Adodc6.Refresh
'Adodc5.RecordSource = "SELECT round(sum(isnull(�����ۼ�Ӧ��,0)),2) as �ϼ������ۼ�Ӧ��,round(sum(isnull(����Ӧ�տ�,0)),2) as �ϼƱ���Ӧ�տ�,round(sum(isnull(�����ۼ�Ӧ�տ�,0)),2) as �ϼƱ����ۼ�Ӧ�տ�,round(sum(isnull(�������տ�,0)),2) as �ϼƱ������տ�,round(sum(isnull(�����ۼƿ�Ʊ,0)),2) as �ϼ������ۼƿ�Ʊ,round(sum(isnull(���ڿ�Ʊ,0)),2) as �ϼƱ��ڿ�Ʊ,round(sum(isnull(�����ۼƿ�Ʊ,0)),2) as �ϼƱ����ۼƿ�Ʊ,round(sum(isnull(�����ۼ�δ��Ʊ,0)),2) as �ϼ������ۼ�δ��Ʊ,round(sum(isnull(����δ��,0)),2) as �ϼƱ���δ��,round(sum(isnull(�����ۼ�δ��,0)),2) as �ϼƱ����ۼ�δ��,round(sum(isnull(Ƿ��,0)),2) as �ϼ�Ƿ�� FROM YSKZCX WHERE ҵ�� like '%'+'" & DataCombo1 & "'+'%'"
'Adodc5.Refresh
End Sub


Private Sub Form_Load()
On Error Resume Next


DTPicker1.value = Date
DTPicker3.value = Date
DTPicker2.value = Date
DTPicker4.value = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
Text1 = ""
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ����,��� from KHZL  GROUP BY ����,���"
Adodc1.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT * FROM YSKZCX WHERE �ͻ�='" & DataCombo2.Text & "'  order by �ͻ�"
Adodc6.Refresh
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select ip from KHZL  GROUP BY ip"
Adodc8.Refresh
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select distinct ˾�� from YSKZCX  "
Adodc11.Refresh
VSFlexGrid1.ColWidth(0) = 300
For i = 1 To 11
VSFlexGrid1.ColWidth(i) = 1300
Next
VSFlexGrid1.ColWidth(12) = 2000

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
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ��� from KHZL where ����  like '%'+'" & Text1 & "'+'%'  group by ���"
Adodc1.Refresh
End Sub

Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
S1 = VSFlexGrid1.RowSel
End Sub

Private Sub VSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
S2 = VSFlexGrid1.RowSel
End Sub

Private Sub CPFHPZ(DT1 As Date, dt2 As Date, dt3 As Date)
'On Error Resume Next
Dim djs As Integer

Adodc4.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN '" & DT1 & "' AND '" & dt2 & "' and �Ƶ� like '%�Զ�-����%'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
If MsgBox("����Ӧ������ƾ֤���Ƿ��������ɣ�", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CLZZPZ WHERE �Ƶ� like '%�Զ�-����%' and ���� BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If


Adodc9.RecordSource = "SELECT * FROM YSKZCX where ����Ӧ�տ�<>0 order by �ͻ�"
Adodc9.Refresh

If Adodc9.Recordset.EOF Then Exit Sub

Adodc10.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc10.Refresh
If Adodc10.Recordset.EOF Then
PZH = "5-1"
Else
Adodc10.RecordSource = "SELECT MAX(right(ƾ֤��,len(ƾ֤��)-2)) FROM CLZZPZ WHERE ���� BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc10.Refresh
PZH = "5-" + Trim(Adodc10.Recordset.Fields(0) + 1)
End If
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 7

sql1 = "INSERT INTO CLZZPZ(ժҪ,���˿�Ŀ,��ϸ��Ŀ,�跽���,�������,ƾ֤��,����,ԭʼ����,����,����,�Ƶ�,ԭʼ������,���ȷ��,���˱��) Values('��Ӫ����','Ӧ���˿�','" & Adodc9.Recordset.Fields(1) & "','" & Adodc9.Recordset.Fields(3) & "','','" & PZH & "','" & dt3 & "','','','','�Զ�-����','','','')"
sql2 = "INSERT INTO CLZZPZ(ժҪ,���˿�Ŀ,��ϸ��Ŀ,�跽���,�������,ƾ֤��,����,ԭʼ����,����,����,�Ƶ�,ԭʼ������,���ȷ��,���˱��) Values('��Ӫ����','��Ӫҵ������','','','" & Adodc9.Recordset.Fields(3) & "','" & PZH & "','" & dt3 & "','','','','�Զ�-����','','','')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
Exit Sub
End If
Next
Adodc10.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc10.Refresh
If Adodc10.Recordset.EOF Then
PZH = "5-1"
Else
Adodc10.RecordSource = "SELECT MAX(right(ƾ֤��,len(ƾ֤��)-2)) FROM CLZZPZ WHERE ���� BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc10.Refresh
PZH = "5-" + Trim(Adodc10.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("��Ʒ������ת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
End Sub





