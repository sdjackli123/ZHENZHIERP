VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma188 
   BackColor       =   &H00C0E0FF&
   Caption         =   "ë����ϸ"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20250
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6960
      TabIndex        =   36
      Text            =   "Text6"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6960
      TabIndex        =   33
      Top             =   1800
      Width           =   1215
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma188.frx":0000
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   19815
      _cx             =   34951
      _cy             =   11880
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
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4320
      TabIndex        =   29
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1935
      Left            =   9720
      TabIndex        =   9
      Top             =   240
      Width           =   3975
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "֯��"
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   32
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "��ͬ"
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1440
         TabIndex        =   23
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "��桷"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�ͻ�"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ʒ��"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      Height          =   495
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ѯ"
      Height          =   495
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      Height          =   495
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ת"
      Height          =   615
      Left            =   16680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   7440
      Top             =   8880
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
      Left            =   7440
      Top             =   9000
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
      Left            =   7440
      Top             =   8880
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
      Height          =   375
      Left            =   7440
      Top             =   9000
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma188.frx":0015
      Height          =   330
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   15240
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   329187329
      CurrentDate     =   36892
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   960
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   329187329
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   329187329
      CurrentDate     =   39177
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "Forma188.frx":002A
      Height          =   330
      Left            =   960
      TabIndex        =   26
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "xm"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma188.frx":003F
      Height          =   1215
      Left            =   240
      TabIndex        =   31
      Top             =   9000
      Width           =   19815
      _cx             =   34951
      _cy             =   2143
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
      FormatString    =   $"Forma188.frx":0054
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
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6000
      TabIndex        =   35
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "֯��"
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
      Index           =   6
      Left            =   6000
      TabIndex        =   34
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��ͬ"
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
      Index           =   4
      Left            =   3360
      TabIndex        =   30
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
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
      Index           =   5
      Left            =   360
      TabIndex        =   27
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
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
      Index           =   3
      Left            =   3360
      TabIndex        =   17
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��������"
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
      Left            =   2520
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��ʼ����"
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
      Left            =   2520
      TabIndex        =   15
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
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
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��ѡ��ͻ�"
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
      Left            =   360
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ת����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   15240
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Forma188"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BAR As Integer
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Private Sub Command1_Click()
On Error Resume Next

sql = ""

If Check1(2).value = 1 Then
sql = sql + "�ͻ����� like '%'+ '" & DataCombo1.Text & "' +'%'" + " and "
End If

If Check1(1).value = 1 Then
sql = sql + "���� like '%'+ '" & DataCombo2.Text & "'+'%'" + " and "
End If

If Check1(0).value = 1 Then
sql = sql + "���� between  cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If Check1(3).value = 1 Then
sql = sql + "���� like '%'+ '" & Text1 & "'+'%'" + " and "
End If

If Check1(4).value = 1 Then
sql = sql + "���ƥ��> cast('" & Text2 & "' as real) and "
End If

If Check1(5).value = 1 Then
sql = sql + "������='" & DataCombo5 & "' and "
End If

If Check1(6).value = 1 Then
sql = sql + "��Լ�� like '%'+'" & Text4 & "'+'%' and "
End If

If Check1(7).value = 1 Then
sql = sql + "֯�� like '%'+'" & Text5 & "'+'%' and "
End If

If Check1(8).value = 1 Then
sql = sql + "���� like '%'+'" & Text6 & "'+'%' and "
End If

If Len(sql) > 1 Then
sql = Left$(Trim(sql), Len(Trim(sql)) - 3)
Adodc2.RecordSource = "select ���ݺ�,����,����,�ͻ�����,����,���ƥ��,�������,����ƥ��,��������,���ƥ��,�������,��ת,�˿�ƥ��,�˿�����,����ƥ��,��������,����ƥ��,��������,�������ƥ��,�����������,֯��,���λ��  from v_mp_kc where (" + sql + ")  order by ����,���ݺ�,����"
Adodc2.Refresh
Adodc4.RecordSource = "select round(sum(isnull(���ƥ��,0)),1) as ���ƥ��,round(sum(�������),2) as �������,round(sum(isnull(����ƥ��,0)),1) as ����ƥ��,round(sum(��������),2) as ��������,round(sum(isnull(���ƥ��,0)),1) as ���ƥ��,round(sum(�������),2) as �������,round(sum(��ת),2) as ��ת����,round(sum(��������),2) as ��������,round(sum(����ƥ��),2) as �������ƥ��,round(sum(��������),2) as �����������,round(sum(����ƥ��),2) as ��������ƥ��,round(sum(��������),2) as ������������,round(sum(�������ƥ��),2) as �������ƥ��,round(sum(�����������),2) as ����������� from v_mp_kc where (" + sql + ") "
Adodc4.Refresh
End If

Call gssx
End Sub

Private Sub Command2_Click()
Call OutadodcToExcel22(VSFlexGrid1, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, DataCombo1.Text + "ë�����")
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If MsgBox("ȷ����ת�𣿽�ת��������Ϊ" + Trim(DTPicker3.value), vbYesNo) = vbNo Then Exit Sub
    Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "MPKCJZ('" & DTPicker3.value & "')"       ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
MsgBox ("��ת�ɹ���")
End Sub

Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Form_Load()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
cdbhf = cdbh
Call SetDeviceIndependentWindow(Me) '�жϵ�ǰ�ֱ��ʺ����ʱ�ķֱ����Ƿ���ͬ
suiping = Screen.Width / Screen.TwipsPerPixelX  '���㵱ǰ��ˮƽ�ֱ���
cuizhi = Screen.Height / Screen.TwipsPerPixelY '���㵱ǰ�Ĵ�ֱ�ֱ���
If fbl = 1 Then    '��ǰ�ֱ��ʺ����ʱ�ķֱ��ʲ���ͬ
Call ResizeInit(Me)    '����ԭ��������ֵ
Call ResizeForm(Me)    '����������
VSFlexGrid1.FontSize = VSFlexGrid1.FontSize * (suiping / 1366)  ' ��������Ӧ�ĵ���
VSFlexGrid2.FontSize = VSFlexGrid2.FontSize * (suiping / 1366)  ' ��������Ӧ�ĵ���

For i = 0 To 7
Label1(i).FontSize = Label1(i).FontSize * suiping / 1366
Next
Label6(0).FontSize = Label6(0).FontSize * suiping / 1366
Label7(0).FontSize = Label7(0).FontSize * suiping / 1366

DataCombo1.Font.Size = DataCombo1.Font.Size * suiping / 1366
DataCombo2.Font.Size = DataCombo2.Font.Size * suiping / 1366
DataCombo5.Font.Size = DataCombo5.Font.Size * suiping / 1366

DTPicker1.Font.Size = DTPicker1.Font.Size * (suiping / 1366)
DTPicker2.Font.Size = DTPicker2.Font.Size * (suiping / 1366)
DTPicker3.Font.Size = DTPicker3.Font.Size * (suiping / 1366)

Frame1.FontSize = Frame1.FontSize * (suiping / 1366)

For i = 0 To 8
Check1(i).FontSize = Check1(i).FontSize * (suiping / 1366)
Next

Command1.FontSize = Command1.FontSize * (suiping / 1366)
Command2.FontSize = Command2.FontSize * (suiping / 1366)
Command3.FontSize = Command3.FontSize * (suiping / 1366)
Command4.FontSize = Command4.FontSize * (suiping / 1366)

Text1.FontSize = Text1.FontSize * (suiping / 1366)
Text2.FontSize = Text2.FontSize * (suiping / 1366)
Text3.FontSize = Text3.FontSize * (suiping / 1366)
Text4.FontSize = Text4.FontSize * (suiping / 1366)
Text5.FontSize = Text5.FontSize * (suiping / 1366)
Text6.FontSize = Text6.FontSize * (suiping / 1366)

End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
DTPicker1.value = Date
DTPicker2.value = Date
Text1 = ""
Text2 = 0
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
cdbhf = cdbh
Check1(0).value = 1
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ��� from KHZL where ip like '%'+'" & yhxx & "'+'%' GROUP BY ���"
Adodc1.Refresh
Adodc2.CommandTimeout = 10000
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select xm  from fzr group by xm"
Adodc3.Refresh
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc2.RecordSource = "select ���ݺ�,����,����,�ͻ�����,����,���ƥ��,�������,����ƥ��,��������,���ƥ��,�������,��ת,�˿�ƥ��,�˿�����,����ƥ��,��������,����ƥ��,��������,�������ƥ��,�����������,֯��,���λ�� from v_mp_kc where ����=cast('" & Date & "' as datetime)  order by ����,���ݺ�,����"
Adodc2.Refresh
Adodc4.RecordSource = "select round(sum(isnull(���ƥ��,0)),1) as ���ƥ��,round(sum(�������),2) as �������,round(sum(isnull(����ƥ��,0)),1) as ����ƥ��,round(sum(��������),2) as ��������,round(sum(isnull(���ƥ��,0)),1) as ���ƥ��,round(sum(�������),2) as �������,round(sum(��ת),2) as ��ת����,round(sum(��������),2) as ��������,round(sum(����ƥ��),2) as �������ƥ��,round(sum(��������),2) as �����������,round(sum(����ƥ��),2) as ��������ƥ��,round(sum(��������),2) as ������������,round(sum(�������ƥ��),2) as �������ƥ��,round(sum(�����������),2) as ����������� from v_mp_kc where ����=cast('" & Date & "' as datetime)"
Adodc4.Refresh

DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo5.Text = ""
Text1.TabIndex = 0
Call gssx
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
sql2 = "insert into yhcd(�û�,�˵�,���) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic


End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sql2 = "delete from yhcd where �û�='" & yhm & "' and ���='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

End Sub

Private Sub Text3_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ��� from KHZL where ip like '%'+'" & yhxx & "'+'%' and ���� like '%'+'" & Text3 & "'+'%' GROUP BY ���"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
DataCombo1 = Adodc1.Recordset.Fields(0)
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub VSFlexGrid1_dblClick()
rs = VSFlexGrid1.Row
cl = VSFlexGrid1.col
If mmkc = 1 And cl = 1 Then
Forma11.DataCombo1 = VSFlexGrid1.TextMatrix(rs, 3) '''�ͻ�
'Forma11.DataCombo8 = VSFlexGrid1.TextMatrix(rs, 1)   '''���ݺ�
'Forma11.Text7 = VSFlexGrid1.TextMatrix(rs, 1)    ''����=ë�����ĵ��ݺ�
Forma11.Text16(0) = VSFlexGrid1.TextMatrix(rs, 1)  ''���ݺ�
Forma11.Text16(2) = VSFlexGrid1.TextMatrix(rs, 10) ''�������
Forma11.Text15 = VSFlexGrid1.TextMatrix(rs, 17) ''������ϸ
Forma11.Text10 = VSFlexGrid1.TextMatrix(rs, 12) ''���ϵ�λ
Forma11.DataCombo4(1) = VSFlexGrid1.TextMatrix(rs, 4)  ''Ʒ��
Forma11.DataCombo4(2) = VSFlexGrid1.TextMatrix(rs, 18)  ''ë������
Forma11.DataCombo4(4) = VSFlexGrid1.TextMatrix(rs, 9) ''�ƻ�ƥ��=���ƥ��
Forma11.DataCombo4(5) = VSFlexGrid1.TextMatrix(rs, 10) ''�ƻ�����=�������
'Forma11.DataCombo4(6) = VSFlexGrid1.TextMatrix(rs, 19) '''��ɫ
Forma11.Timer1.Enabled = False
Unload Me
End If
End Sub


Private Sub gssx()
With VSFlexGrid1

.BackColorAlternate = &HCDEEC6
.SelectionMode = flexSelectionListBox

.ColWidth(0) = 100 * (suiping / 1366)
.ColWidth(1) = 900 * (suiping / 1366)
.ColWidth(2) = 1000 * (suiping / 1366)
.ColWidth(3) = 900 * (suiping / 1366)
.ColWidth(4) = 1500 * (suiping / 1366)
For i = 5 To 19
.ColWidth(i) = 1000 * (suiping / 1366)
Next

.ColFormat(6) = "#0.#"
.ColFormat(8) = "#0.#"
.ColFormat(10) = "#0.#"
.TextMatrix(0, 0) = "��¼��"

'.SubtotalPosition = flexSTBelow
'.Subtotal flexSTSum, -1, 12, , vbWhite
'.Subtotal flexSTSum, -1, 13, , vbWhite
'.Subtotal flexSTCount, -1, 5, , vbWhite

'If .Rows > 2 Then
'.TextMatrix(.Rows - 1, 1) = "�ϼ�"
'End If

.RowHeight(0) = 400 * (cuizhi / 768)
If .Rows > 0 Then
For i = 1 To .Rows - 1
.RowHeight(i) = 400 * (cuizhi / 768)
.TextMatrix(i, 0) = i
Next
End If
If .Rows > 1 Then
.Row = 1
.GridLinesFixed = 14
Else
.GridLinesFixed = 1
End If
End With
For i = 1 To 14
VSFlexGrid2.ColWidth(i) = 1800
Next
End Sub

