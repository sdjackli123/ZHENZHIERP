VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formy59 
   BackColor       =   &H00C0E0FF&
   Caption         =   "ɴ����Ϣ"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   13530
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   25
      Text            =   "Text3"
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
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
      TabIndex        =   14
      Top             =   2415
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
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
      TabIndex        =   13
      Top             =   1935
      Width           =   1095
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1935
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   5520
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   255
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   10320
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ѯ����"
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "��Ӧ"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ʒ��"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2415
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   6360
      Top             =   10455
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
      Left            =   6840
      Top             =   10455
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   7920
      TabIndex        =   15
      Top             =   495
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formy59.frx":0000
      Height          =   330
      Left            =   5520
      TabIndex        =   16
      Top             =   495
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   7800
      Top             =   10575
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
      Height          =   495
      Left            =   7560
      Top             =   10455
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
      Left            =   8520
      Top             =   10455
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   495
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   330039297
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   1215
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   330039297
      CurrentDate     =   39883
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formy59.frx":0015
      Height          =   330
      Left            =   3000
      TabIndex        =   19
      Top             =   495
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Formy59.frx":002A
      Height          =   330
      Left            =   3000
      TabIndex        =   20
      Top             =   1215
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "Formy59.frx":003F
      Height          =   330
      Left            =   5520
      TabIndex        =   21
      Top             =   1215
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Bindings        =   "Formy59.frx":0054
      Height          =   330
      Left            =   7920
      TabIndex        =   22
      Top             =   1215
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Bindings        =   "Formy59.frx":0069
      Height          =   330
      Left            =   10320
      TabIndex        =   23
      Top             =   495
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "��������"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formy59.frx":007E
      Height          =   6255
      Left            =   720
      TabIndex        =   24
      Top             =   3120
      Width           =   11895
      _cx             =   20981
      _cy             =   11033
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
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "��Ӧ��λ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   35
      Top             =   255
      Width           =   1455
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   34
      Top             =   255
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   33
      Top             =   975
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ϼ�"
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
      Left            =   8880
      TabIndex        =   32
      Top             =   8895
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   31
      Top             =   255
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   30
      Top             =   255
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   29
      Top             =   975
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "���Ϲ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5520
      TabIndex        =   28
      Top             =   975
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   27
      Top             =   975
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   9
      Left            =   11400
      TabIndex        =   26
      Top             =   255
      Width           =   1215
   End
End
Attribute VB_Name = "Formy59"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r  As Long


Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Command3_Click()
On Error Resume Next
sql1 = ""
If Check2(0).value = 1 Then
sql1 = sql1 + "���� between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "��Ӧ��λ like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "���� like '%'+'" & DataCombo6.Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "��Լ�� like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "���� like '%'+'" & DataCombo3.Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "���� like '%'+'" & DataCombo4.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "���Ϲ�� like '%'+'" & DataCombo5.Text & "'+'%' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "�������� like '%'+'" & DataCombo7.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("��ѡ���ѯ����")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "SELECT ��Ӧ��λ,����,���ݺ�,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,�ϼƽ��  FROM clgl where (" + sql1 + ") ORDER BY ����,���ݺ�"
Adodc1.Refresh

Adodc3.RecordSource = "SELECT round(sum(�ϼƽ��),2) FROM CLGL where (" + sql1 + ")"
Adodc3.Refresh

Text3.Text = "0"
If Adodc3.Recordset.EOF Then
Text3.Text = "0"
Else
Text3.Text = Adodc3.Recordset.Fields(0)
End If


End Sub


Private Sub Command6_Click()
Call OutadodcToExcel(VSFlexGrid1, 11, DataCombo1.Text)
End Sub


Private Sub Form_Load()
Me.Caption = Me.Caption + "������� " + xdljb
Text3.Text = "0"
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select MC from CLKL  GROUP BY MC"
Adodc4.Refresh

Text4.Text = ""
Text5.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""
DataCombo7.Text = ""

DTPicker1.value = Date
DTPicker2.value = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1500
VSFlexGrid1.ColWidth(3) = 1500
VSFlexGrid1.ColWidth(4) = 1200
VSFlexGrid1.ColWidth(5) = 1200
VSFlexGrid1.ColWidth(6) = 1200
End Sub

Private Sub Text4_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select ��� from gys where ���� like '%'+'" & Text4.Text & "'+'%' group by ���"
Adodc2.Refresh

End Sub

Private Sub Text5_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select �������� from CLMC where ���� like '%'+'" & Text5.Text & "'+'%'  GROUP BY ��������"
Adodc5.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
Forma105.Text1(9).Text = Adodc1.Recordset.Fields(6)
Forma105.Text1(10).Text = Adodc1.Recordset.Fields(4)
Forma105.Text1(2).Text = Adodc1.Recordset.Fields(3)
Forma105.Text1(6).Text = Adodc1.Recordset.Fields(7)
Unload Me
End Sub


