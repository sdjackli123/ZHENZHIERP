VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formy150 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�ܲ����̵����"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form49"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "����ģʽ"
      Height          =   1335
      Left            =   4320
      TabIndex        =   20
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "��ѯ"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   8640
      Style           =   1  'Simple Combo
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ĺ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   9000
      Top             =   10560
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
      Height          =   495
      Left            =   9120
      Top             =   10320
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Left            =   8400
      Top             =   10440
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Left            =   9000
      Top             =   10440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   9120
      Top             =   10440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Height          =   495
      Left            =   9360
      Top             =   10440
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formy150.frx":0000
      Height          =   330
      Left            =   1680
      TabIndex        =   15
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "YS"
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formy150.frx":0015
      Height          =   330
      Left            =   1680
      TabIndex        =   14
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   "��������"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formy150.frx":002A
      Height          =   330
      Left            =   1680
      TabIndex        =   13
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�鿴����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���½��ת����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   10200
      TabIndex        =   9
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   330760193
      CurrentDate     =   39921
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
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
      Height          =   495
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ۿ��תʵ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���½���ת"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ղ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formy150.frx":003F
      Height          =   8175
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   17295
      _cx             =   30506
      _cy             =   14420
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
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "ת������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   10200
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������ɫ"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "���������"
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
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ѡ�����"
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
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Formy150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public c, r As Integer
Dim cdbhf As Integer
Private Sub Command1_Click()
Call OutadodcToExcel3(VSFlexGrid1, 11, 13, 15, "�̴��ӡ")
End Sub
Private Sub Command11_Click()
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "select * from CLPC WHERE ���� in(select mc from clkl where yh='" & yhm & "') order by ����,����"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select * from CLPC where ����='" & DataCombo1.Text & "' order by ����"
Adodc1.Refresh
End If

If Option1(0).value = True Then
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(4) = 0
VSFlexGrid1.ColWidth(5) = 0
VSFlexGrid1.ColWidth(6) = 0
VSFlexGrid1.ColWidth(7) = 1000
VSFlexGrid1.ColWidth(8) = 1000
VSFlexGrid1.ColWidth(9) = 1000
VSFlexGrid1.ColWidth(10) = 1000
VSFlexGrid1.ColWidth(11) = 1000
VSFlexGrid1.ColWidth(12) = 1000
VSFlexGrid1.ColWidth(13) = 1000
VSFlexGrid1.ColWidth(14) = 1000
VSFlexGrid1.ColWidth(15) = 1000
VSFlexGrid1.ColWidth(16) = 800
VSFlexGrid1.ColWidth(17) = 800
VSFlexGrid1.ColWidth(18) = 800
VSFlexGrid1.ColWidth(19) = 800
VSFlexGrid1.ColWidth(20) = 800
VSFlexGrid1.ColWidth(21) = 800
VSFlexGrid1.ColWidth(22) = 800
VSFlexGrid1.ColWidth(23) = 800
End If

If Option1(1).value = True Then
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 2500
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(4) = 0
VSFlexGrid1.ColWidth(5) = 0
VSFlexGrid1.ColWidth(6) = 0
VSFlexGrid1.ColWidth(7) = 0
VSFlexGrid1.ColWidth(8) = 0
VSFlexGrid1.ColWidth(9) = 0
VSFlexGrid1.ColWidth(10) = 0
VSFlexGrid1.ColWidth(11) = 0
VSFlexGrid1.ColWidth(12) = 0
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(15) = 0
VSFlexGrid1.ColWidth(16) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 1200
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(21) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
End If

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
VSFlexGrid1.Cell(flexcpBackColor, i, 8, i, 8) = &HFFFFC0
VSFlexGrid1.Cell(flexcpBackColor, i, 11, i, 11) = &HC0E0FF
VSFlexGrid1.Cell(flexcpBackColor, i, 14, i, 14) = &HFFFFC0
VSFlexGrid1.Cell(flexcpBackColor, i, 17, i, 17) = &HC0E0FF
VSFlexGrid1.Cell(flexcpBackColor, i, 20, i, 20) = &HFFFFC0
VSFlexGrid1.Cell(flexcpBackColor, i, 23, i, 23) = vbRed
Next
End If
End Sub


Private Sub Command2_Click()
sql1 = "UPDATE CLPC SET ʵ�ʽ��=round(ʵ�ʵ���*ʵ�ʿ��,2),�������=round(���ۿ��-ʵ�ʿ��,2) WHERE ���� in(select mc from clkl where yh='" & yhm & "')"
sql2 = "UPDATE CLPC SET ��Ľ��=���۽��-ʵ�ʽ�� WHERE ���� in(select mc from clkl where yh='" & yhm & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
MsgBox ("��ĺ���ɹ���")
End Sub

Private Sub Command3_Click()
If MsgBox("ȷ������̴�������", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CLPC "
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
On Error Resume Next

Adodc1.RecordSource = "SELECT * FROM CLPC"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
If MsgBox("��ת���¼���Ƿ��ת", vbYesNo) = vbNo Then Exit Sub
Else
If MsgBox("��ȷ�ϣ�ת�����¼������Ϊ��" + Str(DTPicker1.value), vbYesNo) = vbYes Then
If MsgBox("ȷ����ת��", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete  from cljl where ����=CAST('" & DTPicker1.value & "' AS DATETIME) and ���� in(select mc from clkl where yh='" & yhm & "')"
sql2 = "INSERT INTO  cljl (��Ӧ��λ,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,���,����,����)  SELECT ��Ӧ��λ,��������,���Ϲ��,���ϵ�λ,��ɫ,����,isnull(���۵���,0),isnull(ʵ�ʿ��,0),isnull(ʵ�ʽ��,0),����,'" & DTPicker1.value & "' FROM CLPC where isnull(ʵ�ʿ��,0)<>0 and ���� in(select mc from clkl where yh='" & yhm & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("ת��ɹ���")
Else
MsgBox ("ת��δ�ɹ���")
End If
End If
End Sub

Private Sub Command5_Click()
sql1 = "UPDATE CLPC SET ʵ�ʿ��=���ۿ�� where  ���� in(select mc from clkl where yh='" & yhm & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
MsgBox ("ת��ɹ���")
End Sub


Private Sub Command6_Click()
If DataCombo1.Text <> "" Then
Adodc1.RecordSource = "select * from CLPC where �������� like '%'+'" & DataCombo2.Text & "'+'%' and ���� in(select mc from clkl where yh='" & yhm & "') order by ����,����"
Adodc1.Refresh
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
'On Error Resume Next
Adodc1.RecordSource = "SELECT * FROM CLPC where ���� in(select mc from clkl where yh='" & yhm & "') ORDER BY ��������,��ɫ"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox ("��ת���¼����ֹ")
Exit Sub
Else
If MsgBox("��ȷ�ϣ�ת�뱨������Ϊ��" + Str(DTPicker1.value), vbYesNo) = vbYes Then
If MsgBox("ȷ��ת����", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete  from clbbzcw where ����=cast('" & DTPicker1.value & "' as datetime) and ���� in(select mc from clkl where yh='" & yhm & "')"
sql2 = "INSERT INTO clbbzcw(��Ӧ��λ,����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,���½������,���½�浥��,���½����,�����������,������ⵥ��,���������,���³�������,���³��ⵥ��,���³�����,���½������,���½�浥��,���½����,�������,��Ľ��,����,����)  SELECT ��Ӧ��λ,����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,���½������,���½�浥��,���½����,�����������,������ⵥ��,���������,���³�������,���³��ⵥ��,���³�����,ʵ�ʿ��,���۵���,ʵ�ʽ��,�������,��Ľ��,'" & DTPicker1.value & "',���� FROM CLPC where ���� in(select mc from clkl where yh='" & yhm & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("ת��ɹ���")
Else
MsgBox ("ת��δ�ɹ���")
End If
End If
End Sub

Private Sub Command9_Click()
Adodc1.RecordSource = "select * from clbbzcw where ����=CAST('" & DTPicker1.value & "' AS datetime) and ���� in(select mc from clkl where yh='" & yhm & "') order by ����,����"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(4) = 0
VSFlexGrid1.ColWidth(5) = 0
VSFlexGrid1.ColWidth(6) = 0
VSFlexGrid1.ColWidth(7) = 1000
VSFlexGrid1.ColWidth(9) = 1000
VSFlexGrid1.ColWidth(10) = 1000
VSFlexGrid1.ColWidth(11) = 1000
VSFlexGrid1.ColWidth(12) = 1000
VSFlexGrid1.ColWidth(13) = 1000
VSFlexGrid1.ColWidth(14) = 1000
VSFlexGrid1.ColWidth(15) = 1000
VSFlexGrid1.ColWidth(16) = 800
VSFlexGrid1.ColWidth(17) = 800
VSFlexGrid1.ColWidth(18) = 800
VSFlexGrid1.ColWidth(19) = 800
VSFlexGrid1.ColWidth(20) = 800
VSFlexGrid1.ColWidth(21) = 800
VSFlexGrid1.ColWidth(22) = 800
VSFlexGrid1.ColWidth(23) = 800
End Sub
Private Sub Form_Load()
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DTPicker1.value = Date
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Option1(1).value = True
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from CLPC where  ���� in(select mc from clkl where yh='" & yhm & "')"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT MC FROM CLKL where yh='" & yhm & "' GROUP BY MC"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT �������� FROM CLPC GROUP BY ��������"
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT YS.YS FROM YS GROUP BY YS.YS"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(4) = 0
VSFlexGrid1.ColWidth(5) = 0
VSFlexGrid1.ColWidth(6) = 0
VSFlexGrid1.ColWidth(7) = 800
VSFlexGrid1.ColWidth(9) = 1000
VSFlexGrid1.ColWidth(10) = 1000
VSFlexGrid1.ColWidth(11) = 1000
VSFlexGrid1.ColWidth(12) = 1000
VSFlexGrid1.ColWidth(13) = 1000
VSFlexGrid1.ColWidth(14) = 1000
VSFlexGrid1.ColWidth(15) = 1000
VSFlexGrid1.ColWidth(16) = 800
VSFlexGrid1.ColWidth(17) = 800
VSFlexGrid1.ColWidth(18) = 800
VSFlexGrid1.ColWidth(19) = 800
VSFlexGrid1.ColWidth(20) = 800
VSFlexGrid1.ColWidth(21) = 800
VSFlexGrid1.ColWidth(22) = 800


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

Private Sub VSFlexGrid1_Click()
FD = VSFlexGrid1.col
End Sub

Private Sub VSFlexGrid1_dblClick()
With VSFlexGrid1
    c = .col: r = .Row
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call VSFlexGrid1_dblClick
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
If c > 8 Then
Adodc1.Recordset.Fields(c - 1) = Val(Combo1111.Text)
VSFlexGrid1.Text = Val(Combo1111.Text)
Else
Adodc1.Recordset.Fields(c - 1) = Combo1111.Text
VSFlexGrid1.Text = Combo1111.Text
End If
Adodc1.Recordset.Update
Combo1111.Visible = False
VSFlexGrid1.SetFocus
End If
End Sub

