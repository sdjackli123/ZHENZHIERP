VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formr332 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�ɱ�����"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   9000
      Style           =   1  'Simple Combo
      TabIndex        =   27
      Text            =   "Combo1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   11280
      TabIndex        =   26
      Text            =   "Text3"
      Top             =   1080
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8160
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
      Left            =   7800
      Top             =   9840
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
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   10680
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   600
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formr332.frx":0000
      Height          =   330
      Left            =   11280
      TabIndex        =   22
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command7 
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
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ѯ����"
      Height          =   1095
      Left            =   12960
      TabIndex        =   18
      Top             =   360
      Width           =   1215
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�ͻ�"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���º���"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "������Ϣ"
      Height          =   1095
      Left            =   5040
      TabIndex        =   13
      Top             =   480
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "ȫ��"
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "�Ѻ���"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "δ����"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command4 
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ת"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   960
      Width           =   1335
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formr332.frx":0015
      Height          =   7215
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   15975
      _cx             =   28178
      _cy             =   12726
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
      FormatString    =   $"Formr332.frx":002A
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5040
      Top             =   9960
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
      Left            =   4800
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4680
      Top             =   10080
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
      Left            =   4680
      Top             =   9840
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   12840
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   329777153
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12840
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   329777153
      CurrentDate     =   36892
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   25
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�ͻ�"
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   24
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
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
      Index           =   1
      Left            =   11640
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
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
      Index           =   1
      Left            =   11640
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����·�"
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
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Formr332"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim sdf, cbfy, rl, zj, rks, xss, xse, zzfy, qlz, ql, dgql As Double
Dim c, r As Integer
Private Sub Command1_Click()
'On Error Resume Next
If MsgBox("ȷ�������𣿣���ǰ�ļ�¼���������ȷ��", vbYesNo) = vbNo Then Exit Sub

Adodc2.RecordSource = "SELECT * FROM rqsd where �·�='" & Text1 & "'"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
DTPicker1.value = Adodc2.Recordset.Fields(0)
DTPicker2.value = Adodc2.Recordset.Fields(1)
Else
MsgBox ("�ڼ�������û�д��·���Ϣ")
Exit Sub
End If
sql1 = "delete from cbfxb"
sql2 = "insert into cbfxb(�ͻ�,����,�׺�,��ɫ,����,����ֵ,����,��������,�·�,Ʒ��) select '',��̨,����,'',����,��ֵ,round(����*cast(��ֵ as real)*0.005,6),round(����*cast(��ֵ as real)*0.005*����,6),'" & Text1 & "','' from pld where cast(CONVERT(varchar,����, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and ��Ϣ='����' and ���� not like '%-%' and len(����)=7 and ���� not in(select distinct �׺� from XSCBFXJZ)"
sql3 = "insert into cbfxb(�ͻ�,����,�׺�,��ɫ,����,����ֵ,����,��������,�·�,Ʒ��) select '',��̨, SUBSTRING(����, 1,PATINDEX('%-%', ����)-1) ,'',����,��ֵ,round(����*cast(��ֵ as real)*0.005,6),round(����*cast(��ֵ as real)*0.005*����,6),'" & Text1 & "','' from pld where cast(CONVERT(varchar,����, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and ��Ϣ='����' and ���� like '%-%' and len(����)>7 and SUBSTRING(����, 1,PATINDEX('%-%', ����)-1) not in(select distinct �׺� from XSCBFXJZ)"
sql4 = "update cbfxb set �ͻ�='1'"
sql5 = "insert into cbfxb(����,�׺�,��ɫ,����,����ֵ,�·�,Ʒ��) select ����,�׺�,��ɫ,����,sum(����ֵ),�·�,Ʒ�� from cbfxb group by ����,�׺�,��ɫ,����,�·�,Ʒ��"
sql6 = "delete from cbfxb where �ͻ�='1'"
sql7 = "update v_cbfxb_qlsx set ����ֵ=��ֵ"
sql8 = "update cbfxb set �ͻ�='',����=round(����*cast(����ֵ as real)*0.005,6),��������=round(����*cast(����ֵ as real)*0.005*����,6)"
sql9 = "delete from XSCBFX"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic
RD.Open sql7, conn, adOpenStatic, adLockOptimistic
RD.Open sql8, conn, adOpenStatic, adLockOptimistic
RD.Open sql9, conn, adOpenStatic, adLockOptimistic


Adodc3.RecordSource = "select �׺�,����,����ֵ,����,�������� from cbfxb where ����ֵ>0 order by �׺�"
Adodc3.Refresh


If Not Adodc3.Recordset.EOF Then
i = 1
ProgressBar1.Visible = True
sl = Adodc3.Recordset.RecordCount
ProgressBar1.value = i / sl * 100

Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
ProgressBar1.value = i / sl * 100
Adodc4.RecordSource = "select sum(isnull(��ֵ,0)) from pld where ����='" & Adodc3.Recordset.Fields(0) & "'"
Adodc4.Refresh
If Not IsNull(Adodc4.Recordset.Fields(0)) Then
qlz = Adodc4.Recordset.Fields(0)
ql = Val(Adodc3.Recordset.Fields(1)) * qlz * 0.005
dgql = Val(Adodc3.Recordset.Fields(1)) * qlz * 0.005 * Val(Adodc3.Recordset.Fields(1))
sql1 = "update cbfxb set ����ֵ='" & qlz & "',����='" & ql & "',��������='" & dgql & "' where �׺�='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Adodc3.Recordset.MoveNext
i = i + 1
Loop
ProgressBar1.Visible = False
End If

Adodc3.RecordSource = "select ����,sum(��������) from cbfxb group by ���� order by ����"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
sql1 = "update cbfxb set �����ܼ�='" & Adodc3.Recordset.Fields(1) & "' where ����='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.Recordset.MoveNext
Loop
End If

Adodc3.RecordSource = "select �ɱ�����,�ɱ����� from cbfy where �ɱ��ڼ�='" & Text1 & "' order by �ɱ�����"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
sql1 = "update cbfxb set ����='" & Adodc3.Recordset.Fields(1) & "' where ����='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.Recordset.MoveNext
Loop
End If

Adodc3.RecordSource = "select ��̨���,���� from ct  order by ��̨���"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
sql1 = "update cbfxb set ϵ��='" & Adodc3.Recordset.Fields(1) & "' where ����='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.Recordset.MoveNext
Loop
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�����õ����
sql1 = "delete from cbfxbdx"
sql2 = "insert into cbfxbdx(�ͻ�,�׺�,����) select '', SUBSTRING(����, 1, PATINDEX('%-%', ����)-1),round(sum(��β���),2) from ddcl where cast(CONVERT(varchar,ʱ��, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and ���ձ��='����' and ���� like '%-%' and len(����)>7 and SUBSTRING(����, 1, PATINDEX('%-%', ����)-1) not in(select distinct �׺� from XSCBFXJZ) group by  SUBSTRING(����, 1, PATINDEX('%-%', ����)-1)"
sql3 = "insert into cbfxbdx(�ͻ�,�׺�,����) select '', ����,round(sum(��β���),2) from ddcl where cast(CONVERT(varchar,ʱ��, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and ���ձ��='����' and ���� not like '%-%' and len(����)=7 and ���� not in(select distinct �׺� from XSCBFXJZ) group by  ����"
sql4 = "delete from cbfxbdx where len(�׺�)<>7"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic

Adodc3.RecordSource = "select sum(�ɱ�����) from cbfy where �ɱ��ڼ�='" & Text1 & "' and �ɱ����� like '%�����õ�%'"
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.EOF) Then
sql1 = "update cbfxbdx set ����='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Else
sql1 = "update cbfxbdx set ����=0"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''����
Adodc3.RecordSource = "select sum(�ɱ�����) from cbfy where �ɱ��ڼ�='" & Text1 & "' and �ɱ�����='��������'"
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.EOF) Then
sql1 = "update cbfxbdx set ����='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Else
sql1 = "update cbfxbdx set ����=0"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

sql1 = "update cbfxbdx set ����=0 where ���� is null"
sql2 = "update cbfxbdx set ����=0 where ���� is null"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�����ܲ���
Adodc3.RecordSource = "select * from cbfxbdx"
Adodc3.Refresh
zdxsl = 0
If Not Adodc3.Recordset.EOF Then
Adodc3.RecordSource = "select round(sum(isnull(����,0)),2) from cbfxbdx"
Adodc3.Refresh
zdxsl = Val(Adodc3.Recordset.Fields(0))
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���͸׷�̯
sql1 = "update cbfxbdx set ����=round(����/'" & zdxsl & "'*����,6) where cast('" & zdxsl & "' as real)<>0"
sql2 = "update cbfxbdx set ����=round(����/'" & zdxsl & "'*����,6) where cast('" & zdxsl & "' as real)<>0"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql2 = "update cbfxb set ����=round(��������/�����ܼ�*����,6),ˮ��=round(��������/�����ܼ�*����*ϵ��,6) where isnull(�����ܼ�,0)<>0"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic


Adodc3.RecordSource = "select sum(�ɱ�����) from cbfy where �ɱ�����='Ⱦɫ����' and �ɱ��ڼ�='" & Text1 & "'"           '''''''''''''''���ʷ���
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.Fields(0)) Then
cbqyl = Adodc3.Recordset.Fields(0)
Else
cbqyl = 0
End If

Adodc4.RecordSource = "select sum(��������) from cbfxb"           '''''''''''''''���ʷ���
Adodc4.Refresh
If Not IsNull(Adodc4.Recordset.Fields(0)) Then
qzyl = Adodc4.Recordset.Fields(0)
Else
qzyl = 0
End If
If qzyl = 0 Then
qxs = 0 '''''''''''''''''����ϵ��
Else
qxs = cbqyl / qzyl '''''''''''''''''����ϵ��
End If
Adodc3.RecordSource = "select �ɱ����� from cbfy where �ɱ�����='��' and �ɱ��ڼ�='" & Text1 & "'"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
sql1 = "update cbfxb set ����=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(��������,0)*'" & qxs & "'"
sql2 = "update cbfxbdx set ����=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(����,0)"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If

Adodc3.RecordSource = "select �ɱ����� from cbfy where �ɱ�����='ˮ' and �ɱ��ڼ�='" & Text1 & "'"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
sql1 = "update cbfxb set ˮ��=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(ˮ��,0)"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

Adodc3.RecordSource = "select �ɱ����� from cbfy where �ɱ�����='��' and �ɱ��ڼ�='" & Text1 & "'"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
sql1 = "update cbfxb set ���=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(����,0)"
sql2 = "update cbfxbdx set ����=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(����,0)"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''ת��ɱ����۱�
sql1 = "delete from XSCBFX where �·�='" & Text1 & "'"
sql2 = "insert into XSCBFX(�׺�,��ɫ,�·�,ˮ��,���,����,��Ϣ,�ͻ�,Ʒ��,����) select �׺�,��ɫ,�·�,round(isnull(ˮ��,0),2),round(isnull(���,0),2),round(isnull(����,0),2),'����',�ͻ�,Ʒ��,'��' from cbfxb where �·�='" & Text1 & "'"
sql3 = "insert into XSCBFX(�׺�,��ɫ,�·�,ˮ��,���,����,��Ϣ,�ͻ�,Ʒ��,����) select �׺�,'','" & Text1 & "',0,����,0,'����',�ͻ�,'','��' from cbfxbdx"
sql4 = "insert into XSCBFX(�׺�,��ɫ,�·�,ˮ��,���,����,��Ϣ,�ͻ�,Ʒ��,����) select �׺�,'','" & Text1 & "',0,0,����,'����',�ͻ�,'','��' from cbfxbdx"
sql5 = "update XSCBFX set �ͻ�='1'"
sql6 = "insert into XSCBFX(�׺�,��ɫ,�·�,ˮ��,���,����,��Ϣ,�ͻ�,Ʒ��,����) select �׺�,��ɫ,�·�,sum(ˮ��),sum(���),sum(����),��Ϣ,'',Ʒ��,���� from XSCBFX group by �׺�,��ɫ,�·�,��Ϣ,Ʒ��,����"
sql7 = "delete from XSCBFX where �ͻ�='1'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic
RD.Open sql7, conn, adOpenStatic, adLockOptimistic


Adodc3.RecordSource = "select distinct �׺� from XSCBFX where �·�='" & Text1 & "' order by �׺�"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
sl = Adodc3.Recordset.RecordCount
Adodc3.Recordset.MoveFirst
i = 1
Do While Not Adodc3.Recordset.EOF
ProgressBar1.Visible = True
ProgressBar1.value = i / sl * 100
Adodc4.RecordSource = "select round(sum(isnull(ë������,0)),2) from jgmxkf where ����='" & Adodc3.Recordset.Fields(0) & "'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
rks = 0
Else
rks = Val(Adodc4.Recordset.Fields(0))
End If
Adodc4.RecordSource = "select round(sum(isnull(����,0)),2),round(sum(isnull(���,0)),2) from jgmx where (����='" & Adodc3.Recordset.Fields(0) & "' or ���� like '" & Adodc3.Recordset.Fields(0) & "'+'-%') and ����<= cast('" & DTPicker2.value & "' as datetime) and �ӹ���� not in('������','ӡ����','��ӡ��')"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
xss = 0
xse = 0
Else
xss = Val(Adodc4.Recordset.Fields(0))
xse = Val(Adodc4.Recordset.Fields(1))
End If
Adodc4.RecordSource = "select round(sum(isnull(�ϼƽ��,0)),2) from v_pld_tj_xx_hs where (����='" & Adodc3.Recordset.Fields(0) & "' or ���� like '" & Adodc3.Recordset.Fields(0) & "'+'-%') and ����='Ⱦ�Ͽ�'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
rl = 0
Else
rl = Val(Adodc4.Recordset.Fields(0))
End If
Adodc4.RecordSource = "select round(sum(isnull(�ϼƽ��,0)),2) from v_pld_tj_xx_hs where (����='" & Adodc3.Recordset.Fields(0) & "' or ���� like '" & Adodc3.Recordset.Fields(0) & "'+'-%') and ����='������'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
zj = 0
Else
zj = Val(Adodc4.Recordset.Fields(0))
End If
'Adodc4.RecordSource = "select sum(����) from cbfxbdx where �׺�='" & Adodc3.Recordset.Fields(0) & "'"
'Adodc4.Refresh
'If Not IsNull(Adodc4.Recordset.Fields(0)) Then
'df = Val(Adodc4.Recordset.Fields(0))
'Else
'df = 0
'End If
sql1 = "update XSCBFX set �����='" & rks & "',������='" & xss & "',���۶�='" & xse & "',Ⱦ��='" & rl & "',����='" & zj & "' where �·�='" & Text1 & "' and �׺�='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.Recordset.MoveNext
i = i + 1
Loop
End If


''''''''''''''''''''''''''''''''''''''''''''''''''''����ת��
If Len(Text1) = 4 Then         ''''�ж��ڼ��Ƿ���ȷ
l1 = Mid(Text1, 1, 2)         ''''���
L2 = Mid(Text1, 3, 2)         ''''�·�
If Val(L2) = 12 Then          ''��������һ���·�
l1 = Val(l1) - 1              ''
L3 = Trim(l1) + "12"
Else                         ''��������·�
L2 = Val(L2) - 1
If Len(Trim(L2)) = 1 Then    ''����2λ  ��0
L3 = l1 + "0" + Trim(L2)
Else
L3 = l1 + Trim(L2)
End If
End If
End If


sql1 = "INSERT into XSCBFX SELECT * FROM XSCBFXQM where �·�='" & L3 & "' and len(�׺�)=7"
sql2 = "update XSCBFX set �·�='" & Text1 & "',����='��' where �·�='" & L3 & "' and ��Ϣ='��ת'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc4.RecordSource = "select distinct �׺� from XSCBFX where  �·�='" & Text1 & "' and ��Ϣ='��ת'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
i = 1
sl = Adodc4.Recordset.RecordCount
ProgressBar1.value = i / sl * 100
Adodc4.Recordset.MoveFirst
Do While Not Adodc4.Recordset.EOF
sql1 = "update XSCBFX set ������=(select sum(����) from jgmx where ���� like '" & Adodc4.Recordset.Fields(0) & "'+'%' and ����<=cast('" & DTPicker2.value & "' as datetime) and �ӹ���� not in('������','ӡ����','��ӡ��')) where �׺�='" & Adodc4.Recordset.Fields(0) & "' and ��Ϣ='��ת'"
sql2 = "update XSCBFX set ���۶�=(select sum(���) from jgmx where ���� like '" & Adodc4.Recordset.Fields(0) & "'+'%' and ����<=cast('" & DTPicker2.value & "' as datetime) and �ӹ���� not in('������','ӡ����','��ӡ��')) where �׺�='" & Adodc4.Recordset.Fields(0) & "' and ��Ϣ='��ת'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc4.Recordset.MoveNext
i = i + 1
Loop
End If

sql3 = "update XSCBFX set ��Ϣ='1' where �·�='" & Text1 & "'"
sql4 = "insert into XSCBFX(�·�,�׺�,Ⱦ��,����,ˮ��,���,����,����,�������,�����,������,���۶�) select �·�,�׺�,sum(isnull(Ⱦ��,0)),sum(isnull(����,0)),sum(isnull(ˮ��,0)),sum(isnull(���,0)),sum(isnull(����,0)),sum(isnull(����,0)),sum(isnull(�������,0)),sum(isnull(�����,0)),sum(isnull(������,0)),sum(isnull(���۶�,0)) from XSCBFX where �·�='" & Text1 & "' group by �·�,�׺�"
sql5 = "delete from XSCBFX where ��Ϣ='1' and �·�='" & Text1 & "'"
sql6 = "update  XSCBFX set ��Ϣ='����',����='��' where �·�='" & Text1 & "'"
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic

sql1 = "update XSCBFX set ����='��' where �·�='" & Text1 & "' and (������+5)>=����� and �����>0"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

Adodc4.RecordSource = "select round(sum(isnull(���۶�,0)),2) from XSCBFX where �·�='" & Text1 & "' and ����='��'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
zse = 0
Else
zse = Val(Adodc4.Recordset.Fields(0))
End If

Adodc4.RecordSource = "select �ɱ����� from cbfy where �ɱ�����='�������' and �ɱ��ڼ�='" & Text1 & "'"         '''''''''''''''�������
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
cbfy = Adodc4.Recordset.Fields(0)
Else
cbfy = 0
End If

Adodc3.RecordSource = "select �ɱ����� from cbfy where �ɱ�����='����' and �ɱ��ڼ�='" & Text1 & "'"           '''''''''''''''���ʷ���
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
gzfy = Adodc3.Recordset.Fields(0)
Else
gzfy = 0
End If

Adodc3.RecordSource = "select sum(isnull(ˮ��,0)+isnull(���,0)+isnull(����,0)) from XSCBFX where  �·�='" & Text1 & "' and ����='��'"          ''''���ʷ���
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.Fields(0)) Then
sdf = Adodc3.Recordset.Fields(0)
gzxs = gzfy / sdf
sql4 = "update XSCBFX set ����=round('" & gzxs & "'*(isnull(ˮ��,0) + isnull(���,0) + isnull(����,0)),2),�������=round(���۶�/'" & zse & "'*'" & cbfy & "',2) where �·�='" & Text1 & "' and cast('" & sdf & "' as real)<>0 and cast('" & zse & "' as real)<>0 and ����='��'"
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
End If

sql2 = "update XSCBFX set ���۳ɱ�=round(Ⱦ��+����+�������+ˮ��+���+����+����,2),���ۼ�=round(���۶�/������,2) where �·�='" & Text1 & "' and ����='��' and ������<>0 and �����>0"
sql3 = "update XSCBFX set ë����=round((���۶�-���۳ɱ�)/���۶�*100,1) where �·�='" & Text1 & "' and ����='��' and ���۳ɱ�<>0 and ���۶�<>0 and �����>0"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic

ProgressBar1.Visible = False
Adodc3.RecordSource = "select �׺� from XSCBFX where �·�='" & Text1 & "' order by �׺�"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
i = 1
sl = Adodc3.Recordset.RecordCount
ProgressBar1.Visible = True
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
ProgressBar1.value = i / sl * 100
Adodc4.RecordSource = "select �ͻ�����,Ʒ��,ɫ��+ɫ�� from kpd where ����='" & Adodc3.Recordset.Fields(0) & "' order by ���� desc"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
sql2 = "update XSCBFX set �ͻ�='" & Adodc4.Recordset.Fields(0) & "',Ʒ��='" & Adodc4.Recordset.Fields(1) & "',��ɫ='" & Adodc4.Recordset.Fields(2) & "' where �׺�='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
Adodc3.Recordset.MoveNext
i = i + 1
Loop
ProgressBar1.Visible = False
End If
Adodc1.RecordSource = "select * from XSCBFX where �·�='" & Text1 & "' order by �׺�"
Adodc1.Refresh

End Sub

Private Sub Command3_Click()

If MsgBox("��ȷ�Ͻ�ת�ڼ䣺" + Text1 + "��ȷ��?", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete from XSCBFXQM where �·�='" & L3 & "'"
sql2 = "insert into XSCBFXQM(�׺�,��ɫ,Ⱦ��,����,ˮ��,���,����,����,�������,�����,�ͻ�,������,���۶�,���ۼ�,���۳ɱ�,ë����,�·�,��Ϣ,Ʒ��) select �׺�,��ɫ,Ⱦ��,����,ˮ��,���,����,����,�������,�����,�ͻ�,������,���۶�,���ۼ�,���۳ɱ�,ë����,'" & Text1 & "','��ת',Ʒ�� from XSCBFX where ����='��' and �·�='" & Text1 & "'"
sql3 = "delete from XSCBFXJZ where �·�='" & Text1 & "'"
sql4 = "insert into XSCBFXJZ(�׺�,��ɫ,Ⱦ��,����,ˮ��,���,����,����,�������,�����,�ͻ�,������,���۶�,���ۼ�,���۳ɱ�,ë����,�·�,��Ϣ,Ʒ��) select �׺�,��ɫ,Ⱦ��,����,ˮ��,���,����,����,�������,�����,�ͻ�,������,���۶�,���ۼ�,���۳ɱ�,ë����,'" & Text1 & "','��ת',Ʒ�� from XSCBFX where ����='��' and �·�='" & Text1 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
MsgBox ("��ת�ɹ�!")
End Sub


Private Sub Command4_Click()
If Option1(2).value = True Then
Adodc1.RecordSource = "select * from XSCBFX where �·�='" & Text1 & "' order by �׺�"
Adodc1.Refresh
End If
If Option1(0).value = True Then
Adodc1.RecordSource = "select * from XSCBFX where �·�='" & Text1 & "' and  ����='��' order by �׺�"
Adodc1.Refresh
End If
If Option1(1).value = True Then
Adodc1.RecordSource = "select * from XSCBFX where �·�='" & Text1 & "' and  ����='��' order by �׺�"
Adodc1.Refresh
End If
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 0, 5, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 6, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 7, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 8, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 9, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 10, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 11, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 12, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 13, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 14, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 16, , vbGreen
End Sub

Private Sub Command5_Click()
If MsgBox("ȷ�����º���ɱ���", vbYesNo) = vbNo Then Exit Sub
Adodc4.RecordSource = "select round(sum(isnull(���۶�,0)),2) from XSCBFX where �·�='" & Text1 & "' and ����='��'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
zse = 0
Else
zse = Val(Adodc4.Recordset.Fields(0))
End If

Adodc4.RecordSource = "select �ɱ����� from cbfy where �ɱ�����='�������' and �ɱ��ڼ�='" & Text1 & "'"         '''''''''''''''�������
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
cbfy = Adodc4.Recordset.Fields(0)
Else
cbfy = 0
End If

Adodc3.RecordSource = "select �ɱ����� from cbfy where �ɱ�����='����' and �ɱ��ڼ�='" & Text1 & "'"           '''''''''''''''���ʷ���
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
gzfy = Adodc3.Recordset.Fields(0)
Else
gzfy = 0
End If

Adodc3.RecordSource = "select sum(isnull(ˮ��,0)+isnull(���,0)+isnull(����,0)) from XSCBFX where  �·�='" & Text1 & "' and ����='��'"          ''''���ʷ���
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.Fields(0)) Then
sdf = Adodc3.Recordset.Fields(0)
gzxs = gzfy / sdf
sql4 = "update XSCBFX set ����=round('" & gzxs & "'*(isnull(ˮ��,0) + isnull(���,0) + isnull(����,0)),2),�������=round(���۶�/'" & zse & "'*'" & cbfy & "',2) where �·�='" & Text1 & "' and cast('" & zse & "' as real)<>0 and ����='��'"
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
End If

sql2 = "update XSCBFX set ���۳ɱ�=round(Ⱦ��+����+�������+ˮ��+���+����+����,2),���ۼ�=round(���۶�/������,2) where �·�='" & Text1 & "' and ����='��' and ������<>0"
sql3 = "update XSCBFX set ë����=round((���۶�-���۳ɱ�)/���۶�*100,1) where �·�='" & Text1 & "' and  ���۶�<>0 and ����='��'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
MsgBox ("���·����ɹ���")

End Sub

Private Sub Command6_Click()
Call MXCBFX(VSFlexGrid1, "�ɱ�����")
End Sub

Private Sub Command7_Click()
sql1 = ""
If Check2(1).value = 1 Then
sql1 = sql1 + "�ͻ� like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "�׺� like '%'+'" & Text3 & "'+'%' and "
End If

If sql1 = "" Then
MsgBox ("��ѡ���ѯ����")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "select * from XSCBFX where (" + sql1 + ") and �·�='" & Text1 & "'"
Adodc1.Refresh

End Sub

Private Sub Text1_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from XSCBFX where �·�='" & Text1 & "' order by �׺�"
Adodc1.Refresh
End Sub

Private Sub Form_Load()

On Error Resume Next
Text1 = ""
Text2 = ""
Text3 = ""
DataCombo1 = ""
Option1(2).value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from XSCBFX where �·�='" & Text1 & "'"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(1) = 1500
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text2_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select ��� from KHZL where ����  like '%'+'" & Text2 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by ���"
Adodc5.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
End With
If r = 2 Then
Formc23.DataCombo4 = VSFlexGrid1.TextMatrix(r, 2)
Formc23.Check2(7).value = 1
Formc23.Show
End If

If c = 5 Or c = 6 Then
Formr309.Text1 = VSFlexGrid1.TextMatrix(r, 2)
Formr309.Show
End If

If c = 8 Then
Formr307.Text1 = VSFlexGrid1.TextMatrix(r, 2)
Formr307.Show
End If

If c = 9 Then
Formr308.Text1 = VSFlexGrid1.TextMatrix(r, 2)
Formr308.Show
End If

End Sub

Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
End With

S2 = VSFlexGrid1.TextMatrix(r, 2)   '''�׺�

    If Button = 2 And c = 2 Then
    If MsgBox("ȷ���������е���Ϣ��" + S2, vbYesNo) = vbNo Then  '''PopupMenu mnu_manager  '�����ڴ��������õ�һ�������˵�����
    Exit Sub
    Else
    sql2 = "update XSCBFX set ����='��' where �׺�='" & S2 & "' and �·�='" & Text1 & "'"
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic
    End If
    Call Command4_Click
    End If

    If Button = 2 And c = 1 Then
    If MsgBox("ȷ��ȡ���������е���Ϣ��" + S2, vbYesNo) = vbNo Then  '''PopupMenu mnu_manager  '�����ڴ��������õ�һ�������˵�����
    Exit Sub
    Else
    sql2 = "update XSCBFX set ����='��' where �׺�='" & S2 & "' and �·�='" & Text1 & "'"
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic
    End If
    Call Command4_Click
    End If

End Sub


Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
    If c = 12 Or c = 13 Or c = 14 Then
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

