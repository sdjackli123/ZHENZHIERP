VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw211 
   BackColor       =   &H00C0E0FF&
   Caption         =   "���������ϸ"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   480
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ѯ����"
      Height          =   975
      Left            =   9600
      TabIndex        =   14
      Top             =   480
      Width           =   2535
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "��Ӧ"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   6480
      Style           =   1  'Simple Combo
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw211.frx":0000
      Height          =   8175
      Left            =   600
      TabIndex        =   10
      Top             =   1800
      Width           =   17535
      _cx             =   30930
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw211.frx":0015
      Height          =   330
      Left            =   5520
      TabIndex        =   9
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command3 
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
      Height          =   375
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   328925185
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   328925185
      CurrentDate     =   36892
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8160
      Top             =   10320
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
      Left            =   8040
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   7680
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
      Left            =   7680
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   7920
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
      Left            =   7440
      Top             =   10200
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
      Left            =   7920
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   375
      Left            =   7560
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formw211.frx":002A
      Height          =   330
      Left            =   5040
      TabIndex        =   11
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formw211.frx":003F
      Height          =   330
      Left            =   7680
      TabIndex        =   21
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
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
      Index           =   0
      Left            =   4440
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ѡ�����ڷ�Χ"
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
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "��Ӧ"
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
      Left            =   4440
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Formw211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Public c, r As Integer
Private Sub Command3_Click()

sql1 = ""
t1 = Format(DTPicker3.value, "yyyy-mm-dd")
t2 = Format(DTPicker4.value, "yyyy-mm-dd")

If Check2(0).value = 1 Then
sql1 = sql1 + "��� like '%'+'" & DataCombo3 & "'+'%' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "��Ӧ��λ like '%'+'" & DataCombo1 & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "���� like '%'+'" & DataCombo2 & "'+'%' and "
End If

If sql1 = "" Then
MsgBox ("��ѡ���ѯ����")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc3.RecordSource = "SELECT ��Ӧ��λ,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,�ϼƽ��,����,�Ƿ�Ʊ,�Ƿ񸶿�,���ݺ�,���� FROM CLGL WHERE (" + sql1 + ")  AND ���='�ɹ����' ORDER BY ����,���ݺ�"
Adodc3.Refresh

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Call OutadodcToExcel(VSFlexGrid1, 9, "  �ͻ���" + DataCombo1.Text + "     ���������ϸ")
End Sub


Private Sub Command5_Click()

End Sub

Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ��� from GYS where ip like '%'+'" & yhxx & "'+'%' GROUP BY ���"
Adodc1.Refresh

DTPicker3.value = Date
DTPicker4.value = Date
cdbhf = cdbh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select MC from CLKB   group by MC"
Adodc2.Refresh

Check2(4).value = 1
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc3.CommandTimeout = 10000
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT ��Ӧ��λ,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,�ϼƽ��,����,�Ƿ�Ʊ,�Ƿ񸶿�,���ݺ�,���� FROM CLGL WHERE ���� BETWEEN '" & DTPicker3.value & "' AND '" & DTPicker4.value & "' order by ����,���ݺ�"
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select MC from CLKL   group by MC"
Adodc4.Refresh

VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1500

DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
Text1 = ""
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

Private Sub Label2_Click()
DataCombo1.Text = ""
End Sub


Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
    If c = 8 Then
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


Private Sub Text1_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ��� from GYS where ip like '%'+'" & yhxx & "'+'%' and ���� like '%'+'" & Text1 & "'+'%' GROUP BY ���"
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
Adodc3.Recordset.MoveFirst
Adodc3.Recordset.Move r - 1
Adodc3.Recordset.Fields(c - 1) = Combo1111.Text
Adodc3.Recordset.Fields(c) = Format(Val(Adodc3.Recordset.Fields(c - 2)) * Val(Combo1111.Text), "#0.00")
L = Format(Val(Adodc3.Recordset.Fields(c - 2)) * Val(Combo1111.Text), "#0.00")
Adodc3.Recordset.Update
VSFlexGrid1.Text = Combo1111.Text
VSFlexGrid1.TextMatrix(r, c + 1) = L
Combo1111.Visible = False
VSFlexGrid1.SetFocus
End If
End Sub


