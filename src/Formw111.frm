VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formw111 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�˻��ڳ��������"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9000
      Top             =   10440
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
      Left            =   9360
      Top             =   10320
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
      Bindings        =   "Formw111.frx":0000
      Height          =   7095
      Left            =   480
      TabIndex        =   34
      Top             =   3000
      Width           =   14295
      _cx             =   25215
      _cy             =   12515
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
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����׼��"
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   10920
      TabIndex        =   32
      Text            =   "Text2"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   9
      Left            =   8880
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   8
      Left            =   8880
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formw111.frx":0015
      Left            =   2280
      List            =   "Formw111.frx":0022
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�޸�"
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
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ɾ��"
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
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "¼��"
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
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ѯ"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   7
      Left            =   6840
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   6
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   5
      Left            =   4800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   4
      Left            =   4800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   3
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   329711617
      CurrentDate     =   39903
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ѡȡ"
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
      Left            =   3960
      TabIndex        =   29
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Index           =   10
      Left            =   8880
      TabIndex        =   28
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Index           =   9
      Left            =   8880
      TabIndex        =   27
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "�˻�����"
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
      Left            =   480
      TabIndex        =   24
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Index           =   7
      Left            =   6840
      TabIndex        =   15
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Index           =   6
      Left            =   6840
      TabIndex        =   14
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����������"
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
      Left            =   4800
      TabIndex        =   13
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ƿ�Ŀ"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "ժҪ"
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
      Left            =   2640
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "ƾ֤��"
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
      Left            =   480
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "�跽�������"
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
      Left            =   4800
      TabIndex        =   9
      Top             =   840
      Width           =   1815
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
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Formw111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public r, c As Integer
Private Sub Combo1_Click()
On Error Resume Next
If Combo1.Text = "" Then
MsgBox ("��ѡ���˻�����")
Exit Sub
End If

If Combo1.Text = "�����˻�" Then
Adodc1.RecordSource = "SELECT * FROM PMZJZ WHERE ����='" & Text1(0).Text & "' ORDER BY ���  DESC"
Adodc1.Refresh
End If

If Combo1.Text = "��ϸ�˻�" Then
If Text1(3).Text = "" Then
Adodc1.RecordSource = "SELECT * FROM PMMXJZ WHERE ����='" & Text1(0).Text & "'  ORDER BY ���  DESC"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM PMMXJZ WHERE ����='" & Text1(0).Text & "'  ORDER BY ���  DESC"
Adodc1.Refresh
End If
End If

If Combo1.Text = "�����˻�" Then
Adodc1.RecordSource = "SELECT * FROM PMTZ WHERE ����='" & Text1(0).Text & "' ORDER BY ���  DESC"
Adodc1.Refresh
End If
Text1(1).Text = "��-" + Trim(Month(Text1(0).Text) - 1)
Text1(2).Text = "�ڳ����"
Text1(9).Text = Adodc1.Recordset.RecordCount + 1
End Sub


Private Sub Command12_Click()
If Combo1.Text = "�����˻�" Then
Adodc1.RecordSource = "SELECT * FROM PMZJZ WHERE ����='" & Text1(0).Text & "' AND ��ƿ�Ŀ like '%'+'" & Text1(3).Text & "'+'%' ORDER BY ���  DESC"
Adodc1.Refresh
End If

If Combo1.Text = "��ϸ�˻�" Then
Adodc1.RecordSource = "SELECT * FROM PMMXJZ WHERE ����='" & Text1(0).Text & "' AND ��ƿ�Ŀ like '%'+'" & Text1(3).Text & "'+'%' ORDER BY ���  DESC"
Adodc1.Refresh
End If

If Combo1.Text = "�����˻�" Then
Adodc1.RecordSource = "SELECT * FROM PMTZ WHERE ����='" & Text1(0).Text & "' AND ��ƿ�Ŀ like '%'+'" & Text1(3).Text & "'+'%' ORDER BY ���  DESC"
Adodc1.Refresh
End If
Text1(9).Text = Adodc1.Recordset.RecordCount + 1

End Sub

Private Sub Command2_Click()
Call OutadodcToExcel(VSFlexGrid1, 8, Combo1.Text + "�ڳ����")
End Sub


Private Sub Command6_Click()
Unload Me
End Sub


Private Sub Command8_Click()
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1500
VSFlexGrid1.ColWidth(3) = 1500
VSFlexGrid1.ColWidth(4) = 2500
VSFlexGrid1.ColWidth(5) = 0
VSFlexGrid1.ColWidth(6) = 0
VSFlexGrid1.ColWidth(7) = 0
VSFlexGrid1.ColWidth(8) = 1500
VSFlexGrid1.ColWidth(9) = 0
VSFlexGrid1.ColWidth(10) = 0
End Sub

Private Sub DTPicker1_Change()
Text1(0).Text = DTPicker1.value
End Sub

Private Sub DTPicker1_CloseUp()
Text1(0).Text = DTPicker1.value
End Sub

Private Sub Form_Load()
For i = 0 To 9
Text1(i).Text = ""
Next
Text1(0).Text = Date
DTPicker1.value = Date
Combo1.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1500
VSFlexGrid1.ColWidth(3) = 1500
VSFlexGrid1.ColWidth(4) = 2500
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

End Sub

Private Sub Label1_DblClick()
KMMC = 1
Formw6.Show
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 1 To 9
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
End Sub

Private Sub Command3_Click()
On Error Resume Next

For i = 0 To 9
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh

For i = 3 To 9
Text1(i).Text = ""
Next
Text1(9).Text = Adodc1.Recordset.RecordCount + 1
End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.Delete
Adodc1.Refresh
For i = 3 To 9
Text1(i).Text = ""
Next
Text1(9).Text = Adodc1.Recordset.RecordCount + 1
End Sub

Private Sub Command1_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
For i = 0 To 9
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
For i = 3 To 9
Text1(i).Text = ""
Next
Text1(9).Text = Adodc1.Recordset.RecordCount + 1
End Sub
'''''''''''''''''''''''
Private Sub MSFlex_DBLClick()
With VSFlexGrid1
    c = .col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex_DBLClick
End If
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 3
If Combo1.Text = "��ϸ�˻�" Then
Adodc1.RecordSource = "SELECT * FROM PMMXJZ WHERE ��ƿ�Ŀ like '%'+'" & Text1(3).Text & "'+'%' AND ����='" & Text1(0).Text & "' ORDER BY ���  DESC"
Adodc1.Refresh
End If
End Select
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid1.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    VSFlexGrid1.Text = Text1111.Text
    Text1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move r - 1
Adodc1.Recordset.Fields(c - 1) = Text1111.Text
Adodc1.Recordset.Update
Text1111.Visible = False
End Sub



