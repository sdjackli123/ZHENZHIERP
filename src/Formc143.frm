VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formc143 
   BackColor       =   &H00C0E0FF&
   Caption         =   "������Ϣ"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   14685
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
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
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�뵥��ӡ"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȫ��"
      Height          =   495
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȫѡ"
      Height          =   495
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      ItemData        =   "Formc143.frx":0000
      Left            =   10080
      List            =   "Formc143.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      Height          =   855
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ѡȡȷ��"
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   1
      Left            =   5280
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   2
      Left            =   8640
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ѯ��Ϣ"
      Height          =   1095
      Left            =   720
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   4815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�ѷ�"
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "δ��"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȡ��ȷ��"
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ѡ�񷢻�"
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10920
      Top             =   240
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formc143.frx":0004
      Height          =   2655
      Left            =   720
      TabIndex        =   15
      Top             =   1320
      Width           =   13215
      _cx             =   23310
      _cy             =   4683
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   10920
      Top             =   240
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   11040
      Top             =   120
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
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   20
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   19
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����ƥ��"
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
      Index           =   2
      Left            =   4200
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Height          =   495
      Index           =   3
      Left            =   7560
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Formc143"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public gygh As String

Private Sub Command1_Click()
On Error Resume Next
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
ph = Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1)
gpps = gpps + 1
gpzl = gpzl + Val(Mid(List2.List(i), InStr(List2.List(i), "-") + 1))
sql1 = "update bmd set ����='' where ����='" & Text1.Text & "' and ƥ��='" & ph & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
MsgBox ("ȡ���ɹ���")
End Sub

Private Sub Command11_Click()
Call fhdmd(Adodc3, Text1, Text2(0))
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If fhxz = 15 Then
Formc15.DataCombo1.Text = Adodc1.Recordset.Fields(0)   ''''�ͻ�
Formc15.DataCombo2.Text = Adodc1.Recordset.Fields(4)   ''Ʒ��
Formc15.DataCombo3.Text = Adodc1.Recordset.Fields(6)   '��ɫ
Formc15.DataCombo4.Text = Adodc1.Recordset.Fields(3)   '����
Formc15.DataCombo5.Text = Adodc1.Recordset.Fields(15)   'ë������
Formc15.DataCombo11.Text = Adodc1.Recordset.Fields(2)  '���
Formc15.Text7.Text = Text2(1)      ' ë��ƥ��
Formc15.DataCombo16.Text = Adodc1.Recordset.Fields(1) '����
Formc15.Text11(0).Text = Adodc1.Recordset.Fields(13)  ''����
Formc15.Text11(1).Text = Adodc1.Recordset.Fields(14)  ''���
Formc15.Text11(2).Text = Adodc1.Recordset.Fields(5)  ''�ɷ�
Formc15.Text11(3).Text = Adodc1.Recordset.Fields(9)  ''��λ
Formc15.Text8.Text = Adodc1.Recordset.Fields(12)   ''''����
Formc15.Text9.Text = Adodc1.Recordset.Fields(7)   ''''ɫ��
Formc15.Text10.Text = Text2(2)   ''''��������
Unload Me
End If
End Sub

Private Sub Command4_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = False
Next
End Sub

Private Sub Command5_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = True
Next
End Sub

Private Sub Command6_Click()
If Option1(1).Value = True Then
Adodc2.RecordSource = "select ƥ��,�������� from bmd where ����='" & Text1 & "' and len(isnull(����,''))=0 order by ƥ��"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List2.Clear
Do While Not Adodc2.Recordset.EOF
List2.AddItem Trim(Adodc2.Recordset.Fields(0)) + "-" + Trim(Adodc2.Recordset.Fields(1))
Adodc2.Recordset.MoveNext
Loop
End If
If Option1(0).Value = True Then
Adodc2.RecordSource = "select ƥ��,�������� from bmd where ����='" & Text1 & "' and isnull(����,'')='" & Text2(0) & "' order by ƥ��"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List2.Clear
gpps = 0
gpzl = 0
Do While Not Adodc2.Recordset.EOF
List2.AddItem Trim(Adodc2.Recordset.Fields(0)) + "-" + Trim(Adodc2.Recordset.Fields(1))
gpps = gpps + 1
gpzl = gpzl + Val(Adodc2.Recordset.Fields(1))
Adodc2.Recordset.MoveNext
Loop
Text2(1) = gpps
Text2(2) = gpzl
End If
End Sub

Private Sub Command7_Click()
gpps = 0
gpzl = 0
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
ph = Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1)
gpps = gpps + 1
gpzl = gpzl + Val(Mid(List2.List(i), InStr(List2.List(i), "-") + 1))
sql1 = "update bmd set ����='" & Text2(0) & "' where ����='" & Text1.Text & "' and ƥ��='" & ph & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
Text2(1) = gpps
Text2(2) = gpzl
MsgBox ("ȷ�ϳɹ���")
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=fydn"
Set RD = New ADODB.Recordset
Option1(1).Value = True
For i = 1 To 2
Text2(i) = ""
Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=fydn"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=fydn"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=fydn"
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(11) = 1500
VSFlexGrid1.ColWidth(12) = 0
VSFlexGrid1.ColWidth(13) = 0
End Sub

Private Sub VSFlexGrid1_DblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
If fhxz = 15 Then
Formc15.DataCombo1.Text = Adodc1.Recordset.Fields(0)   ''''�ͻ�
Formc15.DataCombo2.Text = Adodc1.Recordset.Fields(4)   ''Ʒ��
Formc15.DataCombo3.Text = Adodc1.Recordset.Fields(6)   '��ɫ
Formc15.DataCombo4.Text = Adodc1.Recordset.Fields(3)   '����
Formc15.DataCombo5.Text = Adodc1.Recordset.Fields(15)   'ë������
Formc15.DataCombo11.Text = Adodc1.Recordset.Fields(2)  '���
Formc15.Text7.Text = Adodc1.Recordset.Fields(10)      ' ë��ƥ��
Formc15.DataCombo16.Text = Adodc1.Recordset.Fields(1) '����
Formc15.Text11(0).Text = Adodc1.Recordset.Fields(13)  ''����
Formc15.Text11(1).Text = Adodc1.Recordset.Fields(14)  ''���
Formc15.Text11(2).Text = Adodc1.Recordset.Fields(5)  ''�ɷ�
Formc15.Text11(3).Text = Adodc1.Recordset.Fields(9)  ''��λ
Formc15.Text8.Text = Adodc1.Recordset.Fields(12)   ''''����
Formc15.Text9.Text = Adodc1.Recordset.Fields(7)   ''''ɫ��
Formc15.Text10.Text = Adodc1.Recordset.Fields(8)   ''''��������
Unload Me
End If

If fhxz = 39 Then
Formc39.DataCombo1.Text = Adodc1.Recordset.Fields(0)   ''''�ͻ�
Formc39.DataCombo2.Text = Adodc1.Recordset.Fields(4)   ''Ʒ��
Formc39.DataCombo3.Text = Adodc1.Recordset.Fields(6)   '��ɫ
Formc39.DataCombo4.Text = Adodc1.Recordset.Fields(3)   '����
Formc39.DataCombo5.Text = Adodc1.Recordset.Fields(15)   'ë������
Formc39.DataCombo11.Text = Adodc1.Recordset.Fields(2)  '���
Formc39.Text7.Text = Adodc1.Recordset.Fields(10)      ' ë��ƥ��
Formc39.DataCombo16.Text = Adodc1.Recordset.Fields(1) '����
Formc39.Text11(0).Text = Adodc1.Recordset.Fields(13)  ''����
Formc39.Text11(1).Text = Adodc1.Recordset.Fields(14)  ''���
Formc39.Text11(2).Text = Adodc1.Recordset.Fields(5)  ''�ɷ�
Formc39.Text11(3).Text = Adodc1.Recordset.Fields(9)  ''��λ
Formc39.Text8.Text = Adodc1.Recordset.Fields(12)   ''''����
Formc39.Text9.Text = Adodc1.Recordset.Fields(7)   ''''ɫ��
Formc39.Text10.Text = Adodc1.Recordset.Fields(8)   ''''��������
Unload Me
End If
End Sub

Private Sub Text1_Change()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=fydn"
Adodc1.RecordSource = "select �ͻ�,����,���,����,Ʒ��,�ɷ�,��ɫ,ɫ��,��������,��λ,����ƥ��,����,����,����,���,ë������ from v_jgmx_fh  WHERE ����='" & Text1.Text & "'"
Adodc1.Refresh
End Sub


