VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy196 
   BackColor       =   &H00C0E0FF&
   Caption         =   "��Э�ӹ���ϸ"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy196.frx":0000
      Height          =   7575
      Left            =   1200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ɲ�ѯ"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ƾ֤����"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��λ��ѯ"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���Ų�ѯ"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��λ���"
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Data Data4 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data5 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data6 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy196.frx":0014
      Height          =   330
      Left            =   3840
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "���"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22806529
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22806529
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   11040
      TabIndex        =   16
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   22806529
      CurrentDate     =   36892
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   0
      Left            =   11040
      TabIndex        =   17
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��λ"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   1080
      Width           =   495
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
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   12
      Top             =   600
      Width           =   495
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
      Index           =   1
      Left            =   1200
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Formy196"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call WXCX(MSFlexGrid1, "��Э��ѯ")
End Sub

Private Sub Command2_Click()
Data2.RecordSource = "select * from wxjl where ����='" & Text2.Text & "' and VAL(����)=0 order by ���,��ɫ,���,λ��"
Data2.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data2.RecordSource = "select ��λ,���,��ɫ,���,λ��,���,����,����,format(val(����)*val(����),'#0.00') as ���,���� from wxjl where ����='" & Text2.Text & "' order by ���,��ɫ,���,λ��"
Data2.Refresh
End Sub

Private Sub Command5_Click()
If DBCombo1.Text = "" Then
Data2.RecordSource = "select ��λ,���,��ɫ,���,λ��,���,����,����,format(val(����)*val(����),'#0.00') as ���,���� from wxjl WHERE ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') order by ����,���,��ɫ,���,λ��"
Data2.Refresh
Else
Data2.RecordSource = "select ��λ,���,��ɫ,���,λ��,���,����,����,format(val(����)*val(����),'#0.00') as ���,���� from wxjl where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and  ��λ='" & DBCombo1.Text & "'  order by ����,���,��ɫ,���,λ��"
Data2.Refresh
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
If DBCombo1.Text = "" Then
Data2.RecordSource = "select * from wxjl WHERE ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') AND VAL(����)=0 order by ����,���,��ɫ,���,λ��"
Data2.Refresh
Else
Data2.RecordSource = "select * from wxjl where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') AND VAL(����)=0 and  ��λ='" & DBCombo1.Text & "'  order by ����,���,��ɫ,���,λ��"
Data2.Refresh
End If

End Sub


Private Sub Command7_Click()
If MsgBox("��������Ϊ��" + Trim(DTPicker1.Value) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("�����ڼ�Ϊ��" + Trim(Month(DTPicker1.Value)) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("ȷ�����ɼӹ�ϵ�е�ƾ֤��", vbYesNo) = vbNo Then Exit Sub
Call WXJGPZ(CDate(DTPicker3.Value), CDate(DTPicker4.Value), CDate(DTPicker1.Value))
End Sub

Private Sub Command8_Click()
Formw332.Combo1.Text = "ת��ƾ֤"
Formw332.Show
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Text2.Text = ""
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\CW.mdb"

Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select ��� from gys where instr(����,'��')>0 group by ���"
Data3.Refresh

Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data5.DatabaseName = "d:\���ݿ�\\htgl\2011\CW.mdb"


MSFlexGrid1.ColWidth(0) = 300
For i = 1 To 3
MSFlexGrid1.ColWidth(i) = 1200
Next

For i = 4 To 5
MSFlexGrid1.ColWidth(i) = 0
Next

End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
       khbl = 4
Formw202.Show
End Select
End Sub

Private Sub Text2_Change()
Data2.RecordSource = "select * from wxjl where ����='" & Text2.Text & "' order by ���,��ɫ,���,λ��"
Data2.Refresh
End Sub


Private Sub WXJGPZ(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Data1.RecordSource = "SELECT * FROM CLZZPZ WHERE instr(�Ƶ�,'��Э-����')>0 AND ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data1.Refresh
If Not Data1.Recordset.EOF Then
If MsgBox("���мӹ�����ƾ֤���Ƿ��������ɣ�", vbYesNo) = vbNo Then Exit Sub
Data5.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(�Ƶ�,'��Э-����')>0 and ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data4.RecordSource = "select * from WXJL where ���� between cdate('" & dt1 & "') and cdate('" & dt2 & "')"
Data4.Refresh
If Not Data4.Recordset.EOF Then
Data4.RecordSource = "select ��λ,format(SUM(val(����)*val(����)),'#0.00') from WXJL where ���� between cdate('" & dt1 & "') and cdate('" & dt2 & "') GROUP BY ��λ"
Data4.Refresh

Data1.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data1.Refresh

If Not Data1.Recordset.EOF Then
Data5.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data5.Refresh
PZH = "5-" + Trim(Data5.Recordset.Fields(0) + 1)
Else
PZH = "5-1"
End If

Data4.Recordset.MoveFirst
KLLLL = 1

Do While Not Data4.Recordset.EOF
For i = 1 To 3
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = "�ӹ���"
Data1.Recordset.Fields(1) = "ԭ����"
Data1.Recordset.Fields(2) = ""
Data1.Recordset.Fields(3) = "Ӧ���˿�"
Data1.Recordset.Fields(4) = Data4.Recordset.Fields(0)
Data1.Recordset.Fields(5) = Format(Data4.Recordset.Fields(1) * 0.83, "#0.00")
Data1.Recordset.Fields(6) = PZH
Data1.Recordset.Fields(7) = CDate(dt3)
Data1.Recordset.Fields(8) = ""
Data1.Recordset.Fields(9) = ""
Data1.Recordset.Fields(10) = ""
Data1.Recordset.Fields(11) = "��Э-����"
Data1.Recordset.Update


Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = "�ӹ���"
Data1.Recordset.Fields(1) = "Ӧ��˰��"
Data1.Recordset.Fields(2) = "˰�����"
Data1.Recordset.Fields(3) = "Ӧ���˿�"
Data1.Recordset.Fields(4) = Data4.Recordset.Fields(0)
Data1.Recordset.Fields(5) = Format(Data4.Recordset.Fields(1) * 0.17, "#0.00")
Data1.Recordset.Fields(6) = PZH
Data1.Recordset.Fields(7) = CDate(dt3)
Data1.Recordset.Fields(8) = ""
Data1.Recordset.Fields(9) = ""
Data1.Recordset.Fields(10) = ""
Data1.Recordset.Fields(11) = "��Э-����"
Data1.Recordset.Update


Data4.Recordset.MoveNext
If Data4.Recordset.EOF Then
MsgBox ("��Э�ӹ���ת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Data1.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data1.Refresh

If Not Data1.Recordset.EOF Then
Data5.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data5.Refresh
PZH = "5-" + Trim(Data5.Recordset.Fields(0) + 1)
Else
PZH = "5-1"
End If
Loop
MsgBox ("��Э�ӹ���ת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
End If

End Sub

