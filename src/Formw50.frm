VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw50 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�����̴汨��"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ƾ֤����"
      Height          =   855
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ɲ�ѯ"
      Height          =   855
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��ĺ���"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton Command5 
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw50.frx":0000
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw50.frx":0014
      Height          =   8055
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   14208
      _Version        =   393216
      BackColorFixed  =   8421631
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
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
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw50.frx":0028
      Height          =   330
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "��������"
      BoundColumn     =   "��������"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   80740353
      CurrentDate     =   39921
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   12120
      TabIndex        =   13
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   80740353
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   80740353
      CurrentDate     =   39921
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   495
      Left            =   9240
      TabIndex        =   17
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   80740353
      CurrentDate     =   39921
   End
   Begin VB.Label Label4 
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
      Index           =   1
      Left            =   9240
      TabIndex        =   18
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
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
      Index           =   0
      Left            =   7560
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   0
      Left            =   12120
      TabIndex        =   14
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
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
      Index           =   3
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "����������"
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
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Formw50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call OutDataToExcel5(MSFlexGrid1, 4, 6, 8, 10, 12, "�̴��ӡ")
End Sub

Private Sub Command2_Click()
Data1.Database.Execute "UPDATE kcbbjl SET �������=format(���ۿ��-ʵ�ʿ��,'#0.000') where �̴�����=cdate('" & DTPicker1.Value & "')"
Data1.Database.Execute "UPDATE kcbbjl SET ��Ľ��=FORMAT(���۽��-ʵ�ʽ��,'#0.00') where �̴�����=cdate('" & DTPicker1.Value & "')"
MsgBox ("�ɹ���")
End Sub

Private Sub Command3_Click()
If MsgBox("��������Ϊ��" + Trim(DTPicker2.Value) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("�����ڼ�Ϊ��" + Trim(Month(DTPicker2.Value)) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("ȷ�����ɳɱ�ϵ�е�ƾ֤��", vbYesNo) = vbNo Then Exit Sub
Call CLCKpz(CDate(DTPicker3.Value), CDate(DTPicker4.Value), CDate(DTPicker2.Value))
End Sub

Private Sub Command4_Click()
Formw332.Combo1.Text = "�ɱ�ƾ֤"
Formw332.Show
End Sub

Private Sub Command5_Click()
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT ����,��������,���Ϲ��,���ϵ�λ,���½������,���½����,�����������,���������,���³�������,���³�����,ʵ�ʿ��,ʵ�ʽ��,�������,��Ľ�� FROM kcbbjl where �̴�����=cdate('" & DTPicker1.Value & "') ORDER BY BL,��������"
Data1.Refresh
Else
Data1.RecordSource = "SELECT ����,��������,���Ϲ��,���ϵ�λ,���½������,���½����,�����������,���������,���³�������,���³�����,ʵ�ʿ��,ʵ�ʽ��,�������,��Ľ�� FROM kcbbjl where �̴�����=cdate('" & DTPicker1.Value & "') and bl='" & DBCombo1.Text & "' ORDER BY BL,��������"
Data1.Refresh
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub DBCombo1_Click(Area As Integer)
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT ����,��������,���Ϲ��,���ϵ�λ,���½������,���½����,�����������,���������,���³�������,���³�����,ʵ�ʿ��,ʵ�ʽ��,�������,��Ľ�� FROM kcbbjl where �̴�����=cdate('" & DTPicker1.Value & "') ORDER BY BL,��������"
Data1.Refresh
Data3.RecordSource = "SELECT �������� FROM kcbbjl where �̴�����=cdate('" & DTPicker1.Value & "') GROUP BY ��������"
Data3.Refresh
Else
Data1.RecordSource = "SELECT ����,��������,���Ϲ��,���ϵ�λ,���½������,���½����,�����������,���������,���³�������,���³�����,ʵ�ʿ��,ʵ�ʽ��,�������,��Ľ�� FROM kcbbjl WHERE BL='" & DBCombo1.Text & "' and  �̴�����=cdate('" & DTPicker1.Value & "') ORDER BY ��������"
Data1.Refresh
Data3.RecordSource = "SELECT �������� FROM kcbbjl WHERE BL='" & DBCombo1.Text & "' and �̴�����=cdate('" & DTPicker1.Value & "') GROUP BY ��������"
Data3.Refresh
End If
End Sub
Private Sub DBCombo2_Change()
Data1.RecordSource = "SELECT ����,��������,���Ϲ��,���ϵ�λ,���½������,���½����,�����������,���������,���³�������,���³�����,ʵ�ʿ��,ʵ�ʽ��,�������,��Ľ�� FROM kcbbjl WHERE  INSTR(��������,'" & DBCombo2.Text & "')>0 and �̴�����=cdate('" & DTPicker1.Value & "') ORDER BY ��������"
Data1.Refresh
End Sub

Private Sub DBCombo2_Click(Area As Integer)
Data1.RecordSource = "SELECT ����,��������,���Ϲ��,���ϵ�λ,���½������,���½����,�����������,���������,���³�������,���³�����,ʵ�ʿ��,ʵ�ʽ��,�������,��Ľ�� FROM kcbbjl WHERE  INSTR(��������,'" & DBCombo2.Text & "')>0 and �̴�����=cdate('" & DTPicker1.Value & "') ORDER BY ��������"
Data1.Refresh
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
DBCombo1.Text = ""
DBCombo2.Text = ""
Data1.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\ckgl.MDB"
Data1.RecordSource = "SELECT ����,��������,���Ϲ��,���ϵ�λ,���½������,���½����,�����������,���������,���³�������,���³�����,ʵ�ʿ��,ʵ�ʽ��,�������,��Ľ�� FROM kcbbjl where �̴�����=cdate('" & DTPicker1.Value & "') ORDER BY BL,��������"
Data1.Refresh
Data2.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\ckgl.MDB"
Data2.RecordSource = "select KL.MC from KL   group by KL.MC"
Data2.Refresh
Data3.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\ckgl.MDB"

Data4.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\cw.MDB"
Data5.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\cw.MDB"


MSFlexGrid1.ColWidth(0) = 200
For i = 1 To 14
MSFlexGrid1.ColWidth(i) = 1200
Next
End Sub

Public Sub CLCKpz(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Data4.RecordSource = "SELECT * FROM CLSCCB WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data4.Refresh
If Not Data4.Recordset.EOF Then
If MsgBox("���гɱ�����ƾ֤���Ƿ��������ɣ�", vbYesNo) = vbNo Then Exit Sub
Data5.Database.Execute "DELETE * FROM CLSCCB WHERE instr(�Ƶ�,'�ɱ�-����')>0 and ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If



Data4.RecordSource = "SELECT * FROM CLSCCB WHERE INSTR(ƾ֤��,'S-')>0 AND ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data4.Refresh
If Not Data4.Recordset.EOF Then
Data4.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSCCB WHERE INSTR(ƾ֤��,'S-')>0 AND ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data4.Refresh
PZH = "S-" + Trim(Data4.Recordset.Fields(0) + 1)
Else
PZH = "S-1"
End If


Data3.RecordSource = "SELECT *  FROM kcbbjl where �̴�����=cdate('" & DTPicker1.Value & "')"
Data3.Refresh
If Data3.Recordset.EOF Then Exit Sub

Data3.RecordSource = "SELECT format(sum(���³�����),'#0.00') FROM kcbbjl where �̴�����=cdate('" & DTPicker1.Value & "')"
Data3.Refresh

If Not Data3.Recordset.EOF Then
Data3.Recordset.MoveFirst
KLLLL = 1
Do While Not Data3.Recordset.EOF
For i = 1 To 3
Data4.Recordset.AddNew
Data4.Recordset.Fields(0) = "����ԭ����"
Data4.Recordset.Fields(1) = "�����ɱ�"
Data4.Recordset.Fields(2) = "ֱ�������ɱ�"
Data4.Recordset.Fields(3) = "ԭ����"
Data4.Recordset.Fields(4) = ""
Data4.Recordset.Fields(5) = Data3.Recordset.Fields(0)
Data4.Recordset.Fields(6) = PZH
Data4.Recordset.Fields(7) = dt3
Data4.Recordset.Fields(8) = ""
Data4.Recordset.Fields(9) = ""
Data4.Recordset.Fields(10) = ""
Data4.Recordset.Fields(11) = "�ɱ�-����"
Data4.Recordset.Update
Data3.Recordset.MoveNext
If Data3.Recordset.EOF Then
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ɱ�ƾ֤")
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Data4.RecordSource = "SELECT * FROM CLSCCB WHERE INSTR(ƾ֤��,'S-')>0 AND ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data4.Refresh
If Not Data4.Recordset.EOF Then
Data4.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSCCB WHERE INSTR(ƾ֤��,'S-')>0 AND ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data4.Refresh
PZH = "S-" + Trim(Data4.Recordset.Fields(0) + 1)
Else
PZH = "S-1"
End If
Loop
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ɱ�ƾ֤")
End If
End Sub

