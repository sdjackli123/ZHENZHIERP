VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw49 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�����̵����"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form49"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0FF&
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
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3960
      Width           =   495
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ϲ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
      Width           =   495
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   6495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   9960
      TabIndex        =   18
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   81461249
      CurrentDate     =   39921
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���½��ת�뱨��"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����̴��ӡ"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ۿ��ˢ��"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   1095
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȡƽ������"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   9600
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw49.frx":0000
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
   End
   Begin VB.TextBox Text1111 
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ʼˢ��"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1095
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���½��ת���¿�"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1095
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ܿ��̴��ӡ"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw49.frx":0014
      Height          =   8055
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   14208
      _Version        =   393216
      BackColorFixed  =   8421631
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw49.frx":0028
      Height          =   330
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "��������"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formw49.frx":003C
      Height          =   330
      Left            =   1560
      TabIndex        =   13
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formw49.frx":0050
      Height          =   330
      Left            =   3480
      TabIndex        =   24
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "���Ϲ��"
      Text            =   "DBCombo1"
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ת����"
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
      Left            =   9960
      TabIndex        =   25
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   23
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
      TabIndex        =   14
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
      Index           =   0
      Left            =   360
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Formw49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public r, c, FD As Integer: Public K1, K2 As String

Private Sub Command1_Click()
Data1.RecordSource = "SELECT KCBBLSH.��������,KCBBLSH.���Ϲ��,KCBBLSH.���ϵ�λ,KCBBLSH.��ɫ,KCBBLSH.����,KCBBLSH.����,KCBBLSH.���½������,KCBBLSH.���½����,KCBBLSH.�����������,KCBBLSH.���������,KCBBLSH.���³�������,KCBBLSH.���³�����,KCBBLSH.���ۿ�� as ���½������,KCBBLSH.���۽�� AS ���½���� from KCBBLSH ORDER BY KCBBLSH.��������,KCBBLSH.���Ϲ��"
Data1.Refresh
Call OutDataToExcel4(MSFlexGrid1, 8, 10, 12, 14, "�̴��ӡ")
End Sub

Private Sub Command10_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET ���۽��=format(���½����+���������-���³�����,'#0.00'),���ۿ��=format(���½������+�����������-���³�������,'#0.00')"
Data1.RecordSource = "KCBBLSH"
Data1.Refresh
End Sub

Private Sub Command11_Click()
If DBCombo1.Text = "" Then
MsgBox ("ѡ�����")
Exit Sub
End If

Data1.RecordSource = "SELECT KCBBLSH.��������,KCBBLSH.���Ϲ��,KCBBLSH.���ϵ�λ,KCBBLSH.��ɫ,KCBBLSH.����,KCBBLSH.����,KCBBLSH.���½������,KCBBLSH.���½����,KCBBLSH.�����������,KCBBLSH.���������,KCBBLSH.���³�������,KCBBLSH.���³�����,KCBBLSH.���ۿ�� as ���½������,KCBBLSH.���۽�� AS ���½���� from KCBBLSH WHERE KCBBLSH.BL='" & DBCombo1.Text & "' ORDER BY KCBBLSH.��������,KCBBLSH.���Ϲ��"
Data1.Refresh
FD = 9
Call OutDataToExcel3(MSFlexGrid1, 10, 12, 14, DBCombo1.Text + "  �̴��ӡ")
End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command13_Click()
If MsgBox("��ȷ�ϣ���������Ϊ��" + Str(DTPicker1.Value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("����ȷ�ϣ���������Ϊ��" + Str(DTPicker1.Value), vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "DELETE * FROM kcbbjl WHERE �̴�����=CDATE('" & DTPicker1.Value & "')"
Data1.Database.Execute "INSERT INTO kcbbjl(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,���½������,���½����,�����������,���������,���³�������,���³�����,BL,���ۿ��,���۽��,ʵ�ʿ��,ʵ�ʽ��)  SELECT ����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,���½������,���½����,�����������,���������,���³�������,���³�����,BL,���ۿ��,���۽��,ʵ�ʿ��,ʵ�ʽ�� FROM KCBBLSH"
Data1.Database.Execute "UPDATE kcbbjl SET �̴�����=CDATE('" & DTPicker1.Value & "') where �̴�����=null"
MsgBox ("�����ɹ���")
End Sub

Private Sub Command14_Click()

End Sub

Private Sub Command15_Click()
If MsgBox("ȷ�������ϲ��𣿣�", vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "DELETE * FROM KCBBLSH1"
Data1.Database.Execute "INSERT INTO KCBBLSH1(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,BL,���½������,���½����,�����������,���������,���³�������,���³�����,���ۿ��,���۽��,ʵ�ʿ��,ʵ�ʽ��) SELECT ����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,BL,format(SUM(���½������),'#0.00'),format(SUM(���½����),'#0.00'),format(SUM(�����������),'#0.00'),format(SUM(���������),'#0.00'),format(SUM(���³�������),'#0.00'),format(SUM(���³�����),'#0.00'),format(SUM(���ۿ��),'#0.00'),format(SUM(���۽��),'#0.00'),format(SUM(ʵ�ʿ��),'#0.00'),format(SUM(ʵ�ʽ��),'#0.00') FROM KCBBLSH GROUP BY ����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,BL"
Data1.Database.Execute "DELETE * FROM KCBBLSH"
Data1.Database.Execute "INSERT INTO KCBBLSH SELECT * FROM KCBBLSH1 "
MsgBox ("�����ɹ�����")
End Sub


Private Sub Command16_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET ����=''"
Data1.Refresh
MsgBox ("����ɹ���")
End Sub

Private Sub Command17_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET ����=''"
Data1.Refresh
MsgBox ("������ճɹ���")
End Sub

Private Sub Command18_Click()
Call PCOutDataToExcel(MSFlexGrid1)
End Sub


Private Sub Command3_Click()
Data1.Database.Execute "DELETE * FROM KCBBLSH "
Data1.Refresh
End Sub

Private Sub Command4_Click()
On Error Resume Next

Data1.RecordSource = "SELECT * FROM KCBBLSH ORDER BY ��������,��ɫ"
Data1.Refresh
If Data1.Recordset.EOF Then
MsgBox ("��ת���¼����ֹ")
Exit Sub
Else
If MsgBox("��ȷ�ϣ���������Ϊ��" + Str(DTPicker1.Value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("ȷ��ת����", vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "delete * from kcjl where ����=CDATE('" & DTPicker1.Value & "')"
Data1.Database.Execute "INSERT INTO  KCJL (����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,���,BL)  SELECT ����,KCBBLSH.��������,KCBBLSH.���Ϲ��,KCBBLSH.���ϵ�λ,KCBBLSH.��ɫ,KCBBLSH.����,KCBBLSH.����,KCBBLSH.ʵ�ʿ��,KCBBLSH.ʵ�ʽ��,KCBBLSH.BL FROM KCBBLSH"
Data1.Database.Execute "UPDATE KCJL SET KCJL.����=CDATE('" & DTPicker1.Value & "') WHERE kcjl.����=NULL "
Data1.RecordSource = "kcjl"
Data1.Refresh
MsgBox ("ת��ɹ���")
End If
End Sub

Private Sub Command5_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET ʵ�ʿ��=���ۿ��"
Data1.Database.Execute "UPDATE KCBBLSH SET ʵ�ʽ��=���۽��"
Data1.Refresh
MsgBox ("ת��ɹ���")

End Sub

Private Sub Command6_Click()
On Error Resume Next
Data1.Database.Execute "UPDATE KCBBLSH SET ����='' WHERE ����=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET ��������='' WHERE ��������=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET ���Ϲ��='' WHERE ���Ϲ��=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET ���ϵ�λ='' WHERE ���ϵ�λ=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET ��ɫ='' WHERE ��ɫ=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET ����='' WHERE ����=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET BL='' WHERE BL=NULL"
Data1.RecordSource = "KCBBLSH"
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (Data1.Recordset.Fields(8) + Data1.Recordset.Fields(10)) > 0 Then
Data1.Recordset.Edit
Data1.Recordset.Fields(6) = (Data1.Recordset.Fields(9) + Data1.Recordset.Fields(11)) / (Data1.Recordset.Fields(8) + Data1.Recordset.Fields(10))
Data1.Recordset.Update
End If
Data1.Recordset.MoveNext
Loop
Data1.Refresh
End Sub

Private Sub Command7_Click()
Unload Me
End Sub


Private Sub Command9_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET ����=format((���½����+���������-���³�����)/(���½������+�����������-���³�������),'#0.00') where (���½������+�����������-���³�������)<>0"
Data1.Database.Execute "UPDATE KCBBLSH SET ����=0.00 where (���½������+�����������-���³�������)=0"
Data1.Refresh
End Sub

Private Sub DBCombo1_Change()
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM KCBBLSH ORDER BY ��������,��ɫ"
Data1.Refresh
Data3.RecordSource = "SELECT KCBBLSH.�������� FROM KCBBLSH  GROUP BY KCBBLSH.��������"
Data3.Refresh
Else
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE KCBBLSH.BL='" & DBCombo1.Text & "' ORDER BY ��������,��ɫ"
Data1.Refresh
Data3.RecordSource = "SELECT KCBBLSH.�������� FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' GROUP BY KCBBLSH.��������"
Data3.Refresh
End If

End Sub

Private Sub DBCombo1_Click(Area As Integer)
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM KCBBLSH ORDER BY ��������,��ɫ"
Data1.Refresh
Data3.RecordSource = "SELECT KCBBLSH.�������� FROM KCBBLSH  GROUP BY KCBBLSH.��������"
Data3.Refresh
Else
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE KCBBLSH.BL='" & DBCombo1.Text & "' ORDER BY ��������,��ɫ"
Data1.Refresh
Data3.RecordSource = "SELECT KCBBLSH.�������� FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' GROUP BY KCBBLSH.��������"
Data3.Refresh
End If
End Sub

Private Sub DBCombo2_Change()
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' AND INSTR(KCBBLSH.��������,'" & DBCombo2.Text & "')>0 ORDER BY ��������,��ɫ"
Data1.Refresh
Data7.RecordSource = "SELECT KCBBLSH.���Ϲ�� FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' AND ��������='" & DBCombo2.Text & "' GROUP BY KCBBLSH.���Ϲ��"
Data7.Refresh

End Sub

Private Sub DBCombo2_Click(Area As Integer)
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' AND INSTR(KCBBLSH.��������,'" & DBCombo2.Text & "')>0 ORDER BY ��������,��ɫ"
Data1.Refresh
Data7.RecordSource = "SELECT KCBBLSH.���Ϲ�� FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' AND ��������='" & DBCombo2.Text & "' GROUP BY KCBBLSH.���Ϲ��"
Data7.Refresh

End Sub

Private Sub DBCombo3_Change()
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE  INSTR(KCBBLSH.��ɫ,'" & DBCombo3.Text & "')>0 AND INSTR(KCBBLSH.��������,'" & DBCombo2.Text & "')>0 ORDER BY ��������,��ɫ"
Data1.Refresh
End Sub

Private Sub DBCombo3_Click(Area As Integer)
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE  INSTR(KCBBLSH.��ɫ,'" & DBCombo3.Text & "')>0 AND INSTR(KCBBLSH.��������,'" & DBCombo2.Text & "')>0 ORDER BY ��������,��ɫ"
Data1.Refresh
End Sub

Private Sub DBCombo4_Change()
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE  INSTR(���Ϲ��,'" & DBCombo4.Text & "')>0 AND INSTR(KCBBLSH.��������,'" & DBCombo2.Text & "')>0 ORDER BY ��������,���Ϲ��"
Data1.Refresh
End Sub

Private Sub DBCombo4_Click(Area As Integer)
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE  INSTR(���Ϲ��,'" & DBCombo4.Text & "')>0 AND INSTR(KCBBLSH.��������,'" & DBCombo2.Text & "')>0 ORDER BY ��������,���Ϲ��"
Data1.Refresh
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""
DBCombo4.Text = ""
DTPicker1.Value = Date
Data1.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CKGL.MDB"
Data1.RecordSource = "KCBBLSH"
Data1.Refresh
Data2.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CKGL.MDB"
Data2.RecordSource = "SELECT KL.MC FROM KL GROUP BY KL.MC"
Data2.Refresh
Data3.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CKGL.MDB"
Data3.RecordSource = "SELECT KCBBLSH.�������� FROM KCBBLSH GROUP BY KCBBLSH.��������"
Data3.Refresh

Data4.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\SCZYJHD.MDB"
Data4.RecordSource = "SELECT YS.YS FROM YS GROUP BY YS.YS"
Data4.Refresh

Data5.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\SCZYJHD.MDB"
Data6.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CJBB.MDB"
Data6.RecordSource = "RQSD"
Data6.Refresh

Data7.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CKGL.MDB"
MSFlexGrid1.ColWidth(0) = 200

End Sub

Private Sub MSFlexGrid1_Click()
FD = MSFlexGrid1.Col
End Sub

Private Sub MSFlexGrid1_dblClick()
With MSFlexGrid1
    c = .Col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlexGrid1_dblClick
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid1.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid1.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid1.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data1.Recordset.MoveFirst
Data1.Recordset.Move r - 1
Data1.Recordset.Edit
Data1.Recordset.Fields(c - 1) = Text1111.Text
Data1.Recordset.Update
Text1111.Visible = False
End Sub

