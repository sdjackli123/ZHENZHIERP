VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formy307 
   BackColor       =   &H00C0E0FF&
   Caption         =   "ʵ�ʽ���"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form32"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   7440
      TabIndex        =   13
      Top             =   600
      Width           =   3135
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Data Data14 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data13 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data12 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9930
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "״̬"
      Height          =   855
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000C0C0&
         Caption         =   "ȫ��"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C0C0&
         Caption         =   "����"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "����"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
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
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80609281
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80609281
      CurrentDate     =   36892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy307.frx":0000
      Height          =   7455
      Left            =   3360
      TabIndex        =   8
      Top             =   1680
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   8421631
      BackColorBkg    =   4109501
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6975
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   12303
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ���ڣ�"
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
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "�������ڣ�"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Formy307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call JHJD(MSFlexGrid1, "�ƻ�����")
End Sub

Private Sub dhjd(DH As String)
On Error Resume Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�������
lo = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data3.Database.Execute "delete * from sjjd"
Data1.Database.Execute "insert into SJJD(����,���,��ɫ,Ʒ��,����,��������,���) in'" & lo & "' select ����,���,��ɫ,Ʒ��,����,����,��� from sczy_xdh where ����='" & DH & "'"
Data7.RecordSource = "select * from SJJD"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.Recordset.MoveFirst
Do While Not Data7.Recordset.EOF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ü�
Data7.Recordset.Edit
Data3.RecordSource = "select ���,��ɫ from cjrb where ����='" & Data7.Recordset.Fields(0) & "' and  ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.RecordSource = "select sum(val(�ü�)) from cjrb where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
Data7.Recordset.Fields(17) = Data3.Recordset.Fields(0)
Else
Data7.Recordset.Fields(17) = "0"
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''��ӡ
Data3.RecordSource = "select ���,��ɫ from wxrk where ����='" & Data7.Recordset.Fields(0) & "' and  ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ���='��ӡ'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(����)) from wxrk where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ���='��ӡ'"
Data8.Refresh
Data7.Recordset.Fields(18) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(18) = "0"
Else
Data7.Recordset.Fields(18) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(18) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''��ת��

Data3.RecordSource = "select ���,��ɫ from cpk where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(����)) from cpk where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data8.Refresh
Data7.Recordset.Fields(19) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(19) = "0"
Else
Data7.Recordset.Fields(19) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(19) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''����
Data6.RecordSource = "select ��ʽ,��ɫ from clb where ����='" & Data7.Recordset.Fields(0) & "' and ��ʽ='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ����='��װ'"
Data6.Refresh
If Not Data6.Recordset.EOF Then
Data9.RecordSource = "select sum(����) from clb where ����='" & Data7.Recordset.Fields(0) & "' and ��ʽ='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ����='��װ'"
Data9.Refresh
Data7.Recordset.Fields(20) = "0"
If Data9.Recordset.EOF Then
Data7.Recordset.Fields(20) = "0"
Else
Data7.Recordset.Fields(20) = Data9.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(20) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''��Ʒ���
Data10.RecordSource = "select ���,��� from LSRK where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ���='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(����) from LSRK where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ���='" & Data7.Recordset.Fields(2) & "'"
Data5.Refresh
Data7.Recordset.Fields(21) = "0"
If Data5.Recordset.EOF Then
Data7.Recordset.Fields(21) = "0"
Else
Data7.Recordset.Fields(21) = Data5.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(21) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''��Ʒ����
Data10.RecordSource = "select ���,��� from lsfh where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ���='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(����) from lsfh where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data5.Refresh
Data7.Recordset.Fields(22) = "0"
If Data5.Recordset.EOF Then
Data7.Recordset.Fields(22) = "0"
Else
Data7.Recordset.Fields(22) = Data5.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(22) = "0"
End If
Data7.Recordset.Update
Data7.Recordset.MoveNext
Loop
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���ϲ���
Data7.Database.Execute "update sjjd set ���='0' where ���=null"

Data8.RecordSource = "select ����,����ü�,��Ʒ��� from sjjd where ����='" & DH & "'"
Data8.Refresh
If Not Data8.Recordset.EOF Then
Data8.Recordset.MoveFirst
pd = 0
Do While Not Data8.Recordset.EOF
If Data8.Recordset.Fields(1) >= Data8.Recordset.Fields(0) And Data8.Recordset.Fields(1) = Data8.Recordset.Fields(2) Then
Else
pd = pd + 1
End If
Data8.Recordset.MoveNext
Loop
If pd = 0 Then
Data12.Database.Execute "update sczy_xdh set ����='����' where ����='" & DH & "'"
End If
End If


Data2.RecordSource = "select ����,���,Ʒ��,��ɫ,����,����ü�,�����ӡ,�������,�������,��Ʒ���,��Ʒ����,�������� FROM sjjd  order by ����,���,��ɫ,���"
Data2.Refresh

End Sub

Private Sub dhcl(DH As String)
On Error Resume Next
Data8.Database.Execute "DELETE * FROM cljd"
Data1.Database.Execute "DELETE * FROM CKGL"
Data4.Database.Execute "INSERT INTO CKGL(����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������) IN'd:\���ݿ�\\htgl\2011\SCZYJHD.MDB' SELECT ����,����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,SUM(����) FROM CKGL WHERE CKGL.����='" & DH & "' GROUP BY ����,����,��������,���Ϲ��,���ϵ�λ,��ɫ,���� "
Data11.Database.Execute "insert into ckgl(����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������) IN'd:\���ݿ�\\htgl\2011\SCZYJHD.MDB' select ����,'1���Ͽ�',��������,��������,��λ,��ɫ,��������,�������� from rsrk where ����='" & DH & "'"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',�ɹ�����=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�ɹ�����) SELECT CGCLB.����,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������,SUM(CGCLB.��������) AS �ɹ����� FROM CGCLB WHERE instr(����,'" & DH & "')>0 GROUP BY CGCLB.����,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',�������=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO cljd(����,����,����,���,ɫ��,����,����,����,Ƿ��) IN'd:\���ݿ�\\htgl\2011\scjd.MDB' SELECT ����,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.������ɫ,CKGL.��������,Format(SUM(CKGL.�ɹ�����),'#0.00') AS �ɹ���,Format(SUM(CKGL.�������),'#0.00') AS �����,Format(SUM(CKGL.�ɹ�����-CKGL.�������),'#0.00') FROM CKGL  GROUP BY ����,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.������ɫ,CKGL.��������"
Data2.RecordSource = "select * FROM cljd  order by ����,��ɫ"
Data2.Refresh
End Sub


Private Sub khcl(kh As String)
On Error Resume Next
If Option1.Value = True Then
Data13.RecordSource = "select ���� from sczy_xdh where �ͻ�='" & kh & "' and ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
Data13.Refresh
End If
If Option2.Value = True Then
Data13.RecordSource = "select ���� from sczy_xdh where �ͻ�='" & kh & "' and ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
Data13.Refresh
End If
If Option3.Value = True Then
Data13.RecordSource = "select ���� from sczy_xdh where �ͻ�='" & kh & "' and ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data13.Refresh
End If

If Not Data13.Recordset.EOF Then
Data8.Database.Execute "DELETE * FROM cljd"
Data1.Database.Execute "DELETE * FROM CKGL"
Data13.Recordset.MoveFirst
Do While Not Data13.Recordset.EOF
Data4.Database.Execute "INSERT INTO CKGL(����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������) IN'd:\���ݿ�\\htgl\2011\SCZYJHD.MDB' SELECT ����,����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,SUM(����) FROM CKGL WHERE ����='" & Data13.Recordset.Fields(0) & "' GROUP BY ����,����,��������,���Ϲ��,���ϵ�λ,��ɫ,���� "
Data11.Database.Execute "insert into ckgl(����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������) IN'd:\���ݿ�\\htgl\2011\SCZYJHD.MDB' select ����,'1���Ͽ�',��������,��������,��λ,��ɫ,��������,�������� from rsrk where instr(����,'" & Data13.Recordset.Fields(0) & "')>0"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',�ɹ�����=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�ɹ�����) SELECT CGCLB.����,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������,SUM(CGCLB.��������) AS �ɹ����� FROM CGCLB WHERE instr(����,'" & Data13.Recordset.Fields(0) & "')>0 GROUP BY CGCLB.����,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',�������=0 WHERE LX=NULL"
Data13.Recordset.MoveNext
Loop
End If
Data1.Database.Execute "INSERT INTO cljd(����,����,����,���,ɫ��,����,����,����,Ƿ��) IN'd:\���ݿ�\\htgl\2011\scjd.MDB' SELECT ����,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.������ɫ,CKGL.��������,Format(SUM(CKGL.�ɹ�����),'#0.00') AS �ɹ���,Format(SUM(CKGL.�������),'#0.00') AS �����,Format(SUM(CKGL.�ɹ�����-CKGL.�������),'#0.00') FROM CKGL  GROUP BY ����,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.������ɫ,CKGL.��������"
Data2.RecordSource = "select * FROM cljd  order by ����,��ɫ"
Data2.Refresh
End Sub

Private Sub khjd(kh As String)
On Error Resume Next

On Error Resume Next
If Option1.Value = True Then
Data13.RecordSource = "select ���� from sczy_xdh where �ͻ�='" & kh & "' and ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
Data13.Refresh
End If
If Option2.Value = True Then
Data13.RecordSource = "select ���� from sczy_xdh where �ͻ�='" & kh & "' and ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
Data13.Refresh
End If
If Option3.Value = True Then
Data13.RecordSource = "select ���� from sczy_xdh where �ͻ�='" & kh & "' and ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data13.Refresh
End If

lo = "d:\���ݿ�\\htgl\2011\scjd.mdb"

If Not Data13.Recordset.EOF Then
Data3.Database.Execute "delete * from sjjd"
Data13.Recordset.MoveFirst
Do While Not Data13.Recordset.EOF
Data1.Database.Execute "insert into SJJD(����,���,��ɫ,Ʒ��,����,��������,���) in'" & lo & "' select ����,���,��ɫ,Ʒ��,����,����,��� from sczy_xdh where instr(����,'" & Data13.Recordset.Fields(0) & "')>0"
Data13.Recordset.MoveNext
Loop
End If

lo = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data7.RecordSource = "select * from SJJD"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.Recordset.MoveFirst
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ü�
Do While Not Data7.Recordset.EOF
Data7.Recordset.Edit
Data3.RecordSource = "select ���,��ɫ from cjrb where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.RecordSource = "select sum(val(�ü�)) from cjrb where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' group by ���,��ɫ"
Data3.Refresh
Data7.Recordset.Fields(17) = Data3.Recordset.Fields(0)
Else
Data7.Recordset.Fields(17) = "0"
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''��ӡ
Data3.RecordSource = "select ���,��ɫ from wxrk where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ���='��ӡ'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(����)) from wxrk where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ���='��ӡ'"
Data8.Refresh
Data7.Recordset.Fields(18) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(18) = "0"
Else
Data7.Recordset.Fields(18) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(18) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''��ת��

Data3.RecordSource = "select ���,��ɫ from cpk where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(����)) from cpk where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data8.Refresh
Data7.Recordset.Fields(19) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(19) = "0"
Else
Data7.Recordset.Fields(19) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(19) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''����
Data6.RecordSource = "select ��ʽ,��ɫ from clb where ����='" & Data7.Recordset.Fields(0) & "' and ��ʽ='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ����='��װ'"
Data6.Refresh
If Not Data6.Recordset.EOF Then
Data9.RecordSource = "select sum(����) from clb where ����='" & Data7.Recordset.Fields(0) & "' and ��ʽ='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ����='��װ'"
Data9.Refresh
Data7.Recordset.Fields(20) = "0"
If Data9.Recordset.EOF Then
Data7.Recordset.Fields(20) = "0"
Else
Data7.Recordset.Fields(20) = Data9.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(20) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''��Ʒ���
Data10.RecordSource = "select ���,��� from LSRK where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ���='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(����) from LSRK where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ���='" & Data7.Recordset.Fields(2) & "'"
Data5.Refresh
Data7.Recordset.Fields(21) = "0"
If Data5.Recordset.EOF Then
Data7.Recordset.Fields(21) = "0"
Else
Data7.Recordset.Fields(21) = Data5.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(21) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''��Ʒ����
Data3.RecordSource = "select ���,��ɫ from lsfh where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(����) from lsfh where ����='" & Data7.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data8.Refresh
Data7.Recordset.Fields(22) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(22) = "0"
Else
Data7.Recordset.Fields(22) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(22) = "0"
End If
Data7.Recordset.Update
Data7.Recordset.MoveNext
Loop
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���ϲ���
Data7.Database.Execute "update sjjd set ���='0' where ���=null"
Data2.RecordSource = "select ����,���,Ʒ��,��ɫ,����,����ü�,�����ӡ,�������,�������,��Ʒ���,��Ʒ����,�������� FROM sjjd  order by ����,���,��ɫ,���"
Data2.Refresh

End Sub

Private Sub Command7_Click()
Call tree
Call zk
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.mdb"
Data5.DatabaseName = "d:\���ݿ�\\htgl\2011\CPCK.MDB"
Data6.DatabaseName = "d:\���ݿ�\\htgl\2011\db.mdb"
Data7.DatabaseName = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data8.DatabaseName = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data9.DatabaseName = "d:\���ݿ�\\htgl\2011\db.mdb"
Data10.DatabaseName = "d:\���ݿ�\\htgl\2011\CPCK.MDB"
Data11.DatabaseName = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data12.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data13.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data14.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Option4.Value = True
Option3.Value = True
MSFlexGrid1.ColWidth(0) = 200
End Sub

Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex

   TreeView1.Nodes.Clear
 
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox ("��ѡ������״̬")
Exit Sub
End If

If Option1.Value = True Then
    Data12.RecordSource = "select distinct �ͻ� from sczy_xdh where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data13.RecordSource = "select distinct ���� from sczy_xdh where �ͻ�='" & Data12.Recordset.Fields(0) & "' and  ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
        Data13.Refresh
        
        If Not Data13.Recordset.EOF Then
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data13.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data14.RecordSource = "select distinct ��� from sczy_xdh where ����='" & Data13.Recordset.Fields(0) & "' and ����='����'"
        Data14.Refresh
        
        If Not Data14.Recordset.EOF Then
        Data14.Recordset.MoveFirst
        Do While Not Data14.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data14.Recordset.Fields(0))
        Data14.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data13.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data12.Recordset.MoveNext
        Loop
    End If
End If

If Option3.Value = True Then
    Data12.RecordSource = "select distinct �ͻ� from sczy_xdh where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data13.RecordSource = "select distinct ���� from sczy_xdh where �ͻ�='" & Data12.Recordset.Fields(0) & "' and  ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
        Data13.Refresh
        
        If Not Data13.Recordset.EOF Then
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data13.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data14.RecordSource = "select distinct ��� from sczy_xdh where ����='" & Data13.Recordset.Fields(0) & "'"
        Data14.Refresh
        
        If Not Data14.Recordset.EOF Then
        Data14.Recordset.MoveFirst
        Do While Not Data14.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data14.Recordset.Fields(0))
        Data14.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data13.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data12.Recordset.MoveNext
        Loop
    End If
End If

If Option2.Value = True Then
    Data12.RecordSource = "select distinct �ͻ� from sczy_xdh where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "x" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data13.RecordSource = "select distinct ���� from sczy_xdh where �ͻ�='" & Data12.Recordset.Fields(0) & "' and  ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
        Data13.Refresh
        
        If Not Data13.Recordset.EOF Then
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "x" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data13.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data14.RecordSource = "select distinct ��� from sczy_xdh where ����='" & Data13.Recordset.Fields(0) & "' and ����='����'"
        Data14.Refresh
        
        If Not Data14.Recordset.EOF Then
        Data14.Recordset.MoveFirst
        Do While Not Data14.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "x" + Trim(m) + "w" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data14.Recordset.Fields(0))
        Data14.Recordset.MoveNext
        Loop
        End If
        
        Data13.Recordset.MoveNext
        Loop
        End If
        Data12.Recordset.MoveNext
        Loop
    End If
End If

End Sub


'Ȼ��ô����ֻ��Խ�С�ļ�¼������ѭ�������Ч�ʱȽϸߡ��޸ĺ�Ĵ������£�

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next



If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") = 0 Then
If Option4.Value = True Then
Call khcl(TreeView1.Nodes(Node.Index).FullPath)
End If

If Option5.Value = True Then
Call khjd(TreeView1.Nodes(Node.Index).FullPath)
End If

Else

l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If

If Option4.Value = True Then
Call dhcl(Trim(l1))
End If

If Option5.Value = True Then
Call dhjd(Trim(l1))
End If

End If

'DBCombo2.Text = Node.Index
'DBCombo3.Text = TreeView1.Nodes(Node.Index).FullPath
End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True 'չ�����нڵ�
  Next i
End Sub





