VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw307 
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
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��Ų��Ͻ���"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1455
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "״̬"
      Height          =   855
      Left            =   11160
      TabIndex        =   3
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000C0C0&
         Caption         =   "ȫ��"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C0C0&
         Caption         =   "����"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "����"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���Ų��Ͻ���"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1455
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
      Left            =   9960
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
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   93192193
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   93192193
      CurrentDate     =   36892
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   360
      Left            =   4440
      TabIndex        =   11
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw307.frx":0000
      Height          =   7455
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   8421631
      BackColorBkg    =   4109501
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   360
      Left            =   4440
      TabIndex        =   13
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ѡ����"
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
      Left            =   3360
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ѡ�񵥺�"
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
      Left            =   3360
      TabIndex        =   14
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Formw307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public C, R As Integer
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call jhjd(MSFlexGrid1, "�ƻ�����")
End Sub

Private Sub Command3_Click()
On Error Resume Next
If DBCombo1.text = "" Then
MsgBox ("�����뵥��")
Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''��Ʒ����
LO = "D:\���ݿ�\htgl\2011\scjd.mdb"
Data3.Database.Execute "delete * from sjjd"
Data1.Database.Execute "insert into SJJD(����,���,��ɫ,����,��������,���) in'" & LO & "' select ����,���,��ɫ,����,������,��� from sczy_xdh where ����='" & DBCombo1.text & "'"
Data7.RecordSource = "select * from SJJD"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.Recordset.MoveFirst
Do While Not Data7.Recordset.EOF
Data3.RecordSource = "select * from ypjd where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
Data7.Recordset.Edit
If Not Data3.Recordset.EOF Then
Data7.Recordset.Fields(4) = Data3.Recordset.Fields(5)
Data7.Recordset.Fields(5) = Data3.Recordset.Fields(6)
Data7.Recordset.Fields(14) = Data3.Recordset.Fields(12)
Data7.Recordset.Fields(15) = Data3.Recordset.Fields(11)
Data7.Recordset.Fields(16) = Data3.Recordset.Fields(14)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ü�
Data3.RecordSource = "select ���,��ɫ from cjrb where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.RecordSource = "select sum(val(�ü�)) from cjrb where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' group by ���,��ɫ"
Data3.Refresh
Data7.Recordset.Fields(17) = Data3.Recordset.Fields(0)
Else
Data7.Recordset.Fields(17) = "0"
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''��ӡ
Data3.RecordSource = "select ���,��ɫ from wxrk where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ���='��ӡ'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(����)) from wxrk where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ���='��ӡ'"
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

Data3.RecordSource = "select ���,��ɫ from cpk where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(����)) from cpk where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
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
Data6.RecordSource = "select ��ʽ,��ɫ from clb where ��ʽ='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ����='��װ'"
Data6.Refresh
If Not Data6.Recordset.EOF Then
Data9.RecordSource = "select sum(����) from clb where ��ʽ='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ����='��װ'"
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
Data10.RecordSource = "select ���,��� from LSRK where ���='" & Data7.Recordset.Fields(1) & "' and ���='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(����) from LSRK where ���='" & Data7.Recordset.Fields(1) & "' and ���='" & Data7.Recordset.Fields(2) & "'"
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
Data3.RecordSource = "select ���,��ɫ from zxd where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(�ϼƼ�)) from zxd where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
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
Data2.RecordSource = "select ����,���,��ɫ,����,����,����,��ӡ,������,��ǰ��,����ü�,�����ӡ,�������,�������,��Ʒ���,��Ʒ����,�������� FROM sjjd  order by ���,��ɫ,���"
Data2.Refresh

End Sub

Private Sub Command4_Click()
On Error Resume Next
If DBCombo1.text = "" Then
MsgBox ("�����뵥��")
Exit Sub
End If
Data8.Database.Execute "DELETE * FROM cljd"
Data1.Database.Execute "DELETE * FROM CKGL"
Data4.Database.Execute "INSERT INTO CKGL(����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������) IN'D:\���ݿ�\htgl\2011\SCZYJHD.MDB' SELECT ����,��Լ��,����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,SUM(����) FROM CKGL WHERE CKGL.����='" & DBCombo1.text & "' GROUP BY ����,��Լ��,����,��������,���Ϲ��,���ϵ�λ,��ɫ,���� "
Data11.Database.Execute "insert into ckgl(����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������) IN'D:\���ݿ�\htgl\2011\SCZYJHD.MDB' select ����,���,'1���Ͽ�',��������,��������,��λ,��ɫ,��������,�������� from rsrk where ����='" & DBCombo1.text & "'"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',�ɹ�����=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�ɹ�����) SELECT CGCLB.����,���,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������,SUM(CGCLB.��������) AS �ɹ����� FROM CGCLB WHERE CGCLB.���='" & DBCombo2.text & "' GROUP BY CGCLB.����,���,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',�������=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO cljd(����,���,����,����,���,ɫ��,����,����,����,Ƿ��) IN'D:\���ݿ�\htgl\2011\scjd.MDB' SELECT ����,CKGL.���,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.������ɫ,CKGL.��������,format(SUM(CKGL.�ɹ�����),'#0.00') AS �ɹ���,format(SUM(CKGL.�������),'#0.00') AS �����,format(SUM(CKGL.�ɹ�����-CKGL.�������),'#0.00') FROM CKGL  GROUP BY ����,CKGL.���,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.������ɫ,CKGL.��������"
Data2.RecordSource = "select * FROM cljd  order by ����,���,��ɫ"
Data2.Refresh
End Sub


Private Sub Command5_Click()
On Error Resume Next
If DBCombo2.text = "" Then
MsgBox ("��������")
Exit Sub
End If
Data8.Database.Execute "DELETE * FROM cljd"
Data1.Database.Execute "DELETE * FROM CKGL"
Data4.Database.Execute "INSERT INTO CKGL(����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������) IN'D:\���ݿ�\htgl\2011\SCZYJHD.MDB' SELECT ����,��Լ��,����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,SUM(����) FROM CKGL WHERE CKGL.��Լ��='" & DBCombo2.text & "' GROUP BY ����,��Լ��,����,��������,���Ϲ��,���ϵ�λ,��ɫ,���� "
Data11.Database.Execute "insert into ckgl(����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������) IN'D:\���ݿ�\htgl\2011\SCZYJHD.MDB' select ����,���,'1���Ͽ�',��������,��������,��λ,��ɫ,��������,�������� from rsrk where ���='" & DBCombo2.text & "'"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',�ɹ�����=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�ɹ�����) SELECT CGCLB.����,���,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������,SUM(CGCLB.��������) AS �ɹ����� FROM CGCLB WHERE CGCLB.���='" & DBCombo2.text & "' GROUP BY CGCLB.����,���,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',�������=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO cljd(����,���,����,����,���,ɫ��,����,����,����,Ƿ��) IN'D:\���ݿ�\htgl\2011\scjd.MDB' SELECT ����,CKGL.���,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.������ɫ,CKGL.��������,format(SUM(CKGL.�ɹ�����),'#0.00') AS �ɹ���,format(SUM(CKGL.�������),'#0.00') AS �����,format(SUM(CKGL.�ɹ�����-CKGL.�������),'#0.00') FROM CKGL  GROUP BY ����,CKGL.���,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.������ɫ,CKGL.��������"
Data2.RecordSource = "select * FROM cljd  order by ����,���,��ɫ"
Data2.Refresh
End Sub

Private Sub Command6_Click()
On Error Resume Next
If DBCombo2.text = "" Then
MsgBox ("��������")
Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''��Ʒ����
LO = "D:\���ݿ�\htgl\2011\scjd.mdb"
Data3.Database.Execute "delete * from sjjd"
Data1.Database.Execute "insert into SJJD(����,���,��ɫ,����,��������,���) in'" & LO & "' select ����,���,��ɫ,����,������,��� from sczy_xdh where ���='" & DBCombo2.text & "'"
Data7.RecordSource = "select * from SJJD"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.Recordset.MoveFirst
Do While Not Data7.Recordset.EOF
Data3.RecordSource = "select * from ypjd where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
Data7.Recordset.Edit
If Not Data3.Recordset.EOF Then
Data7.Recordset.Fields(4) = Data3.Recordset.Fields(5)
Data7.Recordset.Fields(5) = Data3.Recordset.Fields(6)
Data7.Recordset.Fields(14) = Data3.Recordset.Fields(12)
Data7.Recordset.Fields(15) = Data3.Recordset.Fields(11)
Data7.Recordset.Fields(16) = Data3.Recordset.Fields(14)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ü�
Data3.RecordSource = "select ���,��ɫ from cjrb where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.RecordSource = "select sum(val(�ü�)) from cjrb where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' group by ���,��ɫ"
Data3.Refresh
Data7.Recordset.Fields(17) = Data3.Recordset.Fields(0)
Else
Data7.Recordset.Fields(17) = "0"
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''��ӡ
Data3.RecordSource = "select ���,��ɫ from wxrk where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(����)) from wxrk where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
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

Data3.RecordSource = "select ���,��ɫ from cpk where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(����)) from cpk where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ���='01'"
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
Data6.RecordSource = "select ��ʽ,��ɫ from clb where ��ʽ='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ����='��װ'"
Data6.Refresh
If Not Data6.Recordset.EOF Then
Data9.RecordSource = "select sum(����) from clb where ��ʽ='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "' and ����='��װ'"
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
Data10.RecordSource = "select ���,��� from LSRK where ���='" & Data7.Recordset.Fields(1) & "' and ���='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(����) from LSRK where ���='" & Data7.Recordset.Fields(1) & "' and ���='" & Data7.Recordset.Fields(2) & "'"
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
Data3.RecordSource = "select ���,��ɫ from zxd where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(�ϼƼ�)) from zxd where ���='" & Data7.Recordset.Fields(1) & "' and ��ɫ='" & Data7.Recordset.Fields(2) & "'"
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
Data2.RecordSource = "select ����,���,��ɫ,����,����,����,��ӡ,������,��ǰ��,����ü�,�����ӡ,�������,�������,��Ʒ���,��Ʒ����,�������� FROM sjjd  order by ���,��ɫ,���"
Data2.Refresh

End Sub

Private Sub Form_Load()
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
DBCombo1.text = ""
DBCombo2.text = ""
Data1.DatabaseName = "D:\���ݿ�\htgl\2011\SCZYJHD.mdb"
Data2.DatabaseName = "D:\���ݿ�\htgl\2011\scjd.mdb"
Data3.DatabaseName = "D:\���ݿ�\htgl\2011\scjd.mdb"
Data4.DatabaseName = "D:\���ݿ�\htgl\2011\ckgl.mdb"
Data5.DatabaseName = "D:\���ݿ�\htgl\2011\CPCK.MDB"
Data6.DatabaseName = "D:\���ݿ�\htgl\2011\db.mdb"
Data7.DatabaseName = "D:\���ݿ�\htgl\2011\scjd.mdb"
Data8.DatabaseName = "D:\���ݿ�\htgl\2011\scjd.mdb"
Data9.DatabaseName = "D:\���ݿ�\htgl\2011\db.mdb"
Data10.DatabaseName = "D:\���ݿ�\htgl\2011\CPCK.MDB"
Data11.DatabaseName = "D:\���ݿ�\htgl\2011\scjd.mdb"

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 0
End Sub

Private Sub MSFlex()
On Error Resume Next
With MSFlexGrid1
    C = .Col: R = .Row    '''''C�У���R��
        Text1111.Left = .Left + .ColPos(C)
        Text1111.Top = .Top + .RowPos(R)
        Text1111.Width = .ColWidth(C)
        Text1111.Height = .RowHeight(R)
        Text1111 = .text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 2
khbl = 12
Formw202.Show
End Select
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data2.Recordset.MoveFirst
Data2.Recordset.Move R - 1
Data2.Recordset.Edit
Data2.Recordset.Fields(C - 1) = Text1111.text
Data2.Recordset.Update
Text1111.Visible = False
MSFlexGrid1.text = Text1111.text
MSFlexGrid1.SetFocus
End If
End Sub




