VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Formw731 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ⱦ�����˲�ѯ"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form39"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ɲ�ѯ"
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ƾ֤����"
      Height          =   855
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.Data Data10 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��ת���"
      Height          =   375
      Left            =   5760
      TabIndex        =   17
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ׼��"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ת����"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw731.frx":0000
      Height          =   7455
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   10790143
      BackColorBkg    =   44718
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw731.frx":0014
      Height          =   330
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "���"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   80871425
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   80871425
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   80871425
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   12000
      TabIndex        =   20
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   80871425
      CurrentDate     =   36892
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   2
      Left            =   12000
      TabIndex        =   21
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "�������"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   15
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   240
      Width           =   855
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
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Formw731"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer

Private Sub Command1_Click()
'On Error Resume Next
Command1.Enabled = False

Data6.Database.Execute "DELETE * FROM JGZCX2"
lo = "d:\���ݿ�\bfrz\" + ljb + "\FP.MDB"       '''''''''''''''''''''''����
''''   ����Data4.Database.Execute "insert into JGZCX2(�ͻ�,�����ۼ�Ӧ��) IN'" & LO &"' SELECT MID(��ƿ�Ŀ,INSTR(��ƿ�Ŀ,'-')+1),format(SUM(VAL(���)),'#0.00') FROM PMMXJZ WHERE �������='��' AND ����=CDATE('" & RQQ & "') GROUP BY MID(��ƿ�Ŀ,INSTR(��ƿ�Ŀ,'-')+1)"
Data4.Database.Execute "insert into JGZCX2(�ͻ�,�����ۼ�Ӧ��) IN'" & lo & "' SELECT MID(��ƿ�Ŀ,INSTR(��ƿ�Ŀ,'-')+1),format(SUM(VAL(���)),'#0.00') FROM PMMXJZ WHERE �������='��' AND ����=CDATE('" & Text1.Text & "') GROUP BY MID(��ƿ�Ŀ,INSTR(��ƿ�Ŀ,'-')+1)"
'Data3.Database.Execute "insert into JGZCX2(�ͻ�,����Ӧ����) in'" & LO & "' SELECT ��Ӧ��λ,format(SUM(�ϼƽ��),'#0.00') FROM CKGL WHERE  ���� between cdate('" & Text1 & "') and cdate('" & Text2.text & "') AND ���='�ɹ����' GROUP BY ��Ӧ��λ"
Data5.Database.Execute "insert into JGZCX2(�ͻ�,����Ӧ����) in'" & lo & "' SELECT ��Ӧ��λ,format(SUM(�ϼƽ��),'#0.00') FROM MX WHERE  ���ʱ�� between cdate('" & Text1 & "') and cdate('" & Text2.Text & "') AND ���='�ɹ����' GROUP BY ��Ӧ��λ"
Data5.Database.Execute "insert into JGZCX2(�ͻ�,����Ӧ����) in'" & lo & "' SELECT ���ⵥλ,format(SUM(-�ϼƽ��),'#0.00') FROM ckMX WHERE  ����ʱ�� between cdate('" & Text1 & "') and cdate('" & Text2.Text & "')  GROUP BY ���ⵥλ"
Data3.Database.Execute "insert into JGZCX2(�ͻ�,����Ӧ����) in'" & lo & "' SELECT ��Ӧ��λ,format(SUM(�ϼƽ��),'#0.00') FROM CKGL WHERE  ���� between cdate('" & Text1 & "') and cdate('" & Text2.Text & "') AND ���='�ɹ����' and �Ƿ񸶿�<>'�Ѹ�' GROUP BY ��Ӧ��λ"
rqq = CDate(Text2.Text) + 1
Data6.Database.Execute "insert into JGZCX2(�ͻ�,���ڿ�Ʊ)  SELECT �ͻ�,��Ʊ��� FROM JHFP WHERE ��Ʊ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & rqq & "')"
Data4.Database.Execute "insert into JGZCX2(�ͻ�,�����Ѹ���) IN'" & lo & "' SELECT MID(�Է���Ŀ,INSTR(�Է���Ŀ,'-')+1),format(SUM(VAL(�������)),'#0.00') FROM TZJZMX WHERE ���� between cdate('" & Text1.Text & "') and cdate('" & Text2.Text & "') AND �������<>'0' GROUP BY MID(�Է���Ŀ,INSTR(�Է���Ŀ,'-')+1)"
Data6.Database.Execute "insert into JGZCX2(�ͻ�,�����ۼƿ�Ʊ) SELECT �ͻ�,��Ʊ��� FROM PMJHFP WHERE  ��ת����=CDATE('" & Text1.Text & "')"
Data6.Database.Execute "insert into JGZCX2(�ͻ�,�����ۼ�δ��Ʊ) SELECT �ͻ�,δ����� FROM PMJHFP WHERE  ��ת����=CDATE('" & Text1.Text & "')"

Data4.Database.Execute "insert into JGZCX2(�ͻ�,δ����) IN'" & lo & "' SELECT MID(�Է���Ŀ,INSTR(�Է���Ŀ,'-')+1),format(SUM(VAL(�������)),'#0.00') FROM TZJZMX WHERE ���� between cdate('" & Text1.Text & "') and cdate('" & Text2.Text & "') AND �������<>'0' GROUP BY MID(�Է���Ŀ,INSTR(�Է���Ŀ,'-')+1)"
Data4.Database.Execute "insert into JGZCX2(�ͻ�,δ����) IN'" & lo & "' SELECT �ͻ�,format(SUM(VAL(���)),'#0.00') FROM WDZSZ WHERE ����=cdate('" & Text1.Text & "')  GROUP BY �ͻ�"
Data6.Database.Execute "insert into JGZCX2(�ͻ�,δ����) SELECT �ͻ�,format(SUM(VAL(��Ʊ���)),'#0.00') FROM JHFP WHERE  ��Ʊ���� between cdate('" & Text1.Text & "') and cdate('" & rqq & "') GROUP BY �ͻ�"


Data6.Database.Execute "UPDATE JGZCX2 SET ���='1'"
Data6.Database.Execute "UPDATE JGZCX2 SET ���ڷ�Χ='" & Text1.Text & "'+'--'+'" & Text2.Text & "'"
Data6.Database.Execute "UPDATE JGZCX2 SET �����ۼ�Ӧ��='0' WHERE �����ۼ�Ӧ��=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET ����Ӧ����='0' WHERE ����Ӧ����=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET �����ۼ�Ӧ����='0' WHERE �����ۼ�Ӧ����=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET �����Ѹ���='0' WHERE �����Ѹ���=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET �����ۼƿ�Ʊ='0' WHERE �����ۼƿ�Ʊ=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET ���ڿ�Ʊ='0' WHERE ���ڿ�Ʊ=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET �����ۼ�δ��Ʊ='0' WHERE �����ۼ�δ��Ʊ=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET �����ۼƿ�Ʊ='0' WHERE �����ۼƿ�Ʊ=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET ����δ��='0' WHERE ����δ��=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET �����ۼ�δ��='0' WHERE �����ۼ�δ��=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET δ����='0' WHERE δ����=NULL"

Data6.Database.Execute "insert into JGZCX2(�ͻ�,���ڷ�Χ,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,����δ��,�����ۼ�δ��,δ����) SELECT �ͻ�,���ڷ�Χ,FORMAT(SUM(VAL(�����ۼ�Ӧ��)),'#0.00'),FORMAT(SUM(VAL(����Ӧ����)),'#0.00'),FORMAT(SUM(VAL(�����ۼ�Ӧ����)),'#0.00'),FORMAT(SUM(VAL(�����Ѹ���)),'#0.00'),FORMAT(SUM(VAL(�����ۼƿ�Ʊ)),'#0.00'),FORMAT(SUM(VAL(���ڿ�Ʊ)),'#0.00'),FORMAT(SUM(VAL(�����ۼƿ�Ʊ)),'#0.00'),FORMAT(SUM(VAL(�����ۼ�δ��Ʊ)),'#0.00'),FORMAT(SUM(VAL(����δ��)),'#0.00'),FORMAT(SUM(VAL(�����ۼ�δ��)),'#0.00'),FORMAT(SUM(VAL(δ����)),'#0.00') FROM JGZCX2 GROUP BY �ͻ�,���ڷ�Χ "
Data6.Database.Execute "DELETE *  FROM  JGZCX2 WHERE ���='1'"
Data6.Database.Execute "UPDATE JGZCX2 SET ����δ��=FORMAT(VAL(����Ӧ����)-VAL(���ڿ�Ʊ),'#0.00')"
Data6.Database.Execute "UPDATE JGZCX2 SET Ƿ��=FORMAT(VAL(�����ۼ�Ӧ��)+VAL(����Ӧ����)-VAL(�����Ѹ���),'#0.00'),�����ۼ�Ӧ����=FORMAT(VAL(�����ۼ�Ӧ��)+VAL(����Ӧ����),'#0.00'),�����ۼƿ�Ʊ=FORMAT(VAL(�����ۼƿ�Ʊ)+VAL(���ڿ�Ʊ),'#0.00'),�����ۼ�δ��=FORMAT(VAL(�����ۼ�δ��Ʊ)+VAL(����δ��),'#0.00')"
Data6.Database.Execute "DELETE *  FROM  JGZCX2 WHERE val(����Ӧ����)=0 and val(�����Ѹ���)=0 and val(Ƿ��)=0"
 
 
Data8.RecordSource = "select ��� from GYS WHERE INSTR(����,'R')>0"
Data8.Refresh

Data6.RecordSource = "SELECT �ͻ� FROM JGZCX2"
Data6.Refresh

If Not Data6.Recordset.EOF Then
Data6.Recordset.MoveFirst
Do While Not Data6.Recordset.EOF
Data8.Recordset.FindFirst "���='" & Data6.Recordset.Fields(0) & "'"
If Data8.Recordset.NoMatch Then
Data9.Database.Execute "DELETE *  FROM  JGZCX2 WHERE �ͻ�='" & Data6.Recordset.Fields(0) & "'"
End If
Data6.Recordset.MoveNext
Loop
End If
Command1.Enabled = True

Data6.RecordSource = "SELECT �ͻ�,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,�����ۼ�δ��,Ƿ��,δ����,���ڷ�Χ FROM JGZCX2  order by �ͻ�"
Data6.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call OutDataToExcel11(VSFlexGrid1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, "��ׯ��¡��֯�������޹�˾ �ͻ���Ŀ��ѯ--����" + "��ֹ����:" + Text2.Text)
End Sub

Private Sub Command4_Click()
Data9.Database.Execute "update JGZCX2 set �����ۼ�Ӧ��='' where �����ۼ�Ӧ��='0.00'"
Data9.Database.Execute "update JGZCX2 set ����Ӧ����='' where ����Ӧ����='0.00'"
Data9.Database.Execute "update JGZCX2 set �����ۼ�Ӧ����='' where �����ۼ�Ӧ����='0.00'"
Data9.Database.Execute "update JGZCX2 set �����Ѹ���='' where �����Ѹ���='0.00'"
Data9.Database.Execute "update JGZCX2 set �����ۼƿ�Ʊ='' where �����ۼƿ�Ʊ='0.00'"
Data9.Database.Execute "update JGZCX2 set ���ڿ�Ʊ='' where ���ڿ�Ʊ='0.00'"
Data9.Database.Execute "update JGZCX2 set �����ۼƿ�Ʊ='' where �����ۼƿ�Ʊ='0.00'"
Data9.Database.Execute "update JGZCX2 set �����ۼ�δ��Ʊ='' where �����ۼ�δ��Ʊ='0.00'"
Data9.Database.Execute "update JGZCX2 set �����ۼ�δ��='' where �����ۼ�δ��='0.00'"
Data9.Database.Execute "update JGZCX2 set Ƿ��='' where Ƿ��='0.00'"
Data9.Database.Execute "update JGZCX2 set �����ۼ�Ӧ��='' where �����ۼ�Ӧ��='0.00'"
Data6.Refresh
End Sub

Private Sub Command5_Click()
'On Error Resume Next
If MsgBox("ȷ����ת�������������Ϊ��" + Trim(DTPicker1.Value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("ȷ����ת������?", vbYesNo) = vbNo Then Exit Sub

lo = "d:\���ݿ�\bfrz\" + ljb + "\zcw.mdb"
Data6.RecordSource = "SELECT �ͻ�,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,�����ۼ�δ��,Ƿ��,���ڷ�Χ FROM JGZCX2  order by �ͻ�"
Data6.Refresh

If Not Data6.Recordset.EOF Then
Data6.Recordset.MoveFirst
Do While Not Data6.Recordset.EOF
Data10.Database.Execute "delete * from  PMMXJZ where instr(ժҪ,'Ⱦ��')>0 and ����='" & DTPicker1.Value & "' and instr(��ƿ�Ŀ,'Ӧ���˿�')>0 and mid(��ƿ�Ŀ,instr(��ƿ�Ŀ,'-')+1)='" & Data6.Recordset.Fields(0) & "'"
Data9.Database.Execute "INSERT INTO PMMXJZ(��ƿ�Ŀ,���) in'" & lo & "' SELECT 'Ӧ���˿�-'+trim(�ͻ�) as ll,Ƿ�� from JGZCX2 where �ͻ�='" & Data6.Recordset.Fields(0) & "'"
Data10.Database.Execute "update PMMXJZ set ժҪ='�ڳ����Ⱦ��',ƾ֤��='��',�������='��',���='1',����='" & DTPicker1.Value & "' where ����=null"

Data9.Database.Execute "delete * from  PMJHFP where  ��ת����='" & DTPicker1.Value & "' and �ͻ�='" & Data6.Recordset.Fields(0) & "'"
Data9.Database.Execute "insert into PMJHFP(�ͻ�,��Ʊ���,δ�����) select �ͻ�,�����ۼƿ�Ʊ,�����ۼ�δ�� from JGZCX2 where �ͻ�='" & Data6.Recordset.Fields(0) & "'"
Data9.Database.Execute "update PMJHFP set ��ת����='" & DTPicker1.Value & "' where ��ת����=null"
Data6.Recordset.MoveNext
Loop
End If



MsgBox ("��ת�ɹ��������ڳ������п��Բ�ѯ��")
End Sub

Private Sub Command6_Click()
Data10.Database.Execute "delete * from  PMMXJZ where ����='" & DTPicker1.Value & "' and instr(��ƿ�Ŀ,'Ӧ���˿�')>0"
Data9.Database.Execute "delete * from  PMJHFP where ��ת����='" & DTPicker1.Value & "'"
MsgBox ("����ɹ�!")
End Sub

Private Sub Command7_Click()
If MsgBox("��������Ϊ��" + Trim(DTPicker2.Value) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("�����ڼ�Ϊ��" + Trim(Month(DTPicker2.Value)) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("ȷ������Ӧ��ϵ�е�ƾ֤��", vbYesNo) = vbNo Then Exit Sub
Call CLRKPZ(CDate(Text1.Text), CDate(Text2.Text), CDate(DTPicker2.Value))
End Sub

Private Sub Command8_Click()
Form1132.DTPicker1.Value = DTPicker2.Value
Form1132.Show
Unload Me
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If DataCombo1.Text = "" Then
Data6.RecordSource = "SELECT �ͻ�,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,�����ۼ�δ��,Ƿ��,δ����,���ڷ�Χ FROM JGZCX2  order by �ͻ�"
Data6.Refresh
Else
Data6.RecordSource = "SELECT �ͻ�,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,�����ۼ�δ��,Ƿ��,δ����,���ڷ�Χ FROM JGZCX2 WHERE �ͻ�='" & DataCombo1.Text & "' and val(Ƿ��)<>0 order by �ͻ�"
Data6.Refresh
End If
End Sub

Private Sub DTPicker3_Change()
Text1.Text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = DTPicker3.Value
Text1.SetFocus
End Sub

Private Sub DTPicker4_Change()
Text2.Text = DTPicker4.Value
End Sub

Private Sub DTPicker4_CloseUp()
Text2.Text = DTPicker4.Value
Text2.SetFocus
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption + "������� " + ljb
Text1.Text = Date
Text2.Text = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
DataCombo1.Text = ""
Data1.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\SCZYJHD.mdb"
Data1.RecordSource = "select GYS.��� from GYS  GROUP BY ���"
Data1.Refresh
Data2.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CLCK.MDB"
Data2.RecordSource = "select ��Ӧ��λ,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,�ϼƽ��,���ݺ�,����,�Ƿ�Ʊ,��Ʊ,��Ʊ���� from ckgl where ��Ӧ��λ='" & DataCombo1.Text & "' and ���� between cdate('" & Text1 & "') and cdate('" & Text2.Text & "') AND ���='�ɹ����'"
Data2.Refresh
Data3.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CLCK.MDB"
Data4.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\ZCW.MDB"
Data5.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CW.MDB"
Data6.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\FP.MDB"
Data7.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CJBB.MDB"
Data7.RecordSource = "rqsd"
Data7.Refresh

Data8.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\SCZYJHD.mdb"
Data9.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\fp.MDB"
Data10.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\zcw.MDB"

For i = 2 To 12
VSFlexGrid1.ColWidth(i) = 1200
Next
VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(13) = 2200

End Sub

Private Sub vSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
S1 = VSFlexGrid1.RowSel
End Sub

Private Sub vSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
S2 = VSFlexGrid1.RowSel
End Sub

Private Sub CLRKPZ(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next
If InStr(ljb, "wx") > 0 Then
Data4.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "') and instr(�Ƶ�,'�Զ�-Ⱦ��')>0"
Data4.Refresh
If Not Data4.Recordset.EOF Then
If MsgBox("����Ӧ������ƾ֤���Ƿ��������ɣ�", vbYesNo) = vbNo Then Exit Sub
Data10.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(�Ƶ�,'�Զ�-Ⱦ��')>0 and ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data9.RecordSource = "SELECT * FROM JGZCX2 where val(����Ӧ����)<>0"
Data9.Refresh

If Data9.Recordset.EOF Then Exit Sub
Data4.RecordSource = "SELECT * FROM CLZZPZ"
Data4.Refresh

Data10.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
If Data10.Recordset.EOF Then
PZH = "R5-1"
Else
Data10.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,4))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
PZH = "R5-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 7
Data4.Recordset.AddNew
Data4.Recordset.Fields(0) = "��Ⱦ��"
Data4.Recordset.Fields(1) = "ԭ����"
Data4.Recordset.Fields(2) = ""
Data4.Recordset.Fields(3) = "Ӧ���˿�"
Data4.Recordset.Fields(4) = Data9.Recordset.Fields(0)
Data4.Recordset.Fields(5) = Format(Data9.Recordset.Fields(2), "#0.00")
Data4.Recordset.Fields(6) = PZH
Data4.Recordset.Fields(7) = CDate(dt3)
Data4.Recordset.Fields(8) = ""
Data4.Recordset.Fields(9) = ""
Data4.Recordset.Fields(10) = ""
Data4.Recordset.Fields(11) = "�Զ�-Ⱦ��"
Data4.Recordset.Update


Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
Exit Sub
End If
Next
Data10.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
If Data10.Recordset.EOF Then
PZH = "R5-1"
Else
Data10.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,4))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
PZH = "R5-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
End If


If InStr(ljb, "nx") > 0 Then
Data4.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "') and instr(�Ƶ�,'�Զ�-Ⱦ��')>0"
Data4.Refresh
If Not Data4.Recordset.EOF Then
If MsgBox("����Ӧ������ƾ֤���Ƿ��������ɣ�", vbYesNo) = vbNo Then Exit Sub
Data10.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(�Ƶ�,'�Զ�-Ⱦ��')>0 and ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data9.RecordSource = "SELECT * FROM JGZCX2 where val(����Ӧ����)>0"
Data9.Refresh

If Data9.Recordset.EOF Then Exit Sub
Data4.RecordSource = "SELECT * FROM CLZZPZ"
Data4.Refresh

Data10.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
If Data10.Recordset.EOF Then
PZH = "I5-1"
Else
Data10.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,4))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
PZH = "I5-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 7
Data4.Recordset.AddNew
Data4.Recordset.Fields(0) = "��Ⱦ��"
Data4.Recordset.Fields(1) = "ԭ����"
Data4.Recordset.Fields(2) = ""
Data4.Recordset.Fields(3) = "Ӧ���˿�"
Data4.Recordset.Fields(4) = Data9.Recordset.Fields(0)
Data4.Recordset.Fields(5) = Format(Data9.Recordset.Fields(2), "#0.00")
Data4.Recordset.Fields(6) = PZH
Data4.Recordset.Fields(7) = CDate(dt3)
Data4.Recordset.Fields(8) = ""
Data4.Recordset.Fields(9) = ""
Data4.Recordset.Fields(10) = ""
Data4.Recordset.Fields(11) = "�Զ�-Ⱦ��"
Data4.Recordset.Update


Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
Exit Sub
End If
Next
Data10.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
If Data10.Recordset.EOF Then
PZH = "I5-1"
Else
Data10.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,4))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
PZH = "I5-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
End If


End Sub





