VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Formw73 
   BackColor       =   &H00C0E0FF&
   Caption         =   "���Ͽͻ��˲�ѯ---����"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ƾ֤����"
      Height          =   855
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ɲ�ѯ"
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   960
      Width           =   1335
   End
   Begin VB.Data Data11 
      Caption         =   "Data5"
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data10 
      Caption         =   "Data5"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ׼��"
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ת����"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Width           =   3855
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3975
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   3615
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   3735
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   3735
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1215
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
      Top             =   10920
      Visible         =   0   'False
      Width           =   3495
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
      Top             =   10920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw73.frx":0000
      Height          =   7455
      Left            =   240
      TabIndex        =   5
      Top             =   2160
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
      Bindings        =   "Formw73.frx":0014
      Height          =   330
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
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
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   82313217
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   82313217
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   82313217
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   12000
      TabIndex        =   19
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   82313217
      CurrentDate     =   36892
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   2
      Left            =   12000
      TabIndex        =   20
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "�������"
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   15
      Top             =   1080
      Width           =   1455
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
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   960
      Width           =   855
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
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "Formw73"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer

Private Sub Command1_Click()
'On Error Resume Next
Command1.Enabled = False
Data6.Database.Execute "DELETE * FROM JGZCX1"
lo = "d:\���ݿ�\bfrz\" + ljb + "\FP.MDB"       '''''''''''''''''''''''����
''''   ����Data4.Database.Execute "insert into JGZCX1(�ͻ�,�����ۼ�Ӧ��) IN'" & LO &"' SELECT MID(��ƿ�Ŀ,INSTR(��ƿ�Ŀ,'-')+1),format(SUM(VAL(���)),'#0.00') FROM PMMXJZ WHERE �������='��' AND ����=CDATE('" & RQQ & "') GROUP BY MID(��ƿ�Ŀ,INSTR(��ƿ�Ŀ,'-')+1)"
Data4.Database.Execute "insert into JGZCX1(�ͻ�,�����ۼ�Ӧ��) IN'" & lo & "' SELECT MID(��ƿ�Ŀ,INSTR(��ƿ�Ŀ,'-')+1),format(SUM(VAL(���)),'#0.00') FROM PMMXJZ WHERE �������='��' AND ����=CDATE('" & Text1.Text & "') GROUP BY MID(��ƿ�Ŀ,INSTR(��ƿ�Ŀ,'-')+1)"
Data3.Database.Execute "insert into JGZCX1(�ͻ�,����Ӧ����) in'" & lo & "' SELECT ��Ӧ��λ,format(SUM(�ϼƽ��),'#0.00') FROM CKGL WHERE  ���� between cdate('" & Text1 & "') and cdate('" & Text2.Text & "') AND ���='�ɹ����' and �Ƿ񸶿�<>'�Ѹ�' GROUP BY ��Ӧ��λ"
Data10.Database.Execute "insert into JGZCX1(�ͻ�,����Ӧ����) in'" & lo & "' SELECT ��Ӧ��λ,format(SUM(�ϼƽ��),'#0.00') FROM MX WHERE  ���ʱ�� between cdate('" & Text1 & "') and cdate('" & Text2.Text & "') AND ���='�ɹ����' GROUP BY ��Ӧ��λ"
Data10.Database.Execute "insert into JGZCX1(�ͻ�,����Ӧ����) in'" & lo & "' SELECT ���ⵥλ,format(SUM(-�ϼƽ��),'#0.00') FROM ckMX WHERE  ����ʱ�� between cdate('" & Text1 & "') and cdate('" & Text2.Text & "')  GROUP BY ���ⵥλ"
'Data5.Database.Execute "insert into JGZCX1(�ͻ�,����Ӧ����) in'" & LO & "' SELECT ��Ӧ��λ,format(SUM(�ϼƽ��),'#0.00') FROM MX WHERE  ���ʱ�� between cdate('" & Text1 & "') and cdate('" & Text2.text & "') AND ���='�ɹ����' GROUP BY ��Ӧ��λ"
rqq = CDate(Text2.Text) + 1
Data6.Database.Execute "insert into JGZCX1(�ͻ�,���ڿ�Ʊ)  SELECT �ͻ�,��Ʊ��� FROM JHFP WHERE ��Ʊ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & rqq & "')"
Data4.Database.Execute "insert into JGZCX1(�ͻ�,�����Ѹ���) IN'" & lo & "' SELECT MID(�Է���Ŀ,INSTR(�Է���Ŀ,'-')+1),format(SUM(VAL(�������)),'#0.00') FROM TZJZMX WHERE ���� between cdate('" & Text1.Text & "') and cdate('" & Text2.Text & "') AND �������<>'0' GROUP BY MID(�Է���Ŀ,INSTR(�Է���Ŀ,'-')+1)"
Data6.Database.Execute "insert into JGZCX1(�ͻ�,�����ۼƿ�Ʊ) SELECT �ͻ�,��Ʊ��� FROM PMJHFP WHERE  ��ת����=CDATE('" & Text1.Text & "')"
Data6.Database.Execute "insert into JGZCX1(�ͻ�,�����ۼ�δ��Ʊ) SELECT �ͻ�,δ����� FROM PMJHFP WHERE  ��ת����=CDATE('" & Text1.Text & "')"

Data4.Database.Execute "insert into JGZCX1(�ͻ�,δ����) IN'" & lo & "' SELECT MID(�Է���Ŀ,INSTR(�Է���Ŀ,'-')+1),format(SUM(VAL(�������)),'#0.00') FROM TZJZMX WHERE ���� between cdate('" & Text1.Text & "') and cdate('" & Text2.Text & "') AND �������<>'0' GROUP BY MID(�Է���Ŀ,INSTR(�Է���Ŀ,'-')+1)"
Data4.Database.Execute "insert into JGZCX1(�ͻ�,δ����) IN'" & lo & "' SELECT �ͻ�,format(SUM(VAL(���)),'#0.00') FROM WDZSZ WHERE ����=cdate('" & Text1.Text & "')  GROUP BY �ͻ�"
Data6.Database.Execute "insert into JGZCX1(�ͻ�,δ����) SELECT �ͻ�,format(SUM(VAL(��Ʊ���)),'#0.00') FROM JHFP WHERE  ��Ʊ���� between cdate('" & Text1.Text & "') and cdate('" & rqq & "') GROUP BY �ͻ�"


Data6.Database.Execute "UPDATE JGZCX1 SET ���='1'"
Data6.Database.Execute "UPDATE JGZCX1 SET ���ڷ�Χ='" & Text1.Text & "'+'--'+'" & Text2.Text & "'"
Data6.Database.Execute "UPDATE JGZCX1 SET �����ۼ�Ӧ��='0' WHERE �����ۼ�Ӧ��=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET ����Ӧ����='0' WHERE ����Ӧ����=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET �����ۼ�Ӧ����='0' WHERE �����ۼ�Ӧ����=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET �����Ѹ���='0' WHERE �����Ѹ���=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET �����ۼƿ�Ʊ='0' WHERE �����ۼƿ�Ʊ=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET ���ڿ�Ʊ='0' WHERE ���ڿ�Ʊ=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET �����ۼ�δ��Ʊ='0' WHERE �����ۼ�δ��Ʊ=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET �����ۼƿ�Ʊ='0' WHERE �����ۼƿ�Ʊ=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET ����δ��='0' WHERE ����δ��=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET �����ۼ�δ��='0' WHERE �����ۼ�δ��=NULL"
Data6.Database.Execute "UPDATE JGZCX1 SET δ����='0' WHERE δ����=NULL"


Data6.Database.Execute "insert into JGZCX1(�ͻ�,���ڷ�Χ,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,����δ��,�����ۼ�δ��,δ����) SELECT �ͻ�,���ڷ�Χ,FORMAT(SUM(VAL(�����ۼ�Ӧ��)),'#0.00'),FORMAT(SUM(VAL(����Ӧ����)),'#0.00'),FORMAT(SUM(VAL(�����ۼ�Ӧ����)),'#0.00'),FORMAT(SUM(VAL(�����Ѹ���)),'#0.00'),FORMAT(SUM(VAL(�����ۼƿ�Ʊ)),'#0.00'),FORMAT(SUM(VAL(���ڿ�Ʊ)),'#0.00'),FORMAT(SUM(VAL(�����ۼƿ�Ʊ)),'#0.00'),FORMAT(SUM(VAL(�����ۼ�δ��Ʊ)),'#0.00'),FORMAT(SUM(VAL(����δ��)),'#0.00'),FORMAT(SUM(VAL(�����ۼ�δ��)),'#0.00'),FORMAT(SUM(VAL(δ����)),'#0.00') FROM JGZCX1 GROUP BY �ͻ�,���ڷ�Χ "
Data6.Database.Execute "DELETE *  FROM  JGZCX1 WHERE ���='1'"
Data6.Database.Execute "UPDATE JGZCX1 SET ����δ��=FORMAT(VAL(����Ӧ����)-VAL(���ڿ�Ʊ),'#0.00')"
Data6.Database.Execute "UPDATE JGZCX1 SET Ƿ��=FORMAT(VAL(�����ۼ�Ӧ��)+VAL(����Ӧ����)-VAL(�����Ѹ���),'#0.00'),�����ۼ�Ӧ����=FORMAT(VAL(�����ۼ�Ӧ��)+VAL(����Ӧ����),'#0.00'),�����ۼƿ�Ʊ=FORMAT(VAL(�����ۼƿ�Ʊ)+VAL(���ڿ�Ʊ),'#0.00'),�����ۼ�δ��=FORMAT(VAL(�����ۼ�δ��Ʊ)+VAL(����δ��),'#0.00')"
Data6.Database.Execute "DELETE *  FROM  JGZCX1 WHERE val(����Ӧ����)=0 and val(�����Ѹ���)=0 and val(Ƿ��)=0"

Data8.RecordSource = "select ��� from GYS WHERE INSTR(����,'P')>0"
Data8.Refresh
Data6.RecordSource = "SELECT �ͻ� FROM JGZCX1"
Data6.Refresh

If Not Data6.Recordset.EOF Then
Data6.Recordset.MoveFirst
Do While Not Data6.Recordset.EOF
Data8.Recordset.FindFirst "���='" & Data6.Recordset.Fields(0) & "'"
If Data8.Recordset.NoMatch Then
Data9.Database.Execute "DELETE *  FROM  JGZCX1 WHERE �ͻ�='" & Data6.Recordset.Fields(0) & "'"
End If
Data6.Recordset.MoveNext
Loop
End If
Command1.Enabled = True

Data6.RecordSource = "SELECT �ͻ�,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,�����ۼ�δ��,Ƿ��,���ڷ�Χ,δ���� FROM JGZCX1  order by �ͻ�"
Data6.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call OutDataToExcel11(VSFlexGrid1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, "��ׯ��¡��֯�������޹�˾ �ͻ���Ŀ��ѯ--����" + "��ֹ����:" + Text2.Text)
End Sub

Private Sub Command4_Click()
Data9.Database.Execute "update JGZCX1 set �����ۼ�Ӧ��='' where �����ۼ�Ӧ��='0.00'"
Data9.Database.Execute "update JGZCX1 set ����Ӧ����='' where ����Ӧ����='0.00'"
Data9.Database.Execute "update JGZCX1 set �����ۼ�Ӧ����='' where �����ۼ�Ӧ����='0.00'"
Data9.Database.Execute "update JGZCX1 set �����Ѹ���='' where �����Ѹ���='0.00'"
Data9.Database.Execute "update JGZCX1 set �����ۼƿ�Ʊ='' where �����ۼƿ�Ʊ='0.00'"
Data9.Database.Execute "update JGZCX1 set ���ڿ�Ʊ='' where ���ڿ�Ʊ='0.00'"
Data9.Database.Execute "update JGZCX1 set �����ۼƿ�Ʊ='' where �����ۼƿ�Ʊ='0.00'"
Data9.Database.Execute "update JGZCX1 set �����ۼ�δ��Ʊ='' where �����ۼ�δ��Ʊ='0.00'"
Data9.Database.Execute "update JGZCX1 set �����ۼ�δ��='' where �����ۼ�δ��='0.00'"
Data9.Database.Execute "update JGZCX1 set Ƿ��='' where Ƿ��='0.00'"
Data9.Database.Execute "update JGZCX1 set �����ۼ�Ӧ��='' where �����ۼ�Ӧ��='0.00'"
Data6.Refresh
End Sub

Private Sub Command5_Click()
If MsgBox("ȷ����ת�������������Ϊ��" + Trim(DTPicker1.Value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("ȷ����ת������?", vbYesNo) = vbNo Then Exit Sub

lo = "d:\���ݿ�\bfrz\" + ljb + "\zcw.mdb"

Data6.RecordSource = "SELECT �ͻ�,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,�����ۼ�δ��,Ƿ��,���ڷ�Χ FROM JGZCX1  order by �ͻ�"
Data6.Refresh

If Not Data6.Recordset.EOF Then
Data6.Recordset.MoveFirst
Do While Not Data6.Recordset.EOF
Data4.Database.Execute "delete * from  PMMXJZ where instr(ժҪ,'����')>0 and ����='" & DTPicker1.Value & "' and instr(��ƿ�Ŀ,'Ӧ���˿�')>0 and mid(��ƿ�Ŀ,instr(��ƿ�Ŀ,'-')+1)='" & Data6.Recordset.Fields(0) & "'"
Data5.Database.Execute "INSERT INTO PMMXJZ(��ƿ�Ŀ,���) in'" & lo & "' SELECT 'Ӧ���˿�-'+trim(�ͻ�) as ll,Ƿ�� from JGZCX1 where �ͻ�='" & Data6.Recordset.Fields(0) & "'"
Data4.Database.Execute "update PMMXJZ set ժҪ='�ڳ�������',ƾ֤��='��',�������='��',���='1',����='" & DTPicker1.Value & "' where ����=null"

Data5.Database.Execute "delete * from  PMJHFP where ��ת����='" & DTPicker1.Value & "' and �ͻ�='" & Data6.Recordset.Fields(0) & "'"
Data5.Database.Execute "insert into PMJHFP(�ͻ�,��Ʊ���,δ�����) select �ͻ�,�����ۼƿ�Ʊ,�����ۼ�δ�� from JGZCX1 where �ͻ�='" & Data6.Recordset.Fields(0) & "'"
Data5.Database.Execute "update PMJHFP set ��ת����='" & DTPicker1.Value & "' where ��ת����=null"
Data6.Recordset.MoveNext
Loop
End If

MsgBox ("��ת�ɹ��������ڳ������п��Բ�ѯ��")
End Sub

Private Sub Command6_Click()
Form1132.DTPicker1.Value = DTPicker2.Value
Form1132.Show
Unload Me
End Sub

Private Sub Command7_Click()
If MsgBox("��������Ϊ��" + Trim(DTPicker2.Value) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("�����ڼ�Ϊ��" + Trim(Month(DTPicker2.Value)) + "��ȷ��", vbYesNo) = vbNo Then Exit Sub
If MsgBox("ȷ������Ӧ��ϵ�е�ƾ֤��", vbYesNo) = vbNo Then Exit Sub
Call CLRKPZ(CDate(Text1.Text), CDate(Text2.Text), CDate(DTPicker2.Value))
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If DataCombo1.Text = "" Then
Data6.RecordSource = "SELECT �ͻ�,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,�����ۼ�δ��,Ƿ��,δ����,���ڷ�Χ FROM JGZCX1  order by �ͻ�"
Data6.Refresh
Else
Data6.RecordSource = "SELECT �ͻ�,�����ۼ�Ӧ��,����Ӧ����,�����ۼ�Ӧ����,�����Ѹ���,�����ۼƿ�Ʊ,���ڿ�Ʊ,�����ۼƿ�Ʊ,�����ۼ�δ��Ʊ,�����ۼ�δ��,Ƿ��,δ����,���ڷ�Χ FROM JGZCX1 WHERE �ͻ�='" & DataCombo1.Text & "' and  val(Ƿ��)<>0 order by �ͻ�"
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
Data5.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\FP.MDB"
Data6.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\FP.MDB"
Data7.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CJBB.MDB"
Data7.RecordSource = "rqsd"
Data7.Refresh
Data10.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CW.MDB"

Data8.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\SCZYJHD.mdb"
Data9.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\FP.mdb"
Data11.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\ZCW.MDB"

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
Data4.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "') and instr(�Ƶ�,'�Զ�-����')>0"
Data4.Refresh
If Not Data4.Recordset.EOF Then
If MsgBox("����Ӧ������ƾ֤���Ƿ��������ɣ�", vbYesNo) = vbNo Then Exit Sub
Data11.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(�Ƶ�,'�Զ�-����')>0 and ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data5.RecordSource = "SELECT * FROM JGZCX1 where val(����Ӧ����)<>0"
Data5.Refresh

If Data5.Recordset.EOF Then Exit Sub
Data4.RecordSource = "SELECT * FROM CLZZPZ"
Data4.Refresh

Data11.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
If Data11.Recordset.EOF Then
PZH = "R5-1"
Else
Data11.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,4))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
PZH = "R5-" + Trim(Data11.Recordset.Fields(0) + 1)
End If
Data5.Recordset.MoveFirst
KLLLL = 1
Do While Not Data5.Recordset.EOF
For i = 1 To 7
Data4.Recordset.AddNew
Data4.Recordset.Fields(0) = "������"
Data4.Recordset.Fields(1) = "ԭ����"
Data4.Recordset.Fields(2) = ""
Data4.Recordset.Fields(3) = "Ӧ���˿�"
Data4.Recordset.Fields(4) = Data5.Recordset.Fields(0)
Data4.Recordset.Fields(5) = Format(Data5.Recordset.Fields(2), "#0.00")
Data4.Recordset.Fields(6) = PZH
Data4.Recordset.Fields(7) = CDate(dt3)
Data4.Recordset.Fields(8) = ""
Data4.Recordset.Fields(9) = ""
Data4.Recordset.Fields(10) = ""
Data4.Recordset.Fields(11) = "�Զ�-����"
Data4.Recordset.Update


'Data4.Recordset.AddNew
'Data4.Recordset.Fields(0) = "������"
'Data4.Recordset.Fields(1) = "Ӧ��˰��"
'Data4.Recordset.Fields(2) = "˰�����"
'Data4.Recordset.Fields(3) = "Ӧ���˿�"
'Data4.Recordset.Fields(4) = Data5.Recordset.Fields(0)
'Data4.Recordset.Fields(5) = Format(Data5.Recordset.Fields(2) * 0.17, "#0.00")
'Data4.Recordset.Fields(6) = PZH
'Data4.Recordset.Fields(7) = CDate(dt3)
'Data4.Recordset.Fields(8) = ""
'Data4.Recordset.Fields(9) = ""
'Data4.Recordset.Fields(10) = ""
'Data4.Recordset.Fields(11) = "�Զ�-����"
'Data4.Recordset.Update


Data5.Recordset.MoveNext
If Data5.Recordset.EOF Then
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
Exit Sub
End If
Next
Data11.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
If Data11.Recordset.EOF Then
PZH = "R5-1"
Else
Data11.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,4))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
PZH = "R5-" + Trim(Data11.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
End If


If InStr(ljb, "nx") > 0 Then
Data4.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "') and instr(�Ƶ�,'�Զ�-����')>0"
Data4.Refresh
If Not Data4.Recordset.EOF Then
If MsgBox("����Ӧ������ƾ֤���Ƿ��������ɣ�", vbYesNo) = vbNo Then Exit Sub
Data11.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(�Ƶ�,'�Զ�-����')>0 and ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data5.RecordSource = "SELECT * FROM JGZCX1 where val(����Ӧ����)>0"
Data5.Refresh

If Data5.Recordset.EOF Then Exit Sub
Data4.RecordSource = "SELECT * FROM CLZZPZ"
Data4.Refresh

Data11.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
If Data11.Recordset.EOF Then
PZH = "I5-1"
Else
Data11.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,4))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
PZH = "I5-" + Trim(Data11.Recordset.Fields(0) + 1)
End If
Data5.Recordset.MoveFirst
KLLLL = 1
Do While Not Data5.Recordset.EOF
For i = 1 To 7
Data4.Recordset.AddNew
Data4.Recordset.Fields(0) = "������"
Data4.Recordset.Fields(1) = "ԭ����"
Data4.Recordset.Fields(2) = ""
Data4.Recordset.Fields(3) = "Ӧ���˿�"
Data4.Recordset.Fields(4) = Data5.Recordset.Fields(0)
Data4.Recordset.Fields(5) = Format(Data5.Recordset.Fields(2), "#0.00")
Data4.Recordset.Fields(6) = PZH
Data4.Recordset.Fields(7) = CDate(dt3)
Data4.Recordset.Fields(8) = ""
Data4.Recordset.Fields(9) = ""
Data4.Recordset.Fields(10) = ""
Data4.Recordset.Fields(11) = "�Զ�-����"
Data4.Recordset.Update


'Data4.Recordset.AddNew
'Data4.Recordset.Fields(0) = "������"
'Data4.Recordset.Fields(1) = "Ӧ��˰��"
'Data4.Recordset.Fields(2) = "˰�����"
'Data4.Recordset.Fields(3) = "Ӧ���˿�"
'Data4.Recordset.Fields(4) = Data5.Recordset.Fields(0)
'Data4.Recordset.Fields(5) = Format(Data5.Recordset.Fields(2) * 0.17, "#0.00")
'Data4.Recordset.Fields(6) = PZH
'Data4.Recordset.Fields(7) = CDate(dt3)
'Data4.Recordset.Fields(8) = ""
'Data4.Recordset.Fields(9) = ""
'Data4.Recordset.Fields(10) = ""
'Data4.Recordset.Fields(11) = "�Զ�-����"
'Data4.Recordset.Update


Data5.Recordset.MoveNext
If Data5.Recordset.EOF Then
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
Exit Sub
End If
Next
Data11.RecordSource = "SELECT * FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
If Data11.Recordset.EOF Then
PZH = "I5-1"
Else
Data11.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,4))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
PZH = "I5-" + Trim(Data11.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
End If


End Sub


