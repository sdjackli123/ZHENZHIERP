VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Formw337 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�˱����"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form37"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data13 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data14 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
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
      Height          =   1095
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
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
      ItemData        =   "Formw337.frx":0000
      Left            =   4560
      List            =   "Formw337.frx":0010
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ƾ֤"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ܷ�����"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "������"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����ռ���"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ϸ��"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����ڡ�ƾ֤"
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
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw337.frx":003C
      Height          =   7815
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13785
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   8421631
      BackColorBkg    =   34952
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   83755009
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   83755009
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   83755009
      CurrentDate     =   39883
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�˱���ѯ"
      Height          =   1215
      Left            =   6960
      TabIndex        =   19
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�������"
      Height          =   1215
      Left            =   10320
      TabIndex        =   20
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label5 
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
      Index           =   2
      Left            =   2400
      TabIndex        =   18
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Index           =   1
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ƾ֤���"
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
      Left            =   4560
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����·�"
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
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Formw337"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2 As String

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("��������")
Exit Sub
End If
If Combo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY PZDZ.����,PZDZ.ƾ֤��"
Data1.Refresh
Else
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.ƾ֤���='" & Combo1.Text & "' AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY PZDZ.����,PZDZ.ƾ֤��"
Data1.Refresh
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("����������ڷ�Χ")
Exit Sub
End If
Data1.Database.Execute "INSERT INTO MXFLZ(����,ƾ֤��,ժҪ,��ƿ�Ŀ,�跽���,�������) SELECT PZDZ.����,PZDZ.ƾ֤��,PZDZ.ժҪ,PZDZ.��ƿ�Ŀ,PZDZ.�跽���,PZDZ.������� FROM PZDZ WHERE (PZDZ.��ϸ����='' OR PZDZ.��ϸ����=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.��ƿ�Ŀ,'-')>0"
Data1.Database.Execute "update  PZDZ SET ��ϸ����='��' WHERE (PZDZ.��ϸ����='' OR PZDZ.��ϸ����=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
MsgBox ("��ϸ����˳ɹ�")
Data1.Refresh
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("����������ڷ�Χ")
Exit Sub
End If
Data1.Database.Execute "DELETE * FROM TZJZ WHERE ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"

Data1.Database.Execute "INSERT INTO TZJZ(����,ƾ֤��,ժҪ,�Է���Ŀ,�跽���,�������) SELECT PZDZ.����,PZDZ.ƾ֤��,PZDZ.ժҪ,PZDZ.��ƿ�Ŀ,PZDZ.�跽���,PZDZ.������� FROM PZDZ WHERE (PZDZ.��������='' OR PZDZ.��������=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.ƾ֤��,'1-')>0 AND ��ƿ�Ŀ<>'�ֽ�'"
Data1.Database.Execute "update  TZJZ SET ���='�ֽ�' WHERE ���=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(����,ƾ֤��,ժҪ,�Է���Ŀ,�跽���,�������) SELECT PZDZ.����,PZDZ.ƾ֤��,PZDZ.ժҪ,PZDZ.��ƿ�Ŀ,PZDZ.�跽���,PZDZ.������� FROM PZDZ WHERE (PZDZ.��������='' OR PZDZ.��������=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.ƾ֤��,'2-')>0  AND ��ƿ�Ŀ<>'�ֽ�'"
Data1.Database.Execute "update  TZJZ SET ���='�ֽ�' WHERE ���=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(����,ƾ֤��,ժҪ,�Է���Ŀ,�������,�跽���) SELECT PZDZ.����,PZDZ.ƾ֤��,PZDZ.ժҪ,PZDZ.��ƿ�Ŀ,PZDZ.�跽���,PZDZ.������� FROM PZDZ WHERE (PZDZ.��������='' OR PZDZ.��������=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.ƾ֤��,'2-')>0  AND ��ƿ�Ŀ='���д��'"
Data1.Database.Execute "update  TZJZ SET ���='���д��' WHERE ���=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(����,ƾ֤��,ժҪ,�Է���Ŀ,�跽���,�������) SELECT PZDZ.����,PZDZ.ƾ֤��,PZDZ.ժҪ,PZDZ.��ƿ�Ŀ,PZDZ.�跽���,PZDZ.������� FROM PZDZ WHERE (PZDZ.��������='' OR PZDZ.��������=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.ƾ֤��,'3-')>0 AND ��ƿ�Ŀ<>'���д��'"
Data1.Database.Execute "update  TZJZ SET ���='���д��' WHERE ���=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(����,ƾ֤��,ժҪ,�Է���Ŀ,�跽���,�������) SELECT PZDZ.����,PZDZ.ƾ֤��,PZDZ.ժҪ,PZDZ.��ƿ�Ŀ,PZDZ.�跽���,PZDZ.������� FROM PZDZ WHERE (PZDZ.��������='' OR PZDZ.��������=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.ƾ֤��,'4-')>0 AND ��ƿ�Ŀ<>'���д��'"
Data1.Database.Execute "update  TZJZ SET ���='���д��' WHERE ���=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(����,ƾ֤��,ժҪ,�Է���Ŀ,�������,�跽���) SELECT PZDZ.����,PZDZ.ƾ֤��,PZDZ.ժҪ,PZDZ.��ƿ�Ŀ,PZDZ.�跽���,PZDZ.������� FROM PZDZ WHERE (PZDZ.��������='' OR PZDZ.��������=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.ƾ֤��,'4-')>0 AND ��ƿ�Ŀ='�ֽ�'"
Data1.Database.Execute "update  TZJZ SET ���='�ֽ�' WHERE ���=NULL"

Data1.Database.Execute "update  PZDZ SET ��������='��' WHERE (PZDZ.��������='' OR PZDZ.��������=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"



MsgBox ("�ռ��˳ɹ�")
Data1.Refresh
End Sub

Private Sub Command4_Click()
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY PZDZ.����,PZDZ.ƾ֤��"
Data1.Refresh
End Sub

Private Sub Command5_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("����������ڷ�Χ")
Exit Sub
End If
Data1.Database.Execute "INSERT INTO ZFLZ(����,ƾ֤��,ժҪ,��ƿ�Ŀ,�跽���,�������) SELECT PZDZ.����,PZDZ.ƾ֤��,PZDZ.ժҪ,PZDZ.��ƿ�Ŀ,PZDZ.�跽���,PZDZ.������� FROM PZDZ WHERE (PZDZ.�ܷ�����='' OR PZDZ.�ܷ�����=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data1.Database.Execute "update  ZFLZ SET ��ƿ�Ŀ=LEFT(ZFLZ.��ƿ�Ŀ,INSTR(ZFLZ.��ƿ�Ŀ,'-')-1) WHERE INSTR(ZFLZ.��ƿ�Ŀ,'-')>0"
Data1.Database.Execute "update  PZDZ SET �ܷ�����='��' WHERE (PZDZ.�ܷ�����='' OR PZDZ.�ܷ�����=NULL) AND PZDZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
MsgBox ("��������˳ɹ�")
Data1.Refresh
End Sub

Private Sub Command6_Click()
If Combo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.���� BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') ORDER BY PZDZ.����,PZDZ.ƾ֤��"
Data1.Refresh
Else
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.ƾ֤���='" & Combo1.Text & "' AND PZDZ.���� BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') ORDER BY PZDZ.����,PZDZ.ƾ֤��"
Data1.Refresh
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Text1.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text1.Text = DTPicker1.Value
End Sub
Private Sub DTPicker2_Change()
Text2.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text2.Text = DTPicker2.Value
End Sub


Private Sub DTPicker3_Change()
Data13.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CJBB.mdb"
Data13.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between ��ʼ���� and ��������"
Data13.Refresh
If Data13.Recordset.EOF Then
MsgBox ("�ڼ�����")
Else
K1 = Data13.Recordset.Fields(0)
K2 = Data13.Recordset.Fields(1)
Text3.Text = Data13.Recordset.Fields(2)
End If
Text1.Text = K1
Text2.Text = K2
DTPicker1.Value = K1
DTPicker2.Value = K2
End Sub

Private Sub DTPicker3_CloseUp()
Data13.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CJBB.mdb"
Data13.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between ��ʼ���� and ��������"
Data13.Refresh
If Data13.Recordset.EOF Then
MsgBox ("�ڼ�����")
Else
K1 = Data13.Recordset.Fields(0)
K2 = Data13.Recordset.Fields(1)
Text3.Text = Data13.Recordset.Fields(2)
End If
Text1.Text = K1
Text2.Text = K2
DTPicker1.Value = K1
DTPicker2.Value = K2
End Sub

Private Sub Form_Load()
'On Error Resume Next
Text1.Text = Date
DTPicker3.Value = Date
DTPicker1.Value = Date
Text2.Text = Date
DTPicker2.Value = Date

Data13.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CJBB.mdb"
Data13.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between ��ʼ���� and ��������"
Data13.Refresh
If Data13.Recordset.EOF Then
MsgBox ("�ڼ�����")
Else
K1 = Data13.Recordset.Fields(0)
K2 = Data13.Recordset.Fields(1)
Text3.Text = Data13.Recordset.Fields(2)
End If
Text1.Text = K1
Text2.Text = K2
DTPicker1.Value = K1
DTPicker2.Value = K2

Combo1.Text = ""
Data1.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CW.MDB"
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.���� BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') ORDER BY PZDZ.����,PZDZ.ƾ֤��"
Data1.Refresh

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 2500
MSFlexGrid1.ColWidth(7) = 700
MSFlexGrid1.ColWidth(8) = 700
MSFlexGrid1.ColWidth(9) = 700
End Sub

