VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formb16 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�������"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form16"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��żƻ�����"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   720
      Width           =   1575
   End
   Begin VB.Data Data12 
      Caption         =   "Data4"
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
      Width           =   3855
   End
   Begin VB.Data Data11 
      Caption         =   "Data4"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "������������"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   390
      Left            =   8520
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���żƻ�����"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ѯ"
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Width           =   3855
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ϵ��ˢ��"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ˢ��"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "FormB16.frx":0000
      Height          =   8055
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   14208
      _Version        =   393216
      Cols            =   12
      BackColorFixed  =   8421631
      BackColorBkg    =   4109501
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   36892
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   12720
      TabIndex        =   9
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Left            =   3120
      TabIndex        =   15
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "FormB16.frx":0014
      Height          =   6375
      Left            =   11400
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   16777215
      FocusRect       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "FormB16.frx":0029
      Height          =   1575
      Left            =   11400
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   16777215
      FocusRect       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Left            =   5160
      TabIndex        =   20
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ѡ���ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5160
      TabIndex        =   21
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ѡ�񵥺ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   16
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   11400
      TabIndex        =   10
      Top             =   360
      Width           =   1335
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
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1335
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
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Formb16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer: Public YPDH As String
Private Sub Combo1_Change()
End Sub

Private Sub Combo1_Click()
End Sub

Private Sub Command1_Click()
On Error Resume Next
If MsgBox("ȷ��ˢ����", vbYesNo) = vbNo Then Exit Sub

Data3.RecordSource = "SELECT * FROM CLB WHERE ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
Data3.Refresh
If Data3.Recordset.EOF Then Exit Sub
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
Data9.Recordset.FindFirst "������='" & Data3.Recordset.Fields(11) & "' and ������='" & Data3.Recordset.Fields(2) & "'"
Data3.Recordset.Edit
If Data9.Recordset.NoMatch Then
Data3.Recordset.Fields(7) = 0
Else
Data3.Recordset.Fields(7) = Data9.Recordset.Fields(2)
End If
Data3.Recordset.Update
Data3.Recordset.MoveNext
Loop
MsgBox ("ϵ����ˢ��")
End Sub


Private Sub Command10_Click()
If DBCombo2.Text = "" Then
MsgBox ("�����뵥��")
Exit Sub
End If

If DBCombo3.Text = "" Then
MsgBox ("��������")
Exit Sub
End If

Data10.RecordSource = "SELECT ���,��ɫ,����,����,���� FROM cmb WHERE ����='" & DBCombo2.Text & "' and ���='" & DBCombo3.Text & "'"
Data10.Refresh
If Data10.Recordset.EOF Then Exit Sub
lo = "d:\���ݿ�\\htgl\2011\DB.MDB"
Data7.Database.Execute "delete * from clbsc"
Data10.Recordset.MoveFirst
Do While Not Data10.Recordset.EOF
Data7.Database.Execute "UPDATE CLB SET ����='" & Data10.Recordset.Fields(3) & "' WHERE  ��ʽ='" & Data10.Recordset.Fields(0) & "' AND ��ɫ='" & Data10.Recordset.Fields(1) & "' and ����='" & Data10.Recordset.Fields(2) & "' and ����='" & Data10.Recordset.Fields(4) & "'"
Data7.Database.Execute "insert into CLBSC(��ǩ,�ͻ�����,Ʒ��,���,����Ա,ɴ��,���) select ��ʽ,��ɫ,����,����,������,����,���� from clb where  ��ʽ='" & Data10.Recordset.Fields(0) & "' AND ��ɫ='" & Data10.Recordset.Fields(1) & "'  and ����='" & Data10.Recordset.Fields(2) & "' and ����='" & Data10.Recordset.Fields(4) & "'"
Data10.Recordset.MoveNext
Loop
Data6.RecordSource = "SELECT ��ǩ as ���,�ͻ����� as ��ɫ,Ʒ�� AS ����,����Ա as ������,��� as ����,FORMAT(SUM(val(ɴ��)),'#0.00') as �ƻ���,FORMAT(SUM(val(���)),'#0.00') as ������ from CLBSC group by ��ǩ,�ͻ�����,Ʒ��,���,����Ա order by ��ǩ,�ͻ�����,Ʒ��,����Ա"
Data6.Refresh
Call sx

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If DBCombo4.Text = "" Then
Data6.RecordSource = "SELECT * FROM CLB WHERE ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') AND CLB.����Ա='" & DBCombo1.Text & "' ORDER BY ����"
Data6.Refresh
Data11.RecordSource = "SELECT ����,count(������) as ����,sum(����) as ͳ���� FROM CLB WHERE ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') AND CLB.����Ա='" & DBCombo1.Text & "' group by ���� ORDER BY ����"
Data11.Refresh
Data12.RecordSource = "SELECT count(������) as ����,sum(����) as ͳ���� FROM CLB WHERE ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') AND CLB.����Ա='" & DBCombo1.Text & "'"
Data12.Refresh
Else
Data6.RecordSource = "SELECT * FROM CLB WHERE ��ʽ='" & DBCombo4.Text & "' AND ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') AND CLB.����Ա='" & DBCombo1.Text & "' ORDER BY ����"
Data6.Refresh
Data11.RecordSource = "SELECT ����,count(������) as ����,sum(����) as ͳ���� FROM CLB WHERE ��ʽ='" & DBCombo4.Text & "' AND ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') AND CLB.����Ա='" & DBCombo1.Text & "' group by ���� ORDER BY ����"
Data11.Refresh
Data12.RecordSource = "SELECT count(������) as ����,sum(����) as ͳ���� FROM CLB WHERE ��ʽ='" & DBCombo4.Text & "' AND ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') AND CLB.����Ա='" & DBCombo1.Text & "'"
Data12.Refresh
End If
End Sub

Private Sub Command4_Click()
Call OutDataToExcel(MSFlexGrid3, 10, "��������")
End Sub

Private Sub Command5_Click()
On Error Resume Next
Data2.RecordSource = "SELECT * FROM CLB WHERE ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') AND LEN(����Ա)<>3"
Data2.Refresh
If Data2.Recordset.EOF Then
MsgBox ("�����ȷ")
Exit Sub
Else
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
MsgBox (Data6.Recordset.Fields(6))
Data2.Recordset.MoveNext
Loop
End If
End Sub

Private Sub Command6_Click()
If MsgBox("ȷ��ˢ����", vbYesNo) = vbNo Then Exit Sub
Data3.RecordSource = "SELECT * FROM CLB WHERE ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
Data3.Refresh
If Data3.Recordset.EOF Then Exit Sub
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
Data1.Recordset.FindFirst "���='" & Data3.Recordset.Fields(4) & "'"
Data3.Recordset.Edit
If Data1.Recordset.NoMatch Then
Data3.Recordset.Fields(8) = "��"
Else
Data3.Recordset.Fields(8) = Data1.Recordset.Fields(0)
End If
Data3.Recordset.Update
Data3.Recordset.MoveNext
Loop
MsgBox ("������ˢ��")
End Sub





Private Sub Command7_Click()
If DBCombo2.Text = "" Then
MsgBox ("�����뵥��")
Exit Sub
End If

Data10.RecordSource = "SELECT ���,��ɫ,����,����,���� FROM cmb WHERE ����='" & DBCombo2.Text & "'"
Data10.Refresh
If Data10.Recordset.EOF Then Exit Sub
lo = "d:\���ݿ�\\htgl\2011\DB.MDB"
Data7.Database.Execute "delete * from clbsc"
Data10.Recordset.MoveFirst
Do While Not Data10.Recordset.EOF
Data7.Database.Execute "UPDATE CLB SET ����='" & Data10.Recordset.Fields(3) & "' WHERE  ��ʽ='" & Data10.Recordset.Fields(0) & "' AND ��ɫ='" & Data10.Recordset.Fields(1) & "' and ����='" & Data10.Recordset.Fields(2) & "' and ����='" & Data10.Recordset.Fields(4) & "'"
Data7.Database.Execute "insert into CLBSC(��ǩ,�ͻ�����,Ʒ��,���,����Ա,ɴ��,���) select ��ʽ,��ɫ,����,����,������,����,���� from clb where  ��ʽ='" & Data10.Recordset.Fields(0) & "' AND ��ɫ='" & Data10.Recordset.Fields(1) & "'  and ����='" & Data10.Recordset.Fields(2) & "' and ����='" & Data10.Recordset.Fields(4) & "'"
Data10.Recordset.MoveNext
Loop
Data6.RecordSource = "SELECT ��ǩ as ���,�ͻ����� as ��ɫ,Ʒ�� AS ����,����Ա as ������,��� as ����,FORMAT(SUM(val(ɴ��)),'#0.00') as �ƻ���,FORMAT(SUM(val(���)),'#0.00') as ������ from CLBSC group by ��ǩ,�ͻ�����,Ʒ��,���,����Ա order by ��ǩ,�ͻ�����,Ʒ��,����Ա"
Data6.Refresh
Call sx
End Sub

Private Sub Command8_Click()
On Error Resume Next
If DBCombo2.Text = "" Then
MsgBox ("�����뵥��")
Exit Sub
End If

Data7.Database.Execute "delete * from clbsc"
Data7.Database.Execute "insert into CLBSC(��ǩ,�ͻ�����,Ʒ��,���,����Ա,ɴ��,���) select ��ʽ,��ɫ,����,����,������,����,���� from clb where ����='" & DBCombo2.Text & "'"

Data6.RecordSource = "SELECT ��ǩ as ���,�ͻ����� as ��ɫ,Ʒ�� AS ����,����Ա as ������,��� as ����,FORMAT(SUM(val(���)),'#0.00') as ������ from CLBSC group by ��ǩ,�ͻ�����,Ʒ��,���,����Ա order by ��ǩ,�ͻ�����,Ʒ��,����Ա"
Data6.Refresh

End Sub

Private Sub DBCombo8_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Command9_Click()
If DBCombo2.Text = "" Then
MsgBox ("�����뵥��")
Exit Sub
End If

Data7.Database.Execute "delete * from clbsc"
Data7.Database.Execute "insert into CLBSC(��ǩ,�ͻ�����,Ʒ��,���,����Ա,ɴ��,���) select ��ʽ,��ɫ,����,����,������,����,���� from clb where ����='" & DBCombo2.Text & "' and ��ʽ='" & DBCombo3.Text & "'"

Data6.RecordSource = "SELECT ��ǩ as ���,�ͻ����� as ��ɫ,Ʒ�� AS ����,����Ա as ������,��� as ����,FORMAT(SUM(val(���)),'#0.00') as ������ from CLBSC group by ��ǩ,�ͻ�����,Ʒ��,���,����Ա order by ��ǩ,�ͻ�����,Ʒ��,����Ա"
Data6.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
DBCombo4.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""
Text1.Text = 0
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\cw.MDB"
Data1.RecordSource = "SELECT * FROM WORKS"
Data1.Refresh

Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\DB.MDB"

Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\DB.MDB"
Data3.RecordSource = "SELECT * FROM CLB "
Data3.Refresh

Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\sczyjhd.mdb"
Data4.RecordSource = "select ���  from KHZL group by ���"
Data4.Refresh

Data5.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"
Data5.RecordSource = "select ct.������  from ct group by ct.������ ORDER BY VAL(CT.������)"
Data5.Refresh

Data6.DatabaseName = "d:\���ݿ�\\htgl\2011\DB.MDB"
Data6.RecordSource = "SELECT * FROM CLB WHERE CLB.��ʽ='" & DBCombo4.Text & "' ORDER BY VAL(������)"
Data6.Refresh

Data7.DatabaseName = "d:\���ݿ�\\htgl\2011\db.MDB"

Data8.DatabaseName = "d:\���ݿ�\\htgl\2011\cw.MDB"
Data8.RecordSource = "SELECT GDINGXSHU.������ FROM GDINGXSHU GROUP BY GDINGXSHU.������"
Data8.Refresh

Data9.DatabaseName = "d:\���ݿ�\\htgl\2011\cw.MDB"
Data9.RecordSource = "GDINGXSHU"
Data9.Refresh

Data10.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"
Data11.DatabaseName = "d:\���ݿ�\\htgl\2011\db.MDB"
Data12.DatabaseName = "d:\���ݿ�\\htgl\2011\db.MDB"


MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid3.ColWidth(0) = 200
MSFlexGrid3.ColWidth(1) = 1500
MSFlexGrid3.ColWidth(2) = 1200
MSFlexGrid3.ColWidth(3) = 1200
MSFlexGrid3.ColWidth(4) = 1200
MSFlexGrid3.ColWidth(5) = 1200
MSFlexGrid3.ColWidth(6) = 1200
MSFlexGrid3.ColWidth(7) = 2200
MSFlexGrid3.ColWidth(8) = 1200
MSFlexGrid3.ColWidth(9) = 1200
MSFlexGrid3.ColWidth(10) = 1200
MSFlexGrid2.ColWidth(0) = 200
MSFlexGrid2.ColWidth(1) = 1500

End Sub

Private Sub sx()
On Error Resume Next
    Dim i     As Integer
      With MSFlexGrid3
                 .AllowBigSelection = True           '   ����������ʽ
                 .FillStyle = flexFillRepeat
                For i = 1 To .Rows - 1
                        .Row = i:       .Col = .FixedCols
                        .ColSel = .Cols() - .FixedCols - 1
                         If (Val(MSFlexGrid3.TextMatrix(i, 6)) + Val(Text1.Text)) < Val(MSFlexGrid3.TextMatrix(i, 7)) Then
                              .CellBackColor = vbGreen          ' ��ɫ
                        Else
                              .CellBackColor = vbBlack       '��ɫ
                        End If
                Next i
        End With
End Sub


Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 2
       khbl = 21
Formb202.Show
End Select
End Sub

