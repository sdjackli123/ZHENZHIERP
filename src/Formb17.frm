VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formb17 
   BackColor       =   &H00C0E0FF&
   Caption         =   "��������"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   2160
      Style           =   1  'Simple Combo
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   4815
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
      Height          =   375
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
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
      Height          =   375
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "������ϸ"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����ӡ"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ű���"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   960
      Width           =   1695
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   1815
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   6495
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   3855
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3855
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   1815
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   2415
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
      Top             =   9960
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
      Height          =   855
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   1
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
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ڱ���"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3255
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   3855
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formb17.frx":0000
      Height          =   8175
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   14420
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   36892
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
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
      Left            =   5880
      TabIndex        =   12
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   0
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   1695
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
      TabIndex        =   6
      Top             =   480
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
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Formb17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer: Public YPDH As String


Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox ("�����뵥��")
Exit Sub
End If

Data7.Database.Execute "delete * from clbsc"
Data7.Database.Execute "insert into CLBSC(��ǩ,�ͻ�����,Ʒ��,���,����Ա,ɴ��,���) select ��ʽ,��ɫ,����,����,������,����,���� from clb where ����='" & Text1.Text & "'"

Data6.RecordSource = "SELECT ��ǩ as ���,�ͻ����� as ��ɫ,Ʒ�� AS ����,����Ա as ������,��� as ����,FORMAT(SUM(val(���)),'#0.00') as ������ from CLBSC group by ��ǩ,�ͻ�����,Ʒ��,���,����Ա order by ��ǩ,�ͻ�����,Ʒ��,����Ա"
Data6.Refresh

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call OutDataToExcel(MSFlexGrid3, 10, "��������")
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Text1.Text = "" Then
Data6.RecordSource = "SELECT * FROM CLB WHERE ����='" & Text1.Text & "' and ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
Data6.Refresh
Else
Data6.RecordSource = "SELECT * FROM CLB WHERE  ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
Data6.Refresh
End If
End Sub

Private Sub Command6_Click()
If MsgBox("ȷ������ˢ����", vbYesNo) = vbNo Then Exit Sub
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





Private Sub Command5_Click()
On Error Resume Next
If MsgBox("ȷ�ϵ���ˢ����", vbYesNo) = vbNo Then Exit Sub

Data3.RecordSource = "SELECT * FROM CLB WHERE ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
Data3.Refresh
If Data3.Recordset.EOF Then Exit Sub
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF

Data9.Recordset.FindFirst "������='" & Data3.Recordset.Fields(11) & "' and ������='" & Data3.Recordset.Fields(2) & "' and ��λ����='" & Data3.Recordset.Fields(13) & "'"
Data3.Recordset.Edit
If Data9.Recordset.NoMatch Then
Data9.Recordset.FindFirst "������='" & Data3.Recordset.Fields(11) & "' and ������='" & Data3.Recordset.Fields(2) & "'"
If Data9.Recordset.NoMatch Then
Data3.Recordset.Fields(7) = 0
Else
Data3.Recordset.Fields(7) = Data9.Recordset.Fields(2)
End If
Else
Data3.Recordset.Fields(7) = Data9.Recordset.Fields(2)
End If

Data3.Recordset.Update
Data3.Recordset.MoveNext
Loop
Data7.Database.Execute "UPDATE CLB SET �ϼƽ��=TRIM(format(VAL(ϵ��)*����,'#0.00')) WHERE ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
MsgBox ("������ˢ��")

End Sub

Private Sub Command7_Click()
On Error Resume Next
Data7.Database.Execute "delete * from clbsc"
Data7.Database.Execute "insert into CLBSC(��ǩ,�ͻ�����,Ʒ��,���,����Ա,ɴ��,���) select ��ʽ,��ɫ,����,����,������,����,���� from clb where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data6.RecordSource = "SELECT ��ǩ as ���,�ͻ����� as ��ɫ,Ʒ�� AS ����,����Ա as ������,��� as ����,FORMAT(SUM(val(���)),'#0.00') as ������ from CLBSC group by ��ǩ,�ͻ�����,Ʒ��,���,����Ա order by ��ǩ,�ͻ�����,Ʒ��,����Ա"
Data6.Refresh
End Sub

Private Sub DBCombo8_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Form_Load()
On Error Resume Next

DTPicker1.Value = Date
DTPicker2.Value = Date

Text1.Text = ""
Text2.Text = ""
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\cw.MDB"
Data1.RecordSource = "SELECT * FROM WORKS"
Data1.Refresh

Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\DB.MDB"

Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\DB.MDB"

Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\sczyjhd.mdb"
Data4.RecordSource = "select ���  from KHZL group by ���"
Data4.Refresh

Data5.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"
Data5.RecordSource = "select ct.������  from ct group by ct.������ ORDER BY VAL(CT.������)"
Data5.Refresh

Data6.DatabaseName = "d:\���ݿ�\\htgl\2011\DB.MDB"

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

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
khbl = 18
Formb202.Show
End Select
End Sub


Private Sub MSFlex()
With MSFlexGrid3
    c = .Col: r = .Row    '''''C�У���R��
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
End With
End Sub


Private Sub MSFlexGrid3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    MSFlexGrid3.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data6.Recordset.MoveFirst
Data6.Recordset.Move r - 1
Data6.Recordset.Edit
Data6.Recordset.Fields(c - 1) = Combo1111.Text
Data6.Recordset.Update
MSFlexGrid3.Text = Combo1111.Text
Combo1111.Visible = False
MSFlexGrid3.SetFocus
End If
End Sub


