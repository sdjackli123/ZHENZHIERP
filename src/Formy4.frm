VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formy4 
   BackColor       =   &H00C0E0FF&
   Caption         =   "����ͻ�������Ϣ"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   840
      TabIndex        =   14
      Top             =   2760
      Width           =   2775
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1320
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   600
      Width           =   2415
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox Text1111 
      Height          =   270
      Left            =   7200
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   5280
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ܱ��ϱ�"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
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
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Top             =   10080
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
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   2775
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
      Top             =   9960
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
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formy4.frx":0000
      Height          =   8895
      Left            =   5280
      TabIndex        =   1
      Top             =   600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   15690
      _Version        =   393216
      Cols            =   8
      BackColorFixed  =   8421631
      BackColorBkg    =   4109501
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   36892
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6255
      Left            =   840
      TabIndex        =   9
      Top             =   3840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11033
      _Version        =   393217
      Style           =   7
      Appearance      =   1
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
      Height          =   615
      Index           =   1
      Left            =   840
      TabIndex        =   12
      Top             =   600
      Width           =   495
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
      Left            =   840
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
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
      Left            =   840
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "Formy4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command10_Click()
If Data4.Recordset.EOF Then
MsgBox ("û�����ݣ����ܴ�ӡ��")
Exit Sub
End If
Call MXOutDataToExcel(MSFlexGrid3, "�������ϱ�                     " + "���ţ�" + Text1.Text + "��Լ�ţ�  " + HYH + "           ������ţ�" + GZBH + " ���ڣ�" + Str(JHRQ) + "      �ƻ����ڣ�" + Str(JHQ) + "     ��ӡ����" + DYRQ)
End Sub




Private Sub Command14_Click()
Data4.Database.Execute "UPDATE DHCLB SET ��������='' WHERE ��������=NULL AND ����='" & Text1.Text & "'"
Data4.Database.Execute "update dhclb set ��������=TRIM(��������) where ����='" & Text1.Text & "'"
Data4.Database.Execute "UPDATE DHCLB SET ���Ϲ��='' WHERE ���Ϲ��=NULL AND ����='" & Text1.Text & "'"
Data4.Database.Execute "update dhclb set ���Ϲ��=TRIM(���Ϲ��) where ����='" & Text1.Text & "'"
Data4.Database.Execute "UPDATE DHCLB SET ������ɫ='' WHERE ������ɫ=NULL AND ����='" & Text1.Text & "'"
Data4.Database.Execute "update dhclb set ������ɫ=TRIM(������ɫ) where ����='" & Text1.Text & "'"
'Data4.Database.Execute "delete * from dhclb where ����='" & text1.Text & "' and trim(��������)='A'"
'Data4.RecordSource = "SELECT DHCLB.���Ͽ���,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ,SUM(DHCLB.��������) AS ���� FROM DHCLB WHERE  DHCLB.����='" & text1.Text & "' GROUP BY DHCLB.���Ͽ���,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ"
'Data4.Refresh
Data4.Database.Execute "UPDATE DHCLB SET ��������='' WHERE LEN(TRIM(��������))=0 AND ����='" & Text1.Text & "'"
Data4.Database.Execute "update dhclb set ��������=TRIM(��������) where ����='" & Text1.Text & "'"
Data4.RecordSource = "SELECT ����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,Format(SUM(��������),'#0.00') AS ���� FROM DHCLB WHERE  ����='" & Text1.Text & "' GROUP BY ����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������"
Data4.Refresh
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Call tree
Call zk
End Sub

Private Sub Command8_Click()
'On Error Resume Next
If MsgBox("ȷ���Զ������𣿣�һ��ѡ���Զ����ϣ���ô��ǰ���ɵĽ�ɾ�����������µı��ϱ�", vbYesNo) = vbNo Then Exit Sub
Data2.RecordSource = "select distinct ��� from sczy_xdh where ����='" & Text1.Text & "'"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data2.Recordset.MoveFirst
Data3.Database.Execute "delete * from zdlclb"
Do While Not Data2.Recordset.EOF
Data3.Database.Execute "insert into zdlclb(���,������ɫ,��������,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,��������,���Ͽ���) select ���,������ɫ,��������,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,sum(��������),���Ͽ��� from dlclb where ���='" & Data2.Recordset.Fields(0) & "' group by ���,������ɫ,��������,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,���Ͽ���"
Data2.Recordset.MoveNext
Loop
End If

Data2.RecordSource = "select ���,��ɫ,����,���� from cmb where ����='" & Text1.Text & "' order by ���"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data7.Database.Execute "DELETE * FROM DHCLB WHERE ����='" & Text1.Text & "'"
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Data7.RecordSource = "select * from zdlclb where ���='" & Data2.Recordset.Fields(0) & "' and ������ɫ='" & Data2.Recordset.Fields(1) & "' and ��������='" & Data2.Recordset.Fields(2) & "'"
Data7.Refresh
If Data7.Recordset.EOF Then
MsgBox ("���" + Data2.Recordset.Fields(0) + "��ɫ" + Data2.Recordset.Fields(1) + "����" + Data2.Recordset.Fields(2) + "û�е���")
Exit Sub
End If

Data3.Database.Execute "insert into dhclb(����,���,������ɫ,��������,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,��������,���Ͽ���) select '" & Text1.Text & "', ���,������ɫ,��������,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,��������*val('" & Data2.Recordset.Fields(3) & "'),���Ͽ��� from zdlclb where ���='" & Data2.Recordset.Fields(0) & "' and ������ɫ='" & Data2.Recordset.Fields(1) & "' and ��������='" & Data2.Recordset.Fields(2) & "'"

Data2.Recordset.MoveNext
Loop
End If

Data7.Database.Execute "INSERT INTO DHCLBY(����,���,������ɫ,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,��������) SELECT ����,���,������ɫ,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,Format(SUM(��������),'#0.00') AS SL FROM DHCLB WHERE ����='" & Text1.Text & "' GROUP BY ����,���,������ɫ,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������"
Data7.Database.Execute "UPDATE DHCLBY SET ���ϵ�λ='��',��������=��������/1500 WHERE ���Ͽ���='2���Ͽ�' AND ��������='�ַ�����' AND ���ϵ�λ='��'"
Data7.Database.Execute "UPDATE DHCLBY SET ���ϵ�λ='��',��������=��������/2700 WHERE ���Ͽ���='2���Ͽ�' AND ��������='������' AND ���ϵ�λ='��'"
Data7.Database.Execute "UPDATE DHCLBY SET ��������=INT(��������)+1 WHERE ���Ͽ���='2���Ͽ�'AND INT(��������/2)<>��������/2 and ���ϵ�λ='��' and (��������='������' or ��������='�ַ�����')"
Data7.Database.Execute "DELETE * FROM DHCLB WHERE ����='" & Text1.Text & "'"
Data7.Database.Execute "INSERT INTO DHCLB(����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,��������) SELECT ����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,Format(SUM(��������),'#0.00') AS SL FROM DHCLBY WHERE ����='" & Text1.Text & "' GROUP BY ����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������"
Data7.Database.Execute "DELETE * FROM DHCLBY WHERE ����='" & Text1.Text & "'"

Data4.RecordSource = "SELECT ����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������� FROM DHCLB WHERE DHCLB.����='" & Text1.Text & "'"
Data4.Refresh

End Sub



Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
DTPicker1.Value = Date - 15
DTPicker2.Value = Date
Option1.Value = True
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select ��� from khZL group by ���"
Data1.Refresh
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"
Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"

Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"
Data4.RecordSource = "SELECT * FROM DHCLB WHERE DHCLB.����='" & Text1.Text & "'"
Data4.Refresh

Data5.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"
Data6.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"

Data7.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"


MSFlexGrid3.ColWidth(0) = 200
MSFlexGrid3.ColWidth(1) = 1500

End Sub


Private Sub MSFlex()
With MSFlexGrid3
    c = .Col: r = .Row    '''''C�У���R��
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid3.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid3.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid3.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data4.Recordset.MoveFirst
Data4.Recordset.Move r - 1
Data4.Recordset.Edit

Data4.Recordset.Fields(c - 1) = Text1111.Text
Data4.Recordset.Update

Text1111.Visible = False
MSFlexGrid3.SetFocus
End Sub


Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex

   TreeView1.Nodes.Clear
   If Option1.Value = True Then
    Data5.RecordSource = "select distinct �ͻ� from sczy_xdh where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
    Data5.Refresh
    m = 1
    If Not Data5.Recordset.EOF Then  'make sure there are records in the table
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data5.Recordset.Fields(0)
        intIndex = mNode.Index
        Data6.RecordSource = "select distinct ���� from sczy_xdh where �ͻ�='" & Data5.Recordset.Fields(0) & "' and  ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data7.RecordSource = "select distinct ��� from sczy_xdh where ����='" & Data6.Recordset.Fields(0) & "' and ����='����'"
        Data7.Refresh
        
        If Not Data7.Recordset.EOF Then
        Data7.Recordset.MoveFirst
        Do While Not Data7.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data7.Recordset.Fields(0))
        Data7.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data6.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data5.Recordset.MoveNext
        Loop
    End If
    End If
 
    If Option2.Value = True Then
    Data5.RecordSource = "select distinct �ͻ� from sczy_xdh where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
    Data5.Refresh
    m = 1
    If Not Data5.Recordset.EOF Then  'make sure there are records in the table
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data5.Recordset.Fields(0)
        intIndex = mNode.Index
        Data6.RecordSource = "select distinct ���� from sczy_xdh where �ͻ�='" & Data5.Recordset.Fields(0) & "' and  ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and ����='����'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data7.RecordSource = "select distinct ��� from sczy_xdh where ����='" & Data6.Recordset.Fields(0) & "' and ����='����'"
        Data7.Refresh
        
        If Not Data7.Recordset.EOF Then
        Data7.Recordset.MoveFirst
        Do While Not Data7.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data7.Recordset.Fields(0))
        Data7.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data6.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data5.Recordset.MoveNext
        Loop
    End If
    End If

End Sub


'Ȼ��ô����ֻ��Խ�С�ļ�¼������ѭ�������Ч�ʱȽϸߡ��޸ĺ�Ĵ������£�

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next



If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") > 0 Then
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
Text1.Text = l1
Data4.RecordSource = "SELECT ����,���,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������� FROM DHCLB WHERE DHCLB.����='" & Text1.Text & "'"
Data4.Refresh
End If

End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True 'չ�����нڵ�
  Next i
End Sub


