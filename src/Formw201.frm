VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formw201 
   BackColor       =   &H00C0E0FF&
   Caption         =   "���ۼƻ���Ϣ"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ɸѡ����"
      Height          =   855
      Left            =   8160
      TabIndex        =   10
      Top             =   600
      Width           =   2895
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���ۿ��С��"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Data Data9 
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
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data8 
      Caption         =   "Data1"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data7 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data6 
      Caption         =   "Data1"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data5 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data4 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data3 
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
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
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
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data2 
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
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81526785
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81526785
      CurrentDate     =   39557
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw201.frx":0000
      Height          =   7815
      Left            =   3360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1920
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   13785
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7455
      Left            =   480
      TabIndex        =   8
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   13150
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Formw201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call zxd(MSFlexGrid1, "���ۼƻ�")
End Sub

Private Sub Command4_Click()
Command4.Enabled = False
Call tree
Call zk
Command4.Enabled = True
End Sub

Private Sub Command6_Click()
Command6.Enabled = False
'''''''''''''''''''''''''''''''''����
Data3.Database.Execute "delete * from xsjh"
Data1.RecordSource = "select * from sczy_xdd where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data1.Refresh
If Not Data1.Recordset.EOF Then
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF

For k = 0 To 7
If Data1.Recordset.Fields(2 * k + 6) <> "" Then
Data3.Database.Execute "INSERT INTO xsjh(���,��λ,��ɫ,���,��������) VALUES('" & Data1.Recordset.Fields(1) & "','" & Data1.Recordset.Fields(4) & "','" & Data1.Recordset.Fields(2) & "','" & Data1.Recordset.Fields(2 * k + 6) & "','" & Data1.Recordset.Fields(2 * k + 7) & "')"
End If
Next
Data1.Recordset.MoveNext
Loop
End If

'Data3.Database.Execute "delete * from xsjh"     '�˵�
Data1.RecordSource = "select * from sczy_xtd where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data1.Refresh
If Not Data1.Recordset.EOF Then
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF

For k = 0 To 7
If Data1.Recordset.Fields(2 * k + 6) <> "" Then
Data3.Database.Execute "INSERT INTO xsjh(���,��λ,��ɫ,���,��������) VALUES('" & Data1.Recordset.Fields(1) & "','" & Data1.Recordset.Fields(4) & "','" & Data1.Recordset.Fields(2) & "','" & Data1.Recordset.Fields(2 * k + 6) & "',-val('" & Data1.Recordset.Fields(2 * k + 7) & "'))"
End If
Next
Data1.Recordset.MoveNext
Loop
End If

''''''''''''''''''''''''''''''''����
Data2.RecordSource = "select * from zxd where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data2.Refresh
If Not Data2.Recordset.EOF Then
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF

For k = 0 To 4
If Data2.Recordset.Fields(2 * k + 6) <> "" Then
Data3.Database.Execute "INSERT INTO xsjh(���,��λ,���,��ɫ,��������) VALUES('" & Data2.Recordset.Fields(1) & "','" & Data2.Recordset.Fields(3) & "','" & Data2.Recordset.Fields(2) & "','" & Data2.Recordset.Fields(2 * k + 5) & "','" & Data2.Recordset.Fields(2 * k + 6) & "')"
End If
Next
Data2.Recordset.MoveNext
Loop
End If
''''''''''''''''''''''''''''''''���
lo = "e:\Excel\Ⱦ��\��¡\sjzz.mdb"
Data2.Database.Execute "insert into xsjh(���,��λ,��ɫ,���,�������) in'" & lo & "' select ���,��λ,���,�ͺ�,sum(����) from lsrk where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') group by ���,��λ,���,�ͺ�"

''''''''''''''''''''''''''''''''���
Call cpkc

Data2.Database.Execute "insert into xsjh(���,��λ,��ɫ,���,�������) in'" & lo & "' select ���,��λ,���,�ͺ�,sum(����) from lskc group by ���,��λ,���,�ͺ�"

'''''''''''''''''''''''''''''''������
Data5.Database.Execute "insert into xsjh(���,��ɫ,���,��������) in'" & lo & "' select ���,��ɫ,���,sum(val(�ü�)) from cjrb where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') group by ���,��ɫ,���"

'''''''''''''''''''''''''''''''�ƻ���
Data1.RecordSource = "select * from sczy_xdh where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data1.Refresh
If Not Data1.Recordset.EOF Then
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF

For k = 0 To 7
If Data1.Recordset.Fields(2 * k + 6) <> "" Then
Data3.Database.Execute "INSERT INTO xsjh(���,��λ,��ɫ,���,�ƻ�����) VALUES('" & Data1.Recordset.Fields(1) & "','" & Data1.Recordset.Fields(4) & "','" & Data1.Recordset.Fields(2) & "','" & Data1.Recordset.Fields(2 * k + 6) & "','" & Data1.Recordset.Fields(2 * k + 7) & "')"
End If
Next
Data1.Recordset.MoveNext
Loop
End If
''''''''''''''''''''''''''''''''���ڼ���
Data4.Database.Execute "insert into xsjh(���,��λ,��ɫ,���,��������,��������) in'" & lo & "' select ���,��λ,��ɫ,���,ת����������,תδ�������� from xsjh where ����=cdate('" & DTPicker1.Value & "') "
''''''''''''''''''''''''''''''''''���
Data3.Database.Execute "update xsjh set ���='' where ���=null"
Data3.Database.Execute "update xsjh set ��������='0' where ��������=null"
Data3.Database.Execute "update xsjh set ��������='0' where ��������=null"
Data3.Database.Execute "update xsjh set ��������='0' where ��������=null"
Data3.Database.Execute "update xsjh set �������='0' where �������=null"
Data3.Database.Execute "update xsjh set �������='0' where �������=null"
Data3.Database.Execute "update xsjh set ��������='0' where ��������=null"
Data3.Database.Execute "update xsjh set �ƻ�����='0' where �ƻ�����=null"
Data3.Database.Execute "update xsjh set δ������='0' where δ������=null"
Data3.Database.Execute "update xsjh set ��������='0' where ��������=null"
Data3.Database.Execute "update xsjh set ���ۿ��='0' where ���ۿ��=null"
Data3.Database.Execute "update xsjh set ��������='0' where ��������=null"

Data3.Database.Execute "update xsjh set ����='0'"
Data3.Database.Execute "insert into xsjh(���,��ɫ,���,��������,��������,��������,�������,�������,��������,�ƻ�����,δ������,���ۿ��) select ���,��ɫ,���,sum(val(��������)),sum(val(��������)),sum(val(��������)),sum(val(�������)),sum(val(�������)),sum(val(��������)),sum(val(�ƻ�����)),sum(val(δ������)),sum(val(���ۿ��)) from xsjh group by ���,��ɫ,���"
Data3.Database.Execute "delete * from xsjh where ����='0' or ���=''"

Data3.Database.Execute "update xsjh set δ������='0' where δ������=null"
Data3.Database.Execute "update xsjh set ��������='0' where ��������=null"
Data3.Database.Execute "update xsjh set ���ۿ��='0' where ���ۿ��=null"
Data3.Database.Execute "update xsjh set ��������='0' where ��������=null"

Data3.Database.Execute "update xsjh set δ������=val(�ƻ�����)-val(��������)+val(��������)"
Data3.Database.Execute "update xsjh set ��������=val(��������)+val(��������)-val(�������)"
Data3.Database.Execute "update xsjh set ��������=val(��������)-val(��������),���ۿ��=val(�������)+val(��������)+val(δ������)-val(��������)+val(��������)"


Data3.RecordSource = "select * from xsjh order by ���,���"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.Recordset.MoveFirst
i = 1
Do While Not Data3.Recordset.EOF
Data3.Recordset.Edit
Data3.Recordset.Fields(16) = i
Data3.Recordset.Update
Data3.Recordset.MoveNext
i = i + 1
Loop
End If
Command6.Enabled = True
End Sub

Private Sub Command7_Click()
If Val(Text1.Text) > 0 Then
Data3.RecordSource = "select * from xsjh where val(���ۿ��)<val('" & Text1.Text & "') order by ���,���"
Data3.Refresh
Else
Data3.RecordSource = "select * from xsjh order by ���,���"
Data3.Refresh
End If
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date
Text1.Text = "0"
Data1.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\SCZYJHD.mdb"
Data2.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data3.DatabaseName = "e:\Excel\Ⱦ��\��¡\sjzz.mdb"
Data4.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\SCZYJHD.mdb"
Data5.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\SCJD.mdb"
Data6.DatabaseName = "e:\Excel\Ⱦ��\��¡\sjzz.mdb"
Data7.DatabaseName = "e:\Excel\Ⱦ��\��¡\sjzz.mdb"
Data8.DatabaseName = "e:\Excel\Ⱦ��\��¡\sjzz.mdb"
Data9.DatabaseName = "e:\Excel\Ⱦ��\��¡\sjzz.mdb"
MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(2) = 0
MSFlexGrid1.ColWidth(3) = 0
MSFlexGrid1.ColWidth(15) = 0
MSFlexGrid1.ColWidth(16) = 0
MSFlexGrid1.ColWidth(17) = 0
MSFlexGrid1.ColWidth(18) = 0
End Sub

Private Sub cpkc()
       Data2.Database.Execute "DELETE * FROM LSKC"
       Data2.Database.Execute "INSERT INTO LSKC(���,Ʒ��,���,�ͺ�,��λ,����) SELECT ���,Ʒ��,���,�ͺ�,��λ,-���� FROM LSFH where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data2.Database.Execute "INSERT INTO LSKC(���,Ʒ��,���,�ͺ�,��λ,����) SELECT ���,Ʒ��,���,�ͺ�,��λ,���� FROM LSRK where  ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data2.Database.Execute "INSERT INTO LSKC(���,Ʒ��,���,�ͺ�,��λ,����) SELECT ���,Ʒ��,���,�ͺ�,��λ,���� FROM LSJL where  ����=cdate('" & DTPicker1.Value & "')"
       Data2.Database.Execute "UPDATE LSKC SET ���='1'"
       Data2.Database.Execute "INSERT INTO LSKC(���,Ʒ��,���,�ͺ�,��λ,����) SELECT ���,Ʒ��,���,�ͺ�,��λ,format(SUM(����),'#0') FROM LSKC GROUP BY ���,Ʒ��,���,�ͺ�,��λ"
       Data2.Database.Execute "DELETE * FROM LSKC WHERE  ���='1'"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") = 0 Then
Data3.RecordSource = "select * from xsjh where ���='" & TreeView1.Nodes(Node.Index).FullPath & "' order by ���,���"
Data3.Refresh
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, 1, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") - 1)
l2 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l2 = Mid(TreeView1.Nodes(Node.Index).FullPath, 1, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") - 1)
End If
Data3.RecordSource = "select * from xsjh where ���='" & l1 & "' and  ���='" & l2 & "' order by ���,���"
Data3.Refresh
End If
End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True 'չ�����нڵ�
  Next i
End Sub

Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex

   TreeView1.Nodes.Clear
   
   TreeView1.Sorted = True

If Val(Text1.Text) > 0 Then

    Data9.RecordSource = "select distinct ��� from xsjh where val(���ۿ��)<val('" & Text1.Text & "')"
    Data9.Refresh
    m = 1
    If Not Data9.Recordset.EOF Then  'make sure there are records in the table
        Data9.Recordset.MoveFirst
        Do While Not Data9.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data9.Recordset.Fields(0)
        intIndex = mNode.Index
        Data7.RecordSource = "select distinct ��� from xsjh where ���='" & Data9.Recordset.Fields(0) & "' and val(���ۿ��)<val('" & Text1.Text & "')"
        Data7.Refresh
        
        If Not Data7.Recordset.EOF Then
        Data7.Recordset.MoveFirst
        k = 1
        Do While Not Data7.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "t" + Trim(intIndex)
        mNode.Text = Trim(Data7.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data8.RecordSource = "select distinct ��ɫ from xsjh where ���='" & Data9.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(0) & "' and val(���ۿ��)<val('" & Text1.Text & "')"
        Data8.Refresh

        If Not Data8.Recordset.EOF Then
        Data8.Recordset.MoveFirst
        Do While Not Data8.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "t" + Trim(intIndex) + "x" + Trim(xintindex)
        mNode.Text = Trim(Data8.Recordset.Fields(0))
        m = m + 1
        Data8.Recordset.MoveNext
        Loop
        m = m + 1
        End If
        m = m + 1
        Data7.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data9.Recordset.MoveNext
        Loop
    End If
Else
    Data9.RecordSource = "select distinct ��� from xsjh"
    Data9.Refresh
    m = 1
    If Not Data9.Recordset.EOF Then  'make sure there are records in the table
        Data9.Recordset.MoveFirst
        Do While Not Data9.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data9.Recordset.Fields(0)
        intIndex = mNode.Index
        Data7.RecordSource = "select distinct ��� from xsjh where ���='" & Data9.Recordset.Fields(0) & "'"
        Data7.Refresh
        
        If Not Data7.Recordset.EOF Then
        Data7.Recordset.MoveFirst
        k = 1
        Do While Not Data7.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "t" + Trim(intIndex)
        mNode.Text = Trim(Data7.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data8.RecordSource = "select distinct ��ɫ from xsjh where ���='" & Data9.Recordset.Fields(0) & "' and ���='" & Data7.Recordset.Fields(0) & "'"
        Data8.Refresh

        If Not Data8.Recordset.EOF Then
        Data8.Recordset.MoveFirst
        Do While Not Data8.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "t" + Trim(intIndex) + "x" + Trim(xintindex)
        mNode.Text = Trim(Data8.Recordset.Fields(0))
        m = m + 1
        Data8.Recordset.MoveNext
        Loop
        m = m + 1
        End If
        m = m + 1
        Data7.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data9.Recordset.MoveNext
        Loop
    End If
End If




End Sub
