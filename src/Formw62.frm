VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Formw62 
   BackColor       =   &H00C0E0FF&
   Caption         =   "��ƿ�Ŀ"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form58"
   ScaleHeight     =   8475
   ScaleWidth      =   7350
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Formw62.frx":0000
      Left            =   4800
      List            =   "Formw62.frx":0016
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   5880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ѡȡ"
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin MSComctlLib.TreeView tvwDB 
      Height          =   6855
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   12091
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "���"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ϸ��Ŀ"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ƿ�Ŀ"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   3135
   End
End
Attribute VB_Name = "Formw62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mDbBiblio As Database
Private mNode As Node


Private Sub Combo1_Click()
On Error Resume Next
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex
   tvwDB.Sorted = True
   tvwDB.Nodes.Clear
   
   Set mNode = tvwDB.Nodes.Add()
   mNode.Text = "��ƿ�Ŀ"
   mNode.Tag = "��ƿ�Ŀ"   '���� Tag ���ԡ�
  ' mNode.Image = "closed"         '���� Image
  If Combo1.Text = "ȫ��" Then
   Set rsPublishers = mDbBiblio. _
   OpenRecordset("SELECT * FROM CWMC WHERE LEN(��Ŀ���)=4 order by ��Ŀ���")
  Else
   Set rsPublishers = mDbBiblio. _
   OpenRecordset("SELECT * FROM CWMC WHERE LEN(��Ŀ���)=4 and ��Ŀ����='" & Combo1.Text & "' order by ��Ŀ���")
  End If
  
  m = 1
   Do Until rsPublishers.EOF
      Set mNode = tvwDB.Nodes.Add(1, tvwChild)
      mNode.Text = rsPublishers!��Ŀ����
      mNode.Tag = "Publisher" '��ʶ��
      mNode.Key = "w" + Trim(m)
     ' mNode.Image = "closed"
      intIndex = mNode.Index
      '��������¼��ʹ�ò�ѯ���� Title ��ļ�¼����
      '��ѯ���������а�����ͬ PubID �ļ�¼���Խ����¼����
      '��ÿһ����¼���� TreeView �ؼ��м���һ�� Node ����
      '���ü�¼�� Title�� ISBN �� Author �ֶ�Ϊ��
      'Node ��������Ը�ֵ��
      Set rsTitles = mDbBiblio.OpenRecordset("select * from CWMC Where mid(��Ŀ���,1,4)='" & rsPublishers!��Ŀ��� & "' AND LEN(��Ŀ���)>4 and ��Ŀ����='" & Combo1.Text & "' order by ��Ŀ���")
      
      Do Until rsTitles.EOF
         Set mNode = tvwDB.Nodes. _
         Add(intIndex, tvwChild)
         mNode.Text = rsTitles!��Ŀ����   '�ı���
         mNode.Key = "w" + Trim(m) + "t" + Trim(intIndex)
         mNode.Tag = "Authors"      '������
       '  mNode.Image = "smlBook"      'ͼ��
         '�ƶ��� rsTitles �е���һ����¼��
         rsTitles.MoveNext
      Loop
      m = m + 1
      rsPublishers.MoveNext
   Loop

End Sub

Private Sub Command2_Click()
If KMMC = 0 Then Exit Sub
If KMMC = 1 Then
Data2.Recordset.FindFirst "��Ŀ����='" & Text3.Text & "'"
If Data2.Recordset.NoMatch Then
MsgBox ("��ƿ�Ŀ��������")
Exit Sub
Else
Formw111.Text1(6).Text = Data2.Recordset.Fields("��Ŀ����")
End If

If Text1.Text <> "" Then
Formw111.Text1(3).Text = Text3.Text + "-" + Text1.Text
Else
Formw111.Text1(3).Text = Text3.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 2 Then
If Text1.Text = "" Then
Formw35.DataCombo2(KMBL).Text = Text3.Text
Else
Formw35.DataCombo2(KMBL).Text = Text3.Text
Formw35.DataCombo3(KMBL).Text = Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 4 Then
If Text1.Text = "" Then
Formw35.DataCombo4(KMBL).Text = Text3.Text
Else
Formw35.DataCombo4(KMBL).Text = Text3.Text
Formw35.DataCombo5(KMBL).Text = Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 14 Then
If Text1.Text = "" Then
Formw114.DataCombo1(0).Text = Text3.Text
Else
Formw114.DataCombo1(0).Text = Text3.Text + "-" + Text1.Text
End If
KMMC = 0
Unload Me
End If


If KMMC = 101 Then
Data2.Recordset.FindFirst "��Ŀ����='" & Text3.Text & "'"
If Data2.Recordset.NoMatch Then
MsgBox ("��ƿ�Ŀ��������")
Exit Sub
Else
Formw11111.Text1(6).Text = Data2.Recordset.Fields("��Ŀ����")
End If

If Text1.Text <> "" Then
Formw11111.Text1(3).Text = Text3.Text + "-" + Text1.Text
Else
Formw11111.Text1(3).Text = Text3.Text
End If
KMMC = 0
Unload Me
End If

End Sub

Private Sub Form_Load()
   '�� Form_Load �¼��У����ö��������
   '������ TreeView �ؼ��ĵ�һ�� Node ����
On Error Resume Next
   Dim rsPublishers As Recordset
   Dim rsTitles As Recordset
   Dim intIndex
Data2.DatabaseName = "d:\���ݿ�\bfrz\" + ljb + "\CW.MDB"
Data2.RecordSource = "SELECT * FROM CWMC"
Data2.Refresh

Combo1.Text = ""

   Set mDbBiblio = DBEngine.Workspaces(0). _
   OpenDatabase("d:\���ݿ�\bfrz\" + ljb + "\CW.MDB")

   tvwDB.Sorted = True
   Set mNode = tvwDB.Nodes.Add()
   mNode.Text = "��ƿ�Ŀ"
   mNode.Tag = "��ƿ�Ŀ"   '���� Tag ���ԡ�
  ' mNode.Image = "closed"         '���� Image
   Set rsPublishers = mDbBiblio. _
   OpenRecordset("SELECT * FROM CWMC WHERE LEN(��Ŀ���)=4 order by ��Ŀ���")
   Do Until rsPublishers.EOF
      Set mNode = tvwDB.Nodes.Add(1, tvwChild)
      mNode.Text = rsPublishers!��Ŀ����
      mNode.Tag = "Publisher" '��ʶ��
      mNode.Key = rsPublishers!��Ŀ��� & " ID"
     ' mNode.Image = "closed"
      intIndex = mNode.Index
      '��������¼��ʹ�ò�ѯ���� Title ��ļ�¼����
      '��ѯ���������а�����ͬ PubID �ļ�¼���Խ����¼����
      '��ÿһ����¼���� TreeView �ؼ��м���һ�� Node ����
      '���ü�¼�� Title�� ISBN �� Author �ֶ�Ϊ��
      'Node ��������Ը�ֵ��
      Set rsTitles = mDbBiblio.OpenRecordset("select * from CWMC Where INSTR(��Ŀ���,'" & rsPublishers!��Ŀ��� & "')>0 AND LEN(��Ŀ���)>4 order by ��Ŀ���")
      Do Until rsTitles.EOF
         Set mNode = tvwDB.Nodes. _
         Add(intIndex, tvwChild)
         mNode.Text = rsTitles!��Ŀ����   '�ı���
         mNode.Key = rsTitles!��Ŀ���      'Ψһ�� ID��
         mNode.Tag = "Authors"      '������
       '  mNode.Image = "smlBook"      'ͼ��
         '�ƶ��� rsTitles �е���һ����¼��
         rsTitles.MoveNext
      Loop
      '�ƶ�����һ�� Publishers ��¼��
      rsPublishers.MoveNext
   Loop
Text3.Text = ""
Text2.Text = ""
Text1.Text = ""
End Sub


'Ȼ��ô����ֻ��Խ�С�ļ�¼������ѭ�������Ч�ʱȽϸߡ��޸ĺ�Ĵ������£�




Private Sub tvwDB_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
'Text3.Text = ""                                '����
   Text2.Text = Node.Text
   
If Text2.Text = "��ƿ�Ŀ" Then Exit Sub
   Text2.Text = tvwDB.Nodes(Node.Index).Parent.Text

If Text2.Text = "��ƿ�Ŀ" Then
Text3.Text = Node.Text
Text1.Text = ""
Else
Text1.Text = Node.Text
End If

End Sub


