VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formw61 
   BackColor       =   &H00C0E0FF&
   Caption         =   "会计科目"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form58"
   ScaleHeight     =   8475
   ScaleWidth      =   7350
   StartUpPosition =   2  '屏幕中心
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   5640
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
      Caption         =   "选取"
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
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "会计科目"
      BeginProperty Font 
         Name            =   "宋体"
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
Attribute VB_Name = "Formw61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mDbBiblio As Database
Private mNode As Node

Private Sub Command2_Click()
If KMMC = 0 Then Exit Sub
If KMMC = 7 Then
Formw113.Text1(1).Text = Text3.Text
KMMC = 0
Unload Me
End If

If KMMC = 8 Then
Formw115.Text1(1).Text = Text3.Text
KMMC = 0
Unload Me
End If

If KMMC = 9 Then
Formw116.Text1(1).Text = Text3.Text
KMMC = 0
Unload Me
End If



End Sub

Private Sub Form_Load()

   '在 Formw_Load 事件中，设置对象变量，
   '并创建 TreeView 控件的第一个 Node 对象。
On Error Resume Next
   Dim rsPublishers As Recordset
   Dim rsTitles As Recordset
   Dim intIndex

   Set mDbBiblio = DBEngine.Workspaces(0). _
   OpenDatabase("d:\数据库\bfrz\" + ljb + "\ZCW.MDB")

   tvwDB.Sorted = True
   Set mNode = tvwDB.Nodes.Add()
   mNode.Text = "会计科目"
   mNode.Tag = "会计科目"   '设置 Tag 属性。
  ' mNode.Image = "closed"         '设置 Image
   Set rsPublishers = mDbBiblio. _
   OpenRecordset("SELECT * FROM CWMC WHERE LEN(科目编号)=4")
   Do Until rsPublishers.EOF
      Set mNode = tvwDB.Nodes.Add(1, tvwChild)
      mNode.Text = rsPublishers!科目名称
      mNode.Tag = "Publisher" '标识表。
      mNode.Key = rsPublishers!科目编号 & " ID"
     ' mNode.Image = "closed"
      intIndex = mNode.Index
      '对这条记录，使用查询创建 Title 表的记录集，
      '查询条件是所有包含相同 PubID 的记录。对结果记录集中
      '的每一条记录，在 TreeView 控件中加入一个 Node 对象，
      '并用记录的 Title、 ISBN 和 Author 字段为新
      'Node 对象的属性赋值。
      Set rsTitles = mDbBiblio.OpenRecordset("select * from CWMC Where INSTR(科目编号,'" & rsPublishers!科目编号 & "')>0 AND LEN(科目编号)>4 ORDER BY mid(科目名称,instr(科目名称,'-')+1)")
      Do Until rsTitles.EOF
         Set mNode = tvwDB.Nodes. _
         Add(intIndex, tvwChild)
         mNode.Text = rsTitles!科目名称   '文本。
         mNode.Key = rsTitles!科目编号      '唯一的 ID。
         mNode.Tag = "Authors"      '表名。
       '  mNode.Image = "smlBook"      '图象。
         '移动到 rsTitles 中的下一个记录。
         rsTitles.MoveNext
      Loop
      '移动到下一个 Publishers 记录。
      rsPublishers.MoveNext
   Loop
Text3.Text = ""
Text2.Text = ""
Text1.Text = ""
End Sub


'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：




Private Sub tvwDB_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Text3.Text = ""                                '属性
   Text1.Text = Node.Text
   Text2.Text = tvwDB.Nodes(Node.Index).Parent.Text
If Text1.Text = "会计科目" Then Exit Sub
If Text2.Text = "会计科目" Then
Text3.Text = Text1.Text
Else
Text3.Text = Trim(Text1.Text)
End If
End Sub


