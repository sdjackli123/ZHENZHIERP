VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formw6 
   BackColor       =   &H00C0E0FF&
   Caption         =   "会计科目"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form6"
   ScaleHeight     =   7410
   ScaleWidth      =   9435
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   1200
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6120
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6120
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6240
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formw6.frx":0000
      Left            =   6720
      List            =   "Formw6.frx":0016
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选取"
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   4920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComctlLib.TreeView tvwDB 
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   12091
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "快捷"
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
      Index           =   2
      Left            =   5520
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "类别"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "明细科目"
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
      Index           =   1
      Left            =   5520
      TabIndex        =   6
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "总账科目"
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
      Index           =   0
      Left            =   5520
      TabIndex        =   5
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Formw6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mNode As Node


Private Sub Combo1_Click()
On Error Resume Next
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
   tvwDB.Sorted = True
   tvwDB.Nodes.Clear
   
   Set mNode = tvwDB.Nodes.Add()
   mNode.Text = "会计科目"
   mNode.Tag = "会计科目"   '设置 Tag 属性。
  ' mNode.Image = "closed"         '设置 Image
  
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
   
If Combo1.Text = "全部" Then
Adodc1.RecordSource = "SELECT * FROM CWMC WHERE LEN(科目编号)=4 ORDER BY 科目编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM CWMC WHERE LEN(科目编号)=4 and 科目类型='" & Combo1.Text & "' ORDER BY 科目编号"
Adodc1.Refresh
End If

   Do While Not Adodc1.Recordset.EOF
      Set mNode = tvwDB.Nodes.Add(1, tvwChild)
      mNode.Text = Adodc1.Recordset.Fields("科目名称")
'      mNode.Tag = "Publisher" '标识表。
      mNode.Key = Adodc1.Recordset.Fields("科目编号")
     ' mNode.Image = "closed"
      intIndex = mNode.Index
      '对这条记录，使用查询创建 Title 表的记录集，
      '查询条件是所有包含相同 PubID 的记录。对结果记录集中
      '的每一条记录，在 TreeView 控件中加入一个 Node 对象，
      '并用记录的 Title、 ISBN 和 Author 字段为新
      'Node 对象的属性赋值。
Adodc3.RecordSource = "select * from CWMC Where left(科目编号,4)='" & Adodc1.Recordset.Fields("科目编号") & "' AND LEN(科目编号)>4  and 科目类型='" & Combo1.Text & "' ORDER BY 科目名称"
Adodc3.Refresh
      
      Do While Not Adodc3.Recordset.EOF
         Set mNode = tvwDB.Nodes. _
         Add(intIndex, tvwChild)
         mNode.Text = Adodc3.Recordset.Fields("科目名称")  '文本。
         mNode.Key = Adodc3.Recordset.Fields("科目编号")      '唯一的 ID。
         mNode.Tag = "Authors"      '表名。
       '  mNode.Image = "smlBook"      '图象。
         '移动到 rsTitles 中的下一个记录。
         Adodc3.Recordset.MoveNext
      Loop
      '移动到下一个 Publishers 记录。
      Adodc1.Recordset.MoveNext
   Loop

  
End Sub

Private Sub Command2_Click()
If KMMC = 0 Then Exit Sub

If KMMC = 1 Then
Adodc2.RecordSource = "select * from cwmc where 科目名称='" & Text3.Text & "'"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
MsgBox ("会计科目设置有误")
Exit Sub
Else
Formw111.Text1(6).Text = Adodc2.Recordset.Fields("科目方向")
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
Formw1135.DataCombo2(KMBL).Text = Text3.Text
Else
Formw1135.DataCombo2(KMBL).Text = Text3.Text
Formw1135.DataCombo3(KMBL).Text = Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 4 Then
If Text1.Text = "" Then
Formw1135.DataCombo4(KMBL).Text = Text3.Text
Else
Formw1135.DataCombo4(KMBL).Text = Text3.Text
Formw1135.DataCombo5(KMBL).Text = Text1.Text
End If
KMMC = 0
Unload Me
End If


If KMMC = 2 Then
If Text1.Text = "" Then
Formw1135.DataCombo2(KMBL).Text = Text3.Text
Else
Formw1135.DataCombo2(KMBL).Text = Text3.Text
Formw1135.DataCombo3(KMBL).Text = Text1.Text
End If
KMMC = 0
Unload Me
End If



If KMMC = 7 Then
Formw113.Text1(1).Text = Text3.Text
KMMC = 0
Unload Me
End If

If KMMC = 8 Then
Formw116.Text1(1).Text = Text3.Text
KMMC = 0
Unload Me
End If

If KMMC = 9 Then
If Text1.Text = "" Then
Formw8.DataCombo1(4).Text = Text3.Text
Else
Formw8.DataCombo1(4).Text = Text3.Text + "-" + Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 10 Then
If Text1.Text = "" Then
Formw52.Text1(2).Text = Text3.Text
Else
Formw52.Text1(2).Text = Text3.Text
Formw52.Text1(3).Text = Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 11 Then
If Text1.Text = "" Then
Formw52.Text1(4).Text = Text3.Text
Else
Formw52.Text1(4).Text = Text3.Text
Formw52.Text1(5).Text = Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 12 Then
If Text1.Text = "" Then
Formw58.Text1(2).Text = Text3.Text
Else
Formw58.Text1(2).Text = Text3.Text
Formw58.Text1(3).Text = Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 13 Then
If Text1.Text = "" Then
Formw58.Text1(4).Text = Text3.Text
Else
Formw58.Text1(4).Text = Text3.Text
Formw58.Text1(5).Text = Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 14 Then
If Text1.Text = "" Then
Formw8.DataCombo1(0).Text = Text3.Text
Else
Formw8.DataCombo1(0).Text = Text3.Text + "-" + Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 15 Then
If Text1.Text = "" Then
Formw22.DataCombo1(31).Text = Text3.Text
Else
Formw22.DataCombo1(31).Text = Text3.Text + "-" + Text1.Text
End If
KMMC = 0
Unload Me
End If

If KMMC = 16 Then
If Text1.Text = "" Then
Formw22.DataCombo1(33).Text = Text3.Text
Else
Formw22.DataCombo1(33).Text = Text3.Text + "-" + Text1.Text
End If
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
Combo1.Text = "全部"
Text4 = ""
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT * FROM CWMC"
Adodc2.Refresh

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

   tvwDB.Sorted = True
   Set mNode = tvwDB.Nodes.Add()
   mNode.Text = "会计科目"
   mNode.Tag = "会计科目"   '设置 Tag 属性。
  ' mNode.Image = "closed"         '设置 Image
   
Adodc1.RecordSource = "SELECT * FROM CWMC WHERE LEN(科目编号)=4 ORDER BY 科目编号"
Adodc1.Refresh

   Do While Not Adodc1.Recordset.EOF
      Set mNode = tvwDB.Nodes.Add(1, tvwChild)
      mNode.Text = Adodc1.Recordset.Fields("科目名称")
'      mNode.Tag = "Publisher" '标识表。
      mNode.Key = Adodc1.Recordset.Fields("科目编号")
     ' mNode.Image = "closed"
      intIndex = mNode.Index
      '对这条记录，使用查询创建 Title 表的记录集，
      '查询条件是所有包含相同 PubID 的记录。对结果记录集中
      '的每一条记录，在 TreeView 控件中加入一个 Node 对象，
      '并用记录的 Title、 ISBN 和 Author 字段为新
      'Node 对象的属性赋值。
Adodc3.RecordSource = "select * from CWMC Where left(科目编号,4)='" & Adodc1.Recordset.Fields("科目编号") & "' AND LEN(科目编号)>4 ORDER BY 科目名称"
Adodc3.Refresh
      
      Do While Not Adodc3.Recordset.EOF
         Set mNode = tvwDB.Nodes. _
         Add(intIndex, tvwChild)
         mNode.Text = Adodc3.Recordset.Fields("科目名称")  '文本。
         mNode.Key = Adodc3.Recordset.Fields("科目编号")      '唯一的 ID。
         mNode.Tag = "Authors"      '表名。
       '  mNode.Image = "smlBook"      '图象。
         '移动到 rsTitles 中的下一个记录。
         Adodc3.Recordset.MoveNext
      Loop
      '移动到下一个 Publishers 记录。
      Adodc1.Recordset.MoveNext
   Loop
Text3.Text = ""
Text2.Text = ""
Text1.Text = ""
End Sub


Private Sub Text4_Change()
On Error Resume Next
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
   tvwDB.Sorted = True
   tvwDB.Nodes.Clear
   
   Set mNode = tvwDB.Nodes.Add()
   mNode.Text = "会计科目"
   mNode.Tag = "会计科目"   '设置 Tag 属性。
  ' mNode.Image = "closed"         '设置 Image
  
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
   
If Combo1.Text = "全部" Then
Adodc1.RecordSource = "SELECT * FROM CWMC WHERE LEN(科目编号)=4 ORDER BY 科目编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM CWMC WHERE LEN(科目编号)=4 and 科目类型='" & Combo1.Text & "' ORDER BY 科目编号"
Adodc1.Refresh
End If

   Do While Not Adodc1.Recordset.EOF
      Set mNode = tvwDB.Nodes.Add(1, tvwChild)
      mNode.Text = Adodc1.Recordset.Fields("科目名称")
'      mNode.Tag = "Publisher" '标识表。
      mNode.Key = Adodc1.Recordset.Fields("科目编号")
     ' mNode.Image = "closed"
      intIndex = mNode.Index
      '对这条记录，使用查询创建 Title 表的记录集，
      '查询条件是所有包含相同 PubID 的记录。对结果记录集中
      '的每一条记录，在 TreeView 控件中加入一个 Node 对象，
      '并用记录的 Title、 ISBN 和 Author 字段为新
      'Node 对象的属性赋值。
Adodc3.RecordSource = "select * from CWMC Where left(科目编号,4)='" & Adodc1.Recordset.Fields("科目编号") & "' AND LEN(科目编号)>4  and 科目类型='" & Combo1.Text & "' and 科目名称 like '%'+'" & Text4 & "'+'%' ORDER BY 科目名称"
Adodc3.Refresh
      
      Do While Not Adodc3.Recordset.EOF
         Set mNode = tvwDB.Nodes. _
         Add(intIndex, tvwChild)
         mNode.Text = Adodc3.Recordset.Fields("科目名称")  '文本。
         mNode.Key = Adodc3.Recordset.Fields("科目编号")      '唯一的 ID。
         mNode.Tag = "Authors"      '表名。
       '  mNode.Image = "smlBook"      '图象。
         '移动到 rsTitles 中的下一个记录。
         Adodc3.Recordset.MoveNext
      Loop
      '移动到下一个 Publishers 记录。
      Adodc1.Recordset.MoveNext
   Loop


End Sub

'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：
Private Sub tvwDB_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
'Text3.Text = ""                                '属性
   Text2.Text = Node.Text
   
If Text2.Text = "会计科目" Then Exit Sub
   Text2.Text = tvwDB.Nodes(Node.Index).Parent.Text

If Text2.Text = "会计科目" Then
Text3.Text = Node.Text
Text1.Text = ""
Else
Text1.Text = Node.Text
End If
End Sub
