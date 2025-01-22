VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formy306 
   BackColor       =   &H00C0E0FF&
   Caption         =   "样品进度"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form32"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data5 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6495
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   11456
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Data Data4 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data3 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "状态"
      Height          =   975
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   3855
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "进行"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C0C0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000C0C0&
         Caption         =   "全部"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   36892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy306.frx":0000
      Height          =   7455
      Left            =   3360
      TabIndex        =   9
      Top             =   1680
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   8421631
      BackColorBkg    =   4109501
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Formy306"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call JHJD(MSFlexGrid1, "样品进度")
End Sub
Private Sub Command6_Click()
Call tree
Call zk
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select 简称 from khZL group by 简称"
Data1.Refresh
Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
Data5.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
MSFlexGrid1.ColWidth(0) = 200
Option3.Value = True
For i = 10 To 25
MSFlexGrid1.ColWidth(i) = 1500
Next
End Sub

Private Sub MSFlex()
On Error Resume Next
With MSFlexGrid1
    c = .Col: r = .Row    '''''C列，，R行
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 2, 3
khbl = 5
Formy202.Show
End Select
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data2.Recordset.MoveFirst
Data2.Recordset.Move r - 1
Data2.Recordset.Edit
Data2.Recordset.Fields(c - 1) = Text1111.Text
Data2.Recordset.Update
Text1111.Visible = False
MSFlexGrid1.Text = Text1111.Text
MSFlexGrid1.SetFocus
End If
End Sub


Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex
   TreeView1.Nodes.Clear
 
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox ("请选择生产状态")
Exit Sub
End If

If Option1.Value = True Then
    Data3.RecordSource = "select distinct 客户 from ypjd where 接单日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
    Data3.Refresh
    m = 1
    If Not Data3.Recordset.EOF Then  'make sure there are records in the table
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data3.Recordset.Fields(0)
        intIndex = mNode.Index
        Data4.RecordSource = "select distinct 单号 from ypjd where 客户='" & Data3.Recordset.Fields(0) & "' and  接单日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
        Data4.Refresh
        
        If Not Data4.Recordset.EOF Then
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "t" + Trim(intIndex)
        mNode.Text = Trim(Data4.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data5.RecordSource = "select distinct 款号 from ypjd where 单号='" & Data4.Recordset.Fields(0) & "' and 进度='进行'"
        Data5.Refresh
        
        If Not Data5.Recordset.EOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "t" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data5.Recordset.Fields(0))
        Data5.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data4.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data3.Recordset.MoveNext
        Loop
    End If
End If

If Option3.Value = True Then
    Data3.RecordSource = "select distinct 客户 from ypjd where 接单日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
    Data3.Refresh
    m = 1
    If Not Data3.Recordset.EOF Then  'make sure there are records in the table
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "x" + Trim(m)
        mNode.Text = Data3.Recordset.Fields(0)
        intIndex = mNode.Index
        Data4.RecordSource = "select distinct 单号 from ypjd where 客户='" & Data3.Recordset.Fields(0) & "' and  接单日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
        Data4.Refresh
        
        If Not Data4.Recordset.EOF Then
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "x" + Trim(m) + "t" + Trim(intIndex)
        mNode.Text = Trim(Data4.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data5.RecordSource = "select distinct 款号 from ypjd where 单号='" & Data4.Recordset.Fields(0) & "'"
        Data5.Refresh
        
        If Not Data5.Recordset.EOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "x" + Trim(m) + "t" + Trim(intIndex) + "w" + Trim(xntIndex)
        mNode.Text = Trim(Data5.Recordset.Fields(0))
        Data5.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data4.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data3.Recordset.MoveNext
        Loop
    End If
End If

If Option2.Value = True Then
    Data3.RecordSource = "select distinct 客户 from ypjd where 接单日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
    Data3.Refresh
    m = 1
    If Not Data3.Recordset.EOF Then  'make sure there are records in the table
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data3.Recordset.Fields(0)
        intIndex = mNode.Index
        Data4.RecordSource = "select distinct 单号 from ypjd where 客户='" & Data3.Recordset.Fields(0) & "' and  接单日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
        Data4.Refresh
        
        If Not Data4.Recordset.EOF Then
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data4.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data5.RecordSource = "select distinct 款号 from ypjd where 单号='" & Data4.Recordset.Fields(0) & "' and 进度='结束'"
        Data5.Refresh
        
        If Not Data5.Recordset.EOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data5.Recordset.Fields(0))
        Data5.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data4.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data3.Recordset.MoveNext
        Loop
    End If
End If

End Sub


'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") = 0 Then
Data2.RecordSource = "select * FROM ypjd WHERE instr(客户,'" & TreeView1.Nodes(Node.Index).FullPath & "')>0  and 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') order by 单号,款号"
Data2.Refresh
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
Data2.RecordSource = "select * FROM ypjd WHERE instr(单号,'" & l1 & "')>0 order by 单号,款号"
Data2.Refresh
End If

'DBCombo2.Text = Node.Index
'DBCombo3.Text = TreeView1.Nodes(Node.Index).FullPath
End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub


