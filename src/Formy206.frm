VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formy206 
   BackColor       =   &H00C0E0FF&
   Caption         =   "客户订单查询"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data14 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data13 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data12 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
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
      TabIndex        =   6
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9810
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   3735
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1095
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "状态"
      Height          =   855
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000C0C0&
         Caption         =   "全部"
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C0C0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "进行"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81395713
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81395713
      CurrentDate     =   36892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy206.frx":0000
      Height          =   7455
      Left            =   3360
      TabIndex        =   9
      Top             =   1560
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
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6975
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   12303
      _Version        =   393217
      Style           =   7
      Appearance      =   1
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
      TabIndex        =   12
      Top             =   480
      Width           =   1095
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
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Formy206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call JHJD(MSFlexGrid1, "生产订单")
End Sub

Private Sub Command7_Click()
Call tree
Call zk
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data12.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data13.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data14.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Option3.Value = True
MSFlexGrid1.ColWidth(0) = 200
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
    Data12.RecordSource = "select distinct 客户 from sczy_xdd where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data13.RecordSource = "select distinct 单号 from sczy_xdd where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
        Data13.Refresh
        
        If Not Data13.Recordset.EOF Then
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data13.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data14.RecordSource = "select distinct 款号 from sczy_xdd where 单号='" & Data13.Recordset.Fields(0) & "' and 进度='进行'"
        Data14.Refresh
        
        If Not Data14.Recordset.EOF Then
        Data14.Recordset.MoveFirst
        Do While Not Data14.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data14.Recordset.Fields(0))
        Data14.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data13.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data12.Recordset.MoveNext
        Loop
    End If
End If

If Option3.Value = True Then
    Data12.RecordSource = "select distinct 客户 from sczy_xdd where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data13.RecordSource = "select distinct 单号 from sczy_xdd where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
        Data13.Refresh
        
        If Not Data13.Recordset.EOF Then
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data13.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data14.RecordSource = "select distinct 款号 from sczy_xdd where 单号='" & Data13.Recordset.Fields(0) & "'"
        Data14.Refresh
        
        If Not Data14.Recordset.EOF Then
        Data14.Recordset.MoveFirst
        Do While Not Data14.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data14.Recordset.Fields(0))
        Data14.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data13.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data12.Recordset.MoveNext
        Loop
    End If
End If

If Option2.Value = True Then
    Data12.RecordSource = "select distinct 客户 from sczy_xdd where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "x" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data13.RecordSource = "select distinct 单号 from sczy_xdd where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
        Data13.Refresh
        
        If Not Data13.Recordset.EOF Then
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "x" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data13.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data14.RecordSource = "select distinct 款号 from sczy_xdd where 单号='" & Data13.Recordset.Fields(0) & "' and 进度='结束'"
        Data14.Refresh
        
        If Not Data14.Recordset.EOF Then
        Data14.Recordset.MoveFirst
        Do While Not Data14.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "x" + Trim(m) + "w" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data14.Recordset.Fields(0))
        m = m + 1
        Data14.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data13.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data12.Recordset.MoveNext
        Loop
    End If
End If

End Sub


Private Sub MSFlexGrid1_dblClick()
If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data2.Recordset.Move rs - 1
If yhmk = "销售" Then
Formy201.DBCombo1.Text = Data2.Recordset.Fields(27)
Formy201.DBCombo2.Text = Data2.Recordset.Fields(0)
Formy201.Show
End If
End Sub

'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next



If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") = 0 Then
Data2.RecordSource = "select * from sczy_xdd where 客户='" & TreeView1.Nodes(Node.Index).FullPath & "'"
Data2.Refresh
Else

l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
Data2.RecordSource = "select * from sczy_xdd where 单号='" & l1 & "'"
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







