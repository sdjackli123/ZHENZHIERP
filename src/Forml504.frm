VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Forml504 
   BackColor       =   &H00C0E0FF&
   Caption         =   "漂染信息"
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
      Caption         =   "金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   3960
      Style           =   1  'Simple Combo
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
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
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
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
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
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
      Height          =   735
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "进行"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Data Data10 
      Caption         =   "Data4"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data11 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data12 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   3600
      TabIndex        =   7
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Forml504.frx":0000
      Height          =   8295
      Left            =   3600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   14631
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6735
      Left            =   360
      TabIndex        =   9
      Top             =   2880
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11880
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80674817
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80674817
      CurrentDate     =   39557
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择单号"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
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
      Index           =   27
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
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
      Index           =   28
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Forml504"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer
Private Sub Command1_Click()
Data4.RecordSource = "select * from rsrk where  单号='" & DBCombo1.Text & "' order by 日期,染色单位,材料名称"
Data4.Refresh
End Sub

Private Sub Command10_Click()
Call tree
Call zk
End Sub

Private Sub Command2_Click()
If DBCombo1.Text = "" Then Exit Sub
Data5.Database.Execute "update rsrk set 金额=format(val(毛坯重量)*val(单价),'#0.00'),织布金额=format(val(毛坯重量)*val(织布单价),'#0.00'),印花金额=format(val(印花单价)*val(印花数量),'#0.00') where 单号='" & DBCombo1.Text & "'"
MsgBox ("金额已刷新")
Data4.RecordSource = "select * from rsrk where  单号='" & DBCombo1.Text & "' order by 日期,染色单位,材料名称"
Data4.Refresh
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DTPicker3.Value = Date - 30
DTPicker4.Value = Date
Option4.Value = True
Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data5.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"

Data10.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data11.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data12.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(4) = 1600
MSFlexGrid1.ColWidth(12) = 1300

For i = 4 To 7
MSFlexGrid1.ColWidth(i) = 0
Next

MSFlexGrid1.ColWidth(10) = 0



End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Data1.DatabaseName = "e:\excel\sjzz.MDB"
If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") = 0 Then
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
DBCombo1.Text = l1
End If


'DBCombo2.Text = Node.Index
'DBCombo3.Text = TreeView1.Nodes(Node.Index).FullPath
End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub


Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex
   TreeView1.Nodes.Clear
 

If Option4.Value = True Then
    Data12.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data10.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
        Data10.Refresh
        
        If Not Data10.Recordset.EOF Then
        Data10.Recordset.MoveFirst
        Do While Not Data10.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data10.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data11.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data10.Recordset.Fields(0) & "' and 进度='进行'"
        Data11.Refresh
        
        If Not Data11.Recordset.EOF Then
        Data11.Recordset.MoveFirst
        Do While Not Data11.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data11.Recordset.Fields(0))
        Data11.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        
        Data10.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        Data12.Recordset.MoveNext
        m = m + 1
        Loop
    End If
End If


If Option5.Value = True Then
    Data12.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='结束'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(, , Data12.Recordset.Fields(0), Data12.Recordset.Fields(0))
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data10.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='结束'"
        Data10.Refresh
        
        If Not Data10.Recordset.EOF Then
        Data10.Recordset.MoveFirst
        Do While Not Data10.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data10.Recordset.Fields(0))
        intIndex = mNode.Index
        Data11.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data10.Recordset.Fields(0) & "' and 进度='结束'"
        Data11.Refresh
        
        If Not Data11.Recordset.EOF Then
        Data11.Recordset.MoveFirst
        Do While Not Data11.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data11.Recordset.Fields(0))
        Data11.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data10.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data12.Recordset.MoveNext
        Loop
    End If
End If

End Sub


Private Sub MSF()
With MSFlexGrid1
    c = .Col: r = .Row    '''''C列，，R行
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSF
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    MSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data4.Recordset.MoveFirst
Data4.Recordset.Move r - 1
Data4.Recordset.Edit

Data4.Recordset.Fields(c - 1) = Combo1111.Text
Data4.Recordset.Update

MSFlexGrid1.Text = Combo1111.Text
Combo1111.Visible = False
MSFlexGrid1.SetFocus
End If
End Sub



