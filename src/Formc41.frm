VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formc41 
   BackColor       =   &H00C0E0FF&
   Caption         =   "单号领料查询"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form41"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data7 
      Caption         =   "Data3"
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data6 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data5 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   2775
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "进行"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command8 
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
      TabIndex        =   14
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data15 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细打印"
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "零价调整"
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "金额刷新"
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4560
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   3000
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   11400
      Top             =   120
   End
   Begin VB.CommandButton Command4 
      Caption         =   "刷新库存"
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按单号详细查询"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按单号总体查询"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc41.frx":0000
      Height          =   7455
      Left            =   3600
      TabIndex        =   4
      Top             =   2040
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   360
      Left            =   4680
      TabIndex        =   0
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formc41.frx":0014
      Height          =   360
      Left            =   4680
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5655
      Left            =   360
      TabIndex        =   18
      Top             =   3720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9975
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81199105
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81199105
      CurrentDate     =   39557
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
      Index           =   14
      Left            =   240
      TabIndex        =   22
      Top             =   1440
      Width           =   1215
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
      Index           =   15
      Left            =   240
      TabIndex        =   21
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择单号"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择库类"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Formc41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2 As String: Public BAR, r, c As Integer

Private Sub Command1_Click()
End Sub

Private Sub Command10_Click()
Data2.RecordSource = "SELECT * FROM KPD WHERE 单号='" & DBCombo1.Text & "' AND (单价=0 OR 单价=NULL)"
Data2.Refresh
End Sub

Private Sub Command11_Click()
Call OutDataToExcel(MSFlexGrid1, 10, "单号：" + DBCombo1.Text + "备料表出库明细")
End Sub

Private Sub Command2_Click()
If DBCombo2.Text = "" Then
Data2.RecordSource = "select 单号,库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,供应单位,日期,单据号 from KPD WHERE 单号='" & DBCombo1.Text & "'  ORDER BY 库类"
Data2.Refresh
Else
Data2.RecordSource = "select 单号,库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,供应单位,日期,单据号 from KPD WHERE 单号='" & DBCombo1.Text & "' and 库类='" & DBCombo2.Text & "' ORDER BY 库类"
Data2.Refresh
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Data2.Database.Execute "DELETE * FROM CLRCZZ"
Data2.Database.Execute "DELETE * FROM CLRCZZHZ"
Data2.Database.Execute "INSERT INTO CLRCZZ(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类) select CKGL.材料名称,CKGL.材料规格,CKGL.材料单位,CKGL.颜色,CKGL.批次,CKGL.数量,CKGL.单价,CKGL.库类 from ckgl WHERE CKGL.库别='清库库存' "
Data2.Database.Execute "UPDATE CLRCZZ SET 库别='入库' where 库别=NULL"
Data2.Database.Execute "INSERT INTO CLRCZZ(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类) select CKBL.材料名称,CKBL.材料规格,CKBL.材料单位,CKBL.颜色,CKBL.批次,CKBL.数量,CKBL.单价,CKBL.库类 from ckBL WHERE CKBL.库别='清库库存' "
Data2.Database.Execute "UPDATE CLRCZZ SET 库别='出库',数量=-数量 WHERE 库别=NULL"
Data2.Database.Execute "INSERT INTO CLRCZZHZ(库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价) SELECT CLRCZZ.库类,CLRCZZ.材料名称,CLRCZZ.材料规格,CLRCZZ.材料单位,CLRCZZ.颜色,CLRCZZ.批次,SUM(CLRCZZ.数量) AS L,AVG(CLRCZZ.单价) AS D FROM CLRCZZ GROUP BY CLRCZZ.库类,CLRCZZ.材料名称,CLRCZZ.材料规格,CLRCZZ.材料单位,CLRCZZ.颜色,CLRCZZ.批次"

End Sub

Private Sub Command5_Click()
If DBCombo2.Text = "" Then
Data2.RecordSource = "select 客户,库类,材料名称,材料规格,材料单位,颜色,SUM(数量) AS 累计数量 from KPD WHERE 单号='" & DBCombo1.Text & "'  GROUP BY 客户,库类,材料名称,材料规格,材料单位,颜色 ORDER BY 库类"
Data2.Refresh
Else
Data2.RecordSource = "select 客户,库类,材料名称,材料规格,材料单位,颜色,SUM(数量) AS 累计数量 from KPD WHERE 单号='" & DBCombo1.Text & "' and 库类='" & DBCombo2.Text & "' GROUP BY 客户,库类,材料名称,材料规格,材料单位,颜色 ORDER BY 库类"
Data2.Refresh
End If
Call SX2(Data2, MSFlexGrid1, 7)
End Sub




Private Sub Command8_Click()
Call tree
Call zk
End Sub

Private Sub Command9_Click()
On Error Resume Next
Data2.RecordSource = "SELECT * FROM KPD WHERE 单号='" & DBCombo1.Text & "' AND (单价=0 OR 单价=NULL)"
Data2.Refresh
If Not Data2.Recordset.EOF Then
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Data4.RecordSource = "select * from ckgl where  单号='" & Data2.Recordset.Fields(1) & "' and 材料名称='" & Data2.Recordset.Fields(3) & "'  and 颜色='" & Data2.Recordset.Fields(6) & "'  and 材料规格='" & Data2.Recordset.Fields(4) & "' and 批次='" & Data2.Recordset.Fields(7) & "'"
Data4.Refresh
If Data4.Recordset.EOF Then
Else
Data2.Recordset.Edit
Data2.Recordset.Fields(9) = Data4.Recordset.Fields(9)
Data2.Recordset.Update
End If
Data2.Recordset.MoveNext
Loop
End If
Data2.Database.Execute "UPDATE KPD SET 合计金额=Format(数量 * 单价, '#0.00') WHERE 单号='" & DBCombo1.Text & "'"
Data2.Refresh
End Sub

Private Sub DBCombo4_Click(Area As Integer)
Data2.RecordSource = "select * from KPD WHERE 单据号='" & DBCombo4.Text & "' "
Data2.Refresh
'Call SX2(Data2, MSFlexGrid1, 9)
End Sub

Private Sub DBCombo1_Change()
Data2.DatabaseName = "d:\数据库\\htgl\2011\CKGL.mdb"
If DBCombo2.Text = "" Then
Data2.RecordSource = "select 单号,库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,供应单位,日期,单据号 from KPD WHERE 单号='" & DBCombo1.Text & "'  ORDER BY 库类"
Data2.Refresh
Else
Data2.RecordSource = "select 单号,库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,供应单位,日期,单据号 from KPD WHERE 单号='" & DBCombo1.Text & "' and 库类='" & DBCombo2.Text & "' ORDER BY 库类"
Data2.Refresh
End If
End Sub

Private Sub DTPicker3_Change()
Text3.Text = Month(DTPicker3.Value)
End Sub

Private Sub DTPicker3_CloseUp()
Text3.Text = Month(DTPicker3.Value)
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DBCombo2.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data1.RecordSource = "select KL.MC FROM KL GROUP BY KL.MC"
Data1.Refresh
ProgressBar1.Visible = False
Timer1.Enabled = False
DTPicker3.Value = Date - 30
DTPicker4.Value = Date
Option4.Value = True
Data3.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data4.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"

Data5.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data6.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data7.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"

Data2.DatabaseName = "d:\数据库\\htgl\2011\CKGL.mdb"
Data2.Refresh
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(12) = 0
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 1
       khbl = 4
Form202.Text1.Text = DBCombo1.Text
Form202.Show

End Select

End Sub


Private Sub Timer1_Timer()
If BAR = 100 Then
DataEnvironment1.Command3 DBCombo4.Text
DataReport9.Show 1
DataEnvironment1.rsCommand3.Close
Timer1.Enabled = False
ProgressBar1.Visible = False
Exit Sub
End If
BAR = BAR + 1
ProgressBar1.Value = BAR


End Sub

Private Sub Timer2_Timer()
End Sub
Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid1.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid1.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data2.Recordset.MoveFirst
Data2.Recordset.Move r - 1
Data2.Recordset.Edit
Data2.Recordset.Fields(c - 1) = Text1111.Text
Data2.Recordset.Update
Text1111.Visible = False
MSFlexGrid1.SetFocus
End Sub


Private Sub MSF()
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

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSF
End If
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
    Data7.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
    Data7.Refresh
    m = 1
    If Not Data7.Recordset.EOF Then  'make sure there are records in the table
        Data7.Recordset.MoveFirst
        Do While Not Data7.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data7.Recordset.Fields(0)
        intIndex = mNode.Index
        Data5.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data7.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
        Data5.Refresh
        
        If Not Data5.Recordset.EOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data5.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data6.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data5.Recordset.Fields(0) & "' and 进度='进行'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        Data6.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        
        Data5.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        Data7.Recordset.MoveNext
        m = m + 1
        Loop
    End If
End If


If Option5.Value = True Then
    Data7.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='结束'"
    Data7.Refresh
    m = 1
    If Not Data7.Recordset.EOF Then  'make sure there are records in the table
        Data7.Recordset.MoveFirst
        Do While Not Data7.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(, , Data7.Recordset.Fields(0), Data7.Recordset.Fields(0))
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data7.Recordset.Fields(0)
        intIndex = mNode.Index
        Data5.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data7.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='结束'"
        Data5.Refresh
        
        If Not Data5.Recordset.EOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data5.Recordset.Fields(0))
        intIndex = mNode.Index
        Data6.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data5.Recordset.Fields(0) & "' and 进度='结束'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        Data6.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data5.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data7.Recordset.MoveNext
        Loop
    End If
End If

End Sub



