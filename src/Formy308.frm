VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formy308 
   BackColor       =   &H00C0E0FF&
   Caption         =   "订单采购"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "采购进度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   2175
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
      Width           =   6855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出本操作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查看采购表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   2175
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
      Width           =   7095
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Text            =   "Text1111"
      Top             =   10320
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command9 
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39177
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Left            =   5160
      TabIndex        =   6
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy308.frx":0000
      Height          =   7935
      Left            =   4080
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13996
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "记录号 "
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7455
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   13150
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label6 
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
      Index           =   0
      Left            =   600
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      Index           =   0
      Left            =   600
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
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
      Index           =   1
      Left            =   4080
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Formy308"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2, M1, M2, M3, M4, M5, C1, C2, C3, C4, C5, C6, c7 As String: Public c, r, S1, S2 As Integer




Private Sub Command1_Click()
On Error Resume Next
Data1.Database.Execute "DELETE * FROM CKGL"
Data5.Database.Execute "INSERT INTO CKGL(单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'd:\数据库\\htgl\2011\SCZYJHD.MDB' SELECT 单号,合约号,库类,材料名称,材料规格,材料单位,颜色,批次,SUM(数量) FROM CKGL WHERE 单号='" & DBCombo1.Text & "'  GROUP BY 单号,合约号,库类,材料名称,材料规格,材料单位,颜色,批次 "
Data1.Database.Execute "UPDATE CKGL SET LX=CK,采购数量=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,采购数量) SELECT 单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,SUM(材料数量) AS 采购数量 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "'  GROUP BY 单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号"
Data1.Database.Execute "UPDATE CKGL SET LX=CK,入库数量=0 WHERE LX=NULL"
Data2.RecordSource = "SELECT 单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,format(SUM(采购数量),'#0.00') AS 采购量,format(SUM(入库数量),'#0.00') AS 入库量 FROM CKGL WHERE  单号='" & DBCombo1.Text & "' GROUP BY 单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号"
Data2.Refresh
End Sub

Private Sub Command4_Click()
Data2.RecordSource = "SELECT 单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量 AS 采购量 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "'  ORDER BY 单号,材料库类,材料名称,材料规格,材料颜色"
Data2.Refresh
End Sub

Private Sub Command5_Click()
Unload Me
End Sub



Private Sub Command9_Click()
Call tree
Call zk
End Sub

Private Sub DBCombo1_Change()
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data2.RecordSource = "SELECT 单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量 AS 采购量 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "'  ORDER BY 单号,材料库类,材料名称,材料规格,材料颜色"
Data2.Refresh
End Sub

Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text4.Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Change()
Text5.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text5.Text = DTPicker2.Value
End Sub


Private Sub Form_Load()
DTPicker1.Value = Date - 15
DTPicker2.Value = Date
DBCombo1.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data4.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data5.DatabaseName = "d:\数据库\\htgl\2011\ckgl.MDB"
Data7.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data8.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
MSFlexGrid2.ColWidth(0) = 200
MSFlexGrid2.ColWidth(1) = 1500
End Sub


Private Sub MSFlexGrid1_dblClick()
rs = MSFlexGrid1.Row
If Data1.Recordset.EOF Then
DBCombo1.Text = ""
Exit Sub
End If

Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
DBCombo1.Text = Data1.Recordset.Fields(7)
End Sub

Private Sub MSFlexGrid2_Click()
On Error Resume Next
rs = MSFlexGrid2.Row
'If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
C1 = Data2.Recordset.Fields(0)
C2 = Data2.Recordset.Fields(1)
C3 = Data2.Recordset.Fields(2)
C4 = Data2.Recordset.Fields(3)
C5 = Data2.Recordset.Fields(4)
C6 = Data2.Recordset.Fields(5)
c7 = Data2.Recordset.Fields(6)
End Sub

Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid2.RowSel
End Sub

Private Sub MSFlexGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid2.RowSel
End Sub
Private Sub MSFlexGrid3_DBLClick()
On Error Resume Next
rs = MSFlexGrid3.Row
Data4.Recordset.MoveFirst
Data4.Recordset.Move rs - 1
Formy52.DBCombo1(12).Text = Data4.Recordset.Fields(7)
Formy52.DBCombo1(3).Text = Data4.Recordset.Fields(0)
Formy52.DBCombo2.Text = Data4.Recordset.Fields(3)
Formy52.DBCombo1(1).Text = DBCombo1.Text
End Sub

Private Sub MSFlexGrid2_dblClick()
With MSFlexGrid2
    c = .Col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlexGrid2_dblClick
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid2.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid2.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid2.SetFocus
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
End Sub


Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex

   TreeView1.Nodes.Clear
 
    Data3.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
    Data3.Refresh
    m = 1
    If Not Data3.Recordset.EOF Then  'make sure there are records in the table
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data3.Recordset.Fields(0)
        intIndex = mNode.Index
        Data4.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data3.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
        Data4.Refresh
        
        If Not Data4.Recordset.EOF Then
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data4.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data7.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data4.Recordset.Fields(0) & "' and 进度='进行'"
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
        Data4.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data3.Recordset.MoveNext
        Loop
    End If

End Sub


'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next

If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") > 0 Then
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
DBCombo1.Text = l1
End If

End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub


