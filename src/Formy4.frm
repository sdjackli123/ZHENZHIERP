VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formy4 
   BackColor       =   &H00C0E0FF&
   Caption         =   "大货客户定单信息"
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
         Caption         =   "进行"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "结束"
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
      Caption         =   "总备料表"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
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
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
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
      Caption         =   "备料"
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
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
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
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
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
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
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
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
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
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
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
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
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
      Caption         =   "单号"
      BeginProperty Font 
         Name            =   "宋体"
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
      Left            =   840
      TabIndex        =   11
      Top             =   2040
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
MsgBox ("没有内容，不能打印！")
Exit Sub
End If
Call MXOutDataToExcel(MSFlexGrid3, "订单备料表                     " + "单号：" + Text1.Text + "合约号：  " + HYH + "           工作编号：" + GZBH + " 交期：" + Str(JHRQ) + "      计划日期：" + Str(JHQ) + "     打印日期" + DYRQ)
End Sub




Private Sub Command14_Click()
Data4.Database.Execute "UPDATE DHCLB SET 材料批号='' WHERE 材料批号=NULL AND 单号='" & Text1.Text & "'"
Data4.Database.Execute "update dhclb set 材料批号=TRIM(材料批号) where 单号='" & Text1.Text & "'"
Data4.Database.Execute "UPDATE DHCLB SET 材料规格='' WHERE 材料规格=NULL AND 单号='" & Text1.Text & "'"
Data4.Database.Execute "update dhclb set 材料规格=TRIM(材料规格) where 单号='" & Text1.Text & "'"
Data4.Database.Execute "UPDATE DHCLB SET 材料颜色='' WHERE 材料颜色=NULL AND 单号='" & Text1.Text & "'"
Data4.Database.Execute "update dhclb set 材料颜色=TRIM(材料颜色) where 单号='" & Text1.Text & "'"
'Data4.Database.Execute "delete * from dhclb where 单号='" & text1.Text & "' and trim(材料批号)='A'"
'Data4.RecordSource = "SELECT DHCLB.材料库类,DHCLB.材料名称,DHCLB.材料规格,DHCLB.材料单位,DHCLB.材料颜色,SUM(DHCLB.材料数量) AS 数量 FROM DHCLB WHERE  DHCLB.单号='" & text1.Text & "' GROUP BY DHCLB.材料库类,DHCLB.材料名称,DHCLB.材料规格,DHCLB.材料单位,DHCLB.材料颜色"
'Data4.Refresh
Data4.Database.Execute "UPDATE DHCLB SET 材料批号='' WHERE LEN(TRIM(材料批号))=0 AND 单号='" & Text1.Text & "'"
Data4.Database.Execute "update dhclb set 材料批号=TRIM(材料批号) where 单号='" & Text1.Text & "'"
Data4.RecordSource = "SELECT 单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,Format(SUM(材料数量),'#0.00') AS 数量 FROM DHCLB WHERE  单号='" & Text1.Text & "' GROUP BY 单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号"
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
If MsgBox("确定自动备料吗？，一旦选择自动备料，那么以前生成的将删除，会生成新的备料表", vbYesNo) = vbNo Then Exit Sub
Data2.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Text1.Text & "'"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data2.Recordset.MoveFirst
Data3.Database.Execute "delete * from zdlclb"
Do While Not Data2.Recordset.EOF
Data3.Database.Execute "insert into zdlclb(款号,订单颜色,主辅名称,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量,材料库类) select 款号,订单颜色,主辅名称,材料名称,材料规格,材料单位,材料颜色,材料批号,sum(材料数量),材料库类 from dlclb where 款号='" & Data2.Recordset.Fields(0) & "' group by 款号,订单颜色,主辅名称,材料名称,材料规格,材料单位,材料颜色,材料批号,材料库类"
Data2.Recordset.MoveNext
Loop
End If

Data2.RecordSource = "select 款号,颜色,尺码,数量 from cmb where 单号='" & Text1.Text & "' order by 款号"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data7.Database.Execute "DELETE * FROM DHCLB WHERE 单号='" & Text1.Text & "'"
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Data7.RecordSource = "select * from zdlclb where 款号='" & Data2.Recordset.Fields(0) & "' and 订单颜色='" & Data2.Recordset.Fields(1) & "' and 主辅名称='" & Data2.Recordset.Fields(2) & "'"
Data7.Refresh
If Data7.Recordset.EOF Then
MsgBox ("款号" + Data2.Recordset.Fields(0) + "颜色" + Data2.Recordset.Fields(1) + "尺码" + Data2.Recordset.Fields(2) + "没有单耗")
Exit Sub
End If

Data3.Database.Execute "insert into dhclb(单号,款号,订单颜色,主辅名称,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量,材料库类) select '" & Text1.Text & "', 款号,订单颜色,主辅名称,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量*val('" & Data2.Recordset.Fields(3) & "'),材料库类 from zdlclb where 款号='" & Data2.Recordset.Fields(0) & "' and 订单颜色='" & Data2.Recordset.Fields(1) & "' and 主辅名称='" & Data2.Recordset.Fields(2) & "'"

Data2.Recordset.MoveNext
Loop
End If

Data7.Database.Execute "INSERT INTO DHCLBY(单号,款号,订单颜色,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量) SELECT 单号,款号,订单颜色,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,Format(SUM(材料数量),'#0.00') AS SL FROM DHCLB WHERE 单号='" & Text1.Text & "' GROUP BY 单号,款号,订单颜色,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号"
Data7.Database.Execute "UPDATE DHCLBY SET 材料单位='个',材料数量=材料数量/1500 WHERE 材料库类='2辅料库' AND 材料名称='粗缝纫线' AND 材料单位='米'"
Data7.Database.Execute "UPDATE DHCLBY SET 材料单位='个',材料数量=材料数量/2700 WHERE 材料库类='2辅料库' AND 材料名称='缝纫线' AND 材料单位='米'"
Data7.Database.Execute "UPDATE DHCLBY SET 材料数量=INT(材料数量)+1 WHERE 材料库类='2辅料库'AND INT(材料数量/2)<>材料数量/2 and 材料单位='个' and (材料名称='缝纫线' or 材料名称='粗缝纫线')"
Data7.Database.Execute "DELETE * FROM DHCLB WHERE 单号='" & Text1.Text & "'"
Data7.Database.Execute "INSERT INTO DHCLB(单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量) SELECT 单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,Format(SUM(材料数量),'#0.00') AS SL FROM DHCLBY WHERE 单号='" & Text1.Text & "' GROUP BY 单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号"
Data7.Database.Execute "DELETE * FROM DHCLBY WHERE 单号='" & Text1.Text & "'"

Data4.RecordSource = "SELECT 单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量 FROM DHCLB WHERE DHCLB.单号='" & Text1.Text & "'"
Data4.Refresh

End Sub



Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
DTPicker1.Value = Date - 15
DTPicker2.Value = Date
Option1.Value = True
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select 简称 from khZL group by 简称"
Data1.Refresh
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

Data4.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data4.RecordSource = "SELECT * FROM DHCLB WHERE DHCLB.单号='" & Text1.Text & "'"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data6.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

Data7.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"


MSFlexGrid3.ColWidth(0) = 200
MSFlexGrid3.ColWidth(1) = 1500

End Sub


Private Sub MSFlex()
With MSFlexGrid3
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
    Data5.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 备料='进行'"
    Data5.Refresh
    m = 1
    If Not Data5.Recordset.EOF Then  'make sure there are records in the table
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data5.Recordset.Fields(0)
        intIndex = mNode.Index
        Data6.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data5.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 备料='进行'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data7.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data6.Recordset.Fields(0) & "' and 备料='进行'"
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
    Data5.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 备料='结束'"
    Data5.Refresh
    m = 1
    If Not Data5.Recordset.EOF Then  'make sure there are records in the table
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data5.Recordset.Fields(0)
        intIndex = mNode.Index
        Data6.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data5.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 备料='结束'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data7.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data6.Recordset.Fields(0) & "' and 备料='结束'"
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
Text1.Text = l1
Data4.RecordSource = "SELECT 单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量 FROM DHCLB WHERE DHCLB.单号='" & Text1.Text & "'"
Data4.Refresh
End If

End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub


