VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormW209 
   BackColor       =   &H00C0E0FF&
   Caption         =   "退单明细"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data11 
      Caption         =   "Data2"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4920
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7920
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   1320
      Width           =   6015
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成明细"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "编号刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "编号查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      TabIndex        =   5
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "进行"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Data Data5 
      Caption         =   "Data1"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data6 
      Caption         =   "Data2"
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Data Data7 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data8 
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data9 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   6360
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data10 
      Caption         =   "Data2"
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81330177
      CurrentDate     =   39557
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "FormW209.frx":0000
      Height          =   5535
      Left            =   3120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1680
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6255
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11033
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81330177
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81330177
      CurrentDate     =   39557
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "FormW209.frx":0014
      Height          =   2175
      Left            =   3120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7320
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "运费"
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
      Index           =   3
      Left            =   3120
      TabIndex        =   28
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
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
      Index           =   4
      Left            =   6360
      TabIndex        =   27
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发货编号"
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
      Index           =   18
      Left            =   4920
      TabIndex        =   22
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
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
      Index           =   17
      Left            =   3120
      TabIndex        =   21
      Top             =   360
      Width           =   1575
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
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   360
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
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发货单号"
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
      Left            =   6360
      TabIndex        =   18
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FormW209"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public h1, h2 As String
Public c, r As Integer

Private Sub Command1_Click()
Call Thmx(Text2.Text)
End Sub

Private Sub Command2_Click()
Call tree
Call zk
End Sub

Private Sub Command3_Click()
Data1.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data1.RecordSource = "select 客户,单号,编号,款号,规格,单位,价格,颜色1,-val(数量1) as A,颜色2,-VAL(数量2) AS B,颜色3,-VAL(数量3) AS C,颜色4,-VAL(数量4) AS D,颜色5,-VAL(数量5) AS E,-VAL(小计) AS 小计数,提货,-VAL(金额) AS 合计金额,折扣 from zxd where 编号='" & Text2.Text & "' order by 序号"
Data1.Refresh
Data11.RecordSource = "select 编号,-val(费用) as 运费,备注,单号,客户,序号,类别 from zxbz where 编号='" & Text2.Text & "'"
Data11.Refresh
End Sub

Private Sub Command4_Click()
Data9.Database.Execute "update zxd set 小计=val(数量1)+val(数量2)+val(数量3)+val(数量4)+val(数量5),金额=format((val(数量1)+val(数量2)+val(数量3)+val(数量4)+val(数量5))*val(提货),'#0.0') where 编号='" & Text2.Text & "'"
Data1.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data1.RecordSource = "select 客户,单号,编号,款号,规格,单位,价格,颜色1,-val(数量1) as A,颜色2,-VAL(数量2) AS B,颜色3,-VAL(数量3) AS C,颜色4,-VAL(数量4) AS D,颜色5,-VAL(数量5) AS E,-VAL(小计) AS 小计数,提货,-VAL(金额) AS 合计金额,折扣 from zxd where 编号='" & Text2.Text & "' order by 编号,序号"
Data1.Refresh
Data11.RecordSource = "select 编号,-val(费用) as 运费,备注,单号,客户,序号,类别 from zxbz where 编号='" & Text2.Text & "'"
Data11.Refresh
End Sub

Private Sub Command5_Click()
Data1.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data1.RecordSource = "select 客户,单号,编号,款号,规格,单位,价格,颜色1,-val(数量1) as A,颜色2,-VAL(数量2) AS B,颜色3,-VAL(数量3) AS C,颜色4,-VAL(数量4) AS D,颜色5,-VAL(数量5) AS E,-VAL(小计) AS 小计数,提货,-VAL(金额) AS 合计金额,折扣 from zxd where 单号='" & Text1.Text & "' order by 编号,序号"
Data1.Refresh
Data11.RecordSource = "select 编号,-val(费用) as 运费,备注,单号,客户,序号,类别 from zxbz where 单号='" & Text1.Text & "' order by 编号"
Data11.Refresh
End Sub

Private Sub Command6_Click()
If MsgBox("装箱编号为：" + Text2.Text + "添加备注吗？", vbYesNo) = vbNo Then Exit Sub
Data9.Database.Execute "delete * from zxbz where 编号='" & Text2.Text & "'"
Data9.Database.Execute "insert into zxbz(编号,费用,备注,单号,客户,序号,类别) VALUES('" & Text2.Text & "',-val('" & Text3.Text & "'),'" & Text4.Text & "','" & Text1.Text & "','" & h2 & "',1,'应收类')"
MsgBox ("已确定！")
Data11.RecordSource = "select 编号,-val(费用) as 运费,备注,单号,客户,序号,类别 from zxbz where 编号='" & Text2.Text & "'"
Data11.Refresh
End Sub

Private Sub Command7_Click()
If Text1.Text = "" Then
MsgBox ("输入单号")
Exit Sub
End If
Data7.RecordSource = "select * from zxdb where 单号='" & Text1.Text & "' order by 序号"
Data7.Refresh
If Data7.Recordset.EOF Then
MsgBox ("目前此单号没有生成余量表，需要生成与量表才可以生成发货明细")
Exit Sub
End If

Data10.RecordSource = "select 编号 from zxd where 单号='" & Text1.Text & "'"
Data10.Refresh

If Not Data10.Recordset.EOF Then
Data10.Recordset.MoveFirst
Do While Not Data10.Recordset.EOF
Data9.RecordSource = "select * from lsfh where 单据号='" & Data10.Recordset.Fields(0) & "'"
Data9.Refresh
If Data9.Recordset.EOF Then
MsgBox ("以前的发货明细单中有没有从仓库发货的，因此禁止在此开发货单，请核实原因")
Exit Sub
End If
Data10.Recordset.MoveNext
Loop
End If

lo = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data7.Recordset.MoveFirst
Do While Not Data7.Recordset.EOF
For i = 0 To 4
If Val(Data7.Recordset.Fields(i + 8)) > 0 Then
Data2.Database.Execute "insert into zxd(客户,单号,款号,规格,单位,序号,编号) VALUES('" & Data7.Recordset.Fields(0) & "','" & Data7.Recordset.Fields(1) & "','" & Data7.Recordset.Fields(2) & "','" & Data7.Recordset.Fields(3) & "','" & Data7.Recordset.Fields(4) & "','" & Data7.Recordset.Fields(25) & "','" & Text2.Text & "')"
i = 5
End If
Next
Data7.Recordset.MoveNext
Loop

Data2.RecordSource = "select * from zxd where 编号='" & Text2.Text & "' order by 序号"
Data2.Refresh
If Data2.Recordset.EOF Then
MsgBox ("此单号：" + Text1.Text + "发货结束")
Exit Sub
End If

Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Data7.RecordSource = "select * from zxdb where 序号=val('" & Data2.Recordset.Fields(21) & "')"
Data7.Refresh

l1 = 5
l2 = 6
Data2.Recordset.Edit
For i = 0 To 4
If Val(Data7.Recordset.Fields(i * 4 + 8)) > 0 Then
Data2.Recordset.Fields(l1) = Data7.Recordset.Fields(i * 4 + 8 - 3)
Data2.Recordset.Fields(l2) = Data7.Recordset.Fields(i * 4 + 8 - 2)
l1 = l1 + 2
l2 = l2 + 2
End If
Next

Data9.RecordSource = "select 销售,折扣,提货 from KSBJ where 客户='" & Data2.Recordset.Fields(0) & "' and 款号='" & Data2.Recordset.Fields(1) & "' and  规格='" & Data2.Recordset.Fields(2) & "'"
Data9.Refresh
If Data9.Recordset.EOF Then
Data2.Recordset.Fields(4) = ""
Data2.Recordset.Fields(18) = ""
Data2.Recordset.Fields(22) = ""
Else
Data2.Recordset.Fields(4) = Data9.Recordset.Fields(0)
Data2.Recordset.Fields(18) = Data9.Recordset.Fields(2)
Data2.Recordset.Fields(22) = Data9.Recordset.Fields(1)
End If

Data2.Recordset.Update
Data2.Recordset.MoveNext
Loop

Data9.Database.Execute "update zxd set 数量1='' where 编号='" & Text2.Text & "' and 数量1=null"
Data9.Database.Execute "update zxd set 颜色1='' where 编号='" & Text2.Text & "' and 颜色1=null"
Data9.Database.Execute "update zxd set 数量2='' where 编号='" & Text2.Text & "' and 数量2=null"
Data9.Database.Execute "update zxd set 颜色2='' where 编号='" & Text2.Text & "' and 颜色2=null"
Data9.Database.Execute "update zxd set 数量3='' where 编号='" & Text2.Text & "' and 数量3=null"
Data9.Database.Execute "update zxd set 颜色3='' where 编号='" & Text2.Text & "' and 颜色3=null"
Data9.Database.Execute "update zxd set 数量4='' where 编号='" & Text2.Text & "' and 数量4=null"
Data9.Database.Execute "update zxd set 颜色4='' where 编号='" & Text2.Text & "' and 颜色4=null"
Data9.Database.Execute "update zxd set 数量5='' where 编号='" & Text2.Text & "' and 数量5=null"
Data9.Database.Execute "update zxd set 颜色5='' where 编号='" & Text2.Text & "' and 颜色5=null"

Data9.Database.Execute "update zxd set 数量1=-数量1 where 编号='" & Text2.Text & "' and val(数量1)>0"
Data9.Database.Execute "update zxd set 颜色1='' where 编号='" & Text2.Text & "' and 颜色1=null"
Data9.Database.Execute "update zxd set 数量2=-数量2 where 编号='" & Text2.Text & "' and val(数量2)>0"
Data9.Database.Execute "update zxd set 颜色2='' where 编号='" & Text2.Text & "' and 颜色2=null"
Data9.Database.Execute "update zxd set 数量3=-数量3 where 编号='" & Text2.Text & "' and val(数量3)>0"
Data9.Database.Execute "update zxd set 颜色3='' where 编号='" & Text2.Text & "' and 颜色3=null"
Data9.Database.Execute "update zxd set 数量4=-数量4 where 编号='" & Text2.Text & "' and val(数量4)>0"
Data9.Database.Execute "update zxd set 颜色4='' where 编号='" & Text2.Text & "' and 颜色4=null"
Data9.Database.Execute "update zxd set 数量5=-数量5 where 编号='" & Text2.Text & "' and val(数量5)>0"
Data9.Database.Execute "update zxd set 颜色5='' where 编号='" & Text2.Text & "' and 颜色5=null"

Data9.Database.Execute "update zxd set 小计=val(数量1)+val(数量2)+val(数量3)+val(数量4)+val(数量5),金额=format((val(数量1)+val(数量2)+val(数量3)+val(数量4)+val(数量5))*val(提货),'#0.0'),日期=cdate('" & DTPicker3.Value & "') where 编号='" & Text2.Text & "'"

Data1.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data1.RecordSource = "select 客户,单号,编号,款号,规格,单位,价格,颜色1,-val(数量1) as A,颜色2,-VAL(数量2) AS B,颜色3,-VAL(数量3) AS C,颜色4,-VAL(数量4) AS D,颜色5,-VAL(数量5) AS E,-VAL(小计) AS 小计数,提货,-VAL(金额) AS 合计金额,折扣 from zxd where 编号='" & Text2.Text & "'"
Data1.Refresh

End Sub

Private Sub Command8_Click()
On Error Resume Next
Data3.RecordSource = "SELECT MAX(VAL(MID(编号,3))) FROM zxd"
Data3.Refresh
Text2.Text = yhdm + "X000001"
If Data3.Recordset.EOF Then
Text2.Text = yhdm + "X000001"
Else
Text2.Text = Left(yhdm + "X000000", 8 - Len(Trim(Data3.Recordset.Fields(0) + 1))) + Trim(Data3.Recordset.Fields(0) + 1)
End If

Data1.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data1.RecordSource = "select 客户,单号,编号,款号,规格,单位,价格,颜色1,数量1,颜色2,数量2,颜色3,数量3,颜色4,数量4,颜色5,数量5,小计,提货,金额,折扣 from zxd where 编号='" & Text2.Text & "' order by 序号"
Data1.Refresh

Text3.Text = ""
Text4.Text = ""

Data11.RecordSource = "select 编号,-val(费用) as 运费,备注,单号,客户,序号,类别 from zxbz where 编号='" & Text2.Text & "'"
Data11.Refresh

End Sub


Private Sub Form_Load()
On Error Resume Next
DTPicker1.Value = Date
DTPicker2.Value = Date - 30
DTPicker3.Value = Date
Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Data1.DatabaseName = "e:\Excel\染整\宝隆\sjzz.MDB"
Data2.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data3.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data3.RecordSource = "SELECT MAX(VAL(MID(编号,3))) FROM zxd"
Data3.Refresh
Data4.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\sczyjhd.mdb"
Data5.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\sczyjhd.mdb"
Data6.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\sczyjhd.mdb"
Data7.DatabaseName = "e:\Excel\染整\宝隆\sjzz.MDB"
Data8.DatabaseName = "e:\Excel\染整\宝隆\sjzz.MDB"
Data9.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data10.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"
Data11.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\ZCW.MDB"

Text2.Text = yhdm + "X000001"
If Data3.Recordset.EOF Then
Text2.Text = yhdm + "X000001"
Else
Text2.Text = Left(yhdm + "X000000", 8 - Len(Trim(Data3.Recordset.Fields(0) + 1))) + Trim(Data3.Recordset.Fields(0) + 1)
End If
Option4.Value = True
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 0
MSFlexGrid1.ColWidth(2) = 0
MSFlexGrid1.ColWidth(3) = 1000
MSFlexGrid1.ColWidth(4) = 1000
For i = 5 To 27
MSFlexGrid1.ColWidth(i) = 450
Next
End Sub


Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex
   TreeView1.Nodes.Clear
 

If Option4.Value = True Then
    Data6.RecordSource = "select distinct 客户 from sczy_xtd where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
    Data6.Refresh
    m = 1
    If Not Data6.Recordset.EOF Then  'make sure there are records in the table
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data6.Recordset.Fields(0)
        intIndex = mNode.Index
        Data4.RecordSource = "select distinct 单号 from sczy_xtd where 客户='" & Data6.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
        Data4.Refresh
        
        If Not Data4.Recordset.EOF Then
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data4.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data5.RecordSource = "select distinct 款号 from sczy_xtd where 单号='" & Data4.Recordset.Fields(0) & "' and 进度='进行'"
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
        
        Data4.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        Data6.Recordset.MoveNext
        m = m + 1
        Loop
    End If
End If


If Option5.Value = True Then
    Data6.RecordSource = "select distinct 客户 from sczy_xtd where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
    Data6.Refresh
    m = 1
    If Not Data6.Recordset.EOF Then  'make sure there are records in the table
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data6.Recordset.Fields(0)
        intIndex = mNode.Index
        Data4.RecordSource = "select distinct 单号 from sczy_xtd where 客户='" & Data6.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
        Data4.Refresh
        
        If Not Data4.Recordset.EOF Then
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data4.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data5.RecordSource = "select distinct 款号 from sczy_xtd where 单号='" & Data4.Recordset.Fields(0) & "' and 进度='结束'"
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
        Data6.Recordset.MoveNext
        Loop
    End If
End If

End Sub


Private Sub MSFlexGrid1_Click()
With MSFlexGrid1
    c = .Col: r = .Row    '''''C列，，R行
End With
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If c = 1 Then
Data1.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data1.Recordset.Move rs - 1
If MsgBox("确定删除" + "第" + Trim(rs) + "行内容吗？", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
End If
End Sub

'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Data1.DatabaseName = "e:\Excel\染整\宝隆\sjzz.MDB"
If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") = 0 Then
Else
h1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
h2 = Mid(TreeView1.Nodes(Node.Index).FullPath, 1, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") - 1)
If InStr(h1, "\") > 0 Then
h1 = Mid(h1, 1, InStr(h1, "\") - 1)
Else
h1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
Text1.Text = h1
Call khtj(Trim(h1))
Call zxdfj
Call ck(Trim(h1))
End If
Data1.RecordSource = "select * from zxdb order by 序号"
Data1.Refresh

'DBCombo2.Text = Node.Index
'DBCombo3.Text = TreeView1.Nodes(Node.Index).FullPath
End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub

Private Sub zxdfj()
Dim lll As String
Dim k As Integer
Data7.Database.Execute "delete * from zxdf"

Data10.RecordSource = "select * from zxd where 单号='" & Text1.Text & "'"
Data10.Refresh

If Not Data10.Recordset.EOF Then
Data10.Recordset.MoveFirst
Do While Not Data10.Recordset.EOF
For k = 0 To 4
If -Val(Data10.Recordset.Fields(2 * k + 6)) > 0 Then
Data7.Database.Execute "INSERT INTO zxdf(客户,单号,款号,规格,颜色,数量) VALUES('" & Data10.Recordset.Fields(0) & "','" & Data10.Recordset.Fields(20) & "','" & Data10.Recordset.Fields(1) & "','" & Data10.Recordset.Fields(2) & "','" & Data10.Recordset.Fields(2 * k + 6 - 1) & "',-VAL('" & Data10.Recordset.Fields(2 * k + 6) & "'))"
End If
Next
Data10.Recordset.MoveNext
Loop
End If
End Sub

Private Sub ck(DH As String)
Dim lll As String
Data7.RecordSource = "select * from zxdb where 单号='" & DH & "' order by 序号"
Data7.Refresh
If Data7.Recordset.EOF Then Exit Sub
Data7.Recordset.MoveFirst
Do While Not Data7.Recordset.EOF
For i = 0 To 4
L = 5 + i * 4
Data8.RecordSource = "select 单号 from zxdf where 单号='" & Data7.Recordset.Fields(1) & "' and 款号='" & Data7.Recordset.Fields(2) & "' and 规格='" & Data7.Recordset.Fields(3) & "' and 颜色='" & Data7.Recordset.Fields(L) & "'"
Data8.Refresh
If Data8.Recordset.EOF Then
lll = ""
Else
Data8.RecordSource = "select sum(val(数量)) from zxdf where 单号='" & Data7.Recordset.Fields(1) & "' and 款号='" & Data7.Recordset.Fields(2) & "' and 规格='" & Data7.Recordset.Fields(3) & "' and 颜色='" & Data7.Recordset.Fields(L) & "'"
Data8.Refresh
lll = Trim(Data8.Recordset.Fields(0))
End If
Data7.Recordset.Edit
Data7.Recordset.Fields(L + 2) = lll
Data7.Recordset.Update
Next
Data7.Recordset.MoveNext
Loop
Data7.Database.Execute "update zxdb set 剩余1=val(数量1)-val(出库1),剩余2=val(数量2)-val(出库2),剩余3=val(数量3)-val(出库3),剩余4=val(数量4)-val(出库4),剩余5=val(数量5)-val(出库5)"

Data7.RecordSource = "select  单号 from zxdb where val(剩余1)>0 or val(剩余2)>0 or val(剩余3)>0 or val(剩余4)>0 or val(剩余5)>0"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data6.Database.Execute "update sczy_xtd set 进度='进行' where 单号='" & DH & "'"
Else
Data6.Database.Execute "update sczy_xtd set 进度='结束' where 单号='" & DH & "'"
End If

Data7.Database.Execute "update zxdb set 剩余1='' where 剩余1='0'"
Data7.Database.Execute "update zxdb set 剩余2='' where 剩余2='0'"
Data7.Database.Execute "update zxdb set 剩余3='' where 剩余3='0'"
Data7.Database.Execute "update zxdb set 剩余4='' where 剩余4='0'"
Data7.Database.Execute "update zxdb set 剩余5='' where 剩余5='0'"
End Sub

Private Sub MSFlex()
On Error Resume Next
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
    Call MSFlex
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    MSFlexGrid1.Text = ms
    MSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data1.Recordset.MoveFirst
Data1.Recordset.Move r - 1
Data1.Recordset.Edit
Data1.Recordset.Fields(c - 1) = Combo1111.Text
Data1.Recordset.Update
Combo1111.Visible = False
MSFlexGrid1.Text = Combo1111.Text
MSFlexGrid1.SetFocus
End If
End Sub


