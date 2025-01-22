VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formy307 
   BackColor       =   &H00C0E0FF&
   Caption         =   "实际进度"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form32"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   7440
      TabIndex        =   13
      Top             =   600
      Width           =   3135
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "生产"
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "材料"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
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
      Top             =   10560
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
      Top             =   10080
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
      Top             =   9960
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
      TabIndex        =   12
      Top             =   1680
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
      Top             =   9600
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
      Top             =   9720
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
      Top             =   9840
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
      Top             =   9840
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
      Top             =   9840
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
      Top             =   9600
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
      Top             =   9720
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
      Top             =   9930
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
      Top             =   10200
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "状态"
      Height          =   855
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000C0C0&
         Caption         =   "全部"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C0C0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "进行"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   0
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
      Top             =   10560
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80609281
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80609281
      CurrentDate     =   36892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy307.frx":0000
      Height          =   7455
      Left            =   3360
      TabIndex        =   8
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
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6975
      Left            =   240
      TabIndex        =   11
      Top             =   2160
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
      TabIndex        =   10
      Top             =   600
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
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Formy307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call JHJD(MSFlexGrid1, "计划进度")
End Sub

Private Sub dhjd(DH As String)
On Error Resume Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''大货订单
lo = "d:\数据库\\htgl\2011\scjd.mdb"
Data3.Database.Execute "delete * from sjjd"
Data1.Database.Execute "insert into SJJD(单号,款号,颜色,品名,数量,出货日期,序号) in'" & lo & "' select 单号,款号,颜色,品名,数量,交期,序号 from sczy_xdh where 单号='" & DH & "'"
Data7.RecordSource = "select * from SJJD"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.Recordset.MoveFirst
Do While Not Data7.Recordset.EOF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''裁剪
Data7.Recordset.Edit
Data3.RecordSource = "select 款号,颜色 from cjrb where 单号='" & Data7.Recordset.Fields(0) & "' and  款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.RecordSource = "select sum(val(裁剪)) from cjrb where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
Data7.Recordset.Fields(17) = Data3.Recordset.Fields(0)
Else
Data7.Recordset.Fields(17) = "0"
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''绣印
Data3.RecordSource = "select 款号,颜色 from wxrk where 单号='" & Data7.Recordset.Fields(0) & "' and  款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 类别='绣印'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(数量)) from wxrk where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 类别='绣印'"
Data8.Refresh
Data7.Recordset.Fields(18) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(18) = "0"
Else
Data7.Recordset.Fields(18) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(18) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''缝转包

Data3.RecordSource = "select 款号,颜色 from cpk where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(数量)) from cpk where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data8.Refresh
Data7.Recordset.Fields(19) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(19) = "0"
Else
Data7.Recordset.Fields(19) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(19) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''后整
Data6.RecordSource = "select 款式,颜色 from clb where 单号='" & Data7.Recordset.Fields(0) & "' and 款式='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 工序='包装'"
Data6.Refresh
If Not Data6.Recordset.EOF Then
Data9.RecordSource = "select sum(产量) from clb where 单号='" & Data7.Recordset.Fields(0) & "' and 款式='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 工序='包装'"
Data9.Refresh
Data7.Recordset.Fields(20) = "0"
If Data9.Recordset.EOF Then
Data7.Recordset.Fields(20) = "0"
Else
Data7.Recordset.Fields(20) = Data9.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(20) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''成品入库
Data10.RecordSource = "select 款号,规格 from LSRK where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 规格='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(数量) from LSRK where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 规格='" & Data7.Recordset.Fields(2) & "'"
Data5.Refresh
Data7.Recordset.Fields(21) = "0"
If Data5.Recordset.EOF Then
Data7.Recordset.Fields(21) = "0"
Else
Data7.Recordset.Fields(21) = Data5.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(21) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''成品出库
Data10.RecordSource = "select 款号,规格 from lsfh where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 规格='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(数量) from lsfh where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data5.Refresh
Data7.Recordset.Fields(22) = "0"
If Data5.Recordset.EOF Then
Data7.Recordset.Fields(22) = "0"
Else
Data7.Recordset.Fields(22) = Data5.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(22) = "0"
End If
Data7.Recordset.Update
Data7.Recordset.MoveNext
Loop
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''主料材料
Data7.Database.Execute "update sjjd set 序号='0' where 序号=null"

Data8.RecordSource = "select 数量,大货裁剪,成品入库 from sjjd where 单号='" & DH & "'"
Data8.Refresh
If Not Data8.Recordset.EOF Then
Data8.Recordset.MoveFirst
pd = 0
Do While Not Data8.Recordset.EOF
If Data8.Recordset.Fields(1) >= Data8.Recordset.Fields(0) And Data8.Recordset.Fields(1) = Data8.Recordset.Fields(2) Then
Else
pd = pd + 1
End If
Data8.Recordset.MoveNext
Loop
If pd = 0 Then
Data12.Database.Execute "update sczy_xdh set 进度='结束' where 单号='" & DH & "'"
End If
End If


Data2.RecordSource = "select 单号,款号,品名,颜色,数量,大货裁剪,大货绣印,大货缝制,大货后整,成品入库,成品出库,出货日期 FROM sjjd  order by 单号,款号,颜色,序号"
Data2.Refresh

End Sub

Private Sub dhcl(DH As String)
On Error Resume Next
Data8.Database.Execute "DELETE * FROM cljd"
Data1.Database.Execute "DELETE * FROM CKGL"
Data4.Database.Execute "INSERT INTO CKGL(单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'd:\数据库\\htgl\2011\SCZYJHD.MDB' SELECT 单号,库类,材料名称,材料规格,材料单位,颜色,批次,SUM(数量) FROM CKGL WHERE CKGL.单号='" & DH & "' GROUP BY 单号,库类,材料名称,材料规格,材料单位,颜色,批次 "
Data11.Database.Execute "insert into ckgl(单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'd:\数据库\\htgl\2011\SCZYJHD.MDB' select 单号,'1主料库',材料名称,光坯幅宽,单位,颜色,光坯克重,光坯重量 from rsrk where 单号='" & DH & "'"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',采购数量=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,采购数量) SELECT CGCLB.单号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号,SUM(CGCLB.材料数量) AS 采购数量 FROM CGCLB WHERE instr(单号,'" & DH & "')>0 GROUP BY CGCLB.单号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',入库数量=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO cljd(单号,库类,材料,规格,色别,批次,用料,来料,欠料) IN'd:\数据库\\htgl\2011\scjd.MDB' SELECT 单号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料颜色,CKGL.材料批号,Format(SUM(CKGL.采购数量),'#0.00') AS 采购量,Format(SUM(CKGL.入库数量),'#0.00') AS 入库量,Format(SUM(CKGL.采购数量-CKGL.入库数量),'#0.00') FROM CKGL  GROUP BY 单号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料颜色,CKGL.材料批号"
Data2.RecordSource = "select * FROM cljd  order by 库类,颜色"
Data2.Refresh
End Sub


Private Sub khcl(kh As String)
On Error Resume Next
If Option1.Value = True Then
Data13.RecordSource = "select 单号 from sczy_xdh where 客户='" & kh & "' and 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
Data13.Refresh
End If
If Option2.Value = True Then
Data13.RecordSource = "select 单号 from sczy_xdh where 客户='" & kh & "' and 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
Data13.Refresh
End If
If Option3.Value = True Then
Data13.RecordSource = "select 单号 from sczy_xdh where 客户='" & kh & "' and 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data13.Refresh
End If

If Not Data13.Recordset.EOF Then
Data8.Database.Execute "DELETE * FROM cljd"
Data1.Database.Execute "DELETE * FROM CKGL"
Data13.Recordset.MoveFirst
Do While Not Data13.Recordset.EOF
Data4.Database.Execute "INSERT INTO CKGL(单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'd:\数据库\\htgl\2011\SCZYJHD.MDB' SELECT 单号,库类,材料名称,材料规格,材料单位,颜色,批次,SUM(数量) FROM CKGL WHERE 单号='" & Data13.Recordset.Fields(0) & "' GROUP BY 单号,库类,材料名称,材料规格,材料单位,颜色,批次 "
Data11.Database.Execute "insert into ckgl(单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'd:\数据库\\htgl\2011\SCZYJHD.MDB' select 单号,'1主料库',材料名称,光坯幅宽,单位,颜色,光坯克重,光坯重量 from rsrk where instr(单号,'" & Data13.Recordset.Fields(0) & "')>0"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',采购数量=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,采购数量) SELECT CGCLB.单号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号,SUM(CGCLB.材料数量) AS 采购数量 FROM CGCLB WHERE instr(单号,'" & Data13.Recordset.Fields(0) & "')>0 GROUP BY CGCLB.单号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',入库数量=0 WHERE LX=NULL"
Data13.Recordset.MoveNext
Loop
End If
Data1.Database.Execute "INSERT INTO cljd(单号,库类,材料,规格,色别,批次,用料,来料,欠料) IN'd:\数据库\\htgl\2011\scjd.MDB' SELECT 单号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料颜色,CKGL.材料批号,Format(SUM(CKGL.采购数量),'#0.00') AS 采购量,Format(SUM(CKGL.入库数量),'#0.00') AS 入库量,Format(SUM(CKGL.采购数量-CKGL.入库数量),'#0.00') FROM CKGL  GROUP BY 单号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料颜色,CKGL.材料批号"
Data2.RecordSource = "select * FROM cljd  order by 库类,颜色"
Data2.Refresh
End Sub

Private Sub khjd(kh As String)
On Error Resume Next

On Error Resume Next
If Option1.Value = True Then
Data13.RecordSource = "select 单号 from sczy_xdh where 客户='" & kh & "' and 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
Data13.Refresh
End If
If Option2.Value = True Then
Data13.RecordSource = "select 单号 from sczy_xdh where 客户='" & kh & "' and 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
Data13.Refresh
End If
If Option3.Value = True Then
Data13.RecordSource = "select 单号 from sczy_xdh where 客户='" & kh & "' and 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data13.Refresh
End If

lo = "d:\数据库\\htgl\2011\scjd.mdb"

If Not Data13.Recordset.EOF Then
Data3.Database.Execute "delete * from sjjd"
Data13.Recordset.MoveFirst
Do While Not Data13.Recordset.EOF
Data1.Database.Execute "insert into SJJD(单号,款号,颜色,品名,数量,出货日期,序号) in'" & lo & "' select 单号,款号,颜色,品名,数量,交期,序号 from sczy_xdh where instr(单号,'" & Data13.Recordset.Fields(0) & "')>0"
Data13.Recordset.MoveNext
Loop
End If

lo = "d:\数据库\\htgl\2011\scjd.mdb"
Data7.RecordSource = "select * from SJJD"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.Recordset.MoveFirst
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''裁剪
Do While Not Data7.Recordset.EOF
Data7.Recordset.Edit
Data3.RecordSource = "select 款号,颜色 from cjrb where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.RecordSource = "select sum(val(裁剪)) from cjrb where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' group by 款号,颜色"
Data3.Refresh
Data7.Recordset.Fields(17) = Data3.Recordset.Fields(0)
Else
Data7.Recordset.Fields(17) = "0"
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''绣印
Data3.RecordSource = "select 款号,颜色 from wxrk where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 类别='绣印'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(数量)) from wxrk where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 类别='绣印'"
Data8.Refresh
Data7.Recordset.Fields(18) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(18) = "0"
Else
Data7.Recordset.Fields(18) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(18) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''缝转包

Data3.RecordSource = "select 款号,颜色 from cpk where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(数量)) from cpk where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data8.Refresh
Data7.Recordset.Fields(19) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(19) = "0"
Else
Data7.Recordset.Fields(19) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(19) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''后整
Data6.RecordSource = "select 款式,颜色 from clb where 单号='" & Data7.Recordset.Fields(0) & "' and 款式='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 工序='包装'"
Data6.Refresh
If Not Data6.Recordset.EOF Then
Data9.RecordSource = "select sum(产量) from clb where 单号='" & Data7.Recordset.Fields(0) & "' and 款式='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 工序='包装'"
Data9.Refresh
Data7.Recordset.Fields(20) = "0"
If Data9.Recordset.EOF Then
Data7.Recordset.Fields(20) = "0"
Else
Data7.Recordset.Fields(20) = Data9.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(20) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''成品入库
Data10.RecordSource = "select 款号,规格 from LSRK where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 规格='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(数量) from LSRK where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 规格='" & Data7.Recordset.Fields(2) & "'"
Data5.Refresh
Data7.Recordset.Fields(21) = "0"
If Data5.Recordset.EOF Then
Data7.Recordset.Fields(21) = "0"
Else
Data7.Recordset.Fields(21) = Data5.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(21) = "0"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''成品出库
Data3.RecordSource = "select 款号,颜色 from lsfh where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(数量) from lsfh where 单号='" & Data7.Recordset.Fields(0) & "' and 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data8.Refresh
Data7.Recordset.Fields(22) = "0"
If Data8.Recordset.EOF Then
Data7.Recordset.Fields(22) = "0"
Else
Data7.Recordset.Fields(22) = Data8.Recordset.Fields(0)
End If
Else
Data7.Recordset.Fields(22) = "0"
End If
Data7.Recordset.Update
Data7.Recordset.MoveNext
Loop
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''主料材料
Data7.Database.Execute "update sjjd set 序号='0' where 序号=null"
Data2.RecordSource = "select 单号,款号,品名,颜色,数量,大货裁剪,大货绣印,大货缝制,大货后整,成品入库,成品出库,出货日期 FROM sjjd  order by 单号,款号,颜色,序号"
Data2.Refresh

End Sub

Private Sub Command7_Click()
Call tree
Call zk
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
Data4.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data5.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data6.DatabaseName = "d:\数据库\\htgl\2011\db.mdb"
Data7.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
Data8.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
Data9.DatabaseName = "d:\数据库\\htgl\2011\db.mdb"
Data10.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data11.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
Data12.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data13.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data14.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Option4.Value = True
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
    Data12.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data13.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
        Data13.Refresh
        
        If Not Data13.Recordset.EOF Then
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data13.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data14.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data13.Recordset.Fields(0) & "' and 进度='进行'"
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
    Data12.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data13.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
        Data13.Refresh
        
        If Not Data13.Recordset.EOF Then
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data13.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data14.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data13.Recordset.Fields(0) & "'"
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
    Data12.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
    Data12.Refresh
    m = 1
    If Not Data12.Recordset.EOF Then  'make sure there are records in the table
        Data12.Recordset.MoveFirst
        Do While Not Data12.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "x" + Trim(m)
        mNode.Text = Data12.Recordset.Fields(0)
        intIndex = mNode.Index
        Data13.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data12.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='结束'"
        Data13.Refresh
        
        If Not Data13.Recordset.EOF Then
        Data13.Recordset.MoveFirst
        Do While Not Data13.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "x" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data13.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data14.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data13.Recordset.Fields(0) & "' and 进度='结束'"
        Data14.Refresh
        
        If Not Data14.Recordset.EOF Then
        Data14.Recordset.MoveFirst
        Do While Not Data14.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "x" + Trim(m) + "w" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data14.Recordset.Fields(0))
        Data14.Recordset.MoveNext
        Loop
        End If
        
        Data13.Recordset.MoveNext
        Loop
        End If
        Data12.Recordset.MoveNext
        Loop
    End If
End If

End Sub


'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next



If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") = 0 Then
If Option4.Value = True Then
Call khcl(TreeView1.Nodes(Node.Index).FullPath)
End If

If Option5.Value = True Then
Call khjd(TreeView1.Nodes(Node.Index).FullPath)
End If

Else

l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If

If Option4.Value = True Then
Call dhcl(Trim(l1))
End If

If Option5.Value = True Then
Call dhjd(Trim(l1))
End If

End If

'DBCombo2.Text = Node.Index
'DBCombo3.Text = TreeView1.Nodes(Node.Index).FullPath
End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub





