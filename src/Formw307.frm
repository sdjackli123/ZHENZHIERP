VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw307 
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
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号总体进度"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号材料进度"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   285
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单号总体进度"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1455
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "状态"
      Height          =   855
      Left            =   11160
      TabIndex        =   3
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000C0C0&
         Caption         =   "全部"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C0C0&
         Caption         =   "结束"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "进行"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单号材料进度"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1455
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
      Left            =   9960
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
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   93192193
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   93192193
      CurrentDate     =   36892
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   360
      Left            =   4440
      TabIndex        =   11
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw307.frx":0000
      Height          =   7455
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   8421631
      BackColorBkg    =   4109501
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   360
      Left            =   4440
      TabIndex        =   13
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择款号"
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
      Left            =   3360
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   1080
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
      Index           =   1
      Left            =   3360
      TabIndex        =   14
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Formw307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public C, R As Integer
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call jhjd(MSFlexGrid1, "计划进度")
End Sub

Private Sub Command3_Click()
On Error Resume Next
If DBCombo1.text = "" Then
MsgBox ("请输入单号")
Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''样品进度
LO = "D:\数据库\htgl\2011\scjd.mdb"
Data3.Database.Execute "delete * from sjjd"
Data1.Database.Execute "insert into SJJD(单号,款号,颜色,数量,出货日期,序号) in'" & LO & "' select 单号,款号,颜色,数量,交货期,序号 from sczy_xdh where 单号='" & DBCombo1.text & "'"
Data7.RecordSource = "select * from SJJD"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.Recordset.MoveFirst
Do While Not Data7.Recordset.EOF
Data3.RecordSource = "select * from ypjd where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
Data7.Recordset.Edit
If Not Data3.Recordset.EOF Then
Data7.Recordset.Fields(4) = Data3.Recordset.Fields(5)
Data7.Recordset.Fields(5) = Data3.Recordset.Fields(6)
Data7.Recordset.Fields(14) = Data3.Recordset.Fields(12)
Data7.Recordset.Fields(15) = Data3.Recordset.Fields(11)
Data7.Recordset.Fields(16) = Data3.Recordset.Fields(14)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''裁剪
Data3.RecordSource = "select 款号,颜色 from cjrb where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.RecordSource = "select sum(val(裁剪)) from cjrb where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' group by 款号,颜色"
Data3.Refresh
Data7.Recordset.Fields(17) = Data3.Recordset.Fields(0)
Else
Data7.Recordset.Fields(17) = "0"
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''绣印
Data3.RecordSource = "select 款号,颜色 from wxrk where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 类别='绣印'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(数量)) from wxrk where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 类别='绣印'"
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

Data3.RecordSource = "select 款号,颜色 from cpk where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(数量)) from cpk where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
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
Data6.RecordSource = "select 款式,颜色 from clb where 款式='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 工序='包装'"
Data6.Refresh
If Not Data6.Recordset.EOF Then
Data9.RecordSource = "select sum(产量) from clb where 款式='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 工序='包装'"
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
Data10.RecordSource = "select 款号,规格 from LSRK where 款号='" & Data7.Recordset.Fields(1) & "' and 规格='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(数量) from LSRK where 款号='" & Data7.Recordset.Fields(1) & "' and 规格='" & Data7.Recordset.Fields(2) & "'"
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
Data3.RecordSource = "select 款号,颜色 from zxd where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(合计件)) from zxd where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
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
Data2.RecordSource = "select 单号,款号,颜色,数量,主料,配料,绣印,缝制样,产前样,大货裁剪,大货绣印,大货缝制,大货后整,成品入库,成品出库,出货日期 FROM sjjd  order by 款号,颜色,序号"
Data2.Refresh

End Sub

Private Sub Command4_Click()
On Error Resume Next
If DBCombo1.text = "" Then
MsgBox ("请输入单号")
Exit Sub
End If
Data8.Database.Execute "DELETE * FROM cljd"
Data1.Database.Execute "DELETE * FROM CKGL"
Data4.Database.Execute "INSERT INTO CKGL(单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'D:\数据库\htgl\2011\SCZYJHD.MDB' SELECT 单号,合约号,库类,材料名称,材料规格,材料单位,颜色,批次,SUM(数量) FROM CKGL WHERE CKGL.单号='" & DBCombo1.text & "' GROUP BY 单号,合约号,库类,材料名称,材料规格,材料单位,颜色,批次 "
Data11.Database.Execute "insert into ckgl(单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'D:\数据库\htgl\2011\SCZYJHD.MDB' select 单号,款号,'1主料库',材料名称,光坯幅宽,单位,颜色,光坯克重,光坯重量 from rsrk where 单号='" & DBCombo1.text & "'"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',采购数量=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,采购数量) SELECT CGCLB.单号,款号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号,SUM(CGCLB.材料数量) AS 采购数量 FROM CGCLB WHERE CGCLB.款号='" & DBCombo2.text & "' GROUP BY CGCLB.单号,款号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',入库数量=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO cljd(单号,款号,库类,材料,规格,色别,批次,用料,来料,欠料) IN'D:\数据库\htgl\2011\scjd.MDB' SELECT 单号,CKGL.款号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料颜色,CKGL.材料批号,format(SUM(CKGL.采购数量),'#0.00') AS 采购量,format(SUM(CKGL.入库数量),'#0.00') AS 入库量,format(SUM(CKGL.采购数量-CKGL.入库数量),'#0.00') FROM CKGL  GROUP BY 单号,CKGL.款号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料颜色,CKGL.材料批号"
Data2.RecordSource = "select * FROM cljd  order by 库类,款号,颜色"
Data2.Refresh
End Sub


Private Sub Command5_Click()
On Error Resume Next
If DBCombo2.text = "" Then
MsgBox ("请输入款号")
Exit Sub
End If
Data8.Database.Execute "DELETE * FROM cljd"
Data1.Database.Execute "DELETE * FROM CKGL"
Data4.Database.Execute "INSERT INTO CKGL(单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'D:\数据库\htgl\2011\SCZYJHD.MDB' SELECT 单号,合约号,库类,材料名称,材料规格,材料单位,颜色,批次,SUM(数量) FROM CKGL WHERE CKGL.合约号='" & DBCombo2.text & "' GROUP BY 单号,合约号,库类,材料名称,材料规格,材料单位,颜色,批次 "
Data11.Database.Execute "insert into ckgl(单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'D:\数据库\htgl\2011\SCZYJHD.MDB' select 单号,款号,'1主料库',材料名称,光坯幅宽,单位,颜色,光坯克重,光坯重量 from rsrk where 款号='" & DBCombo2.text & "'"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',采购数量=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,采购数量) SELECT CGCLB.单号,款号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号,SUM(CGCLB.材料数量) AS 采购数量 FROM CGCLB WHERE CGCLB.款号='" & DBCombo2.text & "' GROUP BY CGCLB.单号,款号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号"
Data1.Database.Execute "UPDATE CKGL SET LX='CK',入库数量=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO cljd(单号,款号,库类,材料,规格,色别,批次,用料,来料,欠料) IN'D:\数据库\htgl\2011\scjd.MDB' SELECT 单号,CKGL.款号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料颜色,CKGL.材料批号,format(SUM(CKGL.采购数量),'#0.00') AS 采购量,format(SUM(CKGL.入库数量),'#0.00') AS 入库量,format(SUM(CKGL.采购数量-CKGL.入库数量),'#0.00') FROM CKGL  GROUP BY 单号,CKGL.款号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料颜色,CKGL.材料批号"
Data2.RecordSource = "select * FROM cljd  order by 库类,款号,颜色"
Data2.Refresh
End Sub

Private Sub Command6_Click()
On Error Resume Next
If DBCombo2.text = "" Then
MsgBox ("请输入款号")
Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''样品进度
LO = "D:\数据库\htgl\2011\scjd.mdb"
Data3.Database.Execute "delete * from sjjd"
Data1.Database.Execute "insert into SJJD(单号,款号,颜色,数量,出货日期,序号) in'" & LO & "' select 单号,款号,颜色,数量,交货期,序号 from sczy_xdh where 款号='" & DBCombo2.text & "'"
Data7.RecordSource = "select * from SJJD"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.Recordset.MoveFirst
Do While Not Data7.Recordset.EOF
Data3.RecordSource = "select * from ypjd where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
Data7.Recordset.Edit
If Not Data3.Recordset.EOF Then
Data7.Recordset.Fields(4) = Data3.Recordset.Fields(5)
Data7.Recordset.Fields(5) = Data3.Recordset.Fields(6)
Data7.Recordset.Fields(14) = Data3.Recordset.Fields(12)
Data7.Recordset.Fields(15) = Data3.Recordset.Fields(11)
Data7.Recordset.Fields(16) = Data3.Recordset.Fields(14)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''裁剪
Data3.RecordSource = "select 款号,颜色 from cjrb where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.RecordSource = "select sum(val(裁剪)) from cjrb where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' group by 款号,颜色"
Data3.Refresh
Data7.Recordset.Fields(17) = Data3.Recordset.Fields(0)
Else
Data7.Recordset.Fields(17) = "0"
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''绣印
Data3.RecordSource = "select 款号,颜色 from wxrk where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(数量)) from wxrk where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
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

Data3.RecordSource = "select 款号,颜色 from cpk where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(数量)) from cpk where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 编号='01'"
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
Data6.RecordSource = "select 款式,颜色 from clb where 款式='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 工序='包装'"
Data6.Refresh
If Not Data6.Recordset.EOF Then
Data9.RecordSource = "select sum(产量) from clb where 款式='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "' and 工序='包装'"
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
Data10.RecordSource = "select 款号,规格 from LSRK where 款号='" & Data7.Recordset.Fields(1) & "' and 规格='" & Data7.Recordset.Fields(2) & "'"
Data10.Refresh
If Not Data10.Recordset.EOF Then
Data5.RecordSource = "select sum(数量) from LSRK where 款号='" & Data7.Recordset.Fields(1) & "' and 规格='" & Data7.Recordset.Fields(2) & "'"
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
Data3.RecordSource = "select 款号,颜色 from zxd where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data8.RecordSource = "select sum(val(合计件)) from zxd where 款号='" & Data7.Recordset.Fields(1) & "' and 颜色='" & Data7.Recordset.Fields(2) & "'"
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
Data2.RecordSource = "select 单号,款号,颜色,数量,主料,配料,绣印,缝制样,产前样,大货裁剪,大货绣印,大货缝制,大货后整,成品入库,成品出库,出货日期 FROM sjjd  order by 款号,颜色,序号"
Data2.Refresh

End Sub

Private Sub Form_Load()
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
DBCombo1.text = ""
DBCombo2.text = ""
Data1.DatabaseName = "D:\数据库\htgl\2011\SCZYJHD.mdb"
Data2.DatabaseName = "D:\数据库\htgl\2011\scjd.mdb"
Data3.DatabaseName = "D:\数据库\htgl\2011\scjd.mdb"
Data4.DatabaseName = "D:\数据库\htgl\2011\ckgl.mdb"
Data5.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"
Data6.DatabaseName = "D:\数据库\htgl\2011\db.mdb"
Data7.DatabaseName = "D:\数据库\htgl\2011\scjd.mdb"
Data8.DatabaseName = "D:\数据库\htgl\2011\scjd.mdb"
Data9.DatabaseName = "D:\数据库\htgl\2011\db.mdb"
Data10.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"
Data11.DatabaseName = "D:\数据库\htgl\2011\scjd.mdb"

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 0
End Sub

Private Sub MSFlex()
On Error Resume Next
With MSFlexGrid1
    C = .Col: R = .Row    '''''C列，，R行
        Text1111.Left = .Left + .ColPos(C)
        Text1111.Top = .Top + .RowPos(R)
        Text1111.Width = .ColWidth(C)
        Text1111.Height = .RowHeight(R)
        Text1111 = .text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 2
khbl = 12
Formw202.Show
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
Data2.Recordset.Move R - 1
Data2.Recordset.Edit
Data2.Recordset.Fields(C - 1) = Text1111.text
Data2.Recordset.Update
Text1111.Visible = False
MSFlexGrid1.text = Text1111.text
MSFlexGrid1.SetFocus
End If
End Sub




