VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Forml505 
   BackColor       =   &H00C0E0FF&
   Caption         =   "加工明细查询"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Forml505.frx":0000
      Height          =   7815
      Left            =   360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13785
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单号刷新"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "凭证生成"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成查询"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   480
      Width           =   2655
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "印花"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "织布"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "染色"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   1140
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
      Top             =   10200
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
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
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   1695
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   1140
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   4335
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单位刷新"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1095
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
      Top             =   10200
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
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   2655
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1095
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   3735
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Forml505.frx":0014
      Height          =   330
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   11640
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   22937601
      CurrentDate     =   36892
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Left            =   5160
      TabIndex        =   18
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
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
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   19
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      Height          =   375
      Index           =   0
      Left            =   11640
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
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
      TabIndex        =   9
      Top             =   960
      Width           =   855
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
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "加工单位"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Forml505"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer
Private Sub Command1_Click()
If Option4.Value = True Then
If DBCombo1.Text = "" Then
Data4.RecordSource = "select 织布单位,单号,材料名称,光坯克重,光坯幅宽,毛坯幅宽,毛坯重量,织布单价,织布金额,锅号,染厂色别,日期 from rsrk where 织布单位<>'' and 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') order by 日期"
Data4.Refresh
Else
Data4.RecordSource = "select 织布单位,单号,材料名称,光坯克重,光坯幅宽,毛坯幅宽,毛坯重量,织布单价,织布金额,锅号,染厂色别,日期 from rsrk where  织布单位='" & DBCombo1.Text & "' and 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') order by 日期"
Data4.Refresh
End If
End If

If Option5.Value = True Then
If DBCombo1.Text = "" Then
Data4.RecordSource = "select 染色单位,单号,材料名称,光坯克重,光坯幅宽,光坯重量,毛坯重量,单价,金额,锅号,染厂色别,染耗,日期 from rsrk where  染色单位<>'' and 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') order by 日期"
Data4.Refresh
Else
Data4.RecordSource = "select 染色单位,单号,材料名称,光坯克重,光坯幅宽,光坯重量,毛坯重量,单价,金额,锅号,染厂色别,染耗,日期 from rsrk where  染色单位='" & DBCombo1.Text & "' and 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') order by 日期"
Data4.Refresh
End If
End If

If Option1.Value = True Then
If DBCombo1.Text = "" Then
Data4.RecordSource = "select 印花单位,单号,材料名称,光坯克重,毛坯幅宽,光坯重量,印花单价,印花数量,印花金额,锅号,染厂色别,日期 from rsrk where  印花单位<>'' and 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') order by 日期"
Data4.Refresh
Else
Data4.RecordSource = "select 印花单位,单号,材料名称,光坯克重,毛坯幅宽,光坯重量,印花单价,印花数量,印花金额,锅号,染厂色别,日期 from rsrk where  印花单位='" & DBCombo1.Text & "' and 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') order by 日期"
Data4.Refresh
End If
End If

End Sub

Private Sub Command2_Click()
Formw332.Combo1.Text = "转账凭证"
Formw332.Show
End Sub

Private Sub Command3_Click()
If MsgBox("操作日期为：" + Trim(DTPicker1.Value) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("操作期间为：" + Trim(Month(DTPicker1.Value)) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定生成加工系列的凭证吗？", vbYesNo) = vbNo Then Exit Sub
Call ZBJGPZ(CDate(DTPicker3.Value), CDate(DTPicker4.Value), CDate(DTPicker1.Value))
Call RSJGPZ(CDate(DTPicker3.Value), CDate(DTPicker4.Value), CDate(DTPicker1.Value))
Call YHJGPZ(CDate(DTPicker3.Value), CDate(DTPicker4.Value), CDate(DTPicker1.Value))
End Sub

Private Sub Command4_Click()
Call qtmx(MSFlexGrid1, 9, "加工对账明细")
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
If Option4.Value = True Then
Data4.RecordSource = "select 织布单位,单号,材料名称,光坯克重,光坯幅宽,毛坯幅宽,毛坯重量,织布单价,织布金额,锅号,染厂色别,日期 from rsrk where  单号='" & DBCombo2.Text & "' order by 日期"
Data4.Refresh
End If

If Option5.Value = True Then
Data4.RecordSource = "select 染色单位,单号,材料名称,光坯克重,光坯幅宽,光坯重量,毛坯重量,单价,金额,锅号,染厂色别,染耗,日期 from rsrk where  单号='" & DBCombo2.Text & "' order by 日期"
Data4.Refresh
End If

If Option1.Value = True Then
Data4.RecordSource = "select 印花单位,单号,材料名称,光坯克重,毛坯幅宽,光坯重量,印花单价,印花数量,印花金额,锅号,染厂色别,日期 from rsrk where  单号='" & DBCombo2.Text & "' order by 日期"
Data4.Refresh
End If
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DTPicker1.Value = Date
DTPicker3.Value = Date - 30
DTPicker4.Value = Date
Option4.Value = True
Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

Data10.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data11.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data12.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.DatabaseName = "d:\数据库\\htgl\2011\cw.mdb"
Data3.DatabaseName = "d:\数据库\\htgl\2011\cw.mdb"
DBCombo2.Text = ""
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(12) = 1300

End Sub

Private Sub Option1_Click()
Data1.RecordSource = "select 简称 from gys where instr(代码,'印')>0 group by 简称"
Data1.Refresh
End Sub

Private Sub Option4_Click()
Data1.RecordSource = "select 简称 from gys where instr(代码,'织')>0 group by 简称"
Data1.Refresh
End Sub

Private Sub Option5_Click()
Data1.RecordSource = "select 简称 from gys where instr(代码,'染')>0 group by 简称"
Data1.Refresh
End Sub

Private Sub ZBJGPZ(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE instr(制单,'加工-材料')>0 AND 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data2.Refresh
If Not Data2.Recordset.EOF Then
If MsgBox("已有加工生成凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
Data3.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(制单,'加工-材料')>0 and 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data4.RecordSource = "select 织布单位,单号,材料名称,光坯克重,光坯幅宽,毛坯幅宽,毛坯重量,织布单价,织布金额,锅号,染厂色别,日期 from rsrk where 织布单位<>'' and 日期 between cdate('" & dt1 & "') and cdate('" & dt2 & "')"
Data4.Refresh
If Not Data4.Recordset.EOF Then
Data4.RecordSource = "select 织布单位,FORMAT(SUM(VAL(织布金额)),'#0.00') from rsrk where 织布单位<>'' and 日期 between cdate('" & dt1 & "') and cdate('" & dt2 & "') GROUP BY 织布单位"
Data4.Refresh

Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data3.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data3.Refresh
PZH = "5-" + Trim(Data3.Recordset.Fields(0) + 1)
Else
PZH = "5-1"
End If

Data4.Recordset.MoveFirst
KLLLL = 1

Do While Not Data4.Recordset.EOF
For i = 1 To 3
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = "加工费"
Data2.Recordset.Fields(1) = "原材料"
Data2.Recordset.Fields(2) = ""
Data2.Recordset.Fields(3) = "应付账款"
Data2.Recordset.Fields(4) = Data4.Recordset.Fields(0)
Data2.Recordset.Fields(5) = Format(Data4.Recordset.Fields(1) * 0.83, "#0.00")
Data2.Recordset.Fields(6) = PZH
Data2.Recordset.Fields(7) = CDate(dt3)
Data2.Recordset.Fields(8) = ""
Data2.Recordset.Fields(9) = ""
Data2.Recordset.Fields(10) = ""
Data2.Recordset.Fields(11) = "加工-材料"
Data2.Recordset.Update


Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = "加工费"
Data2.Recordset.Fields(1) = "应缴税金"
Data2.Recordset.Fields(2) = "税金进项"
Data2.Recordset.Fields(3) = "应付账款"
Data2.Recordset.Fields(4) = Data4.Recordset.Fields(0)
Data2.Recordset.Fields(5) = Format(Data4.Recordset.Fields(1) * 0.17, "#0.00")
Data2.Recordset.Fields(6) = PZH
Data2.Recordset.Fields(7) = CDate(dt3)
Data2.Recordset.Fields(8) = ""
Data2.Recordset.Fields(9) = ""
Data2.Recordset.Fields(10) = ""
Data2.Recordset.Fields(11) = "加工-材料"
Data2.Recordset.Update


Data4.Recordset.MoveNext
If Data4.Recordset.EOF Then
MsgBox ("材料加工单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data3.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data3.Refresh
PZH = "5-" + Trim(Data3.Recordset.Fields(0) + 1)
Else
PZH = "5-1"
End If
Loop
MsgBox ("材料加工单转账成功！" + "生成" + Str(KLLLL) + "凭证")
End If

End Sub


Private Sub YHJGPZ(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE instr(制单,'加工-材料')>0 AND 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data2.Refresh
If Not Data2.Recordset.EOF Then
If MsgBox("已有加工生成凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
Data3.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(制单,'加工-材料')>0 and 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data4.RecordSource = "select 印花单位,单号,材料名称,光坯克重,光坯幅宽,毛坯幅宽,毛坯重量,织布单价,印花金额,锅号,染厂色别,日期 from rsrk where 印花单位<>'' and 日期 between cdate('" & dt1 & "') and cdate('" & dt2 & "')"
Data4.Refresh
If Not Data4.Recordset.EOF Then
Data4.RecordSource = "select 印花单位,FORMAT(SUM(VAL(印花金额)),'#0.00') from rsrk where 印花单位<>'' and 日期 between cdate('" & dt1 & "') and cdate('" & dt2 & "') GROUP BY 印花单位"
Data4.Refresh

Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data3.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data3.Refresh
PZH = "5-" + Trim(Data3.Recordset.Fields(0) + 1)
Else
PZH = "5-1"
End If

Data4.Recordset.MoveFirst
KLLLL = 1

Do While Not Data4.Recordset.EOF
For i = 1 To 3
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = "加工费"
Data2.Recordset.Fields(1) = "原材料"
Data2.Recordset.Fields(2) = ""
Data2.Recordset.Fields(3) = "应付账款"
Data2.Recordset.Fields(4) = Data4.Recordset.Fields(0)
Data2.Recordset.Fields(5) = Format(Data4.Recordset.Fields(1) * 0.83, "#0.00")
Data2.Recordset.Fields(6) = PZH
Data2.Recordset.Fields(7) = CDate(dt3)
Data2.Recordset.Fields(8) = ""
Data2.Recordset.Fields(9) = ""
Data2.Recordset.Fields(10) = ""
Data2.Recordset.Fields(11) = "加工-材料"
Data2.Recordset.Update


Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = "加工费"
Data2.Recordset.Fields(1) = "应缴税金"
Data2.Recordset.Fields(2) = "税金进项"
Data2.Recordset.Fields(3) = "应付账款"
Data2.Recordset.Fields(4) = Data4.Recordset.Fields(0)
Data2.Recordset.Fields(5) = Format(Data4.Recordset.Fields(1) * 0.17, "#0.00")
Data2.Recordset.Fields(6) = PZH
Data2.Recordset.Fields(7) = CDate(dt3)
Data2.Recordset.Fields(8) = ""
Data2.Recordset.Fields(9) = ""
Data2.Recordset.Fields(10) = ""
Data2.Recordset.Fields(11) = "加工-材料"
Data2.Recordset.Update


Data4.Recordset.MoveNext
If Data4.Recordset.EOF Then
MsgBox ("材料加工单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data3.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data3.Refresh
PZH = "5-" + Trim(Data3.Recordset.Fields(0) + 1)
Else
PZH = "5-1"
End If
Loop
MsgBox ("材料加工单转账成功！" + "生成" + Str(KLLLL) + "凭证")
End If

End Sub



Private Sub RSJGPZ(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE instr(制单,'加工-材料')>0 AND 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data2.Refresh
If Not Data2.Recordset.EOF Then
If MsgBox("已有加工生成凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
Data3.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(制单,'加工-材料')>0 and 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data4.RecordSource = "select 染色单位,单号,材料名称,光坯克重,光坯幅宽,毛坯幅宽,毛坯重量,织布单价,金额,锅号,染厂色别,日期 from rsrk where 染色单位<>'' and 日期 between cdate('" & dt1 & "') and cdate('" & dt2 & "')"
Data4.Refresh
If Not Data4.Recordset.EOF Then
Data4.RecordSource = "select 染色单位,FORMAT(SUM(VAL(金额)),'#0.00') from rsrk where 染色单位<>'' and 日期 between cdate('" & dt1 & "') and cdate('" & dt2 & "') GROUP BY 染色单位"
Data4.Refresh

Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data3.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data3.Refresh
PZH = "5-" + Trim(Data3.Recordset.Fields(0) + 1)
Else
PZH = "5-1"
End If

Data4.Recordset.MoveFirst
KLLLL = 1

Do While Not Data4.Recordset.EOF
For i = 1 To 3
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = "加工费"
Data2.Recordset.Fields(1) = "原材料"
Data2.Recordset.Fields(2) = ""
Data2.Recordset.Fields(3) = "应付账款"
Data2.Recordset.Fields(4) = Data4.Recordset.Fields(0)
Data2.Recordset.Fields(5) = Format(Data4.Recordset.Fields(1) * 0.83, "#0.00")
Data2.Recordset.Fields(6) = PZH
Data2.Recordset.Fields(7) = CDate(dt3)
Data2.Recordset.Fields(8) = ""
Data2.Recordset.Fields(9) = ""
Data2.Recordset.Fields(10) = ""
Data2.Recordset.Fields(11) = "加工-材料"
Data2.Recordset.Update


Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = "加工费"
Data2.Recordset.Fields(1) = "应缴税金"
Data2.Recordset.Fields(2) = "税金进项"
Data2.Recordset.Fields(3) = "应付账款"
Data2.Recordset.Fields(4) = Data4.Recordset.Fields(0)
Data2.Recordset.Fields(5) = Format(Data4.Recordset.Fields(1) * 0.17, "#0.00")
Data2.Recordset.Fields(6) = PZH
Data2.Recordset.Fields(7) = CDate(dt3)
Data2.Recordset.Fields(8) = ""
Data2.Recordset.Fields(9) = ""
Data2.Recordset.Fields(10) = ""
Data2.Recordset.Fields(11) = "加工-材料"
Data2.Recordset.Update


Data4.Recordset.MoveNext
If Data4.Recordset.EOF Then
MsgBox ("材料加工单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data3.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data3.Refresh
PZH = "5-" + Trim(Data3.Recordset.Fields(0) + 1)
Else
PZH = "5-1"
End If
Loop
MsgBox ("材料加工单转账成功！" + "生成" + Str(KLLLL) + "凭证")
End If

End Sub



