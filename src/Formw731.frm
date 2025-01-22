VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Formw731 
   BackColor       =   &H00C0E0FF&
   Caption         =   "染化客账查询"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form39"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成查询"
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "凭证生成"
      Height          =   855
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.Data Data10 
      Caption         =   "Data8"
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "结转清除"
      Height          =   375
      Left            =   5760
      TabIndex        =   17
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印准备"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "结转下期"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   840
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
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
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
      Top             =   9360
      Visible         =   0   'False
      Width           =   3975
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
      Top             =   9360
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw731.frx":0000
      Height          =   7455
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   10790143
      BackColorBkg    =   44718
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw731.frx":0014
      Height          =   330
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   80871425
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   80871425
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   80871425
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   12000
      TabIndex        =   20
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   80871425
      CurrentDate     =   36892
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      Height          =   375
      Index           =   2
      Left            =   12000
      TabIndex        =   21
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "下期起初"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   15
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "单位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择日期范围"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Formw731"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer

Private Sub Command1_Click()
'On Error Resume Next
Command1.Enabled = False

Data6.Database.Execute "DELETE * FROM JGZCX2"
lo = "d:\数据库\bfrz\" + ljb + "\FP.MDB"       '''''''''''''''''''''''经典
''''   考察Data4.Database.Execute "insert into JGZCX2(客户,上期累计应付) IN'" & LO &"' SELECT MID(会计科目,INSTR(会计科目,'-')+1),format(SUM(VAL(余额)),'#0.00') FROM PMMXJZ WHERE 借贷方向='贷' AND 日期=CDATE('" & RQQ & "') GROUP BY MID(会计科目,INSTR(会计科目,'-')+1)"
Data4.Database.Execute "insert into JGZCX2(客户,上期累计应付) IN'" & lo & "' SELECT MID(会计科目,INSTR(会计科目,'-')+1),format(SUM(VAL(余额)),'#0.00') FROM PMMXJZ WHERE 借贷方向='贷' AND 日期=CDATE('" & Text1.Text & "') GROUP BY MID(会计科目,INSTR(会计科目,'-')+1)"
'Data3.Database.Execute "insert into JGZCX2(客户,本期应付款) in'" & LO & "' SELECT 供应单位,format(SUM(合计金额),'#0.00') FROM CKGL WHERE  日期 between cdate('" & Text1 & "') and cdate('" & Text2.text & "') AND 库别='采购入库' GROUP BY 供应单位"
Data5.Database.Execute "insert into JGZCX2(客户,本期应付款) in'" & lo & "' SELECT 供应单位,format(SUM(合计金额),'#0.00') FROM MX WHERE  入库时间 between cdate('" & Text1 & "') and cdate('" & Text2.Text & "') AND 库别='采购入库' GROUP BY 供应单位"
Data5.Database.Execute "insert into JGZCX2(客户,本期应付款) in'" & lo & "' SELECT 出库单位,format(SUM(-合计金额),'#0.00') FROM ckMX WHERE  出库时间 between cdate('" & Text1 & "') and cdate('" & Text2.Text & "')  GROUP BY 出库单位"
Data3.Database.Execute "insert into JGZCX2(客户,本期应付款) in'" & lo & "' SELECT 供应单位,format(SUM(合计金额),'#0.00') FROM CKGL WHERE  日期 between cdate('" & Text1 & "') and cdate('" & Text2.Text & "') AND 库别='采购入库' and 是否付款<>'已付' GROUP BY 供应单位"
rqq = CDate(Text2.Text) + 1
Data6.Database.Execute "insert into JGZCX2(客户,本期开票)  SELECT 客户,开票金额 FROM JHFP WHERE 开票日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & rqq & "')"
Data4.Database.Execute "insert into JGZCX2(客户,本期已付款) IN'" & lo & "' SELECT MID(对方科目,INSTR(对方科目,'-')+1),format(SUM(VAL(贷方金额)),'#0.00') FROM TZJZMX WHERE 日期 between cdate('" & Text1.Text & "') and cdate('" & Text2.Text & "') AND 贷方金额<>'0' GROUP BY MID(对方科目,INSTR(对方科目,'-')+1)"
Data6.Database.Execute "insert into JGZCX2(客户,上期累计开票) SELECT 客户,开票金额 FROM PMJHFP WHERE  结转日期=CDATE('" & Text1.Text & "')"
Data6.Database.Execute "insert into JGZCX2(客户,上期累计未开票) SELECT 客户,未开金额 FROM PMJHFP WHERE  结转日期=CDATE('" & Text1.Text & "')"

Data4.Database.Execute "insert into JGZCX2(客户,未达账) IN'" & lo & "' SELECT MID(对方科目,INSTR(对方科目,'-')+1),format(SUM(VAL(贷方金额)),'#0.00') FROM TZJZMX WHERE 日期 between cdate('" & Text1.Text & "') and cdate('" & Text2.Text & "') AND 贷方金额<>'0' GROUP BY MID(对方科目,INSTR(对方科目,'-')+1)"
Data4.Database.Execute "insert into JGZCX2(客户,未达账) IN'" & lo & "' SELECT 客户,format(SUM(VAL(余额)),'#0.00') FROM WDZSZ WHERE 日期=cdate('" & Text1.Text & "')  GROUP BY 客户"
Data6.Database.Execute "insert into JGZCX2(客户,未达账) SELECT 客户,format(SUM(VAL(开票金额)),'#0.00') FROM JHFP WHERE  开票日期 between cdate('" & Text1.Text & "') and cdate('" & rqq & "') GROUP BY 客户"


Data6.Database.Execute "UPDATE JGZCX2 SET 类别='1'"
Data6.Database.Execute "UPDATE JGZCX2 SET 日期范围='" & Text1.Text & "'+'--'+'" & Text2.Text & "'"
Data6.Database.Execute "UPDATE JGZCX2 SET 上期累计应付='0' WHERE 上期累计应付=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 本期应付款='0' WHERE 本期应付款=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 本期累计应付款='0' WHERE 本期累计应付款=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 本期已付款='0' WHERE 本期已付款=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 上期累计开票='0' WHERE 上期累计开票=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 本期开票='0' WHERE 本期开票=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 上期累计未开票='0' WHERE 上期累计未开票=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 本期累计开票='0' WHERE 本期累计开票=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 本期未开='0' WHERE 本期未开=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 本期累计未开='0' WHERE 本期累计未开=NULL"
Data6.Database.Execute "UPDATE JGZCX2 SET 未达账='0' WHERE 未达账=NULL"

Data6.Database.Execute "insert into JGZCX2(客户,日期范围,上期累计应付,本期应付款,本期累计应付款,本期已付款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期未开,本期累计未开,未达账) SELECT 客户,日期范围,FORMAT(SUM(VAL(上期累计应付)),'#0.00'),FORMAT(SUM(VAL(本期应付款)),'#0.00'),FORMAT(SUM(VAL(本期累计应付款)),'#0.00'),FORMAT(SUM(VAL(本期已付款)),'#0.00'),FORMAT(SUM(VAL(上期累计开票)),'#0.00'),FORMAT(SUM(VAL(本期开票)),'#0.00'),FORMAT(SUM(VAL(本期累计开票)),'#0.00'),FORMAT(SUM(VAL(上期累计未开票)),'#0.00'),FORMAT(SUM(VAL(本期未开)),'#0.00'),FORMAT(SUM(VAL(本期累计未开)),'#0.00'),FORMAT(SUM(VAL(未达账)),'#0.00') FROM JGZCX2 GROUP BY 客户,日期范围 "
Data6.Database.Execute "DELETE *  FROM  JGZCX2 WHERE 类别='1'"
Data6.Database.Execute "UPDATE JGZCX2 SET 本期未开=FORMAT(VAL(本期应付款)-VAL(本期开票),'#0.00')"
Data6.Database.Execute "UPDATE JGZCX2 SET 欠款=FORMAT(VAL(上期累计应付)+VAL(本期应付款)-VAL(本期已付款),'#0.00'),本期累计应付款=FORMAT(VAL(上期累计应付)+VAL(本期应付款),'#0.00'),本期累计开票=FORMAT(VAL(上期累计开票)+VAL(本期开票),'#0.00'),本期累计未开=FORMAT(VAL(上期累计未开票)+VAL(本期未开),'#0.00')"
Data6.Database.Execute "DELETE *  FROM  JGZCX2 WHERE val(本期应付款)=0 and val(本期已付款)=0 and val(欠款)=0"
 
 
Data8.RecordSource = "select 简称 from GYS WHERE INSTR(传真,'R')>0"
Data8.Refresh

Data6.RecordSource = "SELECT 客户 FROM JGZCX2"
Data6.Refresh

If Not Data6.Recordset.EOF Then
Data6.Recordset.MoveFirst
Do While Not Data6.Recordset.EOF
Data8.Recordset.FindFirst "简称='" & Data6.Recordset.Fields(0) & "'"
If Data8.Recordset.NoMatch Then
Data9.Database.Execute "DELETE *  FROM  JGZCX2 WHERE 客户='" & Data6.Recordset.Fields(0) & "'"
End If
Data6.Recordset.MoveNext
Loop
End If
Command1.Enabled = True

Data6.RecordSource = "SELECT 客户,上期累计应付,本期应付款,本期累计应付款,本期已付款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期累计未开,欠款,未达账,日期范围 FROM JGZCX2  order by 客户"
Data6.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call OutDataToExcel11(VSFlexGrid1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, "枣庄宝隆针织制衣有限公司 客户账目查询--付款" + "截止日期:" + Text2.Text)
End Sub

Private Sub Command4_Click()
Data9.Database.Execute "update JGZCX2 set 上期累计应付='' where 上期累计应付='0.00'"
Data9.Database.Execute "update JGZCX2 set 本期应付款='' where 本期应付款='0.00'"
Data9.Database.Execute "update JGZCX2 set 本期累计应付款='' where 本期累计应付款='0.00'"
Data9.Database.Execute "update JGZCX2 set 本期已付款='' where 本期已付款='0.00'"
Data9.Database.Execute "update JGZCX2 set 上期累计开票='' where 上期累计开票='0.00'"
Data9.Database.Execute "update JGZCX2 set 本期开票='' where 本期开票='0.00'"
Data9.Database.Execute "update JGZCX2 set 本期累计开票='' where 本期累计开票='0.00'"
Data9.Database.Execute "update JGZCX2 set 上期累计未开票='' where 上期累计未开票='0.00'"
Data9.Database.Execute "update JGZCX2 set 本期累计未开='' where 本期累计未开='0.00'"
Data9.Database.Execute "update JGZCX2 set 欠款='' where 欠款='0.00'"
Data9.Database.Execute "update JGZCX2 set 上期累计应付='' where 上期累计应付='0.00'"
Data6.Refresh
End Sub

Private Sub Command5_Click()
'On Error Resume Next
If MsgBox("确定结转下期吗，下期起初为：" + Trim(DTPicker1.Value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定结转下期吗?", vbYesNo) = vbNo Then Exit Sub

lo = "d:\数据库\bfrz\" + ljb + "\zcw.mdb"
Data6.RecordSource = "SELECT 客户,上期累计应付,本期应付款,本期累计应付款,本期已付款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期累计未开,欠款,日期范围 FROM JGZCX2  order by 客户"
Data6.Refresh

If Not Data6.Recordset.EOF Then
Data6.Recordset.MoveFirst
Do While Not Data6.Recordset.EOF
Data10.Database.Execute "delete * from  PMMXJZ where instr(摘要,'染料')>0 and 日期='" & DTPicker1.Value & "' and instr(会计科目,'应付账款')>0 and mid(会计科目,instr(会计科目,'-')+1)='" & Data6.Recordset.Fields(0) & "'"
Data9.Database.Execute "INSERT INTO PMMXJZ(会计科目,余额) in'" & lo & "' SELECT '应付账款-'+trim(客户) as ll,欠款 from JGZCX2 where 客户='" & Data6.Recordset.Fields(0) & "'"
Data10.Database.Execute "update PMMXJZ set 摘要='期初余额染料',凭证号='结',借贷方向='贷',序号='1',日期='" & DTPicker1.Value & "' where 日期=null"

Data9.Database.Execute "delete * from  PMJHFP where  结转日期='" & DTPicker1.Value & "' and 客户='" & Data6.Recordset.Fields(0) & "'"
Data9.Database.Execute "insert into PMJHFP(客户,开票金额,未开金额) select 客户,本期累计开票,本期累计未开 from JGZCX2 where 客户='" & Data6.Recordset.Fields(0) & "'"
Data9.Database.Execute "update PMJHFP set 结转日期='" & DTPicker1.Value & "' where 结转日期=null"
Data6.Recordset.MoveNext
Loop
End If



MsgBox ("结转成功！，在期初设置中可以查询！")
End Sub

Private Sub Command6_Click()
Data10.Database.Execute "delete * from  PMMXJZ where 日期='" & DTPicker1.Value & "' and instr(会计科目,'应付账款')>0"
Data9.Database.Execute "delete * from  PMJHFP where 结转日期='" & DTPicker1.Value & "'"
MsgBox ("清除成功!")
End Sub

Private Sub Command7_Click()
If MsgBox("操作日期为：" + Trim(DTPicker2.Value) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("操作期间为：" + Trim(Month(DTPicker2.Value)) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定生成应付系列的凭证吗？", vbYesNo) = vbNo Then Exit Sub
Call CLRKPZ(CDate(Text1.Text), CDate(Text2.Text), CDate(DTPicker2.Value))
End Sub

Private Sub Command8_Click()
Form1132.DTPicker1.Value = DTPicker2.Value
Form1132.Show
Unload Me
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If DataCombo1.Text = "" Then
Data6.RecordSource = "SELECT 客户,上期累计应付,本期应付款,本期累计应付款,本期已付款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期累计未开,欠款,未达账,日期范围 FROM JGZCX2  order by 客户"
Data6.Refresh
Else
Data6.RecordSource = "SELECT 客户,上期累计应付,本期应付款,本期累计应付款,本期已付款,上期累计开票,本期开票,本期累计开票,上期累计未开票,本期累计未开,欠款,未达账,日期范围 FROM JGZCX2 WHERE 客户='" & DataCombo1.Text & "' and val(欠款)<>0 order by 客户"
Data6.Refresh
End If
End Sub

Private Sub DTPicker3_Change()
Text1.Text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = DTPicker3.Value
Text1.SetFocus
End Sub

Private Sub DTPicker4_Change()
Text2.Text = DTPicker4.Value
End Sub

Private Sub DTPicker4_CloseUp()
Text2.Text = DTPicker4.Value
Text2.SetFocus
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption + "操作年度 " + ljb
Text1.Text = Date
Text2.Text = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
DataCombo1.Text = ""
Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\SCZYJHD.mdb"
Data1.RecordSource = "select GYS.简称 from GYS  GROUP BY 简称"
Data1.Refresh
Data2.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CLCK.MDB"
Data2.RecordSource = "select 供应单位,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,单据号,日期,是否开票,开票,开票日期 from ckgl where 供应单位='" & DataCombo1.Text & "' and 日期 between cdate('" & Text1 & "') and cdate('" & Text2.Text & "') AND 库别='采购入库'"
Data2.Refresh
Data3.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CLCK.MDB"
Data4.DatabaseName = "d:\数据库\bfrz\" + ljb + "\ZCW.MDB"
Data5.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.MDB"
Data6.DatabaseName = "d:\数据库\bfrz\" + ljb + "\FP.MDB"
Data7.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.MDB"
Data7.RecordSource = "rqsd"
Data7.Refresh

Data8.DatabaseName = "d:\数据库\bfrz\" + ljb + "\SCZYJHD.mdb"
Data9.DatabaseName = "d:\数据库\bfrz\" + ljb + "\fp.MDB"
Data10.DatabaseName = "d:\数据库\bfrz\" + ljb + "\zcw.MDB"

For i = 2 To 12
VSFlexGrid1.ColWidth(i) = 1200
Next
VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(13) = 2200

End Sub

Private Sub vSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
S1 = VSFlexGrid1.RowSel
End Sub

Private Sub vSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
S2 = VSFlexGrid1.RowSel
End Sub

Private Sub CLRKPZ(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next
If InStr(ljb, "wx") > 0 Then
Data4.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "') and instr(制单,'自动-染料')>0"
Data4.Refresh
If Not Data4.Recordset.EOF Then
If MsgBox("已有应付生成凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
Data10.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(制单,'自动-染料')>0 and 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data9.RecordSource = "SELECT * FROM JGZCX2 where val(本期应付款)<>0"
Data9.Refresh

If Data9.Recordset.EOF Then Exit Sub
Data4.RecordSource = "SELECT * FROM CLZZPZ"
Data4.Refresh

Data10.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
If Data10.Recordset.EOF Then
PZH = "R5-1"
Else
Data10.RecordSource = "SELECT MAX(VAL(MID(凭证号,4))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
PZH = "R5-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 7
Data4.Recordset.AddNew
Data4.Recordset.Fields(0) = "购染料"
Data4.Recordset.Fields(1) = "原材料"
Data4.Recordset.Fields(2) = ""
Data4.Recordset.Fields(3) = "应付账款"
Data4.Recordset.Fields(4) = Data9.Recordset.Fields(0)
Data4.Recordset.Fields(5) = Format(Data9.Recordset.Fields(2), "#0.00")
Data4.Recordset.Fields(6) = PZH
Data4.Recordset.Fields(7) = CDate(dt3)
Data4.Recordset.Fields(8) = ""
Data4.Recordset.Fields(9) = ""
Data4.Recordset.Fields(10) = ""
Data4.Recordset.Fields(11) = "自动-染料"
Data4.Recordset.Update


Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
Data10.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
If Data10.Recordset.EOF Then
PZH = "R5-1"
Else
Data10.RecordSource = "SELECT MAX(VAL(MID(凭证号,4))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
PZH = "R5-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
End If


If InStr(ljb, "nx") > 0 Then
Data4.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "') and instr(制单,'自动-染料')>0"
Data4.Refresh
If Not Data4.Recordset.EOF Then
If MsgBox("已有应付生成凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
Data10.Database.Execute "DELETE * FROM CLZZPZ WHERE instr(制单,'自动-染料')>0 and 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
End If

Data9.RecordSource = "SELECT * FROM JGZCX2 where val(本期应付款)>0"
Data9.Refresh

If Data9.Recordset.EOF Then Exit Sub
Data4.RecordSource = "SELECT * FROM CLZZPZ"
Data4.Refresh

Data10.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
If Data10.Recordset.EOF Then
PZH = "I5-1"
Else
Data10.RecordSource = "SELECT MAX(VAL(MID(凭证号,4))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
PZH = "I5-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 7
Data4.Recordset.AddNew
Data4.Recordset.Fields(0) = "购染料"
Data4.Recordset.Fields(1) = "原材料"
Data4.Recordset.Fields(2) = ""
Data4.Recordset.Fields(3) = "应付账款"
Data4.Recordset.Fields(4) = Data9.Recordset.Fields(0)
Data4.Recordset.Fields(5) = Format(Data9.Recordset.Fields(2), "#0.00")
Data4.Recordset.Fields(6) = PZH
Data4.Recordset.Fields(7) = CDate(dt3)
Data4.Recordset.Fields(8) = ""
Data4.Recordset.Fields(9) = ""
Data4.Recordset.Fields(10) = ""
Data4.Recordset.Fields(11) = "自动-染料"
Data4.Recordset.Update


Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
Data10.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
If Data10.Recordset.EOF Then
PZH = "I5-1"
Else
Data10.RecordSource = "SELECT MAX(VAL(MID(凭证号,4))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data10.Refresh
PZH = "I5-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
End If


End Sub





