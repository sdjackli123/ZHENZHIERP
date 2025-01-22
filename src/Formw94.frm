VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw94 
   BackColor       =   &H00C0E0FF&
   Caption         =   "成品库存"
   ClientHeight    =   11115
   ClientLeft      =   -435
   ClientTop       =   3810
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转库"
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号查询"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "日期查询"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "品名查询"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw94.frx":0000
      Height          =   390
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "品名"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全部库存"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw94.frx":0014
      Height          =   7575
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   9
      BackColorFixed  =   8421631
      BackColorBkg    =   50372
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
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
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   390
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   3015
      _ExtentX        =   5318
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   11040
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   39177
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "转库日期"
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
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
      Left            =   600
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
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
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Formw94"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer

Private Sub Command1_Click()
On Error Resume Next
Data2.RecordSource = "SELECT 单号,款号,品名,条码,型号,规格,单位,结存,入库,出库,库存 FROM  lskcmx order by 款号,品名,型号,规格,单号"
Data2.Refresh
       Data4.RecordSource = "SELECT 品名 FROM lskcmx GROUP BY 品名"
       Data4.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Data2.RecordSource = "SELECT 单号,款号,品名,条码,型号,规格,单位,结存,入库,出库,库存 FROM  lskcmx where instr(品名,'" & DBCombo1.Text & "')>0  order by 款号,品名,型号,规格,单号"
Data2.Refresh
End Sub

Private Sub Command4_Click()
If MsgBox("确定把库存转入到以往库存吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确实要把库存转入到以往库存吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("库存转入的库存记录日期为" + Trim(DTPicker3.Value), vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "DELETE * FROM LSJL WHERE 日期=CDATE('" & DTPicker3.Value & "')"
Data1.Database.Execute "insert into LSJL(单号,款号,品名,条码,规格,型号,单位,数量,日期,序号) select 单号,款号,品名,条码,规格,型号,单位,库存,cdate('" & DTPicker3.Value & "'),'1' from lskcmx"
MsgBox ("转入成功!,在库存记录中可以查询")
End Sub

Private Sub Command5_Click()
Call OutDataToExcel4(MSFlexGrid1, 8, 9, 10, 11, "成品库存")
End Sub

Private Sub Command6_Click()
If S1 = 0 Or S2 = 0 Then
MsgBox ("请选择记录！")
Exit Sub
End If


If DBCombo3.Text = "" Then
MsgBox ("输入类别")
Exit Sub
End If


If S1 < 1 Or S2 < 1 Then
MsgBox ("选择记录")
Exit Sub
End If

If S1 > S2 Then
MsgBox ("注意选择顺序！")
Exit Sub
End If

k = S2 - S1
If k = 0 Then
Data2.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data2.Recordset.Move S1 - 1
Data2.Recordset.Edit
Data2.Recordset.Fields(0) = DBCombo3.Text
Data2.Recordset.Update
Data2.Refresh

Else

Data2.Recordset.MoveFirst
Data2.Recordset.Move S1 - 1
For l = 1 To k + 1
Data2.Recordset.Edit
Data2.Recordset.Fields(0) = DBCombo3.Text
Data2.Recordset.Update
Data2.Recordset.MoveNext
Next
End If


Data2.Refresh


End Sub

Private Sub Command7_Click()
'On Error Resume Next
       Data1.Database.Execute "DELETE * FROM lskcmx"
       Data3.Database.Execute "INSERT INTO lskcmx(单号,款号,品名,规格,型号,单位,条码,出库) SELECT 单号,款号,品名,规格,型号,单位,条码,数量 FROM LSFH where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data1.Database.Execute "INSERT INTO lskcmx(单号,款号,品名,规格,型号,单位,条码,入库) SELECT 单号,款号,品名,规格,型号,单位,条码,数量 FROM LSRK where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data1.Database.Execute "INSERT INTO lskcmx(单号,款号,品名,规格,型号,单位,条码,入库) SELECT 单号,款号,品名,规格,型号,单位,条码,数量 FROM LSTK where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data1.Database.Execute "INSERT INTO lskcmx(单号,款号,品名,规格,型号,单位,条码,结存) SELECT 单号,款号,品名,规格,型号,单位,条码,数量 FROM LSJL where 日期=cdate('" & DTPicker1.Value & "')"
       Data1.Database.Execute "UPDATE lskcmx SET 审核='1'"
       Data1.Database.Execute "UPDATE lskcmx SET 出库='0' where 出库=null or 出库=''"
       Data1.Database.Execute "UPDATE lskcmx SET 入库='0' where 入库=null or 入库=''"
       Data1.Database.Execute "UPDATE lskcmx SET 结存=0 where 结存=null"
       Data1.Database.Execute "INSERT INTO lskcmx(单号,款号,品名,规格,型号,单位,条码,入库,出库,结存,库存) SELECT 单号,款号,品名,规格,型号,单位,条码,FORMAT(SUM(val(入库)),'#0'),format(sum(val(出库)),'#0'),format(sum(结存),'#0'),format(sum(val(入库)-val(出库)+结存),'#0') FROM lskcmx GROUP BY 单号,款号,品名,规格,型号,单位,条码"
       Data1.Database.Execute "DELETE * FROM lskcmx WHERE  审核='1'"
       Data2.RecordSource = "SELECT 单号,款号,品名,条码,型号,规格,单位,结存,入库,出库,库存 FROM  lskcmx  order by 款号,品名,型号,规格,单号"
       Data2.Refresh
End Sub

Private Sub Command8_Click()
Data2.RecordSource = "SELECT 单号,款号,品名,条码,型号,规格,单位,结存,入库,出库,库存 FROM  lskcmx where instr(款号,'" & DBCombo2.Text & "')>0  order by 款号,品名,型号,规格,单号"
Data2.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
DTPicker3.Value = Date
Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"

Data2.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"

Data3.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"

Data4.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"

Data5.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"

Data6.DatabaseName = "d:\数据库\\htgl\2011\CPCK"
Data6.RecordSource = "select mc from lb GROUP BY mc"
Data6.Refresh

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.ColWidth(6) = 1500

End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid1.RowSel
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid1.RowSel
End Sub


