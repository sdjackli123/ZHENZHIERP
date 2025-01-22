VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw49 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料盘点操作"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form49"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Width           =   6375
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0FF&
      Caption         =   "表打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0FF&
      Caption         =   "批次清空"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单号整理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3960
      Width           =   495
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0FF&
      Caption         =   "合并调整"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
      Width           =   495
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   6495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   9960
      TabIndex        =   18
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   81461249
      CurrentDate     =   39921
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "本月结存转入报表"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库类盘存打印"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "理论库存刷新"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "取平均单价"
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
      TabIndex        =   12
      Top             =   720
      Width           =   1095
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   1095
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw49.frx":0000
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
   End
   Begin VB.TextBox Text1111 
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "初始刷新"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "理论库存转实库"
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "本月结存转次月库"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "清空操作库"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "总库盘存打印"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw49.frx":0014
      Height          =   8055
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   14208
      _Version        =   393216
      BackColorFixed  =   8421631
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw49.frx":0028
      Height          =   330
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formw49.frx":003C
      Height          =   330
      Left            =   1560
      TabIndex        =   13
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formw49.frx":0050
      Height          =   330
      Left            =   3480
      TabIndex        =   24
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料规格"
      Text            =   "DBCombo1"
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结转日期"
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
      Left            =   9960
      TabIndex        =   25
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入规格"
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
      Left            =   3480
      TabIndex        =   23
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入颜色"
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
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入材料"
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
      Left            =   360
      TabIndex        =   9
      Top             =   600
      Width           =   1215
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
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Formw49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public r, c, FD As Integer: Public K1, K2 As String

Private Sub Command1_Click()
Data1.RecordSource = "SELECT KCBBLSH.材料名称,KCBBLSH.材料规格,KCBBLSH.材料单位,KCBBLSH.颜色,KCBBLSH.批次,KCBBLSH.单价,KCBBLSH.上月结存数量,KCBBLSH.上月结存金额,KCBBLSH.本月入库数量,KCBBLSH.本月入库金额,KCBBLSH.本月出库数量,KCBBLSH.本月出库金额,KCBBLSH.理论库存 as 本月结存数量,KCBBLSH.理论金额 AS 本月结存金额 from KCBBLSH ORDER BY KCBBLSH.材料名称,KCBBLSH.材料规格"
Data1.Refresh
Call OutDataToExcel4(MSFlexGrid1, 8, 10, 12, 14, "盘存打印")
End Sub

Private Sub Command10_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET 理论金额=format(上月结存金额+本月入库金额-本月出库金额,'#0.00'),理论库存=format(上月结存数量+本月入库数量-本月出库数量,'#0.00')"
Data1.RecordSource = "KCBBLSH"
Data1.Refresh
End Sub

Private Sub Command11_Click()
If DBCombo1.Text = "" Then
MsgBox ("选择库类")
Exit Sub
End If

Data1.RecordSource = "SELECT KCBBLSH.材料名称,KCBBLSH.材料规格,KCBBLSH.材料单位,KCBBLSH.颜色,KCBBLSH.批次,KCBBLSH.单价,KCBBLSH.上月结存数量,KCBBLSH.上月结存金额,KCBBLSH.本月入库数量,KCBBLSH.本月入库金额,KCBBLSH.本月出库数量,KCBBLSH.本月出库金额,KCBBLSH.理论库存 as 本月结存数量,KCBBLSH.理论金额 AS 本月结存金额 from KCBBLSH WHERE KCBBLSH.BL='" & DBCombo1.Text & "' ORDER BY KCBBLSH.材料名称,KCBBLSH.材料规格"
Data1.Refresh
FD = 9
Call OutDataToExcel3(MSFlexGrid1, 10, 12, 14, DBCombo1.Text + "  盘存打印")
End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command13_Click()
If MsgBox("请确认：次月日期为：" + Str(DTPicker1.Value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("请再确认：次月日期为：" + Str(DTPicker1.Value), vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "DELETE * FROM kcbbjl WHERE 盘存日期=CDATE('" & DTPicker1.Value & "')"
Data1.Database.Execute "INSERT INTO kcbbjl(单号,材料名称,材料规格,材料单位,颜色,批次,单价,上月结存数量,上月结存金额,本月入库数量,本月入库金额,本月出库数量,本月出库金额,BL,理论库存,理论金额,实际库存,实际金额)  SELECT 单号,材料名称,材料规格,材料单位,颜色,批次,单价,上月结存数量,上月结存金额,本月入库数量,本月入库金额,本月出库数量,本月出库金额,BL,理论库存,理论金额,实际库存,实际金额 FROM KCBBLSH"
Data1.Database.Execute "UPDATE kcbbjl SET 盘存日期=CDATE('" & DTPicker1.Value & "') where 盘存日期=null"
MsgBox ("操作成功！")
End Sub

Private Sub Command14_Click()

End Sub

Private Sub Command15_Click()
If MsgBox("确定调整合并吗？？", vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "DELETE * FROM KCBBLSH1"
Data1.Database.Execute "INSERT INTO KCBBLSH1(单号,材料名称,材料规格,材料单位,颜色,批次,日期,BL,上月结存数量,上月结存金额,本月入库数量,本月入库金额,本月出库数量,本月出库金额,理论库存,理论金额,实际库存,实际金额) SELECT 单号,材料名称,材料规格,材料单位,颜色,批次,日期,BL,format(SUM(上月结存数量),'#0.00'),format(SUM(上月结存金额),'#0.00'),format(SUM(本月入库数量),'#0.00'),format(SUM(本月入库金额),'#0.00'),format(SUM(本月出库数量),'#0.00'),format(SUM(本月出库金额),'#0.00'),format(SUM(理论库存),'#0.00'),format(SUM(理论金额),'#0.00'),format(SUM(实际库存),'#0.00'),format(SUM(实际金额),'#0.00') FROM KCBBLSH GROUP BY 单号,材料名称,材料规格,材料单位,颜色,批次,日期,BL"
Data1.Database.Execute "DELETE * FROM KCBBLSH"
Data1.Database.Execute "INSERT INTO KCBBLSH SELECT * FROM KCBBLSH1 "
MsgBox ("调整成功！！")
End Sub


Private Sub Command16_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET 单号=''"
Data1.Refresh
MsgBox ("整理成功！")
End Sub

Private Sub Command17_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET 批次=''"
Data1.Refresh
MsgBox ("批次清空成功！")
End Sub

Private Sub Command18_Click()
Call PCOutDataToExcel(MSFlexGrid1)
End Sub


Private Sub Command3_Click()
Data1.Database.Execute "DELETE * FROM KCBBLSH "
Data1.Refresh
End Sub

Private Sub Command4_Click()
On Error Resume Next

Data1.RecordSource = "SELECT * FROM KCBBLSH ORDER BY 材料名称,颜色"
Data1.Refresh
If Data1.Recordset.EOF Then
MsgBox ("无转库记录，终止")
Exit Sub
Else
If MsgBox("请确认：次月日期为：" + Str(DTPicker1.Value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定转库吗？", vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "delete * from kcjl where 日期=CDATE('" & DTPicker1.Value & "')"
Data1.Database.Execute "INSERT INTO  KCJL (单号,材料名称,材料规格,材料单位,颜色,批次,单价,数量,金额,BL)  SELECT 单号,KCBBLSH.材料名称,KCBBLSH.材料规格,KCBBLSH.材料单位,KCBBLSH.颜色,KCBBLSH.批次,KCBBLSH.单价,KCBBLSH.实际库存,KCBBLSH.实际金额,KCBBLSH.BL FROM KCBBLSH"
Data1.Database.Execute "UPDATE KCJL SET KCJL.日期=CDATE('" & DTPicker1.Value & "') WHERE kcjl.日期=NULL "
Data1.RecordSource = "kcjl"
Data1.Refresh
MsgBox ("转库成功！")
End If
End Sub

Private Sub Command5_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET 实际库存=理论库存"
Data1.Database.Execute "UPDATE KCBBLSH SET 实际金额=理论金额"
Data1.Refresh
MsgBox ("转库成功！")

End Sub

Private Sub Command6_Click()
On Error Resume Next
Data1.Database.Execute "UPDATE KCBBLSH SET 单号='' WHERE 单号=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET 材料名称='' WHERE 材料名称=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET 材料规格='' WHERE 材料规格=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET 材料单位='' WHERE 材料单位=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET 颜色='' WHERE 颜色=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET 批次='' WHERE 批次=NULL"
Data1.Database.Execute "UPDATE KCBBLSH SET BL='' WHERE BL=NULL"
Data1.RecordSource = "KCBBLSH"
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (Data1.Recordset.Fields(8) + Data1.Recordset.Fields(10)) > 0 Then
Data1.Recordset.Edit
Data1.Recordset.Fields(6) = (Data1.Recordset.Fields(9) + Data1.Recordset.Fields(11)) / (Data1.Recordset.Fields(8) + Data1.Recordset.Fields(10))
Data1.Recordset.Update
End If
Data1.Recordset.MoveNext
Loop
Data1.Refresh
End Sub

Private Sub Command7_Click()
Unload Me
End Sub


Private Sub Command9_Click()
Data1.Database.Execute "UPDATE KCBBLSH SET 单价=format((上月结存金额+本月入库金额-本月出库金额)/(上月结存数量+本月入库数量-本月出库数量),'#0.00') where (上月结存数量+本月入库数量-本月出库数量)<>0"
Data1.Database.Execute "UPDATE KCBBLSH SET 单价=0.00 where (上月结存数量+本月入库数量-本月出库数量)=0"
Data1.Refresh
End Sub

Private Sub DBCombo1_Change()
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM KCBBLSH ORDER BY 材料名称,颜色"
Data1.Refresh
Data3.RecordSource = "SELECT KCBBLSH.材料名称 FROM KCBBLSH  GROUP BY KCBBLSH.材料名称"
Data3.Refresh
Else
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE KCBBLSH.BL='" & DBCombo1.Text & "' ORDER BY 材料名称,颜色"
Data1.Refresh
Data3.RecordSource = "SELECT KCBBLSH.材料名称 FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' GROUP BY KCBBLSH.材料名称"
Data3.Refresh
End If

End Sub

Private Sub DBCombo1_Click(Area As Integer)
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM KCBBLSH ORDER BY 材料名称,颜色"
Data1.Refresh
Data3.RecordSource = "SELECT KCBBLSH.材料名称 FROM KCBBLSH  GROUP BY KCBBLSH.材料名称"
Data3.Refresh
Else
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE KCBBLSH.BL='" & DBCombo1.Text & "' ORDER BY 材料名称,颜色"
Data1.Refresh
Data3.RecordSource = "SELECT KCBBLSH.材料名称 FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' GROUP BY KCBBLSH.材料名称"
Data3.Refresh
End If
End Sub

Private Sub DBCombo2_Change()
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' AND INSTR(KCBBLSH.材料名称,'" & DBCombo2.Text & "')>0 ORDER BY 材料名称,颜色"
Data1.Refresh
Data7.RecordSource = "SELECT KCBBLSH.材料规格 FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' AND 材料名称='" & DBCombo2.Text & "' GROUP BY KCBBLSH.材料规格"
Data7.Refresh

End Sub

Private Sub DBCombo2_Click(Area As Integer)
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' AND INSTR(KCBBLSH.材料名称,'" & DBCombo2.Text & "')>0 ORDER BY 材料名称,颜色"
Data1.Refresh
Data7.RecordSource = "SELECT KCBBLSH.材料规格 FROM KCBBLSH WHERE BL='" & DBCombo1.Text & "' AND 材料名称='" & DBCombo2.Text & "' GROUP BY KCBBLSH.材料规格"
Data7.Refresh

End Sub

Private Sub DBCombo3_Change()
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE  INSTR(KCBBLSH.颜色,'" & DBCombo3.Text & "')>0 AND INSTR(KCBBLSH.材料名称,'" & DBCombo2.Text & "')>0 ORDER BY 材料名称,颜色"
Data1.Refresh
End Sub

Private Sub DBCombo3_Click(Area As Integer)
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE  INSTR(KCBBLSH.颜色,'" & DBCombo3.Text & "')>0 AND INSTR(KCBBLSH.材料名称,'" & DBCombo2.Text & "')>0 ORDER BY 材料名称,颜色"
Data1.Refresh
End Sub

Private Sub DBCombo4_Change()
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE  INSTR(材料规格,'" & DBCombo4.Text & "')>0 AND INSTR(KCBBLSH.材料名称,'" & DBCombo2.Text & "')>0 ORDER BY 材料名称,材料规格"
Data1.Refresh
End Sub

Private Sub DBCombo4_Click(Area As Integer)
Data1.RecordSource = "SELECT * FROM KCBBLSH WHERE  INSTR(材料规格,'" & DBCombo4.Text & "')>0 AND INSTR(KCBBLSH.材料名称,'" & DBCombo2.Text & "')>0 ORDER BY 材料名称,材料规格"
Data1.Refresh
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""
DBCombo4.Text = ""
DTPicker1.Value = Date
Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CKGL.MDB"
Data1.RecordSource = "KCBBLSH"
Data1.Refresh
Data2.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CKGL.MDB"
Data2.RecordSource = "SELECT KL.MC FROM KL GROUP BY KL.MC"
Data2.Refresh
Data3.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CKGL.MDB"
Data3.RecordSource = "SELECT KCBBLSH.材料名称 FROM KCBBLSH GROUP BY KCBBLSH.材料名称"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\bfrz\" + ljb + "\SCZYJHD.MDB"
Data4.RecordSource = "SELECT YS.YS FROM YS GROUP BY YS.YS"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\bfrz\" + ljb + "\SCZYJHD.MDB"
Data6.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.MDB"
Data6.RecordSource = "RQSD"
Data6.Refresh

Data7.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CKGL.MDB"
MSFlexGrid1.ColWidth(0) = 200

End Sub

Private Sub MSFlexGrid1_Click()
FD = MSFlexGrid1.Col
End Sub

Private Sub MSFlexGrid1_dblClick()
With MSFlexGrid1
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

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlexGrid1_dblClick
End If
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
Data1.Recordset.MoveFirst
Data1.Recordset.Move r - 1
Data1.Recordset.Edit
Data1.Recordset.Fields(c - 1) = Text1111.Text
Data1.Recordset.Update
Text1111.Visible = False
End Sub

