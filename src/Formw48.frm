VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Formw48 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料盘点"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   LinkTopic       =   "Form48"
   ScaleHeight     =   10695
   ScaleWidth      =   12090
   StartUpPosition =   2  '屏幕中心
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
      Top             =   0
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "本月入库"
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "本月出库"
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "本月结存"
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "盘存操作"
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
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
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "上月结存入本月库"
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
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
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw48.frx":0000
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Top             =   2400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "日期"
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw48.frx":0014
      Height          =   2415
      Left            =   960
      TabIndex        =   11
      Top             =   7560
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   7
      BackColorFixed  =   8421631
      BackColorBkg    =   43690
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   80150529
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   80150529
      CurrentDate     =   36892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formw48.frx":0028
      Height          =   2775
      Left            =   960
      TabIndex        =   14
      Top             =   4080
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   7
      BackColorFixed  =   8421631
      BackColorBkg    =   43690
      AllowUserResizing=   3
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   19
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   18
      Top             =   360
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
      Index           =   5
      Left            =   2040
      TabIndex        =   17
      Top             =   360
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   6240
      X2              =   6960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "库存记录信息"
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
      Left            =   960
      TabIndex        =   16
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "待结库信息"
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
      Index           =   1
      Left            =   960
      TabIndex        =   15
      Top             =   7080
      Width           =   1095
   End
End
Attribute VB_Name = "Formw48"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
DataCombo1.Visible = False
'DTPicker1.Visible = False
Command1.Visible = False

If MsgBox("请确定上月结存日期是否正确？", vbYesNo) = vbYes Then

Data1.Database.Execute "INSERT INTO  kcbb (单号,材料名称,材料规格,材料单位,颜色,批次,单价,日期,上月结存数量,上月结存金额,BL)  SELECT 单号,kcjl.材料名称,KCJL.材料规格,KCJL.材料单位,KCJL.颜色,KCJL.批次,KCJL.单价,KCJL.日期,KCJL.数量,KCJL.金额,KCJL.BL From KCJL WHERE KCJL.日期 = CDATE('" & DataCombo1.Text & "')"
Data1.Database.Execute "UPDATE KCBB SET BZH=0,本月入库数量=0,本月入库金额=0,本月出库数量=0,本月出库金额=0,日期=CDATE('" & Text1.Text & "')  WHERE BZH=null or bzh=''"
Data1.RecordSource = "KCBB"
Data1.Refresh
 
End If

End Sub

Private Sub Command10_Click()

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Data1.Database.Execute "INSERT INTO  kcbb(单号,BL,材料名称,材料规格,材料单位,颜色,批次,本月入库数量,本月入库金额)  SELECT 单号,库类,材料名称,材料规格,材料单位,颜色,批次,format(SUM(数量),'#0.00') AS S,format(SUM(合计金额),'#0.00') AS D From CKGL WHERE 库别='采购入库' AND 数量<>NULL AND  日期 BETWEEN '" & Text1.Text & "' AND '" & Text2.Text & "' GROUP BY 单号,库类,材料名称,材料规格,材料单位,颜色,批次"
Data1.Database.Execute "UPDATE KCBB SET 日期=CDATE('" & Text1.Text & "'),BZH=1,本月出库数量=0,本月出库金额=0,上月结存数量=0,上月结存金额=0 WHERE BZH=null or bzh=''"
Data1.RecordSource = "KCBB"
Data1.Refresh
End Sub
Private Sub Command5_Click()
'Data2.Database.Execute "INSERT INTO  kcbb (单号,BL,材料名称,材料规格,材料单位,颜色,批次,本月出库数量,本月出库金额)  SELECT 单号,库类,材料名称,材料规格,材料单位,颜色,批次,format(sum(数量),'#0.00') AS 出库数量,format(sum(合计金额),'#0.00') AS 金额合计  From KPD WHERE 标签<>'库存料' and  日期 BETWEEN '" & Text1.Text & "' AND '" & Text2.Text & "'  GROUP BY 单号,库类,材料名称,材料规格,材料单位,颜色,批次"
Data2.Database.Execute "INSERT INTO  kcbb (单号,BL,材料名称,材料规格,材料单位,颜色,批次,本月出库数量,本月出库金额)  SELECT 备注,库类,材料名称,材料规格,材料单位,颜色,批次,format(sum(数量),'#0.00') AS 出库数量,format(sum(合计金额),'#0.00') AS 金额合计  From KPD WHERE  日期 BETWEEN '" & Text1.Text & "' AND '" & Text2.Text & "'  GROUP BY 备注,库类,材料名称,材料规格,材料单位,颜色,批次"
Data1.Database.Execute "UPDATE KCBB SET 日期=CDATE('" & Text1.Text & "'),BZH=2,本月入库数量=0,本月入库金额=0,上月结存数量=0,上月结存金额=0 WHERE BZH=null or bzh=''"
Data1.RecordSource = "KCBB"
Data1.Refresh
End Sub

Private Sub Command6_Click()
On Error Resume Next
If MsgBox("请确定本月结存日期是否正确？", vbYesNo) = vbNo Then Exit Sub

Data3.Database.Execute "DELETE * FROM KCBBLSH "


Call Command4_Click
Call Command5_Click

Data3.Database.Execute "DELETE * FROM KCBB WHERE KCBB.上月结存数量=null OR KCBB.上月结存金额=NULL OR KCBB.本月入库数量=null OR KCBB.本月入库金额=null or KCBB.本月出库数量=null OR KCBB.本月出库金额=NULL"
Data3.Database.Execute "INSERT INTO KCBBLSH(单号,BL,材料名称,材料规格,材料单位,颜色,批次,本月出库数量,本月出库金额,本月入库数量,本月入库金额,上月结存数量,上月结存金额,理论库存,理论金额) SELECT 单号,KCBB.BL,KCBB.材料名称,KCBB.材料规格,KCBB.材料单位,KCBB.颜色,KCBB.批次,format(SUM(kcbb.本月出库数量),'#0.00') AS 本月出库数量,format(SUM(kcbb.本月出库金额),'#0.00') AS 本月出库金额,format(SUM(kcbb.本月入库数量),'#0.00') AS 本月入库数量,format(SUM(kcbb.本月入库金额),'#0.00') AS 本月入库金额,format(SUM(kcbb.上月结存数量),'#0.00') AS 上月结存数量,format(SUM(kcbb.上月结存金额),'#0.00') AS 上月结存金额,format(SUM(KCBB.上月结存数量+KCBB.本月入库数量-kcbb.本月出库数量),'#0.00') AS 本月结存数量,format(SUM(KCBB.上月结存金额+KCBB.本月入库金额-kcbb.本月出库金额),'#0.00') AS 本月结存金额 FROM KCBB GROUP BY 单号,KCBB.BL,KCBB.材料名称,KCBB.材料规格,KCBB.材料单位,KCBB.颜色,KCBB.批次"
Data3.Database.Execute "UPDATE KCBBLSH SET 日期=CDATE('" & Text2.Text & "') "
'Data3.Database.Execute "DELETE * FROM KCBBLSH WHERE KCBBLSH.上月结存数量=null OR KCBBLSH.上月结存金额=NULL OR KCBBLSH.本月入库数量=null OR KCBBLSH.本月入库金额=null or KCBBLSH.本月出库数量=null OR KCBBLSH.本月出库金额=NULL"
Data3.RecordSource = "KCBBLSH"
Data3.Refresh
Data1.Database.Execute "DELETE * FROM  kcbb"
End Sub



Private Sub Command8_Click()
Unload Me
Formw49.Show
End Sub

Private Sub Command9_Click()

DataCombo1.Visible = True
Command1.Visible = True

Data3.RecordSource = "SELECT KCJL.日期 FROM KCJL GROUP BY KCJL.日期"
Data3.Refresh
End Sub

Private Sub DTPicker3_Change()
Text1.Text = DTPicker3.Value
Text1.SetFocus
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = DTPicker3.Value
End Sub

Private Sub DTPicker4_Change()
Text2.Text = DTPicker4.Value
Text2.SetFocus
End Sub

Private Sub DTPicker4_CloseUp()
Text2.Text = DTPicker4.Value
End Sub

Private Sub Form_Load()
DataCombo1.Text = ""
DataCombo1.Visible = False
Command1.Visible = False
Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\MDB"
Data1.Refresh
Text1.Text = Date
Text2.Text = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
Data2.DatabaseName = "d:\数据库\bfrz\" + ljb + "\MDB"
Data2.Refresh

Data3.DatabaseName = "d:\数据库\bfrz\" + ljb + "\MDB"
Data3.RecordSource = "KCBBLSH"
Data3.Refresh


End Sub

Private Sub JILU()
On Error Resume Next
Dim i As Single

If Data2.Recordset.EOF Then
Else
Data2.Recordset.MoveFirst
For i = 1 To Data2.Recordset.RecordCount
VSFlexGrid2.Col = 7

VSFlexGrid2.Row = i
VSFlexGrid2.Text = Format(Data2.Recordset.Fields(5), "0.0")
Data2.Recordset.MoveNext
Next
End If

If Data1.Recordset.RecordCount = 0 Then
Else
Data1.Recordset.MoveFirst
For i = 1 To Data1.Recordset.RecordCount
VSFlexGrid1.Col = 7

VSFlexGrid1.Row = i
VSFlexGrid1.Text = Format(Data1.Recordset.Fields(6), "0.0")
Data1.Recordset.MoveNext
Next
End If

End Sub



