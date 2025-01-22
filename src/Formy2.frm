VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy2 
   BackColor       =   &H00C0E0FF&
   Caption         =   "计划单"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "Formy2.frx":0000
   LinkTopic       =   "计划单"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   1680
      Style           =   1  'Simple Combo
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1320
      Width           =   855
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Left            =   1440
      TabIndex        =   16
      Top             =   2040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy2.frx":440A
      Height          =   330
      Left            =   1440
      TabIndex        =   15
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "下单"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9480
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   9090
      Left            =   10680
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "进度"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.Data Data15 
      Caption         =   "Data15"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
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
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Data Data8 
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
      Top             =   10800
      Visible         =   0   'False
      Width           =   4695
   End
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Width           =   5055
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Width           =   5415
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command6 
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
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成"
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy2.frx":441E
      Height          =   6495
      Left            =   600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3000
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11456
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   36892
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "交期"
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   18
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "日期"
      Height          =   375
      Index           =   16
      Left            =   3720
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  成衣制作(样品)工艺单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "Formy2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer
Private Sub Command12_Click()
Unload Me
Formy4.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command1_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Command2_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub Command4_Click()
On Error Resume Next
Data3.RecordSource = "select MAX(VAL(MID(单号,8))) from SCZY_X WHERE INSTR(单号,'L')>0 AND 日期=CDATE('" & DTPicker1.Value & "')"
Data3.Refresh
DBCombo2.Text = "L" + Trim(Format(Date, "YYMMDD")) + "1"
If Data3.Recordset.EOF Then
DBCombo2.Text = "L" + Trim(Format(Date, "YYMMDD")) + "1"
Else
DBCombo2.Text = "L" + Trim(Format(Date, "YYMMDD")) + Trim(Data3.Recordset.Fields(0) + 1)
End If
End Sub

Private Sub Command5_Click()
Data6.Database.Execute "update sczy_x set 数量=(val(数量1)+val(数量2)+val(数量3)+val(数量4)+val(数量5)+val(数量6)+val(数量7)+val(数量8)) where 单号='" & DBCombo2.Text & "'"
Data1.RecordSource = "select * from SCZY_x WHERE 单号='" & DBCombo2.Text & "' ORDER BY 款号,序号"
Data1.Refresh
End Sub

Private Sub Command6_Click()
Call ypdy(MSFlexGrid1, "样品：" + DBCombo2.Text)
Data1.RecordSource = "select * from SCZY_x WHERE 单号='" & DBCombo2.Text & "' ORDER BY 款号,序号"
Data1.Refresh
End Sub


Private Sub Command8_Click()
If MsgBox("确定生成吗？", vbYesNo) = vbNo Then Exit Sub
Data6.Database.Execute "delete * from sczy_x  where 单号='" & DBCombo2.Text & "'"
For i = 0 To List1.ListCount - 1
l1 = List1.List(i)
l2 = List1.List(i)
If List1.Selected(i) = True Then
Data6.Database.Execute "insert into sczy_x(单号,款号,颜色,数量,单位,品名,尺码1,数量1,尺码2,数量2,尺码3,数量3,尺码4,数量4,尺码5,数量5,尺码6,数量6,尺码7,数量7,尺码8,数量8,备注,日期,进度,图片,客户,交期,序号) select '" & DBCombo2.Text & "',款号,颜色,'0',单位,品名,尺码1,'',尺码2,'',尺码3,'',尺码4,'',尺码5,'',尺码6,'',尺码7,'',尺码8,'',备注,'" & DTPicker1.Value & "','进行','','" & DBCombo1.Text & "','" & DTPicker2.Value & "',序号 from ksnr where 款号='" & l1 & "'"
End If
Next
MsgBox ("生成成功！")
Data1.RecordSource = "select * from SCZY_x WHERE 单号='" & DBCombo2.Text & "' ORDER BY 款号,序号"
Data1.Refresh
End Sub

Private Sub Command9_Click()
On Error Resume Next
Data6.Database.Execute "insert into ypjd(客户,单号,款号,颜色,品名,数量,接单日期,样品交期) in'd:\数据库\\htgl\2011\scjd.mdb' SELECT 客户,单号,款号,颜色,品名,SUM(数量),日期,交期 FROM SCZY_X WHERE 单号='" & DBCombo2.Text & "' GROUP BY 客户,单号,款号,颜色,品名,日期,交期"
Data5.Database.Execute "UPDATE ypjd SET 进度='进行' where 单号='" & DBCombo2.Text & "'"
MsgBox ("列入样品进度！")
End Sub

Private Sub DBCombo2_Change()
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from SCZY_x WHERE 单号='" & DBCombo2.Text & "' ORDER BY 款号,序号"
Data1.Refresh
End Sub

Private Sub DBCombo2_Click(Area As Integer)
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from SCZY_x WHERE 单号='" & DBCombo2.Text & "' ORDER BY 款号,序号"
Data1.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next


Text1.Text = ""
DBCombo1.Text = ""
DBCombo2.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date

Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from SCZY_x WHERE 单号='" & DBCombo2.Text & "' ORDER BY 款号,序号"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"


Data3.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data3.RecordSource = "select MAX(VAL(MID(单号,8))) from SCZY_X WHERE INSTR(单号,'L')>0 AND 日期=CDATE('" & DTPicker1.Value & "')"
Data3.Refresh
DBCombo2.Text = "L" + Trim(Format(Date, "YYMMDD")) + "1"
If Data3.Recordset.EOF Then
DBCombo2.Text = "L" + Trim(Format(Date, "YYMMDD")) + "1"
Else
DBCombo2.Text = "L" + Trim(Format(Date, "YYMMDD")) + Trim(Data3.Recordset.Fields(0) + 1)
End If

Data4.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data4.RecordSource = "SELECT 简称  from khzl GROUP BY 简称"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"

Data6.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 1200
DBCombo1.TabIndex = 0
End Sub

Private Sub Label2_DBLClick(Index As Integer)
Select Case Index
   Case 18
   Formy38.Show
   End Select
End Sub

Private Sub Label3_dblClick(Index As Integer)
Select Case Index
       Case 7
DBCombo6.Enabled = True
End Select
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
If Data1.Recordset.EOF Then
Exit Sub
End If
If c = 0 Then
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
Data1.Recordset.Delete
Data1.Refresh
End If

End Sub

Private Sub Text1_Change()
Data2.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data2.RecordSource = "select distinct 款号 from ksnr where instr(款号,'" & Text1.Text & "')>0 order by 款号"
Data2.Refresh
List1.Clear
If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
List1.AddItem Data2.Recordset.Fields(0)
Data2.Recordset.MoveNext
Loop
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub MSFlex()
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
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    MSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data1.Recordset.MoveFirst
Data1.Recordset.Move r - 1
Data1.Recordset.Edit
Data1.Recordset.Fields(c - 1) = Combo1111.Text
Data1.Recordset.Update
MSFlexGrid1.Text = Combo1111.Text
Combo1111.Visible = False
MSFlexGrid1.SetFocus
End If
End Sub



