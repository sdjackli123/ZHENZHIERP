VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Forml605 
   BackColor       =   &H00C0E0FF&
   Caption         =   "传票卡操作"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form41"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   4200
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   5280
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "条码打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "包装条码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "工序查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "A4/2打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "传票作废"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   5880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data6 
      Caption         =   "Data5"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "下号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   6480
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   7320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "A4打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成传票"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Forml605.frx":0000
      Height          =   4695
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4320
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Forml605.frx":0014
      Height          =   3255
      Left            =   8040
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   3240
      TabIndex        =   18
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单位"
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
      Index           =   9
      Left            =   4200
      TabIndex        =   32
      Top             =   1680
      Width           =   1095
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
      Index           =   8
      Left            =   5280
      TabIndex        =   30
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
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
      Index           =   7
      Left            =   2880
      TabIndex        =   28
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "卡号"
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
      Left            =   3240
      TabIndex        =   17
      Top             =   600
      Width           =   1935
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
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
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
      Left            =   1440
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
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
      Left            =   5400
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格"
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
      Index           =   5
      Left            =   6480
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "数量"
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
      Index           =   6
      Left            =   7320
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "输入款号"
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
      Width           =   2775
   End
End
Attribute VB_Name = "Forml605"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gxk, c, r As Integer


Private Sub Combo1_Click()
Text1(9).Text = Combo1.Text
End Sub
Private Sub Command1_Click()
Data1.RecordSource = "select * from CPK where 卡号='" & DBCombo1.Text & "' order by 编号"
Data1.Refresh
Data2.Refresh
If Text1(3).Text <> "" Then
Data4.Database.Execute "delete * from cjfz"
Data4.Database.Execute "insert into CJFZ(单号,款号,颜色,规格,裁剪,缝制,欠发) select 单号,款号,颜色,规格,sum(val(裁剪)),'0','0' from cjrb where 款号='" & Text1(1).Text & "' and 颜色='" & Text1(2).Text & "' and 规格='" & Text1(3).Text & "' group by 单号,款号,颜色,规格"
Data4.Database.Execute "insert into CJFZ(共量,单号,款号,颜色,规格,裁剪,缝制,欠发) select distinct 卡号,单号,款号,颜色,规格,'0',数量,'0' from cpk where 款号='" & Text1(1).Text & "' and 颜色='" & Text1(2).Text & "' and 规格='" & Text1(3).Text & "'"
Data4.Database.Execute "update cjfz set 共量='1'"
Data4.Database.Execute "insert into CJFZ(单号,款号,颜色,规格,裁剪,缝制,欠发) select 单号,款号,颜色,规格,sum(val(裁剪)),sum(val(缝制)),sum(val(裁剪)-val(缝制)) from cjfz group by 单号,款号,颜色,规格"
Data4.Database.Execute "delete * from cjfz where 共量='1'"
Data5.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data5.RecordSource = "select * from cjfz"
Data5.Refresh
End If
End Sub

Private Sub Command10_Click()
Call bzcm(Data8, DBCombo1.Text)
End Sub

Private Sub Command2_Click()
Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data3.RecordSource = "SELECT MAX(val(mid(卡号,7,3))) FROM CPK where  日期=cdate('" & Date & "')"
Data3.Refresh

DBCombo1.Text = Format(Date, "YYMMDD") + "001"
If Data3.Recordset.EOF Then
DBCombo1.Text = Format(Date, "YYMMDD") + "001"
Else
Select Case Len(Data3.Recordset.Fields(0) + 1)
       Case 1
DBCombo1.Text = Format(Date, "YYMMDD") + "00" + Trim(Data3.Recordset.Fields(0) + 1)
       Case 2
DBCombo1.Text = Format(Date, "YYMMDD") + "0" + Trim(Data3.Recordset.Fields(0) + 1)
       Case 3
DBCombo1.Text = Format(Date, "YYMMDD") + Trim(Data3.Recordset.Fields(0) + 1)
End Select
End If

Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from CPK where 卡号='" & DBCombo1.Text & "' order by 编号"
Data1.Refresh

End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text1(3).Text = "" Then
MsgBox ("请输入部位尺码")
Exit Sub
End If
Data2.DatabaseName = "d:\数据库\\htgl\2011\CW.MDB"
Data2.RecordSource = "select * from GDINGXSHU where instr(工序款号,'" & Text1(1).Text & "')>0"
Data2.Refresh
If Data2.Recordset.EOF Then
MsgBox ("没有设置工序系数，设置后再继续！")
Exit Sub
End If
If Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Data6.Database.Execute "insert into cpk(单号,款号,颜色,规格,数量,编号,工序,日期,条码,卡号,品名,备注,单位) VALUES('" & Text1(0).Text & "','" & Text1(1).Text & "','" & Text1(2).Text & "','" & Text1(3).Text & "','" & Text1(4).Text & "','" & Data2.Recordset.Fields(0) & "','" & Data2.Recordset.Fields(1) & "','" & Date & "',trim('" & DBCombo1.Text & "')+trim('" & Data2.Recordset.Fields(0) & "'),'" & DBCombo1.Text & "','" & Text1(5).Text & "','" & Text1(6).Text & "','" & Text1(7).Text & "')"
Data2.Recordset.MoveNext
Loop
Data1.RecordSource = "select * from CPK where 卡号='" & DBCombo1.Text & "' order by 编号"
Data1.Refresh
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If MsgBox("确定作废卡号为：" + DBCombo1.Text + "吗？", vbYesNo) = vbNo Then Exit Sub
Data6.Database.Execute "delete * from cpk where 卡号='" & DBCombo1.Text & "'"
MsgBox ("卡号：" + DBCombo1.Text + "已废除")
End Sub

Private Sub Command6_Click()
Call cpk1(Data8, DBCombo1.Text)
End Sub

Private Sub Command7_Click()
Call cpk(Data8, DBCombo1.Text)
End Sub

Private Sub DTPicker1_Change()
On Error Resume Next
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from cjrb where 日期=cdate('" & Text1(7).Text & "') order by 序号 desc"
Data1.Refresh
Text1(7).Text = DTPicker1.Value

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM cjrb where 日期=cdate('" & Text1(7).Text & "')"
Data2.Refresh

Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub DTPicker1_CloseUp()
On Error Resume Next
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from cjrb where 日期=cdate('" & Text1(7).Text & "') order by 序号 desc"
Data1.Refresh
Text1(7).Text = DTPicker1.Value

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM cjrb where 日期=cdate('" & Text1(7).Text & "')"
Data2.Refresh

Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If
End Sub


Private Sub DTPicker2_Change()
Text1(10).Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text1(10).Text = DTPicker2.Value
End Sub

Private Sub Command8_Click()
Forml8.Text1.Text = Text1(1).Text
Forml8.Show
End Sub

Private Sub Command9_Click()
Data5.RecordSource = "select * from bztm WHERE 卡号='" & DBCombo1.Text & "'"
Data5.Refresh
For i = 0 To Val(Text1(4).Text)
If Len(i) = 1 Then
bz(i) = DBCombo1.Text + "0" + Trim(i)
End If
If Len(i) = 2 Then
bz(i) = DBCombo1.Text + Trim(i)
End If
Next

If Data5.Recordset.EOF Then
Data6.Database.Execute "INSERT INTO bztm(单号,款号,颜色,尺码,数量,卡号,日期,N1,N2,N3,N4,N5,N6,N7,N8,N9,N10,N11,N12,N13,N14,N15,N16,N17,N18,N19,N20,N21,N22,N23,N24,N25,N26,N27,N28,N29,N30,N31,N32,N33,N34,N35,N36,N37,N38,N39,N40,N41,N42,N43,N44,N45,N46,N47,N48,N49,N50) VALUES('" & Text1(0).Text & "','" & Text1(1).Text & "','" & Text1(2).Text & "','" & Text1(3).Text & "','" & Text1(4).Text & "','" & DBCombo1.Text & "',cdate('" & Date & "'),'" & bz(1) & "','" & bz(2) & "','" & bz(3) & "','" & bz(4) & "','" & bz(5) & "','" & bz(6) & "','" & bz(7) & "','" & bz(8) & "','" & bz(9) & "','" & bz(10) & "',  " & _
                                                                        "'" & bz(11) & "','" & bz(12) & "','" & bz(13) & "','" & bz(14) & "','" & bz(15) & "','" & bz(16) & "','" & bz(17) & "','" & bz(18) & "','" & bz(19) & "','" & bz(20) & "','" & bz(21) & "', " & _
                                                                        "'" & bz(22) & "','" & bz(23) & "','" & bz(24) & "','" & bz(25) & "','" & bz(26) & "','" & bz(27) & "','" & bz(28) & "','" & bz(29) & "','" & bz(30) & "','" & bz(31) & "','" & bz(32) & "', " & _
                                                                        "'" & bz(33) & "','" & bz(34) & "','" & bz(35) & "','" & bz(36) & "','" & bz(37) & "','" & bz(38) & "','" & bz(39) & "','" & bz(40) & "','" & bz(41) & "','" & bz(42) & "','" & bz(43) & "', " & _
                                                                        "'" & bz(44) & "','" & bz(45) & "','" & bz(46) & "','" & bz(47) & "','" & bz(48) & "','" & bz(49) & "','" & bz(50) & "')"
Else
Data6.Database.Execute "delete * from bztm where 卡号='" & DBCombo1.Text & "'"
Data6.Database.Execute "INSERT INTO bztm(单号,款号,颜色,尺码,数量,卡号,日期,N1,N2,N3,N4,N5,N6,N7,N8,N9,N10,N11,N12,N13,N14,N15,N16,N17,N18,N19,N20,N21,N22,N23,N24,N25,N26,N27,N28,N29,N30,N31,N32,N33,N34,N35,N36,N37,N38,N39,N40,N41,N42,N43,N44,N45,N46,N47,N48,N49,N50) VALUES('" & Text1(0).Text & "','" & Text1(1).Text & "','" & Text1(2).Text & "','" & Text1(3).Text & "','" & Text1(4).Text & "','" & DBCombo1.Text & "',cdate('" & Date & "'),'" & bz(1) & "','" & bz(2) & "','" & bz(3) & "','" & bz(4) & "','" & bz(5) & "','" & bz(6) & "','" & bz(7) & "','" & bz(8) & "','" & bz(9) & "','" & bz(10) & "',  " & _
                                                                        "'" & bz(11) & "','" & bz(12) & "','" & bz(13) & "','" & bz(14) & "','" & bz(15) & "','" & bz(16) & "','" & bz(17) & "','" & bz(18) & "','" & bz(19) & "','" & bz(20) & "','" & bz(21) & "', " & _
                                                                        "'" & bz(22) & "','" & bz(23) & "','" & bz(24) & "','" & bz(25) & "','" & bz(26) & "','" & bz(27) & "','" & bz(28) & "','" & bz(29) & "','" & bz(30) & "','" & bz(31) & "','" & bz(32) & "', " & _
                                                                        "'" & bz(33) & "','" & bz(34) & "','" & bz(35) & "','" & bz(36) & "','" & bz(37) & "','" & bz(38) & "','" & bz(39) & "','" & bz(40) & "','" & bz(41) & "','" & bz(42) & "','" & bz(43) & "', " & _
                                                                        "'" & bz(44) & "','" & bz(45) & "','" & bz(46) & "','" & bz(47) & "','" & bz(48) & "','" & bz(49) & "','" & bz(50) & "')"
End If
Data5.RecordSource = "select * from bztm WHERE 卡号='" & DBCombo1.Text & "'"
Data5.Refresh
End Sub

Private Sub DBCombo1_Change()
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from CPK where 卡号='" & DBCombo1.Text & "' order by 编号"
Data1.Refresh
Data5.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data5.RecordSource = "select * from bztm WHERE 卡号='" & DBCombo1.Text & "'"
Data5.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
DBCombo2.Text = ""
For i = 0 To 7
Text1(i).Text = ""
Next

For i = 0 To 50
bz(i) = ""
Next

Data2.DatabaseName = "d:\数据库\\htgl\2011\CW.MDB"

Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data5.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data6.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data8.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"

Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data3.RecordSource = "SELECT MAX(val(mid(卡号,7,3))) FROM CPK where  日期=cdate('" & Date & "')"
Data3.Refresh

DBCombo1.Text = Format(Date, "YYMMDD") + "001"
If Data3.Recordset.EOF Then
DBCombo1.Text = Format(Date, "YYMMDD") + "001"
Else
Select Case Len(Data3.Recordset.Fields(0) + 1)
       Case 1
DBCombo1.Text = Format(Date, "YYMMDD") + "00" + Trim(Data3.Recordset.Fields(0) + 1)
       Case 2
DBCombo1.Text = Format(Date, "YYMMDD") + "0" + Trim(Data3.Recordset.Fields(0) + 1)
       Case 3
DBCombo1.Text = Format(Date, "YYMMDD") + Trim(Data3.Recordset.Fields(0) + 1)
End Select
End If

Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from CPK where 卡号='" & DBCombo1.Text & "' order by 编号"
Data1.Refresh

Data7.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data7.RecordSource = "SELECT MAX(val(条码)) FROM CPK where  日期=cdate('" & Date & "')"
Data7.Refresh
gxk = Format(Date, "YYMMDD") + "001"
If Data3.Recordset.EOF Then
gxk = Format(Date, "YYMMDD") + "001"
Else
Select Case Len(Data3.Recordset.Fields(0))
       Case 1
gxk = Format(Date, "YYMMDD") + "00" + Data7.Recordset.Fields(0)
       Case 2
gxk = Format(Date, "YYMMDD") + "0" + Data7.Recordset.Fields(0)
       Case 3
gxk = Format(Date, "YYMMDD") + Data7.Recordset.Fields(0)
End Select
End If




MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid2.ColWidth(0) = 300
For i = 1 To 5
MSFlexGrid2.ColWidth(i) = 1200
Next
MSFlexGrid2.ColWidth(10) = 1200
MSFlexGrid2.ColWidth(11) = 1300
MSFlexGrid2.ColWidth(12) = 1300

End Sub


Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
khbl = 18
Forml202.Text1.Text = DBCombo2.Text
Forml202.Show
End Select
End Sub

Private Sub MSFlexGrid2_dblClick()
On Error Resume Next
If Data1.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To 3
Text1(i).Text = Data1.Recordset.Fields(i)
Next
DBCombo2.Text = Data1.Recordset.Fields(1)
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 3
If Text1(3).Text <> "" Then
Data4.Database.Execute "delete * from cjfz"
Data4.Database.Execute "insert into CJFZ(单号,款号,颜色,规格,裁剪,缝制,欠发) select 单号,款号,颜色,规格,sum(val(裁剪)),'0','0' from cjrb where 款号='" & Text1(1).Text & "' and 颜色='" & Text1(2).Text & "' and 规格='" & Text1(3).Text & "'"
Data4.Database.Execute "insert into CJFZ(单号,款号,颜色,规格,裁剪,缝制,欠发) select 单号,款号,颜色,规格,'0',sum(val(数量)),'0' from cpk where 款号='" & Text1(1).Text & "' and 颜色='" & Text1(2).Text & "' and 规格='" & Text1(3).Text & "'"
Data4.Database.Execute "update cjfz set 共量='1'"
Data4.Database.Execute "insert into CJFZ(单号,款号,颜色,规格,裁剪,缝制,欠发) select 单号,款号,颜色,规格,sum(val(裁剪)),sum(val(缝制)),sum(val(裁剪)-val(缝制)) from cjfz group by 单号,款号,颜色,规格"
Data4.Database.Execute "delete * from cjfz where 共量='1'"
Data5.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data5.RecordSource = "select * from cjfz"
Data5.Refresh
End If
End Select
End Sub

Private Sub MSF()
With MSFlexGrid2
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

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSF
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid2.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data1.Recordset.MoveFirst
Data1.Recordset.Move r - 1
Data1.Recordset.Edit

Data1.Recordset.Fields(c - 1) = Text1111.Text
Data1.Recordset.Update

MSFlexGrid2.Text = Text1111.Text
Text1111.Visible = False
MSFlexGrid2.SetFocus
End If
End Sub

