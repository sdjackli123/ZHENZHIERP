VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy502 
   BackColor       =   &H00C0E0FF&
   Caption         =   "织布入库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form42"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   14400
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command9 
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
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
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   12960
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   8880
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   10200
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   7800
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   6840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   10440
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
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4455
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按款号"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
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
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
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
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   1440
      TabIndex        =   16
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy502.frx":0000
      Height          =   2295
      Left            =   4200
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy502.frx":0014
      Height          =   330
      Left            =   1440
      TabIndex        =   18
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "款号"
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy502.frx":0028
      Height          =   5055
      Left            =   360
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4560
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   13
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12960
      TabIndex        =   20
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81068033
      CurrentDate     =   36892
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy502.frx":003C
      Height          =   330
      Left            =   10200
      TabIndex        =   21
      Top             =   3480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DBCombo4"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
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
      Left            =   14400
      TabIndex        =   34
      Top             =   3120
      Width           =   495
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
      Index           =   14
      Left            =   12960
      TabIndex        =   32
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单价"
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
      Index           =   11
      Left            =   8880
      TabIndex        =   31
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "加工"
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
      Left            =   10200
      TabIndex        =   30
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料"
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
      Left            =   4920
      TabIndex        =   29
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "幅宽"
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
      Left            =   6840
      TabIndex        =   28
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "入库"
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
      Left            =   7800
      TabIndex        =   27
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择单号"
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
      Left            =   360
      TabIndex        =   26
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择款号"
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
      Left            =   2160
      TabIndex        =   25
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择颜色"
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
      Left            =   3720
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择款号"
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
      TabIndex        =   23
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择单号"
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
      Left            =   360
      TabIndex        =   22
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Formy502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data2.RecordSource = "SELECT 单号,款号,颜色,材料名称,毛坯幅宽,织布量 FROM zbjh WHERE instr(款号,'" & DBCombo2.Text & "')>0"
Data2.Refresh
Data4.RecordSource = "select * from zbrk where  instr(款号,'" & DBCombo2.Text & "')>0 order by 日期"
Data4.Refresh
Text1(9).Text = 1
If Not Data4.Recordset.EOF Then
Text1(9).Text = Data4.Recordset.RecordCount + 1
End If
End Sub

Private Sub Command2_Click()
Data2.RecordSource = "SELECT 单号,款号,颜色,材料名称,毛坯幅宽,织布量 FROM zbjh WHERE 单号='" & DBCombo1.Text & "'"
Data2.Refresh
Data1.RecordSource = "SELECT 款号 FROM zbfl WHERE 单号='" & DBCombo1.Text & "' group BY 款号"
Data1.Refresh
Data4.RecordSource = "select * from zbrk where  单号='" & DBCombo1.Text & "' order by 日期"
Data4.Refresh
Text1(9).Text = 1
If Not Data4.Recordset.EOF Then
Text1(9).Text = Data4.Recordset.RecordCount + 1
End If
End Sub
Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
On Error Resume Next
If MsgBox("删除不能回复！确定删除吗？", vbYesNo) = vbNo Then Exit Sub
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.Delete
Data4.Refresh
Text1(9).Text = 1
If Not Data4.Recordset.EOF Then
Text1(9).Text = Data4.Recordset.RecordCount + 1
End If
Text1(4).Text = ""
Text1(5).Text = ""
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command7_Click()
On Error Resume Next
If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(7).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data4.Recordset.AddNew
For i = 0 To 9
Data4.Recordset.Fields(i) = Text1(i).Text
Next
Data4.Recordset.Update
Data4.Refresh
Text1(9).Text = 1
If Not Data4.Recordset.EOF Then
Text1(9).Text = Data4.Recordset.RecordCount + 1
End If

Text1(4).SetFocus
End Sub

Private Sub Command8_Click()
On Error Resume Next
If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(7).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data4.Recordset.Edit
For i = 0 To 9
Data4.Recordset.Fields(i) = Text1(i).Text
Next
Data4.Recordset.Update
Data4.Refresh
Text1(9).Text = 1
If Not Data4.Recordset.EOF Then
Text1(9).Text = Data4.Recordset.RecordCount + 1
End If

Text1(4).Text = ""
Text1(5).Text = ""
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command9_Click()
Data4.RecordSource = "select * from zbrk where 单号='" & DBCombo1.Text & "'  order by 日期"
Data4.Refresh
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(9).Text = 1
If Not Data4.Recordset.EOF Then
Text1(9).Text = Data4.Recordset.RecordCount + 1
End If
Text1(4).SetFocus
End Sub


Private Sub DTPicker1_Change()
Text1(11).Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text1(11).Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Change()
Text1(12).Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text1(12).Text = DTPicker2.Value
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo4.Text = ""
For i = 0 To 8
Text1(i).Text = ""
Next
Text1(8).Text = Date
DTPicker1.Value = Date

Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data5.RecordSource = "select 简称 from ZBZL group by 简称"
Data5.Refresh
Data6.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data6.RecordSource = "select 简称 from RSZL group by 简称"
Data6.Refresh
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(4) = 2000
MSFlexGrid1.ColWidth(12) = 1300

MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(4) = 1800
End Sub
Private Sub DBCombo4_Click(Area As Integer)
Text1(6).Text = DBCombo4.Text
End Sub
Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
khbl = 6
Formy202.Show
End Select
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data4.Recordset.Move rs - 1
For i = 0 To 8
Text1(i).Text = Data4.Recordset.Fields(i)
Next
DTPicker1.Value = Text1(8).Text
Command7.Enabled = False
Command8.Enabled = True
Command6.Enabled = True
End Sub

Private Sub MSFlexGrid2_dblClick()
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
For i = 0 To 5
Text1(i).Text = Data2.Recordset.Fields(i)
Next
End Sub


