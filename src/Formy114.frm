VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy114 
   BackColor       =   &H00C0E0FF&
   Caption         =   "款式管理"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data5 
      Caption         =   "Data1"
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "尺码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      TabIndex        =   38
      Text            =   "Text3"
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单耗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4080
      Width           =   855
   End
   Begin VB.Data Data4 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9840
      TabIndex        =   35
      Text            =   "Text2"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   9840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Text            =   "Formy114.frx":0000
      Top             =   1800
      Width           =   4575
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3135
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   3015
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy114.frx":0006
      Height          =   4695
      Left            =   840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4800
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   8281
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
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   1800
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy114.frx":001A
      Height          =   330
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy114.frx":002E
      Height          =   330
      Index           =   3
      Left            =   1680
      TabIndex        =   9
      Top             =   2760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy114.frx":0042
      Height          =   330
      Index           =   4
      Left            =   5520
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "名称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy114.frx":0056
      Height          =   330
      Index           =   5
      Left            =   5520
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "名称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   6
      Left            =   5520
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "名称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   7
      Left            =   5520
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "名称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy114.frx":006A
      Height          =   330
      Index           =   8
      Left            =   7680
      TabIndex        =   14
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "名称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   9
      Left            =   7680
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "名称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   0
      Left            =   1680
      TabIndex        =   20
      Top             =   1320
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   10
      Left            =   7680
      TabIndex        =   29
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "名称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   11
      Left            =   7680
      TabIndex        =   30
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "名称"
      Text            =   "DBCombo1"
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号快捷"
      Height          =   375
      Index           =   11
      Left            =   840
      TabIndex        =   37
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "尺码8"
      Height          =   375
      Index           =   10
      Left            =   7080
      TabIndex        =   33
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "尺码7"
      Height          =   375
      Index           =   9
      Left            =   7080
      TabIndex        =   32
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "尺码6"
      Height          =   375
      Index           =   8
      Left            =   7080
      TabIndex        =   31
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   1335
      Index           =   7
      Left            =   9240
      TabIndex        =   28
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   6
      Left            =   9240
      TabIndex        =   27
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "尺码5"
      Height          =   375
      Index           =   4
      Left            =   7080
      TabIndex        =   26
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "尺码4"
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   25
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "尺码3"
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   24
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "尺码2"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   23
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "尺码1"
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   22
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单位"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   21
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  成衣制作款式信息"
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
      Left            =   4680
      TabIndex        =   19
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   375
      Index           =   13
      Left            =   840
      TabIndex        =   18
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "品名"
      Height          =   375
      Index           =   14
      Left            =   840
      TabIndex        =   17
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   19
      Left            =   840
      TabIndex        =   16
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "Formy114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X As Integer
Dim ba As Database: Dim rr As Integer
Dim rs As Single: Dim rd1 As Recordset: Dim ba1 As Database: Public ll As Integer: Public RQ As Date
Dim rd As Recordset: Public mm As Date: Public ml As Date

Private Sub Command12_Click()
Unload Me
Formy4.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
If DBCombo1(0).Text = "" Then Exit Sub

Data1.Recordset.AddNew
For i = 0 To 11
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Fields(12) = Text1.Text
Data1.Recordset.Fields(13) = Text2.Text

Data1.Recordset.Update

Data1.Refresh
Data2.Refresh
Data3.Refresh
For i = 4 To 11
DBCombo1(i).Text = ""
Next
Text2.Text = 1
Text2.Text = Data1.Recordset.Fields(13) + 1

End Sub

Private Sub Command2_Click()
On Error Resume Next

If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If

If DBCombo1(0).Text = "" Then Exit Sub

Data1.Recordset.Edit
For i = 0 To 11
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Fields(12) = Text1.Text
Data1.Recordset.Fields(13) = Text2.Text
Data1.Recordset.Update
Data1.Refresh
For i = 4 To 11
DBCombo1(i).Text = ""
Next

Text2.Text = 1
Text2.Text = Data2.Recordset.Fields(0) + 1

End Sub

Private Sub Command4_Click()

On Error Resume Next

If MsgBox("确定删除吗?,删除将不能恢复!", vbYesNo) = vbNo Then Exit Sub


Data1.Recordset.Delete
Data1.Refresh
For i = 4 To 11
DBCombo1(i).Text = ""
Next
Text2.Text = 1
Text2.Text = Data2.Recordset.Fields(0) + 1

End Sub


Private Sub Command5_Click()
If Text3.Text = "" Then
MsgBox ("请输入款号")
Exit Sub
End If

If MsgBox("确定生成尺码吗？", vbYesNo) = vbNo Then Exit Sub
Data4.RecordSource = "SELECT *  from ksnr where instr(款号,'" & Text3.Text & "')>0 order by  款号"
Data4.Refresh
If Data4.Recordset.EOF Then
MsgBox ("没有内容，不能生成尺码")
Exit Sub
Else
Data4.Recordset.MoveFirst
Data5.Database.Execute "delete * from cmxx WHERE 款号='" & Data4.Recordset.Fields(0) & "'"
Do While Not Data4.Recordset.EOF
For i = 4 To 11
If Data4.Recordset.Fields(i) <> "" Then
Data5.Database.Execute "insert into cmxx(款号,品名,颜色,单位,尺码) VALUES('" & Data4.Recordset.Fields(0) & "','" & Data4.Recordset.Fields(1) & "','" & Data4.Recordset.Fields(2) & "','" & Data4.Recordset.Fields(3) & "','" & Data4.Recordset.Fields(i) & "')"
End If
Next
Data4.Recordset.MoveNext
Loop
MsgBox ("尺码生成成功！")
End If
End Sub

Private Sub Command6_Click()
Data1.RecordSource = "select * from KSNR WHERE 款号='" & DBCombo1(0).Text & "' ORDER BY 序号"
Data1.Refresh
Call ksdy(MSFlexGrid1, DBCombo1(0).Text)
Data1.RecordSource = "select * from KSNR WHERE 款号='" & DBCombo1(0).Text & "' ORDER BY 序号 desc"
Data1.Refresh
End Sub

Private Sub Command7_Click()
Formy18.DBCombo1(1).Text = DBCombo1(0).Text
Formy18.DBCombo1(2).Text = DBCombo1(2).Text
Formy18.Show
End Sub

Private Sub Command8_Click()
On Error Resume Next
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
Data1.Refresh
Text2.Text = 1
Text2.Text = Data1.Recordset.Fields(13) + 1
End Sub



Private Sub DBCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Form_Load()
On Error Resume Next

For i = 0 To 14
DBCombo1(i).Text = ""
Next

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

Text2.Text = 1
Text2.Text = Data2.Recordset.Fields(0) + 1

Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from KSNR WHERE instr(款号,'" & Text3.Text & "')>0 ORDER BY 序号 DESC "
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data2.RecordSource = "SELECT YS.YS  from YS GROUP BY YS.YS"
Data2.Refresh

Data3.DatabaseName = "d:\数据库\\htgl\2011\cpck.MDB"
Data3.RecordSource = "SELECT mc  from cldw GROUP BY mc"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"

Data5.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"

MSFlexGrid1.ColWidth(0) = 100
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1200
DBCombo1(0).TabIndex = 0
End Sub


Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row

If Data1.Recordset.EOF Then
Exit Sub
End If

Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To Data1.Recordset.Fields.Count - 1
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next
Text1.Text = Data1.Recordset.Fields(14)
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text3_Change()
On Error Resume Next
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from KSNR WHERE instr(款号,'" & Text3.Text & "')>0 ORDER BY 序号 DESC "
Data1.Refresh
Text2.Text = 1
Text2.Text = Data1.Recordset.Fields(13) + 1
End Sub
