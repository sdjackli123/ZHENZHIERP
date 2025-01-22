VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Forms203 
   BackColor       =   &H00C0E0FF&
   Caption         =   "条码扫描"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form20"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data11 
      Caption         =   "Data4"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data12 
      Caption         =   "Data5"
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
      Width           =   2175
   End
   Begin VB.Data Data6 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   2535
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   240
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1320
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Forms203.frx":0000
      Height          =   2895
      Left            =   4440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   8
      BackColorFixed  =   8421631
      BackColorBkg    =   4109501
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Forms203.frx":0014
      Height          =   4575
      Left            =   4440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   16777215
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Forms203.frx":0029
      Height          =   4575
      Left            =   10680
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   16777215
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
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "请扫描工号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "扫描区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "Forms203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SZ As Integer: Public BH As String
Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\cw.mdb"
Data1.RecordSource = "works1"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\db.mdb"
Data2.RecordSource = "select * from clb where 标签='" & m & "'"
Data2.Refresh

Data3.DatabaseName = "d:\数据库\\htgl\2011\cw.mdb"
Data3.RecordSource = "GDINGXSHU"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\SCjd.mdb"
Data4.RecordSource = "CPK"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\SCjd.mdb"
Data5.Refresh

Data6.DatabaseName = "d:\数据库\\htgl\2011\DB.mdb"
Data6.Refresh

Data11.DatabaseName = "d:\数据库\\htgl\2011\db.MDB"
Data12.DatabaseName = "d:\数据库\\htgl\2011\db.MDB"

Timer1.Enabled = False
MSFlexGrid3.ColWidth(0) = 300
MSFlexGrid3.ColWidth(1) = 1300
End Sub

Private Sub Text1_Change()
On Error Resume Next
Data1.DatabaseName = "d:\数据库\\htgl\2011\cw.mdb"
Data1.RecordSource = "works1"
Data1.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\SCjd.mdb"
Data4.RecordSource = "CPK"
Data4.Refresh

If InStr(Text1.Text, "J") > 0 Then
m = Left(Text1.Text, Len(Text1.Text) - 1)

If Len(m) = 3 Then
Data1.Recordset.FindFirst "编号='" & m & "'"
If Data1.Recordset.NoMatch Then
Text2.Text = ""
Else
Text2.Text = Data1.Recordset.Fields(2)
BH = m
End If
Text1.Text = ""
Text1.SetFocus
Exit Sub
End If

If Len(m) = 11 Then
If Text2.Text = "" Then
SZ = 1
Label2.Caption = "请扫描工号"
Timer1.Enabled = True
Text1.Text = ""
Text1.SetFocus
Else
Data4.Recordset.FindFirst "条码='" & m & "'"
If Data4.Recordset.NoMatch Then
SZ = 1
Label2.Caption = "不存在此条码"
Text1.Text = ""
Timer1.Enabled = True
Exit Sub
Else
Data6.RecordSource = "SELECT SUM(产量) FROM CLB WHERE 标签='" & m & "'"
Data6.Refresh

MMM = 0
If Not Data6.Recordset.EOF Then
MMM = Data6.Recordset.Fields(0)
End If

If MMM >= Val(Data4.Recordset.Fields(4)) Then
Data2.RecordSource = "select 日期,款式,颜色,尺码,工序编号,工序,操作员,姓名,产量 from clb where 标签='" & m & "'"
Data2.Refresh
Text1.Text = ""
Text1.SetFocus
Call sx
Exit Sub
End If

Data5.Database.Execute "INSERT INTO CLB(日期,工序编号,工序,操作员,产量,姓名,标签,款式,颜色,尺码,单号) in'd:\数据库\\htgl\2011\db.mdb' select CDATE('" & Date & "'),编号,工序,'" & BH & "',数量,'" & Text2.Text & "',条码,款号,颜色,规格,单号 from cpk where 条码='" & m & "'"
Data2.RecordSource = "select 日期,款式,颜色,尺码,工序编号,工序,操作员,姓名,产量 from clb where 标签='" & m & "'"
Data2.Refresh
Text1.Text = ""
Text1.SetFocus
End If
Call sx
End If
End If

If Len(m) <> 3 And Len(m) <> 12 Then
Text1.Text = ""
Text1.SetFocus
End If

End If
''''''''''''''''''''''
End Sub

Private Sub Timer1_Timer()
Label2.Visible = True
If SZ = 5 Then
Label2.Visible = False
Timer1.Enabled = False
End If
SZ = SZ + 1
End Sub

Private Sub sx()
'On Error Resume Next
If Text2.Text <> "" Then
Data11.RecordSource = "SELECT 日期,count(工序编号) as 张数,sum(产量) as 统计量 FROM CLB WHERE  日期=CDATE('" & Date & "') AND CLB.姓名='" & Text2.Text & "' group by 日期 ORDER BY 日期"
Data11.Refresh
Data12.RecordSource = "SELECT count(工序编号) as 张数,sum(产量) as 统计量 FROM CLB WHERE  日期=CDATE('" & Date & "') AND CLB.姓名='" & Text2.Text & "'"
Data12.Refresh
End If
End Sub

