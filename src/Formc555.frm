VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formc555 
   BackColor       =   &H00C0E0FF&
   Caption         =   "供应商信息"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form55"
   MDIChild        =   -1  'True
   ScaleHeight     =   10800
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "成批"
      Height          =   375
      Left            =   13680
      TabIndex        =   28
      Top             =   1680
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "单个"
      Height          =   375
      Left            =   13680
      TabIndex        =   27
      Top             =   2280
      Width           =   975
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入"
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   4080
      TabIndex        =   14
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   6600
      TabIndex        =   13
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   9000
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   6600
      TabIndex        =   11
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   9000
      TabIndex        =   10
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   10800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   975
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   2655
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
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   600
      TabIndex        =   1
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   4080
      TabIndex        =   0
      Top             =   3600
      Width           =   2415
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc555.frx":0000
      Height          =   360
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "客户名称"
      BoundColumn     =   "客户名称"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc555.frx":0014
      Height          =   5055
      Left            =   600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4200
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   8916
      _Version        =   393216
      BackColorFixed  =   12632319
      BackColorSel    =   16777088
      ForeColorSel    =   -2147483635
      BackColorBkg    =   49344
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2520
      TabIndex        =   15
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户全称"
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
      Left            =   600
      TabIndex        =   25
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户地址"
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
      Left            =   4080
      TabIndex        =   24
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "联系人"
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
      Index           =   2
      Left            =   6600
      TabIndex        =   23
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "联系电话"
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
      Index           =   3
      Left            =   9000
      TabIndex        =   22
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "联系手机"
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
      Index           =   4
      Left            =   6600
      TabIndex        =   21
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "传真"
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
      Left            =   9000
      TabIndex        =   20
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "地区号"
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
      Index           =   6
      Left            =   10800
      TabIndex        =   19
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "供 应 商 资 料 信 息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6120
      TabIndex        =   18
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户代码"
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
      Index           =   7
      Left            =   600
      TabIndex        =   17
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户简称"
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
      Index           =   8
      Left            =   4080
      TabIndex        =   16
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "Formc555"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ba As Database: Dim rr As Integer
Dim rs As Single: Dim rd1 As Recordset: Dim ba1 As Database: Public ll As Integer
Dim rd As Recordset: Public mm As Date: Public ml As Date: Dim MODIFY As Integer
Private Sub JILU2()
Dim i As Single
Data2.Refresh
If Data2.Recordset.EOF Then
MSFlexGrid2.TextMatrix(0, 0) = "记录号"

Exit Sub
End If
Data2.Recordset.MoveLast
MSFlexGrid2.TextMatrix(0, 0) = "记录号"
For i = 1 To Data2.Recordset.RecordCount
MSFlexGrid2.TextMatrix(i, 0) = i
Next
End Sub
Private Sub Command12_Click()
Unload Me
Form4.Show
End Sub




Private Sub Command1_Click()
On Error Resume Next

rd.AddNew
For i = 0 To rd.Fields.Count - 1
rd.Fields(i) = Text1(i).Text
Next
rd.Update
Data1.Refresh

For i = 0 To rd.Fields.Count - 1
Text1(i).Text = ""
Next
DBCombo1.Text = ""
DBCombo1.SetFocus

Data2.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data2.RecordSource = "select GYS.客户名称 from GYS group by GYS.客户名称"
Data2.Refresh


End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If

Data1.Recordset.Edit
   For i = 0 To rd.Fields.Count - 1
   Data1.Recordset.Fields(i) = Text1(i).Text
   Next
Data1.Recordset.Update
Data1.Refresh

If Data1.Recordset.EOF Then
MSFlexGrid1.TextMatrix(0, 0) = "记录号"
Exit Sub
End If
Data1.Recordset.MoveLast
MSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To Data1.Recordset.RecordCount
MSFlexGrid1.TextMatrix(i, 0) = i
Next
For i = 0 To rd.Fields.Count - 1
Text1(i).Text = ""
Next
DBCombo1.Text = ""
DBCombo1.SetFocus
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定删除吗?,删除将不能恢复!", vbYesNo) = vbNo Then
Exit Sub
End If

Data1.Recordset.Delete
Data1.Refresh

For i = 0 To rd.Fields.Count - 1
Text1(i).Text = ""
Next
DBCombo1.Text = ""
DBCombo1.SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub


Private Sub Command6_Click()
If Option1.Value = True Then
If MsgBox("确定导入吗？", vbYesNo) = vbNo Then Exit Sub
Data3.Database.Execute "DELETE * FROM CWMC WHERE  INSTR(科目编号,'2121')>0 AND LEN(科目编号)>4"
Data3.RecordSource = "CWMC"
Data3.Refresh
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
i = 1
l = "2121"
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
Data3.Recordset.AddNew
If Len(i) = 1 Then
Data3.Recordset.Fields(0) = l + "00" + Trim(i)
End If
If Len(i) = 2 Then
Data3.Recordset.Fields(0) = l + "0" + Trim(i)
End If
If Len(i) = 3 Then
Data3.Recordset.Fields(0) = l + Trim(i)
End If
Data3.Recordset.Fields(1) = Data1.Recordset.Fields("简称")
Data3.Recordset.Fields(2) = "负债"
Data3.Recordset.Fields(3) = "贷"
Data3.Recordset.Fields(4) = "2"
Data3.Recordset.Fields(5) = "是"
Data3.Recordset.Update
i = i + 1
Data1.Recordset.MoveNext
Loop
MsgBox ("导入成功!")
End If

If Option2.Value = True Then
If MsgBox(Text1(8).Text + "   确定导入吗？", vbYesNo) = vbNo Then Exit Sub
Data3.RecordSource = "SELECT MAX(MID(科目编号,5)) FROM CWMC WHERE  INSTR(科目编号,'2121')>0 AND LEN(科目编号)>5"
Data3.Refresh
MC = 0
If Not Data3.Recordset.EOF Then
MC = Data3.Recordset.Fields(0)
End If
If MC = Null Then MC = 0
MC = Val(MC)
Data4.Recordset.AddNew
If Len(MC) = 1 Then
Data4.Recordset.Fields(0) = "2121" + "00" + Trim(MC + 1)
End If
If Len(MC) = 2 Then
Data4.Recordset.Fields(0) = "2121" + "0" + Trim(MC + 1)
End If
If Len(MC) = 3 Then
Data4.Recordset.Fields(0) = "2121" + Trim(MC + 1)
End If
Data4.Recordset.Fields(1) = Data1.Recordset.Fields("简称")
Data4.Recordset.Fields(2) = "负债"
Data4.Recordset.Fields(3) = "贷"
Data4.Recordset.Fields(4) = "2"
Data4.Recordset.Fields(5) = "是"
Data4.Recordset.Update
MsgBox ("导入成功!")
End If

End Sub


Private Sub Command8_Click()
Unload Me
Form2.Show
End Sub



Private Sub Command5_Click()
Data1.Refresh
Data2.Refresh
End Sub

Private Sub DBCombo1_Change()
Text1(0).Text = DBCombo1.Text
End Sub


Private Sub DBCombo1_Click(Area As Integer)
Text1(0).Text = DBCombo1.Text
End Sub

Private Sub DBCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Form_Load()
On Error Resume Next
Set ba = OpenDatabase("d:\数据库\\htgl\2011\SCZYJHD.MDB")
Set rd = ba.OpenRecordset("GYS", dbOpenDynaset)
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from GYS order by GYS.简称"
Data1.Refresh
If Data1.Recordset.EOF Then
MSFlexGrid1.TextMatrix(0, 0) = "记录号"
Else
Data1.Recordset.MoveLast
MSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To Data1.Recordset.RecordCount
MSFlexGrid1.TextMatrix(i, 0) = i
Next
End If

Data4.DatabaseName = "d:\数据库\\htgl\2011\CW.MDB"
Data4.RecordSource = "SELECT * FROM CWMC"
Data4.Refresh
Data3.DatabaseName = "d:\数据库\\htgl\2011\CW.MDB"
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "select GYS.客户名称 from GYS group by GYS.客户名称"
Data2.Refresh
DBCombo1.TabIndex = 0
MSFlexGrid1.ColWidth(1) = 2600

MSFlexGrid1.ColWidth(2) = 2600
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.ColWidth(6) = 1200
Text2.Enabled = False

End Sub
Private Sub JILU()
Dim i As Single
Data1.Refresh
If rd.EOF Then
MSFlexGrid1.TextMatrix(0, 0) = "记录号"

Exit Sub
End If
rd.MoveLast
MSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To rd.RecordCount
MSFlexGrid1.TextMatrix(i, 0) = i
Next
End Sub


Private Sub Label1_dblClick(Index As Integer)
Select Case Index
       Case 6
       Text2.Enabled = True
End Select
End Sub

Private Sub MSFlexGrid1_Click()
On Error Resume Next
rs = MSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To Data1.Recordset.Fields.Count - 1
Text1(i).Text = Data1.Recordset.Fields(i)
Next
DBCombo1.Text = Text1(0).Text
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text1_lostFocus(Index As Integer)
On Error Resume Next
         
         
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

