VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formm111 
   BackColor       =   &H00C0E0FF&
   Caption         =   "备注设置"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10815
   LinkTopic       =   "Form19"
   ScaleHeight     =   10275
   ScaleWidth      =   10815
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选取"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Index           =   3
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Formm111.frx":0000
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formm111.frx":0006
      Height          =   7455
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   8
      RowHeightMin    =   400
      BackColorFixed  =   8421631
      BackColorBkg    =   40863
      WordWrap        =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "代码"
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注信息"
      Height          =   495
      Index           =   2
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "助记码"
      Height          =   495
      Index           =   1
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "编号"
      Height          =   495
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Formm111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1(0).Text = "" Or Text1(1).Text = "" Then
MsgBox ("代码、编号不能为空白")
Exit Sub
End If
Data1.Recordset.AddNew
For i = 0 To 3
Data1.Recordset.Fields(i) = Text1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh
Data2.Refresh
For i = 1 To 3
Text1(i).Text = ""
Next
Text1(1).Text = 1
If Data2.Recordset.EOF Then
Text1(1).Text = 1
Else
Text1(1).Text = Data2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text1(0).Text = "" Or Text1(1).Text = "" Then
MsgBox ("代码、编号不能为空白")
Exit Sub
End If
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Edit
For i = 0 To 3
Data1.Recordset.Fields(i) = Text1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh
Data2.Refresh
For i = 1 To 3
Text1(i).Text = ""
Next
Text1(1).Text = 1
If Data2.Recordset.EOF Then
Text1(1).Text = 1
Else
Text1(1).Text = Data2.Recordset.Fields(0) + 1
End If
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
Data2.Refresh
For i = 1 To 3
Text1(i).Text = ""
Next
Text1(1).Text = 1
If Data2.Recordset.EOF Then
Text1(1).Text = 1
Else
Text1(1).Text = Data2.Recordset.Fields(0) + 1
End If
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
On Error Resume Next
Data2.Refresh
For i = 1 To 3
Text1(i).Text = ""
Next
Text1(1).Text = 1
If Data2.Recordset.EOF Then
Text1(1).Text = 1
Else
Text1(1).Text = Data2.Recordset.Fields(0) + 1
End If
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Form_Load()
For i = 0 To 3
Text1(i).Text = ""
Next
Data1.DatabaseName = "d:\数据库\bfrz\" + LJB + "\sczyjhd.MDB"
Data2.DatabaseName = "d:\数据库\bfrz\" + LJB + "\sczyjhd.MDB"

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 800
MSFlexGrid1.ColWidth(2) = 600
MSFlexGrid1.ColWidth(3) = 800
MSFlexGrid1.ColWidth(4) = 7300
MSFlexGrid1.RowHeight(0) = 200
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub MSFlexGrid1_dblClick()
If Data1.Recordset.EOF Then Exit Sub
rs = MSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 1 To 3
Text1(i).Text = Data1.Recordset.Fields(i)
Next
Command1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data1.RecordSource = "select * from bz where 代码='" & Text1(0).Text & "' order by val(编号) desc"
Data1.Refresh
Data2.RecordSource = "select max(val(编号)) from bz where 代码='" & Text1(0).Text & "'"
Data2.Refresh
Text1(1).Text = 1
If Data2.Recordset.EOF Then
Text1(1).Text = 1
Else
Text1(1).Text = Data2.Recordset.Fields(0) + 1
End If
End Select
End Sub
