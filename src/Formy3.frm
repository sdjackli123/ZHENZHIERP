VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy3 
   BackColor       =   &H00C0E0FF&
   Caption         =   "纱线单位设置"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Palette         =   "Formy3.frx":0000
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      Left            =   1800
      TabIndex        =   15
      Top             =   2280
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
      Index           =   1
      Left            =   3840
      TabIndex        =   14
      Top             =   2280
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
      Index           =   2
      Left            =   6240
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
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
      Left            =   7680
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
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
      Left            =   6240
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
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
      Left            =   7680
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   9120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录      入"
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
      Left            =   11400
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   2055
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
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
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
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
      Left            =   1800
      TabIndex        =   1
      Top             =   3480
      Width           =   1935
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
      Left            =   3840
      TabIndex        =   0
      Top             =   3480
      Width           =   2055
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy3.frx":6844
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
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
      Bindings        =   "Formy3.frx":6858
      Height          =   4695
      Left            =   1800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8281
      _Version        =   393216
      BackColorFixed  =   12632319
      BackColorSel    =   16777088
      ForeColorSel    =   -2147483635
      BackColorBkg    =   49344
      FocusRect       =   0
      AllowUserResizing=   3
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
      Left            =   1800
      TabIndex        =   25
      Top             =   1560
      Width           =   1935
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
      Left            =   3840
      TabIndex        =   24
      Top             =   1560
      Width           =   2295
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
      Left            =   6240
      TabIndex        =   23
      Top             =   1560
      Width           =   1335
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
      Left            =   7680
      TabIndex        =   22
      Top             =   1560
      Width           =   1335
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
      Left            =   6240
      TabIndex        =   21
      Top             =   2760
      Width           =   1335
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
      Left            =   7680
      TabIndex        =   20
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
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
      Left            =   9120
      TabIndex        =   19
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "客 户 资 料 信 息"
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
      Left            =   6360
      TabIndex        =   18
      Top             =   360
      Width           =   3495
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
      Left            =   1800
      TabIndex        =   17
      Top             =   2760
      Width           =   1935
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
      Left            =   3840
      TabIndex        =   16
      Top             =   2760
      Width           =   2055
   End
End
Attribute VB_Name = "Formy3"
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
Formy4.Show
End Sub




Private Sub Command1_Click()
           
         
         rd.FindFirst "简称='" & Text1(8).Text & "'"
         If rd.NoMatch Then
            lll = 0
         Else
         MsgBox ("客户简称重复，请重新输入！")
         Text1(8).SetFocus
         Exit Sub
         End If

rd.AddNew
For i = 0 To rd.Fields.Count - 1
rd.Fields(i) = Text1(i).Text
Next
rd.Update
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
Text1(6).Text = Data1.Recordset.RecordCount + 1
DBCombo1.Text = ""
DBCombo1.SetFocus

Data2.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.mdb"
Data2.RecordSource = "select 客户名称 from KHZL group by 客户名称"
Data2.Refresh


End Sub

Private Sub Command2_Click()
         
         
         
If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If

   Data1.Recordset.Edit
   For i = 0 To Data1.Recordset.Fields.Count - 1
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
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
DBCombo1.Text = ""
Text1(6).Text = Data1.Recordset.RecordCount + 1
DBCombo1.SetFocus
End Sub

Private Sub Command3_Click()
 
On Error Resume Next

Data1.Recordset.Delete
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
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

For i = 0 To rd.Fields.Count - 1
Text1(i).Text = ""
Next
DBCombo1.Text = ""
Text1(6).Text = Data1.Recordset.RecordCount + 1
DBCombo1.SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command8_Click()
Unload Me
Formy2.Show
End Sub



Private Sub Command5_Click()
Data1.Refresh
Data2.Refresh
Text1(6).Text = Data1.Recordset.RecordCount + 1
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub DBCombo1_Change()
Text1(0).Text = DBCombo1.Text
End Sub


Private Sub DBCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Form_Load()
On Error Resume Next
Set ba = OpenDatabase("d:\数据库\\htgl\2011\SCZYJHD.MDB")
Set rd = ba.OpenRecordset("sxZL", dbOpenDynaset)
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select * from SXZL order by val(ip)"
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


Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "select 客户名称 from SXZL group by 客户名称"
Data2.Refresh
Text1(6).Text = Data1.Recordset.RecordCount + 1
DBCombo1.TabIndex = 0
MSFlexGrid1.ColWidth(1) = 2600

MSFlexGrid1.ColWidth(2) = 2600
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.ColWidth(6) = 1200

Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

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

Private Sub MSFlexGrid1_dblClick()
rs = MSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To Data1.Recordset.Fields.Count - 1
Text1(i).Text = Data1.Recordset.Fields(i)
Next
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text1_lostFocus(Index As Integer)
On Error Resume Next
         
         
End Sub

Private Sub Text2_Change()
On Error Resume Next
Text1(6).Text = Text2.Text
rd.FindFirst "ip='" & Text2.Text & "'"
   If rd.NoMatch Then
    Command2.Enabled = False
    Command3.Enabled = False
    Command1.Enabled = True
    For i = 0 To rd.Fields.Count - 1
     Text1(i).Text = ""
     Next
     Text2.Text = Data1.Recordset.RecordCount + 1
     DBCombo1.Text = ""
     Else
     Command3.Enabled = True
     Command1.Enabled = False
     Command2.Enabled = True
     DBCombo1.Text = rd.Fields(0)
     For i = 0 To rd.Fields.Count - 1
     Text1(i).Text = rd.Fields(i)
     Next
  End If
  
  

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

