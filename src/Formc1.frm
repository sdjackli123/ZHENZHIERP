VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formc1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "库存记录"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   600
      Width           =   1215
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   600
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
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   600
      Width           =   1095
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   600
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc1.frx":0000
      Height          =   6495
      Left            =   480
      TabIndex        =   22
      Top             =   3120
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   11456
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc1.frx":0014
      Height          =   390
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   2
      Left            =   1920
      TabIndex        =   13
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc1.frx":0028
      Height          =   390
      Index           =   3
      Left            =   1920
      TabIndex        =   14
      Top             =   2400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   4
      Left            =   7080
      TabIndex        =   15
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   5
      Left            =   7080
      TabIndex        =   16
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   6
      Left            =   7080
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   7
      Left            =   1920
      TabIndex        =   18
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   8
      Left            =   12120
      TabIndex        =   19
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   9
      Left            =   7080
      TabIndex        =   20
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc1.frx":003C
      Height          =   390
      Index           =   10
      Left            =   12120
      TabIndex        =   21
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "库类"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   10680
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "数量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   10680
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5640
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单价"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5640
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "批次"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料规格"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料单位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "Formc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If DBCombo1(7).Text = "" Or DBCombo1(1).Text = "" Or DBCombo1(10).Text = "" Then
Exit Sub
End If
Data1.Recordset.AddNew
For i = 0 To 10
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If DBCombo1(7).Text = "" Or DBCombo1(1).Text = "" Or DBCombo1(10).Text = "" Then
Exit Sub
End If
If MsgBox("确定修改吗", vbYesNo) = vbNo Then Exit Sub
If DBCombo1(7).Text = "" Then Exit Sub
Data1.Recordset.Edit
For i = 0 To 10
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

End Sub

Private Sub Command4_Click()
If Data1.Recordset.EOF Then Exit Sub
If MsgBox("确定删除吗", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub Command5_Click()
Call OutDataToExcel(MSFlexGrid1, 10, DBCombo1(7).Text)
End Sub

Private Sub Command6_Click()
If MsgBox("确认转入日期为：" + DBCombo1(7).Text + " 正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确认转入日期为：" + DBCombo1(7).Text + " 再次确认", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确认转入日期为：" + DBCombo1(7).Text + " 最后确认？", vbYesNo) = vbNo Then Exit Sub
Data1.Database.Execute "DELETE * FROM CKGL WHERE 日期=CDATE('" & DBCombo1(7).Text & "') AND 库别='清库库存'"
Data1.Database.Execute "INSERT INTO CKGL(合约号,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,库类,日期) SELECT 单号,材料名称,材料规格,材料单位,颜色,批次,数量,单价,金额,BL,日期 FROM KCJL WHERE 日期=CDATE('" & DBCombo1(7).Text & "')"
Data1.Database.Execute "UPDATE CKGL SET 库别='清库库存',单据号='00000000',序号='0' WHERE 库别=NULL"
MsgBox ("导入成功！")
End Sub

Private Sub Command7_Click()
Data1.RecordSource = "SELECT * FROM KCJL WHERE 材料名称='" & DBCombo1(1).Text & "' AND 日期=CDATE('" & DBCombo1(7).Text & "')"
Data1.Refresh
End Sub

Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 7
Data1.RecordSource = "SELECT * FROM KCJL WHERE 日期=CDATE('" & DBCombo1(7).Text & "')"
Data1.Refresh
       Case 6
       If DBCombo1(6).Text <> "" And DBCombo1(8).Text <> "" Then
       DBCombo1(9).Text = Format(DBCombo1(6).Text * DBCombo1(8).Text, "#0.00")
       End If
       Case 8
       If DBCombo1(6).Text <> "" And DBCombo1(8).Text <> "" Then
       DBCombo1(9).Text = Format(DBCombo1(6).Text * DBCombo1(8).Text, "#0.00")
       End If
       Case 10
       Data1.RecordSource = "SELECT * FROM KCJL WHERE bl='" & DBCombo1(10).Text & "' and 日期=CDATE('" & DBCombo1(7).Text & "')"
       Data1.Refresh
End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
Select Case Index
       Case 7
       Data1.RecordSource = "SELECT * FROM KCJL WHERE 日期=CDATE('" & DBCombo1(7).Text & "')"
       Data1.Refresh
       Case 10
       Data1.RecordSource = "SELECT * FROM KCJL WHERE bl='" & DBCombo1(10).Text & "' and 日期=CDATE('" & DBCombo1(7).Text & "')"
       Data1.Refresh
End Select
End Sub

Private Sub Form_Load()
For i = 0 To 10
DBCombo1(i).Text = ""
Next
DBCombo1(7).Text = Date
DBCombo1(6).Text = 0

Data1.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
Data1.RecordSource = "SELECT * FROM KCJL WHERE 日期=CDATE('" & DBCombo1(7).Text & "')"
Data1.Refresh
Data2.DatabaseName = "d:\数据库\\htgl\2011\CKGL.mdb"
Data2.RecordSource = "select KL.MC from KL  GROUP BY KL.MC"
Data2.Refresh
Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select 简称 from GYS GROUP BY 简称"
Data3.Refresh
Data4.DatabaseName = "d:\数据库\\htgl\2011\CKGL.mdb"
Data4.RecordSource = "select MC from CLDW GROUP BY MC"
Data4.Refresh
MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1600
For i = 3 To 11
MSFlexGrid1.ColWidth(i) = 1200
Next
MSFlexGrid1.ColWidth(7) = 0
MSFlexGrid1.ColWidth(10) = 0
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data1.Recordset.EOF Then Exit Sub
rs = MSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To 10
If i = 7 Then i = i + 1
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next
End Sub
