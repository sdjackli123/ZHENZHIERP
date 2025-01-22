VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Formw37 
   BackColor       =   &H00C0E0FF&
   Caption         =   "收款发票期初设置"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form4"
   ScaleHeight     =   9315
   ScaleWidth      =   9270
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Text            =   "Text3"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10920
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw37.frx":0000
      Height          =   5895
      Left            =   480
      TabIndex        =   9
      Top             =   3000
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   10790143
      BackColorBkg    =   44718
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw37.frx":0014
      Height          =   330
      Left            =   1680
      TabIndex        =   10
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22806529
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   22806529
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   22806529
      CurrentDate     =   36892
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "未开金额"
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
      TabIndex        =   21
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "开票金额"
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
      Left            =   480
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "期初日期"
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
      TabIndex        =   18
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "加工单位"
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
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   15
      Top             =   960
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
      Index           =   0
      Left            =   6240
      TabIndex        =   14
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Formw37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If DBCombo1.Text = "" Or Text3.Text = "" Then
MsgBox ("输入错误")
Exit Sub
End If
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DBCombo1.Text
Data2.Recordset.Fields(1) = Text3.Text
Data2.Recordset.Fields(2) = Text4.Text
Data2.Recordset.Fields(3) = CDate(DTPicker1.Value)
Data2.Recordset.Update
Data2.Refresh
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command2_Click()
On Error Resume Next
If DBCombo1.Text = "" Or Text3.Text = "" Then
MsgBox ("输入错误")
Exit Sub
End If
If MsgBox("确认修改吗?", vbYesNo) = vbNo Then Exit Sub
Data2.Recordset.Edit
Data2.Recordset.Fields(0) = DBCombo1.Text
Data2.Recordset.Fields(1) = Text3.Text
Data2.Recordset.Fields(2) = Text4.Text
Data2.Recordset.Fields(3) = CDate(DTPicker1.Value)
Data2.Recordset.Update
Data2.Refresh
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确认删除吗?", vbYesNo) = vbNo Then Exit Sub
Data2.Recordset.Delete
Data2.Refresh
Text3.Text = ""
Text4.Text = ""
End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
rqq = CDate(Text2.Text)
Data2.RecordSource = "select * from PMFHFP where  结转日期 between cdate('" & Text1 & "') and cdate('" & rqq & "') ORDER BY 结转日期 DESC"
Data2.Refresh
End Sub

Private Sub Command6_Click()
Call MXOutDataToExcel(MSFlexGrid1, "发票打印")
End Sub

Private Sub DTPicker3_Change()
Text1.Text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = DTPicker3.Value
Text1.SetFocus
End Sub

Private Sub DTPicker4_Change()
Text2.Text = DTPicker4.Value
End Sub

Private Sub DTPicker4_CloseUp()
Text2.Text = DTPicker4.Value
Text2.SetFocus
End Sub

Private Sub Form_Load()
Text1.Text = Date
Text2.Text = Date
Text3.Text = ""
Text4.Text = ""
DTPicker1.Value = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
DBCombo1.Text = ""
Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\SCZYJHD.mdb"
Data1.RecordSource = "select 简称 from KHZL  GROUP BY 简称"
Data1.Refresh
rqq = CDate(Text2.Text)
Data2.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.MDB"
Data2.RecordSource = "select * from PMFHFP where  结转日期 between cdate('" & Text1 & "') and cdate('" & rqq & "') ORDER BY 结转日期 DESC"
Data2.Refresh
MSFlexGrid1.ColWidth(1) = 2500
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 2500
End Sub
Private Sub MSFlexGrid1_Click()
On Error Resume Next
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid1.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
DBCombo1.Text = Data2.Recordset.Fields(0)
Text3.Text = Data2.Recordset.Fields(1)
Text4.Text = Data2.Recordset.Fields(2)
DTPicker1.Value = Data2.Recordset.Fields(3)
End Sub



