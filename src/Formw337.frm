VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Formw337 
   BackColor       =   &H00C0E0FF&
   Caption         =   "账薄浏览"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form37"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data13 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data14 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
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
      Height          =   1095
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formw337.frx":0000
      Left            =   4560
      List            =   "Formw337.frx":0010
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按凭证"
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
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "总分类账"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按日期"
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
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "特种日记账"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细账"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按日期、凭证"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw337.frx":003C
      Height          =   7815
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13785
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   8421631
      BackColorBkg    =   34952
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   83755009
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   83755009
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   83755009
      CurrentDate     =   39883
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "账薄查询"
      Height          =   1215
      Left            =   6960
      TabIndex        =   19
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "分类记账"
      Height          =   1215
      Left            =   10320
      TabIndex        =   20
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   18
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "凭证类别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作月份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Formw337"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2 As String

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("输入日期")
Exit Sub
End If
If Combo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY PZDZ.日期,PZDZ.凭证号"
Data1.Refresh
Else
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.凭证类别='" & Combo1.Text & "' AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY PZDZ.日期,PZDZ.凭证号"
Data1.Refresh
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("输入记账日期范围")
Exit Sub
End If
Data1.Database.Execute "INSERT INTO MXFLZ(日期,凭证号,摘要,会计科目,借方金额,贷方金额) SELECT PZDZ.日期,PZDZ.凭证号,PZDZ.摘要,PZDZ.会计科目,PZDZ.借方金额,PZDZ.贷方金额 FROM PZDZ WHERE (PZDZ.明细类账='' OR PZDZ.明细类账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.会计科目,'-')>0"
Data1.Database.Execute "update  PZDZ SET 明细类账='记' WHERE (PZDZ.明细类账='' OR PZDZ.明细类账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
MsgBox ("明细类记账成功")
Data1.Refresh
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("输入记账日期范围")
Exit Sub
End If
Data1.Database.Execute "DELETE * FROM TZJZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"

Data1.Database.Execute "INSERT INTO TZJZ(日期,凭证号,摘要,对方科目,借方金额,贷方金额) SELECT PZDZ.日期,PZDZ.凭证号,PZDZ.摘要,PZDZ.会计科目,PZDZ.借方金额,PZDZ.贷方金额 FROM PZDZ WHERE (PZDZ.特种日账='' OR PZDZ.特种日账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.凭证号,'1-')>0 AND 会计科目<>'现金'"
Data1.Database.Execute "update  TZJZ SET 类别='现金' WHERE 类别=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(日期,凭证号,摘要,对方科目,借方金额,贷方金额) SELECT PZDZ.日期,PZDZ.凭证号,PZDZ.摘要,PZDZ.会计科目,PZDZ.借方金额,PZDZ.贷方金额 FROM PZDZ WHERE (PZDZ.特种日账='' OR PZDZ.特种日账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.凭证号,'2-')>0  AND 会计科目<>'现金'"
Data1.Database.Execute "update  TZJZ SET 类别='现金' WHERE 类别=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(日期,凭证号,摘要,对方科目,贷方金额,借方金额) SELECT PZDZ.日期,PZDZ.凭证号,PZDZ.摘要,PZDZ.会计科目,PZDZ.借方金额,PZDZ.贷方金额 FROM PZDZ WHERE (PZDZ.特种日账='' OR PZDZ.特种日账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.凭证号,'2-')>0  AND 会计科目='银行存款'"
Data1.Database.Execute "update  TZJZ SET 类别='银行存款' WHERE 类别=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(日期,凭证号,摘要,对方科目,借方金额,贷方金额) SELECT PZDZ.日期,PZDZ.凭证号,PZDZ.摘要,PZDZ.会计科目,PZDZ.借方金额,PZDZ.贷方金额 FROM PZDZ WHERE (PZDZ.特种日账='' OR PZDZ.特种日账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.凭证号,'3-')>0 AND 会计科目<>'银行存款'"
Data1.Database.Execute "update  TZJZ SET 类别='银行存款' WHERE 类别=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(日期,凭证号,摘要,对方科目,借方金额,贷方金额) SELECT PZDZ.日期,PZDZ.凭证号,PZDZ.摘要,PZDZ.会计科目,PZDZ.借方金额,PZDZ.贷方金额 FROM PZDZ WHERE (PZDZ.特种日账='' OR PZDZ.特种日账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.凭证号,'4-')>0 AND 会计科目<>'银行存款'"
Data1.Database.Execute "update  TZJZ SET 类别='银行存款' WHERE 类别=NULL"

Data1.Database.Execute "INSERT INTO TZJZ(日期,凭证号,摘要,对方科目,贷方金额,借方金额) SELECT PZDZ.日期,PZDZ.凭证号,PZDZ.摘要,PZDZ.会计科目,PZDZ.借方金额,PZDZ.贷方金额 FROM PZDZ WHERE (PZDZ.特种日账='' OR PZDZ.特种日账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND INSTR(PZDZ.凭证号,'4-')>0 AND 会计科目='现金'"
Data1.Database.Execute "update  TZJZ SET 类别='现金' WHERE 类别=NULL"

Data1.Database.Execute "update  PZDZ SET 特种日账='记' WHERE (PZDZ.特种日账='' OR PZDZ.特种日账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"



MsgBox ("日记账成功")
Data1.Refresh
End Sub

Private Sub Command4_Click()
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') ORDER BY PZDZ.日期,PZDZ.凭证号"
Data1.Refresh
End Sub

Private Sub Command5_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("输入记账日期范围")
Exit Sub
End If
Data1.Database.Execute "INSERT INTO ZFLZ(日期,凭证号,摘要,会计科目,借方金额,贷方金额) SELECT PZDZ.日期,PZDZ.凭证号,PZDZ.摘要,PZDZ.会计科目,PZDZ.借方金额,PZDZ.贷方金额 FROM PZDZ WHERE (PZDZ.总分类账='' OR PZDZ.总分类账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data1.Database.Execute "update  ZFLZ SET 会计科目=LEFT(ZFLZ.会计科目,INSTR(ZFLZ.会计科目,'-')-1) WHERE INSTR(ZFLZ.会计科目,'-')>0"
Data1.Database.Execute "update  PZDZ SET 总分类账='记' WHERE (PZDZ.总分类账='' OR PZDZ.总分类账=NULL) AND PZDZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
MsgBox ("总账类记账成功")
Data1.Refresh
End Sub

Private Sub Command6_Click()
If Combo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') ORDER BY PZDZ.日期,PZDZ.凭证号"
Data1.Refresh
Else
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.凭证类别='" & Combo1.Text & "' AND PZDZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') ORDER BY PZDZ.日期,PZDZ.凭证号"
Data1.Refresh
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Text1.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text1.Text = DTPicker1.Value
End Sub
Private Sub DTPicker2_Change()
Text2.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text2.Text = DTPicker2.Value
End Sub


Private Sub DTPicker3_Change()
Data13.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data13.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between 起始日期 and 结束日期"
Data13.Refresh
If Data13.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data13.Recordset.Fields(0)
K2 = Data13.Recordset.Fields(1)
Text3.Text = Data13.Recordset.Fields(2)
End If
Text1.Text = K1
Text2.Text = K2
DTPicker1.Value = K1
DTPicker2.Value = K2
End Sub

Private Sub DTPicker3_CloseUp()
Data13.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data13.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between 起始日期 and 结束日期"
Data13.Refresh
If Data13.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data13.Recordset.Fields(0)
K2 = Data13.Recordset.Fields(1)
Text3.Text = Data13.Recordset.Fields(2)
End If
Text1.Text = K1
Text2.Text = K2
DTPicker1.Value = K1
DTPicker2.Value = K2
End Sub

Private Sub Form_Load()
'On Error Resume Next
Text1.Text = Date
DTPicker3.Value = Date
DTPicker1.Value = Date
Text2.Text = Date
DTPicker2.Value = Date

Data13.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data13.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between 起始日期 and 结束日期"
Data13.Refresh
If Data13.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data13.Recordset.Fields(0)
K2 = Data13.Recordset.Fields(1)
Text3.Text = Data13.Recordset.Fields(2)
End If
Text1.Text = K1
Text2.Text = K2
DTPicker1.Value = K1
DTPicker2.Value = K2

Combo1.Text = ""
Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.MDB"
Data1.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') ORDER BY PZDZ.日期,PZDZ.凭证号"
Data1.Refresh

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 2500
MSFlexGrid1.ColWidth(7) = 700
MSFlexGrid1.ColWidth(8) = 700
MSFlexGrid1.ColWidth(9) = 700
End Sub

