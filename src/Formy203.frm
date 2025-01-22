VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy203 
   BackColor       =   &H00C0E0FF&
   Caption         =   "装箱明细单"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form42"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新编号"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   17
      Left            =   1920
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导出"
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
      TabIndex        =   45
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command6 
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4320
      Width           =   975
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3255
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   16
      Left            =   13080
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   15
      Left            =   12360
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   14
      Left            =   11640
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   13
      Left            =   11040
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   12
      Left            =   13560
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   12480
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   11040
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   13200
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   11640
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   10200
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   8760
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   7320
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   3720
      Width           =   1455
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy203.frx":0000
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   13080
      TabIndex        =   41
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39557
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy203.frx":0014
      Height          =   4335
      Left            =   720
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4920
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "装箱编号"
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
      Index           =   18
      Left            =   720
      TabIndex        =   47
      Top             =   2880
      Width           =   1215
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
      Index           =   17
      Left            =   13080
      TabIndex        =   23
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "体积"
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
      Index           =   16
      Left            =   12360
      TabIndex        =   22
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "合计件"
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
      Index           =   15
      Left            =   13560
      TabIndex        =   21
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛重"
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
      Left            =   11040
      TabIndex        =   20
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "净重"
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
      Index           =   13
      Left            =   11640
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "箱数"
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
      Index           =   12
      Left            =   12480
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格5"
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
      Left            =   11640
      TabIndex        =   17
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格6"
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
      Index           =   10
      Left            =   13200
      TabIndex        =   16
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "件/每箱"
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
      Left            =   11040
      TabIndex        =   15
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格4"
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
      Left            =   10200
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格1"
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
      Left            =   5760
      TabIndex        =   13
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格2"
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
      Left            =   7320
      TabIndex        =   12
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格3"
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
      Left            =   8760
      TabIndex        =   11
      Top             =   3360
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
      Index           =   8
      Left            =   4800
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户"
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
      Left            =   720
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
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
      Index           =   4
      Left            =   2280
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "箱号"
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
      Left            =   3600
      TabIndex        =   7
      Top             =   3360
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
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Formy203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data1.Recordset.Edit
For i = 0 To 17
Data1.Recordset.Fields(i) = Text1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

Text1(4).Text = ""
Text1(5).Text = ""
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command2_Click()
Data2.RecordSource = "SELECT SCZY_ZDH.客户,cmb.单号,cmb.款号,cmb.颜色,cmb.尺码,cmb.数量 FROM SCZY_ZDH,cmb WHERE SCZY_ZDH.单号=cmb.单号 and instr(cmb.款号,'" & DBCombo2.Text & "')>0 order BY cmb.颜色,cmb.尺码"
Data2.Refresh
End Sub

Private Sub Command3_Click()
If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data1.Recordset.AddNew
For i = 0 To 17
Data1.Recordset.Fields(i) = Text1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

Text1(4).Text = ""
Text1(5).Text = ""
Text1(4).SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from zxd where 编号='" & Text1(17).Text & "'"
Data1.Refresh
End Sub

Private Sub Command6_Click()
If MsgBox("删除不能回复！确定删除吗？", vbYesNo) = vbNo Then Exit Sub
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
Text1(4).Text = ""
Text1(5).Text = ""
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command8_Click()
On Error Resume Next
Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data3.RecordSource = "SELECT MAX(VAL(MID(编号,3))) FROM zxd"
Data3.Refresh
Text1(17).Text = "ZX000001"
If Data3.Recordset.EOF Then
Text1(17).Text = "ZX000001"
Else
Text1(17).Text = Left("ZX000000", 8 - Len(Trim(Data3.Recordset.Fields(0) + 1))) + Trim(Data3.Recordset.Fields(0) + 1)
End If

Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from zxd where 编号='" & Text1(17).Text & "'"
Data1.Refresh
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False
End Sub

Private Sub DTPicker1_Change()
Text1(16).Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text1(16).Text = DTPicker1.Value
End Sub

Private Sub Form_Load()
On Error Resume Next
DBCombo2.Text = ""
For i = 0 To 16
Text1(i).Text = ""
Next
Text1(16).Text = Date
DTPicker1.Value = Date
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data3.RecordSource = "SELECT MAX(VAL(MID(编号,3))) FROM zxd"
Data3.Refresh

Text1(17).Text = "ZX000001"
If Data3.Recordset.EOF Then
Text1(17).Text = "ZX000001"
Else
Text1(17).Text = Left("ZX000000", 8 - Len(Trim(Data3.Recordset.Fields(0) + 1))) + Trim(Data3.Recordset.Fields(0) + 1)
End If

Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from zxd where 编号='" & Text1(17).Text & "'"
Data1.Refresh


MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(4) = 2000
MSFlexGrid1.ColWidth(12) = 1300

MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(4) = 1800
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
khbl = 2
Formy202.Text1.Text = DBCombo2.Text
Formy202.Show
End Select
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data4.Recordset.Move rs - 1
For i = 0 To 17
Text1(i).Text = Data4.Recordset.Fields(i)
Next
DTPicker1.Value = Text1(16).Text
Command3.Enabled = False
Command1.Enabled = True
Command6.Enabled = True
End Sub

Private Sub MSFlexGrid2_dblClick()
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
Text1(0).Text = Data2.Recordset.Fields(0)
Text1(1).Text = Data2.Recordset.Fields(2)
Text1(3).Text = Data2.Recordset.Fields(3)
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 4
Text1(10).Text = Val(Mid(Text1(4).Text, InStr(Text1(4).Text, "/") + 1)) + Val(Mid(Text1(5).Text, InStr(Text1(5).Text, "/") + 1)) + Val(Mid(Text1(6).Text, InStr(Text1(6).Text, "/") + 1)) + Val(Mid(Text1(7).Text, InStr(Text1(7).Text, "/") + 1)) + Val(Mid(Text1(8).Text, InStr(Text1(8).Text, "/") + 1)) + Val(Mid(Text1(9).Text, InStr(Text1(9).Text, "/") + 1))
       Case 5
Text1(10).Text = Val(Mid(Text1(4).Text, InStr(Text1(4).Text, "/") + 1)) + Val(Mid(Text1(5).Text, InStr(Text1(5).Text, "/") + 1)) + Val(Mid(Text1(6).Text, InStr(Text1(6).Text, "/") + 1)) + Val(Mid(Text1(7).Text, InStr(Text1(7).Text, "/") + 1)) + Val(Mid(Text1(8).Text, InStr(Text1(8).Text, "/") + 1)) + Val(Mid(Text1(9).Text, InStr(Text1(9).Text, "/") + 1))
       Case 6
Text1(10).Text = Val(Mid(Text1(4).Text, InStr(Text1(4).Text, "/") + 1)) + Val(Mid(Text1(5).Text, InStr(Text1(5).Text, "/") + 1)) + Val(Mid(Text1(6).Text, InStr(Text1(6).Text, "/") + 1)) + Val(Mid(Text1(7).Text, InStr(Text1(7).Text, "/") + 1)) + Val(Mid(Text1(8).Text, InStr(Text1(8).Text, "/") + 1)) + Val(Mid(Text1(9).Text, InStr(Text1(9).Text, "/") + 1))
       Case 7
Text1(10).Text = Val(Mid(Text1(4).Text, InStr(Text1(4).Text, "/") + 1)) + Val(Mid(Text1(5).Text, InStr(Text1(5).Text, "/") + 1)) + Val(Mid(Text1(6).Text, InStr(Text1(6).Text, "/") + 1)) + Val(Mid(Text1(7).Text, InStr(Text1(7).Text, "/") + 1)) + Val(Mid(Text1(8).Text, InStr(Text1(8).Text, "/") + 1)) + Val(Mid(Text1(9).Text, InStr(Text1(9).Text, "/") + 1))
       Case 8
Text1(10).Text = Val(Mid(Text1(4).Text, InStr(Text1(4).Text, "/") + 1)) + Val(Mid(Text1(5).Text, InStr(Text1(5).Text, "/") + 1)) + Val(Mid(Text1(6).Text, InStr(Text1(6).Text, "/") + 1)) + Val(Mid(Text1(7).Text, InStr(Text1(7).Text, "/") + 1)) + Val(Mid(Text1(8).Text, InStr(Text1(8).Text, "/") + 1)) + Val(Mid(Text1(9).Text, InStr(Text1(9).Text, "/") + 1))
       Case 9
Text1(10).Text = Val(Mid(Text1(4).Text, InStr(Text1(4).Text, "/") + 1)) + Val(Mid(Text1(5).Text, InStr(Text1(5).Text, "/") + 1)) + Val(Mid(Text1(6).Text, InStr(Text1(6).Text, "/") + 1)) + Val(Mid(Text1(7).Text, InStr(Text1(7).Text, "/") + 1)) + Val(Mid(Text1(8).Text, InStr(Text1(8).Text, "/") + 1)) + Val(Mid(Text1(9).Text, InStr(Text1(9).Text, "/") + 1))
       Case 10
Text1(12).Text = Val(Text1(10).Text) * Val(Text1(11).Text)
       Case 11
Text1(12).Text = Val(Text1(10).Text) * Val(Text1(11).Text)
End Select
End Sub
