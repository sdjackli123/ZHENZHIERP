VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formy301 
   BackColor       =   &H00C0E0FF&
   Caption         =   "纱线分析及采购"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form41"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data7 
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data6 
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data5 
      Caption         =   "Data4"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
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
      TabIndex        =   39
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "类别"
      Height          =   975
      Left            =   3840
      TabIndex        =   35
      Top             =   3000
      Width           =   1215
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "染色"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "色织"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   13560
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   4440
      Width           =   1095
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   31
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4440
      Width           =   975
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4080
      Width           =   975
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4440
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   13800
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   12120
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   10920
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   9720
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   8160
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   8640
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   7320
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   11160
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
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
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
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
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
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
      Top             =   10200
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
      Height          =   855
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   375
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   4320
      TabIndex        =   0
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy301.frx":0000
      Height          =   1935
      Left            =   3840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy301.frx":0014
      Height          =   4575
      Left            =   3840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5160
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   13
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   13560
      TabIndex        =   34
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   36892
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7695
      Left            =   360
      TabIndex        =   40
      Top             =   2040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   13573
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1440
      TabIndex        =   41
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1440
      TabIndex        =   42
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39177
   End
   Begin VB.Label Label6 
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
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   44
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   43
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "期限"
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
      Left            =   13560
      TabIndex        =   33
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "裁耗"
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
      Left            =   8160
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料数量"
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
      Left            =   13800
      TabIndex        =   24
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "织耗"
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
      Left            =   9720
      TabIndex        =   13
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "染耗"
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
      Left            =   10920
      TabIndex        =   12
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "采购数量"
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
      Left            =   12120
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料名称"
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
      Left            =   11160
      TabIndex        =   10
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "纱支"
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
      Left            =   7320
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "配比"
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
      Left            =   8640
      TabIndex        =   8
      Top             =   4080
      Width           =   855
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
      Left            =   5520
      TabIndex        =   7
      Top             =   3000
      Width           =   2055
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
      Left            =   7920
      TabIndex        =   6
      Top             =   3000
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
      Left            =   9600
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单号"
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
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Formy301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer
Private Sub Command1_Click()
Call sxmx(MSFlexGrid1, "购纱明细")
End Sub

Private Sub Command2_Click()
If Option1.Value = False And Option2.Value = False Then
MsgBox ("请选择类别")
Exit Sub
End If


If Option1.Value = True Then
Data2.RecordSource = "SELECT  单号,材料库类,材料名称,材料颜色,材料规格,材料单位,材料批号,材料数量 as 采购量,订单颜色 as 采购类别 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "' and  材料库类='1主料库' and instr(材料名称,'色织')>0  order by 材料库类,材料名称,材料规格,材料单位,材料批号"
Data2.Refresh
Data4.RecordSource = "select * from sxcg where 单号='" & DBCombo1.Text & "' order by 颜色"
Data4.Refresh
End If


If Option2.Value = True Then
Data2.RecordSource = "SELECT  单号,材料库类,材料名称,材料颜色,材料规格,材料单位,材料批号,材料数量 as 采购量,订单颜色 as 采购类别 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "' and  材料库类='1主料库' and instr(材料名称,'色织')=0 order by 材料库类,材料名称,材料规格,材料单位,材料批号 "
Data2.Refresh
Data4.RecordSource = "select * from sxcg where 单号='" & DBCombo1.Text & "' order by 颜色"
Data4.Refresh
End If


End Sub



Private Sub Command3_Click()
Call tree
Call zk
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
If MsgBox("删除不能回复！确定删除吗？", vbYesNo) = vbNo Then Exit Sub
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.Delete
Data4.Refresh
Text1(4).Text = ""
Text1(5).Text = ""
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command7_Click()

If S1 = 0 Or S2 = 0 Then
MsgBox ("请选择记录！")
Exit Sub
End If

If S1 < 1 Or S2 < 1 Then
MsgBox ("选择记录")
Exit Sub
End If

If S1 > S2 Then
MsgBox ("注意选择顺序！")
Exit Sub
End If

k = S2 - S1
If k = 0 Then
Data2.Recordset.MoveFirst
Data2.Recordset.Move S1 - 1
Text1(0).Text = Data2.Recordset.Fields(0)
Text1(1).Text = ""
Text1(2).Text = Data2.Recordset.Fields(3)
Text1(3).Text = Data2.Recordset.Fields(2)
Text1(10).Text = Data2.Recordset.Fields(7)

If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(9).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If

Data4.Recordset.AddNew
For i = 0 To 11
Data4.Recordset.Fields(i) = Text1(i).Text
Next
Data4.Recordset.Update
Else
Data2.Recordset.MoveFirst
Data2.Recordset.Move S1 - 1
For i = 1 To k + 1
Text1(0).Text = Data2.Recordset.Fields(0)
Text1(1).Text = ""
Text1(2).Text = Data2.Recordset.Fields(3)
Text1(3).Text = Data2.Recordset.Fields(2)
Text1(10).Text = Data2.Recordset.Fields(7)

If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(9).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If

Data4.Recordset.AddNew
For i1 = 0 To 11
Data4.Recordset.Fields(i1) = Text1(i1).Text
Next
Data4.Recordset.Update
Data2.Recordset.MoveNext
Next
End If

Data4.Refresh

Text1(4).SetFocus
End Sub

Private Sub Command8_Click()
If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(9).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data4.Recordset.Edit
For i = 0 To 11
Data4.Recordset.Fields(i) = Text1(i).Text
Next
Data4.Recordset.Update
Data4.Refresh

Text1(4).Text = ""
Text1(5).Text = ""
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command9_Click()
Data4.RecordSource = "select * from sxcg where 单号='" & DBCombo1.Text & "'  order by 款号,材料名称,颜色,纱支"
Data4.Refresh
Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub DBCombo1_Change()
Data2.RecordSource = "SELECT  单号,材料库类,材料名称,材料颜色,材料规格,材料单位,材料批号,材料数量 as 采购量,订单颜色 as 采购类别 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "' and  材料库类='1主料库'   order by 材料库类,材料名称,材料规格,材料单位,材料批号"
Data2.Refresh
Data4.RecordSource = "select * from sxcg where 单号='" & DBCombo1.Text & "' order by 颜色"
Data4.Refresh
End Sub
Private Sub DTPicker1_Change()
Text1(11).Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text1(11).Text = DTPicker1.Value
End Sub
Private Sub Form_Load()
DBCombo1.Text = ""
For i = 0 To 11
Text1(i).Text = ""
Next
Text1(11).Text = Date
DTPicker1.Value = Date
'Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
'Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"

Data5.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data6.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data7.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"



Command7.Enabled = True
Command8.Enabled = False
Command6.Enabled = False

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 0
MSFlexGrid1.ColWidth(4) = 1600
MSFlexGrid1.ColWidth(7) = 0
MSFlexGrid1.ColWidth(12) = 1300

MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(4) = 1800
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 1
xqbl = 2
Formy41.Show
End Select
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data4.Recordset.EOF Then Exit Sub
Data4.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data4.Recordset.Move rs - 1
Text1(10).Text = Data4.Recordset.Fields(10)
Text1(11).Text = Data4.Recordset.Fields(11)
DTPicker1.Value = Text1(11).Text
For i = 0 To 9
Text1(i).Text = Data4.Recordset.Fields(i)
Next
Command7.Enabled = False
Command8.Enabled = True
Command6.Enabled = True
End Sub

Private Sub MSFlexGrid2_dblClick()
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
Text1(0).Text = Data2.Recordset.Fields(0)
Text1(1).Text = ""
Text1(2).Text = Data2.Recordset.Fields(3)
Text1(3).Text = Data2.Recordset.Fields(2)
Text1(10).Text = Data2.Recordset.Fields(7)
End Sub


Private Sub Text1_Change(Index As Integer)
Select Case Index
      Case 7
If Val(Text1(5).Text) > 0 Then
Text1(9).Text = Format(Val(Text1(10).Text) * Val(Text1(5).Text) / 100 / (1 - Val(Text1(7).Text) / 100), "#0.00")
Else
Text1(9).Text = Format(Val(Text1(10).Text) / (1 - Val(Text1(7).Text) / 100), "#0.00")
End If
      Case 8
If Val(Text1(5).Text) > 0 Then
Text1(9).Text = Format(Val(Text1(10).Text) * Val(Text1(5).Text) / 100 / (1 - Val(Text1(7).Text) / 100) / (1 - Val(Text1(8).Text) / 100), "#0.00")
Else
Text1(9).Text = Format(Val(Text1(10).Text) / (1 - Val(Text1(7).Text) / 100) / (1 - Val(Text1(8).Text) / 100), "#0.00")
End If
      Case 10
If Val(Text1(5).Text) > 0 Then
Text1(9).Text = Format(Val(Text1(10).Text) * Val(Text1(5).Text) / 100 / (1 - Val(Text1(7).Text) / 100) / (1 - Val(Text1(8).Text) / 100), "#0.00")
Else
Text1(9).Text = Format(Val(Text1(10).Text) / (1 - Val(Text1(7).Text) / 100) / (1 - Val(Text1(8).Text) / 100), "#0.00")
End If

End Select
End Sub

Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid2.RowSel
End Sub

Private Sub MSFlexGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid2.RowSel
End Sub


Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex

   TreeView1.Nodes.Clear
 
    Data6.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
    Data6.Refresh
    m = 1
    If Not Data6.Recordset.EOF Then  'make sure there are records in the table
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "w" + Trim(m)
        mNode.Text = Data6.Recordset.Fields(0)
        intIndex = mNode.Index
        Data5.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data6.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and 进度='进行'"
        Data5.Refresh
        
        If Not Data5.Recordset.EOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex)
        mNode.Text = Trim(Data5.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data7.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data5.Recordset.Fields(0) & "' and 进度='进行'"
        Data7.Refresh
        
        If Not Data7.Recordset.EOF Then
        Data7.Recordset.MoveFirst
        Do While Not Data7.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "w" + Trim(m) + "x" + Trim(intIndex) + "t" + Trim(xntIndex)
        mNode.Text = Trim(Data7.Recordset.Fields(0))
        Data7.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data5.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data6.Recordset.MoveNext
        Loop
    End If

End Sub


'然后该代码就只需对较小的记录集进行循环，因而效率比较高。修改后的代码如下：

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next

If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") > 0 Then
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
DBCombo1.Text = l1
End If

End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub


