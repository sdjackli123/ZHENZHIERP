VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy204 
   BackColor       =   &H00C0E0FF&
   Caption         =   "入库月报表"
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
      Caption         =   "查询"
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3840
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   240
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
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
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
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Width           =   975
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
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   7080
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   8520
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   9960
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   10920
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   12360
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   13800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   10920
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   12
      Left            =   12120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   13
      Left            =   13440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   14
      Left            =   10920
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   15
      Left            =   13080
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   10080
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
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   3255
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
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
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
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   975
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
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy204.frx":0000
      Height          =   1815
      Left            =   600
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   720
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
      Left            =   1680
      TabIndex        =   24
      Top             =   240
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
      TabIndex        =   25
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81133569
      CurrentDate     =   39557
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy204.frx":0014
      Height          =   5055
      Left            =   600
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4320
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
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
      ItemData        =   "Formy204.frx":0028
      Left            =   10920
      List            =   "Formy204.frx":004A
      TabIndex        =   44
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8520
      TabIndex        =   45
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81133569
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   11760
      TabIndex        =   46
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81133569
      CurrentDate     =   36892
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期："
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
      Left            =   10560
      TabIndex        =   48
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期："
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
      Left            =   7440
      TabIndex        =   47
      Top             =   3840
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
      Left            =   600
      TabIndex        =   43
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "计划"
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
      Left            =   5880
      TabIndex        =   42
      Top             =   2760
      Width           =   1095
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
      Left            =   2160
      TabIndex        =   41
      Top             =   2760
      Width           =   1215
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
      Left            =   600
      TabIndex        =   40
      Top             =   2760
      Width           =   1455
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
      Left            =   3480
      TabIndex        =   39
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "出货"
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
      Left            =   8520
      TabIndex        =   38
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "裁剪"
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
      Left            =   7080
      TabIndex        =   37
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格"
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
      Left            =   4680
      TabIndex        =   36
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户带走"
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
      Left            =   9960
      TabIndex        =   35
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "差"
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
      Left            =   13800
      TabIndex        =   34
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "应入库"
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
      Left            =   12360
      TabIndex        =   33
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "实入库"
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
      Left            =   10920
      TabIndex        =   32
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "布残未做"
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
      Left            =   10920
      TabIndex        =   31
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "原因"
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
      Left            =   10920
      TabIndex        =   30
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "从裁剪取"
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
      Left            =   13440
      TabIndex        =   29
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "样板室"
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
      Left            =   12120
      TabIndex        =   28
      Top             =   1800
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
      TabIndex        =   27
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Formy204"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Text1(14).Text = Combo1.Text
End Sub

Private Sub Command1_Click()
If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Then
MsgBox ("输入不完整")
Exit Sub
End If
Data1.Recordset.Edit
For i = 0 To 15
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
For i = 0 To 15
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
Data1.RecordSource = "select * from bzrk where 日期=cdate('" & Text1(15).Text & "')"
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

Private Sub Command7_Click()
Call bzrk(MSFlexGrid1, "包装入库报表")
End Sub

Private Sub Command8_Click()
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from bzrk where 日期 between cdate('" & DTPicker2.Value & "') and cdate('" & DTPicker3.Value & "')"
Data1.Refresh
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
For i = 0 To 15
Text1(i).Text = ""
Next
Text1(15).Text = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"


Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from bzrk where 日期=cdate('" & Text1(15).Text & "')"
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
khbl = 3
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
For i = 0 To 15
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
For i = 0 To 4
Text1(i).Text = Data2.Recordset.Fields(i)
Next
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 5
Text1(9).Text = Val(Text1(5).Text) - Val(Text1(6).Text) - Val(Text1(7).Text)
       Case 6
Text1(9).Text = Val(Text1(5).Text) - Val(Text1(6).Text) - Val(Text1(7).Text)
       Case 7
Text1(9).Text = Val(Text1(5).Text) - Val(Text1(6).Text) - Val(Text1(7).Text)
       Case 8
Text1(10).Text = Val(Text1(9).Text) - Val(Text1(8).Text)
       Case 9
Text1(10).Text = Val(Text1(9).Text) - Val(Text1(8).Text)
End Select
End Sub

