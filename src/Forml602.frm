VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Forml602 
   BackColor       =   &H00C0E0FF&
   Caption         =   "裁剪日报"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form42"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data6 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data5 
      Caption         =   "Data2"
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data4 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command8 
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   720
      TabIndex        =   37
      Top             =   1800
      Width           =   2775
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "结束"
         Height          =   375
         Left            =   1560
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "进行"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   12
      Left            =   11160
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   9480
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   7800
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   9
      Left            =   4080
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3495
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2640
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3240
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
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2640
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3240
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3240
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
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2640
      Width           =   975
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   13440
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   6120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   12360
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   11160
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   9840
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   8280
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   960
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
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6120
      TabIndex        =   17
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
      ItemData        =   "Forml602.frx":0000
      Left            =   4320
      List            =   "Forml602.frx":0010
      TabIndex        =   28
      Text            =   "Combo1"
      Top             =   1800
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7800
      TabIndex        =   35
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
      Bindings        =   "Forml602.frx":002C
      Height          =   1815
      Left            =   8880
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2280
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Forml602.frx":0040
      Height          =   4815
      Left            =   4080
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4080
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   12480
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5655
      Left            =   720
      TabIndex        =   41
      Top             =   3240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9975
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1920
      TabIndex        =   42
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1920
      TabIndex        =   43
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39557
   End
   Begin VB.Label Label1 
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
      Index           =   15
      Left            =   720
      TabIndex        =   45
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Index           =   14
      Left            =   720
      TabIndex        =   44
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
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
      Left            =   11160
      TabIndex        =   34
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发至日期"
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
      Left            =   7800
      TabIndex        =   32
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发至数量"
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
      Left            =   9480
      TabIndex        =   31
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发至"
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
      TabIndex        =   26
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
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
      Left            =   13440
      TabIndex        =   19
      Top             =   600
      Width           =   1095
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
      Index           =   9
      Left            =   6120
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "累计"
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
      Left            =   12480
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "实裁"
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
      Left            =   12360
      TabIndex        =   5
      Top             =   600
      Width           =   975
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
      Index           =   6
      Left            =   11160
      TabIndex        =   4
      Top             =   600
      Width           =   1095
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
      Index           =   5
      Left            =   9840
      TabIndex        =   3
      Top             =   600
      Width           =   1215
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
      Index           =   4
      Left            =   8280
      TabIndex        =   2
      Top             =   600
      Width           =   1335
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
      Index           =   3
      Left            =   6120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
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
      Index           =   2
      Left            =   4080
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Forml602"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public l As Integer


Private Sub Combo1_Click()
Text1(9).Text = Combo1.Text
End Sub

Private Sub Command1_Click()
If Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Or Text1(9).Text = "" Then
MsgBox ("输入不完整!")
Exit Sub
End If

Data1.Recordset.Edit
For i = 0 To 12
Data1.Recordset.Fields(i) = Text1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh

For i = 4 To 5
Text1(i).Text = ""
Next
Data2.Refresh
Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False

End Sub


Private Sub Command3_Click()
On Error Resume Next
If Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Or Text1(9).Text = "" Then
MsgBox ("输入不完整!")
Exit Sub
End If

Data1.Recordset.AddNew
For i = 0 To 12
Data1.Recordset.Fields(i) = Text1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh
For i = 4 To 5
Text1(i).Text = ""
Next
Data2.Refresh
Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If
End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
On Error Resume Next
For i = 4 To 12
If i = 7 Then i = 8
If i = 11 Then i = 12
Text1(i).Text = ""
Next

Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from cjrb where 款号='" & Text1(1).Text & "' AND 颜色='" & Text1(2).Text & "' AND 规格='" & Text1(3).Text & "' order by 日期,序号 desc"
Data1.Refresh


Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM cjrb where  款号='" & Text1(1).Text & "' AND 颜色='" & Text1(2).Text & "' AND 规格='" & Text1(3).Text & "'"
Data2.Refresh

Data3.RecordSource = "select 颜色,规格,计划,sum(val(裁剪)) as 累计裁剪 from cjrb where 款号='" & Text1(1).Text & "' group by 颜色,规格,计划"
Data3.Refresh

Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False
End Sub

Private Sub Command6_Click()
On Error Resume Next
If MsgBox("确定删除吗？，删除不能回复！", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
For i = 4 To 5
Text1(i).Text = ""
Next
Data2.Refresh
Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False

End Sub

Private Sub Command7_Click()
Call cjbb(MSFlexGrid2, "报表日期" + Text1(7).Text)
End Sub


Private Sub Command8_Click()
Call tree
Call zk
End Sub

Private Sub DTPicker1_Change()
On Error Resume Next
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from cjrb where 日期=cdate('" & Text1(7).Text & "') order by 序号 desc"
Data1.Refresh
Text1(7).Text = DTPicker1.Value

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM cjrb where 日期=cdate('" & Text1(7).Text & "')"
Data2.Refresh

Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If
End Sub

Private Sub DTPicker1_CloseUp()
On Error Resume Next
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from cjrb where 日期=cdate('" & Text1(7).Text & "') order by 序号 desc"
Data1.Refresh
Text1(7).Text = DTPicker1.Value

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM cjrb where 日期=cdate('" & Text1(7).Text & "')"
Data2.Refresh

Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If
End Sub


Private Sub DTPicker2_Change()
Text1(10).Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text1(10).Text = DTPicker2.Value
End Sub

Private Sub Form_Load()
On Error Resume Next
DBCombo2.Text = ""
For i = 0 To 12
Text1(i).Text = ""
Next
Text1(7).Text = Date
DTPicker1.Value = Date
Text1(10).Text = Date
DTPicker2.Value = Date
DTPicker3.Value = Date - 30
DTPicker4.Value = Date

Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from cjrb where 款号='" & Text1(1).Text & "' AND 颜色='" & Text1(2).Text & "' AND 规格='" & Text1(3).Text & "' order by 日期,序号 desc"
Data1.Refresh
Option4.Value = True

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM cjrb where  款号='" & Text1(1).Text & "' AND 颜色='" & Text1(2).Text & "' AND 规格='" & Text1(3).Text & "'"
Data2.Refresh

Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If

Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data4.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data5.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data6.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(7) = 0
For i = 1 To 5
MSFlexGrid2.ColWidth(i) = 1200
Next
Command3.Enabled = True
Command1.Enabled = False
Command6.Enabled = False
End Sub


Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 2
khbl = 4
Forml202.Text1.Text = Text1(0).Text
Forml202.Show
End Select
End Sub

Private Sub MSFlexGrid2_dblClick()
On Error Resume Next
If Data1.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To 12
Text1(i).Text = Data1.Recordset.Fields(i)
Next
Text1(7).Text = Data1.Recordset.Fields(7)
Command3.Enabled = False
Command1.Enabled = True
Command6.Enabled = True

End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 1, 2, 3
Data3.RecordSource = "select 颜色,规格,计划,sum(val(裁剪)) as 累计裁剪 from cjrb where 款号='" & Text1(1).Text & "' group by 颜色,规格,计划"
Data3.Refresh

Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data1.RecordSource = "select * from cjrb where 款号='" & Text1(1).Text & "' AND 颜色='" & Text1(2).Text & "' AND 规格='" & Text1(3).Text & "' order by 日期,序号 desc"
Data1.Refresh


Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.RecordSource = "SELECT MAX(序号) FROM cjrb where  款号='" & Text1(1).Text & "' AND 颜色='" & Text1(2).Text & "' AND 规格='" & Text1(3).Text & "'"
Data2.Refresh
Text1(8).Text = 1
If Data2.Recordset.EOF Then
Text1(8).Text = 1
Else
Text1(8).Text = Data2.Recordset.Fields(0) + 1
End If

       Case 5
Text1(6).Text = l + Val(Text1(5).Text)
Text1(11).Text = Val(Text1(5).Text)
End Select

End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Data1.DatabaseName = "e:\excel\sjzz.MDB"
If InStr(TreeView1.Nodes(Node.Index).FullPath, "\") = 0 Then
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
If InStr(l1, "\") > 0 Then
l1 = Mid(l1, 1, InStr(l1, "\") - 1)
Else
l1 = Mid(TreeView1.Nodes(Node.Index).FullPath, InStr(TreeView1.Nodes(Node.Index).FullPath, "\") + 1)
End If
Text1(0).Text = l1
End If


'DBCombo2.Text = Node.Index
'DBCombo3.Text = TreeView1.Nodes(Node.Index).FullPath
End Sub


Private Sub zk()
  For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Expanded = True '展开所有节点
  Next i
End Sub


Private Sub tree()
    Dim mNode As Node
    Dim i As Integer
    Dim intIndex
    Dim xntIndex
   TreeView1.Nodes.Clear
 

If Option4.Value = True Then
    Data4.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
    Data4.Refresh
    m = 1
    If Not Data4.Recordset.EOF Then  'make sure there are records in the table
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add()
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data4.Recordset.Fields(0)
        intIndex = mNode.Index
        Data5.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data4.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='进行'"
        Data5.Refresh
        
        If Not Data5.Recordset.EOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data5.Recordset.Fields(0))
        xntIndex = mNode.Index
        Data6.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data5.Recordset.Fields(0) & "' and 进度='进行'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(xntIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        Data6.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        
        Data5.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        Data4.Recordset.MoveNext
        m = m + 1
        Loop
    End If
End If


If Option5.Value = True Then
    Data4.RecordSource = "select distinct 客户 from sczy_xdh where 日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='结束'"
    Data4.Refresh
    m = 1
    If Not Data4.Recordset.EOF Then  'make sure there are records in the table
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(, , Data4.Recordset.Fields(0), Data4.Recordset.Fields(0))
        mNode.Key = "t" + Trim(m)
        mNode.Text = Data4.Recordset.Fields(0)
        intIndex = mNode.Index
        Data5.RecordSource = "select distinct 单号 from sczy_xdh where 客户='" & Data4.Recordset.Fields(0) & "' and  日期 between cdate('" & DTPicker3.Value & "') and cdate('" & DTPicker4.Value & "') and 进度='结束'"
        Data5.Refresh
        
        If Not Data5.Recordset.EOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
        
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex)
        mNode.Text = Trim(Data5.Recordset.Fields(0))
        intIndex = mNode.Index
        Data6.RecordSource = "select distinct 款号 from sczy_xdh where 单号='" & Data5.Recordset.Fields(0) & "' and 进度='结束'"
        Data6.Refresh
        
        If Not Data6.Recordset.EOF Then
        Data6.Recordset.MoveFirst
        Do While Not Data6.Recordset.EOF
        Set mNode = TreeView1.Nodes.Add(intIndex, tvwChild)
        mNode.Key = "t" + Trim(m) + "w" + Trim(intIndex) + "x" + Trim(xntIndex)
        mNode.Text = Trim(Data6.Recordset.Fields(0))
        Data6.Recordset.MoveNext
        m = m + 1
        Loop
        End If
        m = m + 1
        Data5.Recordset.MoveNext
        Loop
        End If
        m = m + 1
        Data4.Recordset.MoveNext
        Loop
    End If
End If

End Sub


