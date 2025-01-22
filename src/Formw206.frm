VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw206 
   BackColor       =   &H00C0E0FF&
   Caption         =   "装箱查询"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
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
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "日期查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1335
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
      Top             =   10200
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
      Top             =   10080
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
      Top             =   9960
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
      Top             =   9840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formw206.frx":0000
      Height          =   6855
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2520
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   12091
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   84541441
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   84541441
      CurrentDate     =   39557
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
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   480
      Width           =   1455
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
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   3135
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
      Index           =   17
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Formw206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Call zxd(MSFlexGrid2, "装箱单明细")
End Sub

Private Sub Command2_Click()
Data1.RecordSource = "select * from zxd where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
Data1.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Data1.RecordSource = "select * from zxd where instr(款号,'" & DBCombo2.Text & "')>0"
Data1.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
DBCombo2.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Data1.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\scjd.mdb"
Data2.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\scjd.mdb"
Data3.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\scjd.MDB"
MSFlexGrid2.ColWidth(0) = 300
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
khbl = 22
Formw202.Text1.Text = DBCombo2.Text
Formw202.Show
End Select
End Sub


