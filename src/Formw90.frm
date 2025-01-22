VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formw90 
   BackColor       =   &H00C0E0FF&
   Caption         =   "年度设置"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form18"
   ScaleHeight     =   8175
   ScaleWidth      =   7500
   StartUpPosition =   2  '屏幕中心
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
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   1200
      TabIndex        =   11
      Text            =   "Combo3"
      Top             =   2280
      Width           =   5655
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "Formw90.frx":0000
      Left            =   1200
      List            =   "Formw90.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1680
      Width           =   5655
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Formw90.frx":0011
      Left            =   1200
      List            =   "Formw90.frx":001B
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   1080
      Width           =   5655
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
      Left            =   1200
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   480
      Width           =   5655
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7770
      Visible         =   0   'False
      Width           =   5175
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw90.frx":0027
      Height          =   4335
      Left            =   480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
      _Version        =   393216
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "记录号 "
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
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
      Left            =   600
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "默认"
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
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "路径"
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
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "部门"
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
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "Formw90"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim ba As Database: Dim rr As Integer
Dim rs As Single
Dim rd As Recordset


Private Sub Combo3_Click()
Data1.RecordSource = "SELECT MC AS 路径名称,MR AS 是否默认,BM AS 部门,XH AS 序号 FROM LJB WHERE INSTR('" & Combo3.text & "',BM)>0 ORDER BY VAL(XH)"
Data1.Refresh
Combo4.text = Data1.Recordset.RecordCount + 1
End Sub

Private Sub Command5_Click()
Data1.Refresh
Combo1.text = ""
Combo2.text = ""
Combo4.text = Data1.Recordset.RecordCount + 1
Combo1.SetFocus
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
Combo1.text = Data1.Recordset.Fields(0)
Combo2.text = Data1.Recordset.Fields(1)
Combo4.text = Data1.Recordset.Fields(3)
End Sub
Private Sub JILU()
Dim i As Single
Data1.Refresh
rd.MoveLast
MSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To rd.RecordCount
MSFlexGrid1.TextMatrix(i, 0) = i
Next
End Sub

Private Sub Command3_Click()
On Error Resume Next
Data1.Recordset.Edit
Data1.Recordset.Fields(0) = Combo1.text
Data1.Recordset.Fields(1) = Combo2.text
Data1.Recordset.Fields(2) = Combo3.text
Data1.Recordset.Fields(3) = Combo4.text
Data1.Recordset.Update
Data1.Refresh
Combo1.text = ""
Combo2.text = ""
Combo4.text = Data1.Recordset.RecordCount + 1
Combo1.SetFocus
End Sub

Private Sub Command4_Click()
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
Combo1.text = ""
Combo2.text = ""
Combo4.text = Data1.Recordset.RecordCount + 1
Combo1.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Set ba = OpenDatabase("\\ytyyfw\lbjxx$\" + "LJB.MDB")
Set rd = ba.OpenRecordset("LJB", dbOpenDynaset)
Data1.DatabaseName = "\\ytyyfw\lbjxx$\" + "LJB.MDB"
Data1.RecordSource = "SELECT MC AS 路径名称,MR AS 是否默认,BM AS 部门,XH AS 序号 FROM LJB WHERE BM='财务' ORDER BY VAL(XH)"
Data1.Refresh
Combo1.text = ""
Combo2.text = ""
Combo3.text = ""
Combo4.text = Data1.Recordset.RecordCount + 1
Combo1.TabIndex = 0
MSFlexGrid1.ColWidth(1) = 1500
End Sub
Private Sub Command1_Click()
rd.AddNew
rd.Fields(0) = Combo1.text
rd.Fields(1) = Combo2.text
rd.Fields(2) = Combo3.text
rd.Fields(3) = Combo4.text
rd.Update
Data1.Refresh
Combo1.text = ""
Combo2.text = ""
Combo4.text = Data1.Recordset.RecordCount + 1
Combo1.SetFocus
End Sub
Private Sub Command2_Click()
ba.Close
Unload Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub






