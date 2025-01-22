VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Formy222 
   BackColor       =   &H00C0E0FF&
   Caption         =   "车间设置"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form22"
   ScaleHeight     =   7545
   ScaleWidth      =   8400
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   960
      Width           =   1575
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   5175
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy222.frx":0000
      Height          =   3855
      Left            =   1080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6800
      _Version        =   393216
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FormatString    =   "记录号 "
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "车间："
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
      Left            =   2040
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号："
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
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Formy222"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim BA As Database: Dim rr As Integer
Dim rs As Single
Dim RD As Recordset


Private Sub Command5_Click()
Text4.Text = Trim(Val(Text3.Text))
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
rs = VSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
Text1.Text = Data1.Recordset.Fields(0)
Text2.Text = Data1.Recordset.Fields(1)
End Sub
Private Sub JILU()
Dim i As Single
Data1.Refresh
RD.MoveLast
VSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To RD.RecordCount
VSFlexGrid1.TextMatrix(i, 0) = i
Next
End Sub




Private Sub Command3_Click()
On Error Resume Next
Data1.Recordset.Edit
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Update
Data1.Refresh
Text1.Text = ""
Text2.Text = Data1.Recordset.RecordCount + 1
Text1.SetFocus

End Sub

Private Sub Command4_Click()
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
Text1.Text = ""
Text2.Text = Data1.Recordset.RecordCount + 1
Text1.SetFocus

End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption + "操作年度 " + ljb
Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CLCK.MDB"
Data1.RecordSource = "SELECT * from cj"
Data1.Refresh
Text1.Text = ""
Text2.Text = Data1.Recordset.RecordCount + 1
VSFlexGrid1.ColWidth(1) = 1500
Text1.TabIndex = 0
End Sub
Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Update
Data1.Refresh
Text1.Text = ""
Text2.Text = Data1.Recordset.RecordCount + 1
Text1.SetFocus
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

