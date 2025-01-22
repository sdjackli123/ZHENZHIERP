VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formy12 
   BackColor       =   &H00C0E0FF&
   Caption         =   "颜色表"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   ScaleHeight     =   8175
   ScaleWidth      =   8970
   StartUpPosition =   2  '屏幕中心
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1560
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   5415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy12.frx":0000
      Height          =   3855
      Left            =   1680
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3480
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
      Caption         =   "请输入颜色名称："
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
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "IP："
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "Formy12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ba As Database: Dim rr As Integer
Dim rs As Single
Dim rd As Recordset


Private Sub MSFlexGrid1_dblClick()

rs = MSFlexGrid1.Row
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
Text1.Text = Data1.Recordset.Fields(0)
Text2.Text = Data1.Recordset.Fields(1)
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
If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If
Data1.Recordset.Edit
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Update
Data1.Refresh
Command3.Enabled = False
Text1.Text = ""
Text2.Text = Data1.Recordset.RecordCount + 1
Text1.SetFocus

End Sub

Private Sub Command4_Click()
If MsgBox("确定删除吗?,删除将不能恢复!", vbYesNo) = vbNo Then
Exit Sub
End If

Data1.Recordset.Delete
Data1.Refresh
Text1.Text = ""
Text2.Text = Data1.Recordset.RecordCount + 1
Text1.SetFocus

End Sub

Private Sub Form_Load()
Set ba = OpenDatabase("d:\数据库\\htgl\2011\sczyjhd.MDB")
Set rd = ba.OpenRecordset("YS", dbOpenDynaset)
Data1.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data1.RecordSource = "SELECT YS as 颜色,IP from YS ORDER BY VAL(YS.IP)"
Data1.Refresh
Text1.Text = ""
Text2.Text = Data1.Recordset.RecordCount + 1
MSFlexGrid1.ColWidth(1) = 2600
Text1.TabIndex = 0
End Sub
Private Sub Command1_Click()
rd.AddNew


rd.Fields(0) = Text1.Text
rd.Fields(1) = Text2.Text
rd.Update
Data1.Refresh
Text1.Text = ""
Text2.Text = Data1.Recordset.RecordCount + 1
Text1.SetFocus
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


