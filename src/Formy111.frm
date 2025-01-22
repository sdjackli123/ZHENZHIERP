VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formy111 
   BackColor       =   &H00C0E0FF&
   Caption         =   "尺码表"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form59"
   ScaleHeight     =   9930
   ScaleWidth      =   8925
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1320
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
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
      Top             =   2280
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
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7320
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy111.frx":0000
      Height          =   5895
      Left            =   1800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3000
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10398
      _Version        =   393216
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      FormatString    =   "记录号 "
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "尺码名称"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "Formy111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ba As Database: Dim rr As Integer
Dim rs As Single
Dim rd As Recordset


Private Sub Command5_Click()
Text3.Text = CDate(Now) + 1
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
Text1.Text = Data1.Recordset.Fields(0)
Text2.Text = Data1.Recordset.Fields(1)
End Sub



Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定修改吗?", vbYesNo) = vbNo Then
Exit Sub
End If
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
Data1.DatabaseName = "d:\数据库\\htgl\2011\sczyjhd.MDB"
Data1.RecordSource = "SELECT * from cmsz order by 序号 desc"
Data1.Refresh
Text1.Text = ""
Text2.Text = Data1.Recordset.RecordCount + 1
MSFlexGrid1.ColWidth(1) = 2600
Text1.TabIndex = 0
End Sub
Private Sub Command1_Click()
On Error Resume Next
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




