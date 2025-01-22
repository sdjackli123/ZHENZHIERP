VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formw99 
   BackColor       =   &H00C0E0FF&
   Caption         =   "尺码信息"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   9315
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formw99.frx":0000
      Height          =   6375
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
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
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Formw99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = ""
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
MSFlexGrid2.ColWidth(0) = 200
MSFlexGrid2.ColWidth(1) = 1200
MSFlexGrid2.ColWidth(2) = 1200
End Sub

Private Sub MSFlexGrid2_dblClick()
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1


If khbl = 2 Then
For i = 0 To 4
Formw97.DBCombo1(i + 1).Text = Data2.Recordset.Fields(i)
Next
Unload Me
End If

If khbl = 1 Then
For i = 0 To 4
Formw91.DBCombo1(i + 1).Text = Data2.Recordset.Fields(i)
Next
Unload Me
End If

If khbl = 3 Then
For i = 0 To 4
Formw911.DBCombo1(i + 1).Text = Data2.Recordset.Fields(i)
Next
Unload Me
End If

If khbl = 6 Then
For i = 0 To 4
Formw96.DBCombo1(i + 1).Text = Data2.Recordset.Fields(i)
Next
Unload Me
End If

End Sub

Private Sub Text1_Change()
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "SELECT 款号,品名,颜色,尺码,单位 FROM cmxx WHERE instr(款号,'" & Text1.Text & "')>0 order BY 款号,颜色,尺码"
Data2.Refresh
End Sub


