VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formw202 
   BackColor       =   &H00C0E0FF&
   Caption         =   "款号信息"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form42"
   ScaleHeight     =   8850
   ScaleWidth      =   8880
   StartUpPosition =   2  '屏幕中心
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Data Data2 
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
      Top             =   7800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formw202.frx":0000
      Height          =   6375
      Left            =   600
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
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Formw202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = ""
Data2.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\SCZYJHD.mdb"
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
Formw203.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 4 Then
Formw205.Text2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 5 Then
Formw502.Text2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 6 Then
Formw208.Text2.Text = Data2.Recordset.Fields(0)
Unload Me
End If

If khbl = 7 Then
Formw91.DBCombo1(1).Text = Data2.Recordset.Fields(1)
Formw91.DBCombo1(2).Text = Data2.Recordset.Fields(2)
Formw91.DBCombo1(3).Text = Data2.Recordset.Fields(3)
Formw91.DBCombo1(4).Text = Data2.Recordset.Fields(4)
Unload Me
End If

If khbl = 8 Then
Formw92.DBCombo1(1).Text = Data2.Recordset.Fields(1)
Formw92.DBCombo1(2).Text = Data2.Recordset.Fields(2)
Formw92.DBCombo1(3).Text = Data2.Recordset.Fields(3)
Formw92.DBCombo1(4).Text = Data2.Recordset.Fields(4)
Unload Me
End If

If khbl = 9 Then
Formw94.DBCombo2.Text = Data2.Recordset.Fields(1)
Formw94.DBCombo1.Text = Data2.Recordset.Fields(2)
Unload Me
End If

If khbl = 10 Then
Formw97.DBCombo1(1).Text = Data2.Recordset.Fields(1)
Formw97.DBCombo1(2).Text = Data2.Recordset.Fields(2)
Formw97.DBCombo1(3).Text = Data2.Recordset.Fields(3)
Formw97.DBCombo1(4).Text = Data2.Recordset.Fields(4)
Formw97.DBCombo1(9).Text = Data2.Recordset.Fields(0)
Unload Me
End If

If khbl = 11 Then
Formw14.DBCombo1(2).Text = Data2.Recordset.Fields(1)
Formw14.DBCombo1(3).Text = Data2.Recordset.Fields(2)
Unload Me
End If

If khbl = 12 Then
Formw307.DBCombo1.Text = Data2.Recordset.Fields(0)
Formw307.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 18 Then
Formw605.DBCombo2.Text = Data2.Recordset.Fields(1)
For i = 0 To 1
Formw605.Text1(i).Text = Data2.Recordset.Fields(i)
Next
For i = 2 To 4
Formw605.Text1(i).Text = Data2.Recordset.Fields(i + 1)
Next
Unload Me
End If

If khbl = 19 Then
Formw191.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 21 Then
Formw16.DBCombo2.Text = Data2.Recordset.Fields(0)
Formw16.DBCombo4.Text = Data2.Recordset.Fields(1)
Unload Me
End If


End Sub

Private Sub Text1_Change()
Data2.DatabaseName = "\\192.168.1.254\lbj$\" + ljb + "\SCZYJHD.mdb"
Data2.RecordSource = "SELECT distinct 单号,款号 FROM cmb WHERE instr(款号,'" & Text1.Text & "')>0 order BY 单号 desc"
Data2.Refresh
End Sub

