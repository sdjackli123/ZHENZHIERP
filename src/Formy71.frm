VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formy71 
   BackColor       =   &H00C0E0FF&
   Caption         =   "款号信息"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form42"
   ScaleHeight     =   8505
   ScaleWidth      =   8955
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
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
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy71.frx":0000
      Height          =   6375
      Left            =   600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
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
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Formy71"
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
End Sub

Private Sub MSFlexGrid2_dblClick()
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
If bzbl = 1 Then
For i = 0 To 5
Formy602.Text1(i).Text = Data2.Recordset.Fields(i)
Next
Unload Me
Formy602.Text1(6).SetFocus
End If

If bzbl = 2 Then
For i = 0 To 5
Formy602.Text1(i).Text = Data2.Recordset.Fields(i)
Next
Unload Me
Formy602.Text1(6).SetFocus
End If

End Sub

Private Sub Text1_Change()
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "SELECT SCZY_ZDH.客户,cmb.单号,cmb.款号,cmb.颜色,cmb.尺码,cmb.数量 FROM SCZY_ZDH,cmb WHERE SCZY_ZDH.单号=cmb.单号 and instr(cmb.款号,'" & Text1.Text & "')>0 order BY cmb.颜色,cmb.尺码"
Data2.Refresh
End Sub
