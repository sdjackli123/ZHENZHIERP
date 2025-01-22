VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Forma18 
   BackColor       =   &H00C0E0FF&
   Caption         =   "计划"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form18"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   4800
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "备活"
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Forma18.frx":0000
      Height          =   360
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "车台编号"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "调整"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   9720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8880
      Top             =   0
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      ItemData        =   "Forma18.frx":0014
      Left            =   3600
      List            =   "Forma18.frx":001E
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   3480
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Forma18.frx":002E
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   13996
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   12095216
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Forma18.frx":0042
      Height          =   360
      Left            =   1080
      TabIndex        =   8
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   360
      Left            =   3000
      TabIndex        =   15
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "车台"
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
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "总计锅数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户名称"
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
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Forma18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r, Z As Integer: Public GS As Integer: Public DH As String ''''''锅数变量\单号变量

Private Sub Command1_Click()
Z = 1
       Combo1111.Clear
       Combo1111.AddItem "备活"
       Combo1111.AddItem "取消"

End Sub

Private Sub Command2_Click()
Z = 0
       Combo1111.Clear

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
Data2.Refresh
P = 1
L = "备活"
m = "就绪"
GS = 1
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
For i = 1 To 30

If InStr(Trim(Data2.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Data2.Recordset.Fields(i)), m) <> 0 Then
    MSFlexGrid1.Row = P
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.CellForeColor = vbGreen
End If

If InStr(Trim(Data2.Recordset.Fields(i)), L) = 0 And InStr(Trim(Data2.Recordset.Fields(i)), m) <> 0 Then
    MSFlexGrid1.Row = P
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.CellForeColor = vbRed
End If

If InStr(Trim(Data2.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Data2.Recordset.Fields(i)), m) = 0 Then
    MSFlexGrid1.Row = P
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.CellForeColor = vbCyan
End If

If InStr(Trim(Data2.Recordset.Fields(i)), L) = 0 And InStr(Trim(Data2.Recordset.Fields(i)), m) = 0 Then
    MSFlexGrid1.Row = P
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.CellForeColor = vbBlack
End If

Next

For i = 1 To 30
If InStr(Trim(Data2.Recordset.Fields(i)), DBCombo2.Text) > 0 And InStr(Trim(Data2.Recordset.Fields(i)), DBCombo3.Text) > 0 Then
   Else
    MSFlexGrid1.Row = P
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.Text = ""
End If

If InStr(Trim(Data2.Recordset.Fields(i)), DBCombo2.Text) > 0 Then
GS = GS + 1
End If

Next

Data2.Recordset.MoveNext
P = P + 1
Loop
Data1.Refresh
Label3.Caption = GS - 1
End Sub

Private Sub DBCombo1_Change()
On Error Resume Next
If DBCombo1.Text = "" Then
Data2.RecordSource = "SELECT * FROM JHB  ORDER BY JHB.车台编号"
Data2.Refresh
Data1.RecordSource = "SELECT * FROM JHB  ORDER BY JHB.车台编号"
Data1.Refresh
Else
Data2.RecordSource = "SELECT * FROM JHB WHERE JHB.车台编号='" & DBCombo1.Text & "' ORDER BY JHB.车台编号"
Data2.Refresh
Data1.RecordSource = "SELECT * FROM JHB WHERE JHB.车台编号='" & DBCombo1.Text & "' ORDER BY JHB.车台编号"
Data1.Refresh
End If

End Sub

Private Sub DBCombo1_Click(Area As Integer)
On Error Resume Next
If DBCombo1.Text = "" Then
Data2.RecordSource = "SELECT * FROM JHB  ORDER BY JHB.车台编号"
Data2.Refresh
Data1.RecordSource = "SELECT * FROM JHB  ORDER BY JHB.车台编号"
Data1.Refresh
Else
Data2.RecordSource = "SELECT * FROM JHB WHERE JHB.车台编号='" & DBCombo1.Text & "' ORDER BY JHB.车台编号"
Data2.Refresh
Data1.RecordSource = "SELECT * FROM JHB WHERE JHB.车台编号='" & DBCombo1.Text & "' ORDER BY JHB.车台编号"
Data1.Refresh
End If

End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption + "操作年度： " + LJB
On Error Resume Next
Text1.Text = ""
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""
Combo1111.Text = ""
Combo1111.Visible = False
Data1.DatabaseName = "d:\数据库\bfrz\" + LJB + "\JH.MDB"
Data1.RecordSource = "SELECT * FROM JHB  ORDER BY JHB.车台编号"
Data1.Refresh
Data2.DatabaseName = "d:\数据库\bfrz\" + LJB + "\JH.MDB"
Data2.RecordSource = "SELECT * FROM JHB  ORDER BY JHB.车台编号"
Data2.Refresh
Data3.DatabaseName = "d:\数据库\bfrz\" + LJB + "\JH.MDB"
Data3.RecordSource = "SELECT ct.车台编号 FROM CT  GROUP BY CT.车台编号"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\bfrz\" + LJB + "\sczyjhd.MDB"
Data4.RecordSource = "SELECT khzl.简称 FROM khzl  GROUP BY khzl.简称"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\bfrz\" + LJB + "\JH.MDB"
Data5.Refresh



Z = 1
MSFlexGrid1.ColWidth(0) = 100
For i = 2 To 60
MSFlexGrid1.ColWidth(i) = 4000
Next
End Sub

Private Sub Label4_Click()
DBCombo1.Text = ""
End Sub

Private Sub Text1_Change()

If Text1.Text = "" Then
Data1.RecordSource = "SELECT * FROM JHB  ORDER BY JHB.车台编号"
Data1.Refresh
Data2.RecordSource = "SELECT * FROM JHB  ORDER BY JHB.车台编号"
Data2.Refresh
Else
Data1.RecordSource = "SELECT * FROM JHB WHERE JHB.车台编号='" & Text1.Text & "' ORDER BY JHB.车台编号"
Data1.Refresh
Data2.RecordSource = "SELECT * FROM JHB  WHERE JHB.车台编号='" & Text1.Text & "' ORDER BY JHB.车台编号"
Data2.Refresh
End If

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Data2.Refresh
P = 1
L = "备活"
m = "就绪"
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
For i = 1 To 30

If InStr(Trim(Data2.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Data2.Recordset.Fields(i)), m) <> 0 Then
    MSFlexGrid1.Row = P
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.CellForeColor = vbGreen
End If

If InStr(Trim(Data2.Recordset.Fields(i)), L) = 0 And InStr(Trim(Data2.Recordset.Fields(i)), m) <> 0 Then
    MSFlexGrid1.Row = P
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.CellForeColor = vbRed
End If

If InStr(Trim(Data2.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Data2.Recordset.Fields(i)), m) = 0 Then
    MSFlexGrid1.Row = P
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.CellForeColor = vbCyan
End If

If InStr(Trim(Data2.Recordset.Fields(i)), L) = 0 And InStr(Trim(Data2.Recordset.Fields(i)), m) = 0 Then
    MSFlexGrid1.Row = P
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.CellForeColor = vbBlack
End If

Next
Data2.Recordset.MoveNext
P = P + 1
Loop
End Sub
