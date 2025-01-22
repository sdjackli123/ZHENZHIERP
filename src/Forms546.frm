VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms546 
   BackColor       =   &H00C0E0FF&
   Caption         =   "员工信息"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form39"
   ScaleHeight     =   9240
   ScaleWidth      =   7470
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选取"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "Forms546.frx":0000
      Top             =   120
      Width           =   5655
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1560
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2160
      Top             =   8640
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2040
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5310
      ItemData        =   "Forms546.frx":0006
      Left            =   840
      List            =   "Forms546.frx":0008
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   2160
      Width           =   4335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Forms546.frx":000A
      Left            =   3840
      List            =   "Forms546.frx":001D
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "班次"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择信息"
      Height          =   1335
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "工种"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "Forms546"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
Text1(1) = Combo1
End Sub

Private Sub Combo1_Click()
Text1(1) = Combo1
End Sub

Private Sub Command1_Click()
l1 = ""
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
l1 = l1 + Mid(List1.List(i), 1, InStr(List1.List(i), "-") - 1) + "."
End If
Next

If Text2 = "" Then
Text2 = l1
Else
Text2 = Text2 + l1
End If
End Sub

Private Sub Command2_Click()
If YGBL = 0 Then
Forms501.Text1(8).Text = Text2
End If

If YGBL = 1 Then
Forms504.Text1(8).Text = Text2
End If

If YGBL = 2 Then
Forms500.Text1(8).Text = Text2
End If

If YGBL = 3 Then
Forms502.Text1(8).Text = Text2
End If

If YGBL = 4 Then
Forms503.Text1(8).Text = Text2
End If

If YGBL = 5 Then
Forms505.Text1(8).Text = Text2
End If

If YGBL = 6 Then
Forms506.Text1(8).Text = Text2
End If

If YGBL = 7 Then
Forms507.Text1(8).Text = Text2
End If

If YGBL = 8 Then
Forms508.Text1(8).Text = Text2
End If

If YGBL = 9 Then
Forms511.Text9.Text = Text2
End If

If YGBL = 11 Then
Forms509.Text1(3).Text = Text2
End If

Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Forms_Load()

End Sub

Private Sub Command8_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Command9_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub Form_Load()
Text2 = ""
Combo1 = ""
For i = 0 To 1
Text1(i) = ""
Next
Text1(1) = ""
End Sub

Private Sub Label1_Click()
Text2 = ""
End Sub

Private Sub Label3_Click()
Combo1 = ""
Text1(1) = ""
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case Index
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 编号,姓名 FROM works1 WHERE 部门 like '%'+'" & Text1(0) & "'+'%' and 班次 like '%'+'" & Text1(1) & "'+'%'  order by 编号"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc1.Recordset.MoveFirst
List1.Clear
Do While Not Adodc1.Recordset.EOF
List1.AddItem Adodc1.Recordset.Fields(0) + "-" + Adodc1.Recordset.Fields(1)
Adodc1.Recordset.MoveNext
Loop
End Select
End Sub
