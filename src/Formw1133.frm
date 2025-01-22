VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formw1133 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "凭证批量复核"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7185
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3120
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3600
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
         Size            =   10.5
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
      Left            =   3720
      Top             =   6480
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      ItemData        =   "Formw1133.frx":0000
      Left            =   2520
      List            =   "Formw1133.frx":0010
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   4620
      ItemData        =   "Formw1133.frx":003C
      Left            =   720
      List            =   "Formw1133.frx":003E
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "凭证类别："
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "会计期间"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Formw1133"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public DD, BAR, c, r As Integer: Public k1, k2 As String
Private Sub Combo1_Click()
If Combo1.Text = "转账凭证" Then
Adodc2.RecordSource = "select 凭证号 from CLZZPZ WHERE (审核确认 is NULL OR 审核确认<>'已') AND 日期 BETWEEN '" & k1 & "' AND '" & k2 & "' group by 凭证号 ORDER BY 凭证号"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List1.Clear
Exit Sub
Else
Adodc2.Recordset.MoveFirst
List1.Clear
Do While Not Adodc2.Recordset.EOF
List1.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
End If

If Combo1.Text = "付款凭证" Then
Adodc2.RecordSource = "select 凭证号 from CLFKPZ WHERE (审核确认 is NULL OR 审核确认<>'已') AND 日期 BETWEEN '" & k1 & "' AND '" & k2 & "' group by 凭证号 ORDER BY 凭证号"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List1.Clear
Exit Sub
Else
Adodc2.Recordset.MoveFirst
List1.Clear
Do While Not Adodc2.Recordset.EOF
List1.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
End If

If Combo1.Text = "收款凭证" Then
Adodc2.RecordSource = "select 凭证号 from CLSKPZ WHERE (审核确认 is NULL OR 审核确认<>'已') AND 日期 BETWEEN '" & k1 & "' AND '" & k2 & "' group by 凭证号 ORDER BY 凭证号"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List1.Clear
Exit Sub
Else
Adodc2.Recordset.MoveFirst
List1.Clear
Do While Not Adodc2.Recordset.EOF
List1.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
End If

If Combo1.Text = "成本凭证" Then
Adodc2.RecordSource = "select 凭证号 from CLSCCB WHERE (审核确认 is NULL OR 审核确认<>'已') AND 日期 BETWEEN '" & k1 & "' AND '" & k2 & "' group by 凭证号 ORDER BY 凭证号"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List1.Clear
Exit Sub
Else
Adodc2.Recordset.MoveFirst
List1.Clear
Do While Not Adodc2.Recordset.EOF
List1.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
End If

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Formw1132.DataCombo2.Text = "" Then
MsgBox ("请输入复核员")
Unload Me
Exit Sub
Else
If Combo1.Text = "转账凭证" Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "UPDATE CLZZPZ SET 审核确认='已',复核='" & Formw1132.DataCombo2.Text & "' WHERE 凭证号='" & Trim(List1.List(i)) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
End If

If Combo1.Text = "付款凭证" Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "UPDATE CLFKPZ SET 审核确认='已',复核='" & Formw1132.DataCombo2.Text & "' WHERE 凭证号='" & Trim(List1.List(i)) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
End If

If Combo1.Text = "收款凭证" Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "UPDATE CLSKPZ SET 审核确认='已',复核='" & Formw1132.DataCombo2.Text & "' WHERE 凭证号='" & Trim(List1.List(i)) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
End If

If Combo1.Text = "成本凭证" Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "UPDATE CLSCCB SET 审核确认='已',复核='" & Formw1132.DataCombo2.Text & "' WHERE 凭证号='" & Trim(List1.List(i)) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
End If

End If
Unload Me
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

Text1.Text = ""
Combo1.Text = ""

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
End Sub

Private Sub Text1_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from rqsd where 月份='" & Text1.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Exit Sub
Else
k1 = Adodc1.Recordset.Fields(0)
k2 = Adodc1.Recordset.Fields(1)
End If
End Sub
