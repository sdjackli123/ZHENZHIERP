VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms545 
   BackColor       =   &H00C0E0FF&
   Caption         =   "工序选取"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form39"
   ScaleHeight     =   9000
   ScaleWidth      =   6705
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3960
      Top             =   8760
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   4200
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选取"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   7200
      ItemData        =   "Forms545.frx":0000
      Left            =   480
      List            =   "Forms545.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "工艺名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Forms545"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
l1 = ""
L2 = 0
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
l1 = l1 + Mid(List1.List(i), 1, InStr(List1.List(i), "-") - 1)
L2 = L2 + Val(Mid(List1.List(i), InStr(List1.List(i), "-") + 1))
End If
Next
If GXBL = 0 Then
Forms501.Text1(7).Text = l1
Forms501.Text1(13).Text = Format(L2, "#0.000")
End If

If GXBL = 1 Then
Forms504.Text1(7).Text = l1
Forms504.Text1(15).Text = Format(L2, "#0.000")
End If

If GXBL = 2 Then
Forms500.Text1(7).Text = l1
Forms500.Text1(13).Text = Format(L2, "#0.000")
End If

If GXBL = 3 Then
Forms502.Text1(7).Text = l1
Forms502.Text1(13).Text = Format(L2, "#0.000")
End If

If GXBL = 4 Then
Forms503.Text1(7).Text = l1
Forms503.Text1(13).Text = Format(L2, "#0.000")
End If

If GXBL = 5 Then
Forms505.Text1(7).Text = l1
Forms505.Text1(13).Text = Format(L2, "#0.000")
End If

If GXBL = 6 Then
Forms506.Text1(7).Text = l1
Forms506.Text1(13).Text = Format(L2, "#0.000")
End If

If GXBL = 7 Then
Forms507.Text1(7).Text = l1
Forms507.Text1(13).Text = Format(L2, "#0.000")
End If

If GXBL = 8 Then
Forms508.Text1(7).Text = l1
Forms508.Text1(13).Text = Format(L2, "#0.000")
End If

Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
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

Private Sub Text1_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 工序名称,工序工资系数 FROM gyshd WHERE  工序其它系数='" & Text1.Text & "' order by 序号"
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
End Sub
