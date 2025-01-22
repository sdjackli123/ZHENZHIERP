VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms512 
   BackColor       =   &H00C0E0FF&
   Caption         =   "班组选择"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   12525
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   8520
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6810
      ItemData        =   "Forms512.frx":0000
      Left            =   6480
      List            =   "Forms512.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1080
      Width           =   4335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      ItemData        =   "Forms512.frx":0004
      Left            =   720
      List            =   "Forms512.frx":0006
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1320
      Top             =   8400
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "取消选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "信息选取"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   11040
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "编号确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5400
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "班组编号"
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
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "班组信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Forms512"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bzmc As String
Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 班组编号 FROM BZXX order by 班组编号"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc1.Recordset.MoveFirst
List1.Clear
Do While Not Adodc1.Recordset.EOF
List1.AddItem Adodc1.Recordset.Fields(0)
Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub Label1_Click()
List2.Clear
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 班组信息,班组名称 FROM BZXX where 班组编号='" & List1.List(i) & "'"
Adodc2.Refresh

bzmc = ""
bzmc = Adodc2.Recordset.Fields(1)
k = 0
For L = 0 To Len(Adodc2.Recordset.Fields(0)) / 3 - 1
List2.AddItem Mid(Adodc2.Recordset.Fields(0), k * 3 + 1, 3)
k = k + 1

If InStr(List1.List(i), "染色") > 0 Then
For m = 0 To List2.ListCount - 1
List2.Selected(m) = False
Next
Else
For m = 0 To List2.ListCount - 1
List2.Selected(m) = True
Next
End If

Next
End If
Next
End Sub

Private Sub Label2_Click()
l1 = ""
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
l1 = l1 + List2.List(i) + "."
End If
Next
If Len(bzmc) > 1 Then
Forms511.Text12 = bzmc + "/" + l1
bzgrbh = bzmc + "/" + l1
Else
Forms511.Text12 = l1
bzgrbh = l1
End If
Unload Me
End Sub

Private Sub Label3_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = False
Next
End Sub

