VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forma99 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ô¤¾¯¿Í»§Ñ¡Ôñ"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   7935
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.ListBox List1 
      Height          =   7830
      ItemData        =   "Forma99.frx":0000
      Left            =   1080
      List            =   "Forma99.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   840
      Width           =   4575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1680
      Top             =   8880
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
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
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "¿Í»§ÐÅÏ¢"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ë¢ÐÂ"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ñ¡Ôñ"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "È¡Ïû"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "È·¶¨"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "Forma99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.Clear
End Sub

Private Sub Label1_Click()
If khyj = 1 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ¼ò³Æ from khzl where ¼ò³Æ not in(select ¿Í»§ from yj_qfyj) and ip like '%'+'" & yhxx & "'+'%' order by ¼ò³Æ"
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
End If

If khyj = 2 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ¼ò³Æ from khzl where ¼ò³Æ not in(select ¿Í»§ from yj_pbyj) and ip like '%'+'" & yhxx & "'+'%' order by ¼ò³Æ"
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
End If

If khyj = 3 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ¼ò³Æ from khzl where ¼ò³Æ not in(select ¿Í»§ from yj_bkyj) and ip like '%'+'" & yhxx & "'+'%' order by ¼ò³Æ"
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
End If

If khyj = 4 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ¼ò³Æ from khzl where ¼ò³Æ not in(select ¿Í»§ from yj_ddyj) and ip like '%'+'" & yhxx & "'+'%' order by ¼ò³Æ"
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
End If

If khyj = 5 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ¼ò³Æ from khzl where ip like '%'+'" & yhxx & "'+'%' order by ¼ò³Æ"
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
End If

If khyj = 6 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ¼ò³Æ from khzl where ¼ò³Æ not in(select ¿Í»§ from yj_jhyj) and ip like '%'+'" & yhxx & "'+'%' order by ¼ò³Æ"
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
End If

If khyj = 7 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ¼ò³Æ from khzl where ip like '%'+'" & yhxx & "'+'%' order by ¼ò³Æ"
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
End If

End Sub

Private Sub Label3_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Label4_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub Label5_Click()
If khyj = 1 Then
Forma91.Show
Me.Hide
End If
If khyj = 2 Then
Forma94.Show
Me.Hide
End If
If khyj = 3 Then
Forma95.Show
Me.Hide
End If
If khyj = 4 Then
Forma96.Show
Me.Hide
End If
If khyj = 5 Then
Forma97.Show
Me.Hide
End If
If khyj = 6 Then
Forma98.Show
Me.Hide
End If
If khyj = 7 Then
Forma94.Show
Me.Hide
End If
End Sub

