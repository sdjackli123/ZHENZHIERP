VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formw1131 
   BackColor       =   &H00C0E0FF&
   Caption         =   "ƾ֤�������"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   ScaleHeight     =   7620
   ScaleWidth      =   7620
   StartUpPosition =   2  '��Ļ����
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3960
      Top             =   7080
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
         Name            =   "����"
         Size            =   9
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
      Left            =   4200
      Top             =   6840
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   4320
      Top             =   7080
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   4890
      ItemData        =   "Formw1131.frx":0000
      Left            =   960
      List            =   "Formw1131.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1800
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȫѡ"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȫ��"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formw1131.frx":0004
      Left            =   2760
      List            =   "Formw1131.frx":0014
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "����ڼ�"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ƾ֤���"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "Formw1131"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public DD, BAR, c, r As Integer: Public k1, k2 As String


Private Sub Combo1_Click()
If Combo1.Text = "ת��ƾ֤" Then
Adodc2.RecordSource = "select ƾ֤�� from CLZZPZ WHERE ���ȷ��='��'  AND (���˱�� is NULL OR ���˱��<>'��') AND ���� BETWEEN '" & k1 & "' AND '" & k2 & "' group by ƾ֤�� ORDER BY ƾ֤��"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List1.Clear
Do While Not Adodc2.Recordset.EOF
List1.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
If Combo1.Text = "����ƾ֤" Then
Adodc2.RecordSource = "select ƾ֤�� from CLFKPZ WHERE ���ȷ��='��' AND (���˱�� is NULL OR ���˱��<>'��') AND ���� BETWEEN '" & k1 & "' AND '" & k2 & "' group by ƾ֤�� ORDER BY ƾ֤��"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List1.Clear
Do While Not Adodc2.Recordset.EOF
List1.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
If Combo1.Text = "�տ�ƾ֤" Then
Adodc2.RecordSource = "select ƾ֤�� from CLSKPZ WHERE ���ȷ��='��' AND (���˱�� is NULL OR ���˱��<>'��') AND ���� BETWEEN '" & k1 & "' AND '" & k2 & "' group by ƾ֤�� ORDER BY ƾ֤��"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List1.Clear
Do While Not Adodc2.Recordset.EOF
List1.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
If Combo1.Text = "�ɱ�ƾ֤" Then
Adodc2.RecordSource = "select ƾ֤�� from CLSCCB WHERE ���ȷ��='��' AND (���˱�� is NULL OR ���˱��<>'��') AND ���� BETWEEN '" & k1 & "' AND '" & k2 & "' group by ƾ֤�� ORDER BY ƾ֤��"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List1.Clear
Do While Not Adodc2.Recordset.EOF
List1.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Formw1132.DataCombo2.Text = "" Then
MsgBox ("�����븴��Ա")
Unload Me
Exit Sub
Else
If Combo1.Text = "ת��ƾ֤" Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "UPDATE CLZZPZ SET ���ȷ��='δ',����=NULL WHERE ƾ֤��='" & Trim(List1.List(i)) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
End If
If Combo1.Text = "����ƾ֤" Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "UPDATE CLFKPZ SET ���ȷ��='δ',����=NULL WHERE ƾ֤��='" & Trim(List1.List(i)) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
End If
If Combo1.Text = "�տ�ƾ֤" Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "UPDATE CLSKPZ SET ���ȷ��='δ',����=NULL WHERE ƾ֤��='" & Trim(List1.List(i)) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
End If
If Combo1.Text = "�ɱ�ƾ֤" Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "UPDATE CLSCCB SET ���ȷ��='δ',����=NULL WHERE ƾ֤��='" & Trim(List1.List(i)) & "'"
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
Adodc1.RecordSource = "select * from rqsd where �·�='" & Text1.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Exit Sub
Else
k1 = Adodc1.Recordset.Fields(0)
k2 = Adodc1.Recordset.Fields(1)
End If
End Sub

