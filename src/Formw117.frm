VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formw117 
   BackColor       =   &H00C0E0FF&
   Caption         =   "��Ʊ���"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   LinkTopic       =   "Form47"
   ScaleHeight     =   7755
   ScaleWidth      =   3705
   StartUpPosition =   2  '��Ļ����
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   1320
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   600
      Top             =   6720
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   360
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   480
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   331022337
      CurrentDate     =   40054
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
      ItemData        =   "Formw117.frx":0000
      Left            =   720
      List            =   "Formw117.frx":0028
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����ֽ���������ӡ"
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����������ӡ"
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����ʲ���ծ����ӡ"
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ڳ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����ڼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Formw117"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sz(10) As String

Private Sub Combo1_Click()
L = Combo1.Text + 1
Adodc2.RecordSource = "select * from RQSD where �·�='" & L & "'"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
MsgBox ("�ڼ���������")
Exit Sub
Else
DTPicker1.value = Adodc2.Recordset.Fields(0)
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open ("e:\Excel\��ӡģ��\AccessBase\KJBB.xls")
'5)���õ�1��������Ϊ�������
Excelapp.Sheets(3).Activate
Adodc1.RecordSource = "SELECT * FROM PMZJZ WHERE ����=cast('" & DTPicker1.value & "' as datetime)"
Adodc1.Refresh

If Adodc1.Recordset.EOF Then Exit Sub
Adodc3.RecordSource = "SELECT * FROM FZB"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.MoveFirst
KMZJ = 0
Do While Not Adodc3.Recordset.EOF
LZ = "+"
n = 0
If Adodc3.Recordset.Fields(2) = Null Or Trim(Adodc3.Recordset.Fields(2)) = "" Then
Else

lll = Adodc3.Recordset.Fields(2)
'MsgBox (LLL)
For i = 1 To Len(lll)

ll = Mid(lll, i, 1)
'MsgBox (ll)
If ll = "-" Or ll = "+" Then '''''''''''''''''
sz(n) = LZ
LZ = ll
n = n + 1
Else
If i = Len(lll) Then
LZ = LZ + ll
sz(n) = LZ
Else
LZ = LZ + ll
End If
End If
''''''''''''''''''''''''''''''''''''''''''
Next
End If

For i = 0 To 10
If sz(i) <> "" Then
 W = Mid(sz(i), 2, Len(sz(i)) - 3)
 If Mid(sz(i), 1, 1) = "+" Then
 BJ1 = 2
 End If
 If Mid(sz(i), 1, 1) = "-" Then
 BJ1 = 1
 End If
 JDY = Right(sz(i), 1)
Adodc1.RecordSource = "SELECT * FROM CWMC WHERE ��ƿ�Ŀ='" & W & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Else
Select Case JDY
       Case "��"
If BJ1 = 2 Then
KMZJ = KMZJ + Val(Adodc1.Recordset.Fields(4))
Else
KMZJ = KMZJ - Val(Adodc1.Recordset.Fields(4))
End If
       Case "��"
If BJ1 = 2 Then
KMZJ = KMZJ + Val(Adodc1.Recordset.Fields(5))
Else
KMZJ = KMZJ - Val(Adodc1.Recordset.Fields(5))
End If
       Case "��"
If BJ1 = 2 Then
KMZJ = KMZJ + Val(Adodc1.Recordset.Fields(7))
Else
KMZJ = KMZJ - Val(Adodc1.Recordset.Fields(7))
End If
End Select
End If
End If
Next

If KMZJ <> 0 Then
Excelapp.ActiveSheet.Range(Adodc3.Recordset.Fields(4)) = KMZJ
End If
KMZJ = 0
For i = 0 To 9
sz(i) = ""
Next

Adodc3.Recordset.MoveNext
Loop

Excelapp.ActiveWindow.Zoom = 100
Excelapp.Visible = True
'MsgBox ("")
        Excelapp.DisplayAlerts = False
       ' Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub

Private Sub Command2_Click()
On Error Resume Next
        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open ("e:\Excel\��ӡģ��\AccessBase\KJBB.xls")
'5)���õ�1��������Ϊ�������
Excelapp.Sheets(2).Activate
Adodc1.RecordSource = "SELECT * FROM PMZJZ WHERE ����=cast('" & DTPicker1.value & "' as datetime)"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then Exit Sub
Adodc3.RecordSource = "SELECT * FROM LRB"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.MoveFirst
KMZJ = 0
Do While Not Adodc3.Recordset.EOF
LZ = "+"
n = 0
If Adodc3.Recordset.Fields(2) = Null Or Trim(Adodc3.Recordset.Fields(2)) = "" Then
Else

lll = Adodc3.Recordset.Fields(2)
'MsgBox (LLL)
For i = 1 To Len(lll)

ll = Mid(lll, i, 1)
'MsgBox (ll)
If ll = "-" Or ll = "+" Then '''''''''''''''''
sz(n) = LZ
LZ = ll
n = n + 1
Else
If i = Len(lll) Then
LZ = LZ + ll
sz(n) = LZ
Else
LZ = LZ + ll
End If
End If
''''''''''''''''''''''''''''''''''''''''''
Next
End If

For i = 0 To 10
If sz(i) <> "" Then
 W = Mid(sz(i), 2, Len(sz(i)) - 3)
 If Mid(sz(i), 1, 1) = "+" Then
 BJ1 = 2
 End If
 If Mid(sz(i), 1, 1) = "-" Then
 BJ1 = 1
 End If
 JDY = Right(sz(i), 1)
Adodc1.RecordSource = "SELECT * FROM CWMC WHERE  ��ƿ�Ŀ='" & W & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Else
Select Case JDY
       Case "��"
If BJ1 = 2 Then
KMZJ = KMZJ + Val(Adodc1.Recordset.Fields(4))
Else
KMZJ = KMZJ - Val(Adodc1.Recordset.Fields(4))
End If
       Case "��"
If BJ1 = 2 Then
KMZJ = KMZJ + Val(Adodc1.Recordset.Fields(5))
Else
KMZJ = KMZJ - Val(Adodc1.Recordset.Fields(5))
End If
       Case "��"
If BJ1 = 2 Then
KMZJ = KMZJ + Val(Adodc1.Recordset.Fields(7))
Else
KMZJ = KMZJ - Val(Adodc1.Recordset.Fields(7))
End If
End Select
End If
End If
Next

If KMZJ <> 0 Then
Excelapp.ActiveSheet.Range(Adodc3.Recordset.Fields(4)) = KMZJ
End If
KMZJ = 0
For i = 0 To 9
sz(i) = ""
Next

Adodc3.Recordset.MoveNext
Loop
Excelapp.ActiveWindow.Zoom = 100
Excelapp.Visible = True
'MsgBox ("")
        Excelapp.DisplayAlerts = False
       ' Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub

Private Sub Command3_Click()
On Error Resume Next
        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open ("e:\Excel\��ӡģ��\AccessBase\KJBB.xls")
'5)���õ�1��������Ϊ�������
Excelapp.Sheets(1).Activate
Adodc1.RecordSource = "SELECT * FROM PMZJZ WHERE ����=cast('" & DTPicker1.value & "' as datetime)"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then Exit Sub
Adodc3.RecordSource = "SELECT * FROM XJB"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.MoveFirst
KMZJ = 0
Do While Not Adodc3.Recordset.EOF
LZ = "+"
n = 0
If Adodc3.Recordset.Fields(2) = Null Or Trim(Adodc3.Recordset.Fields(2)) = "" Then
Else

lll = Adodc3.Recordset.Fields(2)
'MsgBox (LLL)
For i = 1 To Len(lll)

ll = Mid(lll, i, 1)
'MsgBox (ll)
If ll = "-" Or ll = "+" Then '''''''''''''''''
sz(n) = LZ
LZ = ll
n = n + 1
Else
If i = Len(lll) Then
LZ = LZ + ll
sz(n) = LZ
Else
LZ = LZ + ll
End If
End If
''''''''''''''''''''''''''''''''''''''''''
Next
End If

For i = 0 To 10
If sz(i) <> "" Then
 W = Mid(sz(i), 2, Len(sz(i)) - 3)
 If Mid(sz(i), 1, 1) = "+" Then
 BJ1 = 2
 End If
 If Mid(sz(i), 1, 1) = "-" Then
 BJ1 = 1
 End If
 JDY = Right(sz(i), 1)
Adodc1.RecordSource = "SELECT * FROM CWMC WHERE  ��ƿ�Ŀ='" & W & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Else
Select Case JDY
       Case "��"
If BJ1 = 2 Then
KMZJ = KMZJ + Val(Adodc1.Recordset.Fields(4))
Else
KMZJ = KMZJ - Val(Adodc1.Recordset.Fields(4))
End If
       Case "��"
If BJ1 = 2 Then
KMZJ = KMZJ + Val(Adodc1.Recordset.Fields(5))
Else
KMZJ = KMZJ - Val(Adodc1.Recordset.Fields(5))
End If
       Case "��"
If BJ1 = 2 Then
KMZJ = KMZJ + Val(Adodc1.Recordset.Fields(7))
Else
KMZJ = KMZJ - Val(Adodc1.Recordset.Fields(7))
End If
End Select
End If
End If
Next

If KMZJ <> 0 Then
Excelapp.ActiveSheet.Range(Adodc3.Recordset.Fields(4)) = KMZJ
End If
KMZJ = 0
For i = 0 To 9
sz(i) = ""
Next

Adodc3.Recordset.MoveNext
Loop
Excelapp.ActiveWindow.Zoom = 100
Excelapp.Visible = True
'MsgBox ("")
        Excelapp.DisplayAlerts = False
       ' Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub

Private Sub Form_Load()

Combo1.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT * FROM CWMC"
Adodc4.Refresh
End Sub
