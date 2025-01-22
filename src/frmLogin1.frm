VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登录"
   ClientHeight    =   5580
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3296.849
   ScaleMode       =   0  'User
   ScaleWidth      =   6211.127
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   840
      Top             =   4680
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1440
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "进入主菜单"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确  定"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修  改"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox UserName 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   2325
   End
   Begin VB.TextBox Password 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2325
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退      出"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "用户名称(&U):"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "密码(&P):"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public PP As String: Public pass1 As String
Public LoginSucceeded As Boolean: Public yc  As Long
Dim c1 As New Class1
Dim MDZC As String
Private Sub Command1_Click()
Dim ywj, mwj As String, s As String * 1, asciin As Integer
mwj = ""
If passbiao = 1 Then
ywj = RTrim$(Text1.Text)
Else
ywj = RTrim$(Password.Text)
End If
L = Len(ywj)
For i = 1 To L
    s = Mid$(ywj, i, 1)
      If s >= "A" And s <= "Z" Then
      asciin = Asc(s) + 6
      If asciin > Asc("Z") Then asciin = asciin - 26
       mwj = mwj + Chr$(asciin)
      End If
      If s >= "a" And s <= "z" Then
      asciin = Asc(s) + 6
      If asciin > Asc("z") Then asciin = asciin - 26
       mwj = mwj + Chr$(asciin)
      End If
      If s < "A" Or s > "z" Or (s > "Z" And s < "a") Then
        mwj = mwj + s
      End If
      Next i
PP = mwj
If passbiao = 1 Then
Open "c:\winnt\system\posys.sys" For Output As #1
Write #1, PP
Close #1
Else
Open "c:\winnt\system\sys.sys" For Output As #1
Write #1, PP
Close #1
End If

End Sub


Private Sub Command2_Click()
On Error Resume Next

If DISKNO <> "" Then
MDZC = c1.Md5_String_Calc(DISKNO)
Else
MDZC = c1.Md5_String_Calc(DISKCO)
End If

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from zcb where idname='" & MDZC & "'"
Adodc2.Refresh

If Adodc2.Recordset.EOF Then
FormC7.Show
Unload Me
Exit Sub
End If


 If yc = 2 Then End
   
   Adodc1.RecordSource = "SELECT * FROM yhb where 用户='" & UserName.Text & "'"
   Adodc1.Refresh

    If Adodc1.Recordset.EOF Then
    MsgBox ("用户不存在!")
    yc = yc + 1
    Command2.Enabled = True
    Exit Sub
    End If
   
   Adodc1.RecordSource = "SELECT * FROM yhb where 用户='" & UserName.Text & "'and 密码='" & Password.Text & "'"
   Adodc1.Refresh

    If Adodc1.Recordset.EOF Then
    MsgBox ("密码错误！!")
    yc = yc + 1
    Command2.Enabled = True
    Exit Sub
    End If
Command2.Enabled = True
Command1.Enabled = True
Command3.Enabled = True
Command3.SetFocus
yhm = UserName.Text
yhdm = Adodc1.Recordset.Fields(1)
yhmk = Adodc1.Recordset.Fields(4)
yhxx = Adodc1.Recordset.Fields(5)

sql2 = "delete from yhcd where 用户='" & yhm & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Call GetHardDiskInfo
ypxx = Text1
End Sub

Private Sub Command3_Click()
On Error Resume Next
Adodc1.RecordSource = "select * from qxb where 用户='" & yhm & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox ("没有任何权限")
End
End If
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
If Adodc1.Recordset.Fields(3) = "Y" Then
Formm1.Label1(Val(Adodc1.Recordset.Fields(2))).Enabled = True
Else
Formm1.Label1(Val(Adodc1.Recordset.Fields(2))).Enabled = False
End If
Adodc1.Recordset.MoveNext
Loop
Unload Me
Formm1.Show
End Sub

Private Sub Command4_Click()
End
End Sub


Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then
Unload Me
End
End If

ljbl = 1

Dim ywj, mwj As String, s As String * 1, asciin As Integer
Dim tim, sji As Integer
Dim lbj, gk As String

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

yc = 0
passbiao = 0
Static dat1, dat2 As Date
Static dat3 As String
If Dir("c:\winnt\system\date.sys") = "" Then
End
End If
dat2 = Date
Open "c:\winnt\system\date.sys" For Input As #1
Input #1, dat3
Close #1

If dat3 = "" Then
dat1 = Date
dat3 = Format(CStr(dat1), "yy,mm,dd")
MsgBox (dat3), vbOKOnly

Label1.Visible = False
Text1.Visible = False

Open "c:\winnt\system\date.sys" For Output As #1
Write #1, Str(DateValue(dat3))
Close #1
End If


TM1 = CDate("2019-1-1")
TM2 = DateDiff("d", TM1, Date)
If TM2 > 720 Then
End
End If

Call GetHardDiskInfo
Call GetDiskVolume


If dat2 + 1 < CDate(dat3) Then
MsgBox ("请确定系统时间！,然后再进入！"), vbOKOnly
End
End If

dat1 = dat3

tim = DateDiff("d", dat1, dat2)
If tim >= 365 Then
dat1 = ""
dat3 = ""
gk = ""
For i = 1 To 6
sji = Fix(Rnd * 57) + 65
lbj = Chr(sji)
gk = gk + lbj
Next

If passbiao = 1 Then
Open "c:\winnt\system\posys.sys" For Output As #1
Write #1, gk
Close #1

Else
Open "c:\winnt\system\sys.sys" For Output As #1
Write #1, gk
Close #1
End If
End If

If Dir("c:\windows\inf\step.ini") = "" Then
End
End If

If Dir("c:\winnt\help\wrs.hlp") = "" Then
End
End If

If Dir("c:\winnt\system\sys.sys") = "" Then
End
End If

If Dir("c:\winnt\system\posys.sys") = "" Then
End
End If

If Dir("c:\winnt\system32\drivers\intel.ini") = "" Then
End
End If


Open "c:\winnt\system\sys.sys" For Input As #1
Input #1, pass1
Close #1

mwj = ""
ywj = RTrim$(pass1)
L = Len(ywj)
For i = 1 To L
   s = Mid$(ywj, i, 1)
      If s >= "A" And s <= "Z" Then
      asciin = Asc(s) - 6
      If asciin < Asc("A") Then asciin = asciin + 26
       mwj = mwj + Chr$(asciin)
      End If
      If s >= "a" And s <= "z" Then
      asciin = Asc(s) - 6
      If asciin < Asc("a") Then asciin = asciin + 26
       mwj = mwj + Chr$(asciin)
      End If
      If s < "A" Or s > "z" Or (s > "Z" And s < "a") Then
        mwj = mwj + s
      End If
      Next i
    pass1 = mwj
    Command1.Enabled = False
    Command3.Enabled = False
    UserName.TabIndex = 0
    

End Sub


Private Sub Password_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub UserName_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

