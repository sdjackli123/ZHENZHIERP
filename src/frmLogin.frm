VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ⱦ����ҵ���ERP----��¼"
   ClientHeight    =   4515
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2667.611
   ScaleMode       =   0  'User
   ScaleWidth      =   5802.688
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   840
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Left            =   1920
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��      ��"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Password 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2325
   End
   Begin VB.TextBox UserName 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��  ��"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȷ  ��"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�������˵�"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "����(&P):"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "�û�����(&U):"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public PP As String: Public pass1 As String
Public LoginSucceeded As Boolean: Public yc  As Long
Dim c1 As New Class1
Dim expirationDate As Date
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

sql1 = "update yhb set ����='" & Password.Text & "' where �û�='" & UserName.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("�޸ĳɹ���")
End Sub


Private Sub Command2_Click()
On Error Resume Next
'On Error GoTo errhandle:

'If DISKNO <> "" Then                          'ע�����
'MDZC = c1.Md5_String_Calc(DISKNO)
'Else
'MDZC = c1.Md5_String_Calc(DISKCO)
'End If

'Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc2.RecordSource = "select * from zcb where idname='" & MDZC & "'"
'Adodc2.Refresh

'If Adodc2.Recordset.EOF Then
'FormC7.Show
'Unload Me
'Exit Sub
'End If

'sql2 = "delete from yhcd where �û�='" & yhm & "'"
'RD.Open sql2, conn, adOpenStatic, adLockOptimistic
'Command2.Enabled = False
'tcpClient.SendData sh
'errhandle:
'If Err = "40006" Then
'MsgBox ("������û�����л�������ϣ�")
'End
'End If
'End Sub


 If yc = 2 Then End
   Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
   Adodc1.RecordSource = "SELECT * FROM yhb where �û�='" & UserName.Text & "'"
   Adodc1.Refresh

    If Adodc1.Recordset.EOF Then
    MsgBox ("�û�������!")
    yc = yc + 1
    Command2.Enabled = True
    Exit Sub
    End If
   Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
   Adodc1.RecordSource = "SELECT * FROM yhb where �û�='" & UserName.Text & "'and ����='" & Password.Text & "'"
   Adodc1.Refresh

    If Adodc1.Recordset.EOF Then
    MsgBox ("�������!")
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

sql2 = "delete from yhcd where �û�='" & yhm & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Call GetHardDiskInfo
ypxx = Text1
End Sub


Private Sub Command3_Click()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from qxb where �û�='" & yhm & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox ("û���κ�Ȩ��")
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
If yhmk = "�����ƻ�" Then
Forma11.Show
End If
If yhmk = "����ɨ��" Then
Forms511.Show
End If
If yhmk = "Ⱦ�ϳ���" Then
Formr331.Show
End If
If yhmk = "��������" Then
Formr441.Show
Else
Formm1.Show
End If
End Sub

Private Sub Command4_Click()
End
End Sub
Private Sub Form_Load()
    On Error Resume Next
 
    ' ��ʼ����������
    Dim expirationDate As Date
    expirationDate = DateValue("2025-10-10")

    ' ������뵽��30�������
    Dim sevenDaysBeforeExpiration As Date
    sevenDaysBeforeExpiration = DateAdd("d", -7, expirationDate)

    ' ��鵱ǰ�����Ƿ���뵽�ڲ���7��
    If Now >= sevenDaysBeforeExpiration Then
        ' ������뵽�ڲ���7�죬��������ʱ��ʾ����
        Dim daysLeft As Integer
        daysLeft = DateDiff("d", Now, CDate(expirationDate)) ' ʹ�� CDate ������ֵת��Ϊ Date ��������
        MsgBox "������ " & daysLeft & " ����ڣ�����ϵ����̣�", vbExclamation
        
        ' ��������ʱ Timer �ؼ�
        Timer1.Interval = 1000 ' ���� Timer �ļ��Ϊ1��
        Timer1.Enabled = True
    End If

   
Dim ywj, mwj As String, s As String * 1, asciin As Integer
Dim tim, sji As Integer
Dim lbj, gk As String

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

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Call GetHardDiskInfo
Call GetDiskVolume

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
    
tcpClient.RemoteHost = "192.168.1.254"
tcpClient.RemotePort = "5000"
tcpClient.Connect
' �汾�����Զ�����
Dim currentVersion As String
currentVersion = "2.6" ' ���ذ汾��

Dim serverVersion As String
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

On Error GoTo errorhandler
http.Open "GET", "http://192.168.1.254/updates/version.txt", False
http.Send

If http.Status = 200 Then
    serverVersion = http.ResponseText
Else
    ' �޷���ȡ�������汾
    Exit Sub
End If

' ����ȡ�ķ������汾�Ƿ�Ϊ��
If Trim(serverVersion) = "" Then
    ' ������������صİ汾Ϊ��
    Exit Sub
End If

' ���汾��ת��Ϊ��ֵ���бȽ�
Dim currentVersionValue As Double
Dim serverVersionValue As Double

currentVersionValue = Val(Replace(currentVersion, ".", "")) ' ��С�����滻Ϊ�ղ�ת��Ϊ��ֵ
serverVersionValue = Val(Replace(serverVersion, ".", "")) ' ͬ��

' �ж��Ƿ���Ҫ����
If serverVersionValue > currentVersionValue Then
    ' ��⵽�°汾����ʼ���غ͸���
    
    ' ���ظ����ļ��� D:\ERP �ļ�����
    Dim updateFileURL As String
    updateFileURL = "http://192.168.1.254/updates/Ⱦ����ҵERPϵͳ.exe"
    
    Dim newFilePath As String
    newFilePath = "D:\ERP\Ⱦ����ҵERPϵͳ.exe" ' ���ر���·��

    ' ʹ�� WinHTTP �������ش��ļ���д�뱾��
    On Error GoTo DownloadError
    Call ChunkedDownloadFile(updateFileURL, newFilePath)

    ' �����滻���򣬽�����ǰ�����滻�ļ�
    Shell "�滻����.exe D:\ERP\Ⱦ����ҵERPϵͳ.exe D:\ERP\ERP", vbNormalFocus
    Unload Me
    Exit Sub
Else
    ' ��ǰ�Ѿ������°汾�������и���
    Exit Sub
End If

errorhandler:
    ' �����ȡ�汾ʱ��������
    Resume Next

DownloadError:
    ' �������ʱ��������
    Resume Next

End Sub
Private Sub ChunkedDownloadFile(url As String, savePath As String)
    On Error GoTo DownloadError

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHTTPRequest.5.1")
    
    ' �����󲢷���
    http.Open "GET", url, False
    http.Send

    ' �����Ӧ״̬
    If http.Status = 200 Then
        Dim fileNum As Integer
        fileNum = FreeFile
        
        ' ���ļ����浽ָ��·��
        Dim Buffer() As Byte
        Buffer = http.ResponseBody

        Open savePath For Binary Access Write As #fileNum
        Put #fileNum, , Buffer
        Close #fileNum
        
        ' ��֤�ļ���С
        Dim fileSize As Double ' ʹ�� Double ��֧�ָ�����ļ���С
        fileSize = CDbl(http.getResponseHeader("Content-Length"))
        If FileLen(savePath) = fileSize Then
            MsgBox "�汾������ɣ����ȷ���������", vbInformation
        Else
            MsgBox "���ص��ļ���С��ƥ�䣬�����ԡ�", vbCritical
            Kill savePath ' ɾ�����������ļ�
            Exit Sub
        End If
    Else
        MsgBox "����ʧ�ܣ�����״̬: " & http.Status, vbCritical
        Exit Sub
    End If

    Set http = Nothing
    Exit Sub

DownloadError:
    MsgBox "����ʱ��������: " & Err.Description, vbCritical
    Resume Next
End Sub
Private Sub lblLabels_Click(Index As Integer)
Select Case Index
       Case 0
pmbl = 3
Formr440.Show
       Case 1
pmbl = 4
Formr440.Show
End Select
End Sub

Private Sub Password_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Timer1_Timer()
 ' ���µ���ʱ�߼���������������µ���ʱ����ʾ����������
    
    ' ����Ƿ��ѵ���
    If Now >= expirationDate Then
        ' ����ѵ��ڣ���ʾ��Ϣ���˳�����
        MsgBox "�����ѵ��ڣ��Զ��˳���", vbExclamation
        Unload Me ' ж�ص�ǰ���壨�����Ҫ�Ļ���
        End ' �˳�����
    End If
End Sub

Private Sub UserName_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

'Private Sub tcpClient_dataArrival(ByVal bytesTotal As Long)
'On Error Resume Next
'Dim sdata As String
'tcpClient.GetData sdata


'If Mid(sdata, 1, 2) = "ok" Then
'Date = Format(Mid(sdata, 3, 10), "yyyy-mm-dd")
'sj = Right(sdata, 8)
'Time = TimeSerial(Hour(sj) Mod 24, Minute(sj), Second(sj))

 'If yc = 2 Then End
  ' Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  ' Adodc1.RecordSource = "select * from yhb where �û�='" & UserName.Text & "'"
  ' Adodc1.Refresh
  '  If Adodc1.Recordset.EOF Then
  '  MsgBox ("�û�������!")
  '  yc = yc + 1
   ' Command2.Enabled = True
  '  Exit Sub
  '  End If
  ' Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  ' Adodc1.RecordSource = "select * from yhb where �û�='" & UserName.Text & "'  and ����='" & Password.Text & "'"
  ' Adodc1.Refresh
  '  If Adodc1.Recordset.EOF Then
  '  MsgBox ("�������!")
  '  yc = yc + 1
   ' Command2.Enabled = True
   ' Exit Sub
  '  End If
'Command2.Enabled = True
'Command1.Enabled = True
'Command3.Enabled = True
'Command3.SetFocus
'yhm = UserName.Text
'If InStr(Adodc1.Recordset.Fields(1), "/") > 0 Then
'yhdm = Mid(Adodc1.Recordset.Fields(1), 1, InStr(Adodc1.Recordset.Fields(1), "/") - 1)
'bzdm = Mid(Adodc1.Recordset.Fields(1), InStr(Adodc1.Recordset.Fields(1), "/") + 1)
'Else
'yhdm = Adodc1.Recordset.Fields(1)
'End If
'yhmk = Adodc1.Recordset.Fields(4)
'yhxx = Adodc1.Recordset.Fields(5)
'yhxm = Adodc1.Recordset.Fields(6)
'tcpClient.Close
'Else
'MsgBox ("������Ч�Ŀͻ��ˣ�������ֹ")
'End
'End If

'End Sub

