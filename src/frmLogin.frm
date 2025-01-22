VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "染整行业软件ERP----登录"
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
   StartUpPosition =   2  '屏幕中心
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退      出"
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
      Caption         =   "修  改"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确  定"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "进入主菜单"
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
      Caption         =   "密码(&P):"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "用户名称(&U):"
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

sql1 = "update yhb set 密码='" & Password.Text & "' where 用户='" & UserName.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("修改成功！")
End Sub


Private Sub Command2_Click()
On Error Resume Next
'On Error GoTo errhandle:

'If DISKNO <> "" Then                          '注册界面
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

'sql2 = "delete from yhcd where 用户='" & yhm & "'"
'RD.Open sql2, conn, adOpenStatic, adLockOptimistic
'Command2.Enabled = False
'tcpClient.SendData sh
'errhandle:
'If Err = "40006" Then
'MsgBox ("服务器没有运行或网络故障！")
'End
'End If
'End Sub


 If yc = 2 Then End
   Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
   Adodc1.RecordSource = "SELECT * FROM yhb where 用户='" & UserName.Text & "'"
   Adodc1.Refresh

    If Adodc1.Recordset.EOF Then
    MsgBox ("用户不存在!")
    yc = yc + 1
    Command2.Enabled = True
    Exit Sub
    End If
   Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
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
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
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
If yhmk = "生产计划" Then
Forma11.Show
End If
If yhmk = "触摸扫描" Then
Forms511.Show
End If
If yhmk = "染料称量" Then
Formr331.Show
End If
If yhmk = "助剂称量" Then
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
 
    ' 初始化到期日期
    Dim expirationDate As Date
    expirationDate = DateValue("2025-10-10")

    ' 计算距离到期30天的日期
    Dim sevenDaysBeforeExpiration As Date
    sevenDaysBeforeExpiration = DateAdd("d", -7, expirationDate)

    ' 检查当前日期是否距离到期不足7天
    If Now >= sevenDaysBeforeExpiration Then
        ' 如果距离到期不足7天，弹出倒计时提示窗口
        Dim daysLeft As Integer
        daysLeft = DateDiff("d", Now, CDate(expirationDate)) ' 使用 CDate 将日期值转换为 Date 数据类型
        MsgBox "程序将在 " & daysLeft & " 天后到期，请联系软件商！", vbExclamation
        
        ' 启动倒计时 Timer 控件
        Timer1.Interval = 1000 ' 设置 Timer 的间隔为1秒
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
' 版本检查和自动更新
Dim currentVersion As String
currentVersion = "2.6" ' 本地版本号

Dim serverVersion As String
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

On Error GoTo errorhandler
http.Open "GET", "http://192.168.1.254/updates/version.txt", False
http.Send

If http.Status = 200 Then
    serverVersion = http.ResponseText
Else
    ' 无法获取服务器版本
    Exit Sub
End If

' 检查获取的服务器版本是否为空
If Trim(serverVersion) = "" Then
    ' 如果服务器返回的版本为空
    Exit Sub
End If

' 将版本号转换为数值进行比较
Dim currentVersionValue As Double
Dim serverVersionValue As Double

currentVersionValue = Val(Replace(currentVersion, ".", "")) ' 将小数点替换为空并转换为数值
serverVersionValue = Val(Replace(serverVersion, ".", "")) ' 同上

' 判断是否需要更新
If serverVersionValue > currentVersionValue Then
    ' 检测到新版本，开始下载和更新
    
    ' 下载更新文件到 D:\ERP 文件夹中
    Dim updateFileURL As String
    updateFileURL = "http://192.168.1.254/updates/染整行业ERP系统.exe"
    
    Dim newFilePath As String
    newFilePath = "D:\ERP\染整行业ERP系统.exe" ' 本地保存路径

    ' 使用 WinHTTP 对象下载大文件并写入本地
    On Error GoTo DownloadError
    Call ChunkedDownloadFile(updateFileURL, newFilePath)

    ' 调用替换程序，结束当前程序并替换文件
    Shell "替换程序.exe D:\ERP\染整行业ERP系统.exe D:\ERP\ERP", vbNormalFocus
    Unload Me
    Exit Sub
Else
    ' 当前已经是最新版本，不进行更新
    Exit Sub
End If

errorhandler:
    ' 如果获取版本时发生错误
    Resume Next

DownloadError:
    ' 如果下载时发生错误
    Resume Next

End Sub
Private Sub ChunkedDownloadFile(url As String, savePath As String)
    On Error GoTo DownloadError

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHTTPRequest.5.1")
    
    ' 打开请求并发送
    http.Open "GET", url, False
    http.Send

    ' 检查响应状态
    If http.Status = 200 Then
        Dim fileNum As Integer
        fileNum = FreeFile
        
        ' 将文件保存到指定路径
        Dim Buffer() As Byte
        Buffer = http.ResponseBody

        Open savePath For Binary Access Write As #fileNum
        Put #fileNum, , Buffer
        Close #fileNum
        
        ' 验证文件大小
        Dim fileSize As Double ' 使用 Double 以支持更大的文件大小
        fileSize = CDbl(http.getResponseHeader("Content-Length"))
        If FileLen(savePath) = fileSize Then
            MsgBox "版本更新完成，点击确定进入程序。", vbInformation
        Else
            MsgBox "下载的文件大小不匹配，请重试。", vbCritical
            Kill savePath ' 删除不完整的文件
            Exit Sub
        End If
    Else
        MsgBox "下载失败，错误状态: " & http.Status, vbCritical
        Exit Sub
    End If

    Set http = Nothing
    Exit Sub

DownloadError:
    MsgBox "下载时发生错误: " & Err.Description, vbCritical
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
 ' 更新倒计时逻辑，可以在这里更新倒计时的显示或其他操作
    
    ' 检查是否已到期
    If Now >= expirationDate Then
        ' 如果已到期，显示消息并退出程序
        MsgBox "程序已到期，自动退出！", vbExclamation
        Unload Me ' 卸载当前窗体（如果需要的话）
        End ' 退出程序
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
  ' Adodc1.RecordSource = "select * from yhb where 用户='" & UserName.Text & "'"
  ' Adodc1.Refresh
  '  If Adodc1.Recordset.EOF Then
  '  MsgBox ("用户不存在!")
  '  yc = yc + 1
   ' Command2.Enabled = True
  '  Exit Sub
  '  End If
  ' Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
  ' Adodc1.RecordSource = "select * from yhb where 用户='" & UserName.Text & "'  and 密码='" & Password.Text & "'"
  ' Adodc1.Refresh
  '  If Adodc1.Recordset.EOF Then
  '  MsgBox ("密码错误！!")
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
'MsgBox ("不是有效的客户端，程序终止")
'End
'End If

'End Sub

