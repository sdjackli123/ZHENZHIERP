VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "染整行业软件_ERP系统"
   ClientHeight    =   9915
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   15945
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2040
      Top             =   5040
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9660
      Width           =   15945
      _ExtentX        =   28125
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "用户"
            TextSave        =   "用户"
            Object.Tag             =   "12"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "软件信息"
            TextSave        =   "软件信息"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "现在时间"
            TextSave        =   "现在时间"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "软件版权"
            TextSave        =   "软件版权"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
            Text            =   "杭州垚凡上智信科技有限公司"
            TextSave        =   "杭州垚凡上智信科技有限公司"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":47C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":548BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5AE1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5FD0A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   1560
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu GSWZ 
      Caption         =   "公司网站"
   End
   Begin VB.Menu BZWJ 
      Caption         =   "帮助文件"
   End
   Begin VB.Menu SCZYTC 
      Caption         =   "退出系统"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rjgxjc As Integer
'Private Sub GSWZ_Click()
'Call ShellExecute(Me.hwnd, "open", "http://www.baidu.com", vbNullString, vbNullString, &H0)
'End Sub

Private Sub GSWZ_Click()
Call ShellExecute(Me.hwnd, "open", "http://www.baidu.com", vbNullString, vbNullString, &H0)
End Sub

Private Sub MDIForm_Load()
rjgxjc = 1
Call SetDeviceIndependentWindow(Me) '判断当前分辩率和设计时的分辩率是否相同
suiping = Screen.Width / Screen.TwipsPerPixelX  '计算当前的水平分辩率
cuizhi = Screen.Height / Screen.TwipsPerPixelY '计算当前的垂直分辩率
If fbl = 1 Then    '当前分辩率和设计时的分辩率不相同
Call ResizeInit(Me)    '保存原来的坐标值
Call ResizeForm(Me)    '按比例缩放
Me.Top = 0
Me.Left = 0
Me.Height = Screen.Height
Me.Width = Screen.Width
End If

tcpClient.RemoteHost = "192.168.1.254"
tcpClient.RemotePort = "5000"
tcpClient.Connect

StatusBar1.Panels(2).Text = yhm
StatusBar1.Panels(4).Text = "软件正常"
StatusBar1.Panels(8).Text = "广兴纺织有限公司"
End Sub

Private Sub SCZYTC_Click()
End
End Sub


Private Sub tcpClient_dataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim sdata As String
tcpClient.GetData sdata

StatusBar1.Panels(4).Text = 软件正常

If sdata = "软件失效" Then
End
End If

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
StatusBar1.Panels(6).Text = Now
If rjgxjc = 25552000# Then
tcpClient.SendData xtxxjm
errhandle:
If Err = "40006" Then
MsgBox ("服务器没有运行或网络故障！")
End
End If
rjgxjc = 1
End If
rjgxjc = rjgxjc + 1
End Sub

Private Sub Timer2_Timer()
If ypxx <> "Z4Z4DDYL" Then
End
End If
End Sub
