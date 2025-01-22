VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormC7 
   BackColor       =   &H00C0E0FF&
   Caption         =   "软件注册  温馨提示"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10800
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "注册"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2400
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7320
      Top             =   5640
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
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "本机编码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "注册编码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "公司名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "联系方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "请正确提给信息便于我们售后提供技术支持"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "FormC7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
If Text1 = "" Then Exit Sub
If Trim(Text3) = "" Then
MsgBox ("请提供公司名称")
Exit Sub
End If

If Trim(Text4) = "" Then
MsgBox ("请提供联系方式 便于提供技术支持！")
Exit Sub
End If

If Len(Text2) <> 32 Then
MsgBox ("注册码错误")
Exit Sub
End If
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = Text2.Text
Adodc1.Recordset.Fields(1) = Text3.Text
Adodc1.Recordset.Fields(2) = Text4.Text
Adodc1.Recordset.Fields(3) = Now
Adodc1.Recordset.Update
Adodc1.Refresh
End
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * from ZCB"
Adodc1.Refresh
If DISKNO <> "" Then
Text1 = DISKNO
Else
Text1 = DISKCO
End If
Text2 = ""
Text3 = ""
Text4 = ""
End Sub



