VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Forml8 
   BackColor       =   &H00C0E0FF&
   Caption         =   "工序信息"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form8"
   ScaleHeight     =   9885
   ScaleWidth      =   8820
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8880
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8400
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Data Data3 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data4 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Forml8.frx":0000
      Height          =   6855
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   12091
      _Version        =   393216
      Cols            =   7
      BackColorFixed  =   12171775
      BackColorBkg    =   45232
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      AllowUserResizing=   3
      FormatString    =   "记录号"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label18 
      BackColor       =   &H0000C0C0&
      Caption         =   " 工 序 系 数 信 息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "Forml8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Data1.DatabaseName = "d:\数据库\\htgl\2011\CW.MDB"
Data1.RecordSource = "select * from gdingxshu where 工序款号='" & Text1.Text & "' order by 工序编号"
Data1.Refresh
End Sub

Private Sub Form_Load()
Data1.DatabaseName = "d:\数据库\\htgl\2011\CW.MDB"
Text1.Text = ""
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1600
MSFlexGrid1.ColWidth(3) = 0
MSFlexGrid1.ColWidth(4) = 0
MSFlexGrid1.ColWidth(5) = 0
MSFlexGrid1.ColWidth(6) = 0
End Sub

