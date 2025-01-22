VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formc58 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料信息"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10800
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc58.frx":0000
      Height          =   6375
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11245
      _Version        =   393216
      BackColorFixed  =   8421631
      BackColorBkg    =   35980
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label3 
      Caption         =   "名称查询"
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "编号查询"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "按编号"
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
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "按名称"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Formc58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
End Sub

Private Sub Label2_Click()
Data1.RecordSource = "select * from CLMC where instr(序号,'" & Text1.Text & "')>0"
Data1.Refresh
End Sub

Private Sub Label3_Click()
Data1.RecordSource = "select * from CLMC where instr(材料名称,'" & Text2.Text & "')>0"
Data1.Refresh
End Sub

Private Sub MSFlexGrid1_Click()
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data1.Recordset.Move rs - 1
Formc21.DBCombo1(3).Text = Data1.Recordset.Fields(0)
Formc21.DBCombo1(4).Text = Data1.Recordset.Fields(1)
Formc21.DBCombo1(6).Text = Data1.Recordset.Fields(3)
Formc21.DBCombo1(5).Text = Data1.Recordset.Fields(2)
Formc21.DBCombo1(15).Text = Data1.Recordset.Fields(4)
Unload Me
End Sub

Private Sub Text1_Change()
Data1.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
Data1.RecordSource = "select * from CLMC where instr(序号,'" & Text1.Text & "')>0"
Data1.Refresh
End Sub

Private Sub Text2_Change()
Data1.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
Data1.RecordSource = "select * from CLMC where instr(材料名称,'" & Text2.Text & "')>0"
Data1.Refresh
End Sub
