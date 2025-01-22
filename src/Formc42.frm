VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formc42 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料入库查询"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form42"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号查询"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "材料查询"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库类查询"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc42.frx":0000
      Height          =   6975
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   8421631
      BackColorBkg    =   37265
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc42.frx":0014
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   360
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formc42.frx":0028
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81133569
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81133569
      CurrentDate     =   39177
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择库类"
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
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择款号"
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
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择材料"
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
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "Formc42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.RecordSource = "SELECT * FROM CKGL WHERE 库别='采购入库' and 材料名称='" & DBCombo1.Text & "' and 日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "') order by 库类,日期"
Data1.Refresh
End Sub

Private Sub Command2_Click()
Data1.RecordSource = "SELECT * FROM CKGL WHERE 库别='采购入库' and CKGL.库类='" & DBCombo3.Text & "' and 日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "') order by 库类,日期"
Data1.Refresh
End Sub

Private Sub Command3_Click()
Data1.RecordSource = "SELECT * FROM CKGL WHERE 库别='采购入库' and 合约号='" & DBCombo2.Text & "' and 日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "') order by 库类,日期"
Data1.Refresh
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text4.Text = DTPicker1.Value

End Sub
Private Sub DTPicker2_Change()
Text5.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text5.Text = DTPicker2.Value

End Sub

Private Sub Form_Load()
Text4.Text = Date
Text5.Text = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""

Data1.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\CKGL.mdb"
Data2.RecordSource = "select CKGL.单号,CKGL.单据号,CKGL.材料名称,CKGL.材料规格,CKGL.材料单位,CKGL.颜色,CKGL.批次,SUM(CKGL.数量-CKGL.实领量) AS 库存量 from CKGL WHERE CKGL.数量>CKGL.实领量 GROUP BY CKGL.单号,CKGL.单据号,CKGL.材料名称,CKGL.材料规格,CKGL.材料单位,CKGL.颜色,CKGL.批次"
Data2.Refresh

Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select YS.YS from YS  GROUP BY YS.YS"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\CKGL.mdb"
Data4.RecordSource = "select KL.MC from KL  GROUP BY KL.MC"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data5.RecordSource = "select CKGL.材料名称 from CKGL  GROUP BY CKGL.材料名称"
Data5.Refresh

MSFlexGrid1.ColWidth(0) = 400
MSFlexGrid1.ColWidth(1) = 0
MSFlexGrid1.ColWidth(2) = 0
MSFlexGrid1.ColWidth(12) = 0
MSFlexGrid1.ColWidth(13) = 0
MSFlexGrid1.ColWidth(14) = 0
MSFlexGrid1.ColWidth(16) = 0
MSFlexGrid1.ColWidth(17) = 0
MSFlexGrid1.ColWidth(18) = 0

MSFlexGrid1.ColWidth(25) = 0
MSFlexGrid1.ColWidth(26) = 0
MSFlexGrid1.ColWidth(27) = 0
MSFlexGrid1.ColWidth(28) = 0
MSFlexGrid1.ColWidth(29) = 0

End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 0
       khbl = 3
Form202.Text1.Text = DBCombo2.Text
Form202.Show

End Select

End Sub
