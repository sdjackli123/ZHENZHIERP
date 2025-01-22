VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formc161 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料库存"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form16"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc161.frx":0000
      Height          =   7575
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   13361
      _Version        =   393216
      BackColorFixed  =   8421631
      BackColorBkg    =   4109501
      AllowUserResizing=   3
   End
   Begin VB.Data Data15 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox Text1111 
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库类汇总确定"
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库类查询"
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "材料查询"
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "颜色查询"
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc161.frx":0014
      Height          =   360
      Left            =   4560
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
      _ExtentX        =   4895
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
      Bindings        =   "Formc161.frx":0028
      Height          =   360
      Left            =   5520
      TabIndex        =   6
      Top             =   7320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
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
      Bindings        =   "Formc161.frx":003C
      Height          =   360
      Left            =   4560
      TabIndex        =   10
      Top             =   1080
      Width           =   2775
      _ExtentX        =   4895
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
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formc161.frx":0050
      Height          =   360
      Left            =   5520
      TabIndex        =   13
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      TabIndex        =   16
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22937601
      CurrentDate     =   39883
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
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
      Index           =   5
      Left            =   360
      TabIndex        =   19
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
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
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择库别"
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
      Index           =   3
      Left            =   4440
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
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
      Left            =   3480
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择颜色"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   7320
      Width           =   1095
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
      Left            =   3480
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "Formc161"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data2.RecordSource = "SELECT * FROM KCCXHZ WHERE  KCCXHZ.材料名称='" & DBCombo1.Text & "'"
Data2.Refresh
End Sub

Private Sub Command2_Click()
Data2.RecordSource = "SELECT * FROM KCCXHZ WHERE  KCCXHZ.材料名称='" & DBCombo1.Text & "' AND KCCXHZ.颜色='" & DBCombo2.Text & "'"
Data2.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Command4_Click()
Call MXOutDataToExcel(MSFlexGrid1, "材料库存")
End Sub

Private Sub Command5_Click()
Data2.RecordSource = "SELECT * FROM KCCXHZ WHERE  KCCXHZ.库类='" & DBCombo3.Text & "'"
Data2.Refresh
End Sub

Private Sub Command6_Click()

Data2.Database.Execute "DELETE * FROM KCCX"
Data2.Database.Execute "DELETE * FROM KCCXHZ"
Data2.Database.Execute "INSERT INTO KCCX(单号,材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库型) select 单号,材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库别 from ckgl WHERE  日期 BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
Data2.Database.Execute "UPDATE KCCX SET 库别='入库' where 库别=NULL"
'Data2.Database.Execute "INSERT INTO KCCX(单号,材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库型) select 单号,材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库别 from KPD WHERE  标签<>'库存料' and 日期 BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') "
Data2.Database.Execute "INSERT INTO KCCX(单号,材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库型) select 备注,材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类,库别 from KPD WHERE  日期 BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') "
Data2.Database.Execute "UPDATE KCCX SET 库别='出库',数量=-数量 WHERE 库别=NULL"
Data2.Database.Execute "INSERT INTO KCCXHZ(单号,库型,库类,材料名称,材料规格,材料单位,颜色,数量,单价) SELECT 单号,库型,库类,材料名称,材料规格,材料单位,颜色,format(SUM(数量),'#0.00'),单价 FROM KCCX GROUP BY 单号,库型,库类,材料名称,材料规格,材料单位,颜色,单价"
Data2.RecordSource = "SELECT * FROM KCCXHZ WHERE KCCXHZ.数量<>0 ORDER BY KCCXHZ.库类"
Data2.Refresh

End Sub

Private Sub DBCombo3_Click(Area As Integer)
Data5.RecordSource = "SELECT KCCXHZ.材料名称 FROM KCCXHZ WHERE  KCCXHZ.库类='" & DBCombo3.Text & "' GROUP BY KCCXHZ.材料名称"
Data5.Refresh

End Sub

Private Sub DTPicker3_Change()
Text3.Text = Month(DTPicker3.Value)
End Sub

Private Sub DTPicker3_CloseUp()
Text3.Text = Month(DTPicker3.Value)
End Sub

Private Sub Form_Load()
On Error Resume Next
DTPicker3 = Date
DTPicker1.Value = Date
DTPicker2 = Date
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""
DBCombo4.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data1.RecordSource = "select 材料名称 from CKGL WHERE 数量>实领量 GROUP BY 材料名称"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data2.RecordSource = "SELECT * FROM KCCXHZ WHERE KCCXHZ.数量<>0 ORDER BY KCCXHZ.库类"
Data2.Refresh

Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select YS.YS from YS  GROUP BY YS.YS"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data4.RecordSource = "select KL.MC from KL  GROUP BY KL.MC"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"

Data6.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data6.RecordSource = "select KB.MC from KB  GROUP BY KB.MC"
Data6.Refresh

MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 2600
MSFlexGrid1.ColWidth(4) = 1600
MSFlexGrid1.ColWidth(5) = 1500
MSFlexGrid1.ColWidth(8) = 0
DBCombo1.Text = ""
End Sub

