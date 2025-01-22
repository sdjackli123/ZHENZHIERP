VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw15 
   BackColor       =   &H00C0E0FF&
   Caption         =   "成品库存"
   ClientHeight    =   11115
   ClientLeft      =   -435
   ClientTop       =   3810
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单号查询"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "日期查询"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "品名查询"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw15.frx":0000
      Height          =   390
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "品名"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全部库存"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw15.frx":0014
      Height          =   7575
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   9
      BackColorFixed  =   8421631
      BackColorBkg    =   50372
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
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
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw15.frx":0028
      Height          =   390
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "单号"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81592321
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81592321
      CurrentDate     =   39177
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "单号"
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
      Left            =   600
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
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
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Formw15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
       Data1.Database.Execute "DELETE * FROM CPKC"
       Data1.Database.Execute "DELETE * FROM CPKCZ"
       Data3.Database.Execute "INSERT INTO CPKC(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPFH where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data3.Database.Execute "UPDATE CPKC SET 数量=-数量 "
       Data1.Database.Execute "INSERT INTO CPKC(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPRK where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data2.Database.Execute "INSERT INTO CPKCZ(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,format(SUM(数量),'#0') FROM CPKC GROUP BY 单号,款号,品名,规格,型号,单位"
       Data2.Database.Execute "DELETE * FROM CPKCZ WHERE 数量<=0"
       Data2.RecordSource = "SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPKCZ"
       Data2.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If DBCombo1.text = "" Then
MsgBox ("请输入品名")
Exit Sub
End If
       Data1.Database.Execute "DELETE * FROM CPKC"
       Data1.Database.Execute "DELETE * FROM CPKCZ"
       Data3.Database.Execute "INSERT INTO CPKC(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPFH where 品名='" & DBCombo1.text & "'"
       Data3.Database.Execute "UPDATE CPKC SET 数量=-数量 "
       Data1.Database.Execute "INSERT INTO CPKC(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPRK where 款号<>'00000000' AND 品名='" & DBCombo1.text & "'"
       Data2.Database.Execute "INSERT INTO CPKCZ(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,format(SUM(数量),'#0') FROM CPKC GROUP BY 单号,款号,品名,规格,型号,单位"
       Data2.RecordSource = "SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPKCZ where 数量<>0"
       Data2.Refresh
End Sub

Private Sub Command5_Click()
Call OutDataToExcel(MSFlexGrid1, 6, "成品库存")
End Sub

Private Sub Command7_Click()
       Data1.Database.Execute "DELETE * FROM CPKC"
       Data1.Database.Execute "DELETE * FROM CPKCZ"
       Data3.Database.Execute "INSERT INTO CPKC(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPFH where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data3.Database.Execute "UPDATE CPKC SET 数量=-数量 "
       Data1.Database.Execute "INSERT INTO CPKC(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPRK where 款号<>'00000000' and 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data2.Database.Execute "INSERT INTO CPKCZ(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,format(SUM(数量),'#0') FROM CPKC GROUP BY 单号,款号,品名,规格,型号,单位"
       Data2.RecordSource = "SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPKCZ where 数量<>0"
       Data2.Refresh
End Sub

Private Sub Command8_Click()
If DBCombo2.text = "" Then
MsgBox ("请输入单号")
Exit Sub
End If
       Data1.Database.Execute "DELETE * FROM CPKC"
       Data1.Database.Execute "DELETE * FROM CPKCZ"
       Data3.Database.Execute "INSERT INTO CPKC(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPFH where 单号='" & DBCombo2.text & "'"
       Data3.Database.Execute "UPDATE CPKC SET 数量=-数量 "
       Data1.Database.Execute "INSERT INTO CPKC(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPRK where 单号='" & DBCombo2.text & "' AND 款号<>'00000000'"
       Data2.Database.Execute "INSERT INTO CPKCZ(单号,款号,品名,规格,型号,单位,数量) SELECT 单号,款号,品名,规格,型号,单位,format(SUM(数量),'#0') FROM CPKC GROUP BY 单号,款号,品名,规格,型号,单位"
       Data2.RecordSource = "SELECT 单号,款号,品名,规格,型号,单位,数量 FROM CPKCZ  where 数量<>0"
       Data2.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
DBCombo1.text = ""
DBCombo2.text = ""
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
Data1.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"

Data2.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"

Data3.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"

Data4.DatabaseName = "D:\数据库\htgl\2011\CPCK.MDB"
Data4.RecordSource = "SELECT 品名 FROM CPRK GROUP BY 品名"
Data4.Refresh

Data5.DatabaseName = "D:\数据库\htgl\2011\sczyjhd.mdb"
Data5.RecordSource = "select 单号  from SCZY_z group by 单号"
Data5.Refresh

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 4500
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.ColWidth(6) = 1500

End Sub
