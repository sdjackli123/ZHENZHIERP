VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formc47 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料出入库查询"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form47"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option3 
      BackColor       =   &H008080FF&
      Caption         =   "复位"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command4 
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
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
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
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
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H008080FF&
      Caption         =   "出库"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H008080FF&
      Caption         =   "入库"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc47.frx":0000
      Height          =   330
      Left            =   2640
      TabIndex        =   7
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料名称"
      Text            =   "DBCombo1"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81395713
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81395713
      CurrentDate     =   39883
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formc47.frx":0014
      Height          =   7815
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13785
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "材料名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "Formc47"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Command4_Click()
If Option1.Value = True Then
Data1.RecordSource = "SELECT 日期,单号,合约号,供应单位,单据号,材料名称,材料规格,材料单位,颜色,批次,数量 FROM CKGL  WHERE 库别='采购入库' AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND 材料名称='" & DBCombo1.Text & "'"
Data1.Refresh
End If
If Option2.Value = True Then
Data1.RecordSource = "SELECT 日期,单号,供应单位,领料车间,单据号,材料名称,材料规格,材料单位,颜色,批次,数量 FROM KPD WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND 材料名称='" & DBCombo1.Text & "'"
Data1.Refresh
End If
End Sub

Private Sub Command6_Click()
Call OutDataToExcel(MSFlexGrid2, 13, DBCombo1.Text)
End Sub

Private Sub DTPicker1_Change()
Text1.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text1.Text = DTPicker1.Value
End Sub
Private Sub DTPicker2_Change()
Text2.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text2.Text = DTPicker2.Value
End Sub

Private Sub Form_Load()
Data1.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\CKGL.MDB"

Data3.DatabaseName = "d:\数据库\\htgl\2011\CW.MDB"



Text1.Text = Date
Text2.Text = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
DBCombo1.Text = ""
End Sub

Private Sub Option1_Click()
Data2.RecordSource = "SELECT 材料名称 FROM CKGL WHERE  日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 材料名称"
Data2.Refresh
End Sub

Private Sub Option2_Click()
Data2.RecordSource = "SELECT 材料名称 FROM kpd WHERE  日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 材料名称"
Data2.Refresh
End Sub
