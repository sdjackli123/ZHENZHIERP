VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy41 
   BackColor       =   &H00C0E0FF&
   Caption         =   "大货订单"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   LinkTopic       =   "Form41"
   ScaleHeight     =   8775
   ScaleWidth      =   10695
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号刷新"
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "客户刷新"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81133569
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81133569
      CurrentDate     =   36892
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy41.frx":0000
      Height          =   360
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy41.frx":0014
      Height          =   6615
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   16
      BackColorFixed  =   8421631
      BackColorBkg    =   4109501
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   360
      Left            =   7800
      TabIndex        =   9
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
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
      Index           =   2
      Left            =   6720
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择客户"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期："
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
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期："
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
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Formy41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If DBCombo2.Text = "" Then
Data2.RecordSource = "select SCZY_ZDH.客户,SCZY_ZDH.单号,SCZY_ZDH.款式,SCZY_XDH.款号,SCZY_XDH.颜色,SCZY_XDH.数量,SCZY_ZDH.样品单号 FROM SCZY_ZDH,SCZY_XDH WHERE INSTR(SCZY_XDH.款号,'" & DBCombo2.Text & "')>0 AND SCZY_ZDH.单号=SCZY_XDH.单号 AND SCZY_ZDH.日期 BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') and INSTR(SCZY_ZDH.单号,'L')>0"
Data2.Refresh
Else
Data2.RecordSource = "select SCZY_ZDH.客户,SCZY_ZDH.单号,SCZY_ZDH.款式,SCZY_XDH.款号,SCZY_XDH.颜色,SCZY_XDH.数量,SCZY_ZDH.样品单号 FROM SCZY_ZDH,SCZY_XDH WHERE INSTR(SCZY_XDH.款号,'" & DBCombo2.Text & "')>0 AND SCZY_ZDH.单号=SCZY_XDH.单号 AND SCZY_ZDH.日期 BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') and INSTR(SCZY_ZDH.单号,'L')>0"
Data2.Refresh
End If

End Sub

Private Sub Command4_Click()
If DBCombo1.Text = "" Then
Data2.RecordSource = "select SCZY_ZDH.客户,SCZY_ZDH.单号,SCZY_ZDH.款式,SCZY_XDH.款号,SCZY_XDH.颜色,SCZY_XDH.数量,SCZY_ZDH.样品单号 FROM SCZY_ZDH,SCZY_XDH WHERE INSTR(SCZY_ZDH.单号,'L')>0 AND SCZY_ZDH.单号=SCZY_XDH.单号 AND SCZY_ZDH.日期 BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
Data2.Refresh
Else
Data2.RecordSource = "select SCZY_ZDH.客户,SCZY_ZDH.单号,SCZY_ZDH.款式,SCZY_XDH.款号,SCZY_XDH.颜色,SCZY_XDH.数量,SCZY_ZDH.样品单号 FROM SCZY_ZDH,SCZY_XDH WHERE INSTR(SCZY_ZDH.单号,'L')>0 AND SCZY_ZDH.单号=SCZY_XDH.单号 AND  SCZY_ZDH.客户='" & DBCombo1.Text & "' AND SCZY_ZDH.日期 BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
Data2.Refresh
End If
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date - 15
DTPicker2.Value = Date
DBCombo1.Text = ""
DBCombo2.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select 简称 from khZL group by 简称"
Data1.Refresh
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
MSFlexGrid1.ColWidth(0) = 400
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 1200
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.ColWidth(6) = 1200
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid1.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
If xqbl = 1 Then
Formy302.DBCombo1.Text = Data2.Recordset.Fields(1)
Unload Me
End If
If xqbl = 2 Then
Formy301.DBCombo1.Text = Data2.Recordset.Fields(1)
Unload Me
End If
If xqbl = 3 Then
Formy303.DBCombo1.Text = Data2.Recordset.Fields(1)
Unload Me
End If
If xqbl = 4 Then
Formy304.DBCombo1.Text = Data2.Recordset.Fields(1)
Unload Me
End If
If xqbl = 5 Then
Formy25.DBCombo1(0).Text = Data2.Recordset.Fields(1)
Formy25.DBCombo1(1).Text = Data2.Recordset.Fields(3)
Formy25.DBCombo1(2).Text = Data2.Recordset.Fields(4)
Unload Me
End If
If xqbl = 6 Then
Formy501.DBCombo1.Text = Data2.Recordset.Fields(1)
Unload Me
End If
If xqbl = 7 Then
Formy502.DBCombo1.Text = Data2.Recordset.Fields(1)
Unload Me
End If
If xqbl = 8 Then
Formy503.DBCombo1.Text = Data2.Recordset.Fields(1)
Unload Me
End If

End Sub
