VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw932 
   BackColor       =   &H00C0E0FF&
   Caption         =   "扫描入库箱号信息"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   10515
   StartUpPosition =   2  '屏幕中心
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81395713
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81395713
      CurrentDate     =   39177
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw932.frx":0000
      Height          =   6255
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   9
      BackColorFixed  =   8421631
      BackColorBkg    =   50372
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Formw932"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command7_Click()
Data2.RecordSource = "select *  from LSfh where 日期 between cdate('" & DTPicker4.Value & "') and cdate('" & DTPicker5.Value & "')"
Data2.Refresh

If Not Data2.Recordset.EOF Then
Data1.RecordSource = "select 日期,发货地 as 箱号,sum(数量) as 装箱数 from LSfh where 日期 between cdate('" & DTPicker4.Value & "') and cdate('" & DTPicker5.Value & "') group by 日期,发货地"
Data1.Refresh
End If


End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Form_Load()
DTPicker4.Value = Date
DTPicker5.Value = Date
Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data2.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1800
MSFlexGrid1.ColWidth(2) = 1800
MSFlexGrid1.ColWidth(3) = 1800

End Sub

