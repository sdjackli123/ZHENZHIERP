VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw502 
   BackColor       =   &H00C0E0FF&
   Caption         =   "织布入库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form42"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单位查询"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号查询"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9120
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
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号单价"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单位单价"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "单价调整"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw502.frx":0000
      Height          =   8175
      Left            =   360
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1560
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   14420
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw502.frx":0014
      Height          =   330
      Left            =   3240
      TabIndex        =   10
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81788929
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81788929
      CurrentDate     =   39177
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单位"
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
      Left            =   2640
      TabIndex        =   16
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期范围"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "单价"
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
      Index           =   3
      Left            =   9720
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Formw502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer
Private Sub Command1_Click()
Call WXCX(MSFlexGrid1, "外协查询")
End Sub

Private Sub Command2_Click()
Data2.RecordSource = "select * from ZBRK where 款号='" & Text2.text & "' and VAL(单价)=0 order by 日期"
Data2.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Data2.RecordSource = "select * from ZBRK where 款号='" & Text2.text & "' order by 日期"
Data2.Refresh
End Sub

Private Sub Command5_Click()
If DBCombo1.text = "" Then
Data2.RecordSource = "select 加工,款号,颜色,材料,幅宽,入库,单价,format(val(入库)*val(单价),'#0.00') as 金额,日期 from ZBRK WHERE 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') order by 日期"
Data2.Refresh
Else
Data2.RecordSource = "select 加工,款号,颜色,材料,幅宽,入库,单价,format(val(入库)*val(单价),'#0.00') as 金额,日期 from ZBRK where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') and  单位='" & DBCombo1.text & "'  order by 日期"
Data2.Refresh
End If
End Sub

Private Sub Command6_Click()
If DBCombo1.text = "" Then
Data2.RecordSource = "select 加工,款号,颜色,材料,幅宽,入库,单价,format(val(入库)*val(单价),'#0.00') as 金额,日期 from ZBRK WHERE 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') AND VAL(单价)=0 order by 日期"
Data2.Refresh
Else
Data2.RecordSource = "select 加工,款号,颜色,材料,幅宽,入库,单价,format(val(入库)*val(单价),'#0.00') as 金额,日期 from ZBRK where 日期 between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "') AND VAL(单价)=0 and  单位='" & DBCombo1.text & "'  order by 日期"
Data2.Refresh
End If

End Sub

Private Sub Command7_Click()
If S1 = 0 Or S2 = 0 Then
MsgBox ("请选择记录！")
Exit Sub
End If


If Text1.text = "" Then
MsgBox ("输入单价")
Exit Sub
End If


If S1 < 1 Or S2 < 1 Then
MsgBox ("选择记录")
Exit Sub
End If

If S1 > S2 Then
MsgBox ("注意选择顺序！")
Exit Sub
End If

k = S2 - S1
If k = 0 Then
Data2.Recordset.MoveFirst
rs = MSFlexGrid1.Row
Data2.Recordset.Move S1 - 1
Data2.Recordset.Edit
Data2.Recordset.Fields(11) = Text1.text
Data2.Recordset.Update
Data2.Refresh

Else

Data2.Recordset.MoveFirst
Data2.Recordset.Move S1 - 1
For L = 1 To k + 1
Data2.Recordset.Edit
Data2.Recordset.Fields(11) = Text1.text
Data2.Recordset.Update
Data2.Recordset.MoveNext
Next
End If
Data2.Refresh
End Sub

Private Sub Form_Load()
Text1.text = ""
Text2.text = ""
DBCombo1.text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Data1.DatabaseName = "D:\数据库\htgl\2011\scjd.mdb"

Data2.DatabaseName = "D:\数据库\htgl\2011\scjd.mdb"
Data3.DatabaseName = "D:\数据库\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select 简称 from WXZL group by 简称"
Data3.Refresh
MSFlexGrid1.ColWidth(0) = 300
For i = 1 To 5
MSFlexGrid1.ColWidth(i) = 1200
Next

End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
       khbl = 5
Formw202.Show
End Select
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid1.RowSel
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid1.RowSel
End Sub


