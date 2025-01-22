VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Formw338 
   BackColor       =   &H00C0E0FF&
   Caption         =   "明细账"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form38"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   1200
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data14 
      Caption         =   "Data6"
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
      Top             =   4320
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细打印"
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "期末结转"
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "余额打印"
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按明细汇总"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "总明细账"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw338.frx":0000
      Height          =   330
      Left            =   8640
      TabIndex        =   11
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "LBJ"
      Text            =   "DBCombo1"
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "明细账"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按总账汇总"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formw338.frx":0014
      Left            =   2400
      List            =   "Formw338.frx":001E
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw338.frx":0036
      Height          =   7815
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13785
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   8421631
      BackColorBkg    =   34952
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   84017153
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   84017153
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   84017153
      CurrentDate     =   39883
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "上期结转日"
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
      Left            =   2400
      TabIndex        =   13
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "明细科目"
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
      Left            =   8640
      TabIndex        =   10
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
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
      Left            =   360
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作月份"
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
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "总账科目"
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
      Index           =   3
      Left            =   2400
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Formw338"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2, BT As String

Private Sub Combo1_Click()
Data5.RecordSource = "SELECT RIGHT(MXFLZ.会计科目,LEN(MXFLZ.会计科目)-INSTR(MXFLZ.会计科目,'-')) as lbj FROM MXFLZ WHERE INSTR(MXFLZ.会计科目,'" & Combo1.Text & "')>0  GROUP BY MXFLZ.会计科目"
Data5.Refresh
End Sub

Private Sub Command1_Click()
Data1.RecordSource = "SELECT * FROM MXFLZ WHERE INSTR(MXFLZ.会计科目,'" & Combo1.Text & "')>0 AND MXFLZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') ORDER BY MXFLZ.日期,MXFLZ.凭证号"
Data1.Refresh
End Sub

Private Sub Command2_Click()
On Error Resume Next
Data1.Database.Execute "DELETE * FROM ZLCX"
Data1.Database.Execute "INSERT INTO ZLCX SELECT * FROM MXFLZ WHERE INSTR(MXFLZ.会计科目,'" & Combo1.Text & "')>0 AND MXFLZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "')"
Data1.Database.Execute "UPDATE ZLCX set 序号='2'"
Data1.Database.Execute "INSERT INTO ZLCX(会计科目,借方金额,贷方金额) SELECT MXFLZ.会计科目,SUM(VAL(MXFLZ.借方金额)),SUM(VAL(MXFLZ.贷方金额)) FROM MXFLZ WHERE INSTR(MXFLZ.会计科目,'" & Combo1.Text & "')>0 AND MXFLZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') GROUP BY MXFLZ.会计科目"
Data1.Database.Execute "UPDATE ZLCX set 摘要='本月合计',序号='3' WHERE 摘要=NULL"
Data1.Database.Execute "UPDATE ZLCX set 借方金额=format(借方金额,'#0.00'),贷方金额=format(贷方金额,'#0.00')"

Data3.RecordSource = "ZLCX"
Data3.Refresh

Data1.Database.Execute "INSERT INTO ZLCX(日期,凭证号,摘要,会计科目,借贷方向,余额,类别) SELECT 日期,凭证号,摘要,会计科目,借贷方向,余额,类别 FROM PMMXJZ WHERE INSTR(PMMXJZ.会计科目,'" & Combo1.Text & "')>0 AND PMMXJZ.日期=CDATE('" & DTPicker1.Value & "') "
Data1.Database.Execute "UPDATE ZLCX set 序号='1',摘要='期初余额' WHERE 摘要='期初余额'"

Data2.RecordSource = "SELECT ZLCX.会计科目 FROM ZLCX GROUP BY ZLCX.会计科目"
Data2.Refresh

If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
L = Right(Data2.Recordset.Fields(0), Len(Data2.Recordset.Fields(0)) - InStr(Data2.Recordset.Fields(0), "-"))
Data4.Recordset.FindFirst "科目名称='" & L & "'"
If Data4.Recordset.NoMatch Then
MsgBox (L + "科目设置有错")
Exit Sub
End If

Data3.RecordSource = "SELECT * FROM ZLCX WHERE INSTR(ZLCX.会计科目,'" & L & "')>0"
Data3.Refresh
If Data3.Recordset.EOF Then
MsgBox ("无记录")
Exit Sub
End If
Data3.Recordset.FindFirst "序号='1'"
If Data3.Recordset.NoMatch Then
Data3.RecordSource = "SELECT * FROM ZLCX WHERE ZLCX.摘要='本月合计' AND ZLCX.会计科目='" & Data2.Recordset.Fields(0) & "'"
Data3.Refresh
Data3.Recordset.Edit
If Data4.Recordset.Fields(3) = "贷" Then
Data3.Recordset.Fields(7) = Format(Format(Val(Data3.Recordset.Fields(5)) - Val(Data3.Recordset.Fields(4)), "#0.00"), "#0.00")
Data3.Recordset.Fields(6) = "贷"
Else
Data3.Recordset.Fields(7) = Format(Format(Val(Data3.Recordset.Fields(4)) - Val(Data3.Recordset.Fields(5)), "#0.00"), "#0.00")
Data3.Recordset.Fields(6) = "借"
End If
Data3.Recordset.Update

Else

k = Format(Val(Data3.Recordset.Fields(7)), "#0.00")
Data3.RecordSource = "SELECT * FROM ZLCX WHERE ZLCX.摘要='本月合计' AND ZLCX.会计科目='" & Data2.Recordset.Fields(0) & "'"
Data3.Refresh
Data3.Recordset.Edit
If Data4.Recordset.Fields(3) = "贷" Then
Data3.Recordset.Fields(7) = Format(Format(Val(Data3.Recordset.Fields(5)) - Val(Data3.Recordset.Fields(4)) + k, "#0.00"), "#0.00")
Data3.Recordset.Fields(6) = "贷"
Else
Data3.Recordset.Fields(7) = Format(Format(Val(Data3.Recordset.Fields(4)) - Val(Data3.Recordset.Fields(5)) + k, "#0.00"), "#0.00")
Data3.Recordset.Fields(6) = "借"
End If
Data3.Recordset.Update
End If
Data2.Recordset.MoveNext
Loop


Data2.RecordSource = "SELECT * FROM ZLCX WHERE ZLCX.序号='1'"
Data2.Refresh

Data3.RecordSource = "SELECT * FROM ZLCX "
Data3.Refresh

Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Data3.Recordset.FindFirst "序号='2' AND 会计科目='" & Data2.Recordset.Fields(3) & "'"
If Data3.Recordset.NoMatch Then
Data3.Database.Execute "INSERT INTO ZLCX(日期,摘要,会计科目,借方金额,贷方金额,借贷方向,余额,序号) VALUES('" & DTPicker3.Value & "','本月合计','" & Data2.Recordset.Fields(3) & "','" & Data2.Recordset.Fields(4) & "','" & Data2.Recordset.Fields(5) & "','" & Data2.Recordset.Fields(6) & "','" & Data2.Recordset.Fields(7) & "','3')"
End If
Data2.Recordset.MoveNext
Loop


Data1.Database.Execute "UPDATE ZLCX SET 凭证号='结-'+'" & Text3.Text & "' WHERE 凭证号=NULL"
Data1.RecordSource = "SELECT * FROM ZLCX ORDER BY ZLCX.会计科目,VAL(ZLCX.序号),ZLCX.日期"
Data1.Refresh
BT = "按明细账汇总"
End Sub

Private Sub Command3_Click()
If Combo1.Text = "应付账款" Then
Call YEBDOutDataToExcelSZ(Data2, Data3, Text3.Text)
End If
If Combo1.Text = "应收账款" Then
Call SYEBDOutDataToExcelSZ(Data2, Data3, Text3.Text)
End If
If Combo1.Text <> "应付账款" And Combo1.Text <> "应收账款" Then
Call QYEBDOutDataToExcelSZ(Data2, Data3, Text3.Text)
End If
End Sub

Private Sub Command4_Click()
Data1.RecordSource = "SELECT * FROM MXFLZ WHERE INSTR(MXFLZ.会计科目,'" & Combo1.Text & "')>0 AND INSTR(MXFLZ.会计科目,'" & DBCombo1.Text & "')>0 AND MXFLZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') ORDER BY MXFLZ.日期,MXFLZ.凭证号"
Data1.Refresh
End Sub


Private Sub Command5_Click()
If MsgBox("确定本期期末余额结转吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("本期为: " + Text3.Text + " 期间" + "，正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("结转为次月的: " + Str(DTPicker2.Value) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
Data3.RecordSource = "SELECT * FROM PMMXJZ WHERE 日期=CDATE('" & DTPicker2.Value & "') AND INSTR(会计科目,'" & Combo1.Text & "')>0"
Data3.Refresh
If Not Data3.Recordset.EOF Then
If MsgBox("已有此时间内的记录，即已结转过，覆盖原先记录吗？", vbYesNo) = vbNo Then
Exit Sub
Else
Data1.Database.Execute "DELETE * FROM PMMXJZ WHERE 日期=CDATE('" & DTPicker2.Value & "')  AND INSTR(会计科目,'" & Combo1.Text & "')>0"
End If
End If

Data1.Database.Execute "UPDATE PMMXJZ SET 余额=format(余额,'#0.00') where 日期=CDATE('" & DTPicker2.Value & "')"
Data1.Database.Execute "INSERT INTO PMMXJZ(凭证号,摘要,会计科目,借方金额,贷方金额,借贷方向,余额,类别,序号) SELECT 凭证号,摘要,会计科目,借方金额,贷方金额,借贷方向,余额,类别,序号 FROM ZLCX WHERE 凭证号='结-'+'" & Text3.Text & "'"
Data1.Database.Execute "UPDATE PMMXJZ SET 摘要='期初余额',日期=CDATE('" & DTPicker2.Value & "') WHERE 日期=null"
MsgBox ("结转成功！")

End Sub

Private Sub Command6_Click()
On Error Resume Next
If Combo1.Text = "" Then
MsgBox ("请选择总账科目")
Exit Sub
End If
Data1.Database.Execute "DELETE * FROM ZLCX"
Data1.Database.Execute "INSERT INTO ZLCX(会计科目,借方金额,贷方金额) SELECT MXFLZ.会计科目,format(SUM(VAL(MXFLZ.借方金额)),'#0.00'),format(SUM(VAL(MXFLZ.贷方金额)),'#0.00') FROM MXFLZ WHERE INSTR(MXFLZ.会计科目,'" & Combo1.Text & "')>0 AND MXFLZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') GROUP BY MXFLZ.会计科目"
Data3.RecordSource = "ZLCX"
Data3.Refresh
If Data3.Recordset.EOF Then Exit Sub
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
L = Right(Data3.Recordset.Fields(3), Len(Data3.Recordset.Fields(3)) - InStr(Data3.Recordset.Fields(3), "-"))
m = Left(Data3.Recordset.Fields(3), InStr(Data3.Recordset.Fields(3), "-") - 1)

Data4.Recordset.FindFirst "科目名称='" & m & "'"
If Data4.Recordset.NoMatch Then
MsgBox (m + "科目设置有错")
Exit Sub
Else
n = Data4.Recordset.Fields(2)
End If


Data4.Recordset.FindFirst "科目名称='" & L & "' AND 科目类型='" & n & "'"
If Data4.Recordset.NoMatch Then
MsgBox (m + L + "科目设置有错或类型有误")
Exit Sub
Else
If Data4.Recordset.Fields(3) = "借" Then
Data3.Recordset.Edit
Data3.Recordset.Fields(0) = K2
Data3.Recordset.Fields(1) = "汇总"
Data3.Recordset.Fields(2) = "本期发生额"
Data3.Recordset.Fields(6) = "借"
Data3.Recordset.Fields(9) = "2"
Data3.Recordset.Update
End If
If Data4.Recordset.Fields(3) = "贷" Then
Data3.Recordset.Edit
Data3.Recordset.Fields(0) = K2
Data3.Recordset.Fields(1) = "汇总"
Data3.Recordset.Fields(2) = "本期发生额"
Data3.Recordset.Fields(6) = "贷"
Data3.Recordset.Fields(9) = "2"
Data3.Recordset.Update
End If
End If
Data3.Recordset.MoveNext
Loop
Data1.Database.Execute "INSERT INTO ZLCX(日期,凭证号,会计科目,借贷方向,余额,类别) SELECT 日期,凭证号,会计科目,借贷方向,余额,类别 FROM PMMXJZ WHERE INSTR(PMMXJZ.会计科目,'" & Combo1.Text & "')>0 AND PMMXJZ.日期=CDATE('" & DTPicker1.Value & "') "
Data1.Database.Execute "UPDATE ZLCX set 序号='1',摘要='期初余额' WHERE 摘要=NULL"
''''''''''''''''''''''''''''''''''''''''''''''''
Data3.RecordSource = "SELECT * FROM ZLCX WHERE 序号='1'"
Data3.Refresh
If Not Data3.Recordset.EOF Then
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
L = Right(Data3.Recordset.Fields(3), Len(Data3.Recordset.Fields(3)) - InStr(Data3.Recordset.Fields(3), "-"))
m = Left(Data3.Recordset.Fields(3), InStr(Data3.Recordset.Fields(3), "-") - 1)

Data4.Recordset.FindFirst "科目名称='" & m & "'"
If Data4.Recordset.NoMatch Then
MsgBox (m + "科目设置有错")
Exit Sub
Else
n = Data4.Recordset.Fields(2)
End If


Data4.Recordset.FindFirst "科目名称='" & L & "' AND 科目类型='" & n & "'"
If Data4.Recordset.NoMatch Then
MsgBox (m + L + "科目设置有错或类型有误")
Exit Sub
End If
Data3.Recordset.MoveNext
Loop
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Data2.RecordSource = "SELECT ZLCX.会计科目 FROM ZLCX GROUP BY ZLCX.会计科目"
Data2.Refresh

If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Data3.RecordSource = "SELECT * FROM ZLCX WHERE ZLCX.会计科目='" & Data2.Recordset.Fields(0) & "' ORDER BY VAL(ZLCX.序号)"
Data3.Refresh
Data3.Recordset.FindFirst "序号='1'"
If Data3.Recordset.NoMatch Then
M3 = Data3.Recordset.Fields(3)
M4 = Val(Data3.Recordset.Fields(4))
M5 = Val(Data3.Recordset.Fields(5))
M6 = Data3.Recordset.Fields(6)

Data3.Recordset.Edit
If Data3.Recordset.Fields(6) = "贷" Then
Data3.Recordset.Fields(7) = Format(Val(Data3.Recordset.Fields(5)) - Val(Data3.Recordset.Fields(4)), "#0.00")
M7 = Data3.Recordset.Fields(7)
Else
Data3.Recordset.Fields(7) = Format(Val(Data3.Recordset.Fields(4)) - Val(Data3.Recordset.Fields(5)), "#0.00")
M7 = Data3.Recordset.Fields(7)
End If
Data3.Recordset.Update

Data3.Recordset.AddNew
Data3.Recordset.Fields(2) = "本期发生额及余额"
Data3.Recordset.Fields(3) = M3
Data3.Recordset.Fields(4) = M4
Data3.Recordset.Fields(5) = M5
Data3.Recordset.Fields(6) = M6
Data3.Recordset.Fields(7) = M7
Data3.Recordset.Fields(9) = "3"
Data3.Recordset.Update
Else
L = Format(Val(Data3.Recordset.Fields(7)), "#0.00")
Data3.Recordset.MoveNext
M3 = Data3.Recordset.Fields(3)
M4 = Val(Data3.Recordset.Fields(4))
M5 = Val(Data3.Recordset.Fields(5))
M6 = Data3.Recordset.Fields(6)
Data3.Recordset.Edit
If Data3.Recordset.Fields(6) = "贷" Then
Data3.Recordset.Fields(7) = Format(Format(Val(Data3.Recordset.Fields(5)) - Val(Data3.Recordset.Fields(4)) + L, "#0.00"), "#0.00")
M7 = Data3.Recordset.Fields(7)
Else
Data3.Recordset.Fields(7) = Format(Format(Val(Data3.Recordset.Fields(4)) - Val(Data3.Recordset.Fields(5)) + L, "#0.00"), "#0.00")
M7 = Data3.Recordset.Fields(7)
End If
Data3.Recordset.Update

Data3.Recordset.AddNew
Data3.Recordset.Fields(2) = "本期发生额及余额"
Data3.Recordset.Fields(3) = M3
Data3.Recordset.Fields(4) = M4
Data3.Recordset.Fields(5) = M5
Data3.Recordset.Fields(6) = M6
Data3.Recordset.Fields(7) = M7
Data3.Recordset.Fields(9) = "3"
Data3.Recordset.Update

End If

Data2.Recordset.MoveNext
Loop


Data2.RecordSource = "SELECT * FROM ZLCX WHERE ZLCX.序号='1'"
Data2.Refresh

Data3.RecordSource = "SELECT * FROM ZLCX "
Data3.Refresh

Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Data3.Recordset.FindFirst "序号='2' AND 会计科目='" & Data2.Recordset.Fields(3) & "'"
If Data3.Recordset.NoMatch Then
Data3.Database.Execute "INSERT INTO ZLCX(日期,摘要,会计科目,借方金额,贷方金额,借贷方向,余额,序号) VALUES('" & DTPicker3.Value & "','本月合计','" & Data2.Recordset.Fields(3) & "','" & Data2.Recordset.Fields(4) & "','" & Data2.Recordset.Fields(5) & "','" & Data2.Recordset.Fields(6) & "','" & Data2.Recordset.Fields(7) & "','3')"
End If
Data2.Recordset.MoveNext
Loop

Data1.Database.Execute "UPDATE ZLCX SET 借方金额=format(借方金额,'#0.00'),贷方金额=format(贷方金额,'#0.00'),余额=format(余额,'#0.00')"
Data1.Database.Execute "DELETE * FROM ZLCX WHERE 会计科目=NULL"
Data1.Database.Execute "UPDATE ZLCX SET 凭证号='结-'+'" & Text3.Text & "' WHERE 凭证号=NULL"
Data1.RecordSource = "SELECT * FROM ZLCX ORDER BY ZLCX.会计科目,VAL(ZLCX.序号),ZLCX.日期"
Data1.Refresh



BT = "按总账汇总"

End Sub

Private Sub Command7_Click()
Unload Me
End Sub



Private Sub Command8_Click()
Call OutDataToExcel3(MSFlexGrid1, 5, 6, 8, "明细打印")
End Sub

Private Sub DBCombo1_Click(Area As Integer)
On Error Resume Next
If Combo1.Text = "" Then
MsgBox ("请输入总账科目")
Exit Sub
End If
If Combo1.Text = "应付账款" Then
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM ZLCX WHERE INSTR(ZLCX.会计科目,'" & Combo1.Text & "')>0   ORDER BY ZLCX.会计科目,VAL(ZLCX.序号),ZLCX.日期,VAL(RIGHT(ZLCX.凭证号,LEN(ZLCX.凭证号)-INSTR(ZLCX.凭证号,'-')))"
Data1.Refresh
Else
Data1.RecordSource = "SELECT * FROM ZLCX WHERE INSTR(ZLCX.会计科目,'" & Combo1.Text & "')>0 AND INSTR(ZLCX.会计科目,'" & DBCombo1.Text & "')>0 ORDER BY ZLCX.会计科目,VAL(ZLCX.序号),ZLCX.日期,VAL(RIGHT(ZLCX.凭证号,LEN(ZLCX.凭证号)-INSTR(ZLCX.凭证号,'-')))"
Data1.Refresh
End If
End If

If Combo1.Text = "应收账款" Then
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT * FROM ZLCX WHERE INSTR(ZLCX.会计科目,'" & Combo1.Text & "')>0 ORDER BY ZLCX.会计科目,VAL(ZLCX.序号),ZLCX.日期,VAL(RIGHT(ZLCX.凭证号,LEN(ZLCX.凭证号)-INSTR(ZLCX.凭证号,'-')))"
Data1.Refresh
Else
Data1.RecordSource = "SELECT * FROM ZLCX WHERE INSTR(ZLCX.会计科目,'" & Combo1.Text & "')>0 AND INSTR(ZLCX.会计科目,'" & DBCombo1.Text & "')>0 ORDER BY ZLCX.会计科目,VAL(ZLCX.序号),ZLCX.日期,VAL(RIGHT(ZLCX.凭证号,LEN(ZLCX.凭证号)-INSTR(ZLCX.凭证号,'-')))"
Data1.Refresh
End If
End If



End Sub

Private Sub DTPicker3_Change()
Data14.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data14.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between 起始日期 and 结束日期"
Data14.Refresh
If Data14.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data14.Recordset.Fields(0)
K2 = Data14.Recordset.Fields(1)
Text3.Text = Data14.Recordset.Fields(2)
End If
End Sub

Private Sub DTPicker3_CloseUp()
Data14.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data14.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between 起始日期 and 结束日期"
Data14.Refresh
If Data14.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data14.Recordset.Fields(0)
K2 = Data14.Recordset.Fields(1)
Text3.Text = Data14.Recordset.Fields(2)
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DTPicker3 = Date
DTPicker1 = Date
DTPicker2 = Date
DBCombo1.Text = ""


Data14.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data14.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between 起始日期 and 结束日期"
Data14.Refresh
If Data14.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data14.Recordset.Fields(0)
K2 = Data14.Recordset.Fields(1)
Text3.Text = Data14.Recordset.Fields(2)
End If

Combo1.Text = ""
Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.MDB"
Data1.RecordSource = "SELECT * FROM MXFLZ WHERE MXFLZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') ORDER BY MXFLZ.日期,MXFLZ.凭证号"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.MDB"
Data2.Refresh

Data3.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.MDB"
Data3.Refresh

Data4.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.MDB"
Data4.RecordSource = "CWMC"
Data4.Refresh

Data5.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.MDB"
Data5.Refresh

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(3) = 2000
MSFlexGrid1.ColWidth(4) = 2500
MSFlexGrid1.ColWidth(7) = 700
MSFlexGrid1.ColWidth(8) = 700
MSFlexGrid1.ColWidth(9) = 700
End Sub

Private Sub MSFlexGrid1_dblClick()
With MSFlexGrid1
    c = .Col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlexGrid1_dblClick
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid1.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid1.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid1.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data1.Recordset.MoveFirst
Data1.Recordset.Move r - 1
Data1.Recordset.Edit
Data1.Recordset.Fields(c - 1) = Text1111.Text
Data1.Recordset.Update
Text1111.Visible = False
End Sub

