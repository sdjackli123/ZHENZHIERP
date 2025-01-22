VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formw95 
   BackColor       =   &H00C0E0FF&
   Caption         =   "扫描出库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data8 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Data Data7 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "装箱打印"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "详单刷新"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
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
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8160
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw95.frx":0000
      Height          =   6615
      Left            =   5280
      TabIndex        =   2
      Top             =   480
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw95.frx":0014
      Height          =   390
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "xm"
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
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw95.frx":0028
      Height          =   390
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   688
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "mc"
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formw95.frx":003C
      Height          =   4575
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formw95.frx":0050
      Height          =   2055
      Left            =   5280
      TabIndex        =   13
      Top             =   7200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "箱号"
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
      Left            =   2280
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "购货单位"
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
      Left            =   360
      TabIndex        =   11
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "详单编号"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "保管"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "扫描区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   8160
      Width           =   975
   End
End
Attribute VB_Name = "Formw95"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'On Error Resume Next
If Text2.Text = "" Then
MsgBox ("请输入发货详单编号")
Exit Sub
End If
Data1.RecordSource = "select * from zxd where 编号='" & Text2.Text & "'"
Data1.Refresh
If Data1.Recordset.EOF Then
MsgBox ("不存在发货编号")
Exit Sub
End If
Data1.Recordset.MoveFirst
DBCombo1.Text = Data1.Recordset.Fields(0)
Data4.Database.Execute "delete * from zxdf"    ''''''
Do While Not Data1.Recordset.EOF
For i = 6 To 14
If Val(Data1.Recordset.Fields(i)) > 0 Then
Data4.Database.Execute "insert into zxdf(客户,款号,规格,颜色,发货量,编号) VALUES('" & Data1.Recordset.Fields(0) & "','" & Data1.Recordset.Fields(1) & "','" & Data1.Recordset.Fields(2) & "','" & Data1.Recordset.Fields(i - 1) & "','" & Data1.Recordset.Fields(i) & "','" & Data1.Recordset.Fields(17) & "')"
End If
i = i + 2
Next
Data1.Recordset.MoveNext
Loop
Data4.Database.Execute "insert into zxdf(客户,款号,规格,颜色,出库量,编号) select 购货单位,款号,型号,规格,sum(数量),单据号 from lsfh where 单据号='" & Text2.Text & "' group by 购货单位,款号,型号,规格,单据号"
Data4.Database.Execute "update zxdf set 共量='1'"
Data4.Database.Execute "update zxdf set 出库量='0' where 出库量=null"
Data4.Database.Execute "update zxdf set 发货量='0' where 发货量=null"

Data4.Database.Execute "insert into zxdf(客户,款号,规格,颜色,编号,出库量,发货量) select 客户,款号,规格,颜色,编号,sum(val(出库量)),sum(val(发货量)) from zxdf where 编号='" & Text2.Text & "' group by 客户,款号,规格,颜色,编号"
Data4.Database.Execute "delete * from zxdf where 共量='1'"

Data2.RecordSource = "select 购货单位,单号,款号,品名,规格,型号,单位,数量,条码,仓务员 from lsfh where 单据号='" & Text2.Text & "'"
Data2.Refresh

Data5.RecordSource = "select 款号,规格,颜色,发货量,出库量 from zxdf order by 款号,规格,颜色"
Data5.Refresh
Call sx(MSFlexGrid3)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If Text2.Text = "" Then
MsgBox ("请输入详单编号")
Exit Sub
End If
Call fhmxdy(Data7, Data3, Text2.Text)
End Sub

Private Sub Form_Load()
Dim l As Integer
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
DBCombo1.Text = ""
DBCombo2.Text = ""
m = ""
Data4.DatabaseName = "d:\数据库\\htgl\2011\CPCK.mdb"

Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"

Data2.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data2.RecordSource = "select * from lsfh where 单据号='" & Text2.Text & "' order by 序号 desc"
Data2.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\CPCK.mdb"

Data6.DatabaseName = "d:\数据库\\htgl\2011\cpck.mdb"

Data3.DatabaseName = "d:\数据库\\htgl\2011\cpck.mdb"

Data7.DatabaseName = "d:\数据库\\htgl\2011\cpck.mdb"

Data8.DatabaseName = "d:\数据库\\htgl\2011\ckgl.mdb"
Data8.RecordSource = "select fzr.xm  from fzr group by fzr.xm"
Data8.Refresh

MSFlexGrid1.ColWidth(11) = 1200
MSFlexGrid1.ColWidth(10) = 1200
MSFlexGrid3.ColWidth(0) = 200
MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid2.ColWidth(0) = 200

End Sub


Private Sub Label5_dblClick()
Data7.RecordSource = "SELECT * FROM LSFH WHERE 日期=cdate('" & Date & "')"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.RecordSource = "select max(mid(发货地,7)) from lsfh where 日期=cdate('" & Date & "')"
Data7.Refresh
If Len(Data7.Recordset.Fields(0) + 1) < 2 Then
Text3.Text = "C" + Format(Date, "mmdd") + "-" + "0" + Trim(Data7.Recordset.Fields(0) + 1)
Else
Text3.Text = "C" + Format(Date, "mmdd") + "-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Else
Text3.Text = "C" + Format(Date, "mmdd") + "-" + "01"
End If
End Sub

Private Sub MSFlexGrid1_dblClick()
If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
rs = MSFlexGrid1.Row
rc = MSFlexGrid1.Col
Data2.Recordset.Move rs - 1
If rc = 1 Then
Data2.Recordset.Delete
Data2.Refresh
End If
End Sub

Private Sub MSFlexGrid3_DBLClick()
''款号,规格,颜色
If Data5.Recordset.EOF Then Exit Sub
rs = MSFlexGrid3.Row
Data5.Recordset.MoveFirst
Data5.Recordset.Move rs - 1
Data4.Database.Execute "delete * from lscx"
Data4.Database.Execute "insert into lscx(单号,款号,品名,规格,型号,单位,数量,条码,备注) select 单号,款号,品名,规格,型号,单位,数量,条码,备注 from lsrk where 款号='" & Data5.Recordset.Fields(0) & "' and 规格='" & Data5.Recordset.Fields(2) & "' and 型号='" & Data5.Recordset.Fields(1) & "'"
Data4.Database.Execute "insert into lscx(单号,款号,品名,规格,型号,单位,数量,条码,备注) select 单号,款号,品名,规格,型号,单位,-数量,条码,备注 from lsfh where 款号='" & Data5.Recordset.Fields(0) & "' and 规格='" & Data5.Recordset.Fields(2) & "' and 型号='" & Data5.Recordset.Fields(1) & "'"
Data4.Database.Execute "update lscx set 共量='1'"
Data4.Database.Execute "insert into lscx(单号,款号,品名,规格,型号,单位,数量,条码,备注) select 单号,款号,品名,规格,型号,单位,sum(数量),条码,备注 from lscx group by 单号,款号,品名,规格,型号,单位,条码,备注"
Data4.Database.Execute "delete * from lscx where 共量='1' or 数量<=0"
Data6.RecordSource = "select 单号,款号,品名,规格,型号,单位,数量,条码,备注 from lscx"
Data6.Refresh
End Sub

Private Sub Text1_Change()
If DBCombo1.Text = "" Then Exit Sub

If InStr(Text1.Text, "J") > 0 Then
m = Left(Text1.Text, Len(Text1.Text) - 1)

If Len(m) = 9 Then

If Text3.Text = "" Then
MsgBox ("请输入箱号")
Exit Sub
End If

Data4.RecordSource = "SELECT * FROM LSRK"
Data4.Refresh

Data4.Recordset.FindFirst "条码='" & m & "'"
If Data4.Recordset.NoMatch Then
Label2.Caption = "不存在此条码"
Text1.Text = ""
Timer1.Enabled = True
Exit Sub
Else
Data6.RecordSource = "SELECT * FROM LSFH WHERE 条码='" & m & "'"
Data6.Refresh
If Data6.Recordset.EOF Then

l = 1
Data3.RecordSource = "SELECT 序号 FROM LSFH WHERE 单据号='" & Text2.Text & "' ORDER BY 序号 DESC"
Data3.Refresh
If Data3.Recordset.EOF Then
l = 1
Else
l = Data3.Recordset.Fields(0) + 1
End If
Data5.Database.Execute "INSERT INTO lsfh(日期,单号,款号,品名,规格,型号,单位,数量,备注,条码,序号,购货单位,仓务员,单据号,发货地) select 日期,单号,款号,品名,规格,型号,单位,数量,备注,条码,'" & l & "','" & DBCombo1.Text & "','" & DBCombo2.Text & "','" & Text2.Text & "','" & Text3.Text & "' from lsrk where 条码='" & m & "'"
End If
Data2.RecordSource = "select 购货单位,单号,款号,品名,规格,型号,单位,数量,条码,发货地 as 箱号,仓务员 from lsfh where 单据号='" & Text2.Text & "'"
Data2.Refresh
Text1.Text = ""
Text1.SetFocus
End If

Else
Text1.Text = ""
Text1.SetFocus

End If
End If

End Sub


Private Sub sx(MSF As MSFlexGrid)

    Dim i     As Integer
      With MSF
                 .AllowBigSelection = True           '   设置网格样式
                 .FillStyle = flexFillRepeat
                For i = 1 To .Rows - 1
                        .Row = i:       .Col = .FixedCols
                        .ColSel = .Cols() - .FixedCols - 1
                         If Val(MSF.TextMatrix(i, 4)) < Val(MSF.TextMatrix(i, 5)) Then
                              .CellBackColor = vbGreen           '兰色
                        Else
                              .CellBackColor = vbBlack      ' 黑色
                        End If
                Next i
        End With
End Sub

