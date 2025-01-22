VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy51 
   BackColor       =   &H00C0E0FF&
   Caption         =   "生成采购表"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form30"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1920
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   240
      Top             =   1320
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "采购表不正确"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "采购表正确"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "采购结束"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除记录"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   1200
      TabIndex        =   27
      Text            =   "Text1111"
      Top             =   5280
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "采购表打印"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "采购批次统一"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "进入出库操作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库类信息"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3360
      Width           =   1695
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy51.frx":0000
      Height          =   330
      Left            =   8760
      TabIndex        =   21
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo2"
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库存信息"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查看采购表"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "采购信息"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查看备料表"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出本操作"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   39177
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy51.frx":0014
      Height          =   1935
      Left            =   3600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "记录号 "
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   1680
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy51.frx":0028
      Height          =   5535
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3960
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "记录号 "
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formy51.frx":003C
      Height          =   5535
      Left            =   7560
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3960
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "记录号 "
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   39883
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择库类"
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
      Left            =   7560
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
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
      Left            =   7560
      TabIndex        =   19
      Top             =   2520
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
      Left            =   9600
      TabIndex        =   18
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
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
      Left            =   1680
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择单号"
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
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Formy51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2, M1, M2, M3, M4, M5 As String: Public c, r, S1, S2 As Integer
Private Sub Command1_Click()
Data2.RecordSource = "SELECT DHCLB.材料库类,DHCLB.材料名称,DHCLB.材料规格,DHCLB.材料单位,DHCLB.材料颜色,DHCLB.材料批号,SUM(DHCLB.材料数量) AS 备料量 FROM DHCLB WHERE DHCLB.单号='" & DBCombo1.Text & "' GROUP BY DHCLB.材料库类,DHCLB.材料名称,DHCLB.材料规格,DHCLB.材料单位,DHCLB.材料颜色,DHCLB.材料批号"
Data2.Refresh
Call SX2(Data2, MSFlexGrid2, 7)
End Sub

Private Sub Command10_Click()
On Error Resume Next
Data2.Recordset.MoveFirst
Data2.Recordset.Move S1 - 1
p = S2 - S1 + 1
For i = 1 To p
Data2.Recordset.Delete
Data2.Recordset.MoveNext
Next
Data2.Refresh
End Sub

Private Sub Command11_Click()
If MsgBox("确定采购结束吗 单号：" + DBCombo1.Text, vbYesNo) = vbNo Then Exit Sub
Data2.Database.Execute "UPDATE SCZY_ZDH SET 入库='已' WHERE 单号='" & DBCombo1.Text & "'"
Data1.Refresh
End Sub

Private Sub Command12_Click()
Data2.Database.Execute "UPDATE SCZY_ZDH SET B1='Y' WHERE 单号='" & DBCombo1.Text & "'"
End Sub

Private Sub Command13_Click()
Data2.Database.Execute "UPDATE SCZY_ZDH SET B1='N' WHERE 单号='" & DBCombo1.Text & "'"
End Sub

Private Sub Command14_Click()
Data2.Database.Execute "UPDATE SCZY_ZDH SET B1=''  WHERE B1=NULL AND INSTR(SCZY_ZDH.单号,'L')>0 AND (入库=NULL OR 入库<>'已') AND SCZY_ZDH.日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
Data1.Recordset.MoveFirst
p = 1
Do While Not Data1.Recordset.EOF
If Data1.Recordset.Fields(24) = "Y" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbGreen
End If

If Data1.Recordset.Fields(24) = "N" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbRed
End If

If Data1.Recordset.Fields(24) = "" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbCyan
End If

Data1.Recordset.MoveNext
p = p + 1
Loop

End Sub

Private Sub Command2_Click()
On Error Resume Next
Data1.Database.Execute "DELETE * FROM CKGL"

Data3.Database.Execute "INSERT INTO CKGL(单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,入库数量) IN'd:\数据库\\htgl\2011\SCZYJHD.MDB' SELECT 单号,库类,材料名称,材料规格,材料单位,颜色,批次,SUM(数量) FROM CKGL WHERE CKGL.单号='" & DBCombo1.Text & "' GROUP BY 单号,库类,材料名称,材料规格,材料单位,颜色,批次 "
Data1.Database.Execute "UPDATE CKGL SET LX=CK,采购数量=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(单号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,采购数量) SELECT CGCLB.单号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号,SUM(CGCLB.材料数量) AS 采购数量 FROM CGCLB WHERE CGCLB.单号='" & DBCombo1.Text & "' GROUP BY CGCLB.单号,CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号"
Data1.Database.Execute "UPDATE CKGL SET LX=CK,入库数量=0 WHERE LX=NULL"
Data2.RecordSource = "SELECT CKGL.单号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料单位,CKGL.材料颜色,SUM(CKGL.采购数量) AS 采购量,SUM(CKGL.入库数量) AS 入库量 FROM CKGL WHERE  CKGL.单号='" & DBCombo1.Text & "' GROUP BY CKGL.单号,CKGL.材料库类,CKGL.材料名称,CKGL.材料规格,CKGL.材料单位,CKGL.材料颜色"
Data2.Refresh
Call SX2(Data2, MSFlexGrid2, 7)
Call SX2(Data2, MSFlexGrid2, 8)
End Sub

Private Sub Command3_Click()
Data3.Database.Execute "DELETE * FROM CLRCZZ"
Data3.Database.Execute "DELETE * FROM CLRCZZHZ"
Data3.Database.Execute "INSERT INTO CLRCZZ(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类) select CKGL.材料名称,CKGL.材料规格,CKGL.材料单位,CKGL.颜色,CKGL.批次,CKGL.数量,CKGL.单价,CKGL.库类 from ckgl WHERE CKGL.库别='清库库存' AND CKGL.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data3.Database.Execute "UPDATE CLRCZZ SET 库别='入库' where 库别=NULL"
Data3.Database.Execute "INSERT INTO CLRCZZ(材料名称,材料规格,材料单位,颜色,批次,数量,单价,库类) select CKBL.材料名称,CKBL.材料规格,CKBL.材料单位,CKBL.颜色,CKBL.批次,CKBL.数量,CKBL.单价,CKBL.库类 from ckBL WHERE CKBL.库别='清库库存' AND CKBL.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data3.Database.Execute "UPDATE CLRCZZ SET 库别='出库',数量=-数量 WHERE 库别=NULL"
Data3.Database.Execute "INSERT INTO CLRCZZHZ(库类,材料名称,材料规格,材料单位,颜色,批次,数量,单价) SELECT CLRCZZ.库类,CLRCZZ.材料名称,CLRCZZ.材料规格,CLRCZZ.材料单位,CLRCZZ.颜色,CLRCZZ.批次,SUM(CLRCZZ.数量) AS L,AVG(CLRCZZ.单价) AS D FROM CLRCZZ GROUP BY CLRCZZ.库类,CLRCZZ.材料名称,CLRCZZ.材料规格,CLRCZZ.材料单位,CLRCZZ.颜色,CLRCZZ.批次"
Data4.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.数量>0"
Data4.Refresh
End Sub

Private Sub Command4_Click()
Data2.RecordSource = "SELECT CGCLB.材料库类,CGCLB.材料名称,CGCLB.材料规格,CGCLB.材料单位,CGCLB.材料颜色,CGCLB.材料批号,CGCLB.材料数量 AS 采购量 FROM CGCLB WHERE CGCLB.单号='" & DBCombo1.Text & "' AND CGCLB.材料数量>0 ORDER BY 材料库类,材料名称,CGCLB.材料规格,材料颜色"
Data2.Refresh
Call SX2(Data2, MSFlexGrid2, 7)
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Data4.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.库类='" & DBCombo2.Text & "' AND CLRCZZHZ.数量>0"
Data4.Refresh
End Sub

Private Sub Command7_Click()
On Error Resume Next
Formy52.DBCombo1(12).Text = Data4.Recordset.Fields(7)
Formy52.DBCombo1(3).Text = Data4.Recordset.Fields(0)
Formy52.DBCombo2.Text = Data4.Recordset.Fields(3)
Formy52.DBCombo1(1).Text = DBCombo1.Text
Formy52.Text2.Text = DBCombo1.Text
End Sub

Private Sub Command8_Click()
l = Format(Date, "YYMMDD")
Data2.Database.Execute "UPDATE CGCLB SET 材料批号='" & l & "' WHERE 单号='" & DBCombo1.Text & "'"
Data2.Refresh
End Sub

Private Sub Command9_Click()
On Error Resume Next
If Data2.Recordset.EOF Then
MsgBox ("无记录不能打印！")
Exit Sub
End If
Call MXOutDataToExcel(MSFlexGrid2, "单号： " + DBCombo1.Text + "合约号：" + Data1.Recordset.Fields(8) + "  采购表")
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

Private Sub DTPicker3_Change()
Text3.Text = Month(DTPicker3.Value)
End Sub

Private Sub DTPicker3_CloseUp()
Text3.Text = Month(DTPicker3.Value)
End Sub


Private Sub Form_Load()
DTPicker3.Value = Date
Text3.Text = Month(DTPicker1.Value)
Select Case Text3.Text
       Case 1
K1 = Format(Date, "YYYY") + "-" + "01" + "-01"
K2 = Format(Date, "YYYY") + "-" + "01" + "-31"
       Case 2
If Val(Format(Date, "YYYY")) / 4 = Int(Val(Format(Date, "YYYY")) / 4) Then
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-29"
Else
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-28"
End If
       Case 3
K1 = Format(Date, "YYYY") + "-" + "03" + "-01"
K2 = Format(Date, "YYYY") + "-" + "03" + "-31"
       Case 4
K1 = Format(Date, "YYYY") + "-" + "04" + "-01"
K2 = Format(Date, "YYYY") + "-" + "04" + "-30"
       Case 5
K1 = Format(Date, "YYYY") + "-" + "05" + "-01"
K2 = Format(Date, "YYYY") + "-" + "05" + "-31"
       Case 6
K1 = Format(Date, "YYYY") + "-" + "06" + "-01"
K2 = Format(Date, "YYYY") + "-" + "06" + "-30"
       Case 7
K1 = Format(Date, "YYYY") + "-" + "07" + "-01"
K2 = Format(Date, "YYYY") + "-" + "07" + "-31"
       Case 8
K1 = Format(Date, "YYYY") + "-" + "08" + "-01"
K2 = Format(Date, "YYYY") + "-" + "08" + "-30"
       Case 9
K1 = Format(Date, "YYYY") + "-" + "09" + "-01"
K2 = Format(Date, "YYYY") + "-" + "09" + "-31"
       Case 10
K1 = Format(Date, "YYYY") + "-" + "10" + "-01"
K2 = Format(Date, "YYYY") + "-" + "10" + "-30"
       Case 11
K1 = Format(Date, "YYYY") + "-" + "11" + "-01"
K2 = Format(Date, "YYYY") + "-" + "11" + "-31"
       Case 12
K1 = Format(Date, "YYYY") + "-" + "12" + "-01"
K2 = Format(Date, "YYYY") + "-" + "12" + "-30"
End Select
Text4.Text = Date - 15
Text5.Text = Date
DTPicker1.Value = Date - 15
DTPicker2.Value = Date
DBCombo1.Text = ""
DBCombo2.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data1.RecordSource = "SELECT * FROM SCZY_ZDH WHERE INSTR(SCZY_ZDH.单号,'L')>0 AND (入库=NULL OR 入库<>'已') AND SCZY_ZDH.日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
Data2.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.MDB"
Data3.DatabaseName = "d:\数据库\\htgl\2011\ckgl.MDB"
Data4.DatabaseName = "d:\数据库\\htgl\2011\ckgl.MDB"
Data5.DatabaseName = "d:\数据库\\htgl\2011\ckgl.MDB"
Data5.RecordSource = "SELECT KL.MC FROM KL GROUP BY KL.MC"
Data5.Refresh
MSFlexGrid1.ColWidth(8) = 1500
End Sub

Private Sub Label2_Click()
Data1.RecordSource = "SELECT * FROM SCZY_ZDH WHERE INSTR(SCZY_ZDH.单号,'L')>0 AND (入库=NULL OR 入库<>'已') AND SCZY_ZDH.日期 BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
End Sub

Private Sub MSFlexGrid1_dblClick()
rs = MSFlexGrid1.Row
If Data1.Recordset.EOF Then
DBCombo1.Text = ""
Exit Sub
End If

Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
DBCombo1.Text = Data1.Recordset.Fields(7)
End Sub

Private Sub MSFlexGrid2_Click()
On Error Resume Next
rs = MSFlexGrid2.Row
'If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
DBCombo2.Text = Data2.Recordset.Fields(0)
Data4.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.库类='" & Data2.Recordset.Fields(0) & "' AND CLRCZZHZ.材料名称='" & Data2.Recordset.Fields(1) & "' AND 颜色='" & Data2.Recordset.Fields(4) & "' AND CLRCZZHZ.数量>0"
Data4.Refresh
End Sub

Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid2.RowSel
End Sub

Private Sub MSFlexGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid2.RowSel
End Sub


Private Sub MSFlexGrid3_DBLClick()
On Error Resume Next
rs = MSFlexGrid3.Row
Data4.Recordset.MoveFirst
Data4.Recordset.Move rs - 1
Formy52.DBCombo1(12).Text = Data4.Recordset.Fields(7)
Formy52.DBCombo1(3).Text = Data4.Recordset.Fields(0)
Formy52.DBCombo2.Text = Data4.Recordset.Fields(3)
Formy52.DBCombo1(1).Text = DBCombo1.Text
End Sub

Private Sub Text3_Change()
Select Case Text3.Text
       Case 1
K1 = Format(Date, "YYYY") + "-" + "01" + "-01"
K2 = Format(Date, "YYYY") + "-" + "01" + "-31"
       Case 2
If Val(Format(Date, "YYYY")) / 4 = Int(Val(Format(Date, "YYYY")) / 4) Then
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-29"
Else
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-28"
End If
       Case 3
K1 = Format(Date, "YYYY") + "-" + "03" + "-01"
K2 = Format(Date, "YYYY") + "-" + "03" + "-31"
       Case 4
K1 = Format(Date, "YYYY") + "-" + "04" + "-01"
K2 = Format(Date, "YYYY") + "-" + "04" + "-30"
       Case 5
K1 = Format(Date, "YYYY") + "-" + "05" + "-01"
K2 = Format(Date, "YYYY") + "-" + "05" + "-31"
       Case 6
K1 = Format(Date, "YYYY") + "-" + "06" + "-01"
K2 = Format(Date, "YYYY") + "-" + "06" + "-30"
       Case 7
K1 = Format(Date, "YYYY") + "-" + "07" + "-01"
K2 = Format(Date, "YYYY") + "-" + "07" + "-31"
       Case 8
K1 = Format(Date, "YYYY") + "-" + "08" + "-01"
K2 = Format(Date, "YYYY") + "-" + "08" + "-30"
       Case 9
K1 = Format(Date, "YYYY") + "-" + "09" + "-01"
K2 = Format(Date, "YYYY") + "-" + "09" + "-31"
       Case 10
K1 = Format(Date, "YYYY") + "-" + "10" + "-01"
K2 = Format(Date, "YYYY") + "-" + "10" + "-30"
       Case 11
K1 = Format(Date, "YYYY") + "-" + "11" + "-01"
K2 = Format(Date, "YYYY") + "-" + "11" + "-31"
       Case 12
K1 = Format(Date, "YYYY") + "-" + "12" + "-01"
K2 = Format(Date, "YYYY") + "-" + "12" + "-30"
End Select

End Sub
Private Sub MSFlexGrid2_dblClick()
With MSFlexGrid2
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

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlexGrid2_dblClick
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid2.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid2.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid2.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data2.Recordset.MoveFirst
Data2.Recordset.Move r - 1
Data2.Recordset.Edit
Data2.Recordset.Fields(c - 1) = Text1111.Text
Data2.Recordset.Update
Text1111.Visible = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Data1.Refresh
Data1.Recordset.MoveFirst
p = 1
Do While Not Data1.Recordset.EOF

If Data1.Recordset.Fields(24) = "Y" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbGreen
End If

If Data1.Recordset.Fields(24) = Null Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbRed
End If

If Data1.Recordset.Fields(24) = "N" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbCyan
End If

Data1.Recordset.MoveNext
p = p + 1
Loop

End Sub

