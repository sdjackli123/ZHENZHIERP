VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formw97 
   BackColor       =   &H00C0E0FF&
   Caption         =   "成品库存记录"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form16"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command11 
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
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
      Top             =   10560
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Enabled         =   0   'False
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
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
      Height          =   375
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
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
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
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
   Begin VB.CommandButton Command5 
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw97.frx":0000
      Height          =   6495
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   11456
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
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
      Height          =   390
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
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
      Height          =   390
      Index           =   2
      Left            =   1920
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
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
      Height          =   390
      Index           =   3
      Left            =   6120
      TabIndex        =   9
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
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
      Height          =   390
      Index           =   4
      Left            =   6120
      TabIndex        =   10
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
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
      Bindings        =   "Formw97.frx":0014
      Height          =   390
      Index           =   5
      Left            =   6120
      TabIndex        =   11
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
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
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   390
      Index           =   6
      Left            =   6120
      TabIndex        =   12
      Top             =   2160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
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
      Height          =   390
      Index           =   7
      Left            =   10920
      TabIndex        =   20
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
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
      Height          =   390
      Index           =   8
      Left            =   10920
      TabIndex        =   22
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
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
      Height          =   390
      Index           =   9
      Left            =   1920
      TabIndex        =   24
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
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
      Height          =   390
      Index           =   10
      Left            =   10920
      TabIndex        =   25
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
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
      Height          =   390
      Index           =   11
      Left            =   10920
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
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
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
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
      Index           =   1
      Left            =   9360
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
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
      Left            =   480
      TabIndex        =   28
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Index           =   6
      Left            =   480
      TabIndex        =   27
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "条码"
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
      Left            =   9360
      TabIndex        =   26
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
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
      Index           =   5
      Left            =   9360
      TabIndex        =   23
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
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
      Index           =   4
      Left            =   9360
      TabIndex        =   21
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
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
      Left            =   4680
      TabIndex        =   19
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Left            =   4680
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "规格"
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
      Left            =   4680
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
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
      Left            =   480
      TabIndex        =   14
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "数量"
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
      Left            =   4680
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "Formw97"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer
Private Sub Command1_Click()
If DBCombo1(2).Text = "" Or DBCombo1(6).Text = "" Then
Exit Sub
End If
Data1.Recordset.AddNew
For i = 0 To 11
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh
DBCombo1(6).Text = 0
Data6.RecordSource = "SELECT max(序号) FROM LSJL WHERE 日期=CDATE('" & DBCombo1(0).Text & "')"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If

End Sub


Private Sub Command11_Click()
On Error Resume Next
DBCombo1(6).Text = 0
Data6.RecordSource = "SELECT max(序号) FROM LSJL WHERE 日期=CDATE('" & DBCombo1(0).Text & "')"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
Data1.Refresh
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If DBCombo1(2).Text = "" Or DBCombo1(6).Text = "" Then
Exit Sub
End If
If MsgBox("确定修改吗", vbYesNo) = vbNo Then Exit Sub
If DBCombo1(0).Text = "" Then Exit Sub
Data1.Recordset.Edit
For i = 0 To 10
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Update
Data1.Refresh
DBCombo1(6).Text = 0
Data6.RecordSource = "SELECT max(序号) FROM LSJL WHERE 日期=CDATE('" & DBCombo1(0).Text & "')"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If

End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除吗", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh
Data6.RecordSource = "SELECT max(序号) FROM LSJL WHERE 日期=CDATE('" & DBCombo1(0).Text & "')"
Data6.Refresh
DBCombo1(6).Text = 0
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If
End Sub

Private Sub Command5_Click()
Call OutDataToExcel(MSFlexGrid1, 7, DBCombo1(0).Text)
End Sub


Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data1.RecordSource = "SELECT * FROM LSJL WHERE 日期=CDATE('" & DBCombo1(0).Text & "') order by 序号 DESC"
Data1.Refresh
End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data1.RecordSource = "SELECT * FROM LSJL WHERE 日期=CDATE('" & DBCombo1(0).Text & "') order by 序号"
Data1.Refresh
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
S1 = 0
S2 = 0
For i = 0 To 11
DBCombo1(i).Text = ""
Next
DBCombo1(0).Text = Date
DBCombo1(5).Text = "件"
DBCombo1(11).Text = "0"
Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data1.RecordSource = "SELECT * FROM LSJL WHERE 日期=CDATE('" & DBCombo1(0).Text & "') ORDER BY 序号 DESC"
Data1.Refresh
Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select 简称 from khzl GROUP BY 简称"
Data3.Refresh
Data4.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data4.RecordSource = "select MC from CLDW GROUP BY MC"
Data4.Refresh
Data5.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data6.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data6.RecordSource = "SELECT max(序号) FROM LSJL WHERE 日期=CDATE('" & DBCombo1(0).Text & "')"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If


MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1200
For i = 1 To 7
MSFlexGrid1.ColWidth(i) = 1600
Next
MSFlexGrid1.ColWidth(12) = 0
MSFlexGrid1.ColWidth(13) = 0
MSFlexGrid1.ColWidth(14) = 0
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Label1_dblClick(Index As Integer)
Select Case Index
       Case 1
khbl = 10
Formw202.Text1.Text = DBCombo1(1).Text
Formw202.Show
End Select
End Sub

Private Sub Label3_Click()
khbl = 2
Formw99.Text1 = DBCombo1(1).Text
Formw99.Show
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
If Data1.Recordset.EOF Then Exit Sub
rs = MSFlexGrid1.Row
Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To 10
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next
Command3.Enabled = True
Command4.Enabled = True
Command1.Enabled = False
End Sub
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid1.RowSel
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid1.RowSel
End Sub

