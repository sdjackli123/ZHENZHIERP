VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Formy112 
   BackColor       =   &H00C0E0FF&
   Caption         =   "款式报价"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   14100
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2880
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data6 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data5 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data4 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "重算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   720
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy112.frx":0000
      Height          =   330
      Index           =   0
      Left            =   4680
      TabIndex        =   0
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy112.frx":0014
      Height          =   330
      Index           =   3
      Left            =   4680
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "mc"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy112.frx":0028
      Height          =   330
      Index           =   2
      Left            =   4680
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   4
      Left            =   4680
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   5
      Left            =   8040
      TabIndex        =   9
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   6
      Left            =   8040
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   7
      Left            =   8040
      TabIndex        =   20
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy112.frx":003D
      Height          =   6135
      Left            =   360
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3960
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   8
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      FormatString    =   "记录号 "
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   8
      Left            =   8040
      TabIndex        =   24
      Top             =   2520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DBCombo1"
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   9
      Left            =   6840
      TabIndex        =   23
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "规格"
      Height          =   375
      Index           =   8
      Left            =   3480
      TabIndex        =   22
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   360
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "销售"
      Height          =   375
      Index           =   7
      Left            =   6840
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "提货"
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "折扣"
      Height          =   375
      Index           =   2
      Left            =   6840
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "品名"
      Height          =   375
      Index           =   4
      Left            =   3480
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单位"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户"
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Formy112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strname As String
Dim Stm As New ADODB.Stream
Dim StrPicTemp As String
Private Sub Command1_Click()
If DBCombo1(0).Text = "" Or DBCombo1(1).Text = "" Then
Exit Sub
End If

Stm.Type = adTypeBinary
Stm.Open
If strname <> "" Then
Stm.LoadFromFile strname
End If


Data1.Recordset.AddNew
For i = 0 To 8
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Fields(9) = Stm.Read
Data1.Recordset.Update
Data1.Refresh

Stm.Close
strname = ""
Image1.Picture = Nothing

For i = 1 To 6
DBCombo1(i).Text = ""
Next
Data6.RecordSource = "SELECT max(序号) FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If

End Sub

Private Sub Command2_Click()
If DBCombo1(0).Text = "" Or DBCombo1(1).Text = "" Then
Exit Sub
End If

Stm.Type = adTypeBinary
Stm.Open
If strname <> "" Then
Stm.LoadFromFile strname
End If


Data1.Recordset.Edit
For i = 0 To 8
Data1.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data1.Recordset.Fields(9) = Stm.Read
Data1.Recordset.Update
Data1.Refresh

Stm.Close
strname = ""
Image1.Picture = Nothing

For i = 1 To 6
DBCombo1(i).Text = ""
Next
Data6.RecordSource = "SELECT max(序号) FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Data1.Recordset.EOF Then Exit Sub
If MsgBox("确定删除吗？，删除不能恢复", vbYesNo) = vbNo Then Exit Sub
Data1.Recordset.Delete
Data1.Refresh

For i = 1 To 6
DBCombo1(i).Text = ""
Next

Data6.RecordSource = "SELECT max(序号) FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If

End Sub

Private Sub Command5_Click()
Data5.Database.Execute "update ksbj set 提货=format(val(销售)*val(折扣))"
MsgBox ("重算完成！")
End Sub

Private Sub Command6_Click()
Call ksbj(Data5, DBCombo1(0).Text)
End Sub

Private Sub Command7_Click()
If DBCombo1(1).Text = "" Then
MsgBox ("请输入款号")
Exit Sub
End If
Data1.RecordSource = "SELECT * FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "' and instr(款号,'" & DBCombo1(1).Text & "')>0 ORDER BY 序号 desc"
Data1.Refresh
End Sub

Private Sub Command8_Click()
On Error Resume Next
Data6.RecordSource = "SELECT max(序号) FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If
Data1.RecordSource = "SELECT * FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "' ORDER BY 序号 desc"
Data1.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
Stm.Close
strname = ""
Image1.Picture = Nothing
End Sub

Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data1.RecordSource = "SELECT * FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "' ORDER BY 序号 desc"
Data1.Refresh
Data6.RecordSource = "SELECT max(序号) FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "'"
Data6.Refresh
DBCombo1(7).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(7).Text = Data6.Recordset.Fields(0) + 1
End If
      Case 5
DBCombo1(7).Text = Format(Val(DBCombo1(5).Text) * Val(DBCombo1(6).Text), "#0.00")
      Case 6
DBCombo1(7).Text = Format(Val(DBCombo1(5).Text) * Val(DBCombo1(6).Text), "#0.00")
End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 0
Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data1.RecordSource = "SELECT * FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "' ORDER BY 序号 desc"
Data1.Refresh
Data6.RecordSource = "SELECT max(序号) FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If
      Case 5
DBCombo1(7).Text = Format(Val(DBCombo1(5).Text) * Val(DBCombo1(6).Text), "#0.00")
      Case 6
DBCombo1(7).Text = Format(Val(DBCombo1(5).Text) * Val(DBCombo1(6).Text), "#0.00")
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
For i = 0 To 8
DBCombo1(i).Text = ""
Next

Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data1.RecordSource = "SELECT * FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "' ORDER BY 序号 desc"
Data1.Refresh
Data3.DatabaseName = "d:\数据库\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select 简称 from khzl GROUP BY 简称"
Data3.Refresh
Data4.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data4.RecordSource = "select MC from CLDW GROUP BY MC"
Data4.Refresh
Data5.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data6.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data6.RecordSource = "SELECT max(序号) FROM KSBJ WHERE 客户='" & DBCombo1(0).Text & "'"
Data6.Refresh
DBCombo1(8).Text = 1
If Not Data6.Recordset.EOF Then
DBCombo1(8).Text = Data6.Recordset.Fields(0) + 1
End If

Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1600
MSFlexGrid1.ColWidth(2) = 1600
MSFlexGrid1.ColWidth(3) = 1600

End Sub

Private Sub Image1_dblClick()
Call pict
End Sub

Private Sub pict()
CommonDialog1.Filter = "JPG图片(*.JPG)|*.JPG"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
 Image1.Picture = LoadPicture(CommonDialog1.FileName)
 strname = CommonDialog1.FileName
 Else
 MsgBox "没有选中图样"
End If

End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row

If Data1.Recordset.EOF Then Exit Sub

Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
For i = 0 To 8
DBCombo1(i).Text = Data1.Recordset.Fields(i)
Next

If Data1.Recordset.Fields(9) <> Null Then
''''''''''''''''''''''''''''''''''''''''图片
     StrPicTemp = "c:\temp.tmp"     '临时文件,用来保存读出的图片
     With Stm
        .Type = adTypeBinary
        .Open
        .Write Data1.Recordset.Fields(9)        '写入数据库中的数据至Stream中
        .SaveToFile StrPicTemp, adSaveCreateOverWrite   '将Stream中数据写入临时文件中
        .Close
    End With
    Image1.Picture = LoadPicture(StrPicTemp)
Else
    Image1.Picture = Nothing
End If
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = True

End Sub
