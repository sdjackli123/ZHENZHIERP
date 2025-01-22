VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formw93 
   BackColor       =   &H00C0E0FF&
   Caption         =   "扫描入库"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   960
      Width           =   1335
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
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   2895
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
      Height          =   855
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   5775
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw93.frx":0000
      Height          =   7695
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   13573
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label4 
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
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   1335
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
      Height          =   495
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Formw93"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DBCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Form_Load()
Dim l As Integer
Text1.Text = ""
Text3.Text = ""

m = ""
Data4.DatabaseName = "d:\数据库\\htgl\2011\SCjd.mdb"
Data4.RecordSource = "CPK"
Data4.Refresh

Data1.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data1.RecordSource = "select MC from CLDW GROUP BY MC"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\\htgl\2011\CPCK.MDB"
Data2.RecordSource = "select * from lsrk where 条码='" & m & "' and 日期=cdate('" & Date & "')"
Data2.Refresh

Data5.DatabaseName = "d:\数据库\\htgl\2011\SCJD.mdb"

Data6.DatabaseName = "d:\数据库\\htgl\2011\cpck.mdb"

Data3.DatabaseName = "d:\数据库\\htgl\2011\cpck.mdb"

MSFlexGrid1.ColWidth(11) = 1200
MSFlexGrid1.ColWidth(10) = 1200

End Sub

Private Sub Label4_Click()
Data6.RecordSource = "SELECT * FROM lsrk WHERE 日期=cdate('" & Date & "')"
Data6.Refresh
If Not Data6.Recordset.EOF Then
Data6.RecordSource = "select max(mid(发货地,7)) from lsrk where 日期=cdate('" & Date & "')"
Data6.Refresh
If Len(Data6.Recordset.Fields(0) + 1) < 2 Then
Text3.Text = "R" + Format(Date, "mmdd") + "-" + "0" + Trim(Data6.Recordset.Fields(0) + 1)
Else
Text3.Text = "R" + Format(Date, "mmdd") + "-" + Trim(Data6.Recordset.Fields(0) + 1)
End If
Else
Text3.Text = "R" + Format(Date, "mmdd") + "-" + "01"
End If
End Sub

Private Sub Text1_Change()
If Text3.Text = "" Then Exit Sub

If InStr(Text1.Text, "J") > 0 Then
m = Left(Text1.Text, Len(Text1.Text) - 1)

If Len(m) = 9 Then

Data4.Recordset.FindFirst "卡号='" & m & "'"
If Data4.Recordset.NoMatch Then
Label2.Caption = "不存在此条码"
Text1.Text = ""
Timer1.Enabled = True
Exit Sub
Else
Data6.RecordSource = "SELECT * FROM LSRK WHERE 条码='" & m & "'"
Data6.Refresh
If Data6.Recordset.EOF Then

l = 1
Data3.RecordSource = "SELECT 序号 FROM LSRK WHERE 日期=CDATE('" & Date & "') ORDER BY 序号 DESC"
Data3.Refresh
If Data3.Recordset.EOF Then
l = 1
Else
l = Data3.Recordset.Fields(0) + 1
End If

Data5.Database.Execute "INSERT INTO lsrk(日期,单号,款号,品名,规格,型号,单位,数量,备注,条码,序号) in'd:\数据库\\htgl\2011\cpck.mdb' select distinct CDATE('" & Date & "'),单号,款号,品名,颜色,规格,单位,数量,'" & Text3.Text & "',卡号,'" & l & "' from cpk where 卡号='" & m & "' and 日期=cdate('" & Date & "')"
End If
Data2.RecordSource = "select * from lsrk where 条码='" & m & "'"
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

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub
