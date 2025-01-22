VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy302 
   BackColor       =   &H00C0E0FF&
   Caption         =   "采购类别"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form41"
   ScaleHeight     =   8355
   ScaleWidth      =   11385
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导出"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      ItemData        =   "Formy302.frx":0000
      Left            =   6960
      List            =   "Formy302.frx":000A
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "客供"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "自购"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按库类"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy302.frx":001A
      Height          =   5895
      Left            =   600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy302.frx":002E
      Height          =   330
      Left            =   5280
      TabIndex        =   5
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "材料库类"
      Text            =   "DBCombo1"
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
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Left            =   4200
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Formy302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer
Private Sub Command1_Click()
Data2.RecordSource = "SELECT 单号,款号,材料库类,材料名称,材料规格,材料单位,材料批号,材料数量 as 采购量,订单颜色 as 采购类别 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "' and  材料库类='" & DBCombo2.Text & "' order by 款号,材料库类,材料名称"
Data2.Refresh
End Sub

Private Sub Command2_Click()
Data2.RecordSource = "SELECT 单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量,订单颜色 as 采购类别 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "' order BY 单号,款号,材料库类,材料名称,材料规格,材料单位,材料颜色,材料批号"
Data2.Refresh
Data1.RecordSource = "SELECT 材料库类 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "' group BY 材料库类"
Data1.Refresh
End Sub


Private Sub Command3_Click()
If MsgBox("确定自购吗？" + DBCombo2.Text, vbYesNo) = vbNo Then Exit Sub
Data3.Database.Execute "UPDATE CGCLB SET 订单颜色='自购' WHERE 单号='" & DBCombo1.Text & "' and  材料库类='" & DBCombo2.Text & "'"
Data2.RecordSource = "SELECT 单号,款号,材料库类,材料名称,材料规格,材料单位,材料批号,材料数量 as 采购量,订单颜色 as 采购类别 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "' and  材料库类='" & DBCombo2.Text & "' order by 款号,材料库类,材料名称"
Data2.Refresh
MsgBox ("自购成功！")
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If MsgBox("确定客供吗？" + DBCombo2.Text, vbYesNo) = vbNo Then Exit Sub
Data3.Database.Execute "UPDATE CGCLB SET 订单颜色='客供' WHERE 单号='" & DBCombo1.Text & "' and  材料库类='" & DBCombo2.Text & "'"
Data2.RecordSource = "SELECT 单号,款号,材料库类,材料名称,材料规格,材料单位,材料批号,材料数量 as 采购量,订单颜色 as 采购类别 FROM CGCLB WHERE 单号='" & DBCombo1.Text & "' and  材料库类='" & DBCombo2.Text & "' order by 款号,材料库类,材料名称"
Data2.Refresh
MsgBox ("客供成功！")
End Sub

Private Sub Command6_Click()
Call cgmx(MSFlexGrid2, DBCombo2.Text)
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DBCombo2.Text = ""
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data3.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"
Data4.DatabaseName = "d:\数据库\\htgl\2011\scjd.MDB"

MSFlexGrid2.ColWidth(0) = 300
MSFlexGrid2.ColWidth(4) = 1800
End Sub

Private Sub MSFlex()
On Error Resume Next
With MSFlexGrid2
    c = .Col: r = .Row    '''''C列，，R行
    If c = 9 Then
        Combo1.Left = .Left + .ColPos(c)
        Combo1.Top = .Top + .RowPos(r)
        Combo1.Width = .ColWidth(c)
        Combo1.Height = .RowHeight(r)
        Combo1 = .Text
        Combo1.Visible = True
        Combo1.SetFocus
    End If
End With
End Sub

Private Sub Label1_dblClick(Index As Integer)
Select Case Index
       Case 1
xqbl = 1
Formy41.Show
End Select
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Combo1.Visible = False
    MSFlexGrid2.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Data2.Recordset.MoveFirst
Data2.Recordset.Move r - 1
Data2.Recordset.Edit
Data2.Recordset.Fields(c - 1) = Combo1.Text
Data2.Recordset.Update
Combo1.Visible = False
MSFlexGrid2.Text = Combo1.Text
MSFlexGrid2.SetFocus
End If
End Sub

