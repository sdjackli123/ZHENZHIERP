VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formr14 
   BackColor       =   &H00C0E0FF&
   Caption         =   "染化助库存信息"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10785
   LinkTopic       =   "Form14"
   ScaleHeight     =   10320
   ScaleWidth      =   10785
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入报价"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1440
      Width           =   1575
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formr14.frx":0000
      Height          =   6975
      Left            =   360
      TabIndex        =   21
      Top             =   2160
      Width           =   9975
      _cx             =   17595
      _cy             =   12303
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6360
      Top             =   9840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6240
      Top             =   9720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6360
      Top             =   9600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   6720
      Top             =   9600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Bindings        =   "Formr14.frx":0015
      Height          =   330
      Left            =   2520
      TabIndex        =   20
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "简称"
      Text            =   "DataCombo6"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Left            =   5760
      TabIndex        =   19
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo5"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Left            =   4680
      TabIndex        =   18
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   4680
      TabIndex        =   17
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formr14.frx":002A
      Height          =   330
      Left            =   360
      TabIndex        =   16
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "染化助库名"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formr14.frx":003F
      Height          =   330
      Left            =   2520
      TabIndex        =   15
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "名称"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command6 
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入库存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
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
      Height          =   480
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   141230081
      CurrentDate     =   39883
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "染化助名称"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "供应单位"
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
      Index           =   5
      Left            =   2520
      TabIndex        =   13
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "金额"
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
      Index           =   4
      Left            =   4680
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单价"
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
      Index           =   3
      Left            =   5760
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "库存数量"
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
      Left            =   4680
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "选择染化助剂库"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "库存日期"
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
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Formr14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command1_Click()
On Error Resume Next
If DataCombo1.Text = "" Or DataCombo2.Text = "" Or DataCombo3.Text = "" Or DataCombo4.Text = "" Or DataCombo5.Text = "" Then Exit Sub
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = DataCombo1.Text
Adodc1.Recordset.Fields(1) = DataCombo5.Text
Adodc1.Recordset.Fields(2) = DTPicker1.value
Adodc1.Recordset.Fields(3) = DataCombo3.Text
Adodc1.Recordset.Fields(4) = DataCombo4.Text
Adodc1.Recordset.Fields(5) = DataCombo2.Text
Adodc1.Recordset.Fields(6) = DataCombo6.Text
Adodc1.Recordset.Update
Adodc1.Refresh
DataCombo2.SetFocus
End Sub

Private Sub Command11_Click()
Call OutadodcToExcel(VSFlexGrid1, 5, "库存信息")
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
If DataCombo1.Text = "" Or DataCombo2.Text = "" Or DataCombo3.Text = "" Or DataCombo4.Text = "" Or DataCombo5.Text = "" Then Exit Sub

Adodc1.Recordset.Fields(0) = DataCombo1.Text
Adodc1.Recordset.Fields(1) = DataCombo5.Text
Adodc1.Recordset.Fields(2) = DTPicker1.value
Adodc1.Recordset.Fields(3) = DataCombo3.Text
Adodc1.Recordset.Fields(4) = DataCombo4.Text
Adodc1.Recordset.Fields(5) = DataCombo2.Text
Adodc1.Recordset.Fields(6) = DataCombo6.Text
Adodc1.Recordset.Update
Adodc1.Refresh

End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定把库存导入日期为" + Trim(DTPicker1.value), vbYesNo) = vbNo Then Exit Sub
If MsgBox("再次确认！！把库存导入日期为" + Trim(DTPicker1.value), vbYesNo) = vbNo Then Exit Sub

sql1 = "DELETE  FROM MX  WHERE 入库时间=cast('" & DTPicker1.value & "' as datetime) AND 库别='清库库存'"
sql2 = "INSERT INTO MX(供应单位,名称,入库数量,单价,染化助库名,合计金额,单据号,ip,库别,入库时间) SELECT 单位,名称,round(数量,5),round(单价,2),BL,round(金额,5),'00000000','1','清库库存','" & DTPicker1.value & "' FROM RSJL WHERE 日期=cast('" & DTPicker1.value & "' as datetime)"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("导入成功！")
End Sub


Private Sub Command5_Click()
If MsgBox("确定把库存导入报价吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("再次确认！！把库存导入报价吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM RHZH2"
sql2 = "INSERT INTO RHZH2(名称,单价,TS,IP,标志,染化助库名,简码) SELECT 名称,单价,'10','10','10',bl,'P' FROM RSJL WHERE 日期=cast('" & DTPicker1.value & "' as datetime)"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("导入成功！")
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
DataCombo3.Text = ""
DataCombo5.Text = ""
End Sub
Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub dataCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo3_Change()
DataCombo4.Text = Format(Val(DataCombo3.Text) * Val(DataCombo5.Text), "#0.00")
End Sub

Private Sub dataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub dataCombo4_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo5_Change()
DataCombo4.Text = Format(Val(DataCombo3.Text) * Val(DataCombo5.Text), "#0.00")
End Sub

Private Sub dataCombo5_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode

End Sub


Private Sub DTPicker1_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.2.254"
Adodc1.RecordSource = "SELECT 名称 AS 染化助名称,单价 AS 平均单价,日期 AS 结存日期,数量,金额,BL AS 库类,单位 FROM RSJL WHERE 日期=cast('" & DTPicker1.value & "' as datetime)"
Adodc1.Refresh
End Sub

Private Sub DTPicker1_CloseUp()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.2.254"
Adodc1.RecordSource = "SELECT 名称 AS 染化助名称,单价 AS 平均单价,日期 AS 结存日期,数量,金额,BL AS 库类,单位 FROM RSJL WHERE 日期=cast('" & DTPicker1.value & "' as datetime)"
Adodc1.Refresh
End Sub

Private Sub Form_Load()

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.2.254"
Adodc1.RecordSource = "SELECT 名称 AS 染化助名称,单价 AS 平均单价,日期 AS 结存日期,数量,金额,BL AS 库类,单位 FROM RSJL WHERE 日期=cast('" & DTPicker1.value & "' as datetime)"
Adodc1.Refresh
DTPicker1.value = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.2.254"
Set RD = New ADODB.Recordset

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.2.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.2.254"
Adodc3.RecordSource = "SELECT 简称 FROM GYS GROUP BY 简称"
Adodc3.Refresh
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.2.254"
Adodc4.RecordSource = "select 染化助库名  from RHZH GROUP BY 染化助库名 "
Adodc4.Refresh
VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(1) = 2000
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 1200
VSFlexGrid1.ColWidth(5) = 1200
VSFlexGrid1.ColWidth(6) = 1200
VSFlexGrid1.ColWidth(7) = 1200
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Label1_Click()
rhlbl = 1
Formr27.DataCombo1.Text = DataCombo2.Text
Formr27.Show
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
DataCombo1.Text = Adodc1.Recordset.Fields(0)
DataCombo2.Text = Adodc1.Recordset.Fields(5)
DataCombo3.Text = Adodc1.Recordset.Fields(3)
DataCombo5.Text = Adodc1.Recordset.Fields(1)
DataCombo6.Text = Adodc1.Recordset.Fields(6)
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub
