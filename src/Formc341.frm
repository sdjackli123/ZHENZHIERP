VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formc341 
   BackColor       =   &H00C0E0FF&
   Caption         =   "审核确认"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   14805
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "反结"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1920
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   2040
      Top             =   9000
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   1800
      Top             =   9000
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.ComboBox Combo1111 
      Appearance      =   0  'Flat
      Height          =   300
      ItemData        =   "Formc341.frx":0000
      Left            =   6960
      List            =   "Formc341.frx":0002
      Style           =   1  'Simple Combo
      TabIndex        =   22
      Text            =   "Combo1111"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1095
      Left            =   4080
      TabIndex        =   14
      Top             =   840
      Width           =   3255
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单据"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "已审"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "未审"
         Height          =   255
         Index           =   7
         Left            =   2160
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc341.frx":0004
      Height          =   330
      Left            =   1800
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "锅号"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   480
      Top             =   9120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Left            =   240
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Height          =   330
      Left            =   480
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "审核"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "反审"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   855
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formc341.frx":0019
      Height          =   3615
      Left            =   840
      TabIndex        =   5
      Top             =   3480
      Width           =   13575
      _cx             =   23945
      _cy             =   6376
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
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   327876609
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   327876609
      CurrentDate     =   36892
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formc341.frx":002E
      Height          =   330
      Left            =   1800
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "锅号"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formc341.frx":0043
      Height          =   735
      Left            =   8400
      TabIndex        =   20
      Top             =   7080
      Width           =   6015
      _cx             =   10610
      _cy             =   1296
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
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "输入单据"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光坯发货表:"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "输入锅号"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
End
Attribute VB_Name = "Formc341"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim sql As String
Dim c, r As Integer
Dim cdbhf As Integer
Private Sub Command1_Click()
On Error Resume Next
If DataCombo2 = "" Then Exit Sub
If MsgBox("确定审核吗？  " + DataCombo2, vbYesNo) = vbNo Then Exit Sub
sql1 = "UPDATE JGMX SET 审核='是' WHERE 单号='" & DataCombo2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
End Sub
Private Sub Command2_Click()
On Error Resume Next
If DataCombo2 = "" Then Exit Sub
If MsgBox("确定反审吗？  " + DataCombo2, vbYesNo) = vbNo Then Exit Sub
sql1 = "UPDATE JGMX SET 审核='否' WHERE 单号='" & DataCombo2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
End Sub


Private Sub Command3_Click()
If DataCombo2 = "" Then
MsgBox ("请输入单据号")
Exit Sub
End If
Call CPCKSH(Adodc3, DataCombo2)
End Sub

Private Sub Command4_Click()
Unload Me
End Sub


Private Sub Command5_Click()
On Error Resume Next
If DataCombo2 = "" Then Exit Sub
If MsgBox("确定反结吗？  " + DataCombo2, vbYesNo) = vbNo Then Exit Sub
sql1 = "UPDATE JGMX SET 成分='' WHERE 单号='" & DataCombo2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
sql = ""

If Check2(2).value = 1 Then
sql = sql + "isnull(审核,'')='是' and "
End If

If Check2(6).value = 1 Then
sql = sql + "锅号 like '%'+'" & DataCombo1 & "'+'%' and "
End If

If Check2(4).value = 1 Then
t1 = Format(Trim(DTPicker3.value), "yyyy-mm-dd")
t2 = Format(Trim(DTPicker4.value), "yyyy-mm-dd")
sql = sql + "CONVERT(varchar,日期, 23) between '" & t1 & "' and '" & t2 & "' and "
End If

If Check2(5).value = 1 Then
sql = sql + "单号 like '%'+'" & DataCombo2 & "'+'%' and "
End If

If Check2(7).value = 1 Then
sql = sql + "isnull(审核,'')<>'是' and "
End If

If sql = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql = Left$(Trim(sql), Len(Trim(sql)) - 4)
Adodc1.RecordSource = "select 审核,加工单位 as 客户名称,品名,成分,颜色,锅号,匹数,数量 as 毛坯重量,光坯 as 光坯重量,米数 as 光坯米数,单价,金额,日期,单号,备注,计划号 as 合同编号,和约号 as 客户合同号,核算,成分 as 结账 from jgmx where (" + sql + ") order by 日期,单号"
Adodc1.Refresh
Adodc2.RecordSource = "select sum(匹数) as 合计匹数,round(sum(数量),2) as 合计毛坯,round(sum(光坯),2) as 合计光坯,round(sum(金额),2) as 合计金额 from jgmx where (" + sql + ")"
Adodc2.Refresh
End Sub


Private Sub Form_Load()
DataCombo1.Text = ""
DataCombo2.Text = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
DTPicker3.value = Date
DTPicker4.value = Date
cdbhf = cdbh
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select distinct 部门,姓名 from shbm"
Adodc4.Refresh
VSFlexGrid1.ColWidth(0) = 600
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 1300
VSFlexGrid1.ColWidth(5) = 1600
End Sub

Private Sub MSF()
On Error Resume Next
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
If .Text <> "审核" Then
If c = 7 Or c = 8 Or c = 9 Or c = 10 Or c = 11 Or c = 18 Then
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
End If
End If
End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
sql2 = "insert into yhcd(用户,菜单,编号) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.WindowState = 2
Formm1.Adodc1.Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sql2 = "delete from yhcd where 用户='" & yhm & "' and 编号='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.Adodc1.Refresh
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
'Call MSF
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move r - 1
'm = Adodc1.Recordset.Fields(c - 2)
Adodc1.Recordset.Fields(c - 1) = Combo1111.Text
If (Adodc1.Recordset.Fields(17) = "毛坯" And c = 8) Then
Adodc1.Recordset.Fields(11) = Format(Adodc1.Recordset.Fields(10) * Val(Combo1111), "#0.00")
End If
If (Adodc1.Recordset.Fields(17) = "光坯" And c = 9) Then
Adodc1.Recordset.Fields(11) = Format(Adodc1.Recordset.Fields(10) * Val(Combo1111), "#0.00")
End If
If (Adodc1.Recordset.Fields(17) = "米数" And c = 10) Then
Adodc1.Recordset.Fields(11) = Format(Adodc1.Recordset.Fields(10) * Val(Combo1111), "#0.00")
End If

If c = 11 Then
If Adodc1.Recordset.Fields(17) = "毛坯" Then
Adodc1.Recordset.Fields(11) = Format(Adodc1.Recordset.Fields(7) * Val(Combo1111), "#0.00")
End If
If Adodc1.Recordset.Fields(17) = "光坯" Then
Adodc1.Recordset.Fields(11) = Format(Adodc1.Recordset.Fields(8) * Val(Combo1111), "#0.00")
End If
If Adodc1.Recordset.Fields(17) = "米数" Then
Adodc1.Recordset.Fields(11) = Format(Adodc1.Recordset.Fields(9) * Val(Combo1111), "#0.00")
End If
End If




'If c = 10 Then
'Adodc1.Recordset.Fields(c) = Format(Adodc1.Recordset.Fields(c - 2) * Val(Combo1111.Text), "#0.00")
'End If
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc2.Refresh
'    VSFlexGrid1.Text = Combo1111.Text
'    If c = 10 Then
'    VSFlexGrid1.TextMatrix(r, c + 1) = Format(m * Val(Combo1111.Text), "#0.00")
'    End If
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub



