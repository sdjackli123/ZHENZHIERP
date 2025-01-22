VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw122 
   BackColor       =   &H00C0E0FF&
   Caption         =   "费用明细"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4680
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1560
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   7800
      Top             =   10560
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "Adodc6"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   7680
      Top             =   10320
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
      Left            =   8040
      Top             =   10440
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
      Left            =   8040
      Top             =   10320
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Left            =   8160
      Top             =   10560
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
      Left            =   8160
      Top             =   10320
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw122.frx":0000
      Height          =   7095
      Left            =   720
      TabIndex        =   16
      Top             =   3000
      Width           =   13935
      _cx             =   24580
      _cy             =   12515
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw122.frx":0015
      Height          =   330
      Left            =   5280
      TabIndex        =   15
      Top             =   1560
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Formw122.frx":002A
      Left            =   5280
      List            =   "Formw122.frx":0034
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "客户查询"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "类别查询"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "凭证生成"
      Height          =   375
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成查询"
      Height          =   375
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   423821313
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   11640
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   423821313
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   423821313
      CurrentDate     =   36892
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期："
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
      Index           =   0
      Left            =   840
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期："
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
      Index           =   12
      Left            =   840
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "类别"
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
      Left            =   4080
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      Height          =   375
      Index           =   0
      Left            =   11640
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户"
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
      Left            =   4080
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "Formw122"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Private Sub Combo1_Click()
If Combo1.Text = "应付类" Then
Adodc3.RecordSource = "select 简称 from gys group by 简称 order by 简称"
Adodc3.Refresh
End If
If Combo1.Text = "应收类" Then
Adodc3.RecordSource = "select 简称 from khzl group by 简称 order by 简称"
Adodc3.Refresh
End If
End Sub

Private Sub Command1_Click()
If MsgBox("操作日期为：" + Trim(DTPicker3.value) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("操作期间为：" + Trim(Month(DTPicker3.value)) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定生成费用系列的凭证吗？", vbYesNo) = vbNo Then Exit Sub
Call CPFYPZ(DTPicker2.value, DTPicker3.value, DTPicker1.value)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Combo1.Text = "" Then
Adodc1.RecordSource = "select * from zxbz where 日期 between '" & DTPicker2.value & "' and '" & DTPicker3.value & "' order by 日期,序号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select * from zxbz where 日期 between '" & DTPicker2.value & "' and '" & DTPicker3.value & "' and 类别='" & Combo1.Text & "' order by 日期,序号"
Adodc1.Refresh
End If
End Sub

Private Sub Command4_Click()
Formw1132.DTPicker1.value = DTPicker2.value
Formw1132.Show
End Sub

Private Sub Command5_Click()
Call OutadodcToExcel(VSFlexGrid1, 2, Combo1.Text + "费用查询 日期范围： " + Trim(DTPicker2.value) + "--" + Trim(DTPicker3.value))
End Sub

Private Sub Command6_Click()
Adodc1.RecordSource = "select * from zxbz where 日期 between '" & DTPicker2.value & "' and '" & DTPicker3.value & "' and 客户 like '%'+'" & DataCombo1.Text & "'+'%' order by 日期,序号"
Adodc1.Refresh
End Sub


Private Sub Form_Load()
On Error Resume Next
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
cdbhf = cdbh

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM ZXBZ WHERE 日期 BETWEEN '" & DTPicker2.value & "' AND '" & DTPicker3.value & "' ORDER BY 序号 DESC"
Adodc1.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Combo1.Text = ""
DataCombo1.Text = ""
Text1 = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 1500
VSFlexGrid1.ColWidth(5) = 1500
VSFlexGrid1.ColWidth(8) = 1500
VSFlexGrid1.ColWidth(13) = 0
End Sub

Private Sub CPFYPZ(DT1 As Date, dt2 As Date, dt3 As Date)
'On Error Resume Next

If Combo1.Text = "应收类" Then
Dim djs As Integer

Adodc4.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "' and 制单 like '%自动-应收%'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
If MsgBox("已有应收生成凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CLZZPZ WHERE 制单 like '%自动-应收%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If


Adodc6.RecordSource = "SELECT * FROM JGZCX where 本期应收款<>0 order by 客户"
Adodc6.Refresh

If Adodc6.Recordset.EOF Then Exit Sub

Adodc5.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
If Adodc5.Recordset.EOF Then
PZH = "5-1"
Else
Adodc5.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-2)) FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
PZH = "5-" + Trim(Adodc5.Recordset.Fields(0) + 1)
End If
Adodc6.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc6.Recordset.EOF
For i = 1 To 7

sql1 = "INSERT INTO CLZZPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('应收','应收账款','" & Adodc6.Recordset.Fields(6) & "'),'" & Data5.Recordset.Fields(0) & "','" & PZH & "','" & dt3 & "','','','','自动-应收')"
sql2 = "INSERT INTO CLZZPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('应收','主营业务收入','','','" & Adodc6.Recordset.Fields(2) & "','" & PZH & "','" & dt3 & "','','','','自动-应收')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc6.Recordset.MoveNext
If Adodc6.Recordset.EOF Then
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
Adodc5.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
If Adodc5.Recordset.EOF Then
PZH = "5-1"
Else
Adodc5.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-2)) FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
PZH = "5-" + Trim(Adodc5.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("应收类转账成功！" + "生成" + Str(KLLLL) + "凭证")
End If

If Combo1.Text = "应付类" Then
Adodc4.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "' and 制单 like '%自动-应付%'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
If MsgBox("已有应付生成凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CLZZPZ WHERE 制单 like '%自动-应付%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

Adodc6.RecordSource = "SELECT * FROM JGZCX1 where 本期应付款<>0 order by 客户"
Adodc6.Refresh

If Adodc6.Recordset.EOF Then Exit Sub

Adodc5.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
If Adodc5.Recordset.EOF Then
PZH = "5-1"
Else
Adodc5.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-2)) FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
PZH = "5-" + Trim(Adodc5.Recordset.Fields(0) + 1)
End If
Adodc6.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc6.Recordset.EOF
For i = 1 To 7

sql1 = "INSERT INTO CLZZPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('应付','制造费用','','" & Adodc6.Recordset.Fields(2) & "','','" & PZH & "','" & dt3 & "','','','','应付自动','','','')"
sql2 = "INSERT INTO CLZZPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('应付','应付账款','" & Adodc6.Recordset.Fields(0) & "','','" & Adodc6.Recordset.Fields(2) & "','" & PZH & "','" & dt3 & "','','','','应付自动','','','')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc6.Recordset.MoveNext
If Adodc6.Recordset.EOF Then
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
Adodc5.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
If Adodc5.Recordset.EOF Then
PZH = "5-1"
Else
Adodc5.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-2)) FROM CLZZPZ WHERE 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
PZH = "5-" + Trim(Adodc5.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("应付类转账成功！" + "生成" + Str(KLLLL) + "凭证")
End If
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

Private Sub Text1_Change()
If Combo1.Text = "应收类" Then
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select  简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%' group by 简称 "
Adodc3.Refresh
End If
If Combo1.Text = "应收类" Then
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select  简称 from gys where 代码  like '%'+'" & Text1 & "'+'%' group by 简称 "
Adodc3.Refresh
End If

End Sub
