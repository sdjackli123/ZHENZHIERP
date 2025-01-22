VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw1241 
   BackColor       =   &H00C0E0FF&
   Caption         =   "报价表"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form5"
   ScaleHeight     =   10755
   ScaleWidth      =   12000
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   5280
      Top             =   9480
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
      Left            =   5640
      Top             =   9720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      Left            =   5760
      Top             =   9240
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
      Left            =   5640
      Top             =   9480
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Height          =   3855
      Left            =   360
      TabIndex        =   20
      Top             =   3960
      Width           =   11655
      _cx             =   20558
      _cy             =   6800
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
      AllowUserResizing=   0
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
      AutoSizeMode    =   0
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
      WordWrap        =   0   'False
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   3480
      TabIndex        =   19
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   1320
      TabIndex        =   18
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Left            =   1320
      TabIndex        =   17
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转入内"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转入外"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "未转查询"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转入查询"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "色号查询"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   855
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
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
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "客户查询"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw1241.frx":0000
      Height          =   330
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "简称"
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formw1241.frx":0014
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   6
      BackColorFixed  =   9671679
      BackColorBkg    =   37779
      AllowUserResizing=   3
      FormatString    =   "记录号"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Left            =   4920
      TabIndex        =   14
      Top             =   960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo2"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   15
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择客户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Formw1241"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer
Private Sub Command1_Click()
Call MXOutadodcToExcel(VSFlexGrid2, "客户：" + DataCombo1.Text + "色号" + DataCombo2.Text + "-" + DataCombo5.Text)
End Sub

Private Sub Command2_Click()
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
If DataCombo1.Text = "" Or DataCombo2.Text = "" Or DataCombo3.Text = "" Then
MsgBox ("请输入完整！")
Exit Sub
End If
Adodc3.Database.Execute "delete * from bjb where 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
Adodc2.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If DataCombo1.Text = "" Then
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where instr(色号,'" & DataCombo2.Text & "')>0 order by 色号 desc"
Adodc2.Refresh
Else
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where 客户='" & DataCombo1.Text & "' and instr(色号,'" & DataCombo2.Text & "')>0 order by 色号 desc"
Adodc2.Refresh
End If
End Sub

Private Sub Command5_Click()
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where 转入='已' order by 色号 desc"
Adodc2.Refresh
End Sub

Private Sub Command6_Click()
If DataCombo1.Text = "" Then
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb order by 色号 desc"
Adodc2.Refresh
Else
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where 客户='" & DataCombo1.Text & "' order by 色号 desc"
Adodc2.Refresh
End If
End Sub

Private Sub Command7_Click()
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where 转入='未' or 转入=null order by 色号 desc"
Adodc2.Refresh
End Sub

Private Sub Command8_Click()
If MsgBox("确定转入吗？", vbYesNo) = vbNo Then Exit Sub
If DataCombo1.Text = "" Or DataCombo2.Text = "" Or DataCombo3.Text = "" Then
MsgBox ("请输入完整！")
Exit Sub
End If
lo = "d:\数据库\bfrz\2012\wx\khjg.mdb"
Adodc4.Database.Execute "delete * from dj  WHERE 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
Adodc3.Database.Execute "INSERT INTO DJ(客户,品名,色号,色别,单价,备注) in'" & lo & "' SELECT 客户,品名,色号,颜色,val(报价),备注 FROM bjb WHERE 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
Adodc3.Database.Execute "update bjb set 转入='已' WHERE 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
Adodc4.Database.Execute "update dj set ip='1' where ip=null"
MsgBox ("转入成功!在色别单价中查询")
End Sub

Private Sub Command9_Click()
If MsgBox("确定转入吗？", vbYesNo) = vbNo Then Exit Sub
If DataCombo1.Text = "" Or DataCombo2.Text = "" Or DataCombo3.Text = "" Then
MsgBox ("请输入完整！")
Exit Sub
End If
lo = "d:\数据库\bfrz\2012\nx\khjg.mdb"
Adodc4.Database.Execute "delete * from dj  WHERE 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
Adodc3.Database.Execute "INSERT INTO DJ(客户,品名,色号,色别,单价,备注) in'" & lo & "' SELECT 客户,品名,色号,颜色,val(报价),备注 FROM bjb WHERE 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
Adodc3.Database.Execute "update bjb set 转入='已' WHERE 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
Adodc4.Database.Execute "update dj set ip='1' where ip=null"
MsgBox ("转入成功!在色别单价中查询")
End Sub

Private Sub Form_Load()
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc1.RecordSource = "select 简称 from KHZL  group by 简称"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc2.RecordSource = "select * from bjb order by 色号 desc"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fyrj;Persist Security Info=True;User ID=fyrj;Initial Catalog=zzpr;Data Source=fydnrj"
VSFlexGrid2.ColWidth(0) = 300
VSFlexGrid2.ColWidth(1) = 1600
VSFlexGrid2.ColWidth(2) = 1600
VSFlexGrid2.ColWidth(3) = 1600
VSFlexGrid2.ColWidth(4) = 1600
End Sub

Private Sub vSFlexGrid2_dblClick()
If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc2.Recordset.Move rs - 1
DataCombo1.Text = Adodc2.Recordset.Fields(0)
DataCombo2.Text = Adodc2.Recordset.Fields(2)
DataCombo3.Text = Adodc2.Recordset.Fields(1)
End Sub


Private Sub MSFlex()
With VSFlexGrid2
    c = .Col: r = .Row    '''''C列，，R行
    If c = 6 Or c = 8 Then
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
    End If
End With
End Sub


Private Sub vSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid2.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move r - 1
Adodc2.Recordset.Edit
Adodc2.Recordset.Fields(c - 1) = Text1111.Text
Adodc2.Recordset.Update
VSFlexGrid2.Text = Text1111.Text
Text1111.Visible = False
VSFlexGrid2.SetFocus
End If
End Sub

