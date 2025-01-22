VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formh124 
   BackColor       =   &H00C0E0FF&
   Caption         =   "报价表"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   LinkTopic       =   "Form5"
   ScaleHeight     =   9975
   ScaleWidth      =   10005
   StartUpPosition =   2  '屏幕中心
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formh124.frx":0000
      Height          =   7335
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   9375
      _cx             =   16536
      _cy             =   12938
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   5640
      Top             =   9360
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
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
      Left            =   6000
      Top             =   9360
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
      Left            =   6000
      Top             =   9240
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Left            =   3600
      TabIndex        =   10
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   1440
      TabIndex        =   9
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formh124.frx":0015
      Height          =   330
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入excel"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "色号查询"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "客户查询"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   3120
      X2              =   3600
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号范围"
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
      Left            =   360
      TabIndex        =   2
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
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Formh124"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer
Private Sub Command1_Click()
Call MXOutadodcToExcel(VSFlexGrid2, "客户：" + DataCombo1.Text + "色号" + DataCombo2.Text + "-" + DataCombo5.Text)
End Sub

Private Sub Command2_Click()
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
On Error Resume Next
If S1 < 1 Or S2 < 1 Then
MsgBox ("选择审核记录")
Exit Sub
End If
If S1 > S2 Then
MsgBox ("注意选择顺序！")
Exit Sub
End If
k = S2 - S1
If k = 0 Then
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move S1 - 1
Adodc2.Recordset.Delete
Else
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move S1 - 1
For i = 1 To k + 1
Adodc2.Recordset.Delete
Adodc2.Recordset.MoveNext
Next
End If
Adodc2.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If DataCombo1.Text = "" Then
Adodc2.RecordSource = "select 客户,色号,颜色,品名,染费,负责 as 打样员,备注 from bjb where 色号 between '" & DataCombo2.Text & "' and '" & DataCombo5.Text & "' order by 色号"
Adodc2.Refresh
Else
Adodc2.RecordSource = "select 客户,色号,颜色,品名,染费,负责 as 打样员,备注 from bjb where 客户='" & DataCombo1.Text & "' and 色号 between '" & DataCombo2.Text & "' and '" & DataCombo5.Text & "' order by 色号"
Adodc2.Refresh
End If
End Sub

Private Sub Command5_Click()
Call bjd(VSFlexGrid2, "")
End Sub

Private Sub Command6_Click()
If DataCombo1.Text = "" Then
Adodc2.RecordSource = "select 客户,色号,颜色,品名,染费,负责 as 打样员,备注 from bjb order by 色号"
Adodc2.Refresh
Else
Adodc2.RecordSource = "select 客户,色号,颜色,品名,染费,负责 as 打样员,备注 from bjb where 客户='" & DataCombo1.Text & "' order by 色号"
Adodc2.Refresh
End If
End Sub

Private Sub Form_Load()
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo5.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 简称 from KHZL  group by 简称"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 客户,色号,颜色,品名,染费,负责 as 打样员,备注 from bjb where 转入<>'已' order by 色号"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid2.ColWidth(0) = 300
VSFlexGrid2.ColWidth(1) = 1600
VSFlexGrid2.ColWidth(2) = 1600
VSFlexGrid2.ColWidth(3) = 1600
VSFlexGrid2.ColWidth(4) = 1600
S1 = 1
S2 = 1
End Sub

Private Sub VSFlexGrid2_Click()
If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc2.Recordset.Move rs - 1
If DataCombo2.Text = "" Then
DataCombo2.Text = Adodc2.Recordset.Fields(1)
Else
DataCombo5.Text = Adodc2.Recordset.Fields(1)
End If
End Sub


Private Sub vSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
S1 = VSFlexGrid2.RowSel
End Sub

Private Sub vSFlexGrid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
S2 = VSFlexGrid2.RowSel
End Sub


