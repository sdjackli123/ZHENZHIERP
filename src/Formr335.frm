VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formr335 
   BackColor       =   &H00C0E0FF&
   Caption         =   "用料统计"
   ClientHeight    =   10215
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1200
      Top             =   9840
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "材料刷新"
      Height          =   375
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1080
      Top             =   9960
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formr335.frx":0000
      Height          =   330
      Left            =   7080
      TabIndex        =   11
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "染化助库名"
      Text            =   "DataCombo1"
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Formr335.frx":0015
      Left            =   5280
      List            =   "Formr335.frx":001F
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   600
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1320
      Top             =   9840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   13560
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库类刷新"
      Height          =   375
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330956801
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   330956801
      CurrentDate     =   36892
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formr335.frx":002F
      Height          =   8055
      Left            =   960
      TabIndex        =   8
      Top             =   1440
      Width           =   14295
      _cx             =   25215
      _cy             =   14208
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formr335.frx":0044
      Height          =   330
      Left            =   9120
      TabIndex        =   13
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "染料名称"
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   9120
      TabIndex        =   14
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "库类"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Formr335"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")

If Combo1.Text = "手工" Then
If DataCombo2.Text = "" Then
Adodc1.RecordSource = "select 名称,单位,round(sum(数量),5) as 统计量 from sgtj where CONVERT(varchar(120),日期,23) between '" & t1 & "' and '" & t2 & "' group by 名称,单位 order by 名称"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select 名称,单位,round(sum(数量),5) as 统计量 from sgtj where 名称='" & DataCombo2.Text & "' and CONVERT(varchar(120),日期,23) between '" & t1 & "' and '" & t2 & "' group by 名称,单位 order by 名称"
Adodc1.Refresh
End If
End If

If Combo1.Text = "配料" Then
If DataCombo2.Text = "" Then
Adodc1.RecordSource = "select 染化助名称,配料单位,round(sum(配料用量),5) as 配料量 from pldb where CONVERT(varchar(120),配料日期,23) between '" & t1 & "' and '" & t2 & "' group by 染化助名称,配料单位 order by 染化助名称"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select 染化助名称,配料单位,round(sum(配料用量),5) as 配料量 from pldb where 染化助名称='" & DataCombo2.Text & "' and CONVERT(varchar(120),配料日期,23) between '" & t1 & "' and '" & t2 & "' group by 染化助名称,配料单位 order by 染化助名称"
Adodc1.Refresh
End If
End If
End Sub

Private Sub Command2_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "用料统计")
End Sub

Private Sub Command3_Click()
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")

If Combo1.Text = "手工" Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "select 名称,单位,round(sum(数量),5) as 统计量 from sgtj where CONVERT(varchar(120),日期,23) between '" & t1 & "' and '" & t2 & "' group by 名称,单位 order by 名称"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select 名称,单位,round(sum(数量),5) as 统计量 from sgtj where 库类='" & DataCombo1.Text & "' and CONVERT(varchar(120),日期,23) between '" & t1 & "' and '" & t2 & "' group by 名称,单位 order by 名称"
Adodc1.Refresh
End If
End If

If Combo1.Text = "配料" Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "select 染化助名称,配料单位,round(sum(配料用量),5) as 配料量 from pldb where CONVERT(varchar(120),配料日期,23) between '" & t1 & "' and '" & t2 & "' group by 染化助名称,配料单位 order by 染化助名称"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select 染化助名称,配料单位,round(sum(配料用量),5) as 配料量 from pldb where 染化助库='" & DataCombo1.Text & "' and CONVERT(varchar(120),配料日期,23) between '" & t1 & "' and '" & t2 & "' group by 染化助名称,配料单位 order by 染化助名称"
Adodc1.Refresh
End If
End If

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
DTPicker1.value = Date
DTPicker2.value = Date
Combo1.Text = ""
DataCombo1.Text = ""
DataCombo2.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 染化助库名 from RHZH group by 染化助库名"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT 染料名称 from RHZH group by 染料名称"
Adodc3.Refresh
End Sub
