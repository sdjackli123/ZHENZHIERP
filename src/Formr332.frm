VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formr332 
   BackColor       =   &H00C0E0FF&
   Caption         =   "成本分析"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   9000
      Style           =   1  'Simple Combo
      TabIndex        =   27
      Text            =   "Combo1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   11280
      TabIndex        =   26
      Text            =   "Text3"
      Top             =   1080
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8160
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   7800
      Top             =   9840
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
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   10680
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   600
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formr332.frx":0000
      Height          =   330
      Left            =   11280
      TabIndex        =   22
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1095
      Left            =   12960
      TabIndex        =   18
      Top             =   360
      Width           =   1215
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "重新核算"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "核算信息"
      Height          =   1095
      Left            =   5040
      TabIndex        =   13
      Top             =   480
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "全部"
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "已核算"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "未核算"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command4 
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "结转"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "分析"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formr332.frx":0015
      Height          =   7215
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   15975
      _cx             =   28178
      _cy             =   12726
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
      FormatString    =   $"Formr332.frx":002A
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5040
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4800
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4680
      Top             =   10080
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4680
      Top             =   9840
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   12840
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   329777153
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12840
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   329777153
      CurrentDate     =   36892
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "锅号"
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   25
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "客户"
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   24
      Top             =   600
      Width           =   855
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
      Left            =   11640
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
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
      Left            =   11640
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "分析月份"
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
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Formr332"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim sdf, cbfy, rl, zj, rks, xss, xse, zzfy, qlz, ql, dgql As Double
Dim c, r As Integer
Private Sub Command1_Click()
'On Error Resume Next
If MsgBox("确定分析吗？，以前的记录被清除，请确认", vbYesNo) = vbNo Then Exit Sub

Adodc2.RecordSource = "SELECT * FROM rqsd where 月份='" & Text1 & "'"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
DTPicker1.value = Adodc2.Recordset.Fields(0)
DTPicker2.value = Adodc2.Recordset.Fields(1)
Else
MsgBox ("期间设置中没有此月份信息")
Exit Sub
End If
sql1 = "delete from cbfxb"
sql2 = "insert into cbfxb(客户,缸型,缸号,颜色,公斤,气量值,气量,单缸气量,月份,品名) select '',车台,锅号,'',数量,汽值,round(数量*cast(汽值 as real)*0.005,6),round(数量*cast(汽值 as real)*0.005*数量,6),'" & Text1 & "','' from pld where cast(CONVERT(varchar,日期, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and 信息='正常' and 锅号 not like '%-%' and len(锅号)=7 and 锅号 not in(select distinct 缸号 from XSCBFXJZ)"
sql3 = "insert into cbfxb(客户,缸型,缸号,颜色,公斤,气量值,气量,单缸气量,月份,品名) select '',车台, SUBSTRING(锅号, 1,PATINDEX('%-%', 锅号)-1) ,'',数量,汽值,round(数量*cast(汽值 as real)*0.005,6),round(数量*cast(汽值 as real)*0.005*数量,6),'" & Text1 & "','' from pld where cast(CONVERT(varchar,日期, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and 信息='正常' and 锅号 like '%-%' and len(锅号)>7 and SUBSTRING(锅号, 1,PATINDEX('%-%', 锅号)-1) not in(select distinct 缸号 from XSCBFXJZ)"
sql4 = "update cbfxb set 客户='1'"
sql5 = "insert into cbfxb(缸型,缸号,颜色,公斤,气量值,月份,品名) select 缸型,缸号,颜色,公斤,sum(气量值),月份,品名 from cbfxb group by 缸型,缸号,颜色,公斤,月份,品名"
sql6 = "delete from cbfxb where 客户='1'"
sql7 = "update v_cbfxb_qlsx set 气量值=汽值"
sql8 = "update cbfxb set 客户='',气量=round(公斤*cast(气量值 as real)*0.005,6),单缸气量=round(公斤*cast(气量值 as real)*0.005*公斤,6)"
sql9 = "delete from XSCBFX"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic
RD.Open sql7, conn, adOpenStatic, adLockOptimistic
RD.Open sql8, conn, adOpenStatic, adLockOptimistic
RD.Open sql9, conn, adOpenStatic, adLockOptimistic


Adodc3.RecordSource = "select 缸号,公斤,气量值,气量,单缸气量 from cbfxb where 气量值>0 order by 缸号"
Adodc3.Refresh


If Not Adodc3.Recordset.EOF Then
i = 1
ProgressBar1.Visible = True
sl = Adodc3.Recordset.RecordCount
ProgressBar1.value = i / sl * 100

Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
ProgressBar1.value = i / sl * 100
Adodc4.RecordSource = "select sum(isnull(汽值,0)) from pld where 锅号='" & Adodc3.Recordset.Fields(0) & "'"
Adodc4.Refresh
If Not IsNull(Adodc4.Recordset.Fields(0)) Then
qlz = Adodc4.Recordset.Fields(0)
ql = Val(Adodc3.Recordset.Fields(1)) * qlz * 0.005
dgql = Val(Adodc3.Recordset.Fields(1)) * qlz * 0.005 * Val(Adodc3.Recordset.Fields(1))
sql1 = "update cbfxb set 气量值='" & qlz & "',气量='" & ql & "',单缸气量='" & dgql & "' where 缸号='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Adodc3.Recordset.MoveNext
i = i + 1
Loop
ProgressBar1.Visible = False
End If

Adodc3.RecordSource = "select 缸型,sum(单缸气量) from cbfxb group by 缸型 order by 缸型"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
sql1 = "update cbfxb set 气量总计='" & Adodc3.Recordset.Fields(1) & "' where 缸型='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.Recordset.MoveNext
Loop
End If

Adodc3.RecordSource = "select 成本名称,成本数量 from cbfy where 成本期间='" & Text1 & "' order by 成本名称"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
sql1 = "update cbfxb set 度数='" & Adodc3.Recordset.Fields(1) & "' where 缸型='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.Recordset.MoveNext
Loop
End If

Adodc3.RecordSource = "select 车台编号,主泵 from ct  order by 车台编号"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
sql1 = "update cbfxb set 系数='" & Adodc3.Recordset.Fields(1) & "' where 缸型='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.Recordset.MoveNext
Loop
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''定型用电分配
sql1 = "delete from cbfxbdx"
sql2 = "insert into cbfxbdx(客户,缸号,公斤) select '', SUBSTRING(锅号, 1, PATINDEX('%-%', 锅号)-1),round(sum(班次产量),2) from ddcl where cast(CONVERT(varchar,时间, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and 工艺编号='定型' and 锅号 like '%-%' and len(锅号)>7 and SUBSTRING(锅号, 1, PATINDEX('%-%', 锅号)-1) not in(select distinct 缸号 from XSCBFXJZ) group by  SUBSTRING(锅号, 1, PATINDEX('%-%', 锅号)-1)"
sql3 = "insert into cbfxbdx(客户,缸号,公斤) select '', 锅号,round(sum(班次产量),2) from ddcl where cast(CONVERT(varchar,时间, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and 工艺编号='定型' and 锅号 not like '%-%' and len(锅号)=7 and 锅号 not in(select distinct 缸号 from XSCBFXJZ) group by  锅号"
sql4 = "delete from cbfxbdx where len(缸号)<>7"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic

Adodc3.RecordSource = "select sum(成本数量) from cbfy where 成本期间='" & Text1 & "' and 成本名称 like '%定型用电%'"
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.EOF) Then
sql1 = "update cbfxbdx set 定度='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Else
sql1 = "update cbfxbdx set 定度=0"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''定汽
Adodc3.RecordSource = "select sum(成本数量) from cbfy where 成本期间='" & Text1 & "' and 成本名称='定型用气'"
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.EOF) Then
sql1 = "update cbfxbdx set 定汽='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Else
sql1 = "update cbfxbdx set 定汽=0"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

sql1 = "update cbfxbdx set 定汽=0 where 定汽 is null"
sql2 = "update cbfxbdx set 定电=0 where 定电 is null"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''定型总产量
Adodc3.RecordSource = "select * from cbfxbdx"
Adodc3.Refresh
zdxsl = 0
If Not Adodc3.Recordset.EOF Then
Adodc3.RecordSource = "select round(sum(isnull(公斤,0)),2) from cbfxbdx"
Adodc3.Refresh
zdxsl = Val(Adodc3.Recordset.Fields(0))
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''定型缸分摊
sql1 = "update cbfxbdx set 定电=round(定度/'" & zdxsl & "'*公斤,6) where cast('" & zdxsl & "' as real)<>0"
sql2 = "update cbfxbdx set 定汽=round(定汽/'" & zdxsl & "'*公斤,6) where cast('" & zdxsl & "' as real)<>0"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql2 = "update cbfxb set 电量=round(单缸气量/气量总计*度数,6),水量=round(单缸气量/气量总计*度数*系数,6) where isnull(气量总计,0)<>0"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic


Adodc3.RecordSource = "select sum(成本数量) from cbfy where 成本名称='染色用气' and 成本期间='" & Text1 & "'"           '''''''''''''''工资费用
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.Fields(0)) Then
cbqyl = Adodc3.Recordset.Fields(0)
Else
cbqyl = 0
End If

Adodc4.RecordSource = "select sum(单缸气量) from cbfxb"           '''''''''''''''工资费用
Adodc4.Refresh
If Not IsNull(Adodc4.Recordset.Fields(0)) Then
qzyl = Adodc4.Recordset.Fields(0)
Else
qzyl = 0
End If
If qzyl = 0 Then
qxs = 0 '''''''''''''''''气量系数
Else
qxs = cbqyl / qzyl '''''''''''''''''气量系数
End If
Adodc3.RecordSource = "select 成本单价 from cbfy where 成本名称='气' and 成本期间='" & Text1 & "'"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
sql1 = "update cbfxb set 气费=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(单缸气量,0)*'" & qxs & "'"
sql2 = "update cbfxbdx set 汽费=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(定汽,0)"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If

Adodc3.RecordSource = "select 成本单价 from cbfy where 成本名称='水' and 成本期间='" & Text1 & "'"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
sql1 = "update cbfxb set 水费=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(水量,0)"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

Adodc3.RecordSource = "select 成本单价 from cbfy where 成本名称='电' and 成本期间='" & Text1 & "'"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
sql1 = "update cbfxb set 电费=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(电量,0)"
sql2 = "update cbfxbdx set 定费=cast('" & Adodc3.Recordset.Fields(0) & "' as real)*isnull(定电,0)"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''转入成本销售表
sql1 = "delete from XSCBFX where 月份='" & Text1 & "'"
sql2 = "insert into XSCBFX(缸号,颜色,月份,水费,电费,气费,信息,客户,品名,核算) select 缸号,颜色,月份,round(isnull(水费,0),2),round(isnull(电费,0),2),round(isnull(气费,0),2),'本月',客户,品名,'否' from cbfxb where 月份='" & Text1 & "'"
sql3 = "insert into XSCBFX(缸号,颜色,月份,水费,电费,气费,信息,客户,品名,核算) select 缸号,'','" & Text1 & "',0,定费,0,'本月',客户,'','否' from cbfxbdx"
sql4 = "insert into XSCBFX(缸号,颜色,月份,水费,电费,气费,信息,客户,品名,核算) select 缸号,'','" & Text1 & "',0,0,汽费,'本月',客户,'','否' from cbfxbdx"
sql5 = "update XSCBFX set 客户='1'"
sql6 = "insert into XSCBFX(缸号,颜色,月份,水费,电费,气费,信息,客户,品名,核算) select 缸号,颜色,月份,sum(水费),sum(电费),sum(气费),信息,'',品名,核算 from XSCBFX group by 缸号,颜色,月份,信息,品名,核算"
sql7 = "delete from XSCBFX where 客户='1'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic
RD.Open sql7, conn, adOpenStatic, adLockOptimistic


Adodc3.RecordSource = "select distinct 缸号 from XSCBFX where 月份='" & Text1 & "' order by 缸号"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
sl = Adodc3.Recordset.RecordCount
Adodc3.Recordset.MoveFirst
i = 1
Do While Not Adodc3.Recordset.EOF
ProgressBar1.Visible = True
ProgressBar1.value = i / sl * 100
Adodc4.RecordSource = "select round(sum(isnull(毛坯重量,0)),2) from jgmxkf where 锅号='" & Adodc3.Recordset.Fields(0) & "'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
rks = 0
Else
rks = Val(Adodc4.Recordset.Fields(0))
End If
Adodc4.RecordSource = "select round(sum(isnull(数量,0)),2),round(sum(isnull(金额,0)),2) from jgmx where (锅号='" & Adodc3.Recordset.Fields(0) & "' or 锅号 like '" & Adodc3.Recordset.Fields(0) & "'+'-%') and 日期<= cast('" & DTPicker2.value & "' as datetime) and 加工类别 not in('坯布费','印花费','外印花')"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
xss = 0
xse = 0
Else
xss = Val(Adodc4.Recordset.Fields(0))
xse = Val(Adodc4.Recordset.Fields(1))
End If
Adodc4.RecordSource = "select round(sum(isnull(合计金额,0)),2) from v_pld_tj_xx_hs where (锅号='" & Adodc3.Recordset.Fields(0) & "' or 锅号 like '" & Adodc3.Recordset.Fields(0) & "'+'-%') and 库类='染料库'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
rl = 0
Else
rl = Val(Adodc4.Recordset.Fields(0))
End If
Adodc4.RecordSource = "select round(sum(isnull(合计金额,0)),2) from v_pld_tj_xx_hs where (锅号='" & Adodc3.Recordset.Fields(0) & "' or 锅号 like '" & Adodc3.Recordset.Fields(0) & "'+'-%') and 库类='助剂库'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
zj = 0
Else
zj = Val(Adodc4.Recordset.Fields(0))
End If
'Adodc4.RecordSource = "select sum(定费) from cbfxbdx where 缸号='" & Adodc3.Recordset.Fields(0) & "'"
'Adodc4.Refresh
'If Not IsNull(Adodc4.Recordset.Fields(0)) Then
'df = Val(Adodc4.Recordset.Fields(0))
'Else
'df = 0
'End If
sql1 = "update XSCBFX set 入库数='" & rks & "',销售数='" & xss & "',销售额='" & xse & "',染料='" & rl & "',助剂='" & zj & "' where 月份='" & Text1 & "' and 缸号='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.Recordset.MoveNext
i = i + 1
Loop
End If


''''''''''''''''''''''''''''''''''''''''''''''''''''上月转入
If Len(Text1) = 4 Then         ''''判断期间是否正确
l1 = Mid(Text1, 1, 2)         ''''年份
L2 = Mid(Text1, 3, 2)         ''''月份
If Val(L2) = 12 Then          ''如果是最后一个月份
l1 = Val(l1) - 1              ''
L3 = Trim(l1) + "12"
Else                         ''不是最后月份
L2 = Val(L2) - 1
If Len(Trim(L2)) = 1 Then    ''不足2位  填0
L3 = l1 + "0" + Trim(L2)
Else
L3 = l1 + Trim(L2)
End If
End If
End If


sql1 = "INSERT into XSCBFX SELECT * FROM XSCBFXQM where 月份='" & L3 & "' and len(缸号)=7"
sql2 = "update XSCBFX set 月份='" & Text1 & "',核算='否' where 月份='" & L3 & "' and 信息='结转'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc4.RecordSource = "select distinct 缸号 from XSCBFX where  月份='" & Text1 & "' and 信息='结转'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
i = 1
sl = Adodc4.Recordset.RecordCount
ProgressBar1.value = i / sl * 100
Adodc4.Recordset.MoveFirst
Do While Not Adodc4.Recordset.EOF
sql1 = "update XSCBFX set 销售数=(select sum(数量) from jgmx where 锅号 like '" & Adodc4.Recordset.Fields(0) & "'+'%' and 日期<=cast('" & DTPicker2.value & "' as datetime) and 加工类别 not in('坯布费','印花费','外印花')) where 缸号='" & Adodc4.Recordset.Fields(0) & "' and 信息='结转'"
sql2 = "update XSCBFX set 销售额=(select sum(金额) from jgmx where 锅号 like '" & Adodc4.Recordset.Fields(0) & "'+'%' and 日期<=cast('" & DTPicker2.value & "' as datetime) and 加工类别 not in('坯布费','印花费','外印花')) where 缸号='" & Adodc4.Recordset.Fields(0) & "' and 信息='结转'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc4.Recordset.MoveNext
i = i + 1
Loop
End If

sql3 = "update XSCBFX set 信息='1' where 月份='" & Text1 & "'"
sql4 = "insert into XSCBFX(月份,缸号,染料,助剂,水费,电费,气费,工资,制造费用,入库数,销售数,销售额) select 月份,缸号,sum(isnull(染料,0)),sum(isnull(助剂,0)),sum(isnull(水费,0)),sum(isnull(电费,0)),sum(isnull(气费,0)),sum(isnull(工资,0)),sum(isnull(制造费用,0)),sum(isnull(入库数,0)),sum(isnull(销售数,0)),sum(isnull(销售额,0)) from XSCBFX where 月份='" & Text1 & "' group by 月份,缸号"
sql5 = "delete from XSCBFX where 信息='1' and 月份='" & Text1 & "'"
sql6 = "update  XSCBFX set 信息='本月',核算='否' where 月份='" & Text1 & "'"
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic

sql1 = "update XSCBFX set 核算='是' where 月份='" & Text1 & "' and (销售数+5)>=入库数 and 入库数>0"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

Adodc4.RecordSource = "select round(sum(isnull(销售额,0)),2) from XSCBFX where 月份='" & Text1 & "' and 核算='是'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
zse = 0
Else
zse = Val(Adodc4.Recordset.Fields(0))
End If

Adodc4.RecordSource = "select 成本费用 from cbfy where 成本名称='制造费用' and 成本期间='" & Text1 & "'"         '''''''''''''''制造费用
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
cbfy = Adodc4.Recordset.Fields(0)
Else
cbfy = 0
End If

Adodc3.RecordSource = "select 成本费用 from cbfy where 成本名称='工资' and 成本期间='" & Text1 & "'"           '''''''''''''''工资费用
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
gzfy = Adodc3.Recordset.Fields(0)
Else
gzfy = 0
End If

Adodc3.RecordSource = "select sum(isnull(水费,0)+isnull(电费,0)+isnull(气费,0)) from XSCBFX where  月份='" & Text1 & "' and 核算='是'"          ''''工资费用
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.Fields(0)) Then
sdf = Adodc3.Recordset.Fields(0)
gzxs = gzfy / sdf
sql4 = "update XSCBFX set 工资=round('" & gzxs & "'*(isnull(水费,0) + isnull(电费,0) + isnull(气费,0)),2),制造费用=round(销售额/'" & zse & "'*'" & cbfy & "',2) where 月份='" & Text1 & "' and cast('" & sdf & "' as real)<>0 and cast('" & zse & "' as real)<>0 and 核算='是'"
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
End If

sql2 = "update XSCBFX set 销售成本=round(染料+助剂+制造费用+水费+电费+气费+工资,2),销售价=round(销售额/销售数,2) where 月份='" & Text1 & "' and 核算='是' and 销售数<>0 and 入库数>0"
sql3 = "update XSCBFX set 毛利率=round((销售额-销售成本)/销售额*100,1) where 月份='" & Text1 & "' and 核算='是' and 销售成本<>0 and 销售额<>0 and 入库数>0"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic

ProgressBar1.Visible = False
Adodc3.RecordSource = "select 缸号 from XSCBFX where 月份='" & Text1 & "' order by 缸号"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
i = 1
sl = Adodc3.Recordset.RecordCount
ProgressBar1.Visible = True
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
ProgressBar1.value = i / sl * 100
Adodc4.RecordSource = "select 客户名称,品名,色别+色名 from kpd where 锅号='" & Adodc3.Recordset.Fields(0) & "' order by 重量 desc"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
sql2 = "update XSCBFX set 客户='" & Adodc4.Recordset.Fields(0) & "',品名='" & Adodc4.Recordset.Fields(1) & "',颜色='" & Adodc4.Recordset.Fields(2) & "' where 缸号='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
Adodc3.Recordset.MoveNext
i = i + 1
Loop
ProgressBar1.Visible = False
End If
Adodc1.RecordSource = "select * from XSCBFX where 月份='" & Text1 & "' order by 缸号"
Adodc1.Refresh

End Sub

Private Sub Command3_Click()

If MsgBox("请确认结转期间：" + Text1 + "正确吗?", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete from XSCBFXQM where 月份='" & L3 & "'"
sql2 = "insert into XSCBFXQM(缸号,颜色,染料,助剂,水费,电费,气费,工资,制造费用,入库数,客户,销售数,销售额,销售价,销售成本,毛利率,月份,信息,品名) select 缸号,颜色,染料,助剂,水费,电费,气费,工资,制造费用,入库数,客户,销售数,销售额,销售价,销售成本,毛利率,'" & Text1 & "','结转',品名 from XSCBFX where 核算='否' and 月份='" & Text1 & "'"
sql3 = "delete from XSCBFXJZ where 月份='" & Text1 & "'"
sql4 = "insert into XSCBFXJZ(缸号,颜色,染料,助剂,水费,电费,气费,工资,制造费用,入库数,客户,销售数,销售额,销售价,销售成本,毛利率,月份,信息,品名) select 缸号,颜色,染料,助剂,水费,电费,气费,工资,制造费用,入库数,客户,销售数,销售额,销售价,销售成本,毛利率,'" & Text1 & "','结转',品名 from XSCBFX where 核算='是' and 月份='" & Text1 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
MsgBox ("结转成功!")
End Sub


Private Sub Command4_Click()
If Option1(2).value = True Then
Adodc1.RecordSource = "select * from XSCBFX where 月份='" & Text1 & "' order by 缸号"
Adodc1.Refresh
End If
If Option1(0).value = True Then
Adodc1.RecordSource = "select * from XSCBFX where 月份='" & Text1 & "' and  核算='否' order by 缸号"
Adodc1.Refresh
End If
If Option1(1).value = True Then
Adodc1.RecordSource = "select * from XSCBFX where 月份='" & Text1 & "' and  核算='是' order by 缸号"
Adodc1.Refresh
End If
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 0, 5, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 6, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 7, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 8, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 9, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 10, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 11, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 12, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 13, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 14, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 16, , vbGreen
End Sub

Private Sub Command5_Click()
If MsgBox("确定重新核算成本吗？", vbYesNo) = vbNo Then Exit Sub
Adodc4.RecordSource = "select round(sum(isnull(销售额,0)),2) from XSCBFX where 月份='" & Text1 & "' and 核算='是'"
Adodc4.Refresh
If IsNull(Adodc4.Recordset.Fields(0)) Then
zse = 0
Else
zse = Val(Adodc4.Recordset.Fields(0))
End If

Adodc4.RecordSource = "select 成本费用 from cbfy where 成本名称='制造费用' and 成本期间='" & Text1 & "'"         '''''''''''''''制造费用
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
cbfy = Adodc4.Recordset.Fields(0)
Else
cbfy = 0
End If

Adodc3.RecordSource = "select 成本费用 from cbfy where 成本名称='工资' and 成本期间='" & Text1 & "'"           '''''''''''''''工资费用
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
gzfy = Adodc3.Recordset.Fields(0)
Else
gzfy = 0
End If

Adodc3.RecordSource = "select sum(isnull(水费,0)+isnull(电费,0)+isnull(气费,0)) from XSCBFX where  月份='" & Text1 & "' and 核算='是'"          ''''工资费用
Adodc3.Refresh
If Not IsNull(Adodc3.Recordset.Fields(0)) Then
sdf = Adodc3.Recordset.Fields(0)
gzxs = gzfy / sdf
sql4 = "update XSCBFX set 工资=round('" & gzxs & "'*(isnull(水费,0) + isnull(电费,0) + isnull(气费,0)),2),制造费用=round(销售额/'" & zse & "'*'" & cbfy & "',2) where 月份='" & Text1 & "' and cast('" & zse & "' as real)<>0 and 核算='是'"
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
End If

sql2 = "update XSCBFX set 销售成本=round(染料+助剂+制造费用+水费+电费+气费+工资,2),销售价=round(销售额/销售数,2) where 月份='" & Text1 & "' and 核算='是' and 销售数<>0"
sql3 = "update XSCBFX set 毛利率=round((销售额-销售成本)/销售额*100,1) where 月份='" & Text1 & "' and  销售额<>0 and 核算='是'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
MsgBox ("重新分析成功！")

End Sub

Private Sub Command6_Click()
Call MXCBFX(VSFlexGrid1, "成本分析")
End Sub

Private Sub Command7_Click()
sql1 = ""
If Check2(1).value = 1 Then
sql1 = sql1 + "客户 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "缸号 like '%'+'" & Text3 & "'+'%' and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "select * from XSCBFX where (" + sql1 + ") and 月份='" & Text1 & "'"
Adodc1.Refresh

End Sub

Private Sub Text1_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from XSCBFX where 月份='" & Text1 & "' order by 缸号"
Adodc1.Refresh
End Sub

Private Sub Form_Load()

On Error Resume Next
Text1 = ""
Text2 = ""
Text3 = ""
DataCombo1 = ""
Option1(2).value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from XSCBFX where 月份='" & Text1 & "'"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(1) = 1500
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text2_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text2 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc5.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
End With
If r = 2 Then
Formc23.DataCombo4 = VSFlexGrid1.TextMatrix(r, 2)
Formc23.Check2(7).value = 1
Formc23.Show
End If

If c = 5 Or c = 6 Then
Formr309.Text1 = VSFlexGrid1.TextMatrix(r, 2)
Formr309.Show
End If

If c = 8 Then
Formr307.Text1 = VSFlexGrid1.TextMatrix(r, 2)
Formr307.Show
End If

If c = 9 Then
Formr308.Text1 = VSFlexGrid1.TextMatrix(r, 2)
Formr308.Show
End If

End Sub

Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
End With

S2 = VSFlexGrid1.TextMatrix(r, 2)   '''缸号

    If Button = 2 And c = 2 Then
    If MsgBox("确定核算这行的信息吗？" + S2, vbYesNo) = vbNo Then  '''PopupMenu mnu_manager  '这是在窗体中设置的一个顶级菜单名称
    Exit Sub
    Else
    sql2 = "update XSCBFX set 核算='是' where 缸号='" & S2 & "' and 月份='" & Text1 & "'"
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic
    End If
    Call Command4_Click
    End If

    If Button = 2 And c = 1 Then
    If MsgBox("确定取消核算这行的信息吗？" + S2, vbYesNo) = vbNo Then  '''PopupMenu mnu_manager  '这是在窗体中设置的一个顶级菜单名称
    Exit Sub
    Else
    sql2 = "update XSCBFX set 核算='否' where 缸号='" & S2 & "' and 月份='" & Text1 & "'"
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic
    End If
    Call Command4_Click
    End If

End Sub


Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
    If c = 12 Or c = 13 Or c = 14 Then
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
    End If
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
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

Adodc1.Recordset.Fields(c - 1) = Combo1111.Text
Adodc1.Recordset.Update

    VSFlexGrid1.Text = Combo1111.Text
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub

