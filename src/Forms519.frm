VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms519 
   BackColor       =   &H00C0E0FF&
   Caption         =   "班次能耗登记"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   15735
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   10
      Left            =   11160
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   9
      Left            =   9120
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   8
      Left            =   9120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   7
      Left            =   7080
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   6
      Left            =   7080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   5
      Left            =   5040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   4
      Left            =   5040
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   3
      Left            =   2880
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   9480
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms519.frx":0000
      Height          =   3855
      Left            =   720
      TabIndex        =   0
      Top             =   4200
      Width           =   14295
      _cx             =   25215
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7200
      Top             =   7320
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   720
      TabIndex        =   17
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   330039297
      CurrentDate     =   39903
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   720
      TabIndex        =   33
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   330039297
      CurrentDate     =   39903
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   495
      Left            =   3240
      TabIndex        =   34
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   330039297
      CurrentDate     =   39903
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
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
      Index           =   12
      Left            =   3240
      TabIndex        =   32
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
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
      Index           =   11
      Left            =   720
      TabIndex        =   31
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "班次"
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
      Index           =   8
      Left            =   11160
      TabIndex        =   29
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "选取"
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
      Left            =   4200
      TabIndex        =   28
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "用电"
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
      Index           =   10
      Left            =   9120
      TabIndex        =   27
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "用汽"
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
      Index           =   9
      Left            =   9120
      TabIndex        =   26
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "用水"
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
      Index           =   7
      Left            =   7080
      TabIndex        =   25
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "产量"
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
      Index           =   6
      Left            =   7080
      TabIndex        =   24
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "修色"
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
      Left            =   5040
      TabIndex        =   23
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "漂白"
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
      Left            =   2880
      TabIndex        =   22
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "化纤"
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
      Left            =   2880
      TabIndex        =   21
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "染棉"
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
      Left            =   720
      TabIndex        =   20
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "落水"
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
      Left            =   5040
      TabIndex        =   19
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "日期"
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
      Left            =   720
      TabIndex        =   18
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Forms519"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command12_Click()
Adodc1.RecordSource = "select  * from cjnhfx where 日期 between cast('" & DTPicker2.value & "' as datetime) and cast('" & DTPicker3.value & "' as datetime) order by 日期,班次"
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
Call OutadodcToExcel(VSFlexGrid1, 8, "能耗表")
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Text1(0).Text = DTPicker1.value
End Sub

Private Sub DTPicker1_CloseUp()
Text1(0).Text = DTPicker1.value
End Sub

Private Sub Form_Load()
For i = 0 To 10
Text1(i).Text = ""
Next
Text1(0).Text = Date
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select  * from cjnhfx where 日期=cast('" & DTPicker1.value & "' as datetime) order by 日期,班次"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1000
VSFlexGrid1.ColWidth(3) = 1000
VSFlexGrid1.ColWidth(4) = 1000
VSFlexGrid1.ColWidth(5) = 1000
VSFlexGrid1.ColWidth(6) = 1000
VSFlexGrid1.ColWidth(7) = 1000
VSFlexGrid1.ColWidth(8) = 1000
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 1 To 10
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定修改 报表日期" + Text1(0) + "班次" + Text1(10) + "的能耗信息吗？", vbYesNo) = vbNo Then Exit Sub
If Text1(10) = "" Or Val(Text1(6)) = 0 Or Val(Text1(8)) = 0 Then
MsgBox ("输入的信息 不正确")
Exit Sub
End If
For i = 1 To 9
Text1(i) = Val(Text1(i))
Next
For i = 0 To 10
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
For i = 1 To 10
Text1(i).Text = ""
Next
End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除 报表日期" + Text1(0) + "班次" + Text1(10) + "的能耗信息吗？", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
For i = 1 To 10
Text1(i).Text = ""
Next
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Text1(10) = "" Or Val(Text1(6)) = 0 Or Val(Text1(8)) = 0 Then
MsgBox ("输入的信息 不正确")
Exit Sub
End If
For i = 1 To 9
Text1(i) = Val(Text1(i))
Next
Adodc1.Recordset.AddNew
For i = 0 To 10
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
For i = 1 To 10
Text1(i).Text = ""
Next
End Sub


