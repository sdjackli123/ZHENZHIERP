VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formc146 
   BackColor       =   &H00C0E0FF&
   Caption         =   "成品信息"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1455
      Left            =   6240
      TabIndex        =   13
      Top             =   240
      Width           =   3015
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "款号"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   240
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3960
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全部"
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
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选择"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3840
      Top             =   10080
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      Left            =   4440
      Top             =   9960
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
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
      Height          =   375
      Left            =   5040
      Top             =   9840
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formc146.frx":0000
      Height          =   7935
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   17175
      _cx             =   30295
      _cy             =   13996
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
      FormatString    =   $"Formc146.frx":0015
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
      Editable        =   1
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
      Bindings        =   "Formc146.frx":00EA
      Height          =   330
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   307429377
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   307429377
      CurrentDate     =   36892
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3960
      Top             =   10080
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   4560
      Top             =   9960
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   375
      Left            =   5160
      Top             =   9840
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
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
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
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
      Index           =   2
      Left            =   3240
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "光坯入库"
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
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
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "Formc146"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public gygh As String

Private Sub Command1_Click()
If Adodc1.Recordset.EOF Then Exit Sub

Adodc3.RecordSource = "SELECT 单号 FROM JGMX WHERE 锅号='" & Text1 & "' "
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
MsgBox ("已经开出发货单据," + Adodc3.Recordset.Fields(0) + " 不能重复出库？")
Exit Sub
End If

Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF

Adodc2.RecordSource = "SELECT 顺序号 FROM JGMX WHERE 单号='" & Formc15.Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
ID = Adodc2.Recordset.Fields(0) + 1
SXH = Adodc2.Recordset.Fields(0) + 1
Else
ID = 1
SXH = 1
End If

sql1 = "INSERT INTO dbo.jgmx(入库单据,入库序号,加工单位,计划号,ip,和约号,锅号,品名,颜色,数量,匹数,光坯,加工类别,单价,金额,核算,负责,单号,日期,顺序号,单位) Values('" & Adodc1.Recordset.Fields(0) & "','" & Adodc1.Recordset.Fields(1) & "','" & Adodc1.Recordset.Fields(2) & "','" & Adodc1.Recordset.Fields(3) & "','" & Adodc1.Recordset.Fields(4) & "','" & Adodc1.Recordset.Fields(5) & "','" & Adodc1.Recordset.Fields(6) & "','" & Adodc1.Recordset.Fields(7) & "','" & Adodc1.Recordset.Fields(8) & "','" & Adodc1.Recordset.Fields(9) & "','" & Adodc1.Recordset.Fields(10) & "','" & Adodc1.Recordset.Fields(11) & "','" & Adodc1.Recordset.Fields(12) & "','" & Adodc1.Recordset.Fields(13) & "','" & Adodc1.Recordset.Fields(14) & "','" & Adodc1.Recordset.Fields(15) & "','" & Formc15.DataCombo19 & "','" & Formc15.Label13 & "','" & Date & "','" & SXH & "','公斤')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

Adodc1.Recordset.MoveNext
Loop
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''发货
sql2 = "update dbo.kpd set FH=convert(nvarchar ,'" & Now & "',120),zt='已发货' WHERE 锅号='" & Text1 & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定出库吗？", vbYesNo) = vbNo Then Exit Sub

If Formc15.DataCombo19 = "" Then
MsgBox ("请在发货界面选择负责人再继续！")
Exit Sub
End If

'For i = 1 To VSFlexGrid1.Rows - 1
'If VSFlexGrid1.Cell(flexcpChecked, i, 1) = 1 Then
'If VSFlexGrid1.TextMatrix(i, 13) = "" Then
'MsgBox ("请设置加工项目才能发货")
'Exit Sub
'End If
'End If
'Next

For i = 1 To VSFlexGrid1.Rows - 1
If Formc15.Label13.Caption = "" Then Exit Sub
If VSFlexGrid1.Cell(flexcpChecked, i, 1) = 1 Then

Adodc2.RecordSource = "SELECT 顺序号,加工单位 FROM JGMX WHERE 单号='" & Formc15.Label13.Caption & "' ORDER BY 顺序号 DESC"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
ID = Adodc2.Recordset.Fields(0) + 1
SXH = Adodc2.Recordset.Fields(0) + 1
S17 = Adodc2.Recordset.Fields(1)
Else
ID = 1
SXH = 1
S17 = ""
End If
                                                                                        
S1 = VSFlexGrid1.TextMatrix(i, 1)
S2 = VSFlexGrid1.TextMatrix(i, 2)
s3 = VSFlexGrid1.TextMatrix(i, 3)
s4 = VSFlexGrid1.TextMatrix(i, 4)
s5 = VSFlexGrid1.TextMatrix(i, 5)
s6 = VSFlexGrid1.TextMatrix(i, 6)
s7 = VSFlexGrid1.TextMatrix(i, 7)
s8 = VSFlexGrid1.TextMatrix(i, 8)
s9 = VSFlexGrid1.TextMatrix(i, 9)
s10 = Val(VSFlexGrid1.TextMatrix(i, 10))  '毛坯数量
S11 = Val(VSFlexGrid1.TextMatrix(i, 11))  ''匹数
S12 = Val(VSFlexGrid1.TextMatrix(i, 12))  '''光坯数量
S13 = VSFlexGrid1.TextMatrix(i, 13)       '''加工类别
S14 = Val(VSFlexGrid1.TextMatrix(i, 14))  '''单价
s15 = Val(VSFlexGrid1.TextMatrix(i, 15))  ''金额
s18 = Val(VSFlexGrid1.TextMatrix(i, 17))   ''库存匹数
S12 = Format(S12 / s18 * S11, "#0.0")

If Formc15.Option4.value = True Then
s16 = "毛坯"       ''核算
s15 = Val(VSFlexGrid1.TextMatrix(i, 14)) * Val(VSFlexGrid1.TextMatrix(i, 10)) ''金额
End If
If Formc15.Option5.value = True Then
s16 = "光坯"       ''核算
s15 = Val(VSFlexGrid1.TextMatrix(i, 14)) * Val(VSFlexGrid1.TextMatrix(i, 12)) ''金额
End If
If Formc15.Option6.value = True Then
s16 = "匹数"       ''核算
s15 = Val(VSFlexGrid1.TextMatrix(i, 14)) * Val(VSFlexGrid1.TextMatrix(i, 11)) ''金额
End If

If VSFlexGrid1.TextMatrix(i, 19) = "" Then
s18 = VSFlexGrid1.TextMatrix(i, 18)   ''''备注
Else
s18 = VSFlexGrid1.TextMatrix(i, 19)
End If

If S17 <> s3 And S17 <> "" Then
MsgBox ("不是一个客户的，不能开发货单")
Exit Sub
End If
'Adodc28.RecordSource = "SELECT * from yj_qfts where 客户 = '" & S17 & "' "
'Adodc28.Refresh
'If Not Adodc28.Recordset.EOF Then
'  If Val(Adodc28.Recordset.Fields(3)) >= Val(Adodc28.Recordset.Fields(1)) Then
'  MsgBox ("客户欠费超出预警，不能开发货单")
'Exit Sub
'End If
Else
sql1 = "INSERT INTO dbo.jgmx(入库单据,入库序号,加工单位,缸号,ip,和约号,锅号,品名,颜色,数量,匹数,光坯,加工类别,单价,金额,核算,负责,单号,日期,顺序号,单位,备注,跟单) Values('" & S1 & "','" & S2 & "','" & s3 & "','" & s4 & "','" & s5 & "','" & s6 & "','" & s7 & "','" & s8 & "','" & s9 & "','" & s10 & "','" & S11 & "','" & S12 & "','" & S13 & "','" & S14 & "','" & s15 & "','" & s16 & "','" & Formc15.DataCombo19 & "','" & Formc15.Label13.Caption & "','" & Formc15.Text5 & "','" & SXH & "','公斤','" & s18 & "','" & Formc15.DataCombo17 & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

sql2 = "update dbo.kpd set FH=convert(nvarchar ,'" & Now & "',120),zt='已发货' WHERE 锅号='" & VSFlexGrid1.TextMatrix(i, 7) & "' and 编号='" & VSFlexGrid1.TextMatrix(i, 4) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
'End If
Next
Call Command3_Click   '''''查询按钮
Formc15.Adodc9.Refresh
Formc15.Adodc21.Refresh
End Sub

Private Sub Command3_Click()
sql1 = ""

If Check2(0).value = 1 Then
sql1 = sql1 + "款号 like '%'+'" & Text3 & "'+'%' and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "客户 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "日期 between cast('" & DTPicker3.value & "' as datetime) and cast('" & DTPicker4.value & "' as datetime) and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "锅号 like '%'+'" & Text1 & "'+'%' and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc1.RecordSource = "select distinct '00000000' as 单据,'1' as 序号,客户,缸号,'' as 单号序号,款号,锅号,品名,颜色+色号 as 色别,库存数量,库存匹数 as 发货匹数,光坯数量,收费项目,单价,(case when isnull(核算方式,'')='毛坯' then round(毛坯重量*isnull(单价,0),2) when isnull(核算方式,'')='光坯' then round(光坯数量*isnull(单价,0),2) when isnull(核算方式,'')='匹数' then round(光坯匹数*isnull(单价,0),2) end) as 合计金额,核算方式,库存匹数,备注,图案 as 收费明细,日期 from v_kpd_fh  WHERE (" + sql1 + ") and 库存匹数>0 order by 日期"
Adodc1.Refresh
VSFlexGrid1.AutoSizeMode = flexAutoSizeRowHeight
VSFlexGrid1.AutoSize 0, VSFlexGrid1.Cols - 1, False, 30

If Adodc1.Recordset.EOF Then
hs = 0
Else
hs = Adodc1.Recordset.RecordCount + 1
End If


If hs > 0 Then
    With VSFlexGrid1
        .Editable = flexEDKbdMouse
'        .AutoSize 0
        .Cell(flexcpChecked, 1, 1, hs - 1, 1) = 2
'        .Cell(MergeCells, 1, 2, hs - 1, 2) = True
        End With
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 0, 11, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 12, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 15, , vbGreen
End If

VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(5) = 0

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

End Sub

Private Sub Command4_Click()
fhsx = 1
Formc15.Timer2.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
DataCombo1 = ""
Check2(1) = 1
Check2(4) = 1
Check2(6) = 1
DTPicker3.value = Date - 10
DTPicker4.value = Date
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(0) = 100
End Sub

Private Sub Label2_Click()
Formc34.Show
End Sub

Private Sub Text2_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text2 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'  group by 简称"
Adodc5.Refresh
End Sub
Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    If (col = 14 Or col = 1 Or col = 10 Or col = 11) Then
    Cancel = False
    Else
    Cancel = True
    End If ''不能修改指定的列
End Sub

Private Sub VSFlexGrid1_CellChanged(ByVal Row As Long, ByVal col As Long)
    If col = 14 Or col = 10 Or col = 11 Then
    With VSFlexGrid1
     rs = .Row
     cl = .col
     If .TextMatrix(rs, 16) = "毛坯" Then
    .TextMatrix(rs, 15) = Format(Val(VSFlexGrid1.TextMatrix(rs, 10)) * Val(VSFlexGrid1.TextMatrix(rs, 14)), "#0.00")
     End If
     If .TextMatrix(rs, 16) = "光坯" Then
    .TextMatrix(rs, 15) = Format(Val(VSFlexGrid1.TextMatrix(rs, 12)) * Val(VSFlexGrid1.TextMatrix(rs, 14)), "#0.00")
     End If
     End With
     End If
End Sub

Private Sub VSFlexGrid1_Click()
r = VSFlexGrid1.RowSel
c = VSFlexGrid1.ColSel

If c = 1 Then
If InStr(VSFlexGrid1.TextMatrix(r, 1), "Total") > 0 Then
    If VSFlexGrid1.Cell(flexcpChecked, 1, 1, r - 1, 1) = 2 Then
            VSFlexGrid1.Cell(flexcpChecked, 1, 1, r - 1, 1) = 1
    End If
    
End If
End If

If c = 2 Then
If InStr(VSFlexGrid1.TextMatrix(r, 1), "Total") > 0 Then
    If VSFlexGrid1.Cell(flexcpChecked, 1, 1, r - 1, 1) = 1 Then
            VSFlexGrid1.Cell(flexcpChecked, 1, 1, r - 1, 1) = 2
    
    End If
End If
End If

If c = 7 Then
hssx = 2
Formy85.Text1(0) = VSFlexGrid1.TextMatrix(r, 4)
Formy85.Text1(1) = VSFlexGrid1.TextMatrix(r, 5)
Formy85.Show
End If

End Sub

Private Sub jc()
sl1 = 0
sl2 = 0
For i = 1 To VSFlexGrid1.Rows - 1
If VSFlexGrid1.Cell(flexcpChecked, i, 1) = 1 Then
sl1 = sl1 + 1
sl2 = sl2 + Val(VSFlexGrid1.TextMatrix(i, 4))
End If
Next
End Sub

