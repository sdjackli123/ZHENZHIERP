VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formr327 
   BackColor       =   &H00C0E0FF&
   Caption         =   "输送监控"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   14520
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   360
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "审核"
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   15000
      Left            =   3840
      Top             =   840
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   480
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   9480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Formr327.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   2280
      Top             =   10320
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   1200
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1080
      Top             =   10440
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
      Height          =   375
      Left            =   1320
      Top             =   10320
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formr327.frx":0006
      Height          =   3135
      Left            =   360
      TabIndex        =   7
      Top             =   3480
      Width           =   15855
      _cx             =   27966
      _cy             =   5530
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formr327.frx":001B
      Height          =   1455
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   15855
      _cx             =   27966
      _cy             =   2566
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Formr327.frx":0030
      Height          =   2775
      Left            =   360
      TabIndex        =   9
      Top             =   7080
      Width           =   15855
      _cx             =   27966
      _cy             =   4895
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
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "正输送信息"
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
      Left            =   360
      TabIndex        =   13
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "待输送信息"
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
      Left            =   360
      TabIndex        =   12
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "异常信息"
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
      Left            =   360
      TabIndex        =   11
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "语音选择"
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
      Left            =   360
      TabIndex        =   10
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Formr327"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim plshsx As Integer
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim MXH  As Integer    '''''''''循环读M
Dim ctStart As Long, ctSelLen As Long, ctOnlySel As Boolean
Dim ctRead As Boolean, ctPause As Boolean
Dim WithEvents Voice As SpVoice
Attribute Voice.VB_VarHelpID = -1

Private Sub Combo2_Click()
Set Voice.Voice = Voice.GetVoices().Item(Combo2.ListIndex)
End Sub

Private Sub Command1_Click()
Formr328.Check2(6).value = 1
Formr328.Show
End Sub

Private Sub Command2_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "输送统计")
End Sub

Private Sub Command3_Click()
Adodc4.RecordSource = "SELECT 料单编号,工序名称,染化助库,染化助名称,配料单位,round(配料用量,4) as 配料用量,实际称量,次序号,机台,锅号,管道编号,车台编号,开始输送,输送状态,输送信息 FROM v_pldr_dx WHERE 申请时间 is not null and 开始输送 is not null and 输送结束 is null and isnull(输送信息,'')<> '异常跳过'  ORDER BY 申请时间,工序名称,次序号"
Adodc4.Refresh

Adodc1.RecordSource = "SELECT 料单编号,工序名称,染化助库,染化助名称,配料单位,round(配料用量,4) as 配料用量,实际称量,次序号,机台,锅号,申请时间,管道编号,车台编号 FROM v_pldr_dx WHERE 申请时间 is not null and 开始输送 is null and 输送结束 is null and isnull(称量标记,'')<>'Y' and isnull(车台编号,'')<>'' and isnull(管道编号,'')<>'' ORDER BY 申请时间,工序名称,次序号"
Adodc1.Refresh

Adodc2.RecordSource = "SELECT 料单编号,工序名称,染化助库,染化助名称,配料单位,round(配料用量,4) as 配料用量,实际称量,次序号,机台,锅号,管道编号,车台编号,开始输送,输送状态,输送信息 FROM v_pldr_dx WHERE isnull(输送信息,'') like '%异常%' and isnull(称量标记,'')<>'Y' ORDER BY 申请时间 desc,工序名称,次序号"
Adodc2.Refresh

VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(1) = 1300
VSFlexGrid1.ColWidth(2) = 1300
VSFlexGrid1.ColWidth(3) = 1000
VSFlexGrid1.ColWidth(4) = 1000
VSFlexGrid1.ColWidth(5) = 1000
VSFlexGrid1.ColWidth(6) = 1000
VSFlexGrid1.ColWidth(7) = 1000
VSFlexGrid1.ColWidth(8) = 1000
VSFlexGrid1.ColWidth(9) = 1200
VSFlexGrid1.ColWidth(10) = 1200
VSFlexGrid1.ColWidth(11) = 1900

If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 800
Next
End If

VSFlexGrid2.ColWidth(0) = 100
VSFlexGrid2.ColWidth(1) = 1000
VSFlexGrid2.ColWidth(2) = 1000
VSFlexGrid2.ColWidth(3) = 1000
VSFlexGrid2.ColWidth(4) = 1000
VSFlexGrid2.ColWidth(5) = 1000
VSFlexGrid2.ColWidth(6) = 1000
VSFlexGrid2.ColWidth(7) = 1000
VSFlexGrid2.ColWidth(8) = 1000
VSFlexGrid2.ColWidth(9) = 1000
VSFlexGrid2.ColWidth(10) = 1000

VSFlexGrid3.ColWidth(0) = 100
VSFlexGrid3.ColWidth(1) = 1000
VSFlexGrid3.ColWidth(2) = 1000
VSFlexGrid3.ColWidth(3) = 1000
VSFlexGrid3.ColWidth(4) = 1000
VSFlexGrid3.ColWidth(5) = 1000
VSFlexGrid3.ColWidth(6) = 1000
VSFlexGrid3.ColWidth(7) = 1000
VSFlexGrid3.ColWidth(8) = 1000
VSFlexGrid3.ColWidth(9) = 1000
VSFlexGrid3.ColWidth(10) = 1000

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Formr448.Show
End Sub

Private Sub Form_Load()
plshsx = 1

Set Voice = New SpVoice
    Dim Token As ISpeechObjectToken
    For Each Token In Voice.GetVoices
    Combo2.AddItem (Token.GetDescription())
    Next
    Combo2.ListIndex = 0
    Voice.Volume = 100
    Voice.Rate = -1

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.CommandTimeout = 10000
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.CommandTimeout = 10000
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
End Sub

Private Sub Timer1_Timer()
Call Command3_Click
If Not Adodc4.Recordset.EOF Then
If InStr(Adodc4.Recordset.Fields(13), "二次") > 0 Then
Text1 = "染缸； " + Adodc4.Recordset.Fields(8) + "  ；工序, " + Adodc4.Recordset.Fields(1) + "；" + ",原料； " + Adodc4.Recordset.Fields(3) + "；输送完成！"
Else
Text1 = "染缸； " + Adodc4.Recordset.Fields(8) + "  ；工序, " + Adodc4.Recordset.Fields(1) + "；" + ",原料； " + Adodc4.Recordset.Fields(3) + "；" + Adodc4.Recordset.Fields(13)
End If
Timer2.Enabled = True
Else
Text1 = ""
Timer2.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Call SpeakStr
End Sub

Private Sub VSFlexGrid2_DblClick()
r = VSFlexGrid2.RowSel
c = VSFlexGrid2.ColSel
If r > 0 Then
response = MsgBox("信息操作说明 点击是 作为异常处理 强制结束   点击否 不做任何操作  点击取消 重新输送 ！", vbYesNoCancel)
If response = vbCancel Then
sql1 = "UPDATE pldr SET 输送状态='',输送信息='',开始输送=null WHERE 料单编号='" & VSFlexGrid2.TextMatrix(r, 1) & "' and 次序号='" & VSFlexGrid2.TextMatrix(r, 8) & "' and  工序名称='" & VSFlexGrid2.TextMatrix(r, 2) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
If response = vbYes Then
sql1 = "UPDATE pldr SET 称量标记='',输送状态='',输送信息='异常跳过' WHERE 料单编号='" & VSFlexGrid2.TextMatrix(r, 1) & "' and 次序号='" & VSFlexGrid2.TextMatrix(r, 8) & "' and  工序名称='" & VSFlexGrid2.TextMatrix(r, 2) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
End If
End Sub

Private Sub VSFlexGrid3_dblClick()
r = VSFlexGrid3.RowSel
c = VSFlexGrid3.ColSel
If r > 0 Then
response = MsgBox("异常信息操作说明 点击是 维修正常 可以继续输送 点击否 不做任何操作  点击取消 输送终止 ！", vbYesNoCancel)
If response = vbYes Then
sql1 = "UPDATE pldr SET 输送状态='',输送信息='',开始输送=null,输送结束=null,称量标记='' WHERE 料单编号='" & VSFlexGrid3.TextMatrix(r, 1) & "' and 次序号='" & VSFlexGrid3.TextMatrix(r, 8) & "' and  工序名称='" & VSFlexGrid3.TextMatrix(r, 2) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
If response = vbCancel Then
sql1 = "UPDATE pldr SET 称量标记='Y' WHERE 料单编号='" & VSFlexGrid3.TextMatrix(r, 1) & "' and 次序号='" & VSFlexGrid3.TextMatrix(r, 8) & "' and  工序名称='" & VSFlexGrid3.TextMatrix(r, 2) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
End If
End Sub

Private Sub SpeakStr(Optional OnlySel As Boolean)
    If OnlySel And Text1.SelLength > 0 Then
       ctStart = Text1.SelStart: ctSelLen = Text1.SelLength
       nStr = Text1.SelText
    Else
       ctStart = 0: ctSelLen = Len(Text1.Text): nStr = Text1.Text
    End If
    ctRead = True: ctOnlySel = OnlySel
    Voice.Speak nStr, SVSFlagsAsync   '开始以异步方式朗读
End Sub


