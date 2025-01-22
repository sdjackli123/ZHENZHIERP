VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formr328 
   BackColor       =   &H00C0E0FF&
   Caption         =   "输送查询"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   615
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8520
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1455
      Left            =   10320
      TabIndex        =   5
      Top             =   240
      Width           =   5055
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "未输送"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "已输送"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "申请日期"
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "料单"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "配料日期"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "库类"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "名称"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "输送信息"
         Height          =   255
         Index           =   8
         Left            =   2760
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "机台"
         Height          =   255
         Index           =   9
         Left            =   3960
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Formr328.frx":0000
      Left            =   2280
      List            =   "Formr328.frx":000A
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   15360
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   2040
      Top             =   9840
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
      Left            =   960
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   840
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
      Bindings        =   "Formr328.frx":001A
      Height          =   330
      Left            =   5520
      TabIndex        =   17
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "染化助库名"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1080
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   423034881
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   423034881
      CurrentDate     =   36892
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formr328.frx":002F
      Height          =   6495
      Left            =   720
      TabIndex        =   20
      Top             =   1920
      Width           =   15855
      _cx             =   27966
      _cy             =   11456
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
      Height          =   330
      Left            =   5520
      TabIndex        =   21
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   5520
      TabIndex        =   22
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Formr328.frx":0044
      Height          =   330
      Left            =   8520
      TabIndex        =   23
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "染料名称"
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formr328.frx":0059
      Height          =   1335
      Left            =   720
      TabIndex        =   24
      Top             =   8520
      Width           =   12735
      _cx             =   22463
      _cy             =   2355
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
      Index           =   5
      Left            =   7800
      TabIndex        =   32
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "料单"
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
      Index           =   4
      Left            =   4800
      TabIndex        =   31
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
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
      Left            =   4800
      TabIndex        =   30
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "信息"
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
      Left            =   720
      TabIndex        =   29
      Top             =   1200
      Width           =   1335
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
      Left            =   720
      TabIndex        =   28
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
      Left            =   720
      TabIndex        =   27
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
      Left            =   4800
      TabIndex        =   26
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "机台"
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
      Index           =   6
      Left            =   7800
      TabIndex        =   25
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Formr328"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim plshsx As Integer


Private Sub Command2_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "输送统计")
End Sub

Private Sub Command3_Click()

sql1 = ""

If Check2(0).value = 1 Then
sql1 = sql1 + "配料日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "染化助名称 like '%'+'" & DataCombo4.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "isnull(称量标记,'')='Y' and 申请时间 is not null and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "料单编号 like '%'+'" & DataCombo3.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "染化助库 LIKE '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(9).value = 1 Then
sql1 = sql1 + "机台='" & Text1 & "' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "锅号 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "cast(CONVERT(varchar(120),申请时间,23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "isnull(称量标记,'')<>'Y' and 申请时间 is not null and "
End If

If Check2(8).value = 1 Then
sql1 = sql1 + "isnull(输送信息,'')='" & Combo1 & "' and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc1.RecordSource = "SELECT 料单编号,工序名称,染化助库,染化助名称,配料单位,round(配料用量,4) as 配料用量,实际称量,次序号,机台,锅号,申请时间,开始输送,管道编号,车台编号,输送信息,输送时长 FROM v_pldr_dx WHERE 申请时间 is not null and 是否审核='是'  and (" + sql1 + ") ORDER BY 申请时间 desc,开始输送 desc,工序名称,次序号"
Adodc1.Refresh

Adodc4.RecordSource = "select 染化助名称,配料单位,round(sum(isnull(配料用量,0)),5) as 配料量 from v_pldr_dx where 申请时间 is not null and 是否审核='是' and (" + sql1 + ") group by 染化助名称,配料单位 order by 染化助名称"
Adodc4.Refresh

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 500
If (Val(VSFlexGrid1.TextMatrix(i, 7)) - Val(VSFlexGrid1.TextMatrix(i, 6))) > 0.2 Then
VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, 15) = vbRed
Else
End If
VSFlexGrid1.RowHeight(i) = 500
Next
End If

VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(1) = 1000
VSFlexGrid1.ColWidth(2) = 1000
VSFlexGrid1.ColWidth(3) = 1000
VSFlexGrid1.ColWidth(4) = 1000
VSFlexGrid1.ColWidth(5) = 600
VSFlexGrid1.ColWidth(6) = 1000
VSFlexGrid1.ColWidth(7) = 1000
VSFlexGrid1.ColWidth(8) = 600
VSFlexGrid1.ColWidth(9) = 600
VSFlexGrid1.ColWidth(10) = 1000
VSFlexGrid1.ColWidth(11) = 1900
VSFlexGrid1.ColWidth(12) = 1900
VSFlexGrid1.ColWidth(13) = 600
VSFlexGrid1.ColWidth(14) = 600
VSFlexGrid1.ColWidth(15) = 600
VSFlexGrid1.ColWidth(16) = 600

VSFlexGrid2.ColWidth(0) = 300
VSFlexGrid2.ColWidth(1) = 1300
VSFlexGrid2.ColWidth(2) = 1300
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
Private Sub Command5_Click()
Call MXOutadodcToExcel(VSFlexGrid2, "用料统计")
End Sub

Private Sub Form_Load()
DTPicker1.value = Date
DTPicker2.value = Date
plshsx = 1
Text1 = ""
Combo1.Text = ""
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
Check2(6).value = 1
Adodc1.CommandTimeout = 10000
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 染化助库名 from RHZH group by 染化助库名"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT 染料名称 from RHZH group by 染料名称"
Adodc3.Refresh
Adodc4.CommandTimeout = 10000
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
End Sub



