VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FormJ8 
   BackColor       =   &H00C0E0FF&
   Caption         =   "排缸查询"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1335
      Left            =   9360
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "已染"
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   27
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "未染"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   25
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "操作"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色号"
         Height          =   255
         Index           =   5
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "机台"
         Height          =   255
         Index           =   7
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6720
      Top             =   10320
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
      Left            =   6480
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
      Bindings        =   "FormJ8.frx":0000
      Height          =   9855
      Left            =   480
      TabIndex        =   10
      Top             =   1920
      Width           =   22935
      _cx             =   40455
      _cy             =   17383
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   30
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Left            =   6840
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   4320
      TabIndex        =   12
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FormJ8.frx":0015
      Height          =   330
      Left            =   4320
      TabIndex        =   13
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   327876611
      CurrentDate     =   39961
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   327876611
      CurrentDate     =   39961
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Left            =   6840
      TabIndex        =   16
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo6 
      Bindings        =   "FormJ8.frx":002A
      Height          =   330
      Left            =   6840
      TabIndex        =   17
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "车台编号"
      Text            =   "DataCombo4"
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   6720
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   6720
      Top             =   10320
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "FormJ8.frx":003F
      Height          =   615
      Left            =   480
      TabIndex        =   26
      Top             =   11760
      Width           =   15255
      _cx             =   26908
      _cy             =   1085
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   30
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
   Begin VB.Label Label1 
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
      Index           =   6
      Left            =   480
      TabIndex        =   24
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Index           =   5
      Left            =   480
      TabIndex        =   23
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户名称"
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
      Left            =   3840
      TabIndex        =   22
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作"
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
      Left            =   3840
      TabIndex        =   21
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
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
      Left            =   6360
      TabIndex        =   20
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号"
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
      Left            =   6360
      TabIndex        =   19
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "机台"
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
      Left            =   6360
      TabIndex        =   18
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "FormJ8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public jd As Integer
Dim cdbhf As Integer

Private Sub Command1_Click()
On Error Resume Next
sql1 = ""

If Check2(0).value = 1 Then
sql1 = sql1 + "len(配料)< 9  and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "客户名称 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If


If Check2(3).value = 1 Then
sql1 = sql1 + "操作 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
t1 = Format(DTPicker1.value, "yyyy-mm-dd 00:00:01")
t2 = Format(DTPicker2.value, "yyyy-mm-dd 23:59:59")
sql1 = sql1 + "cast(排产时间 as varchar(19)) between '" & t1 & "' and '" & t2 & "' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "锅号 like '%'+'" & DataCombo4.Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "色号 like '%'+'" & DataCombo5.Text & "'+'%' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "车台='" & DataCombo6.Text & "' and "
End If

If Check2(8).value = 1 Then
sql1 = sql1 + "len(配料)> 9  and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)



Adodc1.RecordSource = "SELECT 车台, 排产编号, 锅号, 客户名称, 品名, 毛坯位置, 色号, 颜色, 重量, 克重, 存放位置, 排产备注, 染色要求, 班次, 操作, 状态, 排产时间, 配料, 打印次数 " & _
    "FROM v_kpdb " & _
    "WHERE cast(排产时间 as varchar(19)) BETWEEN '" & t1 & "' AND '" & t2 & "' " & _
    "ORDER BY CAST(" & _
    "  CASE " & _
    "    WHEN CHARINDEX('-', 车台) > 0 THEN SUBSTRING(车台, 1, CHARINDEX('-', 车台) - 1) " & _
    "    ELSE 车台 " & _
    "  END AS INT), 排产编号"
Adodc1.Refresh

Adodc1.Refresh

Adodc1.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT count(distinct 锅号) as 合计缸数,round(sum(isnull(重量,0)),2) as 合计重量 FROM v_kpdb where  (" + sql1 + ")"
Adodc4.Refresh

If Adodc1.Recordset.EOF Then
hs = 0
Else
hs = Adodc1.Recordset.RecordCount + 1
End If

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600   ''指定行高
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 700
VSFlexGrid1.ColWidth(2) = 1100
VSFlexGrid1.ColWidth(3) = 1100
VSFlexGrid1.ColWidth(4) = 1000
VSFlexGrid1.ColWidth(5) = 1100
VSFlexGrid1.ColWidth(6) = 1500
VSFlexGrid1.ColWidth(7) = 1500
VSFlexGrid1.ColWidth(9) = 800
VSFlexGrid1.ColWidth(10) = 800
VSFlexGrid1.ColWidth(11) = 1100
VSFlexGrid1.ColWidth(12) = 800
VSFlexGrid1.ColWidth(13) = 1700
VSFlexGrid1.ColWidth(15) = 1700
VSFlexGrid1.ColWidth(16) = 1100
VSFlexGrid1.ColWidth(17) = 1200
If i / 2 = Int(i / 2) Then
VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &HCDEEC6  ''指定隔行的颜色
Else
VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &H8000000F
End If
Next
End If


If hs > 0 Then ''合计数量
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 1, 8, , vbGreen
VSFlexGrid1.Subtotal flexSTCount, 1, 3, , vbBlue
End If

'VSFlexGrid1.ColWidth(0) = 200
'VSFlexGrid1.AutoSizeMode = flexAutoSizeRowHeight       自动行高
'VSFlexGrid1.AutoSize 0, VSFlexGrid1.Cols - 1, False, 30

With VSFlexGrid1
    .WordWrap = True
    .MergeCells = 4
    .MergeCol(1) = True '是否上下列合并
    .MergeCol(2) = True '是否上下列合并
    .MergeCol(3) = True '是否上下列合并
    
End With

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Call pcmx(VSFlexGrid1, "排产进度明细")
End Sub
Private Sub DataCombo7_Click(Area As Integer)
DataCombo6.Text = DataCombo7.Text
End Sub

Private Sub Form_Load()
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""
DTPicker1.value = Date - 10
DTPicker2.value = Date
cdbhf = cdbh
t1 = Format(DTPicker1.value, "yyyy-mm-dd 00:00:01")
t2 = Format(DTPicker2.value, "yyyy-mm-dd 23:59:59")
jd = 1
Check2(4).value = 1
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 车台, 排产编号, 锅号, 客户名称, 品名, 毛坯位置, 色号, 颜色, 重量, 克重, 存放位置, 排产备注, 染色要求, 班次, 操作, 状态, 排产时间, 配料, 打印次数 " & _
    "FROM v_kpdb " & _
    "WHERE cast(排产时间 as varchar(19)) BETWEEN '" & t1 & "' AND '" & t2 & "' " & _
    "ORDER BY CAST(" & _
    "  CASE " & _
    "    WHEN CHARINDEX('-', 车台) > 0 THEN SUBSTRING(车台, 1, CHARINDEX('-', 车台) - 1) " & _
    "    ELSE 车台 " & _
    "  END AS INT), 排产编号"
Adodc1.Refresh

Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from khZL  group by 简称"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 车台编号 from ct  order by IP"
Adodc3.Refresh
VSFlexGrid1.RowHeight(i) = 600   ''指定行高
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 700
VSFlexGrid1.ColWidth(2) = 1100
VSFlexGrid1.ColWidth(3) = 1100
VSFlexGrid1.ColWidth(4) = 1000
VSFlexGrid1.ColWidth(5) = 1100
VSFlexGrid1.ColWidth(6) = 1500
VSFlexGrid1.ColWidth(7) = 1500
VSFlexGrid1.ColWidth(9) = 800
VSFlexGrid1.ColWidth(10) = 800
VSFlexGrid1.ColWidth(11) = 1100
VSFlexGrid1.ColWidth(12) = 800
VSFlexGrid1.ColWidth(13) = 1700
VSFlexGrid1.ColWidth(15) = 1700
VSFlexGrid1.ColWidth(16) = 1100
VSFlexGrid1.ColWidth(17) = 1200
VSFlexGrid1.BackColorAlternate = &HCDEEC6
VSFlexGrid1.SelectionMode = flexSelectionListBox
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600   ''指定行高
Next
With VSFlexGrid1
    .WordWrap = True
    .MergeCells = 4
    .MergeCol(1) = True '是否上下列合并
    .MergeCol(2) = True '是否上下列合并
    .MergeCol(3) = True '是否上下列合并
End With
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

Private Sub VSFlexGrid1_dblClick()           ''双击跳到配料界面
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
'If pfjh = 2 Then
Formd331.Text5 = Adodc1.Recordset.Fields(2) ''锅号
Formd331.Text4 = Adodc1.Recordset.Fields(0)
Fromd331.Show
pfjh = 0
Unload Me
'End If
End Sub

