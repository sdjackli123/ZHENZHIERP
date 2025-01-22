VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw119 
   BackColor       =   &H00C0E0FF&
   Caption         =   "特种查询"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "应付类别"
      Height          =   735
      Left            =   8640
      TabIndex        =   26
      Top             =   480
      Width           =   3615
      Begin VB.OptionButton Option4 
         Caption         =   "五金"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "染料"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw119.frx":0000
      Height          =   7335
      Left            =   840
      TabIndex        =   24
      Top             =   2880
      Width           =   17175
      _cx             =   30295
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw119.frx":0015
      Height          =   330
      Index           =   0
      Left            =   5760
      TabIndex        =   23
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "凭证生成"
      Height          =   375
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   4200
      TabIndex        =   18
      Top             =   1920
      Width           =   4095
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "不选"
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "贷方金额=0"
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "借方金额=0"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成查询"
      Height          =   375
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   840
      TabIndex        =   11
      Top             =   1920
      Width           =   3135
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "科目"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         Top             =   120
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "类别"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/M/d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/M/d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   423428097
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   423428097
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   11640
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   423428097
      CurrentDate     =   36892
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7200
      Top             =   10320
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
      Left            =   7080
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6720
      Top             =   10440
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
      Left            =   7560
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   6960
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   6480
      Top             =   10320
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   1
      Left            =   5760
      TabIndex        =   25
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "对方科目"
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
      TabIndex        =   17
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      Height          =   375
      Index           =   0
      Left            =   11640
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
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
      TabIndex        =   10
      Top             =   480
      Width           =   1695
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
      TabIndex        =   9
      Top             =   480
      Width           =   1095
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
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Formw119"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Private Sub Check1_Click(Index As Integer)
Select Case Index
       Case 2
If Check1(2).value = 1 Then
DataCombo1(2).Text = Check1(2).Caption
End If
       Case 3
If Check1(3).value = 1 Then
DataCombo1(3).Text = Check1(3).Caption
End If
End Select
End Sub

Private Sub Command1_Click()
If MsgBox("操作日期为：" + Trim(DTPicker3.value) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("操作期间为：" + Trim(Month(DTPicker3.value)) + "正确吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定生成特种系列的凭证吗？", vbYesNo) = vbNo Then Exit Sub
Call FKHZXJPZ(CDate(Text1.Text), CDate(Text2.Text), CDate(DTPicker3.value))
Call FKHZYHPZ(CDate(Text1.Text), CDate(Text2.Text), CDate(DTPicker3.value))
Call SKHZXJPZ(CDate(Text1.Text), CDate(Text2.Text), CDate(DTPicker3.value))
Call SKHZYHPZ(CDate(Text1.Text), CDate(Text2.Text), CDate(DTPicker3.value))
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
sql = ""
For i = 0 To 1
If Check1(i).value = 1 Then
sql = sql + "" + Label1(i).Caption + " like '%'+ '" & DataCombo1(i).Text & "'+'%'" + " and "
End If
Next

t1 = Format(Text1, "yyyy-mm-dd")
t2 = Format(Text2, "yyyy-mm-dd")

If Option1.value = True Then
If Len(sql) > 1 Then
sql = Left$(Trim(sql), Len(Trim(sql)) - 3)
Adodc1.RecordSource = "SELECT 类别,日期,单据号,摘要,对方科目,借方金额,贷方金额,抹零金额,备注1,备注2,备注3 FROM TZJZMX WHERE (" + sql + ") AND CONVERT(varchar,日期, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' and cast(借方金额 as real)=0 ORDER BY 日期,序号"
Adodc1.Refresh
End If
End If

If Option2.value = True Then
If Len(sql) > 1 Then
sql = Left$(Trim(sql), Len(Trim(sql)) - 3)
Adodc1.RecordSource = "SELECT 类别,日期,单据号,摘要,对方科目,借方金额,贷方金额,抹零金额,备注1,备注2,备注3 FROM TZJZMX WHERE (" + sql + ") AND CONVERT(varchar,日期, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' and cast(贷方金额 as real)=0 ORDER BY 日期,序号"
Adodc1.Refresh
End If
End If

If Option3.value = True Then
If Len(sql) > 1 Then
sql = Left$(Trim(sql), Len(Trim(sql)) - 3)
Adodc1.RecordSource = "SELECT 类别,日期,单据号,摘要,对方科目,借方金额,贷方金额,抹零金额,备注1,备注2,备注3 FROM TZJZMX WHERE (" + sql + ") AND CONVERT(varchar,日期, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' ORDER BY 日期,序号"
Adodc1.Refresh
End If
End If



VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 1, 6, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 1, 7, , &HC0C0&

End Sub

Private Sub Command4_Click()
Formw1132.Show
End Sub

Private Sub Command5_Click()
Call OutadodcToExcel2(VSFlexGrid1, 6, 7, "特种记账 日期范围： " + Text1.Text + "--" + Text2.Text)
End Sub

Private Sub DTPicker1_Change()
Text1.Text = DTPicker1.value
Text1.SetFocus
End Sub

Private Sub DTPicker1_CloseUp()
Text1.Text = DTPicker1.value
Text1.SetFocus
End Sub
Private Sub DTPicker2_Change()
Text2.Text = DTPicker2.value
Text2.SetFocus
End Sub

Private Sub DTPicker2_CloseUp()
Text2.Text = DTPicker2.value
Text2.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = Date
Text2.Text = Date
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset


Option4(0).value = True
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)  ORDER BY 序号 DESC"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT MAX(序号) FROM TZJZMX "
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Combo1.Text = ""
For i = 0 To 11
DataCombo1(i).Text = ""
Next
DataCombo1(11).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc2.Recordset.Fields(0) + 1
End If
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 1500
VSFlexGrid1.ColWidth(5) = 1500
VSFlexGrid1.ColWidth(8) = 1500
VSFlexGrid1.ColWidth(13) = 0
End Sub

Private Sub FKHZXJPZ(DT1 As Date, dt2 As Date, dt3 As Date)   '''''''''付款汇总--现金--材料
'On Error Resume Next
Adodc5.RecordSource = "SELECT * FROM CLFKPZ WHERE 制单 like '%现金材料%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
If Not Adodc5.Recordset.EOF Then
If MsgBox("已有现金付款凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CLFKPZ WHERE 制单 like '%现金材料%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

Adodc4.RecordSource = "SELECT * FROM CLFKPZ WHERE 凭证号 like 'W2-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-3)) FROM CLFKPZ WHERE 凭证号 like 'W2-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
PZH = "W2-" + Trim(Adodc4.Recordset.Fields(0) + 1)
Else
PZH = "W2-1"
End If

Adodc6.RecordSource = "SELECT * FROM TZJZMX where 贷方金额<>0 and 类别 LIKE '现金%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "' order by 对方科目"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
Exit Sub
Else
Adodc6.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc6.Recordset.EOF
For i = 1 To 7

If InStr(Adodc6.Recordset.Fields(4), "-") > 0 Then
l1 = Left(Adodc6.Recordset.Fields(4), InStr(Adodc6.Recordset.Fields(4), "-") - 1)      ''''借方总账科目
L2 = Mid(Adodc6.Recordset.Fields(4), InStr(Adodc6.Recordset.Fields(4), "-") + 1)         ''''借方明细科目
Else
l1 = Adodc6.Recordset.Fields(4)
L2 = ""
End If

If InStr(Adodc6.Recordset.Fields(0), "-") > 0 Then
L3 = Left(Adodc6.Recordset.Fields(0), InStr(Adodc6.Recordset.Fields(0), "-") - 1)    '''''''''''贷方总账科目
L4 = Mid(Adodc6.Recordset.Fields(0), InStr(Adodc6.Recordset.Fields(0), "-") + 1)      '''贷方明细科目
Else
L3 = Adodc6.Recordset.Fields(0)
L4 = ""
End If

sql1 = "INSERT INTO CLFKPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('" & Adodc6.Recordset.Fields(3) & "','" & l1 & "','" & L2 & "','" & Adodc6.Recordset.Fields(6) & "','','" & PZH & "','" & dt3 & "','','','','现金材料','','','')"
sql2 = "INSERT INTO CLFKPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('" & Adodc6.Recordset.Fields(3) & "','" & L3 & "','" & L4 & "','','" & Adodc6.Recordset.Fields(6) & "','" & PZH & "','" & dt3 & "','','','','现金材料','','','')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc6.Recordset.MoveNext
If Adodc6.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金付款凭证")
Exit Sub
End If
Next
Adodc4.RecordSource = "SELECT * FROM CLFKPZ WHERE 凭证号 like 'W2-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-3)) FROM CLFKPZ WHERE 凭证号 like 'W2-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
PZH = "W2-" + Trim(Adodc4.Recordset.Fields(0) + 1)
Else
PZH = "W2-1"
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金付款凭证")
End If

End Sub
''''''''''''''
Private Sub FKHZYHPZ(DT1 As Date, dt2 As Date, dt3 As Date)   '''''''''付款汇总---银行存款
'On Error Resume Next

Adodc5.RecordSource = "SELECT * FROM CLFKPZ WHERE 制单 like '%银行材料%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
If Not Adodc5.Recordset.EOF Then
If MsgBox("已有银行付款凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CLFKPZ WHERE 制单 like '%银行材料%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

Adodc4.RecordSource = "SELECT * FROM CLFKPZ WHERE 凭证号 like 'W4-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-3)) FROM CLFKPZ WHERE 凭证号 like 'W4-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
PZH = "W4-" + Trim(Adodc4.Recordset.Fields(0) + 1)
Else
PZH = "W4-1"
End If

Adodc6.RecordSource = "SELECT * FROM TZJZMX where  贷方金额<>0 and 类别 LIKE '银行存款%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "' order by 对方科目"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
Exit Sub
Else
Adodc6.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc6.Recordset.EOF
For i = 1 To 7

If InStr(Adodc6.Recordset.Fields(4), "-") > 0 Then
l1 = Left(Adodc6.Recordset.Fields(4), InStr(Adodc6.Recordset.Fields(4), "-") - 1)      ''''借方总账科目
L2 = Mid(Adodc6.Recordset.Fields(4), InStr(Adodc6.Recordset.Fields(4), "-") + 1)         ''''借方明细科目
Else
l1 = Adodc6.Recordset.Fields(4)
L2 = ""
End If

If InStr(Adodc6.Recordset.Fields(0), "-") > 0 Then
L3 = Left(Adodc6.Recordset.Fields(0), InStr(Adodc6.Recordset.Fields(0), "-") - 1)    '''''''''''贷方总账科目
L4 = Mid(Adodc6.Recordset.Fields(0), InStr(Adodc6.Recordset.Fields(0), "-") + 1)      '''贷方明细科目
Else
L3 = Adodc6.Recordset.Fields(0)
L4 = ""
End If

sql1 = "INSERT INTO CLFKPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('" & Adodc6.Recordset.Fields(3) & "','" & l1 & "','" & L2 & "','" & Adodc6.Recordset.Fields(6) & "','','" & PZH & "','" & dt3 & "','','','','银行材料','','','')"
sql2 = "INSERT INTO CLFKPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('" & Adodc6.Recordset.Fields(3) & "','" & L3 & "','" & L4 & "','','" & Adodc6.Recordset.Fields(6) & "','" & PZH & "','" & dt3 & "','','','','银行材料','','','')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc6.Recordset.MoveNext
If Adodc6.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行付款凭证")
Exit Sub
End If
Next
Adodc4.RecordSource = "SELECT * FROM CLFKPZ WHERE 凭证号 like 'W4-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-3)) FROM CLFKPZ WHERE 凭证号 like 'W4-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
PZH = "W4-" + Trim(Adodc4.Recordset.Fields(0) + 1)
Else
PZH = "W4-1"
End If
Loop
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行付款凭证")
End If
End Sub

Private Sub SKHZXJPZ(DT1 As Date, dt2 As Date, dt3 As Date)    ''''''''收款汇总----现金
'On Error Resume Next
Adodc5.RecordSource = "SELECT * FROM CLSKPZ WHERE 制单 like '%现金-发货%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
If Not Adodc5.Recordset.EOF Then
If MsgBox("已有现金收款凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CLSKPZ WHERE 制单 like '%现金-发货%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

Adodc4.RecordSource = "SELECT * FROM CLSKPZ WHERE 凭证号 like 'W1-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-3)) FROM CLSKPZ WHERE 凭证号 like 'W1-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
PZH = "W1-" + Trim(Adodc4.Recordset.Fields(0) + 1)
Else
PZH = "W1-1"
End If

Adodc6.RecordSource = "SELECT * FROM TZJZMX where  借方金额<>0 and 类别 LIKE '现金%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "' order by 对方科目"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
Exit Sub
Else
Adodc6.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc6.Recordset.EOF
For i = 1 To 7

If InStr(Adodc6.Recordset.Fields(0), "-") > 0 Then
l1 = Left(Adodc6.Recordset.Fields(0), InStr(Adodc6.Recordset.Fields(0), "-") - 1)    '''''''''''贷方总账科目
L2 = Mid(Adodc6.Recordset.Fields(0), InStr(Adodc6.Recordset.Fields(0), "-") + 1)      '''贷方明细科目
Else
l1 = Adodc6.Recordset.Fields(0)
L2 = ""
End If

If InStr(Adodc6.Recordset.Fields(4), "-") > 0 Then
L3 = Left(Adodc6.Recordset.Fields(4), InStr(Adodc6.Recordset.Fields(4), "-") - 1)      ''''借方总账科目
L4 = Mid(Adodc6.Recordset.Fields(4), InStr(Adodc6.Recordset.Fields(4), "-") + 1)         ''''借方明细科目
Else
L3 = Adodc6.Recordset.Fields(4)
L4 = ""
End If


sql1 = "INSERT INTO CLSKPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('" & Adodc6.Recordset.Fields(3) & "','" & l1 & "','" & L2 & "','" & Adodc6.Recordset.Fields(5) & "','','" & PZH & "','" & dt3 & "','','','','现金-发货','','','')"
sql2 = "INSERT INTO CLSKPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('" & Adodc6.Recordset.Fields(3) & "','" & L3 & "','" & L4 & "','','" & Adodc6.Recordset.Fields(5) & "','" & PZH & "','" & dt3 & "','','','','现金-发货','','','')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc6.Recordset.MoveNext
If Adodc6.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金收款凭证")
Exit Sub
End If
Next
Adodc4.RecordSource = "SELECT * FROM CLSKPZ WHERE 凭证号 like 'W1-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-3)) FROM CLSKPZ WHERE 凭证号 like 'W1-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
PZH = "W1-" + Trim(Adodc4.Recordset.Fields(0) + 1)
Else
PZH = "W1-1"
End If
Loop
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金收款凭证")
End If
End Sub

Private Sub SKHZYHPZ(DT1 As Date, dt2 As Date, dt3 As Date)    ''''''''收款汇总----银行存款
'On Error Resume Next

Adodc5.RecordSource = "SELECT * FROM CLSKPZ WHERE 制单 like '%银行-发货%' and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc5.Refresh
If Not Adodc5.Recordset.EOF Then
If MsgBox("已有现金收款凭证，是否重新生成？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM CLSKPZ WHERE 制单 like '%银行-发货%')>0 and 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

Adodc4.RecordSource = "SELECT * FROM CLSKPZ WHERE 凭证号 like 'W3-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-3)) FROM CLSKPZ WHERE 凭证号 like 'W3-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
PZH = "W3-" + Trim(Adodc4.Recordset.Fields(0) + 1)
Else
PZH = "W3-1"
End If

Adodc6.RecordSource = "SELECT * FROM TZJZMX where  借方金额<>0 and 类别 LIKE '银行%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "' order by 对方科目"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
Exit Sub
Else
Adodc6.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc6.Recordset.EOF
For i = 1 To 7

If InStr(Adodc6.Recordset.Fields(0), "-") > 0 Then
l1 = Left(Adodc6.Recordset.Fields(0), InStr(Adodc6.Recordset.Fields(0), "-") - 1)    '''''''''''贷方总账科目
L2 = Mid(Adodc6.Recordset.Fields(0), InStr(Adodc6.Recordset.Fields(0), "-") + 1)      '''贷方明细科目
Else
l1 = Adodc6.Recordset.Fields(0)
L2 = ""
End If

If InStr(Adodc6.Recordset.Fields(4), "-") > 0 Then
L3 = Left(Adodc6.Recordset.Fields(4), InStr(Adodc6.Recordset.Fields(4), "-") - 1)      ''''借方总账科目
L4 = Mid(Adodc6.Recordset.Fields(4), InStr(Adodc6.Recordset.Fields(4), "-") + 1)         ''''借方明细科目
Else
L3 = Adodc6.Recordset.Fields(4)
L4 = ""
End If


sql1 = "INSERT INTO CLSKPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('" & Adodc6.Recordset.Fields(3) & "','" & l1 & "','" & L2 & "','" & Adodc6.Recordset.Fields(5) & "','','" & PZH & "','" & dt3 & "','','','','银行-发货','','','')"
sql2 = "INSERT INTO CLSKPZ(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) Values('" & Adodc6.Recordset.Fields(3) & "','" & L3 & "','" & L4 & "','','" & Adodc6.Recordset.Fields(5) & "','" & PZH & "','" & dt3 & "','','','','银行-发货','','','')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc6.Recordset.MoveNext
If Adodc6.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行收款凭证")
Exit Sub
End If
Next
Adodc4.RecordSource = "SELECT * FROM CLSKPZ WHERE 凭证号 like 'W3-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Adodc4.RecordSource = "SELECT MAX(right(凭证号,len(凭证号)-3)) FROM CLSKPZ WHERE 凭证号 like 'W3-%' AND 日期 BETWEEN '" & DT1 & "' AND '" & dt2 & "'"
Adodc4.Refresh
PZH = "W3-" + Trim(Adodc4.Recordset.Fields(0) + 1)
Else
PZH = "W3-1"
End If
Loop
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行收款凭证")
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
