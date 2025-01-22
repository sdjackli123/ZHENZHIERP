VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms520 
   BackColor       =   &H00F3D6BE&
   Caption         =   "能耗登记"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15735
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   15735
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   495
      Left            =   10800
      TabIndex        =   48
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      Format          =   328990721
      CurrentDate     =   45411
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   11
      Left            =   9960
      TabIndex        =   46
      Text            =   "Text3"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   10
      Left            =   6360
      TabIndex        =   45
      Text            =   "Text3"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   9
      Left            =   2040
      TabIndex        =   44
      Text            =   "Text3"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   8
      Left            =   9960
      TabIndex        =   41
      Text            =   "Text3"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   7
      Left            =   6360
      TabIndex        =   40
      Text            =   "Text3"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   6
      Left            =   2040
      TabIndex        =   38
      Text            =   "Text3"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   5
      Left            =   9960
      TabIndex        =   37
      Text            =   "Text3"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   4
      Left            =   6360
      TabIndex        =   36
      Text            =   "Text3"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   3
      Left            =   2040
      TabIndex        =   35
      Text            =   "Text3"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   2
      Left            =   9960
      TabIndex        =   34
      Text            =   "Text3"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   1
      Left            =   6360
      TabIndex        =   33
      Text            =   "Text3"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   0
      Left            =   2040
      TabIndex        =   32
      Text            =   "Text3"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   3
      Left            =   9600
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   2
      Left            =   8760
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   1
      Left            =   7920
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   0
      Left            =   6360
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   4440
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   3600
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   720
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   6360
      TabIndex        =   19
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      Format          =   329056257
      CurrentDate     =   45411
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   840
      TabIndex        =   18
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      Format          =   329056257
      CurrentDate     =   45411
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00F3D6BE&
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00F3D6BE&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00F3D6BE&
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F3D6BE&
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
      Height          =   615
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F3D6BE&
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00F3D6BE&
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
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   18600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms520.frx":0000
      Height          =   6015
      Left            =   480
      TabIndex        =   0
      Top             =   6120
      Width           =   23175
      _cx             =   40878
      _cy             =   10610
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
      Left            =   17760
      Top             =   12480
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
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
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
      Height          =   255
      Index           =   5
      Left            =   11160
      TabIndex        =   47
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "汽表截止数"
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
      Index           =   14
      Left            =   4920
      TabIndex        =   43
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "汽表起始数"
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
      Index           =   13
      Left            =   720
      TabIndex        =   42
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "电表截止数"
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
      Left            =   4920
      TabIndex        =   39
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   495
      Index           =   2
      Left            =   9360
      TabIndex        =   31
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   495
      Index           =   1
      Left            =   8520
      TabIndex        =   30
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   495
      Index           =   0
      Left            =   4200
      TabIndex        =   29
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   495
      Index           =   4
      Left            =   3360
      TabIndex        =   28
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "起始时间"
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
      Left            =   1200
      TabIndex        =   17
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "截止时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   6600
      TabIndex        =   16
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "污水表截止数"
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
      Left            =   4920
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "用电数"
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
      TabIndex        =   14
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "用汽数"
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
      TabIndex        =   13
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "用水数"
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
      Left            =   9120
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "污水表起始数"
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
      Left            =   720
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "电表起始数"
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
      Left            =   720
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "水表起始数"
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
      Left            =   720
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "污水数"
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
      Left            =   9120
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3D6BE&
      Caption         =   "水表截止数"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Forms520"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command12_Click()
    ' 构造SQL查询语句，使用Text1(0)和Text2(0)中的日期值进行查询
    Adodc1.RecordSource = "SELECT * FROM cjnhtj WHERE 日期 BETWEEN CAST('" & Text1(0).Text & "' AS DATETIME) AND CAST('" & Text2(0).Text & "' AS DATETIME) ORDER BY 日期"
    ' 刷新Adodc控件以加载新的查询结果
    Adodc1.Refresh
    VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 0, 3, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 6, , vbBlue
VSFlexGrid1.Subtotal flexSTSum, 0, 9, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 0, 12, , vbBlue
End Sub


Private Sub Command2_Click()
Call OutadodcToExcel(VSFlexGrid1, 8, "能耗表")
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
    ' 将DTPicker1的日期值设置到Text1(0)，格式为年-月-日
    Text1(0).Text = Format(DTPicker1.value, "yyyy-mm-dd")
    ' 将DTPicker1的小时值设置到Text1(1)
    Text1(1).Text = Format(DTPicker1.value, "HH")
    ' 将DTPicker1的分钟值设置到Text1(2)
    Text1(2).Text = Format(DTPicker1.value, "nn")
    ' 将DTPicker1的秒数值设置到Text1(3)
    Text1(3).Text = Format(DTPicker1.value, "ss")
End Sub

Private Sub DTPicker2_Change()
 ' 将DTPicker2的日期值设置到Text2(0)，格式为年-月-日
    Text2(0).Text = Format(DTPicker2.value, "yyyy-mm-dd")
    ' 将DTPicker2的小时值设置到Text2(1)
    Text2(1).Text = Format(DTPicker2.value, "HH")
    ' 将DTPicker2的分钟值设置到Text2(2)
    Text2(2).Text = Format(DTPicker2.value, "nn")
    ' 将DTPicker2的秒数值设置到Text2(3)
    Text2(3).Text = Format(DTPicker2.value, "ss")
End Sub

Private Sub Form_Load()

cdbhf = cdbh

For i = 0 To 11
Text3(i).Text = ""
Next

' 设置DTPicker1的日期和时间为当天的7点0分0秒
    DTPicker1.value = Date + TimeSerial(7, 0, 0)
    ' 设置DTPicker2的日期和时间为次日的0点0分0秒
    DTPicker2.value = Date + TimeSerial(23, 59, 59)  ' 这里也可以用 Date + 1

 DTPicker3.value = Date
 Text1(0).Text = Format(DTPicker1.value, "yyyy-mm-dd")
    ' 将DTPicker1的小时值设置到Text1(1)
    Text1(1).Text = Format(DTPicker1.value, "HH")
    ' 将DTPicker1的分钟值设置到Text1(2)
    Text1(2).Text = Format(DTPicker1.value, "nn")
    ' 将DTPicker1的秒数值设置到Text1(3)
    Text1(3).Text = Format(DTPicker1.value, "ss")
     ' 将DTPicker2的日期值设置到Text2(0)，格式为年-月-日
    Text2(0).Text = Format(DTPicker2.value, "yyyy-mm-dd")
    ' 将DTPicker2的小时值设置到Text2(1)
    Text2(1).Text = Format(DTPicker2.value, "HH")
    ' 将DTPicker2的分钟值设置到Text2(2)
    Text2(2).Text = Format(DTPicker2.value, "nn")
    ' 将DTPicker2的秒数值设置到Text2(3)
    Text2(3).Text = Format(DTPicker2.value, "ss")
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select  * from cjnhtj where 日期=cast('" & DTPicker3.value & "' as datetime) order by 日期"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 1500
VSFlexGrid1.ColWidth(3) = 1500
VSFlexGrid1.ColWidth(4) = 1500
VSFlexGrid1.ColWidth(5) = 1500
VSFlexGrid1.ColWidth(6) = 1500
VSFlexGrid1.ColWidth(7) = 1500
VSFlexGrid1.ColWidth(8) = 1500
VSFlexGrid1.ColWidth(9) = 1500
VSFlexGrid1.ColWidth(10) = 1500
VSFlexGrid1.ColWidth(11) = 1500
VSFlexGrid1.ColWidth(12) = 1500
VSFlexGrid1.ColWidth(13) = 1500
VSFlexGrid1.ColWidth(14) = 1500
VSFlexGrid1.ColWidth(15) = 1500

End Sub

Private Sub Text3_Change(Index As Integer)
Select Case Index
       Case 1
       Text3(2) = Val(Text3(1)) - Val(Text3(0))
       Case 4
       Text3(5) = Val(Text3(4)) - Val(Text3(3))
       Case 7
       Text3(8) = Val(Text3(7)) - Val(Text3(6))
       Case 10
       Text3(11) = Val(Text3(10)) - Val(Text3(9))
       End Select
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To 11
Text3(i).Text = Adodc1.Recordset.Fields(i)
Next
End Sub

Private Sub Command3_Click()
'On Error Resume Next

For i = 0 To 11
        If Text3(i).Text = "" Then
            MsgBox "请输入数据"
            Exit Sub
        End If
    Next i
For i = 0 To 11
Text3(i) = Val(Text3(i))
Next i
For i = 0 To 11
Adodc1.Recordset.Fields(i) = Text3(i).Text
Next i
 ' 将Text1中的日期和时间部分组合后赋值给Recordset的第12个字段
    Adodc1.Recordset.Fields(12).value = Text1(0).Text & " " & Text1(1).Text & ":" & Text1(2).Text & ":" & Text1(3).Text

    ' 将Text2中的日期和时间部分组合后赋值给Recordset的第13个字段
    Adodc1.Recordset.Fields(13).value = Text2(0).Text & " " & Text2(1).Text & ":" & Text2(2).Text & ":" & Text2(3).Text

    ' 将DTPicker3的值赋给Recordset的第14个字段
    Adodc1.Recordset.Fields(14).value = DTPicker3.value

Adodc1.Recordset.Update
Adodc1.Refresh
For i = 0 To 11
Text3(i).Text = ""
Next i
End Sub

Private Sub Command4_Click()
On Error Resume Next

Adodc1.Recordset.Delete
Adodc1.Refresh
For i = 0 To 11
Text3(i).Text = ""
Next i
End Sub

Private Sub Command1_Click()
    On Error Resume Next  ' 错误处理，遇到错误时跳过错误继续执行

    ' 检查Text3数组中的每个元素是否为空，如果为空，则提示并退出子程序
    Dim i As Integer
    For i = 0 To 11
        If Text3(i).Text = "" Then
            MsgBox "请输入数据"
            Exit Sub
        End If
    Next i

    ' 将Text3数组中的每个文本值转换为数值
    For i = 0 To 11
        Text3(i).Text = Val(Text3(i).Text)
    Next i

    ' 向Recordset添加新记录
    Adodc1.Recordset.AddNew

    ' 将Text3中的数据保存到Recordset的对应字段
    For i = 0 To 11
        Adodc1.Recordset.Fields(i).value = Text3(i).Text
    Next i

    ' 将Text1中的日期和时间部分组合后赋值给Recordset的第12个字段
    Adodc1.Recordset.Fields(12).value = Text1(0).Text & " " & Text1(1).Text & ":" & Text1(2).Text & ":" & Text1(3).Text

    ' 将Text2中的日期和时间部分组合后赋值给Recordset的第13个字段
    Adodc1.Recordset.Fields(13).value = Text2(0).Text & " " & Text2(1).Text & ":" & Text2(2).Text & ":" & Text2(3).Text

    ' 将DTPicker3的值赋给Recordset的第14个字段
    Adodc1.Recordset.Fields(14).value = DTPicker3.value

    ' 更新并刷新Recordset以保存更改
    Adodc1.Recordset.Update
    Adodc1.Refresh

    ' 清空Text3控件
    For i = 0 To 11
        Text3(i).Text = ""
    Next i
End Sub



