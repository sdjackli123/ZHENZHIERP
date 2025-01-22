VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma901 
   BackColor       =   &H00C0E0FF&
   Caption         =   "欠费预警信息"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   14100
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   9360
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9960
      Top             =   720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "显示信息"
      Height          =   615
      Left            =   960
      TabIndex        =   7
      Top             =   1680
      Width           =   8175
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "低于欠费下限"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "全部"
         Height          =   255
         Left            =   6360
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "欠费上下限之间"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "超出欠费上限"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   615
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma901.frx":0000
      Height          =   450
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2040
      Top             =   8160
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma901.frx":0015
      Height          =   5535
      Left            =   960
      TabIndex        =   4
      Top             =   2520
      Width           =   11655
      _cx             =   20558
      _cy             =   9763
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
      Left            =   2400
      Top             =   7680
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "客户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Forma901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sj As Integer
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "欠费信息")
End Sub

Private Sub Command6_Click()
If Option1.value = True Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * from yj_qfts  where 欠费>=欠费上限 order by 客户"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * from yj_qfts  where  欠费>=欠费上限 and 客户='" & DataCombo1.Text & "' order by 客户"
Adodc1.Refresh
End If
End If
If Option2.value = True Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * from yj_qfts  where 欠费<欠费上限 and  欠费>欠费下限 order by 客户"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * from yj_qfts  where  欠费<欠费上限 and  欠费>欠费下限 and 客户='" & DataCombo1.Text & "' order by 客户"
Adodc1.Refresh
End If
End If
If Option3.value = True Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * from yj_qfts  order by 客户"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * from yj_qfts  where 客户='" & DataCombo1.Text & "' order by 客户"
Adodc1.Refresh
End If
End If
If Option4.value = True Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * from yj_qfts  where 欠费<欠费下限 order by 客户"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * from yj_qfts  where  欠费<欠费下限 and  客户='" & DataCombo1.Text & "' order by 客户"
Adodc1.Refresh
End If
End If

End Sub

Private Sub Form_Load()

Text2.Text = ""
DataCombo1.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * from yj_qfts  where 欠费>=欠费上限 order by 客户"
Adodc1.Refresh
Option1.value = True
VSFlexGrid1.ColWidth(1) = 2000
VSFlexGrid1.ColWidth(2) = 2000
VSFlexGrid1.ColWidth(3) = 2000
VSFlexGrid1.ColWidth(4) = 2000
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then Exit Sub
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select distinct 简称 from KHZL where 代码 like '%'+'" & Text2.Text & "'+'%'  order by 简称"
Adodc2.Refresh
End Sub

Private Sub Timer1_Timer()
If sj = 5 Then
sj = 1
Call Command6_Click
Else
sj = sj + 1
End If
End Sub

Private Sub Timer2_Timer()
    Dim i
    i = FlashWindow(Me.hwnd, 1) '定时闪烁标题栏
End Sub
