VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formc151 
   BackColor       =   &H00C0E0FF&
   Caption         =   "发货客户信息"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   14190
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
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
      Height          =   615
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   11
      Left            =   6840
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   3960
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   7440
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
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   615
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   615
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   615
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      Height          =   615
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   10
      Left            =   6840
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   9
      Left            =   6840
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   8
      Left            =   6840
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   7
      Left            =   6840
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   6
      Left            =   6840
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   2160
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   2160
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   2160
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formc151.frx":0000
      Height          =   2055
      Left            =   600
      TabIndex        =   22
      Top             =   4800
      Width           =   13095
      _cx             =   23098
      _cy             =   3625
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
      FormatString    =   $"Formc151.frx":0015
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
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "收货单位"
      Height          =   495
      Index           =   11
      Left            =   5280
      TabIndex        =   28
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "发货电话"
      Height          =   495
      Index           =   10
      Left            =   5280
      TabIndex        =   20
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "发货人"
      Height          =   495
      Index           =   9
      Left            =   5280
      TabIndex        =   18
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "运费金额"
      Height          =   495
      Index           =   8
      Left            =   5280
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "路途公里"
      Height          =   495
      Index           =   7
      Left            =   5280
      TabIndex        =   14
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "包装形式"
      Height          =   495
      Index           =   6
      Left            =   5280
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "包装体积"
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "联系电话"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "联系人"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "配货自提"
      Height          =   495
      Index           =   5
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "收货地址"
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "发货单据"
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Formc151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
Formc152.Show
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 0
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * from jgmxfhtz where 发货单据='" & Text1(0) & "'"
Adodc1.Refresh
End Select
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub VSFlexGrid2_DblClick()
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc1.Recordset.Move rs - 1
For i = 1 To 10
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
For i = 0 To 11
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Text1(0).SetFocus
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Form_Load()

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * from jgmxfhtz where 发货单据='" & Text1(0) & "'"
Adodc1.Refresh
For i = 0 To 11
Text1(i).Text = ""
Next
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub
Private Sub Command1_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
For i = 0 To 11
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

