VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formw113 
   BackColor       =   &H00C0E0FF&
   Caption         =   "资产负债表设置"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12255
   LinkTopic       =   "Form3"
   ScaleHeight     =   10215
   ScaleWidth      =   12255
   StartUpPosition =   2  '屏幕中心
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw113.frx":0000
      Height          =   7095
      Left            =   720
      TabIndex        =   19
      Top             =   3000
      Width           =   10815
      _cx             =   19076
      _cy             =   12515
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
      AllowUserResizing=   0
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
      WordWrap        =   0   'False
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
      AllowUserFreezing=   0
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
      Left            =   7200
      Top             =   10200
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      ItemData        =   "Formw113.frx":0015
      Left            =   6000
      List            =   "Formw113.frx":0022
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   4
      Left            =   7200
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   0
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   1
      Left            =   3360
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   3
      Left            =   7200
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1575
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1575
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   " -"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   18
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "取数科目公式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   15
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "报表标题"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   720
      TabIndex        =   13
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   7200
      TabIndex        =   12
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "取数会计科目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "单元格"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   7200
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "Formw113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Text1(1).Text = "" Then Exit Sub
Text1(1).Text = Text1(1).Text + "." + Combo1.Text
End Sub

Private Sub Command5_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "资产负债表设置")
End Sub

Private Sub Command6_Click()
Adodc1.Refresh
Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
Text1(3).Text = 1
Text1(3).Text = Adodc1.Recordset.RecordCount + 1
Text1(0).SetFocus
Text1(4).Text = ""
End Sub

Private Sub Label1_Click()
KMMC = 7
Formw61.Show
End Sub

Private Sub Label4_Click()
If Text1(2).Text <> "" Then
Text1(2).Text = Text1(2).Text + "+" + Text1(1).Text
Else
Text1(2).Text = Text1(2).Text + Text1(1).Text
End If
End Sub

Private Sub Label5_Click()
If Text1(2).Text <> "" Then
Text1(2).Text = Text1(2).Text + "-" + Text1(1).Text
Else
Text1(2).Text = Text1(2).Text + Text1(1).Text
End If
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
rs = VSFlexGrid1.Row
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To 4
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
End Sub
Private Sub Command3_Click()
If MsgBox("确认修改吗?", vbYesNo) = vbNo Then Exit Sub

For i = 0 To 4
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
Text1(3).Text = 1
Text1(3).Text = Adodc1.Recordset.RecordCount + 1
Text1(0).SetFocus
Text1(4).Text = ""
End Sub

Private Sub Command4_Click()
If MsgBox("确认删除吗?", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
Text1(3).Text = 1
Text1(3).Text = Adodc1.Recordset.RecordCount + 1
Text1(0).SetFocus
Text1(4).Text = ""
End Sub

Private Sub Form_Load()

On Error Resume Next
Combo1.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from FZB order by 序号 DESC"
Adodc1.Refresh
For i = 0 To 4
Text1(i).Text = ""
Next
Text1(3).Text = 1
Text1(3).Text = Adodc1.Recordset.RecordCount + 1
VSFlexGrid1.ColWidth(1) = 2000
VSFlexGrid1.ColWidth(2) = 3000
VSFlexGrid1.ColWidth(3) = 2000
Text1(0).TabIndex = 0
End Sub
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
For i = 0 To 4
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
Text1(3).Text = 1
Text1(3).Text = Adodc1.Recordset.RecordCount + 1
Text1(0).SetFocus
Text1(4).Text = ""
End Sub
Private Sub Command2_Click()
Unload Me
End Sub




