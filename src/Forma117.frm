VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forma117 
   BackColor       =   &H00C0E0FF&
   Caption         =   "返修原因"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   12600
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1080
      Top             =   7560
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
   Begin VB.CommandButton Command4 
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
      Height          =   495
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Height          =   495
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   8280
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   8280
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   8280
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   8280
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   600
      Width           =   3495
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma117.frx":0000
      Height          =   3375
      Left            =   840
      TabIndex        =   18
      Top             =   3840
      Width           =   10935
      _cx             =   19288
      _cy             =   5953
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
      AllowUserResizing=   4
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
      Caption         =   "返修完成时间"
      Height          =   375
      Index           =   8
      Left            =   6960
      TabIndex        =   16
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "返修开单时间"
      Height          =   375
      Index           =   7
      Left            =   6960
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "厂长确认"
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "生产确认"
      Height          =   375
      Index           =   5
      Left            =   6960
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "计划要求"
      Height          =   375
      Index           =   4
      Left            =   840
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "返修锅号"
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "返修部门"
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "责任部门"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "返修原因"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "Forma117"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
For i = 0 To 8
Adodc1.Recordset.Fields(i) = Text1(i)
Next
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
For i = 0 To 8
Adodc1.Recordset.Fields(i) = Text1(i)
Next
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
For i = 0 To 6
Text1(i).Text = ""
Next
Text1(7) = Now
Text1(8) = Now
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM kpdfx WHERE 返修锅号='" & Text1(0).Text & "'"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(8) = 1800
VSFlexGrid1.ColWidth(9) = 1800
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 7
       Text1(7) = Now
       Case 8
       Text1(8) = Now
End Select
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 0
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM kpdfx WHERE 返修锅号='" & Text1(0).Text & "'"
Adodc1.Refresh
End Select
End Sub

Private Sub VSFlexGrid1_dblClick()
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
For i = 0 To 8
Text1(i) = Adodc1.Recordset.Fields(i)
Next
End Sub
