VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forma105 
   BackColor       =   &H00C0E0FF&
   Caption         =   "配比设置"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   14415
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   7560
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   7560
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   1095
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   7560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   8160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "复制"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   8400
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   7560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   10
      Left            =   7560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2160
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6120
      Top             =   10080
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
      Left            =   6120
      Top             =   9720
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
      Left            =   6120
      Top             =   9840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   330
      Left            =   5880
      Top             =   9960
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
      Bindings        =   "Forma105.frx":0000
      Height          =   4455
      Left            =   720
      TabIndex        =   20
      Top             =   3600
      Width           =   12975
      _cx             =   22886
      _cy             =   7858
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
      GridLines       =   2
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
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "纱量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   13
      Left            =   6000
      TabIndex        =   33
      Top             =   960
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "织号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   720
      TabIndex        =   32
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "配比"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   6000
      TabIndex        =   31
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "织耗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   720
      TabIndex        =   30
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "排产"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   720
      TabIndex        =   29
      Top             =   960
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "纱支"
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
      TabIndex        =   28
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "批次"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   720
      TabIndex        =   27
      Top             =   2160
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "备注"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   6000
      TabIndex        =   26
      Top             =   2760
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "单号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   720
      TabIndex        =   25
      Top             =   8160
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   4320
      TabIndex        =   24
      Top             =   360
      Width           =   585
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFF80&
      Caption         =   "复制织号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   5160
      TabIndex        =   23
      Top             =   8400
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "产地"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   6000
      TabIndex        =   22
      Top             =   2160
      Width           =   1545
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C0C0&
      Caption         =   "颜色"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   10
      Left            =   6000
      TabIndex        =   21
      Top             =   1560
      Width           =   1545
   End
End
Attribute VB_Name = "Forma105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command1_Click()
On Error Resume Next
If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(1).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If
Adodc1.Recordset.AddNew
For i = 0 To 10
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc3.RecordSource = "select isnull(max(序号),0) FROM sxpb WHERE 织号='" & Text1(0).Text & "'"
Adodc3.Refresh
Text1(8) = Adodc3.Recordset.Fields(0) + 1

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定修改吗", vbYesNo) = vbNo Then Exit Sub
If Text1(0).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(1).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If
For i = 0 To 10
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc3.RecordSource = "select isnull(max(序号),0) FROM sxpb WHERE 织号='" & Text1(0).Text & "'"
Adodc3.Refresh
Text1(8) = Adodc3.Recordset.Fields(0) + 1
Command1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True

End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除吗", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Adodc3.RecordSource = "select isnull(max(序号),0) FROM sxpb WHERE 织号='" & Text1(0).Text & "'"
Adodc3.Refresh
Text1(8) = Adodc3.Recordset.Fields(0) + 1
Command1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command5_Click()
sql1 = "insert into  sxpb(织号,排产,纱支,织耗,配比,纱量,批次,备注,序号) select '" & Text1(0) & "','" & Text1(1) & "',纱支,织耗,配比,纱量,批次,备注,序号 from sxpb where 织号='" & Text3 & "'"
sql2 = "update  sxpb set 纱量=排产/(100-cast(isnull(织耗,0) as real))*配比 where 织号='" & Text1(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("复制成功！")
Adodc1.RecordSource = "SELECT * FROM sxpb WHERE 织号='" & Text1(0).Text & "'"
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
Adodc3.RecordSource = "select isnull(max(序号),0) FROM sxpb WHERE 织号='" & Text1(0).Text & "'"
Adodc3.Refresh
Text1(8) = Adodc3.Recordset.Fields(0) + 1
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command7_Click()
Me.Hide
End Sub

Private Sub Form_Load()
For i = 0 To 10
Text1(i).Text = ""
Next
Text2.Text = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM sxpb WHERE 织号='" & Text1(0).Text & "'"
Adodc1.Refresh
Text3 = ""
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select isnull(max(序号),0) FROM sxpb WHERE 织号='" & Text1(0).Text & "'"
Adodc3.Refresh
Text1(8) = Adodc3.Recordset.Fields(0) + 1
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1600
VSFlexGrid1.ColWidth(2) = 2600
VSFlexGrid1.ColWidth(3) = 1000
VSFlexGrid1.ColWidth(4) = 1000
VSFlexGrid1.ColWidth(5) = 1600
VSFlexGrid1.ColWidth(6) = 1600
End Sub


Private Sub Label1_Click()
clbl = 3
Formy59.DataCombo3 = "纱线库"
'Formy59.DataCombo6.Text = Text2.Text
'Formy59.Check2(2).Value = 1
Formy59.Check2(4).value = 1
Formy59.Show
End Sub

Private Sub lblLabels_Click(Index As Integer)
Select Case Index
       Case 8
Forma103.Show
End Select
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM sxpb WHERE 织号='" & Text1(0).Text & "'"
Adodc1.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select isnull(max(序号),0) FROM sxpb WHERE 织号='" & Text1(0).Text & "'"
Adodc3.Refresh
Text1(8) = Adodc3.Recordset.Fields(0) + 1

       Case 1, 3, 4
If Val(Text1(3).Text) > 0 Then
Text1(5).Text = Format(Val(Text1(1).Text) / (100 - Val(Text1(3).Text)) * 100 * Val(Text1(4).Text) / 100, "#0.00")
End If
End Select

End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To 10
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
DTPicker1.value = Adodc1.Recordset.Fields(6)
Command1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
End Sub


