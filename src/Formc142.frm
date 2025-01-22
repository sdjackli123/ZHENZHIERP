VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formc142 
   BackColor       =   &H00C0E0FF&
   Caption         =   "锅号信息"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14220
   LinkTopic       =   "Form14"
   ScaleHeight     =   8580
   ScaleWidth      =   14220
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   1
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Text3"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   0
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text3"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   3
      Left            =   10320
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   10320
      Top             =   120
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "选择发货"
      Height          =   495
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "取消"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询信息"
      Height          =   1095
      Left            =   480
      TabIndex        =   15
      Top             =   4800
      Width           =   4815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "未发"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "已发"
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   2
      Left            =   7680
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   1
      Left            =   5040
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      ItemData        =   "Formc142.frx":0000
      Left            =   9840
      List            =   "Formc142.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10680
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Bindings        =   "Formc142.frx":0004
      Height          =   2655
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   13215
      _cx             =   23310
      _cy             =   4683
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
   Begin VB.TextBox Text1 
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
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   10680
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   10800
      Top             =   -120
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formc142.frx":0019
      Height          =   1695
      Left            =   480
      TabIndex        =   21
      Top             =   6240
      Width           =   9255
      _cx             =   16325
      _cy             =   2990
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
      FormatString    =   $"Formc142.frx":002E
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
      Height          =   495
      Index           =   6
      Left            =   6360
      TabIndex        =   26
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "缸号"
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
      Index           =   5
      Left            =   4080
      TabIndex        =   24
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发货米数"
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
      Index           =   4
      Left            =   9240
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发货重量"
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
      Index           =   3
      Left            =   6600
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发货匹数"
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
      Index           =   2
      Left            =   3960
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "发货单号"
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
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
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
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Formc142"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public gygh As String

Private Sub Command1_Click()
On Error Resume Next
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
ph = Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1)
gpps = gpps + 1
gpzl = gpzl + Val(Mid(List2.List(i), InStr(List2.List(i), "-") + 1))
sql1 = "update bmd set 单据='' where 锅号='" & Text1.Text & "' and 匹号='" & ph & "' and 缸号='" & Text3(0).Text & "' and 序号='" & Text3(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
MsgBox ("取消成功！")
Call Command6_Click
End Sub

Private Sub Command11_Click()
If Text1 = "" Or Text3(0) = "" Then
MsgBox ("请输入锅号和缸号")
Exit Sub
End If
Call fhdmd(Adodc3, Text1, Text3(0), Text2(0))
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If fhxz = 15 Then
Formc15.DataCombo1.Text = Adodc1.Recordset.Fields(0)   ''''客户
Formc15.DataCombo2.Text = Adodc1.Recordset.Fields(4)   ''品名
Formc15.DataCombo3.Text = Adodc1.Recordset.Fields(6)   '颜色
Formc15.DataCombo4.Text = Adodc1.Recordset.Fields(3)   '锅号
Formc15.DataCombo5.Text = Adodc1.Recordset.Fields(15)   '毛坯重量
Formc15.DataCombo11.Text = Adodc1.Recordset.Fields(2)  '款号
Formc15.Text7.Text = Text2(1)      ' 毛坯匹数
Formc15.DataCombo16.Text = Adodc1.Recordset.Fields(1) '单号
Formc15.Text11(2).Text = Adodc1.Recordset.Fields(5)  ''成分
Formc15.Text11(3).Text = Adodc1.Recordset.Fields(9)  ''单位
Formc15.Text8.Text = Adodc1.Recordset.Fields(12)   ''''单价
Formc15.Text9.Text = Adodc1.Recordset.Fields(7)   ''''色号
Formc15.Text10.Text = Text2(2)   ''''光坯重量
Formc15.Text12.Text = Text2(3)   ''''光坯数量
Unload Me
End If
End Sub

Private Sub Command4_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = False
Next
End Sub

Private Sub Command5_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = True
Next
End Sub

Private Sub Command6_Click()
If Option1(1).value = True Then
Adodc2.RecordSource = "select 匹号,光胚重量 from bmd where 锅号='" & Text1 & "' and isnull(单据,'')='' and 缸号='" & Text3(0) & "' and 序号='" & Text3(1) & "' order by 匹号"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List2.Clear
Do While Not Adodc2.Recordset.EOF
List2.AddItem Trim(Adodc2.Recordset.Fields(0)) + "-" + Trim(Adodc2.Recordset.Fields(1))
Adodc2.Recordset.MoveNext
Loop
End If
If Option1(0).value = True Then
Adodc2.RecordSource = "select 匹号,光胚重量 from bmd where 锅号='" & Text1 & "' and isnull(单据,'')<>'' and 缸号='" & Text3(0) & "' and 序号='" & Text3(1) & "' order by 匹号"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Adodc2.Recordset.MoveFirst
List2.Clear
Do While Not Adodc2.Recordset.EOF
List2.AddItem Trim(Adodc2.Recordset.Fields(0)) + "-" + Trim(Adodc2.Recordset.Fields(1))
Adodc2.Recordset.MoveNext
Loop
End If
Adodc4.RecordSource = "select distinct 缸号,序号,单据 from bmd  WHERE 锅号='" & Text1.Text & "'"
Adodc4.Refresh
Command1.Enabled = True
Command7.Enabled = True
End Sub

Private Sub Command7_Click()
If Text2(0) = "" Then
MsgBox ("请输入发货单号")
Exit Sub
End If
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
ph = Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1)
gpps = gpps + 1
gpzl = gpzl + Val(Mid(List2.List(i), InStr(List2.List(i), "-") + 1))
sql1 = "update bmd set 单据='" & Text2(0) & "' where 锅号='" & Text1 & "' and 缸号='" & Text3(0).Text & "' and 序号='" & Text3(1) & "' and 匹号='" & ph & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
Text2(1) = gpps
Text2(2) = gpzl
MsgBox ("确认成功！")
Call Command6_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Option1(1).value = True
For i = 1 To 3
Text2(i) = ""
Text3(i - 1) = ""
Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(11) = 1500
VSFlexGrid1.ColWidth(12) = 0
VSFlexGrid1.ColWidth(13) = 0

VSFlexGrid2.ColWidth(0) = 100
VSFlexGrid2.ColWidth(3) = 1500
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
cl = VSFlexGrid1.col
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
Text3(0) = Adodc1.Recordset.Fields(13)
Text3(1) = Adodc1.Recordset.Fields(14)
End Sub

Private Sub Text1_Change()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 客户名称,锅号,款号,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,技术要求 as 克重,''as 分备注,''as 总备注,'' as 合同部门,编号 as 缸号,序号 from v_kpd_cx  WHERE 锅号='" & Text1.Text & "'"
Adodc1.Refresh
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select distinct 缸号,序号,isnull(单据,'') as 单据 from bmd  WHERE 锅号='" & Text1.Text & "'"
Adodc4.Refresh
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If
End Sub
Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc4.Recordset.EOF Then Exit Sub
Adodc4.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc4.Recordset.Move rs - 1
Text2(0) = Adodc4.Recordset.Fields(2)
Text3(0) = Adodc4.Recordset.Fields(0)
Text3(1) = Adodc4.Recordset.Fields(1)
Command1.Enabled = False
Command7.Enabled = False
End Sub
