VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formy58 
   BackColor       =   &H00C0E0FF&
   Caption         =   "材料信息"
   ClientHeight    =   11730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11730
   ScaleWidth      =   15000
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   360
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1320
      Top             =   7560
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Bindings        =   "Formy58.frx":0000
      Height          =   9735
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   14055
      _cx             =   24791
      _cy             =   17171
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
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
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "库类"
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
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "按名称"
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
      Index           =   2
      Left            =   6960
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "按简码"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "简码查询"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "名称查询"
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Formy58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from v_clmc where   库类 like '%'+'" & Text3.Text & "'+'%' order by 库类"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 4000
VSFlexGrid1.ColWidth(8) = 3000

End Sub

Private Sub Label2_Click()
Adodc1.RecordSource = "select * from v_clmc where 简码 LIKE '%'+'" & Text1.Text & "'+'%' and  库类 like '%'+'" & Text3.Text & "'+'%'"
Adodc1.Refresh
End Sub

Private Sub Label3_Click()
Adodc1.RecordSource = "select * from v_clmc where 材料名称 LIKE '%'+'" & Text2.Text & "'+'%' and  库类 like '%'+'" & Text3.Text & "'+'%'"
Adodc1.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1

If clbl = 1 Then
Formy121.DataCombo1(3).Text = Adodc1.Recordset.Fields(0)
'Formy121.DataCombo1(4).Text = Adodc1.Recordset.Fields(1)
Formy121.DataCombo1(6).Text = Adodc1.Recordset.Fields(3)
Formy121.DataCombo1(5).Text = Adodc1.Recordset.Fields(2)
Formy121.DataCombo1(15).Text = Adodc1.Recordset.Fields(4)
Formy121.DataCombo1(9).Text = Adodc1.Recordset.Fields(6)
Formy121.DataCombo1(3).SetFocus
Unload Me
End If

If clbl = 2 Then
Formy1.DataCombo1(1).Text = Adodc1.Recordset.Fields(0)
Formy1.DataCombo1(2).Text = Adodc1.Recordset.Fields(1)
Formy1.DataCombo1(3).Text = Adodc1.Recordset.Fields(2)
Formy1.DataCombo1(4).Text = Adodc1.Recordset.Fields(3)
Formy1.DataCombo1(6).Text = Adodc1.Recordset.Fields(6)
Formy1.DataCombo1(10).Text = Adodc1.Recordset.Fields(4)
Formy1.DataCombo1(0).Text = Adodc1.Recordset.Fields(7)
Unload Me
End If

If clbl = 3 Then
Formy101.DataCombo1(1).Text = Adodc1.Recordset.Fields(0)
Formy101.DataCombo1(2).Text = Adodc1.Recordset.Fields(1)
Formy101.DataCombo1(3).Text = Adodc1.Recordset.Fields(2)
Formy101.DataCombo1(4).Text = Adodc1.Recordset.Fields(3)
Formy101.DataCombo1(6).Text = Adodc1.Recordset.Fields(6)
Formy101.DataCombo1(10).Text = Adodc1.Recordset.Fields(4)
Formy101.DataCombo1(0).Text = Adodc1.Recordset.Fields(7)
Unload Me
End If

End Sub

Private Sub Text1_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from v_clmc where 简码 LIKE '%'+'" & Text1.Text & "'+'%' and  库类 like '%'+'" & Text3.Text & "'+'%'"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 4000
VSFlexGrid1.ColWidth(8) = 3000

End Sub

Private Sub Text2_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from v_clmc where 材料名称 LIKE '%'+'" & Text2.Text & "'+'%' and  库类 like '%'+'" & Text3.Text & "'+'%'"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 4000
VSFlexGrid1.ColWidth(8) = 3000

End Sub

