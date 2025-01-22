VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormA115 
   BackColor       =   &H00C0E0FF&
   Caption         =   "修色复制"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   13455
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3240
      Top             =   7320
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Height          =   330
      Left            =   840
      Top             =   7320
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "复制修色锅号"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text11"
      Top             =   720
      Width           =   2295
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
      Bindings        =   "FormA115.frx":0000
      Height          =   4935
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   11535
      _cx             =   20346
      _cy             =   8705
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
      MergeCells      =   1
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
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请输入正常锅号"
      Height          =   375
      Index           =   15
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "FormA115"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If Text11.Text = "" Then
MsgBox ("请输入正常锅号")
Exit Sub
End If

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select distinct 锅号 from v_kpdb_xs where 锅号 like '%'+'" & Text11.Text & "'+'%'"
Adodc2.Refresh

ll = Text11.Text + "F" + Trim(Adodc2.Recordset.RecordCount + 1)

If MsgBox("要复制的正常锅号" + Text11.Text + "吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "insert into kpd(客户名称,单号,锅号,色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,日期,备注,技术要求,IP,标签,kp,kp1,CKY,负责人,pb,rs,ts,xdx,ddx,fh) select 客户名称,单号,'" & ll & "',色别,品名,毛胚幅宽,光胚幅宽,匹数,重量,类别,'" & Date & "',备注,技术要求,IP,标签,'N','N',CKY,负责人,'Y','N','N','N','N','N' from kpd where 锅号='" & Text11.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc1.RecordSource = "select 客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,标签 as 合约号,备注,技术要求,类别,CKY as 毛坯备注,车台  from kpd where 锅号='" & ll & "' order by IP"
Adodc1.Refresh

End Sub

Private Sub Form_Load()
Text11.Text = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 客户名称,锅号,IP,品名,毛胚幅宽,光胚幅宽,匹数,重量,色别,标签 as 合约号,备注,技术要求,类别,CKY as 毛坯备注,车台,GX AS 工序  from kpd where 锅号='" & Text11.Text & "' "
Adodc1.Refresh
End Sub
