VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formd221 
   BackColor       =   &H00C0E0FF&
   Caption         =   "大货工艺配方单信息"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form21"
   ScaleHeight     =   9135
   ScaleWidth      =   10605
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1200
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   6240
      Top             =   8400
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   375
      Left            =   6600
      Top             =   8400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   375
      Left            =   3960
      Top             =   8400
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
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Left            =   8640
      Top             =   0
   End
   Begin VB.Data Data10 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "配方单"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formd221.frx":0000
      Height          =   5175
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   9855
      _cx             =   17383
      _cy             =   9128
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
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "吸水率"
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
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "编号"
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
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号快捷"
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
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Formd221"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer: Dim BA As Database: Dim RD As Recordset: Dim sz(6) As String: Dim ZS(7) As String

Private Sub Command1_Click()
Adodc5.RecordSource = "SELECT RQ AS 日期,KH AS 客户,fz AS 工艺负责人,YS AS 颜色,SH AS 色号,BL AS 品名,DH AS 配方编号,ZL AS 生产类别,IP AS 浴比 FROM ZH WHERE SH LIKE '%'+'" & Text1.Text & "'+'%'  ORDER BY SH DESC"
Adodc5.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Data10.Database.Execute "delete * from pfda"

Adodc11.RecordSource = "select * from pfd where  编号='" & Text2.Text & "'"
Adodc11.Refresh
If Not Adodc11.Recordset.EOF Then
Adodc11.Recordset.MoveFirst

For i = 0 To 6
ZS(i) = Adodc11.Recordset.Fields(i)
Next
ZS(7) = "大货"

mb = 0
For i = 7 To 56
If Adodc11.Recordset.Fields(i) <> "" Then
mb = mb + 1
End If
Next

ProgressBar1.Visible = True
Timer1.Enabled = True
For i = 7 To mb + 7
If Adodc11.Recordset.Fields(i) <> "" Then
sz(0) = Mid(Adodc11.Recordset.Fields(i), 1, InStr(Adodc11.Recordset.Fields(i), "(") - 1)
sz(1) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "(") + 1, InStr(Adodc11.Recordset.Fields(i), ")") - InStr(Adodc11.Recordset.Fields(i), "(") - 1)
sz(2) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), ")") + 1, InStr(Adodc11.Recordset.Fields(i), "-") - InStr(Adodc11.Recordset.Fields(i), ")") - 1)
sz(3) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "-") + 1, InStr(Adodc11.Recordset.Fields(i), "\") - InStr(Adodc11.Recordset.Fields(i), "-") - 1)
sz(4) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "\") + 1, InStr(Adodc11.Recordset.Fields(i), "#") - InStr(Adodc11.Recordset.Fields(i), "\") - 1)
sz(5) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "#") + 1, InStr(Adodc11.Recordset.Fields(i), "^") - InStr(Adodc11.Recordset.Fields(i), "#") - 1)
sz(6) = Mid(Adodc11.Recordset.Fields(i), InStr(Adodc11.Recordset.Fields(i), "^") + 1)
L = i - 6

Data10.Database.Execute "insert into pfda(加工单位,品名,色号,颜色,配方编号,负责人,配方日期,生产种类,工序名称,浴比,染化助库,染化助名称,单位,配方,车速,次序号) VALUES('" & ZS(0) & "','" & ZS(1) & "','" & ZS(2) & "','" & ZS(3) & "','" & ZS(4) & "','" & ZS(5) & "','" & ZS(6) & "','" & ZS(7) & "','" & sz(0) & "','" & sz(1) & "','" & sz(2) & "',  " & _
                                                                       "'" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & L & "')"
ProgressBar1.value = 100 / mb * L
End If
Next
ProgressBar1.Visible = False
Timer1.Enabled = False
End If


Formd331.Text1.Text = Text2.Text
Data10.Database.Execute "INSERT INTO  pldd(工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,次序号,批次,车速,料单编号) SELECT pfda.工序名称,pfda.浴比,pfda.染化助库,pfda.染化助名称,pfda.单位,pfda.配方,'1',pfda.次序号,批次,车速,'" & Formd331.Text2 & "' From pfda WHERE 配方编号='" & Text2.Text & "'"

Formd331.Data13.Refresh
Formd331.VSFlexGrid1.ColWidth(0) = 400
Formd331.VSFlexGrid1.ColWidth(1) = 0
Formd331.VSFlexGrid1.ColWidth(2) = 0
Formd331.VSFlexGrid1.ColWidth(5) = 400
Formd331.VSFlexGrid1.ColWidth(7) = 2000
Formd331.VSFlexGrid1.ColWidth(8) = 800
Formd331.VSFlexGrid1.ColWidth(10) = 600
Formd331.VSFlexGrid1.ColWidth(13) = 0
Formd331.VSFlexGrid1.ColWidth(14) = 0
Formd331.VSFlexGrid1.ColWidth(17) = 0
Formd331.VSFlexGrid1.ColWidth(18) = 0
Formd331.VSFlexGrid1.ColWidth(19) = 2600
Formd331.VSFlexGrid1.ColWidth(20) = 0
Formd331.VSFlexGrid1.ColWidth(22) = 0
Formd331.VSFlexGrid1.ColWidth(23) = 0
Formd331.VSFlexGrid1.ColWidth(24) = 0
Formd331.VSFlexGrid1.ColWidth(25) = 0
Unload Me
Formd331.Show
End Sub



Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Form_Load()



Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
       Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Data10.DatabaseName = App.Path & "\AccessBase\db.mdb"
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 1200
VSFlexGrid1.ColWidth(5) = 1000
VSFlexGrid1.ColWidth(6) = 1500
VSFlexGrid1.ColWidth(7) = 1500
VSFlexGrid1.ColWidth(8) = 2000
VSFlexGrid1.ColWidth(9) = 1000

End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move rs - 1
Text2.Text = Adodc5.Recordset.Fields(6)
If IsNull(Adodc5.Recordset.Fields(9)) Then
Text3 = ""
Else
Text3.Text = Adodc5.Recordset.Fields(9)
End If
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 2 Then
       Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc5.RecordSource = "SELECT RQ AS 日期,KH AS 客户,fz AS 工艺负责人,YS AS 颜色,SH AS 色号,BL AS 品名,DH AS 配方编号,ZL AS 生产类别,IP AS 浴比,xs as 吸水率 FROM ZH WHERE SH='" & Text1.Text & "' and isnull(qr,'')='审核'"
       Adodc5.Refresh
End If
End Sub
