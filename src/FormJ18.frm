VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formj18 
   BackColor       =   &H00C0E0FF&
   Caption         =   "计划"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form18"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   840
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   7440
      Top             =   10800
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
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   7200
      Top             =   10920
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Left            =   7800
      Top             =   10920
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
      Height          =   375
      Left            =   7560
      Top             =   10920
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Left            =   7680
      Top             =   10920
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
      Bindings        =   "FormJ18.frx":0000
      Height          =   8895
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   18735
      _cx             =   33046
      _cy             =   15690
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   3000
      TabIndex        =   13
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "FormJ18.frx":0015
      Height          =   330
      Left            =   1080
      TabIndex        =   12
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FormJ18.frx":002A
      Height          =   330
      Left            =   1080
      TabIndex        =   11
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "车台编号"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "备活"
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
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
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "调整"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   9720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8880
      Top             =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号"
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
      Left            =   3000
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "车台"
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
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "总计锅数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户名称"
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
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Formj18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r, Z As Integer: Public GS As Integer: Public DH As String ''''''锅数变量\单号变量

Private Sub Command1_Click()
Z = 1
       Combo1111.Clear
       Combo1111.AddItem "备活"
       Combo1111.AddItem "取消"

End Sub

Private Sub Command2_Click()
Z = 0
       Combo1111.Clear

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc2.Refresh
p = 1
L = "备活"
m = "就绪"
GS = 1
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
For i = 1 To 30

If InStr(Trim(Adodc2.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc2.Recordset.Fields(i)), m) <> 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i + 1
    VSFlexGrid1.CellForeColor = vbGreen
End If

If InStr(Trim(Adodc2.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc2.Recordset.Fields(i)), m) <> 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i + 1
    VSFlexGrid1.CellForeColor = vbRed
End If

If InStr(Trim(Adodc2.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc2.Recordset.Fields(i)), m) = 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i + 1
    VSFlexGrid1.CellForeColor = vbCyan
End If

If InStr(Trim(Adodc2.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc2.Recordset.Fields(i)), m) = 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i + 1
    VSFlexGrid1.CellForeColor = vbBlack
End If

Next

For i = 1 To 30
If InStr(Trim(Adodc2.Recordset.Fields(i)), DataCombo2.Text) > 0 And InStr(Trim(Adodc2.Recordset.Fields(i)), DataCombo3.Text) > 0 Then
   Else
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i + 1
    VSFlexGrid1.Text = ""
End If

If InStr(Trim(Adodc2.Recordset.Fields(i)), DataCombo2.Text) > 0 Then
GS = GS + 1
End If

Next

Adodc2.Recordset.MoveNext
p = p + 1
Loop
Adodc1.Refresh
Label3.Caption = GS - 1

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 800
For i = 2 To 30
VSFlexGrid1.ColWidth(i) = 3000
Next

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 1600
Next
End If

End Sub

Private Sub Command6_Click()
Call jhbOutadodcToExcel(VSFlexGrid1, "日期：" + Trim(Now))
End Sub

Private Sub DataCombo1_Change()
On Error Resume Next
If DataCombo1.Text = "" Then
Adodc2.RecordSource = "SELECT * FROM v_rgjhb  ORDER BY 车台位区,车台编号"
Adodc2.Refresh
Adodc1.RecordSource = "SELECT * FROM v_rgjhb  ORDER BY 车台位区,车台编号"
Adodc1.Refresh
Else
Adodc2.RecordSource = "SELECT * FROM v_rgjhb WHERE 车台编号='" & DataCombo1.Text & "' ORDER BY 车台位区,车台编号"
Adodc2.Refresh
Adodc1.RecordSource = "SELECT * FROM v_rgjhb WHERE 车台编号='" & DataCombo1.Text & "' ORDER BY 车台位区,车台编号"
Adodc1.Refresh
End If

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 800
For i = 2 To 30
VSFlexGrid1.ColWidth(i) = 3000
Next

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 1600
Next
End If
End Sub

Private Sub DataCombo1_Click(Area As Integer)
On Error Resume Next
If DataCombo1.Text = "" Then
Adodc2.RecordSource = "SELECT * FROM v_rgjhb  ORDER BY 车台位区,车台编号"
Adodc2.Refresh
Adodc1.RecordSource = "SELECT * FROM v_rgjhb  ORDER BY 车台位区,车台编号"
Adodc1.Refresh
Else
Adodc2.RecordSource = "SELECT * FROM v_rgjhb WHERE 车台编号='" & DataCombo1.Text & "' ORDER BY 车台位区,车台编号"
Adodc2.Refresh
Adodc1.RecordSource = "SELECT * FROM v_rgjhb WHERE 车台编号='" & DataCombo1.Text & "' ORDER BY 车台位区,车台编号"
Adodc1.Refresh
End If
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 800
For i = 2 To 30
VSFlexGrid1.ColWidth(i) = 3000
Next

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 1600
Next
End If
End Sub


Private Sub Form_Load()

'On Error Resume Next
Text1.Text = ""
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM v_rgjhb  ORDER BY 车台位区,车台编号"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT * FROM v_rgjhb  ORDER BY 车台位区,车台编号"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT 车台编号 FROM CT  GROUP BY 车台编号"
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT distinct 简称 FROM khzl  order BY 简称"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"



Z = 1
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 800
For i = 2 To 7
VSFlexGrid1.ColWidth(i) = 3000
Next

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 1600
Next
End If
End Sub

Private Sub Label4_Click()
DataCombo1.Text = ""
End Sub

Private Sub Text1_Change()

If Text1.Text = "" Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM v_rgjhb  ORDER BY 车台位区,车台编号"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT * FROM v_rgjhb  ORDER BY 车台位区,车台编号"
Adodc2.Refresh
Else
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM v_rgjhb WHERE 车台编号='" & Text1.Text & "' ORDER BY 车台位区,车台编号"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT * FROM v_rgjhb  WHERE 车台编号='" & Text1.Text & "' ORDER BY 车台位区,车台编号"
Adodc2.Refresh
End If

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Adodc2.Refresh
p = 1
L = "备活"
m = "就绪"
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
For i = 1 To 30

If InStr(Trim(Adodc2.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc2.Recordset.Fields(i)), m) <> 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i + 1
    VSFlexGrid1.CellForeColor = vbGreen
End If

If InStr(Trim(Adodc2.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc2.Recordset.Fields(i)), m) <> 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i + 1
    VSFlexGrid1.CellForeColor = vbRed
End If

If InStr(Trim(Adodc2.Recordset.Fields(i)), L) <> 0 And InStr(Trim(Adodc2.Recordset.Fields(i)), m) = 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i + 1
    VSFlexGrid1.CellForeColor = vbCyan
End If

If InStr(Trim(Adodc2.Recordset.Fields(i)), L) = 0 And InStr(Trim(Adodc2.Recordset.Fields(i)), m) = 0 Then
    VSFlexGrid1.Row = p
    VSFlexGrid1.col = i + 1
    VSFlexGrid1.CellForeColor = vbBlack
End If

Next
Adodc2.Recordset.MoveNext
p = p + 1
Loop
End Sub
