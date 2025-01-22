VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formw1128 
   BackColor       =   &H00C0E0FF&
   Caption         =   "复核员信息"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   LinkTopic       =   "Form28"
   ScaleHeight     =   7920
   ScaleWidth      =   7185
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转新账本"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转新凭证"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   8520
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formw1128.frx":0000
      Height          =   4575
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   5655
      _cx             =   9975
      _cy             =   8070
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
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   360
      Width           =   1575
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   480
      Top             =   8280
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   480
      Top             =   8280
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "姓名"
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
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
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
      Left            =   3480
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Formw1128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command5_Click()
If MsgBox("确定结转", vbYesNo) = vbNo Then Exit Sub
If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
sql1 = "insert into clskpz1(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) VALUES('" & Adodc2.Recordset.Fields(0) & "','" & Adodc2.Recordset.Fields(1) & "','" & Adodc2.Recordset.Fields(2) & "','" & Adodc2.Recordset.Fields(5) & "','','" & Adodc2.Recordset.Fields(6) & "','" & Adodc2.Recordset.Fields(7) & "','" & Adodc2.Recordset.Fields(8) & "','" & Adodc2.Recordset.Fields(9) & "','" & Adodc2.Recordset.Fields(10) & "','" & Adodc2.Recordset.Fields(11) & "','" & Adodc2.Recordset.Fields(12) & "','" & Adodc2.Recordset.Fields(13) & "','" & Adodc2.Recordset.Fields(14) & "')"
sql2 = "insert into clskpz1(摘要,总账科目,明细科目,借方金额,贷方金额,凭证号,日期,原始单据,记账,复核,制单,原始单据数,审核确认,记账标记) VALUES('" & Adodc2.Recordset.Fields(0) & "','" & Adodc2.Recordset.Fields(3) & "','" & Adodc2.Recordset.Fields(4) & "','','" & Adodc2.Recordset.Fields(5) & "','" & Adodc2.Recordset.Fields(6) & "','" & Adodc2.Recordset.Fields(7) & "','" & Adodc2.Recordset.Fields(8) & "','" & Adodc2.Recordset.Fields(9) & "','" & Adodc2.Recordset.Fields(10) & "','" & Adodc2.Recordset.Fields(11) & "','" & Adodc2.Recordset.Fields(12) & "','" & Adodc2.Recordset.Fields(13) & "','" & Adodc2.Recordset.Fields(14) & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc2.Recordset.MoveNext
Loop
MsgBox ("ok")
End Sub

Private Sub Command6_Click()
If MsgBox("确定结转", vbYesNo) = vbNo Then Exit Sub
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
If InStr(Adodc3.Recordset.Fields(3), "-") > 0 Then
ZL = Mid(Adodc3.Recordset.Fields(3), 1, InStr(Adodc3.Recordset.Fields(3), "-") - 1)
mX = Mid(Adodc3.Recordset.Fields(3), InStr(Adodc3.Recordset.Fields(3), "-") + 1)
sql1 = "insert into pzdz1(日期,凭证号,摘要,总账科目,明细科目,借方金额,贷方金额,制单,复核,登账,凭证类别,特种日账,总分类账,明细类账) VALUES('" & Adodc3.Recordset.Fields(0) & "','" & Adodc3.Recordset.Fields(1) & "','" & Adodc3.Recordset.Fields(2) & "','" & ZL & "','" & mX & "','" & Adodc3.Recordset.Fields(4) & "','" & Adodc3.Recordset.Fields(5) & "','" & Adodc3.Recordset.Fields(6) & "','" & Adodc3.Recordset.Fields(7) & "','" & Adodc3.Recordset.Fields(8) & "','" & Adodc3.Recordset.Fields(9) & "','" & Adodc3.Recordset.Fields(10) & "','" & Adodc3.Recordset.Fields(11) & "','" & Adodc3.Recordset.Fields(12) & "')"
Else
ZL = Adodc3.Recordset.Fields(3)
mX = ""
sql1 = "insert into pzdz1(日期,凭证号,摘要,总账科目,明细科目,借方金额,贷方金额,制单,复核,登账,凭证类别,特种日账,总分类账,明细类账) VALUES('" & Adodc3.Recordset.Fields(0) & "','" & Adodc3.Recordset.Fields(1) & "','" & Adodc3.Recordset.Fields(2) & "','" & ZL & "','" & mX & "','" & Adodc3.Recordset.Fields(4) & "','" & Adodc3.Recordset.Fields(5) & "','" & Adodc3.Recordset.Fields(6) & "','" & Adodc3.Recordset.Fields(7) & "','" & Adodc3.Recordset.Fields(8) & "','" & Adodc3.Recordset.Fields(9) & "','" & Adodc3.Recordset.Fields(10) & "','" & Adodc3.Recordset.Fields(11) & "','" & Adodc3.Recordset.Fields(12) & "')"
End If
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc3.Recordset.MoveNext
Loop
MsgBox ("ok")
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
Text1.Text = Adodc1.Recordset.Fields(0)
Text2.Text = Adodc1.Recordset.Fields(1)
End Sub
Private Sub JILU()
Dim i As Single
Adodc1.Refresh
Adodc1.Recordset.MoveLast
VSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To Adodc1.Recordset.RecordCount
VSFlexGrid1.TextMatrix(i, 0) = i
Next
End Sub




Private Sub Command3_Click()
On Error Resume Next

Adodc1.Recordset.Fields(0) = Text1.Text
Adodc1.Recordset.Fields(1) = Text2.Text
Adodc1.Recordset.Update
Adodc1.Refresh
Text1.Text = ""
Text2.Text = Adodc1.Recordset.RecordCount + 1
Text1.SetFocus

End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.Delete
Adodc1.Refresh
Text1.Text = ""
Text2.Text = Adodc1.Recordset.RecordCount + 1
Text1.SetFocus

End Sub

Private Sub Form_Load()


Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "FHY"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "clskpz"
Adodc2.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "pzdz"
Adodc3.Refresh

Text1.Text = ""
Text2.Text = Adodc1.Recordset.RecordCount + 1
VSFlexGrid1.ColWidth(1) = 1500
Text1.TabIndex = 0
End Sub
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = Text1.Text
Adodc1.Recordset.Fields(1) = Text2.Text
Adodc1.Recordset.Update
Adodc1.Refresh
Text1.Text = ""
Text2.Text = Adodc1.Recordset.RecordCount + 1
Text1.SetFocus
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub




