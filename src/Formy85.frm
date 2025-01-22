VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formy85 
   BackColor       =   &H00C0E0FF&
   Caption         =   "设置单价"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   11490
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   3720
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   6000
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5040
      Top             =   5280
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
      Left            =   5040
      Top             =   5280
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
      Left            =   5040
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "刷新"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      ItemData        =   "Formy85.frx":0000
      Left            =   720
      List            =   "Formy85.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formy85.frx":0004
      Height          =   3255
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   7215
      _cx             =   12726
      _cy             =   5741
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
   Begin VB.Label Label3 
      Caption         =   "序号"
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "缸号"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "取消"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "全选"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Formy85"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim c, r As Integer: Dim xm As String
Private Sub Command1_Click()
List4.Clear
Adodc2.RecordSource = "select 加工项目 from jgxm order by 序号"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
List4.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
List4.Selected(0) = True
Adodc1.RecordSource = "select 项目,单价 from dhjgxm where 缸号='" & Text1(0) & "'"
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
sql1 = "delete from dhjgxm where 缸号='" & Text1(0) & "' "
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
For i = 0 To List4.ListCount - 1
If List4.Selected(i) = True Then
sql2 = "insert into dhjgxm(缸号,序号,项目,单价) VALUES('" & Text1(0) & "','" & Text1(1) & "','" & List4.List(i) & "',0)"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
Next
If hssx = 1 Then
Formj12.Adodc3.Refresh
hssx = 0
Unload Me
End If
If hssx = 2 Then
Formc146.Adodc1.Refresh
End If
End Sub

Private Sub Command3_Click()
If hssx = 1 Then
Formj12.Adodc3.Refresh
hssx = 0
Unload Me
End If
If hssx = 2 Then
Formc146.Adodc1.Refresh
hssx = 0
Unload Me
End If
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
End Sub

Private Sub Form_Resize()
Call Command1_Click
End Sub

Private Sub Label1_Click()
For i = 0 To List4.ListCount - 1
List4.Selected(i) = True
Next
End Sub

Private Sub Label2_Click()
For i = 0 To List4.ListCount - 1
List4.Selected(i) = False
Next
End Sub

Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
    If c = 2 Then
        xm = .TextMatrix(r, 1)
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
    End If
End With
End Sub


Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 1
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 项目,单价 from dhjgxm where 缸号='" & Text1(0) & "'"
Adodc1.Refresh
End Select
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If hssx = 2 Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
If MsgBox("删除" + "加工项目：" + Trim(Adodc1.Recordset.Fields(0)) + "吗？", vbYesNo) = vbNo Then Exit Sub
sql2 = "delete from dhjgxm where 缸号='" & Text1(0) & "' and 项目='" & Adodc1.Recordset.Fields(0) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    VSFlexGrid1.Text = Combo1111.Text
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub

Private Sub Combo1111_LostFocus()
On Error Resume Next
dj = Val(Combo1111)
sql2 = "update dhjgxm set 单价='" & dj & "' where 缸号='" & Text1(0) & "' and 项目='" & xm & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Combo1111.Visible = False
VSFlexGrid1.SetFocus
End Sub
