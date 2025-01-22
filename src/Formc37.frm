VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formc37 
   BackColor       =   &H00C0E0FF&
   Caption         =   "报价转入"
   ClientHeight    =   10215
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   12780
   StartUpPosition =   3  '窗口缺省
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   3720
      TabIndex        =   15
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   1560
      TabIndex        =   14
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   840
      Top             =   9960
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
      Left            =   1320
      Top             =   9960
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   1680
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc37.frx":0000
      Height          =   330
      Left            =   1560
      TabIndex        =   13
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "客户查询"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "色号查询"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转入查询"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "未转查询"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "转入外"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formc37.frx":0015
      Height          =   8055
      Left            =   480
      TabIndex        =   11
      Top             =   1680
      Width           =   11655
      _cx             =   20558
      _cy             =   14208
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
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择客户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Formc37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public c, r As Integer
Private Sub Command1_Click()
Call MXOutadodcToExcel(VSFlexGrid2, "客户：" + DataCombo1.Text + "色号" + DataCombo2.Text + "-" + DataCombo5.Text)
End Sub

Private Sub Command2_Click()
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
If DataCombo1.Text = "" Or DataCombo2.Text = "" Or DataCombo3.Text = "" Then
MsgBox ("请输入完整！")
Exit Sub
End If
sql1 = "delete  from bjb where 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If DataCombo1.Text = "" Then
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where 色号 LIKE '%'+'" & DataCombo2.Text & "'+'%' order by 色号 desc"
Adodc2.Refresh
Else
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where 客户='" & DataCombo1.Text & "' and 色号 LIKE '%'+'" & DataCombo2.Text & "'+'%' order by 色号 desc"
Adodc2.Refresh
End If
End Sub

Private Sub Command5_Click()
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where 转入='已' order by 色号 desc"
Adodc2.Refresh
End Sub

Private Sub Command6_Click()
If DataCombo1.Text = "" Then
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb order by 色号 desc"
Adodc2.Refresh
Else
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where 客户='" & DataCombo1.Text & "' order by 色号 desc"
Adodc2.Refresh
End If
End Sub

Private Sub Command7_Click()
Adodc2.RecordSource = "select 客户,品名,色号,颜色,染费,报价,转入,备注 from bjb where 转入='未' or 转入 IS null order by 色号 desc"
Adodc2.Refresh
End Sub

Private Sub Command8_Click()
If MsgBox("确定转入吗？", vbYesNo) = vbNo Then Exit Sub
If DataCombo1.Text = "" Or DataCombo2.Text = "" Or DataCombo3.Text = "" Then
MsgBox ("请输入完整！")
Exit Sub
End If
sql1 = "delete  from dj  WHERE 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
sql2 = "INSERT INTO DJ(客户,品名,色号,色别,单价,备注) SELECT 客户,品名,色号,颜色,报价,备注 FROM bjb WHERE 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"
sql3 = "update bjb set 转入='已' WHERE 客户='" & DataCombo1.Text & "' and 色号='" & DataCombo2.Text & "' and  品名='" & DataCombo3.Text & "'"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic

MsgBox ("转入成功!在色别单价中查询")
End Sub


Private Sub Form_Load()
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 简称 from KHZL  group by 简称"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from bjb order by 色号 desc"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid2.ColWidth(0) = 300
VSFlexGrid2.ColWidth(1) = 1600
VSFlexGrid2.ColWidth(2) = 1600
VSFlexGrid2.ColWidth(3) = 1600
VSFlexGrid2.ColWidth(4) = 1600
End Sub

Private Sub VSFlexGrid2_DblClick()
If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc2.Recordset.Move rs - 1
DataCombo1.Text = Adodc2.Recordset.Fields(0)
DataCombo2.Text = Adodc2.Recordset.Fields(2)
DataCombo3.Text = Adodc2.Recordset.Fields(1)
End Sub


Private Sub MSFlex()
With VSFlexGrid2
    c = .col: r = .Row    '''''C列，，R行
    If c = 6 Or c = 8 Then
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
    End If
End With
End Sub


Private Sub VSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid2.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move r - 1

Adodc2.Recordset.Fields(c - 1) = Text1111.Text
Adodc2.Recordset.Update
VSFlexGrid2.Text = Text1111.Text
Text1111.Visible = False
VSFlexGrid2.SetFocus
End If
End Sub


