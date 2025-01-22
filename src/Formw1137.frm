VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formw1137 
   BackColor       =   &H00C0E0FF&
   Caption         =   "账薄浏览"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form37"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "账本导入"
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
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7440
      Top             =   10440
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
      Left            =   7560
      Top             =   10440
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
      Bindings        =   "Formw1137.frx":0000
      Height          =   7695
      Left            =   360
      TabIndex        =   15
      Top             =   1800
      Width           =   13695
      _cx             =   24156
      _cy             =   13573
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
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      ItemData        =   "Formw1137.frx":0015
      Left            =   4560
      List            =   "Formw1137.frx":0025
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按凭证"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按日期"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "按日期、凭证"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   423034881
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   423034881
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   423034881
      CurrentDate     =   39883
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "账薄查询"
      Height          =   1215
      Left            =   6840
      TabIndex        =   13
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "结账"
      Height          =   1215
      Left            =   10320
      TabIndex        =   14
      Top             =   360
      Width           =   3735
      Begin VB.CommandButton Command7 
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
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "期末结转"
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
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   12
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "凭证类别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作月份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Formw1137"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn1 As ADODB.Connection: Dim RD1 As ADODB.Recordset
Dim conn2 As ADODB.Connection: Dim RD2 As ADODB.Recordset
Public k1, k2 As String

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("输入日期")
Exit Sub
End If
If Combo1.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM PZDZ WHERE 日期 BETWEEN '" & DTPicker1 & "' AND '" & DTPicker2 & "' ORDER BY 日期,凭证号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM PZDZ WHERE 凭证类别='" & Combo1.Text & "' AND 日期 BETWEEN '" & DTPicker1 & "' AND '" & DTPicker2 & "' ORDER BY 日期,凭证号"
Adodc1.Refresh
End If
End Sub



Private Sub Command2_Click()
If MsgBox("确定转入吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("日期范围为：" + Trim(DTPicker1) + "---" + Trim(DTPicker2) + "吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("本次操作是" + ljb, vbYesNo) = vbNo Then Exit Sub
sql1 = "delete  from PZDZ WHERE 日期 BETWEEN '" & DTPicker1 & "' AND '" & DTPicker2 & "' and like 凭证号 like 'W%'"
RD1.Open sql1, conn1, adOpenStatic, adLockOptimistic

sql1 = "insert into zzpr.dbo.PZDZ  select * from zzpr.dbo.PZDZ WHERE 日期 BETWEEN '" & DTPicker1 & "' AND '" & DTPicker2 & "'"
RD2.Open sql1, conn2, adOpenStatic, adLockOptimistic

MsgBox ("转入成功！")

End Sub

Private Sub Command3_Click()
If MsgBox("确定期末结转吗？", vbYesNo) = vbNo Then Exit Sub
st3 = CDate(DTPicker2.value) + 1
If Text3.Text = "" Then
MsgBox ("月份不正确")
Exit Sub
End If

If MsgBox("操作月份为" + Text3.Text + "正确吗？", vbYesNo) = vbNo Then Exit Sub

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "zbjzhz('" & DTPicker1.value & "','" & DTPicker2.value & "','" & st3 & "','" & Text3.Text & "')"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel

MsgBox ("结转成功！")
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "SELECT * FROM PZDZ WHERE 日期 BETWEEN '" & DTPicker1 & "' AND '" & DTPicker2 & "' ORDER BY 日期,凭证号"
Adodc1.Refresh
End Sub


Private Sub Command6_Click()
If Combo1.Text = "" Then
Adodc1.RecordSource = "SELECT * FROM PZDZ WHERE 日期 BETWEEN '" & DTPicker1 & "' AND '" & DTPicker2 & "' ORDER BY 凭证号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * FROM PZDZ WHERE 凭证类别='" & Combo1.Text & "' AND 日期 BETWEEN '" & DTPicker1 & "' AND '" & DTPicker2 & "' ORDER BY 凭证号"
Adodc1.Refresh
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub DTPicker3_Change()
Text3.Text = Month(DTPicker3.value)
End Sub

Private Sub DTPicker3_CloseUp()
Text3.Text = Month(DTPicker3.value)
End Sub

Private Sub Form_Load()

On Error Resume Next
DTPicker3 = Date
DTPicker1.value = Date
Text3.Text = Month(Date)
DTPicker2.value = Date
Combo1.Text = ""

Set conn1 = New ADODB.Connection
conn1.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD1 = New ADODB.Recordset

Set conn2 = New ADODB.Connection
conn2.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD2 = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM PZDZ WHERE 日期 BETWEEN '" & DTPicker1 & "' AND '" & DTPicker2 & "' ORDER BY 日期,凭证号"
Adodc1.Refresh

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(3) = 2000
VSFlexGrid1.ColWidth(4) = 2500
VSFlexGrid1.ColWidth(7) = 700
VSFlexGrid1.ColWidth(8) = 700
VSFlexGrid1.ColWidth(9) = 700
End Sub

Private Sub Text3_Change()
On Error Resume Next
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from rqsd where 月份='" & Text3.Text & "'"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
Exit Sub
End If
DTPicker1.value = Adodc2.Recordset.Fields(0)
DTPicker2.value = Adodc2.Recordset.Fields(1)
End Sub
