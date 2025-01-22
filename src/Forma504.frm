VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forma504 
   Caption         =   "数据汇总表"
   ClientHeight    =   13215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   27345
   LinkTopic       =   "Form1"
   Picture         =   "Forma504.frx":0000
   ScaleHeight     =   13215
   ScaleWidth      =   27345
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Forma504.frx":0342
      Left            =   5640
      List            =   "Forma504.frx":034F
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   15
      Text            =   "Text8"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   14
      Text            =   "Text8"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "结  转"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   21480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   735
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   735
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   19800
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   307625985
      CurrentDate     =   45061
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   307625985
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   307625985
      CurrentDate     =   36892
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma504.frx":0365
      Height          =   11535
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   26895
      _cx             =   47440
      _cy             =   20346
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   375
         Left            =   9960
         Top             =   7680
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Left            =   7200
         Top             =   7680
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
         Left            =   4680
         Top             =   7560
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   23
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   22
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   21
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   4
      Left            =   3600
      TabIndex        =   20
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "结转至"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19920
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "请选择日期范围"
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
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Forma504"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public r, c, FD As Integer: Public k1, k2 As String
Private Sub Command6_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "染色产量报表")
End Sub

Private Sub Command1_Click()
Dim rq1 As Date, rq2 As Date, RQ3 As Date
rq1 = DTPicker3.value
rq2 = DTPicker4.value
RQ3 = DTPicker1.value

Adodc2.RecordSource = "SELECT * FROM prbbjzr   ORDER BY 客户"
Adodc2.Refresh
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 1500
For i = 2 To 14
VSFlexGrid1.ColWidth(i) = 1500
Next
VSFlexGrid1.SubtotalPosition = flexSTBelow

VSFlexGrid1.Subtotal flexSTSum, 0, 5, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 6, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 7, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 8, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 9, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 10, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 11, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 12, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 13, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 14, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 15, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 16, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 17, , &HC0C0&
End Sub

Private Sub Combo1_Click()
    Debug.Print "Combo1 Clicked: " & Combo1.Text

    If Combo1.Text = "全天" Then
        Text2(3).Text = "00"
        Text2(1).Text = "00"
        Text2(2).Text = "00"
        Text8(0).Text = "23"
        Text8(1).Text = "59"
        Text8(2).Text = "59"
    ElseIf Combo1.Text = "白班" Then
        Text2(3).Text = "07"
        Text2(1).Text = "30"
        Text2(2).Text = "00"
        Text8(0).Text = "19"
        Text8(1).Text = "29"
        Text8(2).Text = "59"
    ElseIf Combo1.Text = "夜班" Then
        ' 将 DTPicker3 的值加一天，并设置给 DTPicker4
        DTPicker4.value = DTPicker3.value + 1
        
        ' 将 Text1.Text 转换为日期，加一天，然后转换回字符串
        Dim newDate As Date
        newDate = CDate(Text1.Text) + 1
        Text2(0).Text = Format(newDate, "yyyy-M-d")

        Text2(3).Text = "19"
        Text2(1).Text = "30"
        Text2(2).Text = "00"
        Text8(0).Text = "07"
        Text8(1).Text = "29"
        Text8(2).Text = "59"
    End If
End Sub






Private Sub Command2_Click()
Call fhbb1(VSFlexGrid1, 2, 3, 4, 5, 6, 7, 8, 9, "汇总报表" + Text1.Text + "-" + Text2(0).Text)
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim t1 As String
        Dim t2 As String
        t1 = Format(Trim(DTPicker3.value) + Space(2) + Text2(3) + ":" + Text2(1) + ":" + Text2(2), "yyyy-MM-dd hh:mm:ss")
        t2 = Format(Trim(DTPicker4.value) + Space(2) + Text8(0) + ":" + Text8(1) + ":" + Text8(2), "yyyy-MM-dd hh:mm:ss")
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "prbbCZr('" & t1 & "','" & t2 & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel
Adodc1.RecordSource = "SELECT 日期, 毛坯入库重量, 染色重量,染色金额合计,烘干重量,定型重量,蒸汽用量,水用量,染化料助剂金额,成品入库重量, 成品库存重量,成品库存金额, 成品发货重量, 应收款, 回款 FROM prbbjzr   ORDER BY 日期"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 2500
VSFlexGrid1.ColWidth(2) = 2000
VSFlexGrid1.ColWidth(3) = 2000

For i = 4 To 9
VSFlexGrid1.ColWidth(i) = 2000
Next
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 0, 2, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 3, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 4, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 5, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 6, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 7, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 8, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 9, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 10, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 11, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 12, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 13, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 14, , &HC0C0&
VSFlexGrid1.Subtotal flexSTSum, 0, 15, , &HC0C0&
End Sub

Private Sub Command5_Click()
On Error Resume Next
Adodc1.RecordSource = "SELECT 日期, 毛坯入库重量, 染色重量,染色金额合计,烘干重量,定型重量,蒸汽用量,水用量,染化料助剂金额,成品入库重量, 成品库存重量,成品库存金额, 成品发货重量, 应收款, 回款  FROM prbbjzr   ORDER BY 日期"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox ("无转库记录，终止")
Exit Sub
Else
If MsgBox("转至日期为" + Trim(DTPicker1.value) + "确定转至吗？", vbYesNo) = vbNo Then Exit Sub
If MsgBox("确定转库吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM prbbjlr WHERE 日期=cast('" & DTPicker1.value & "' as datetime)"
sql2 = "INSERT INTO  prbbjlr (款号, 品名, 色别,累计毛坯重量,累计毛坯匹数, 累计染色重量, 累计染色匹数,累计入库重量,累计入库匹数,日期)  SELECT 款号, 品名, 色别, 累计毛坯重量,累计毛坯匹数, 累计染色重量, 累计染色匹数,累计入库重量,累计入库匹数,'" & DTPicker1.value & "' FROM prbbjzr "
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
MsgBox ("转库成功！")
End If
End Sub

Private Sub DTPicker3_Change()
Text1.Text = DTPicker3.value
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = DTPicker3.value
Text1.SetFocus
End Sub

Private Sub DTPicker4_Change()
Text2(0).Text = DTPicker4.value
End Sub

Private Sub DTPicker4_CloseUp()
Text2(0).Text = DTPicker4.value
Text2(0).SetFocus
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
 Combo1 = "全天"
Text1.Text = Date
Text2(0).Text = Date
DTPicker1.value = Date
DTPicker3.value = Date
DTPicker4.value = Date
Text2(3) = "00"
Text2(1) = "00"
Text2(2) = "00"
Text8(0) = "23"
Text8(1) = "59"
Text8(2) = "59"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc2.RecordSource = "SELECT 款号, 品名, 色别, 前存毛坯重量,前存毛坯匹数,前存染色重量, 前存染色匹数, 前存入库重量, 前存入库匹数, 当日毛坯重量, 当日毛坯匹数, 累计毛坯重量, 累计毛坯匹数, 当日染色重量, 当日染色匹数,累计染色重量,累计染色匹数,当日入库重量,当日入库匹数,累计入库重量,累计入库匹数,已染未入库重量,已染未入库匹数,未染重量,未染匹数 FROM prbbjzr ORDER BY 款号,品名"
'Adodc2.Refresh
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 日期, 毛坯入库重量, 染色重量,染色金额合计,烘干重量,定型重量,蒸汽用量,水用量,染化料助剂金额,成品入库重量, 成品库存重量,成品库存金额, 成品发货重量, 应收款, 回款 FROM prbbjzr   ORDER BY 日期"
Adodc1.Refresh
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 2000
VSFlexGrid1.ColWidth(2) = 2000
VSFlexGrid1.ColWidth(3) = 2000

For i = 4 To 9
VSFlexGrid1.ColWidth(i) = 2000
Next
End Sub




