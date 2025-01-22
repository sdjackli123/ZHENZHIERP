VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formd334 
   BackColor       =   &H00C0E0FF&
   Caption         =   "染色核算确认"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   11355
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Text            =   "Text4"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "校正"
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   4080
      Width           =   975
   End
   Begin VB.ListBox List3 
      Height          =   3210
      ItemData        =   "Formd334.frx":0000
      Left            =   3960
      List            =   "Formd334.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3840
      ItemData        =   "Formd334.frx":0004
      Left            =   600
      List            =   "Formd334.frx":0006
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "工序确定"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "工序追加"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formd334.frx":0008
      Height          =   2415
      Left            =   600
      TabIndex        =   10
      Top             =   1320
      Width           =   6615
      _cx             =   11668
      _cy             =   4260
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1440
      Top             =   10560
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
      Left            =   1560
      Top             =   10560
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
      Left            =   1800
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8280
      Top             =   10680
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   8400
      Top             =   10680
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8640
      Top             =   10680
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formd334.frx":001D
      Height          =   7455
      Left            =   7680
      TabIndex        =   11
      Top             =   1200
      Width           =   3015
      _cx             =   5318
      _cy             =   13150
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
      FormatString    =   $"Formd334.frx":0032
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "确认"
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
      Left            =   10560
      TabIndex        =   18
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "染色工序"
      Height          =   3375
      Left            =   3720
      TabIndex        =   17
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号序号信息"
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   16
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号信息"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   15
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入锅号"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "染色工序信息"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "工序信息"
      Height          =   375
      Index           =   6
      Left            =   7680
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "Formd334"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc1.Recordset.MoveFirst
List1.Clear
Do While Not Adodc1.Recordset.EOF
List1.AddItem Trim(Adodc1.Recordset.Fields(1))
Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim L2 As String
    If MsgBox("工序已选择，确认此类设置吗？", vbYesNo) = vbNo Then Exit Sub

    If Text1 = "" Then Exit Sub

    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then

            ll = ""
            l1 = ""
            For Q = 0 To List3.ListCount - 1
                If List3.Selected(Q) = True Then
                    l1 = Mid(List3.List(Q), 1, InStr(List3.List(Q), "-") - 1)
                    L2 = Mid(List3.List(Q), InStr(List3.List(Q), "-") + 1) ' 取出 '-' 之后的数据
                    bs = Val(Text3)

                    ' 检查 l2 是否存在于 Adodc3
                    Adodc3.RecordSource = "select 工序名称 from ghgx where 锅号='" & Text1.Text & "' and 工序 between '1001' and '6000'"
                    Adodc3.Refresh

                    Dim exists As Boolean
                    exists = False
                    Adodc3.Recordset.MoveFirst
                    Do While Not Adodc3.Recordset.EOF
                        If Adodc3.Recordset.Fields("工序名称").value = L2 Then
                            exists = True
                            Exit Do
                        End If
                        Adodc3.Recordset.MoveNext
                    Loop

                    If exists Then
                        MsgBox "禁止重复加工序: " & L2
                    Else
                        ' 删除操作移动到外层循环外，确保所有选中的项都能处理
                        'sql2 = "delete from ghgx where 锅号='" & Text1.Text & "' and 序号='" & List1.List(i) & "' and 工序 BETWEEN '1001' AND '6000'"
                        'RD.Open sql2, conn, adOpenStatic, adLockOptimistic

                        ' 执行存储过程
                        Set g_Cmd = New Command
                        g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
                        g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
                        g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
                        g_Cmd.CommandText = "ghgxlr('" & Text1.Text & "','" & List1.List(i) & "','" & l1 & "','" & bs & "','" & L2 & "')"      ' 表示调用哪个存储过程
                        g_Cmd.Execute           ' 执行存储过程
                        g_Cmd.Cancel

                        ' 检查是否需要更新 ll
                        Adodc2.RecordSource = "select * from kpd where CHARINDEX('" & l1 & "',gx)>0"
                        Adodc2.Refresh
                        If Adodc2.Recordset.EOF Then
                            ll = ll + l1 + "-"
                        End If
                    End If
                End If
            Next

            ' 更新 kpd 的 gx 字段
            'sql1 = "update kpd set gx=gx+'" & ll & "' where 锅号='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
            'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
        End If
    Next

    ' 插入操作日志
    sql2 = "insert into czrz(日期,锅号,操作,内容,功能) VALUES('" & Now & "','" & Text1.Text & "','" & yhm & "','" & ll & "','染色工序')"
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic

    ' 刷新相关数据控件
    Adodc1.Refresh
    Adodc6.Refresh

    With VSFlexGrid2
        .Editable = flexEDKbdMouse
        .Cell(flexcpChecked, 1, 2, .Rows - 1, 2) = 2
    End With

    MsgBox ("设置成功！")
End Sub




Private Sub Command4_Click()
On Error Resume Next
Dim sx As Integer
If MsgBox("工序已选择，确认此类设置吗？", vbYesNo) = vbNo Then Exit Sub

If Text1 = "" Then Exit Sub
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then

ll = ""
l1 = ""
For Q = 0 To List3.ListCount - 1

If List3.Selected(Q) = True Then
l1 = Mid(List3.List(Q), 1, InStr(List3.List(Q), "-") - 1)
L2 = Mid(List3.List(Q), InStr(List3.List(Q), "-") + 1) ' 取出 '-' 之后的数据
bs = Val(Text3)
Adodc3.RecordSource = "select isnull(顺序,0) from ghgx where 锅号='" & Text1 & "' and 工序 between '1001' and '6000' order by 顺序 desc"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
sx = 1
Else
sx = Adodc3.Recordset.Fields(0) + 1
End If
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "ghgxlr('" & Text1.Text & "','" & List1.List(i) & "','" & l1 & "','" & bs & "','" & L2 & "')"     ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
    
Adodc2.RecordSource = "select * from kpd where CHARINDEX('" & l1 & "',gx)>0"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
ll = ll + l1 + "-"
End If

End If

Next

'sql1 = "update kpd set gx=gx+'" & ll & "' where 锅号='" & Text1.text & "' and ip='" & List1.List(i) & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic

End If
Next
Adodc1.Refresh
Adodc6.Refresh

sql2 = "insert into czrz(日期,锅号,操作,内容,功能) VALUES('" & Now & "','" & Text1.Text & "','" & yhm & "','" & ll & "','染色追加')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

MsgBox ("设置成功！")

End Sub

Private Sub Command5_Click()
On Error Resume Next
If Val(Text4) > 2 Then
MsgBox ("校正倍数太大 禁止")
Exit Sub
End If
For i = 1 To VSFlexGrid2.Rows - 1
If VSFlexGrid2.Cell(flexcpChecked, i, 2) = 1 Then
bs = Val(Text4)
sql1 = "UPDATE ghgx SET 倍数='" & bs & "' WHERE 锅号='" & Text1 & "' and 工序='" & VSFlexGrid2.TextMatrix(i, 2) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
Adodc6.Refresh
    With VSFlexGrid2
        .Editable = flexEDKbdMouse
        .Cell(flexcpChecked, 1, 2, .Rows - 1, 2) = 2
    End With
End Sub

Private Sub Command8_Click()
On Error Resume Next
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Command9_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub Form_Load()
On Error Resume Next

Label2.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text13.Text = ""
Text3 = 1
Text4 = 1
DataCombo1 = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

If InStr(yhmk, "管理") > 0 Then
Command4.Enabled = True
Else
Command4.Enabled = False
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 锅号,ip as 序号,品名,色别,匹数,重量,备注,gx as 工序 from kpd where 锅号='" & Text1.Text & "' order by ip"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(8) = 3200

End Sub

Private Sub Label3_Click()
On Error Resume Next
ll = ""
For i = 0 To List3.ListCount - 1
If List3.Selected(i) = True Then
ll = ll + Mid(List3.List(i), 1, InStr(List3.List(i), "-") - 1) + "-"
End If
Next
Text13.Text = Text13.Text + "-" + Mid(ll, 1, Len(ll) - 1)
For i = 0 To List3.ListCount - 1
List3.Selected(i) = False
Next
End Sub

Private Sub Label4_Click()
GXBL = 31
  ''''''根据染料的配方用量自动检索染色工序
'FormS4.Text3 = pfyljt '''' 从生成配料把车台传过来
If pfyl = 0 Then
FormS4.Text2 = "水"
End If
If pfyl <= 0.4 And pfyl > 0 Then
FormS4.Text1 = "浅"
FormS4.Text2 = "白"
'FormS4.Text3 = "漂白"
End If
If pfyl > 0.4 And pfyl <= 1.5 Then
FormS4.Text2 = "中"
End If
If pfyl > 1.5 Then
FormS4.Text2 = "深"
End If
FormS4.Show
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 4 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 锅号,ip as 序号,品名,色别,匹数,重量,备注,gx as 工序 from kpd where 锅号='" & Text1.Text & "' order by ip"
Adodc1.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 序号,工序,倍数 from ghgx where 锅号='" & Text1.Text & "' and 工序 between '1001' and '6000' order by 序号,工序"
Adodc6.Refresh

    With VSFlexGrid2
        .Editable = flexEDKbdMouse
        .Cell(flexcpChecked, 1, 2, .Rows - 1, 2) = 2
    End With

End If
Call Command2_Click
Call Command8_Click
End Sub

Private Sub Text13_Change()
List4.Clear
i = 1
For L = 0 To Int(Len(Text13.Text) / 5)
List4.AddItem Mid(Text13.Text, L * 4 + i, 4)
i = i + 1
Next
For i = 0 To List4.ListCount - 1
List4.Selected(i) = True
Next
End Sub

Private Sub Text2_Change()
Formd331.Text9 = Text2
If Text2.Text = "" Then Exit Sub
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 工艺编号,工序名称 from GYSHD where 工序其它系数='" & Text2.Text & "' and 工艺编号 between '1001' and  '6000' GROUP BY 工艺编号,工序名称 order by 工艺编号"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
List3.Clear
Exit Sub
End If
Adodc4.Recordset.MoveFirst
List3.Clear
Do While Not Adodc4.Recordset.EOF
List3.AddItem Adodc4.Recordset.Fields(0) + "-" + Trim(Adodc4.Recordset.Fields(1))
Adodc4.Recordset.MoveNext
Loop
For i = 0 To List3.ListCount - 1
List3.Selected(i) = True
Next
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc6.Recordset.EOF Then Exit Sub
rs = VSFlexGrid2.Row
Adodc6.Recordset.MoveFirst
Adodc6.Recordset.Move rs - 1
If MsgBox("删除" + "序号：" + Trim(Adodc6.Recordset.Fields(0)) + "工序：" + Adodc6.Recordset.Fields(1) + "吗？", vbYesNo) = vbNo Then Exit Sub
sql2 = "delete from ghgx  where 锅号='" & Text1.Text & "' and 序号='" & Adodc6.Recordset.Fields(0) & "' and 工序='" & Adodc6.Recordset.Fields(1) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc6.Refresh
End Sub

