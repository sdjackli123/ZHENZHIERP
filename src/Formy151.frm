VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formy151 
   BackColor       =   &H00C0E0FF&
   Caption         =   "分材料盘存操作"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   15810
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "清空操作库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "本月结存结转"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "理论库存转实库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1111 
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "库类刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "本月结存转报表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查看报表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "损耗核算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "材料刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   9240
      Top             =   10920
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
      Caption         =   "Adodc6"
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
      Height          =   495
      Left            =   9360
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
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
      Height          =   375
      Left            =   8640
      Top             =   10800
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Left            =   9240
      Top             =   10800
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Height          =   495
      Left            =   9360
      Top             =   10800
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
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
      Height          =   495
      Left            =   9600
      Top             =   10800
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "Formy151.frx":0000
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "YS"
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formy151.frx":0015
      Height          =   330
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "材料名称"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formy151.frx":002A
      Height          =   330
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   10440
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   475136001
      CurrentDate     =   39921
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formy151.frx":003F
      Height          =   8415
      Left            =   600
      TabIndex        =   15
      Top             =   1920
      Width           =   14535
      _cx             =   25638
      _cy             =   14843
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
      Caption         =   "请选择库类"
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
      Left            =   600
      TabIndex        =   19
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入材料"
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
      Left            =   600
      TabIndex        =   18
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "请输入颜色"
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
      Left            =   600
      TabIndex        =   17
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "转至日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   10440
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Formy151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public c, r As Integer
Private Sub Command1_Click()
Call OutadodcToExcel3(VSFlexGrid1, 11, 13, 15, "盘存打印")
End Sub



Private Sub Command11_Click()
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "select * from FCLPC WHERE 库类 in(select mc from clkl where yh='" & yhm & "') order by 库类,材料名称"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select * from FCLPC where 库类='" & DataCombo1.Text & "' order by 材料名称"
Adodc1.Refresh
End If
End Sub


Private Sub Command2_Click()
sql1 = "UPDATE FCLPC SET 实际金额=round(单价*实际库存,2),损耗数量=round(理论库存-实际库存,2) WHERE 库类 in(select mc from clkl where yh='" & yhm & "')"
sql2 = "UPDATE FCLPC SET 损耗金额=理论金额-实际金额 WHERE 库类 in(select mc from clkl where yh='" & yhm & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
MsgBox ("损耗核算成功！")
End Sub

Private Sub Command3_Click()
If MsgBox("确定清楚吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM FCLPC "
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
On Error Resume Next

Adodc1.RecordSource = "SELECT * FROM FCLPC"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
If MsgBox("无转库记录，是否空转", vbYesNo) = vbNo Then Exit Sub
Else
If MsgBox("请确认：转入库存记录的日期为：" + Str(DTPicker1.Value), vbYesNo) = vbYes Then
If MsgBox("确定结转吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete  from cljl1 where 日期=CAST('" & DTPicker1.Value & "' AS DATETIME) and 库类 in(select mc from clkl where yh='" & yhm & "')"
sql2 = "INSERT INTO  cljl1 (供应单位,材料名称,材料规格,材料单位,颜色,批次,单价,数量,金额,库类,日期)  SELECT 供应单位,材料名称,材料规格,材料单位,颜色,批次,单价,实际库存,实际金额,库类,'" & DTPicker1.Value & "' FROM FCLPC where 实际库存<>0 and 库类 in(select mc from clkl where yh='" & yhm & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("转库成功！")
Else
MsgBox ("转库未成功！")
End If
End If
End Sub

Private Sub Command5_Click()
sql1 = "UPDATE FCLPC SET 实际库存=理论库存,实际金额=理论金额 where  库类 in(select mc from clkl where yh='" & yhm & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc1.Refresh
MsgBox ("转库成功！")
End Sub


Private Sub Command6_Click()
If DataCombo1.Text <> "" Then
Adodc1.RecordSource = "select * from FCLPC where 材料名称 like '%'+'" & DataCombo2.Text & "'+'%' and 库类 in(select mc from clkl where yh='" & yhm & "') order by 库类,材料名称"
Adodc1.Refresh
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
'On Error Resume Next
Adodc1.RecordSource = "SELECT * FROM FCLPC where 库类 in(select mc from clkl where yh='" & yhm & "') ORDER BY 材料名称,颜色"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox ("无转库记录，终止")
Exit Sub
Else
If MsgBox("请确认：转入报表日期为：" + Str(DTPicker1.Value), vbYesNo) = vbYes Then
If MsgBox("确定转库吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete  from fclbbzcw where 日期=cast('" & DTPicker1.Value & "' as datetime) and 库类 in(select mc from clkl where yh='" & yhm & "')"
sql2 = "INSERT INTO fclbbzcw(供应单位,库类,材料名称,材料规格,材料单位,颜色,批次,单价,上月结存数量,上月结存金额,本月入库数量,本月入库金额,本月出库数量,本月出库金额,本月结存数量,本月结存金额,损耗数量,损耗金额,日期)  SELECT 供应单位,库类,材料名称,材料规格,材料单位,颜色,批次,单价,上月结存数量,上月结存金额,本月入库数量,本月入库金额,本月出库数量,本月出库金额,实际库存,实际金额,损耗数量,损耗金额,'" & DTPicker1.Value & "' FROM FCLPC where 库类 in(select mc from clkl where yh='" & yhm & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("转库成功！")
Else
MsgBox ("转库未成功！")
End If
End If
End Sub

Private Sub Command9_Click()
Adodc1.RecordSource = "select * from fclbbzcw where 日期=CAST('" & DTPicker1.Value & "' AS datetime) and 库类 in(select mc from clkl where yh='" & yhm & "') order by 库类,材料名称"
Adodc1.Refresh
End Sub



Private Sub Form_Load()
Me.Caption = Me.Caption + "操作年度 " + ljb
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DTPicker1.Value = Date

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=sc01\sql2008"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=sc01\sql2008"
Adodc1.RecordSource = "select * from FCLPC where  库类 in(select mc from clkl where yh='" & yhm & "')"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=sc01\sql2008"
Adodc2.RecordSource = "SELECT MC FROM CLKL where yh='" & yhm & "' GROUP BY MC"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=sc01\sql2008"
Adodc3.RecordSource = "SELECT 材料名称 FROM FCLPC GROUP BY 材料名称"
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=sc01\sql2008"
Adodc4.RecordSource = "SELECT YS.YS FROM YS GROUP BY YS.YS"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=sc01\sql2008"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=sc01\sql2008"
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1600
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 800
VSFlexGrid1.ColWidth(4) = 800
VSFlexGrid1.ColWidth(5) = 800
VSFlexGrid1.ColWidth(6) = 800
VSFlexGrid1.ColWidth(7) = 800
VSFlexGrid1.ColWidth(9) = 1200
VSFlexGrid1.ColWidth(10) = 1200
VSFlexGrid1.ColWidth(11) = 1200
VSFlexGrid1.ColWidth(12) = 1200
VSFlexGrid1.ColWidth(13) = 1200
VSFlexGrid1.ColWidth(14) = 1200
VSFlexGrid1.ColWidth(15) = 0
VSFlexGrid1.ColWidth(16) = 800
VSFlexGrid1.ColWidth(17) = 800
VSFlexGrid1.ColWidth(18) = 800
VSFlexGrid1.ColWidth(19) = 800


End Sub

Private Sub VSFlexGrid1_Click()
FD = VSFlexGrid1.col
End Sub

Private Sub VSFlexGrid1_DblClick()
With VSFlexGrid1
    c = .col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call VSFlexGrid1_DblClick
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid1.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move r - 1
Adodc1.Recordset.Fields(c - 1) = Text1111.Text
Adodc1.Recordset.Update
Text1111.Visible = False
VSFlexGrid1.Text = Text1111.Text
VSFlexGrid1.SetFocus
End If
End Sub


