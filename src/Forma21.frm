VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma21 
   BackColor       =   &H00C0E0FF&
   Caption         =   "织布进度"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "生产进度"
      Height          =   975
      Left            =   8280
      TabIndex        =   13
      Top             =   1680
      Width           =   6255
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "有条件"
         Height          =   255
         Left            =   120
         MaskColor       =   &H0000C0C0&
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "无条件"
         Height          =   255
         Left            =   120
         MaskColor       =   &H0000C0C0&
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "出库完"
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "出库中"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "入库完"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "入库中"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "织布完"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "织布中"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "已排产"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "未排产"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "标织中"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "标织完"
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command4 
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
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1095
      Left            =   10560
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单据"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "合同号"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "代码"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "机台"
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "品名"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "织号"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1320
      Top             =   9960
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
   Begin MSDataListLib.DataCombo DataCombo6 
      Height          =   330
      Left            =   1440
      TabIndex        =   27
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo10 
      Height          =   330
      Left            =   4800
      TabIndex        =   28
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo8 
      Height          =   330
      Left            =   4800
      TabIndex        =   29
      Top             =   1680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Height          =   330
      Left            =   1440
      TabIndex        =   30
      Top             =   2640
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   4800
      TabIndex        =   31
      Top             =   2160
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   4800
      TabIndex        =   32
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma21.frx":0000
      Height          =   330
      Left            =   1440
      TabIndex        =   33
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   10200
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   34
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   329252865
      CurrentDate     =   39961
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1440
      TabIndex        =   35
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   329252865
      CurrentDate     =   39961
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma21.frx":0015
      Height          =   6135
      Left            =   360
      TabIndex        =   36
      Top             =   3120
      Width           =   17775
      _cx             =   31353
      _cy             =   10821
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      GridLines       =   2
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
      WallPaperAlignment=   4
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Left            =   4800
      TabIndex        =   37
      Top             =   2640
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "机台"
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
      Index           =   2
      Left            =   4320
      TabIndex        =   47
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "合同号"
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
      Left            =   4320
      TabIndex        =   46
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "客户"
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
      TabIndex        =   45
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "单据"
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
      Left            =   360
      TabIndex        =   44
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "代码"
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
      TabIndex        =   43
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "品名"
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
      Index           =   8
      Left            =   4320
      TabIndex        =   42
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "业务"
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
      Index           =   3
      Left            =   4320
      TabIndex        =   41
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "起始日期"
      Height          =   375
      Index           =   18
      Left            =   360
      TabIndex        =   40
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "结束日期"
      Height          =   375
      Index           =   19
      Left            =   360
      TabIndex        =   39
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "织号"
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
      Index           =   4
      Left            =   4320
      TabIndex        =   38
      Top             =   2640
      Width           =   495
   End
End
Attribute VB_Name = "Forma21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Check1_Click(Index As Integer)
Select Case Index
       Case 0
If Check1(0).value = 1 Then
cxtjsz(0) = "(车台='' or 车台 is null)"
Else
cxtjsz(0) = ""
End If

       Case 1
If Check1(1).value = 1 Then
cxtjsz(1) = "车台<>''"
Else
cxtjsz(1) = ""
End If

       Case 2
If Check1(2).value = 1 Then
cxtjsz(2) = "(欠织>10 or 欠织 is null)"
Else
cxtjsz(2) = ""
End If

       Case 3
If Check1(3).value = 1 Then
cxtjsz(3) = "欠织< 10"
Else
cxtjsz(3) = ""
End If

       Case 4
If Check1(4).value = 1 Then
cxtjsz(4) = "(入库<累计 or 入库 is null)"
Else
cxtjsz(4) = ""
End If

       Case 5
If Check1(5).value = 1 Then
cxtjsz(5) = "入库>=累计"
Else
cxtjsz(5) = ""
End If

       Case 6
If Check1(6).value = 1 Then
cxtjsz(6) = "(出库<入库 or 出库 is null)"
Else
cxtjsz(6) = ""
End If

       Case 7
If Check1(7).value = 1 Then
cxtjsz(7) = "出库>=入库"
Else
cxtjsz(7) = ""
End If

       Case 8
If Check1(8).value = 1 Then
cxtjsz(8) = "织标 is null"
Else
cxtjsz(8) = ""
End If

       Case 9
If Check1(9).value = 1 Then
cxtjsz(9) = "织标 is not null"
Else
cxtjsz(9) = ""
End If

End Select
End Sub



Private Sub Command1_Click()
'On Error Resume Next
If Option1.value = True Then

If Check2(0).value = 1 Then
sql1 = sql1 + "日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "客户 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "单号 like '%'+'" & DataCombo6.Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "合同号 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "left(单号,1)='" & DataCombo7.Text & "' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "织号 like '%'+'" & DataCombo4.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "车台 like '%'+'" & DataCombo3.Text & "'+'%' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "品名 like '%'+'" & DataCombo8.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)


sql2 = ""
For i = 0 To 9
If Check1(i).value = 1 Then
sql2 = sql2 + Trim(cxtjsz(i)) + " AND "
End If
Next

If sql2 = "" Then
MsgBox ("请选择生产条件")
Exit Sub
End If

sql2 = Left$(Trim(sql2), Len(Trim(sql2)) - 4)

Adodc1.RecordSource = "SELECT 单据,织号,客户,品名,颜色,筒颈,计划,排产,累计,欠织,日期,交期,纱别,车台 as 车间 FROM zbkpd where (" + sql1 + ") and (" + sql2 + ") ORDER BY 单号,织号"
Adodc1.Refresh
End If

If Option3.value = True Then

If Check2(0).value = 1 Then
sql1 = sql1 + "日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "客户 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "单号 like '%'+'" & DataCombo6.Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "合同号 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "left(单号,1)='" & DataCombo7.Text & "' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "织号 like '%'+'" & DataCombo4.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "车台 like '%'+'" & DataCombo3.Text & "'+'%' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "品名 like '%'+'" & DataCombo8.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "SELECT 单据,织号,客户,品名,颜色,筒颈,计划,排产,累计,欠织,日期,交期,纱别,车台 as 车间 FROM zbkpd where (" + sql1 + ") ORDER BY 单号,织号"
Adodc1.Refresh
End If
End Sub

Private Sub Command2_Click()
sql1 = "update kpdcljd set 累计=累计产量"
sql2 = "update kpdcljd set 疵布=累计疵布"
sql3 = "update v_kpd_jtjh1_sx set 排产=计划量"
sql4 = "update kpdcljd set 欠织=round(排产-累计,2)"
sql5 = "update clkpdsh set 出纱=出库 where (出库<>出纱)"
sql6 = "update v_kpd_mpsx set 入库=入库量,出库=出库量 where 入库<>入库量 or 出库<>出库量"
sql7 = "update kpd set 实耗=round((出纱-累计)/累计*100,0) where (出库<>出纱)"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic
RD.Open sql7, conn, adOpenStatic, adLockOptimistic

MsgBox ("刷新成功！")
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Call jdmx(VSFlexGrid1, "订单查询明细")
End Sub

Private Sub Form_Load()
DTPicker1.value = Date - 5
DTPicker2.value = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
DataCombo6.Text = ""
DataCombo7.Text = ""
DataCombo8.Text = ""
DataCombo10.Text = ""
Text1 = ""
Option3.value = True
Check2(0).value = 1
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 单据,织号,客户,品名,颜色,筒颈,计划,排产,累计,欠织,日期,交期,纱别,车台 as 车间 FROM zbkpd where 日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime)  ORDER BY 单号,织号"
Adodc1.Refresh


VSFlexGrid1.ColWidth(0) = 200
For i = 1 To 7
VSFlexGrid1.ColWidth(i) = 1200
Next
VSFlexGrid1.ColWidth(6) = 1000
VSFlexGrid1.ColWidth(7) = 1000
VSFlexGrid1.ColWidth(8) = 1000
VSFlexGrid1.ColWidth(9) = 1000
VSFlexGrid1.ColWidth(10) = 1000
VSFlexGrid1.ColWidth(11) = 1500
VSFlexGrid1.ColWidth(12) = 1500
VSFlexGrid1.ColWidth(13) = 1500
VSFlexGrid1.ColWidth(14) = 1200
End Sub

Private Sub Option1_Click()
On Error Resume Next
For i = 0 To 7
Check1(i).value = 0
cxtjsz(i) = ""
Next
End Sub

Private Sub Text1_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from v_khZL where 简码 like '%'+'" & Text1.Text & "'+'%' group by 简称"
Adodc2.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
cl = VSFlexGrid1.col
Adodc1.Recordset.Move rs - 1
DataCombo4.Text = Adodc1.Recordset.Fields(1)
DataCombo6.Text = Adodc1.Recordset.Fields(0)
If cl = 2 Then
Formj17.DataCombo4 = Adodc1.Recordset.Fields(1)
Formj17.Show
End If
End Sub

