VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma25 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1095
      Left            =   14520
      TabIndex        =   14
      Top             =   360
      Width           =   3975
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "车间"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "操作"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "织号"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "开幅"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "款号"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单号"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "验布"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
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
      Height          =   375
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
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
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
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
      Height          =   375
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command2 
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9000
      Width           =   975
   End
   Begin VB.CommandButton Command5 
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
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF80&
      Caption         =   "明细报表"
      Height          =   375
      Left            =   12840
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF80&
      Caption         =   "日报表"
      Height          =   375
      Left            =   12840
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFF80&
      Caption         =   "质检报表"
      Height          =   375
      Left            =   12840
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8760
      Top             =   10440
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
      Left            =   6360
      Top             =   10440
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
   Begin MSDataListLib.DataCombo DataCombo6 
      Height          =   330
      Left            =   6840
      TabIndex        =   23
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo10 
      Height          =   330
      Left            =   9600
      TabIndex        =   24
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo8 
      Height          =   330
      Left            =   6840
      TabIndex        =   25
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Height          =   330
      Left            =   6840
      TabIndex        =   26
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   9600
      TabIndex        =   27
      Top             =   1560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   9600
      TabIndex        =   28
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   10440
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
      Left            =   1560
      TabIndex        =   29
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   327876611
      CurrentDate     =   39961.3333333333
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1560
      TabIndex        =   30
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   327876611
      CurrentDate     =   39961.3333333333
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma25.frx":0000
      Height          =   5295
      Left            =   480
      TabIndex        =   31
      Top             =   2280
      Width           =   11895
      _cx             =   20981
      _cy             =   9340
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma25.frx":0015
      Height          =   1815
      Left            =   480
      TabIndex        =   32
      Top             =   7560
      Width           =   14295
      _cx             =   25215
      _cy             =   3201
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2760
      Top             =   10440
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forma25.frx":002A
      Height          =   1215
      Left            =   12840
      TabIndex        =   33
      Top             =   2280
      Width           =   5655
      _cx             =   9975
      _cy             =   2143
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Left            =   1560
      TabIndex        =   34
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "结束日期"
      Height          =   375
      Index           =   19
      Left            =   480
      TabIndex        =   47
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "起始日期"
      Height          =   375
      Index           =   18
      Left            =   480
      TabIndex        =   46
      Top             =   480
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
      Index           =   3
      Left            =   9120
      TabIndex        =   45
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "车间"
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
      Left            =   6240
      TabIndex        =   44
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "开幅"
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
      Left            =   6240
      TabIndex        =   43
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "单号"
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
      Left            =   6240
      TabIndex        =   42
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "款号"
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
      Left            =   9120
      TabIndex        =   41
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "操作"
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
      Left            =   9120
      TabIndex        =   40
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   39
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   38
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   37
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   36
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "质检"
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
      Left            =   480
      TabIndex        =   35
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "Forma25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
sql1 = ""

t1 = Format(Trim(DTPicker1.value) + Space(2) + Text1(0) + ":" + Text1(1) + ":" + Text1(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text3(0) + ":" + Text3(1) + ":" + Text3(2), "yyyy-MM-dd hh:mm:ss")

If Option1.value = True Then
If Check2(0).value = 1 Then
sql1 = sql1 + "日期 between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "质检 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "单号 like '%'+'" & DataCombo6.Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "款号 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "开幅 like '%'+'" & DataCombo7.Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "织号 like '%'+'" & DataCombo10.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "操作员 like '%'+'" & DataCombo3.Text & "'+'%' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "机台 like '%'+'" & DataCombo8.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "SELECT 单号,织号,款号,品名,筒颈,幅宽,克重,班次,操作员,产量,匹号,条码,支数 as 单价,产量工资,疵布,日期,质检,开幅,机台 as 车间,编号 as 机台,等级 FROM v_clbb_zjbb where (" + sql1 + ")  ORDER BY 日期,机台,单号,织号,cast(匹号 as int)"
Adodc1.Refresh

If Check2(6).value = 1 Then


Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "hzclbb('" & t1 & "','" & t2 & "','" & yhm & "','车工')"          ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel


Adodc2.RecordSource = "SELECT 挡车,布类,白班,白价,白资,夜班,夜价,夜资,一等,一等率,二等,等外,等资,产量 from clbbhz where 用户='" & yhm & "' and 挡车 like '%'+'" & DataCombo3 & "'+'%' order by 挡车"
Adodc2.Refresh

If Adodc2.Recordset.EOF Then
hs = 0
Else
hs = Adodc2.Recordset.RecordCount + 1
End If



'If hs > 0 Then
'    With VSFlexGrid2
'        .Editable = flexEDKbdMouse
'        .AutoSize 0
'        .Cell(flexcpChecked, 1, 3, hs - 1, 3) = 2
'        .Cell(MergeCells, 1, 2, hs - 1, 2) = True
'        End With
'VSFlexGrid2.SubtotalPosition = flexSTBelow
'VSFlexGrid2.Subtotal flexSTSum, 1, 4, , vbGreen
'VSFlexGrid2.Subtotal flexSTSum, 1, 5, , vbGreen
'End If

End If

If Check2(1).value = 1 Then

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cjzjsc('" & t1 & "','" & t2 & "')"          ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel


Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 质检,日期,round(sum(产量),2) as 合计产量 FROM zjbbf where (" + sql1 + ") group by 质检,日期 order by 质检,日期"
Adodc2.Refresh

If Adodc2.Recordset.EOF Then
hs = 0
Else
hs = Adodc2.Recordset.RecordCount + 1
End If



If hs > 0 Then
    With VSFlexGrid2
        .Editable = flexEDKbdMouse
'        .AutoSize 0
'        .Cell(flexcpChecked, 1, 3, hs - 1, 3) = 2
'        .Cell(MergeCells, 1, 2, hs - 1, 2) = True
        End With
VSFlexGrid2.SubtotalPosition = flexSTBelow
VSFlexGrid2.Subtotal flexSTSum, 1, 3, , vbGreen
End If
End If

If Check2(4).value = 1 Then

'Set g_Cmd = New Command
'    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
'    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
'    g_Cmd.CommandText = "cjzjsc('" & DTPicker1.Value & "','" & DTPicker2.Value & "')"          ' 表示调用哪个存储过程
'    g_Cmd.Execute           ' 执行存储过程
'g_Cmd.Cancel


Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 开幅,日期,round(sum(产量),2) as 合计产量 FROM zjbbf where (" + sql1 + ") group by 开幅,日期 order by 开幅,日期"
Adodc2.Refresh

If Adodc2.Recordset.EOF Then
hs = 0
Else
hs = Adodc2.Recordset.RecordCount + 1
End If



If hs > 0 Then
    With VSFlexGrid2
        .Editable = flexEDKbdMouse
'        .AutoSize 0
'        .Cell(flexcpChecked, 1, 3, hs - 1, 3) = 2
'        .Cell(MergeCells, 1, 2, hs - 1, 2) = True
        End With
VSFlexGrid2.SubtotalPosition = flexSTBelow
VSFlexGrid2.Subtotal flexSTSum, 1, 3, , vbGreen
End If
End If

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT round(sum(产量),2) as 合计产量 FROM v_clbb_zjbb where (" + sql1 + ")"
Adodc3.Refresh

End If

If Option2.value = True Then
If Check2(0).value = 1 Then
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text1(0) + ":" + Text1(1) + ":" + Text1(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text3(0) + ":" + Text3(1) + ":" + Text3(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "日期 between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If


If Check2(5).value = 1 Then
sql1 = sql1 + "织号 like '%'+'" & DataCombo10.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "SELECT * from v_clbb_rbb where (" + sql1 + ") ORDER BY 日期,织号"
Adodc1.Refresh


Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT round(sum(日产量),2) as 合计产量 FROM v_clbb_rbb where (" + sql1 + ")"
Adodc3.Refresh

End If

If Option3.value = True Then
If Check2(0).value = 1 Then
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text1(0) + ":" + Text1(1) + ":" + Text1(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text3(0) + ":" + Text3(1) + ":" + Text3(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If


If Check2(5).value = 1 Then
sql1 = sql1 + "织号 like '%'+'" & DataCombo10.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "SELECT 品名,sum(合计产量) as 总产量,等级 from v_clbb_rbb1 where (" + sql1 + ") group by 品名,等级 ORDER BY 品名,等级"
Adodc1.Refresh


Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT round(sum(合计产量),2) as 合计产量 FROM v_clbb_rbb1 where (" + sql1 + ")"
Adodc3.Refresh

End If

End Sub

Private Sub Command2_Click()
Call jdmx(VSFlexGrid2, "汇总产量")
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Call jdmx(VSFlexGrid1, "产量明细")
End Sub

Private Sub Command5_Click()
Call jdmx(VSFlexGrid3, "汇总产量")
End Sub

Private Sub Form_Load()
On Error Resume Next
For i = 0 To 2
Text1(i) = "00"
Text3(i).Text = "00"
Next
Text1(2) = "00"
Text3(0).Text = "23"
Text3(1).Text = "00"
Text3(2).Text = "00"

DTPicker1.value = Date - 1
DTPicker2.value = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo6.Text = ""
DataCombo7.Text = ""
DataCombo8.Text = ""
DataCombo10.Text = ""
Adodc1.CommandTimeout = 10000
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 单号,织号,款号 as 合同号,品名,筒颈,幅宽,克重,操作员,产量,匹号,条码,支数 as 单价,产量工资,疵布,日期,质检,开幅,机台 as 车间,编号 as 机台,等级 FROM v_clbb_zjbb where 日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime)  ORDER BY 单号,织号"
Adodc1.Refresh
Adodc2.CommandTimeout = 10000
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.CommandTimeout = 10000
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid2.ColWidth(0) = 200
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
DataCombo6.Text = Adodc1.Recordset.Fields(0)
End Sub

