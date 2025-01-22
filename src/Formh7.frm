VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formh7 
   BackColor       =   &H00C0E0FF&
   Caption         =   "客户来样管理"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "配方"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   4320
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   1440
      Top             =   10560
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "Adodc12"
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Formh7.frx":0000
      Height          =   450
      Left            =   5880
      TabIndex        =   67
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "代码"
      Text            =   "DataCombo2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   63
      Text            =   "Text2"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1095
      Left            =   4440
      TabIndex        =   56
      Top             =   5160
      Width           =   4335
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户色号"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   64
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "业务员"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   62
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色号"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   61
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   60
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "方式"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   59
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "打样员"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      ItemData        =   "Formh7.frx":0016
      Left            =   5640
      List            =   "Formh7.frx":0018
      Style           =   1  'Simple Combo
      TabIndex        =   55
      Text            =   "Combo1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认日期"
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
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "发样确认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   15960
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   13200
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   2040
      Width           =   2175
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formh7.frx":001A
      Height          =   6615
      Left            =   840
      TabIndex        =   48
      Top             =   6600
      Width           =   18975
      _cx             =   33470
      _cy             =   11668
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
      Bindings        =   "Formh7.frx":002F
      Height          =   330
      Index           =   0
      Left            =   840
      TabIndex        =   32
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   10560
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
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "分样确认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "输入方式"
      Height          =   855
      Left            =   840
      TabIndex        =   28
      Top             =   240
      Width           =   3495
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000C0C0&
         Caption         =   "连输"
         Height          =   495
         Left            =   2040
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "单输"
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   328335363
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   5880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   328335363
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   15960
      TabIndex        =   18
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   328335363
      CurrentDate     =   36892.5
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   13200
      TabIndex        =   17
      Top             =   2040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   328335363
      CurrentDate     =   36892.5
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   328335361
      CurrentDate     =   36892
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1560
      Top             =   10560
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
      Height          =   330
      Left            =   1560
      Top             =   10560
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
      Left            =   1560
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   1560
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   1560
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   1560
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   1560
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   1560
      Top             =   10560
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
      Height          =   330
      Index           =   1
      Left            =   2880
      TabIndex        =   33
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   2
      Left            =   4920
      TabIndex        =   34
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   3
      Left            =   7200
      TabIndex        =   35
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formh7.frx":0044
      Height          =   330
      Index           =   4
      Left            =   8520
      TabIndex        =   36
      Top             =   2040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "PM"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   5
      Left            =   12960
      TabIndex        =   37
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   6
      Left            =   13080
      TabIndex        =   38
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   7
      Left            =   12960
      TabIndex        =   39
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   8
      Left            =   2880
      TabIndex        =   40
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   9
      Left            =   4800
      TabIndex        =   41
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   10
      Left            =   6360
      TabIndex        =   42
      Top             =   3360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   11
      Left            =   9000
      TabIndex        =   43
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formh7.frx":0059
      Height          =   330
      Index           =   12
      Left            =   16680
      TabIndex        =   44
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "负责"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formh7.frx":006E
      Height          =   330
      Index           =   13
      Left            =   12240
      TabIndex        =   45
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "负责人姓名"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formh7.frx":0083
      Height          =   330
      Index           =   14
      Left            =   13800
      TabIndex        =   46
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formh7.frx":0099
      Height          =   330
      Index           =   15
      Left            =   15120
      TabIndex        =   47
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   360
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   480
      Top             =   10800
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
      Height          =   330
      Index           =   16
      Left            =   10680
      TabIndex        =   68
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formh7.frx":00AF
      Height          =   330
      Index           =   17
      Left            =   18240
      TabIndex        =   72
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "负责"
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单价"
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
      Index           =   18
      Left            =   18240
      TabIndex        =   71
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "业务负责"
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
      Index           =   17
      Left            =   16680
      TabIndex        =   70
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "色样编号"
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
      Index           =   16
      Left            =   4800
      TabIndex        =   66
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户名称"
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
      Left            =   840
      TabIndex        =   65
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "小样类别"
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
      Index           =   15
      Left            =   15120
      TabIndex        =   27
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "染色方式"
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
      Index           =   14
      Left            =   13800
      TabIndex        =   26
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Left            =   840
      TabIndex        =   25
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Left            =   840
      TabIndex        =   24
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "打样负责人"
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
      Index           =   13
      Left            =   12240
      TabIndex        =   21
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
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
      Index           =   12
      Left            =   10680
      TabIndex        =   20
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
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
      Index           =   11
      Left            =   6360
      TabIndex        =   16
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   10
      Left            =   9000
      TabIndex        =   15
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "发样次数"
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
      Index           =   9
      Left            =   4800
      TabIndex        =   14
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "确认意见"
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
      Index           =   8
      Left            =   2880
      TabIndex        =   13
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "确认日期"
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
      Index           =   7
      Left            =   840
      TabIndex        =   12
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
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
      Index           =   1
      Left            =   2880
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "色号"
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
      Index           =   2
      Left            =   4920
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户色号"
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
      Index           =   3
      Left            =   7200
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "打样布类"
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
      Index           =   4
      Left            =   8520
      TabIndex        =   8
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "接样日期"
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
      Index           =   5
      Left            =   13200
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "发样日期"
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
      Index           =   6
      Left            =   15960
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "Formh7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c, r As Integer
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer
Private Sub JILU2()
Dim i As Single
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
VSFlexGrid2.TextMatrix(0, 0) = "记录号"

Exit Sub
End If
Adodc2.Recordset.MoveLast
VSFlexGrid2.TextMatrix(0, 0) = "记录号"
For i = 1 To Adodc2.Recordset.RecordCount
VSFlexGrid2.TextMatrix(i, 0) = i
Next
End Sub
Private Sub Command12_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
If DataCombo1(0).Text = "" Or Text1(0).Text = "" Then
MsgBox ("请输入信息！")
Exit Sub
End If
If Option2.value = True Then
For L = 0 To 17
Adodc9.Refresh
If Formh71.Text2(L).Text <> "" Then
DataCombo1(1).Text = Formh71.Text1(L).Text
DataCombo1(2).Text = Formh71.Text2(L).Text
DataCombo1(16).Text = Formh71.Text4(L).Text
DataCombo1(10).Text = Formh71.Text5(L).Text
DataCombo1(6).Text = ""
Adodc9.RecordSource = "SELECT * FROM KHY WHERE sh='" & DataCombo1(2).Text & "' AND KH='" & DataCombo1(0).Text & "'"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
Adodc9.Recordset.AddNew

For i = 0 To Adodc9.Recordset.Fields.count - 1
Adodc9.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc9.Recordset.Fields(5) = Text1(0).Text
Adodc9.Recordset.Fields(6) = Text1(1).Text
Adodc9.Recordset.Fields(18) = DataCombo1(17).Text
Adodc9.Recordset.Update
Adodc1.Refresh
Adodc6.RecordSource = "SELECT COUNT(SH) FROM KHY WHERE KH='" & DataCombo1(0).Text & "'"
Adodc6.Refresh
DataCombo1(11).Text = 1
If Adodc6.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc6.Recordset.Fields(0) + 1
End If
End If
End If
Next
Unload Formh71
Option1.value = True
For i = 1 To 3
DataCombo1(i).Text = ""
Next
Exit Sub
End If

If Option1.value = True Then
Adodc9.Refresh
If DataCombo1(2).Text = "" Then
MsgBox ("请输入色号！")
Exit Sub
End If
Adodc9.RecordSource = "SELECT * FROM KHY WHERE sh='" & DataCombo1(2).Text & "' AND KH='" & DataCombo1(0).Text & "'"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
DataCombo1(6).Text = ""
Adodc9.Recordset.AddNew
For i = 0 To Adodc9.Recordset.Fields.count - 1
Adodc9.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc9.Recordset.Fields(18) = DataCombo1(17).Text
Adodc9.Recordset.Update
Adodc1.Refresh
Adodc6.Refresh
For i = 1 To 3
DataCombo1(i).Text = ""
Next
DataCombo1(11).Text = 1
If Adodc6.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc6.Recordset.Fields(0) + 1
End If
End If
End If
Adodc1.Refresh
End Sub

Private Sub Command10_Click()
Formh77.DTPicker1.value = DTPicker4.value
Formh77.DTPicker2.value = DTPicker5.value
Formh77.DTPicker3.value = Date
Formh77.Text1.Text = DataCombo1(0).Text
Formh77.Show
End Sub

Private Sub Command11_Click()
Formh223.DataCombo1(1) = DataCombo1(0)
Formh223.DataCombo1(5) = DataCombo1(4)
Formh223.DataCombo1(3) = DataCombo1(1)
Formh223.DataCombo1(4) = DataCombo1(2)
Formh223.DataCombo1(8) = 10
Formh223.DataCombo1(2) = DataCombo1(13)
Formh223.DataCombo1(0) = Date
Formh223.Text2 = DataCombo1(2)
Formh223.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
If DataCombo1(0).Text = "" Or Text1(0).Text = "" Then
MsgBox ("请输入信息！")
Exit Sub
End If
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc1.Recordset.Fields(18) = DataCombo1(17).Text
Adodc1.Recordset.Fields(5) = Text1(0).Text
Adodc1.Recordset.Fields(6) = Text1(1).Text
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc6.Refresh
For i = 1 To 3
DataCombo1(i).Text = ""
Next
DataCombo1(11).Text = 1
If Adodc6.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc6.Recordset.Fields(0) + 1
End If
DataCombo4.Text = ""
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Adodc6.Refresh
For i = 1 To Adodc9.Recordset.Fields.count - 1
DataCombo1(i).Text = ""
Next
DataCombo1(11).Text = 1
If Adodc6.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc6.Recordset.Fields(0) + 1
End If
DataCombo4.Text = ""
khx = 0
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Call bhmx(VSFlexGrid1, 9, 10, DataCombo1(0).Text)
End Sub

Private Sub Command7_Click()
Formh74.DTPicker1.value = DTPicker4.value
Formh74.DTPicker2.value = DTPicker5.value
Formh74.DTPicker3.value = Date
Formh74.Text1.Text = DataCombo1(0).Text
Formh74.Show
End Sub

Private Sub Command8_Click()
Adodc6.RecordSource = "SELECT COUNT(SH) FROM KHY WHERE KH='" & DataCombo1(0).Text & "'"
Adodc6.Refresh
DataCombo1(11).Text = 1
If Adodc6.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc6.Recordset.Fields(0) + 1
End If
Adodc1.Refresh
Adodc4.Refresh
Adodc5.Refresh
Adodc6.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub



Private Sub Command5_Click()
On Error Resume Next
sql1 = ""

If Check2(0).value = 1 Then
sql1 = sql1 + "ksh like '%'+'" & DataCombo1(3).Text & "'+'%' and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "kh like '%'+'" & DataCombo1(0).Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "dyfz like '%'+'" & DataCombo1(13).Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "lb like '%'+'" & DataCombo1(15).Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
t1 = Format(Trim(DTPicker4.value), "yyyy-mm-dd")
t2 = Format(Trim(DTPicker5.value), "yyyy-mm-dd")
sql1 = sql1 + "CONVERT(varchar,jyr, 23) between '" & t1 & "' and '" & t2 & "' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "ywf like '%'+'" & DataCombo1(12).Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "sh like '%'+'" & DataCombo1(2).Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If

sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc1.RecordSource = "SELECT kh as 客户名称,ys as 颜色,sh as 色号,ksh as 客户色号,dyb as 打样布类,jyr as 接样日期,fyr as 发样日期,qyr as 确认日期,qry as 确认意见,fyc as 发样次数,sc AS 款号,bz as 备注,ip as 序号,ywf as 业务负责人,dyfz as 打样负责人,FS AS 染色方式,LB AS 小样类别,单价 from khy where (" + sql1 + ")  order  by jyr desc,sh desc "
Adodc1.Refresh

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 500
Next
End If

End Sub



Private Sub dataCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo3_Change()
DataCombo1(13).Text = DataCombo3.Text
End Sub

Private Sub DataCombo3_Click(Area As Integer)
DataCombo1(13).Text = DataCombo3.Text
End Sub

Private Sub dataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo4_Change()
DataCombo1(1).Text = DataCombo4.Text
End Sub

Private Sub DataCombo4_Click(Area As Integer)
DataCombo1(1).Text = DataCombo4.Text
End Sub

Private Sub dataCombo4_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo5_Change()
DataCombo1(4).Text = DataCombo5.Text
End Sub

Private Sub DataCombo5_Click(Area As Integer)
DataCombo1(4).Text = DataCombo5.Text
End Sub

Private Sub dataCombo5_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub Command9_Click()
Formh72.DTPicker1.value = DTPicker4.value
Formh72.DTPicker2.value = DTPicker5.value
Formh72.DTPicker3.value = Date
Formh72.Text1.Text = DataCombo1(0).Text
Formh72.Show
End Sub

Private Sub DataCombo1_Change(Index As Integer)
Select Case Index
       Case 0
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT COUNT(SH) FROM KHY WHERE KH='" & DataCombo1(0).Text & "'"
Adodc6.Refresh
DataCombo1(11).Text = 1
If Adodc6.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc6.Recordset.Fields(0) + 1
End If
'Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc1.RecordSource = "SELECT kh as 客户名称,ys as 颜色,sh as 色号,ksh as 客户色号,dyb as 打样布类,jyr as 接样日期,fyr as 发样日期,qyr as 确认日期,qry as 确认意见,fyc as 发样次数,HT AS 款号,bz as 备注,ip as 序号,ywf as 业务负责人,dyfz as 打样负责人,FS AS 染色方式,LB AS 小样类别 from khy where jyr BETWEEN '" & DTPicker4.Value & "' and '" & DTPicker5.Value & "' AND KH='" & DataCombo1(0).Text & "' order by jyr desc ,IP DESC"
'Adodc1.Refresh
End Select
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
Select Case Index
       Case 0
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT COUNT(SH) FROM KHY WHERE KH='" & DataCombo1(0).Text & "'"
Adodc6.Refresh
DataCombo1(11).Text = 1
If Adodc6.Recordset.EOF Then
DataCombo1(11).Text = 1
Else
DataCombo1(11).Text = Adodc6.Recordset.Fields(0) + 1
End If
'Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc1.RecordSource = "SELECT kh as 客户名称,ys as 颜色,sh as 色号,ksh as 客户色号,dyb as 打样布类,jyr as 接样日期,fyr as 发样日期,qyr as 确认日期,qry as 确认意见,fyc as 发样次数,HT AS 款号,bz as 备注,ip as 序号,ywf as 业务负责人,dyfz as 打样负责人,FS AS 染色方式,LB AS 小样类别 from khy where jyr BETWEEN '" & DTPicker4.Value & "' and '" & DTPicker5.Value & "' AND KH='" & DataCombo1(0).Text & "' order by jyr desc,IP DESC "
'Adodc1.Refresh
End Select
End Sub




Private Sub DTPicker1_Click()
Text1(0).Text = Format(DTPicker1.value)
End Sub

Private Sub DTPicker1_CloseUp()
Text1(0).Text = Format(DTPicker1.value)
End Sub

Private Sub DTPicker2_Click()
Text1(1).Text = Format(DTPicker2.value)
End Sub

Private Sub DTPicker2_CloseUp()
Text1(1).Text = Format(DTPicker2.value)
End Sub

Private Sub DTPicker3_Click()
Text1(2).Text = DTPicker3.value
End Sub

Private Sub DTPicker3_CloseUp()
Text1(2).Text = DTPicker3.value
End Sub


Private Sub Form_Load()
On Error Resume Next

DTPicker4.value = Date
DTPicker5.value = Date
cdbhf = cdbh
For i = 0 To 2
Text1(i).Text = ""
Next
Text2 = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT kh as 客户名称,ys as 颜色,sh as 色号,ksh as 客户色号,dyb as 打样布类,jyr as 接样日期,fyr as 发样日期,qyr as 确认日期,qry as 确认意见,fyc as 发样次数,sc AS 款号,bz as 备注,ip as 序号,ywf as 业务负责人,dyfz as 打样负责人,FS AS 染色方式,LB AS 小样类别,单价 from khy where jyr BETWEEN '" & DTPicker4.value & "' and '" & DTPicker5.value & "'  order by jyr desc "
Adodc1.Refresh

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Option1.value = True
For i = 0 To 17
DataCombo1(i).Text = ""
Next
DataCombo2.Text = Format(Text1(0), "YYMMDD") + "0"
Text1(0) = Date
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from khzl where ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 负责 from ywf group by 负责"
Adodc3.Refresh
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 负责人姓名 from gr group by 负责人姓名"
Adodc4.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select pm from pm group by pm"
Adodc5.Refresh
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "select MC from RSFS group by MC"
Adodc10.Refresh

Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.RecordSource = "select 代码 from dybh group by 代码"
Adodc12.Refresh

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "KHY"
Adodc9.Refresh
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select MC from XYLB group by MC"
Adodc11.Refresh

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(2) = 800
VSFlexGrid1.ColWidth(4) = 1200
VSFlexGrid1.ColWidth(5) = 1200
VSFlexGrid1.ColWidth(6) = 1200

DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
DataCombo1(5).Text = Date
DataCombo1(6).Text = ""
DataCombo1(7).Text = ""
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False


End Sub
Private Sub JILU()
Dim i As Single
Adodc1.Refresh
If Adodc9.Recordset.EOF Then
VSFlexGrid1.TextMatrix(0, 0) = "记录号"

Exit Sub
End If
Adodc9.Recordset.MoveLast
VSFlexGrid1.TextMatrix(0, 0) = "记录号"
For i = 1 To Adodc9.Recordset.RecordCount
VSFlexGrid1.TextMatrix(i, 0) = i
Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
sql2 = "insert into yhcd(用户,菜单,编号) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.WindowState = 2
Formm1.Adodc1.Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sql2 = "delete from yhcd where 用户='" & yhm & "' and 编号='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.Adodc1.Refresh
End Sub

Private Sub Text1_Change(Index As Integer)
DataCombo2.Text = Format(Text1(0), "YYMMDD") + "0"
End Sub

Private Sub Text2_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from KHZL where 代码 like '%'+'" & Text2 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by 简称"
Adodc2.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To Adodc1.Recordset.Fields.count - 1
DataCombo1(i).Text = Adodc1.Recordset.Fields(i)
Next
DataCombo1(17).Text = Adodc1.Recordset.Fields(18)
For i = 0 To 2
Text1(i).Text = Adodc1.Recordset.Fields(5 + i)
Next

Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Option2_Click()
On Error Resume Next
If DataCombo1(0).Text = "" Or DataCombo1(4) = "" Or Text1(0).Text = "" Or DataCombo1(15).Text = "" Then
MsgBox ("信息输入不完整!")
Option2.value = False
Option1.value = True
Exit Sub
End If
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select sh from khy where  jyr =cast( '" & Text1(0) & "' as datetime) and len(sh)=8 order by sh desc"
Adodc7.Refresh
If Adodc7.Recordset.EOF Then
Formh71.Text3 = Format(Text1(0), "YYMMDD") + "00"
Else
Formh71.Text3 = Trim(Adodc7.Recordset.Fields(0))
End If
Formh71.Show
End Sub


Private Sub MSF()
On Error Resume Next
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
If Mid(yhm, 1, 2) = "hy" Or InStr(yhm, "root") > 0 Or InStr(yhm, "jh") > 0 Then
    
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
Else

If c = 9 Or c = 3 Then
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
End If

End If
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Call MSF
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move r - 1


Adodc1.Recordset.Fields(c - 1) = Combo1111.Text
Adodc1.Recordset.Update

    VSFlexGrid1.Text = Combo1111.Text
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub



