VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Formr49 
   BackColor       =   &H00C0E0FF&
   Caption         =   "染料称量操作"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form49"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Index           =   4
      Left            =   13560
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "关闭串口"
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   16080
      Top             =   480
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "称重去皮"
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   15000
      Top             =   480
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   3
      Left            =   13560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   2
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   0
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "配料信息"
      Height          =   975
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   3255
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000FF00&
         Caption         =   "已称量"
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000000FF&
         Caption         =   "未称量"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   5280
      Top             =   10080
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
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
      Caption         =   "Adodc7"
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
      Left            =   6000
      Top             =   9960
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Height          =   330
      Left            =   5640
      Top             =   9960
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
      Height          =   330
      Left            =   4800
      Top             =   9840
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
      Height          =   375
      Left            =   5400
      Top             =   10080
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Height          =   375
      Left            =   5760
      Top             =   9720
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
      Left            =   5880
      Top             =   10080
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
      Bindings        =   "Formr49.frx":0000
      Height          =   6495
      Left            =   480
      TabIndex        =   12
      Top             =   2640
      Width           =   6255
      _cx             =   11033
      _cy             =   11456
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
      FormatString    =   $"Formr49.frx":0015
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
      Height          =   390
      Left            =   5640
      TabIndex        =   13
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330760193
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   330760193
      CurrentDate     =   36892
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formr49.frx":00EC
      Height          =   3855
      Left            =   7680
      TabIndex        =   16
      Top             =   5280
      Width           =   8775
      _cx             =   15478
      _cy             =   6800
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
   Begin MSCommLib.MSComm MSComm1 
      Left            =   15480
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      BaudRate        =   600
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   15480
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      BaudRate        =   600
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   15480
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      BaudRate        =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   27
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "提示信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   7200
      TabIndex        =   26
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "实际称重"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   12360
      TabIndex        =   25
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "需要称重"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   12360
      TabIndex        =   24
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "染料序号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   7320
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "称量染料名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   7320
      TabIndex        =   22
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "料单编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "配单信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   20
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "称量信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   1
      Left            =   7320
      TabIndex        =   19
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   18
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Formr49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim a As String
Dim flag1 As Integer
Dim flag2 As Boolean
Dim flag3 As Boolean     ''''''''染料判断变量
Dim i
Dim ksjs As Integer


Private Sub Command2_Click()
If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        End If
If MSComm2.PortOpen = True Then
            MSComm2.PortOpen = False
        End If
If MSComm3.PortOpen = True Then
            MSComm3.PortOpen = False
        End If
Timer1.Enabled = True
flag2 = False
Unload Me
End Sub


Private Sub Command3_Click()
If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
MSComm1.Output = Chr$(27) + "t"
'MSComm1.PortOpen = False
End Sub

Private Sub Command4_Click()
If Option1.value = True Then
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库<>'助剂库' and (称量标记='N' or 称量标记 is null) AND 配料日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库<>'助剂库' and 称量标记='Y' AND 配料日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
End If
End Sub



Private Sub Command5_Click()
If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        End If
If MSComm2.PortOpen = True Then
            MSComm2.PortOpen = False
        End If
If MSComm3.PortOpen = True Then
            MSComm3.PortOpen = False
        End If
Timer1.Enabled = False
Timer2.Enabled = False
End Sub



Private Sub DataCombo1_Change()
'On Error Resume Next
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & DataCombo1.Text & "' and 染化助库<>'助剂库' ORDER BY 工序名称,次序号"
Adodc2.Refresh
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call VQJC
End Sub

Private Sub Form_Load()
DTPicker1.value = Date - 10
DTPicker2.value = Date

Label4.Caption = ""

MSComm1.CommPort = 1
MSComm1.Settings = "600,n,8,1"

MSComm2.CommPort = 2
MSComm2.Settings = "600,n,8,1"

MSComm3.CommPort = 3
MSComm3.Settings = "600,n,8,1"

DataCombo1.Text = ""
Option1.value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

flag1 = 4 ''''''''不显示称重量

flag2 = True
flag3 = False
For m = 0 To 4
Text1(m) = ""
Next


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
If Option1.value = True Then
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库<>'助剂库' and (称量标记='N' or 称量标记 is null) AND 配料日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库<>'助剂库' and 称量标记='Y' AND 配料日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
End If




VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(3) = 2500


End Sub




Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 3
If Val(Text1(3)) > 0 And Val(Text1(3)) = Val(Text1(2)) And Val(Text1(3)) > 0 Then
Timer2.Enabled = True
ksjs = 0
End If
       Case 4
If Text1(4) = "" Then Exit Sub
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select * from rhzh where ip='" & Text1(4) & "' and 染料名称='" & Text1(0) & "'"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
Timer1.Enabled = False
flag3 = False
Else
Beep 2000, 50
Timer1.Enabled = True
flag3 = True
Text1(4) = ""
End If
End Select
End Sub



Private Sub Timer1_Timer()

If flag1 = 0 Then

If flag3 = True Then
If MSComm2.PortOpen = True Then
            MSComm2.PortOpen = False
End If

If MSComm3.PortOpen = True Then
            MSComm3.PortOpen = False
End If

Label4.Caption = "请用1号称    称量"
If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
MSComm1.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until flag2 And MSComm1.InBufferCount >= 13
a = MSComm1.Input
Text1(3) = Trim(Mid(a, 1, 9))            ''''''称重量

If Val(Text1(3)) < -1 Then    ''''''''''''''''''''''''如果小于0  去皮
If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
MSComm1.Output = Chr$(27) + "t"
End If

End If
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If flag1 = 1 Then
If flag3 = True Then
If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
End If

If MSComm3.PortOpen = True Then
            MSComm3.PortOpen = False
End If

Label4.Caption = "请用2号称    称量"

If MSComm2.PortOpen = False Then
            MSComm2.PortOpen = True
        End If
MSComm2.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until flag2 And MSComm2.InBufferCount >= 13
a = MSComm2.Input
Text1(3) = Trim(Mid(a, 1, 9)) * 1000

If Val(Text1(3)) < -1 Then    ''''''''''''''''''''''''如果小于0  去皮
If MSComm2.PortOpen = False Then
            MSComm2.PortOpen = True
        End If
MSComm2.Output = Chr$(27) + "t"
End If

End If
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If flag1 = 2 Then
If flag3 = True Then
If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
End If

If MSComm2.PortOpen = True Then
            MSComm2.PortOpen = False
End If

Label4.Caption = "请用3号称    称量"

If MSComm3.PortOpen = False Then
            MSComm3.PortOpen = True
        End If
MSComm3.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until flag2 And MSComm3.InBufferCount >= 13
a = MSComm3.Input
Text1(3) = Trim(Mid(a, 1, 9)) * 1000

If Val(Text1(3)) < -1 Then    ''''''''''''''''''''''''如果小于0  去皮
If MSComm3.PortOpen = False Then
            MSComm3.PortOpen = True
        End If
MSComm3.Output = Chr$(27) + "t"
End If


End If
End If

End Sub

Private Sub Timer2_Timer()
If Val(Text1(2)) = Val(Text1(3)) And Val(Text1(2)) > 0 Then
ksjs = ksjs + 1
Beep 1000, 50
If ksjs / 2 = Int(ksjs / 2) Then
Text1(0).ForeColor = &HFF&
Else
Text1(0).ForeColor = &HFF00&
End If
Else
ksjs = 0
Text1(0).ForeColor = &HFF&
End If
If ksjs = 6 Then
Timer1.Enabled = False
Timer2.Enabled = False
flag1 = 4
sql1 = "UPDATE pldr SET 实际称量='" & Text1(3) & "',称量员='" & yhm & "',称量标记='Y',称量日期='" & Now & "' WHERE 料单编号='" & DataCombo1.Text & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & DataCombo1.Text & "' and 染化助库<>'助剂库' ORDER BY 工序名称,次序号"
Adodc2.Refresh

If MSComm1.PortOpen = True Then
MSComm1.Output = Chr$(27) + "t"
End If

If MSComm2.PortOpen = True Then
MSComm2.Output = Chr$(27) + "t"
End If

If MSComm3.PortOpen = True Then
MSComm3.Output = Chr$(27) + "t"
End If

Call VQJC
Call Command4_Click
Text1(0).ForeColor = &HFF&
End If
End Sub

Private Sub VQJC()

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT ISNULL(称量标记,'N'),工序名称,染化助库,染化助名称,配料单位,round(配料用量,2),实际称量,次序号 FROM pldr WHERE (称量标记<>'Y' OR 称量标记 IS NULL) AND 料单编号='" & DataCombo1.Text & "' and 染化助库<>'助剂库' ORDER BY 工序名称,次序号"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
Text1(0) = ""
Text1(1) = ""
Text1(2) = ""
Text1(3) = ""
Text1(4) = ""
Label4.Caption = "称重完成"
Else
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''根据称量转换串口
If Adodc3.Recordset.Fields(0) <> "Y" Then
Text1(0) = Adodc3.Recordset.Fields(3)
Text1(1) = Adodc3.Recordset.Fields(7)
Text1(2) = Adodc3.Recordset.Fields(5)

If Val(Text1(2)) <= 1000 And Val(Text1(2)) >= 0 Then
Text1(2) = Format(Text1(2), "#0.00")
flag1 = 0
flag3 = False
Timer1.Enabled = False
Label4.Caption = "请扫描染料条码"
Text1(3) = 0
Text1(4) = ""
Text1(4).SetFocus
End If
If Val(Text1(2)) > 1000 Then
Text1(2) = Format(Text1(2), "#0")
flag1 = 1
flag3 = False
Timer1.Enabled = False
Label4.Caption = "请扫描染料条码"
Text1(3) = 0
Text1(4) = ""
Text1(4).SetFocus
End If
'If Val(Text1(2)) > 6000 And Val(Text1(2)) <= 30000 Then
'flag1 = 2
'End If

Exit Sub
End If


Adodc3.Recordset.MoveNext
Loop
End If
End Sub

Private Sub VSFlexGrid1_dblClick()
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
DataCombo1.Text = Adodc1.Recordset.Fields(2)
End Sub


