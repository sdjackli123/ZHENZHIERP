VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formc34 
   BackColor       =   &H00C0E0FF&
   Caption         =   "光坯入库"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15900
   Icon            =   "Formc34.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15900
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "码单"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "码单"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "新单据"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4920
      TabIndex        =   55
      Text            =   "Text3"
      Top             =   1680
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   8640
      Top             =   10200
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "Adodc10"
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
      Height          =   495
      Left            =   9480
      Top             =   10200
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Adodc9"
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
      Height          =   375
      Left            =   8760
      Top             =   10320
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "Adodc8"
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
      Left            =   9120
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
      Height          =   495
      Left            =   9720
      Top             =   10080
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Left            =   9000
      Top             =   10320
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
      Left            =   10320
      Top             =   10080
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
      Left            =   9600
      Top             =   10320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   10200
      Top             =   10200
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
      Left            =   10560
      Top             =   10200
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Formc34.frx":0A7A
      Height          =   1815
      Left            =   360
      TabIndex        =   54
      Top             =   5040
      Width           =   15015
      _cx             =   26485
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
      FormatString    =   $"Formc34.frx":0A8F
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formc34.frx":0B64
      Height          =   1935
      Left            =   360
      TabIndex        =   53
      Top             =   7200
      Width           =   15015
      _cx             =   26485
      _cy             =   3413
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
      FormatString    =   $"Formc34.frx":0B79
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
      Index           =   0
      Left            =   2880
      TabIndex        =   35
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421376
      Format          =   330760193
      CurrentDate     =   39274
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   0
      Left            =   3720
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DBCombo1"
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3720
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   7920
      TabIndex        =   23
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330760193
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   330760193
      CurrentDate     =   36892
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc34.frx":0C4E
      Height          =   330
      Index           =   1
      Left            =   4440
      TabIndex        =   36
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   2
      Left            =   11640
      TabIndex        =   37
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   3
      Left            =   5040
      TabIndex        =   38
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   4
      Left            =   8640
      TabIndex        =   39
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   5
      Left            =   11640
      TabIndex        =   40
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   6
      Left            =   10200
      TabIndex        =   41
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   7
      Left            =   840
      TabIndex        =   42
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   8
      Left            =   3000
      TabIndex        =   43
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   9
      Left            =   6600
      TabIndex        =   44
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   10
      Left            =   5280
      TabIndex        =   45
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   11
      Left            =   4320
      TabIndex        =   46
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc34.frx":0C63
      Height          =   330
      Index           =   12
      Left            =   7200
      TabIndex        =   47
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "KL"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc34.frx":0C78
      Height          =   330
      Index           =   13
      Left            =   9840
      TabIndex        =   48
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   14
      Left            =   10920
      TabIndex        =   49
      Top             =   3240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   15
      Left            =   12960
      TabIndex        =   50
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   16
      Left            =   840
      TabIndex        =   51
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   17
      Left            =   5760
      TabIndex        =   52
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   18
      Left            =   2040
      TabIndex        =   57
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   19
      Left            =   11640
      TabIndex        =   59
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   20
      Left            =   12960
      TabIndex        =   61
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formc34.frx":0C8D
      Height          =   330
      Index           =   21
      Left            =   7680
      TabIndex        =   64
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "光坯单位"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   22
      Left            =   10200
      TabIndex        =   66
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   23
      Left            =   8640
      TabIndex        =   71
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "染耗"
      Height          =   375
      Index           =   8
      Left            =   8640
      TabIndex        =   70
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   1920
      TabIndex        =   69
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "成分"
      Height          =   375
      Index           =   7
      Left            =   10200
      TabIndex        =   65
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单位"
      Height          =   375
      Index           =   6
      Left            =   7680
      TabIndex        =   63
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Index           =   3
      Left            =   12960
      TabIndex        =   62
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号序号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   11640
      TabIndex        =   60
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "克重"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   58
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "锅号"
      Height          =   375
      Left            =   8640
      TabIndex        =   34
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H008080FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   33
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H008080FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   32
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "缸号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   31
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单据"
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   28
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   5
      Left            =   2880
      TabIndex        =   27
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   26
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   25
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "幅宽"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   22
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "仓库备活表:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   19
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光坯入库表;"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "序号"
      Height          =   375
      Index           =   15
      Left            =   10920
      TabIndex        =   15
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "班次"
      Height          =   375
      Index           =   14
      Left            =   9840
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "整理"
      Height          =   375
      Index           =   13
      Left            =   7200
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光坯匹数"
      Height          =   375
      Index           =   12
      Left            =   4320
      TabIndex        =   11
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光坯数量"
      Height          =   375
      Index           =   11
      Left            =   5280
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "光坯米数"
      Height          =   375
      Index           =   10
      Left            =   6600
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "毛坯重量"
      Height          =   375
      Index           =   9
      Left            =   3000
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "日期"
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户锅号"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "颜色"
      Height          =   375
      Index           =   3
      Left            =   11640
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "品名"
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   375
      Left            =   11640
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   375
      Index           =   0
      Left            =   12960
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "Formc34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Single: Public ll As Integer: Public LKP As String
Public mm As Date: Public ML As Date: Public KI As String: Public GHB As String '''''锅号变量
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf As Integer


Private Sub Command10_Click()
Forms51.DataCombo4(1) = DataCombo1(4)
Forms51.Show
End Sub

Private Sub Command3_Click()
On Error Resume Next
For i = 1 To 23
If i = 14 Then i = 15
DataCombo1(i) = ""
Next
DataCombo1(4).Text = ""
DataCombo1(14).Enabled = False

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT * FROM jgmxkfdj where 单据编号='" & yhdm & "'"
Adodc6.Refresh

DataCombo1(16).Text = Trim(yhdm) + "0000001"
If Adodc6.Recordset.EOF Then
DataCombo1(16).Text = Trim(yhdm) + "0000001"
Else
uu = Val(Adodc6.Recordset.Fields(1)) + 1
DataCombo1(16).Text = Trim(yhdm) + Left("0000000", 7 - Len(uu)) + Trim(uu)
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM jgmxkf where 单据= '" & DataCombo1(16).Text & "' order by  序号 desc"
Adodc1.Refresh

DataCombo1(14).Text = Adodc1.Recordset.RecordCount + 1
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

End Sub


Private Sub Command7_Click()
If DataCombo1(17) = "" Or DataCombo1(19) = "" Or DataCombo1(4) = "" Or DataCombo1(14) = "" Then
MsgBox ("请选择仓库备活的信息")
Exit Sub
End If
Formc144.Text1 = DataCombo1(4)
Formc144.Text2(0) = DataCombo1(16)
Formc144.Text2(1) = DataCombo1(14)
Formc144.Text3(0) = DataCombo1(17)
Formc144.Text3(1) = DataCombo1(19)
Formc144.Show
End Sub


Private Sub Command8_Click()
Formc345.Check2(4).value = 1
Formc345.Show
End Sub

Private Sub Command9_Click()
Adodc1.RecordSource = "select *  FROM jgmxkf where  单据= '" & DataCombo1(16).Text & "' order by  序号 desc"
Adodc1.Refresh
DataCombo1(14).Text = Adodc1.Recordset.RecordCount + 1
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub adodcCombo1_Click(Index As Integer, Area As Integer)

End Sub

Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 4
If InStr(DataCombo1(4).Text, "J") > 0 Or InStr(DataCombo1(4).Text, "j") > 0 Then DataCombo1(4).Text = Mid(DataCombo1(4), 1, Len(DataCombo1(4).Text) - 1)
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT  日期,锅号,客户名称,色别,色名 as 色号,品名,'' as 毛坯幅宽,isnull(配缸重量,0) as 重量,isnull(配缸匹数,0) as 匹数,克重,类别,序号,生产状态,缸号,款号,'' as 备注,光胚幅宽,光坯匹数,光坯重量 FROM v_kpd_ok where 锅号= '" & DataCombo1(4).Text & "' order by  序号 "
Adodc2.Refresh
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select *  FROM jgmxkf where  锅号= '" & DataCombo1(4).Text & "' order by  缸号 desc"
Adodc1.Refresh
       Case 8
If Val(DataCombo1(8)) <> 0 Then
DataCombo1(23) = Format((Val(DataCombo1(8)) - Val(DataCombo1(10))) / Val(DataCombo1(8)) * 100, "#0.0")
End If
       Case 10
If Val(DataCombo1(8)) <> 0 Then
DataCombo1(23) = Format((Val(DataCombo1(8)) - Val(DataCombo1(10))) / Val(DataCombo1(8)) * 100, "#0.0")
End If
       Case 16
End Select

VSFlexGrid3.ColWidth(0) = 200
End Sub



Private Sub DataCombo5_Change()
DataCombo7.Text = Val(DataCombo5.Text) * Val(DataCombo6.Text)
End Sub


Private Sub dataCombo6_Change()
DataCombo7.Text = Val(DataCombo5.Text) * Val(DataCombo6.Text)
End Sub

Private Sub Command1_Click()
On Error Resume Next

If DataCombo1(1).Text = "" Then
MsgBox ("请输入客户")
Exit Sub
End If

If DataCombo1(21) = "" Then
MsgBox ("单位不正确")
Exit Sub
End If

If DateDiff("d", DTPicker1.value, Date) > 3 Then
If MsgBox("日期可能有误 请检查  是否继续？", vbYesNo) = vbNo Then Exit Sub
End If

Adodc10.RecordSource = "select * from jgmxkf where 锅号='" & DataCombo1(4) & "' and 缸号='" & DataCombo1(17) & "'"
Adodc10.Refresh

If Not Adodc10.Recordset.EOF Then
If MsgBox("已有入库信息，确定是否继续？", vbYesNo) = vbNo Then Exit Sub
End If

If DataCombo1(8).Text = "" Then DataCombo1(8).Text = 0
If DataCombo1(9).Text = "" Then DataCombo1(9).Text = 0
If DataCombo1(10).Text = "" Then DataCombo1(10).Text = 0
If DataCombo1(11).Text = "" Then DataCombo1(11).Text = 0

Adodc1.Recordset.AddNew
For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.RecordSource = "select *  FROM jgmxkf where  锅号= '" & DataCombo1(4).Text & "' order by  缸号 desc"
Adodc1.Refresh

sql1 = "UPDATE kpd SET zt='光坯入库',KP1=convert(nvarchar ,'" & DTPicker1.value & "',23)  WHERE 锅号='" & DataCombo1(4) & "' and 缸号='" & DataCombo1(17) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic


For i = 1 To 23
If i = 16 Then i = 17
If i = 21 Then i = 22
If i <> 4 Then
DataCombo1(i).Text = ""
End If
Next
DataCombo1(14).Text = Adodc1.Recordset.RecordCount + 1
End Sub

Private Sub Command2_Click()
On Error Resume Next
If DataCombo1(1).Text = "" Then
MsgBox ("请输入客户")
Exit Sub
End If

If Len(DataCombo1(16)) <> 8 Then
MsgBox ("单据不正确")
Exit Sub
End If
If DataCombo1(21) = "" Then
MsgBox ("单位不正确")
Exit Sub
End If


If DataCombo1(8).Text = "" Then DataCombo1(8).Text = 0
If DataCombo1(9).Text = "" Then DataCombo1(9).Text = 0
If DataCombo1(10).Text = "" Then DataCombo1(10).Text = 0
If DataCombo1(11).Text = "" Then DataCombo1(11).Text = 0
DataCombo1(0).Text = DTPicker1.value
For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.RecordSource = "select *  FROM jgmxkf where  锅号= '" & DataCombo1(4).Text & "' order by  缸号 desc"
Adodc1.Refresh

sql1 = "UPDATE kpd SET zt='光坯入库',KP1=convert(nvarchar ,'" & DTPicker1.value & "',23)  WHERE 锅号='" & DataCombo1(4) & "' and 缸号='" & DataCombo1(17) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic


For i = 1 To Adodc1.Recordset.Fields.count - 1
If i = 16 Then i = 17
If i = 21 Then i = 22
If i <> 4 Then
DataCombo1(i).Text = ""
End If
Next
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
DataCombo1(14).Text = Adodc1.Recordset.RecordCount + 1
DataCombo1(4).SetFocus
End Sub


Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.Delete
Adodc1.RecordSource = "select *  FROM jgmxkf where  锅号= '" & DataCombo1(4).Text & "' order by  缸号 desc"
Adodc1.Refresh
Adodc10.RecordSource = "select 工序 from v_cjcl where 锅号='" & DataCombo1(4) & "' order by 工艺编号 desc"
Adodc10.Refresh

If Not Adodc10.Recordset.EOF Then
sql1 = "UPDATE kpd SET zt='" & Adodc10.Recordset.Fields(0) & "',KP1=''  WHERE 锅号='" & DataCombo1(4) & "' and 缸号='" & DataCombo1(17) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
For i = 1 To 23
If i = 16 Then i = 17
If i = 21 Then i = 22
If i <> 4 Then
DataCombo1(i).Text = ""
End If
Next
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
DataCombo1(14).Text = Adodc1.Recordset.RecordCount + 1
DataCombo1(4).SetFocus
End Sub

Private Sub Command6_Click()
On Error Resume Next
BA.Close
Unload Me
End Sub


Private Sub DataCombo6_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub




Private Sub DTPicker1_Change()
DataCombo1(0) = DTPicker1.value
End Sub

Private Sub DTPicker1_CloseUp()
DataCombo1(0) = DTPicker1.value
End Sub

Private Sub Form_Load()

On Error Resume Next


For i = 0 To 23
DataCombo1(i) = ""
Next
Text3.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = "1"
DataCombo1(4).Text = ""
DataCombo1(0).Text = Date
DTPicker1.value = Date
DTPicker3.value = Date
DTPicker4.value = Date
Text1.Text = Date
Text2.Text = Date
DataCombo1(14).Enabled = False
cdbhf = cdbh
VSFlexGrid1.ColWidth(1) = 1500
VSFlexGrid1.ColWidth(2) = 800
VSFlexGrid1.ColWidth(3) = 800
VSFlexGrid1.ColWidth(4) = 800
VSFlexGrid1.ColWidth(5) = 1200

VSFlexGrid3.ColWidth(0) = 200
For i = 1 To 18
VSFlexGrid3.ColWidth(i) = 600
Next

DataCombo1(21) = "公斤"
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset


Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT  日期,锅号,客户名称,色别,色名 as 色号,品名,'' as 毛坯幅宽,isnull(配缸重量,0) as 重量,isnull(配缸匹数,0) as 匹数,克重,类别,序号,生产状态,缸号,款号,'' as 备注,光胚幅宽,光坯匹数,光坯重量 FROM v_kpd_ok where 锅号= '" & DataCombo1(4).Text & "' order by  序号 "
Adodc2.Refresh

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT distinct 光坯单位  FROM GPDW"
Adodc8.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT 简称  FROM KHZL where ip like '%'+'" & yhxx & "'+'%'  GROUP BY 简称 "
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT mc  FROM bc  GROUP BY mc "
Adodc5.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT * FROM jgmxkfdj where 单据编号='" & yhdm & "'"
Adodc6.Refresh

DataCombo1(16).Text = Trim(yhdm) + "0000001"
If Adodc6.Recordset.EOF Then
DataCombo1(16).Text = Trim(yhdm) + "0000001"
Else
uu = Val(Adodc6.Recordset.Fields(1)) + 1
DataCombo1(16).Text = Trim(yhdm) + Left("0000000", 7 - Len(uu)) + Trim(uu)
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM jgmxkf where 锅号= '" & DataCombo1(4).Text & "' order by  缸号 desc"
Adodc1.Refresh

DataCombo1(14).Text = Adodc1.Recordset.RecordCount + 1



Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

       DataCombo1(12).Text = ""
       DataCombo1(4).Text = ""
       DataCombo1(4).TabIndex = 0
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

Private Sub Label1_Click(Index As Integer)
DataCombo1(14).Enabled = True
End Sub

Private Sub Label10_DBLClick(Index As Integer)
Select Case Index
       Case 0
Form343.Text3.Text = ""
Form343.Show
End Select
End Sub

Private Sub Label2_Click()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select *  FROM jgmxkf where  单据= '" & DataCombo1(16).Text & "' order by  序号 desc"
Adodc1.Refresh
DataCombo1(14).Text = Adodc1.Recordset.RecordCount + 1
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Text3_Change()
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text3 & "'+'%' and ip like '%'+'" & yhxx & "'+'%'  group by 简称"
Adodc4.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1

Adodc7.RecordSource = "SELECT * FROM JGMX WHERE 锅号='" & Adodc1.Recordset.Fields(4) & "' and 缸号='" & Adodc1.Recordset.Fields(17) & "'"
Adodc7.Refresh
If Adodc7.Recordset.EOF Then
For i = 1 To Adodc1.Recordset.Fields.count - 1
DataCombo1(i).Text = Adodc1.Recordset.Fields(i)
Next
DTPicker1.value = DataCombo1(0)
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
Else
MsgBox ("此入库单据已经出库，不能操作")
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
End If
End Sub

Private Sub VSFlexGrid3_dblClick()
On Error Resume Next
rs = VSFlexGrid3.Row
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
       DataCombo1(1).Text = Adodc2.Recordset.Fields(2)
       DataCombo1(2).Text = Adodc2.Recordset.Fields(14)
       DataCombo1(5).Text = Adodc2.Recordset.Fields(3)
       DataCombo1(6).Text = Adodc2.Recordset.Fields(5)
       DataCombo1(7).Text = Adodc2.Recordset.Fields(16)
       DataCombo1(10).Text = Adodc2.Recordset.Fields(7)   ''''光坯重量=毛坯重量
       DataCombo1(11).Text = Adodc2.Recordset.Fields(8) ''''光坯匹数=毛坯匹数
       DataCombo1(8).Text = Adodc2.Recordset.Fields(7)
       DataCombo1(12).Text = Adodc2.Recordset.Fields(10)  ''''类别
       DataCombo1(15).Text = ""  ''''备注
       DataCombo1(17).Text = Adodc2.Recordset.Fields(13)
       DataCombo1(18).Text = Adodc2.Recordset.Fields(9)
       DataCombo1(19).Text = Adodc2.Recordset.Fields(11)
       DataCombo1(20).Text = Adodc2.Recordset.Fields(4)
End Sub

