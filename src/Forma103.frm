VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma103 
   BackColor       =   &H00C0E0FF&
   Caption         =   "毛坯配缸"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   10320
      Top             =   9720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Forma103.frx":0000
      Left            =   15960
      List            =   "Forma103.frx":001C
      TabIndex        =   51
      Text            =   "Combo1"
      Top             =   2040
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   6720
      Top             =   9720
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
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   615
      Left            =   17040
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   1920
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   7560
      Top             =   10080
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8520
      TabIndex        =   44
      Text            =   "Text2"
      Top             =   5400
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8280
      Top             =   10200
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
      Left            =   8400
      Top             =   10200
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
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "校正"
      Height          =   615
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8520
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   1920
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   0
      Left            =   2760
      TabIndex        =   23
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
      Height          =   615
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   615
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   615
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8760
      Top             =   10320
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
      Left            =   8760
      Top             =   10440
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
      Left            =   8760
      Top             =   10320
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
      Height          =   375
      Left            =   8760
      Top             =   10440
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
      Bindings        =   "Forma103.frx":0040
      Height          =   330
      Index           =   0
      Left            =   7560
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   2
      Left            =   6480
      TabIndex        =   6
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma103.frx":0055
      Height          =   3855
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Width           =   18375
      _cx             =   32411
      _cy             =   6800
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   30
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Forma103.frx":006A
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forma103.frx":02BB
      Height          =   1575
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   18375
      _cx             =   32411
      _cy             =   2778
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   30
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Forma103.frx":02D0
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   1
      Left            =   480
      TabIndex        =   24
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   2
      Left            =   5040
      TabIndex        =   25
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   3
      Left            =   7440
      TabIndex        =   26
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   4
      Left            =   9840
      TabIndex        =   27
      Top             =   3360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   5
      Left            =   11040
      TabIndex        =   28
      Top             =   3360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   6
      Left            =   12240
      TabIndex        =   29
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   7
      Left            =   13320
      TabIndex        =   30
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma103.frx":0521
      Height          =   330
      Index           =   3
      Left            =   2400
      TabIndex        =   31
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   8
      Left            =   15240
      TabIndex        =   34
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo2"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma103.frx":0536
      Height          =   1575
      Left            =   480
      TabIndex        =   35
      Top             =   240
      Width           =   18135
      _cx             =   31988
      _cy             =   2778
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   30
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Forma103.frx":054B
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   14640
      TabIndex        =   39
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330170369
      CurrentDate     =   39177
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "毛坯码单"
      Height          =   615
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   9
      Left            =   13440
      TabIndex        =   41
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Forma103.frx":079C
      Height          =   330
      Index           =   10
      Left            =   16080
      TabIndex        =   46
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "xm"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma103.frx":07B1
      Height          =   330
      Index           =   4
      Left            =   4560
      TabIndex        =   47
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "卡号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   15480
      TabIndex        =   50
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   15
      Left            =   4560
      TabIndex        =   48
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "配缸负责"
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
      Index           =   14
      Left            =   16080
      TabIndex        =   45
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   43
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "实际匹数"
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
      Index           =   13
      Left            =   13440
      TabIndex        =   42
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "出库日期"
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
      Index           =   12
      Left            =   14640
      TabIndex        =   38
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "存放位置"
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
      Index           =   11
      Left            =   15240
      TabIndex        =   33
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯款号"
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
      Index           =   10
      Left            =   5040
      TabIndex        =   22
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯入库单据"
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
      Index           =   9
      Left            =   480
      TabIndex        =   21
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯入库序号"
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
      Left            =   2760
      TabIndex        =   20
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
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
      Left            =   480
      TabIndex        =   19
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
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
      Left            =   2400
      TabIndex        =   18
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Height          =   375
      Index           =   3
      Left            =   6480
      TabIndex        =   17
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯库存"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯简码"
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   15
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯布类"
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
      Index           =   4
      Left            =   7440
      TabIndex        =   14
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯幅宽"
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
      Index           =   5
      Left            =   9840
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯匹数"
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
      Index           =   6
      Left            =   11040
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯重量"
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
      Index           =   7
      Left            =   12240
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "毛坯备注"
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
      Index           =   8
      Left            =   13320
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7560
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "Forma103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BAR As Integer: Dim cdbhf As Integer
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command1_Click()
On Error Resume Next
If MsgBox("确定删除吗？，锅号" + DataCombo1(1) + "" + 序号 + DataCombo1(2), vbYesNo) = vbNo Then Exit Sub
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.Delete
'Adodc2.RecordSource = "select * from v_mp_kc where 库存重量<>0 and 客户名称='" & DataCombo1(0) & "' and 简码 like '%'+'" & Text2 & "'+'%' order by 日期,单据号,序号"
Adodc2.Refresh
Adodc1.RecordSource = "select *  from v_mp_kpd_pg where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc1.Refresh
Adodc3.RecordSource = "select * from mpbh where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc3.Refresh
DataCombo2(5) = ""
DataCombo2(6) = ""
DataCombo2(9) = ""
Adodc6.RecordSource = "select * from v_mp_kc where 库存重量<0 and 客户名称='" & DataCombo1(0) & "' and 简码 like '%'+'" & Text2 & "'+'%' order by 日期,单据号,序号"
Adodc6.Refresh
sql1 = "update kpd  set zt='备布取消',pb='" & Now & "' where 锅号='" & DataCombo1(1) & "' "
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
If Not Adodc6.Recordset.EOF Then
MsgBox ("入库单据： " + Adodc6.Recordset.Fields(1) + "入库序号： " + Adodc6.Recordset.Fields(0) + "出现负库存" + "请确认原因")
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗？，锅号" + DataCombo1(1) + "" + 序号 + DataCombo1(2), vbYesNo) = vbNo Then Exit Sub
If Adodc3.Recordset.EOF Then Exit Sub
If DataCombo2(5) = "" Or DataCombo2(6) = "" Or DataCombo2(10) = "" Or DataCombo1(1) = "" Then
MsgBox ("输入不正确!")
Exit Sub
End If
For i = 0 To 8
Adodc3.Recordset.Fields(i) = DataCombo2(i)
Next
Adodc3.Recordset.Fields(9) = DataCombo1(1)
Adodc3.Recordset.Fields(10) = DataCombo1(2)
Adodc3.Recordset.Fields(12) = DTPicker1.value
Adodc3.Recordset.Fields(13) = Val(DataCombo2(9))
Adodc3.Recordset.Fields(14) = DataCombo2(10)
Adodc3.Recordset.Fields(15) = DataCombo1(4)
Adodc3.Recordset.Update
'Adodc2.RecordSource = "select * from v_mp_kc where 库存重量<>0 and 客户名称='" & DataCombo1(0) & "' and 简码 like '%'+'" & Text2 & "'+'%' order by 日期,单据号,序号"
Adodc2.Refresh
Adodc1.RecordSource = "select *  from v_mp_kpd_pg where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc1.Refresh
Adodc3.RecordSource = "select * from mpbh where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc3.Refresh
DataCombo2(5) = ""
DataCombo2(6) = ""
DataCombo2(9) = ""
Adodc6.RecordSource = "select * from v_mp_kc where 库存重量<0 and 单据号='" & DataCombo2(1) & "' and 序号='" & DataCombo2(0) & "'"
Adodc6.Refresh
If Not Adodc6.Recordset.EOF Then
MsgBox ("入库单据： " + Adodc6.Recordset.Fields(1) + "入库序号： " + Trim(Adodc6.Recordset.Fields(0)) + "出现负库存" + "请确认原因")
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
If DataCombo1(0) = "" Or DataCombo1(1) = "" Or DataCombo1(2) = "" Or DataCombo1(3) = "" Or DataCombo2(10) = "" Then
MsgBox ("输入不完整，不能保存")
Exit Sub
End If

Adodc9.RecordSource = "select 库存重量 from v_mp_kc where 单据号='" & DataCombo2(1) & "' and 序号='" & DataCombo2(0) & "'"
Adodc9.Refresh
If Not Adodc9.Recordset.EOF Then
If Val(DataCombo2(6)) > Val(Adodc9.Recordset.Fields(0)) Then
MsgBox ("出库超出库存")
Exit Sub
End If
Else
MsgBox ("没有库存信息")
Exit Sub
End If

DataCombo2(9) = Val(DataCombo2(9))
'If MsgBox("确定结转吗？结转到的日期为" + Trim(DTPicker3.Value), vbYesNo) = vbNo Then Exit Sub
    Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "mpbhlr2('" & DataCombo2(0) & "','" & DataCombo2(1) & "','" & DataCombo2(2) & "','" & DataCombo2(3) & "','" & DataCombo2(4) & "','" & DataCombo2(5) & "','" & DataCombo2(6) & "','" & DataCombo2(7) & "','" & DataCombo2(8) & "','" & DataCombo1(1) & "','" & DataCombo1(2) & "','" & DTPicker1.value & "','" & DataCombo2(9) & "','" & DataCombo2(10) & "','" & DataCombo1(4) & "')"    ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

Adodc2.Refresh
Adodc1.RecordSource = "select *  from v_mp_kpd_pg where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc1.Refresh
Adodc3.RecordSource = "select * from mpbh where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc3.Refresh
Adodc6.RecordSource = "select * from v_mp_kc where 库存重量<0 and 单据号='" & DataCombo2(1) & "' and 序号='" & DataCombo2(0) & "'"
Adodc6.Refresh

sql1 = "update kpd  set zt='已备布待染',pb='" & Now & "' where 锅号='" & DataCombo1(1) & "' "
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

If Not Adodc6.Recordset.EOF Then
MsgBox ("入库单据： " + Adodc6.Recordset.Fields(1) + "入库序号： " + Trim(Adodc6.Recordset.Fields(0)) + "出现负库存" + "请确认原因")
End If
End Sub


Private Sub Command5_Click()
If Val(DataCombo2(5)) > 0 And DataCombo2(1) <> "" Then
Forma110.Text1(0) = DataCombo1(1)
Forma110.Text1(2) = DataCombo2(4)
Forma110.Text1(1) = DataCombo2(3)
Forma110.Text1(5) = DataCombo2(1)
Forma110.Text1(6) = DataCombo1(2)
Forma110.Show
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
If MsgBox("确定校正吗？", vbYesNo) = vbNo Then Exit Sub
Adodc5.RecordSource = "select * from v_mpbh_pg where 锅号='" & DataCombo1(1) & "'"
Adodc5.Refresh
If Not Adodc5.Recordset.EOF Then
sql1 = "update  kpd set 重量=(select round(sum(isnull(配缸重量,0)),2) from v_mpbh_pg where 锅号='" & DataCombo1(1) & "' and 缸号='" & DataCombo1(4) & "') where 锅号='" & DataCombo1(1) & "' and 编号='" & DataCombo1(4) & "'"
sql2 = "update  kpd set 匹数=(select round(sum(isnull(配缸匹数,0)),2) from v_mpbh_pg where 锅号='" & DataCombo1(1) & "' and 缸号='" & DataCombo1(4) & "') where 锅号='" & DataCombo1(1) & "' and 编号='" & DataCombo1(4) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
MsgBox ("校正成功！")
Adodc1.Refresh
Else
MsgBox ("没有完成出库，不能校正！")
End If
End Sub

Private Sub Command7_Click()
If Combo1 = "" Then
MsgBox ("请选择卡号")
Exit Sub
End If
Call mpckdy(Adodc8, DataCombo1(1), Combo1)
End Sub

Private Sub DataCombo1_Change(Index As Integer)
Select Case Index
       Case 1
If InStr(DataCombo1(1), "j") > 0 Or InStr(DataCombo1(1), "J") > 0 Then
DataCombo1(1) = Mid(DataCombo1(1), 1, Len(DataCombo1(1)) - 1)
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select *  from v_mp_kpd_pg where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc1.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select *  from mpbh where 锅号='" & DataCombo1(1) & "' order by 序号"
Adodc3.Refresh
End Select
End Sub

Private Sub DataCombo2_Change(Index As Integer)
Select Case Index
       Case 5
DataCombo2(9) = DataCombo2(5)
End Select
End Sub

Private Sub dataCombo2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Form_Load()
On Error Resume Next
DTPicker1.value = Date
DTPicker3.value = Date
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select xm  from fzr group by xm"
Adodc7.Refresh
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
For i = 0 To 10
DataCombo1(i).Text = ""
DataCombo2(i).Text = ""
Next
DataCombo2(10).Text = "杨春荣"
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(2) = 1700
VSFlexGrid1.ColWidth(5) = 1700
Text1.TabIndex = 0
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

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 0
Text1 = ""
End Select
End Sub

Private Sub Text1_Change()
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text1 & "'+'%'  group by 简称"
Adodc4.Refresh
End Sub

Private Sub Text2_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from v_mp_kc where  库存重量<>0 and 客户名称='" & DataCombo1(0) & "' and 简码 like '%'+'" & Text2 & "'+'%' order by 日期,单据号,序号"
Adodc2.Refresh
End Sub

Private Sub Text3_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from v_mp_kc where  库存重量<>0 and 客户名称='" & DataCombo1(0) & "' and 和约号 like '%'+'" & Text3 & "'+'%' order by 日期,单据号,序号"
Adodc2.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
For i = 0 To 4
If i = 1 Then i = 2
DataCombo1(i) = Adodc1.Recordset.Fields(i)
Next
Combo1 = Adodc1.Recordset.Fields(19)
'DataCombo2(5) = Adodc1.Recordset.Fields(15)
'DataCombo2(6) = Adodc1.Recordset.Fields(16)
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc2.Recordset.Move rs - 1
DataCombo2(0) = Adodc2.Recordset.Fields(0)
DataCombo2(1) = Adodc2.Recordset.Fields(1)
DataCombo2(2) = Adodc2.Recordset.Fields(4)
DataCombo2(3) = Adodc2.Recordset.Fields(5)
DataCombo2(4) = Adodc2.Recordset.Fields(6)
DataCombo2(5) = Adodc2.Recordset.Fields(14)
DataCombo2(6) = Adodc2.Recordset.Fields(15)
If DataCombo2(8) > Adodc2.Recordset.Fields(15) Then
DataCombo2(5) = Adodc2.Recordset.Fields(14)
DataCombo2(6) = Adodc2.Recordset.Fields(15)
End If
End Sub

Private Sub VSFlexGrid3_dblClick()
On Error Resume Next
If Adodc3.Recordset.EOF Then Exit Sub
rs = VSFlexGrid3.Row
Adodc3.Recordset.MoveFirst
Adodc3.Recordset.Move rs - 1
For i = 0 To 8
DataCombo2(i) = Adodc3.Recordset.Fields(i)
Next
DataCombo2(10) = Adodc3.Recordset.Fields(14)
DataCombo1(1) = Adodc3.Recordset.Fields(9)
DataCombo1(2) = Adodc3.Recordset.Fields(10)
End Sub
